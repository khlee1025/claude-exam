
"""
confluence_md_v2_ocr.py

Confluence child page collector + Markdown exporter + LLM report generator.

Key change from the original version:
  - Qwen is treated as a text model.
  - Images are downloaded, OCR'd, and the extracted text is interpreted by Qwen.
  - This mirrors the common Open WebUI flow where uploaded images are converted
    to text before being sent to a non-VL LLM.
"""

import base64
import glob
import io
import json
import os
import re
import subprocess
import sys
import threading
import traceback
from datetime import datetime
from typing import Dict, List, Optional
from urllib.parse import urljoin

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


# ---------------------------------------------------------------------------
# Optional dependency loader
# ---------------------------------------------------------------------------

def _install(pkg: str):
    subprocess.check_call(
        [sys.executable, "-m", "pip", "install", pkg, "-q"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def _import_or_install(import_name: str, pip_name: Optional[str] = None):
    try:
        return __import__(import_name)
    except ImportError:
        _install(pip_name or import_name)
        return __import__(import_name)


def _need_html2text():
    return _import_or_install("html2text")


def _need_beautifulsoup():
    bs4 = _import_or_install("bs4", "beautifulsoup4")
    return bs4.BeautifulSoup


def _need_openai_class():
    openai_mod = _import_or_install("openai")
    return openai_mod.OpenAI


# ---------------------------------------------------------------------------
# Palette
# ---------------------------------------------------------------------------

BLUE       = "#5A6C7D"
BLUE_DK    = "#4A5568"
BLUE_LT    = "#E2E8F0"
BG         = "#E8E9EB"
CARD       = "#F7F8FA"
BORDER     = "#D1D5DB"
TEXT       = "#2D3748"
TEXT_MUTED = "#718096"
LOG_BG     = "#2D3748"
LOG_FG     = "#E8E9EB"


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

BASE_URL      = os.getenv("CONFLUENCE_BASE_URL", "https://confluence.sec.samsung.net")
USER_DATA_DIR = os.getenv("CONFLUENCE_PROFILE_DIR", "./chrome_profile_confluence_md")
CONFIG_FILE   = os.path.join(os.path.dirname(os.path.abspath(__file__)), "llm_config.json")

# 기본값 초기화 (내부 서버 기본 연결)
LLM_API_KEY      = os.getenv("LLM_API_KEY", "")
LLM_BASE_URL     = os.getenv("LLM_BASE_URL", "http://10.240.246.158:8000/v1")
LLM_MODEL        = os.getenv("LLM_MODEL", "Qwen3.5-122B")
LLM_VISION_MODEL = os.getenv("LLM_VISION_MODEL", "Qwen3.5-122B")
LLM_MAX_TOKENS   = int(os.getenv("LLM_MAX_TOKENS", "4096"))

# OCR settings. EasyOCR is preferred; pytesseract is fallback.
OCR_LANGS = os.getenv("OCR_LANGS", "ko,en")
OCR_AUTO_INSTALL = os.getenv("OCR_AUTO_INSTALL", "0") == "1"
_OCR_READER = None


def load_llm_config():
    global LLM_API_KEY, LLM_BASE_URL, LLM_MODEL, LLM_VISION_MODEL, LLM_MAX_TOKENS
    if not os.path.exists(CONFIG_FILE):
        return
    try:
        with open(CONFIG_FILE, encoding="utf-8") as f:
            cfg = json.load(f)
        LLM_API_KEY = cfg.get("api_key", LLM_API_KEY)
        LLM_BASE_URL = cfg.get("base_url", LLM_BASE_URL)
        LLM_MODEL = cfg.get("model", LLM_MODEL)
        LLM_VISION_MODEL = cfg.get("vision_model", LLM_VISION_MODEL)
        LLM_MAX_TOKENS = int(cfg.get("max_tokens", LLM_MAX_TOKENS))
    except Exception as e:
        print(f"[설정 불러오기 실패] {e}")


def save_llm_config(api_key, base_url, model, vision_model, max_tokens):
    global LLM_API_KEY, LLM_BASE_URL, LLM_MODEL, LLM_VISION_MODEL, LLM_MAX_TOKENS
    LLM_API_KEY = api_key
    LLM_BASE_URL = base_url
    LLM_MODEL = model
    LLM_VISION_MODEL = vision_model
    LLM_MAX_TOKENS = int(max_tokens)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(
            {
                "api_key": api_key,
                "base_url": base_url,
                "model": model,
                "vision_model": vision_model,
                "max_tokens": int(max_tokens),
            },
            f,
            ensure_ascii=False,
            indent=2,
        )


load_llm_config()


def _llm():
    OpenAI = _need_openai_class()
    return OpenAI(
        api_key=LLM_API_KEY if LLM_API_KEY else "sk-ignored",
        base_url=LLM_BASE_URL,
    )


# ---------------------------------------------------------------------------
# OCR + LLM image analysis
# ---------------------------------------------------------------------------

def extract_text_from_image(image_bytes: bytes) -> str:
    """
    Convert an image into text before sending anything to Qwen.

    Recommended local install:
      pip install pillow easyocr numpy

    Fallback:
      pip install pillow pytesseract
      plus local Tesseract engine + Korean language pack.
    """
    errors = []

    try:
        from PIL import Image
    except ImportError:
        if OCR_AUTO_INSTALL:
            _install("pillow")
            from PIL import Image
        else:
            return "[OCR 실패: pillow 미설치. pip install pillow 필요]"

    image = Image.open(io.BytesIO(image_bytes)).convert("RGB")

    try:
        global _OCR_READER
        import numpy as np
        import easyocr

        if _OCR_READER is None:
            langs = [x.strip() for x in OCR_LANGS.split(",") if x.strip()]
            _OCR_READER = easyocr.Reader(langs, gpu=False)

        result = _OCR_READER.readtext(np.array(image), detail=0, paragraph=True)
        text = "\n".join(x.strip() for x in result if x and x.strip())
        if text.strip():
            return text.strip()
    except ImportError as e:
        errors.append(f"easyocr 미설치: {e}")
    except Exception as e:
        errors.append(f"easyocr 실패: {type(e).__name__}: {e}")

    try:
        import pytesseract
        text = pytesseract.image_to_string(image, lang="kor+eng")
        if text.strip():
            return text.strip()
    except ImportError as e:
        errors.append(f"pytesseract 미설치: {e}")
    except Exception as e:
        errors.append(f"pytesseract 실패: {type(e).__name__}: {e}")

    return "[OCR 텍스트 없음]\n" + "\n".join(errors)


def llm_summarize_image_ocr(ocr_text: str, context: str = "") -> str:
    if not ocr_text.strip():
        return "[이미지 분석 실패: OCR 결과가 비어 있습니다.]"
    if ocr_text.startswith("[OCR 실패") or ocr_text.startswith("[OCR 텍스트 없음]"):
        return ocr_text

    try:
        resp = _llm().chat.completions.create(
            model=LLM_MODEL,
            messages=[
                {
                    "role": "system",
                    "content": (
                        "당신은 기업 문서의 이미지, 표, 차트, 화면 캡처를 분석하는 전문가입니다. "
                        "당신에게는 이미지 자체가 아니라 OCR로 추출된 텍스트만 제공됩니다. "
                        "제공된 텍스트만 근거로 핵심 내용을 한국어로 요약하세요. "
                        "수치, 날짜, 상태, 이슈, 리스크가 있으면 명확히 정리하세요. "
                        "OCR 오류로 보이는 부분은 확정하지 말고 '판독 불명확'이라고 표현하세요."
                    ),
                },
                {
                    "role": "user",
                    "content": (
                        f"이미지 참고 정보: {context or '없음'}\n\n"
                        f"이미지 OCR 결과:\n{ocr_text[:6000]}\n\n"
                        "위 OCR 결과를 바탕으로 이미지의 핵심 의미를 3~5문장으로 요약해주세요."
                    ),
                },
            ],
            max_tokens=min(1000, LLM_MAX_TOKENS),
            temperature=0.2,
        )
        return resp.choices[0].message.content
    except Exception as e:
        return f"[OCR 기반 이미지 요약 실패: {type(e).__name__}: {e}]"


def llm_analyze_image(image_bytes: bytes, context: str = "", mime_type: str = "image/png") -> str:
    """
    Non-VL Qwen path:
      image bytes -> OCR text -> Qwen text interpretation
    """
    try:
        ocr_text = extract_text_from_image(image_bytes)
        return llm_summarize_image_ocr(ocr_text, context=context)
    except Exception as e:
        return f"[이미지 분석 실패: {type(e).__name__}: {e}]"


def llm_summarize_page(title: str, markdown_text: str) -> str:
    try:
        resp = _llm().chat.completions.create(
            model=LLM_MODEL,
            messages=[
                {
                    "role": "system",
                    "content": (
                        "당신은 기업 업무 보고서 작성 전문가입니다. "
                        "주어진 Confluence 페이지 내용을 분석하여 핵심만 요약하세요.\n"
                        "형식:\n"
                        "- **핵심 요약** (2~3 문장)\n"
                        "- **완료된 사항**\n"
                        "- **진행 중인 사항**\n"
                        "- **이슈 / 리스크**\n"
                        "한국어로 간결하게 작성. 없는 항목은 생략."
                    ),
                },
                {"role": "user", "content": f"페이지 제목: {title}\n\n내용:\n{markdown_text[:5000]}"},
            ],
            max_tokens=LLM_MAX_TOKENS,
            temperature=0.3,
        )
        return resp.choices[0].message.content
    except Exception as e:
        return f"[LLM 요약 실패: {type(e).__name__}: {e}]"


def llm_generate_report(selected_md_contents: List[Dict]) -> str:
    combined = ""
    for item in selected_md_contents:
        combined += f"\n\n### {item['title']}\n{item['content']}"

    try:
        resp = _llm().chat.completions.create(
            model=LLM_MODEL,
            messages=[
                {
                    "role": "system",
                    "content": (
                        "당신은 400 명 규모 조직의 팀장을 위한 종합 업무 보고서 작성 전문가입니다.\n\n"
                        "【작성 원칙】\n"
                        "1. 반드시 제공된 문서 내용만을 근거로 작성하세요. 추측하거나 내용을 창작하지 마세요.\n"
                        "2. 전문 용어가 나오면 괄호 안에 쉬운 말로 설명을 덧붙이세요.\n"
                        "3. 격식 있는 문어체로 작성하세요. (~하였습니다, ~진행 중에 있습니다)\n"
                        "4. 쉽고 구체적으로 풀어 쓰세요.\n"
                        "5. 수치, 날짜, 완료 여부 등 구체적인 사실은 그대로 포함하세요.\n\n"
                        "【보고서 구성】\n"
                        "# 1. 전체 요약\n"
                        "# 2. 항목별 상세 현황\n"
                        "# 3. 완료된 주요 사항\n"
                        "# 4. 진행 중인 주요 과제\n"
                        "# 5. 이슈 및 리스크\n"
                        "# 6. 다음 단계 / 액션 아이템\n\n"
                        "한국어 격식체로 작성. 문서에 없는 내용은 절대 추가하지 않음."
                    ),
                },
                {
                    "role": "user",
                    "content": (
                        f"아래 {len(selected_md_contents)}개 페이지의 내용을 바탕으로 종합 보고서를 작성해주세요.\n\n"
                        f"반드시 아래 제공된 내용만 사용하고, 없는 내용은 만들지 마세요.\n\n"
                        f"=== 각 페이지 내용 ===\n{combined}"
                    ),
                },
            ],
            max_tokens=LLM_MAX_TOKENS * 3,
            temperature=0.2,
        )
        return resp.choices[0].message.content
    except Exception as e:
        return f"[보고서 생성 실패: {type(e).__name__}: {e}]"


# ---------------------------------------------------------------------------
# Markdown / Confluence helpers
# ---------------------------------------------------------------------------

def html_to_markdown(html: str, page_session=None, base_url: str = "") -> str:
    BeautifulSoup = _need_beautifulsoup()
    html2text = _need_html2text()
    soup = BeautifulSoup(html, "html.parser")

    for img in soup.find_all("img"):
        src = img.get("src", "")
        alt = img.get("alt", "이미지")
        if not src:
            continue

        img_desc = None
        if page_session and src:
            try:
                full_src = urljoin(base_url + "/", src)
                resp = page_session.request.get(full_src)
                if resp.status == 200:
                    img_bytes = resp.body()
                    content_type = resp.headers.get("content-type", "image/png").split(";")[0].strip().lower()
                    if content_type == "image/jpg":
                        content_type = "image/jpeg"

                    if content_type in ("image/png", "image/jpeg", "image/webp", "image/gif") and len(img_bytes) > 300:
                        img_desc = llm_analyze_image(img_bytes, context=alt, mime_type=content_type)
                    else:
                        img_desc = f"[이미지 분석 스킵: content-type={content_type}, size={len(img_bytes)} bytes]"
                else:
                    img_desc = f"[이미지 다운로드 실패: HTTP {resp.status}]"
            except Exception as e:
                img_desc = f"[이미지 다운로드/OCR 실패: {type(e).__name__}: {e}]"

        replacement = soup.new_tag("p")
        if img_desc:
            replacement.string = f"**[이미지 OCR 분석: {alt}]** {img_desc}"
        else:
            replacement.string = f"**[이미지: {alt}]** (분석 불가 - src: {src[:120]})"
        img.replace_with(replacement)

    h = html2text.HTML2Text()
    h.ignore_links = False
    h.ignore_images = True
    h.body_width = 0
    h.protect_links = True
    h.wrap_links = False
    h.unicode_snob = True
    h.ignore_emphasis = False
    h.mark_code = True
    return h.handle(str(soup))


def clean_filename(title: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "_", title).strip()


def extract_page_id_from_url(url: str) -> Optional[str]:
    m = re.search(r"/pages/(\d+)", url)
    if m:
        return m.group(1)
    m = re.search(r"pageId=(\d+)", url)
    if m:
        return m.group(1)
    return None


def get_page_content(page_session, page_id: str) -> Optional[Dict]:
    try:
        resp = page_session.request.get(
            f"{BASE_URL}/rest/api/content/{page_id}",
            params={"expand": "body.storage"},
        )
        if resp.status != 200:
            return None
        d = resp.json()
        return {
            "id": d.get("id"),
            "title": d.get("title", ""),
            "html": d.get("body", {}).get("storage", {}).get("value", ""),
        }
    except Exception:
        return None


def get_children(page_session, parent_id: str) -> List[Dict]:
    pages, start = [], 0
    while True:
        resp = page_session.request.get(
            f"{BASE_URL}/rest/api/content/search",
            params={
                "cql": f"ancestor={parent_id} and type=page",
                "start": str(start),
                "limit": "50",
                "expand": "version",
            },
        )
        if resp.status != 200:
            break
        data = resp.json()
        results = data.get("results", [])
        if not results:
            break
        for doc in results:
            pages.append({"id": doc["id"], "title": doc["title"]})
        if len(pages) >= data.get("size", 0):
            break
        start += 50
        if start > 500:
            break
    return pages


def process_page(
    page_session,
    page_id: str,
    save_dir: str,
    depth: int,
    use_vision: bool,
    use_llm_summary: bool,
    callback=None,
):
    def log(msg):
        if callback:
            callback(msg)

    data = get_page_content(page_session, page_id)
    if not data:
        log(f"  [실패] 페이지 가져오기 실패: {page_id}")
        return

    title = data["title"]
    log(f"  처리 중: {title}")
    ps = page_session if use_vision else None
    md_body = html_to_markdown(data["html"], page_session=ps, base_url=BASE_URL)

    llm_summary = ""
    if use_llm_summary:
        log(f"    LLM 요약 중: {title}")
        llm_summary = llm_summarize_page(title, md_body)

    md_lines = [
        f"# {title}",
        "",
        "---",
        f"페이지 ID: {page_id}",
        f"생성일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "---",
        "",
    ]
    if llm_summary:
        md_lines += ["## AI 요약", "", llm_summary, "", "---", ""]
    md_lines += ["## 원문 내용", "", md_body]

    safe = clean_filename(title)
    path = os.path.join(save_dir, f"{page_id}_{safe}.md")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines))
    log(f"  [저장] {path}")

    if depth > 0:
        children = get_children(page_session, page_id)
        if children:
            child_dir = os.path.join(save_dir, f"sub_{safe}")
            os.makedirs(child_dir, exist_ok=True)
            for child in children:
                process_page(
                    page_session,
                    child["id"],
                    child_dir,
                    depth - 1,
                    use_vision,
                    use_llm_summary,
                    callback,
                )


# ---------------------------------------------------------------------------
# Markdown to Word converter
# ---------------------------------------------------------------------------

class MDConverter:
    def __init__(self):
        docx = _import_or_install("docx", "python-docx")
        self.Document = docx.Document
        self.Pt = __import__("docx.shared", fromlist=["Pt"]).Pt
        self.Cm = __import__("docx.shared", fromlist=["Cm"]).Cm
        self.RGBColor = __import__("docx.shared", fromlist=["RGBColor"]).RGBColor
        self.WD_ALIGN_PARAGRAPH = __import__("docx.enum.text", fromlist=["WD_ALIGN_PARAGRAPH"]).WD_ALIGN_PARAGRAPH

    def convert(self, md_text: str, title: str = "", add_cover: bool = True):
        doc = self.Document()
        for section in doc.sections:
            section.page_width = self.Cm(21)
            section.page_height = self.Cm(29.7)
            section.left_margin = section.right_margin = self.Cm(2.5)
            section.top_margin = section.bottom_margin = self.Cm(2.5)
        style = doc.styles["Normal"]
        style.font.name = "맑은 고딕"
        style.font.size = self.Pt(11)

        if add_cover and title:
            doc.add_paragraph()
            p = doc.add_paragraph()
            r = p.add_run(title)
            r.font.name = "맑은 고딕"
            r.font.size = self.Pt(22)
            r.bold = True
            p.alignment = self.WD_ALIGN_PARAGRAPH.CENTER
            p2 = doc.add_paragraph()
            r2 = p2.add_run(f"작성일: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
            r2.font.name = "맑은 고딕"
            r2.font.size = self.Pt(11)
            p2.alignment = self.WD_ALIGN_PARAGRAPH.CENTER
            doc.add_page_break()

        in_code = False
        code = []
        for line in md_text.splitlines():
            stripped = line.strip()
            if stripped.startswith("```"):
                if in_code:
                    p = doc.add_paragraph()
                    r = p.add_run("\n".join(code))
                    r.font.name = "Consolas"
                    r.font.size = self.Pt(9)
                    code = []
                    in_code = False
                else:
                    in_code = True
                continue
            if in_code:
                code.append(line)
                continue
            m = re.match(r"^(#{1,6})\s+(.*)", stripped)
            if m:
                doc.add_heading(m.group(2), level=min(len(m.group(1)), 4))
            elif re.match(r"^[-*•]\s+", stripped):
                doc.add_paragraph(re.sub(r"^[-*•]\s+", "", stripped), style="List Bullet")
            elif re.match(r"^\d+\.\s+", stripped):
                doc.add_paragraph(re.sub(r"^\d+\.\s+", "", stripped), style="List Number")
            elif stripped:
                doc.add_paragraph(stripped)
            else:
                doc.add_paragraph()
        return doc

    def convert_file(self, md_path: str, out_path: str = None, add_cover: bool = True, callback=None) -> str:
        with open(md_path, encoding="utf-8") as f:
            md_text = f.read()
        title = os.path.basename(md_path).replace(".md", "")
        if callback:
            callback(f"INFO: 변환 중: {title}")
        doc = self.convert(md_text, title=title, add_cover=add_cover)
        out_path = out_path or md_path.replace(".md", ".docx")
        doc.save(out_path)
        if callback:
            callback(f"OK: 저장 완료 -> {out_path}")
        return out_path


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

def setup_style():
    style = ttk.Style()
    style.theme_use("clam")
    style.configure(".", background=BG, foreground=TEXT, font=("Malgun Gothic", 10))
    style.configure("Card.TFrame", background=CARD, relief="flat")
    style.configure("TLabel", background=CARD, foreground=TEXT, font=("Malgun Gothic", 10))
    style.configure("Root.TFrame", background=BG)
    style.configure("Title.TLabel", background=CARD, font=("Malgun Gothic", 13, "bold"), foreground=BLUE_DK)
    style.configure("Muted.TLabel", background=CARD, foreground=TEXT_MUTED)
    style.configure("TButton", background=BLUE, foreground="white", font=("Malgun Gothic", 10, "bold"), padding=(14, 7))
    style.map("TButton", background=[("active", BLUE_DK), ("disabled", BORDER)])
    style.configure("TEntry", fieldbackground="white", foreground=TEXT, padding=(8, 6))
    style.configure("TCheckbutton", background=CARD, foreground=TEXT)
    style.configure("TProgressbar", background=BLUE, troughcolor=BORDER)
    style.configure("TNotebook", background=BG, borderwidth=0)
    style.configure("TNotebook.Tab", background=BG, foreground=TEXT_MUTED, padding=[18, 8])
    style.map("TNotebook.Tab", background=[("selected", CARD)], foreground=[("selected", BLUE_DK)])


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Confluence MD 수집 + OCR 보고서 생성 v2")
        self.geometry("940x720")
        self.resizable(True, True)
        self.configure(bg=BG)
        setup_style()

        nb = ttk.Notebook(self)
        nb.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        self.tab_crawl = CrawlTab(nb)
        self.tab_report = ReportTab(nb)
        self.tab_settings = SettingsTab(nb)
        nb.add(self.tab_crawl, text="수집")
        nb.add(self.tab_report, text="보고서 생성")
        nb.add(self.tab_settings, text="LLM 설정")


class CrawlTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, style="Card.TFrame", padding=20)
        self._build()

    def _build(self):
        self.columnconfigure(1, weight=1)
        self.rowconfigure(9, weight=1)

        ttk.Label(self, text="Confluence URL:", style="Title.TLabel").grid(row=0, column=0, sticky=tk.W, pady=8)
        self.url_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.url_var, width=70).grid(row=0, column=1, columnspan=2, sticky=tk.EW, pady=8)
        self.url_var.trace_add("write", self._on_url)
        self.pid_lbl = ttk.Label(self, text="페이지 ID: -", style="Muted.TLabel")
        self.pid_lbl.grid(row=1, column=1, sticky=tk.W, pady=4)

        ttk.Label(self, text="재귀 깊이:").grid(row=2, column=0, sticky=tk.W, pady=8)
        self.depth_var = tk.IntVar(value=3)
        ttk.Spinbox(self, from_=0, to=10, textvariable=self.depth_var, width=8).grid(row=2, column=1, sticky=tk.W)

        self.vision_var = tk.BooleanVar(value=True)
        self.summary_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self, text="이미지 OCR 분석", variable=self.vision_var).grid(row=3, column=1, sticky=tk.W, pady=4)
        ttk.Checkbutton(self, text="페이지별 LLM 요약 생성", variable=self.summary_var).grid(row=4, column=1, sticky=tk.W, pady=4)

        ttk.Label(self, text="출력 폴더:").grid(row=5, column=0, sticky=tk.W, pady=8)
        self.out_var = tk.StringVar(value="./confluence_output")
        frm = ttk.Frame(self, style="Card.TFrame")
        frm.grid(row=5, column=1, columnspan=2, sticky=tk.EW)
        frm.columnconfigure(0, weight=1)
        ttk.Entry(frm, textvariable=self.out_var).grid(row=0, column=0, sticky=tk.EW)
        ttk.Button(frm, text="찾아보기", command=self._browse).grid(row=0, column=1, padx=8)

        self.btn = ttk.Button(self, text="수집 시작", command=self._run)
        self.btn.grid(row=6, column=0, pady=12)
        self.prog = ttk.Progressbar(self, mode="indeterminate", length=700)
        self.prog.grid(row=7, column=0, columnspan=3, sticky=tk.EW)

        ttk.Label(self, text="로그:", style="Title.TLabel").grid(row=8, column=0, sticky=tk.W, pady=8)
        lf = ttk.Frame(self, style="Card.TFrame")
        lf.grid(row=9, column=0, columnspan=3, sticky=tk.NSEW)
        sb = ttk.Scrollbar(lf)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_box = tk.Text(lf, height=16, yscrollcommand=sb.set, wrap=tk.WORD, bg=LOG_BG, fg=LOG_FG, font=("Consolas", 9))
        self.log_box.pack(fill=tk.BOTH, expand=True)
        sb.config(command=self.log_box.yview)

    def _on_url(self, *_):
        pid = extract_page_id_from_url(self.url_var.get())
        self.pid_lbl.config(text=f"페이지 ID: {pid}" if pid else "페이지 ID: 찾을 수 없음")

    def _browse(self):
        d = filedialog.askdirectory()
        if d:
            self.out_var.set(d)

    def log(self, msg):
        try:
            self.log_box.insert(tk.END, msg + "\n")
            self.log_box.see(tk.END)
            self.update_idletasks()
        except Exception:
            pass

    def _run(self):
        pid = extract_page_id_from_url(self.url_var.get())
        if not pid:
            messagebox.showerror("오류", "URL에서 페이지 ID를 찾을 수 없습니다.")
            return
        self.btn.config(state="disabled")
        self.prog.start()
        threading.Thread(target=self._worker, args=(pid,), daemon=True).start()

    def _worker(self, page_id):
        try:
            from playwright.sync_api import sync_playwright
        except ImportError:
            self.after(0, lambda: messagebox.showerror("오류", "playwright 미설치: pip install playwright 후 playwright install chromium 실행 필요"))
            self.after(0, lambda: (self.prog.stop(), self.btn.config(state="normal")))
            return

        save_dir = os.path.join(self.out_var.get(), f"page_{page_id}")
        os.makedirs(save_dir, exist_ok=True)
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch_persistent_context(
                    user_data_dir=USER_DATA_DIR,
                    headless=False,
                    viewport={"width": 1280, "height": 720},
                )
                page = browser.new_page()
                page.goto(f"{BASE_URL}/pages/viewpage.action?pageId={page_id}")
                page.wait_for_load_state("networkidle", timeout=30000)

                if not messagebox.askokcancel("로그인 확인", "Confluence에 로그인되어 있나요?\n[확인] 진행 / [취소] 중단"):
                    browser.close()
                    return

                page.reload(wait_until="networkidle")
                self.log("=" * 50)
                self.log(f"수집 시작 | 페이지 ID: {page_id}")
                self.log(f"이미지 OCR: {self.vision_var.get()} | LLM 요약: {self.summary_var.get()}")
                self.log("=" * 50)

                process_page(
                    page_session=page,
                    page_id=page_id,
                    save_dir=save_dir,
                    depth=self.depth_var.get(),
                    use_vision=self.vision_var.get(),
                    use_llm_summary=self.summary_var.get(),
                    callback=self.log,
                )
                browser.close()
                self.log("\n수집 완료!")
                self.log(f"저장 위치: {os.path.abspath(save_dir)}")
                messagebox.showinfo("완료", f"수집 완료!\n{os.path.abspath(save_dir)}")
        except Exception as e:
            self.log(f"[오류] {e}\n{traceback.format_exc()}")
            messagebox.showerror("오류", str(e))
        finally:
            self.after(0, lambda: (self.prog.stop(), self.btn.config(state="normal")))


class ReportTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, style="Card.TFrame", padding=20)
        self.md_files: List[str] = []
        self._build()

    def _build(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.rowconfigure(6, weight=1)

        ttk.Label(self, text="MD 파일 선택:", style="Title.TLabel").grid(row=0, column=0, sticky=tk.W, pady=8)
        btn_frm = ttk.Frame(self, style="Card.TFrame")
        btn_frm.grid(row=0, column=1, sticky=tk.W, pady=8)
        ttk.Button(btn_frm, text="폴더에서 MD 불러오기", command=self._load_folder).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frm, text="파일 직접 선택", command=self._load_files).pack(side=tk.LEFT, padx=4)

        lf = ttk.LabelFrame(self, text="선택된 파일 목록")
        lf.grid(row=1, column=0, columnspan=2, sticky=tk.NSEW, pady=10)
        sb = ttk.Scrollbar(lf)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_list = tk.Listbox(lf, selectmode=tk.EXTENDED, height=10, yscrollcommand=sb.set, bg="white", fg=TEXT)
        self.file_list.pack(fill=tk.BOTH, expand=True)
        sb.config(command=self.file_list.yview)
        inner = ttk.Frame(lf)
        inner.pack(anchor=tk.E, padx=4, pady=4)
        ttk.Button(inner, text="선택 항목 제거", command=self._remove_selected).pack(side=tk.LEFT, padx=2)
        ttk.Button(inner, text="Word 로 변환", command=self._export_to_word).pack(side=tk.LEFT, padx=2)

        opt = ttk.LabelFrame(self, text="보고서 옵션")
        opt.grid(row=2, column=0, columnspan=2, sticky=tk.EW, pady=10)
        self.report_title_var = tk.StringVar(value="주간 보고서")
        self.out_var = tk.StringVar(value="./reports")
        ttk.Label(opt, text="보고서 제목:").grid(row=0, column=0, padx=6, pady=6, sticky=tk.W)
        ttk.Entry(opt, textvariable=self.report_title_var, width=40).grid(row=0, column=1, sticky=tk.W)
        ttk.Label(opt, text="저장 폴더:").grid(row=1, column=0, padx=6, pady=6, sticky=tk.W)
        frm2 = ttk.Frame(opt)
        frm2.grid(row=1, column=1, sticky=tk.W)
        ttk.Entry(frm2, textvariable=self.out_var, width=40).pack(side=tk.LEFT)
        ttk.Button(frm2, text="찾아보기", command=lambda: self.out_var.set(filedialog.askdirectory() or self.out_var.get())).pack(side=tk.LEFT, padx=8)

        self.word_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt, text="보고서 생성 후 Word(.docx) 로도 저장", variable=self.word_var).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=6)

        self.btn = ttk.Button(self, text="보고서 생성 (LLM)", command=self._generate)
        self.btn.grid(row=3, column=0, pady=12)
        self.prog = ttk.Progressbar(self, mode="indeterminate", length=700)
        self.prog.grid(row=4, column=0, columnspan=2, sticky=tk.EW)

        ttk.Label(self, text="로그:", style="Title.TLabel").grid(row=5, column=0, sticky=tk.W, pady=8)
        lf2 = ttk.Frame(self, style="Card.TFrame")
        lf2.grid(row=6, column=0, columnspan=2, sticky=tk.NSEW, pady=8)
        sb2 = ttk.Scrollbar(lf2)
        sb2.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_box = tk.Text(lf2, height=10, yscrollcommand=sb2.set, wrap=tk.WORD, bg=LOG_BG, fg=LOG_FG, font=("Consolas", 9))
        self.log_box.pack(fill=tk.BOTH, expand=True)
        sb2.config(command=self.log_box.yview)

    def log(self, msg):
        self.log_box.insert(tk.END, msg + "\n")
        self.log_box.see(tk.END)
        self.update_idletasks()

    def _load_folder(self):
        folder = filedialog.askdirectory(title="MD 파일이 있는 폴더 선택")
        if not folder:
            return
        count = 0
        for f in glob.glob(os.path.join(folder, "**", "*.md"), recursive=True):
            if f not in self.md_files:
                self.md_files.append(f)
                self.file_list.insert(tk.END, os.path.basename(f))
                count += 1
        self.log(f"{count}개 파일을 로드했습니다.")

    def _load_files(self):
        files = filedialog.askopenfilenames(title="MD 파일 선택", filetypes=[("Markdown files", "*.md"), ("All files", "*.*")])
        count = 0
        for path in files:
            if path not in self.md_files:
                self.md_files.append(path)
                self.file_list.insert(tk.END, os.path.basename(path))
                count += 1
        self.log(f"{count}개 파일을 로드했습니다.")

    def _remove_selected(self):
        selected = self.file_list.curselection()
        if not selected:
            messagebox.showinfo("알림", "제거할 파일을 선택해주세요.")
            return
        for i in reversed(selected):
            self.file_list.delete(i)
            if i < len(self.md_files):
                self.md_files.pop(i)
        self.log(f"{len(selected)}개 파일을 제거했습니다.")

    def _generate(self):
        if not self.md_files:
            messagebox.showwarning("경고", "생성할 MD 파일이 없습니다.")
            return
        selected_contents = []
        for path in self.md_files:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    content = f.read()
                title = os.path.basename(path)
                if content.startswith("# "):
                    title = content.split("\n", 1)[0][2:].strip()
                selected_contents.append({"title": title, "content": content})
            except Exception as e:
                self.log(f"파일 읽기 실패: {path} - {e}")
        if not selected_contents:
            messagebox.showerror("오류", "모든 파일 읽기에 실패했습니다.")
            return
        self.btn.config(state="disabled")
        self.prog.start()
        self.log(f"{len(selected_contents)}개 파일을 종합 중...")

        def _worker():
            try:
                report = llm_generate_report(selected_contents)
                self.after(0, lambda: self._save_report(report))
            except Exception as e:
                self.after(0, lambda: (self.log(f"[오류] {e}"), messagebox.showerror("오류", str(e)), self.prog.stop(), self.btn.config(state="normal")))

        threading.Thread(target=_worker, daemon=True).start()

    def _export_to_word(self):
        if not self.md_files:
            messagebox.showwarning("경고", "변환할 MD 파일이 없습니다.")
            return
        self.btn.config(state="disabled")
        self.prog.start()
        self.log(f"{len(self.md_files)}개 파일을 Word 로 변환 중...")

        def _worker():
            try:
                conv = MDConverter()
                out_dir = self.out_var.get().strip() or "./word_output"
                os.makedirs(out_dir, exist_ok=True)
                results = []
                for path in self.md_files:
                    try:
                        fname = os.path.basename(path).replace(".md", "")
                        out_path = os.path.join(out_dir, f"{fname}.docx")
                        conv.convert_file(path, out_path, add_cover=True, callback=self.log)
                        results.append(out_path)
                    except Exception as e:
                        self.log(f"변환 실패: {path} - {e}")
                self.after(0, lambda: (self.prog.stop(), self.btn.config(state="normal"), self.log(f"Word 변환 완료: {len(results)}개"), messagebox.showinfo("완료", f"Word 변환 완료!\n{len(results)}개 파일 생성:\n{out_dir}")))
            except Exception as e:
                self.after(0, lambda: (self.log(f"[오류] {e}"), messagebox.showerror("오류", str(e)), self.prog.stop(), self.btn.config(state="normal")))

        threading.Thread(target=_worker, daemon=True).start()

    def _save_report(self, report: str):
        try:
            os.makedirs(self.out_var.get(), exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{self.report_title_var.get().replace(' ', '_')}_{timestamp}.md"
            path = os.path.join(self.out_var.get(), filename)
            with open(path, "w", encoding="utf-8") as f:
                f.write(f"# {self.report_title_var.get()}\n\n")
                f.write(f"생성일: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                f.write("---\n\n")
                f.write(report)
            if self.word_var.get():
                out_docx = path.replace(".md", ".docx")
                MDConverter().convert_file(path, out_docx, add_cover=True, callback=self.log)
            self.prog.stop()
            self.btn.config(state="normal")
            self.log(f"보고서 생성 완료: {path}")
            messagebox.showinfo("완료", f"보고서 생성 완료!\n\n{path}")
        except Exception as e:
            self.prog.stop()
            self.btn.config(state="normal")
            self.log(f"[오류] 보고서 저장 실패: {e}")
            messagebox.showerror("오류", f"보고서 저장 실패:\n{e}")


class SettingsTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, style="Card.TFrame", padding=20)
        self._build()

    def _build(self):
        self.columnconfigure(1, weight=1)
        ttk.Label(self, text="루코드 LLM 연결 설정", style="Title.TLabel").grid(row=0, column=0, columnspan=2, pady=(0, 16), sticky="w")

        ttk.Label(self, text="API Key:").grid(row=1, column=0, sticky="w", pady=8)
        key_frm = ttk.Frame(self, style="Card.TFrame")
        key_frm.grid(row=1, column=1, sticky="ew", pady=8)
        self.llm_key_var = tk.StringVar(value=LLM_API_KEY)
        self.key_entry = ttk.Entry(key_frm, textvariable=self.llm_key_var, show="*")
        self.key_entry.pack(side="left", fill="x", expand=True)
        self._show_key = False

        def toggle_key():
            self._show_key = not self._show_key
            self.key_entry.config(show="" if self._show_key else "*")
            show_btn.config(text="숨기기" if self._show_key else "보기")

        show_btn = ttk.Button(key_frm, text="보기", width=6, command=toggle_key)
        show_btn.pack(side="left", padx=6)

        ttk.Label(self, text="Base URL:").grid(row=2, column=0, sticky="w", pady=8)
        self.llm_url_var = tk.StringVar(value=LLM_BASE_URL)
        ttk.Entry(self, textvariable=self.llm_url_var).grid(row=2, column=1, sticky="ew", pady=8)
        ttk.Label(self, text="예) http://10.240.246.158:8000/v1", style="Muted.TLabel").grid(row=3, column=1, sticky="w")

        ttk.Label(self, text="모델명:").grid(row=4, column=0, sticky="w", pady=8)
        self.llm_model_var = tk.StringVar(value=LLM_MODEL)
        ttk.Entry(self, textvariable=self.llm_model_var, width=32).grid(row=4, column=1, sticky="w", pady=8)

        ttk.Label(self, text="Vision/OCR 후처리 모델:").grid(row=5, column=0, sticky="w", pady=8)
        self.llm_vision_var = tk.StringVar(value=LLM_VISION_MODEL)
        ttk.Entry(self, textvariable=self.llm_vision_var, width=32).grid(row=5, column=1, sticky="w", pady=8)
        ttk.Label(self, text="현재 코드는 이미지를 직접 보내지 않고 OCR 텍스트를 Qwen에 보냅니다.", style="Muted.TLabel").grid(row=6, column=1, sticky="w")

        ttk.Label(self, text="최대 토큰:").grid(row=7, column=0, sticky="w", pady=8)
        self.llm_tokens_var = tk.StringVar(value=str(LLM_MAX_TOKENS))
        ttk.Entry(self, textvariable=self.llm_tokens_var, width=10).grid(row=7, column=1, sticky="w", pady=8)

        btn_frm = ttk.Frame(self, style="Card.TFrame")
        btn_frm.grid(row=8, column=0, columnspan=2, pady=16, sticky="w")
        ttk.Button(btn_frm, text="설정 저장", command=self._save_settings).pack(side="left", padx=6)
        ttk.Button(btn_frm, text="연결 테스트", command=self._test_llm).pack(side="left", padx=6)

        self.settings_status = ttk.Label(self, text="", style="Muted.TLabel")
        self.settings_status.grid(row=9, column=0, columnspan=2, sticky="w", pady=8)

        ttk.Separator(self, orient="horizontal").grid(row=10, column=0, columnspan=2, sticky="ew", pady=12)
        info = (
            "OCR 사용 안내\n"
            "1. 이미지 분석은 base64 이미지를 Qwen에 직접 보내지 않고 OCR 텍스트를 추출한 뒤 Qwen이 해석합니다.\n"
            "2. 권장 설치: pip install pillow easyocr numpy\n"
            "3. Tesseract를 쓰려면 pip install pillow pytesseract 및 로컬 Tesseract 설치가 필요합니다.\n"
            "4. OCR 미설치 상태에서도 GUI와 텍스트 요약 기능은 실행됩니다."
        )
        ttk.Label(self, text=info, style="Muted.TLabel", justify="left").grid(row=11, column=0, columnspan=2, sticky="w")

    def _save_settings(self):
        key = self.llm_key_var.get().strip()
        url = self.llm_url_var.get().strip()
        model = self.llm_model_var.get().strip()
        vision = self.llm_vision_var.get().strip()
        tokens = self.llm_tokens_var.get().strip()
        if not url:
            self.settings_status.config(text="Base URL 은 필수입니다.", foreground="red")
            return
        try:
            save_llm_config(key, url, model, vision, int(tokens))
            self.settings_status.config(text=f"저장 완료 -> {CONFIG_FILE}", foreground="green")
            messagebox.showinfo("완료", "설정이 저장되었습니다.")
        except Exception as e:
            self.settings_status.config(text=f"저장 실패: {e}", foreground="red")

    def _test_llm(self):
        self.settings_status.config(text="연결 테스트 중...", foreground="blue")
        self.update_idletasks()
        key = self.llm_key_var.get().strip()
        url = self.llm_url_var.get().strip()
        model = self.llm_model_var.get().strip()

        def _test():
            try:
                OpenAI = _need_openai_class()
                client = OpenAI(api_key=key if key else "sk-ignored", base_url=url)
                resp = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": "연결 테스트입니다. 한 문장으로 답해주세요."}],
                    max_tokens=100,
                    temperature=0.1,
                )
                answer = resp.choices[0].message.content.strip()[:120]
                self.after(0, lambda: self.settings_status.config(text=f"연결 성공! 응답: {answer}", foreground="green"))
            except Exception as e:
                self.after(0, lambda: self.settings_status.config(text=f"연결 실패: {type(e).__name__}: {e}", foreground="red"))

        threading.Thread(target=_test, daemon=True).start()


if __name__ == "__main__":
    App().mainloop()
