"""
confluence_md_v2.py - Confluence 차일드 페이지 수집 + LLM 분석 (v5 GUI 적용)

개선사항:
  1. v5 의 모던 Samsung 스타일 GUI 적용
  2. html2text 로 문서 순서 유지한 정확한 MD 변환
  3. 이미지 다운로드 → base64 → Vision LLM 으로 실제 내용 분석
  4. 페이지별 LLM 요약 (단순 텍스트 자르기 X)
  5. MD 파일 선택 → 보고서 생성 탭 유지
"""

import os, re, base64, threading, traceback, json, subprocess, sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, List, Optional
from datetime import datetime
from bs4 import BeautifulSoup
import html2text
from openai import OpenAI

# ─── html2text 및 openai 자동 설치 ────────────
def _install(pkg):
    subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"],
                          stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

try:
    import html2text
except ImportError:
    print("html2text 설치 중..."); _install("html2text"); import html2text

try:
    from openai import OpenAI
except ImportError:
    print("openai 설치 중..."); _install("openai"); from openai import OpenAI

# ─── 색상 팔레트 (그레이/미니멀 스타일) ──────────────
BLUE       = "#5A6C7D"      # 차분한 회색계열 블루 (-muted)
BLUE_DK    = "#4A5568"      # 어두운 회색
BLUE_LT    = "#E2E8F0"      # 연한 회색
CYAN       = "#718096"      # 회색계열 시안
BG         = "#E8E9EB"      # 회색 배경
CARD       = "#F7F8FA"      # 밝은 회색 카드
BORDER     = "#D1D5DB"      # 회색 테두리
TEXT       = "#2D3748"      # 어두운 회색 텍스트
TEXT_MUTED = "#718096"      # muted 텍스트
GREEN      = "#68756E"      # 차분한 회색계열 녹색
RED        = "#A05A5A"      # 차분한 회색계열 빨간색
ORANGE     = "#B58A5A"      # 차분한 회색계열 주황색
LOG_BG     = "#2D3748"      # 어두운 로그 배경
LOG_FG     = "#E8E9EB"      # 밝은 로그 텍스트
LOG_GREEN  = "#68756E"      # 로그 성공
LOG_RED    = "#A05A5A"      # 로그 오류
LOG_BLUE   = "#5A6C7D"      # 로그 정보
LOG_YELLOW = "#B58A5A"      # 로그 경고

# ─────────────────────────────────────────────
# 설정 (회사 환경에 맞게 수정)
# ─────────────────────────────────────────────
BASE_URL        = os.getenv("CONFLUENCE_BASE_URL", "https://confluence.sec.samsung.net")
USER_DATA_DIR   = os.getenv("CONFLUENCE_PROFILE_DIR", "./chrome_profile_confluence_md")

# 설정 파일 (스크립트 폴더 옆에 저장)
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "llm_config.json")

# 기본값 초기화 (내부 서버 기본 연결)
LLM_API_KEY     = os.getenv("LLM_API_KEY", "")
LLM_BASE_URL    = os.getenv("LLM_BASE_URL", "http://10.240.246.158:8000/v1")
LLM_MODEL       = os.getenv("LLM_MODEL", "Qwen3.5-122B")
LLM_VISION_MODEL= os.getenv("LLM_VISION_MODEL", "Qwen3.5-122B")
LLM_MAX_TOKENS  = int(os.getenv("LLM_MAX_TOKENS", "4096"))

def load_llm_config():
    """llm_config.json 에서 설정 불러오기"""
    global LLM_API_KEY, LLM_BASE_URL, LLM_MODEL, LLM_VISION_MODEL, LLM_MAX_TOKENS
    if os.path.exists(CONFIG_FILE):
        try:
            cfg = json.loads(open(CONFIG_FILE, encoding="utf-8").read())
            LLM_API_KEY    = cfg.get("api_key", LLM_API_KEY)
            LLM_BASE_URL   = cfg.get("base_url", LLM_BASE_URL)
            LLM_MODEL      = cfg.get("model", LLM_MODEL)
            LLM_VISION_MODEL = cfg.get("vision_model", LLM_VISION_MODEL)
            LLM_MAX_TOKENS = int(cfg.get("max_tokens", LLM_MAX_TOKENS))
        except Exception as e:
            print(f"[설정 불러오기 실패] {e}")

def save_llm_config(api_key, base_url, model, vision_model, max_tokens):
    """llm_config.json 에 설정 저장"""
    global LLM_API_KEY, LLM_BASE_URL, LLM_MODEL, LLM_VISION_MODEL, LLM_MAX_TOKENS
    LLM_API_KEY    = api_key
    LLM_BASE_URL   = base_url
    LLM_MODEL      = model
    LLM_VISION_MODEL = vision_model
    LLM_MAX_TOKENS = int(max_tokens)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump({
            "api_key": api_key,
            "base_url": base_url,
            "model": model,
            "vision_model": vision_model,
            "max_tokens": max_tokens
        }, f, ensure_ascii=False, indent=2)

# 시작 시 저장된 설정 불러오기
load_llm_config()

def _llm_ready():
    # API 키는 로컬 서버에서는 생략 가능하므로 Base URL 만 검증
    return LLM_BASE_URL not in ("", "https://your-endpoint/v1", "http://10.240.246.158:8000/v1")

# ─────────────────────────────────────────────
# LLM 클라이언트
# ─────────────────────────────────────────────
def _llm():
    # 로컬 서버는 API 키가 필요 없을 수 있음
    if LLM_API_KEY and LLM_API_KEY not in ("", "sk-"):
        return OpenAI(api_key=LLM_API_KEY, base_url=LLM_BASE_URL)
    else:
        # API 키 없이 연결 시도
        return OpenAI(base_url=LLM_BASE_URL, api_key="sk-ignored")


def llm_summarize_page(title: str, markdown_text: str) -> str:
    """페이지 MD 내용을 LLM 으로 요약"""
    try:
        resp = _llm().chat.completions.create(
            model=LLM_MODEL,
            messages=[
                {"role": "system", "content":
                    "당신은 기업 업무 보고서 작성 전문가입니다. "
                    "주어진 Confluence 페이지 내용을 분석하여 핵심만 요약하세요.\n"
                    "형식:\n"
                    "- **핵심 요약** (2~3 문장)\n"
                    "- **완료된 사항**\n"
                    "- **진행 중인 사항**\n"
                    "- **이슈 / 리스크**\n"
                    "한국어로 간결하게 작성. 없는 항목은 생략."},
                {"role": "user", "content":
                    f"페이지 제목: {title}\n\n내용:\n{markdown_text[:5000]}"}
            ],
            max_tokens=LLM_MAX_TOKENS,
            temperature=0.3,
        )
        return resp.choices[0].message.content
    except Exception as e:
        return f"[LLM 요약 실패: {e}]"


def llm_analyze_image(image_bytes: bytes, context: str = "") -> str:
    """이미지 바이트를 Vision LLM 으로 분석"""
    try:
        b64 = base64.b64encode(image_bytes).decode()
        messages = [
            {"role": "system", "content":
                "당신은 기업 문서의 이미지/차트/다이어그램을 분석하는 전문가입니다. "
                "이미지에 담긴 핵심 정보를 한국어로 설명하세요. "
                "수치, 추세, 상태(완료/진행/이슈) 를 중심으로 3~5문장으로 요약하세요."},
            {"role": "user", "content": [
                {"type": "text",
                 "text": f"다음 Confluence 페이지 이미지를 분석해주세요.{(' 참고: ' + context) if context else ''}"},
                {"type": "image_url",
                 "image_url": {"url": f"data:image/png;base64,{b64}", "detail": "high"}}
            ]}
        ]
        resp = _llm().chat.completions.create(
            model=LLM_VISION_MODEL,
            messages=messages,
            max_tokens=500,
            temperature=0.3,
        )
        return resp.choices[0].message.content
    except Exception as e:
        return f"[이미지 분석 실패: {e}]"


def llm_generate_report(selected_md_contents: List[Dict]) -> str:
    """여러 MD 파일을 종합해 보고서 생성"""
    combined = ""
    for item in selected_md_contents:
        combined += f"\n\n### {item['title']}\n{item['content']}"

    try:
        resp = _llm().chat.completions.create(
            model=LLM_MODEL,
            messages=[
                {"role": "system", "content":
                    "당신은 400 명 규모 조직의 팀장을 위한 종합 업무 보고서 작성 전문가입니다.\n\n"
                    "【작성 원칙】\n"
                    "1. 반드시 제공된 문서 내용만을 근거로 작성하세요. 추측하거나 내용을 창작하지 마세요.\n"
                    "2. 팀장은 각 담당자의 전문 기술을 잘 모를 수 있습니다. 전문 용어가 나오면 괄호 안에 쉬운 말로 설명을 덧붙이세요.\n"
                    "3. 격식 있는 문어체로 작성하세요. (~하였습니다, ~진행 중에 있습니다)\n"
                    "4. 누가 읽어도 이해할 수 있도록 쉽고 구체적으로 풀어 쓰세요.\n"
                    "5. 수치, 날짜, 완료 여부 등 구체적인 사실은 그대로 포함하세요.\n\n"
                    "【보고서 구성】\n"
                    "# 1. 전체 요약\n"
                    "(전체 내용을 처음 보는 사람도 이해할 수 있게 3~5 문장으로 요약)\n\n"
                    "# 2. 항목별 상세 현황\n"
                    "(각 페이지/팀별로, 업무 내용 + 현재 상태를 자세히 설명)\n\n"
                    "# 3. 완료된 주요 사항\n"
                    "(이번 기간 내 완료된 것들을 불릿 포인트로 나열, 없으면 '완료된 항목 없음' 표기)\n\n"
                    "# 4. 진행 중인 주요 과제\n"
                    "(현재 진행 중인 것들과 예상 완료 시점, 진행률 포함)\n\n"
                    "# 5. 이슈 및 리스크\n"
                    "(문제가 되거나 주의가 필요한 사항, 없으면 '특이사항 없음' 표기)\n\n"
                    "# 6. 다음 단계 / 액션 아이템\n"
                    "(다음 주/월 동안 진행할 예정인 사항과 담당자)\n\n"
                    "한국어 격식체로 작성. 문서에 없는 내용은 절대 추가하지 않음."},
                {"role": "user", "content":
                    f"아래 {len(selected_md_contents)}개 페이지의 내용을 바탕으로 종합 보고서를 작성해주세요.\n\n"
                    f"반드시 아래 제공된 내용만 사용하고, 없는 내용은 만들지 마세요.\n\n"
                    f"=== 각 페이지 내용 ===\n{combined}"}
            ],
            max_tokens=LLM_MAX_TOKENS * 3,
            temperature=0.2,
        )
        return resp.choices[0].message.content
    except Exception as e:
        return f"[보고서 생성 실패: {e}]"


# ─────────────────────────────────────────────
# HTML → Markdown 변환 (문서 순서 보존)
# ─────────────────────────────────────────────
def html_to_markdown(html: str, page_session=None, base_url: str = "") -> str:
    """
    html2text 로 문서 순서를 유지한 MD 변환.
    이미지는 Playwright 세션으로 다운로드 후 Vision LLM 분석.
    """
    soup = BeautifulSoup(html, "html.parser")

    # ── 이미지 처리 (Vision LLM) ──────────────────
    for img in soup.find_all("img"):
        src = img.get("src", "")
        alt = img.get("alt", "이미지")

        if not src:
            continue

        img_desc = None

        if page_session and src:
            try:
                full_src = src if src.startswith("http") else base_url + src
                # Playwright 세션으로 인증 포함 다운로드
                resp = page_session.request.get(full_src)
                if resp.status == 200:
                    img_bytes = resp.body()
                    if len(img_bytes) > 1000:   # 너무 작은 아이콘은 스킵
                        img_desc = llm_analyze_image(img_bytes, context=alt)
            except Exception as e:
                img_desc = f"[이미지 다운로드 실패: {e}]"

        # 이미지를 설명 블록으로 교체
        if img_desc:
            replacement = soup.new_tag("p")
            replacement.string = f"📸 **[이미지 분석: {alt}]** {img_desc}"
            img.replace_with(replacement)
        else:
            replacement = soup.new_tag("p")
            replacement.string = f"📸 **[이미지: {alt}]** (분석 불가 - src: {src[:80]})"
            img.replace_with(replacement)

    # ── html2text 설정 ────────────────────────────
    h = html2text.HTML2Text()
    h.ignore_links      = False
    h.ignore_images     = True   # 이미 위에서 처리함
    h.body_width        = 0      # 줄바꿈 없음
    h.protect_links     = True
    h.wrap_links        = False
    h.unicode_snob      = True
    h.ignore_emphasis   = False
    h.mark_code         = True

    return h.handle(str(soup))


# ─────────────────────────────────────────────
# Confluence API
# ─────────────────────────────────────────────
def clean_filename(title: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "_", title).strip()

def extract_page_id_from_url(url: str) -> Optional[str]:
    m = re.search(r'/pages/(\d+)', url)
    if m: return m.group(1)
    m = re.search(r'pageId=(\d+)', url)
    if m: return m.group(1)
    return None

def get_page_content(page_session, page_id: str) -> Optional[Dict]:
    try:
        resp = page_session.request.get(
            f"{BASE_URL}/rest/api/content/{page_id}",
            params={"expand": "body.storage"}
        )
        if resp.status != 200:
            return None
        d = resp.json()
        return {
            "id":    d.get("id"),
            "title": d.get("title", ""),
            "html":  d.get("body", {}).get("storage", {}).get("value", "")
        }
    except:
        return None

def get_children(page_session, parent_id: str) -> List[Dict]:
    pages, start = [], 0
    while True:
        resp = page_session.request.get(
            f"{BASE_URL}/rest/api/content/search",
            params={"cql": f"ancestor={parent_id} and type=page",
                    "start": str(start), "limit": "50", "expand": "version"}
        )
        if resp.status != 200: break
        data    = resp.json()
        results = data.get("results", [])
        if not results: break
        for doc in results:
            pages.append({"id": doc["id"], "title": doc["title"]})
        if len(pages) >= data.get("size", 0): break
        start += 50
        if start > 500: break
    return pages


# ─────────────────────────────────────────────
# 페이지 처리 (재귀)
# ─────────────────────────────────────────────
def process_page(page_session, page_id: str, save_dir: str,
                 depth: int, use_vision: bool, use_llm_summary: bool,
                 callback=None):

    def log(msg):
        if callback: callback(msg)

    data = get_page_content(page_session, page_id)
    if not data:
        log(f"  [실패] 페이지 가져오기 실패: {page_id}")
        return

    title = data["title"]
    log(f"  처리 중: {title}")

    # HTML → MD
    ps = page_session if use_vision else None
    md_body = html_to_markdown(data["html"], page_session=ps, base_url=BASE_URL)

    # LLM 요약
    llm_summary = ""
    if use_llm_summary:
        log(f"    LLM 요약 중: {title}")
        llm_summary = llm_summarize_page(title, md_body)

    # 파일 저장
    md_lines = [
        f"# {title}", "",
        f"---",
        f"페이지 ID: {page_id}",
        f"생성일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"---", "",
    ]
    if llm_summary:
        md_lines += ["## 🤖 AI 요약", "", llm_summary, "", "---", ""]
    md_lines += ["## 원문 내용", "", md_body]

    safe = clean_filename(title)
    path = os.path.join(save_dir, f"{page_id}_{safe}.md")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines))
    log(f"  [저장] {path}")

    # 자식 페이지 재귀
    if depth > 0:
        children = get_children(page_session, page_id)
        if children:
            child_dir = os.path.join(save_dir, f"sub_{safe}")
            os.makedirs(child_dir, exist_ok=True)
            for child in children:
                process_page(page_session, child["id"], child_dir,
                             depth - 1, use_vision, use_llm_summary, callback)


# ─────────────────────────────────────────────
# GUI - Apple 스타일
# ─────────────────────────────────────────────

def setup_apple_style():
    """Apple 스타일 ttk 스타일 설정"""
    style = ttk.Style()
    style.theme_use('clam')  # 더 커스터마이징 가능한 테마
    
    # 전역 스타일 설정
    style.configure('.',
                    background=BG,
                    foreground=TEXT,
                    font=('SF Pro Display', 11))
    
    # 프레임 스타일 (카드 효과)
    style.configure('Card.TFrame',
                    background=CARD,
                    relief='flat')
    
    # 라벨 스타일
    style.configure('TLabel',
                    background=BG,
                    foreground=TEXT,
                    font=('SF Pro Display', 10))
    style.configure('Title.TLabel',
                    font=('SF Pro Display', 14, 'bold'),
                    foreground=BLUE)
    style.configure('Muted.TLabel',
                    foreground=TEXT_MUTED)
    
    # 버튼 스타일 (Apple 스타일)
    style.configure('TButton',
                    background=BLUE,
                    foreground='white',
                    font=('SF Pro Display', 10, 'bold'),
                    padding=(16, 8))
    style.map('TButton',
              background=[('active', BLUE_DK), ('disabled', BORDER)])
    
    # 입력 필드 스타일
    style.configure('TEntry',
                    fieldbackground=BG,
                    foreground=TEXT,
                    bordercolor=BLUE,
                    focuscolor=BLUE,
                    padding=(10, 8))
    
    # 체크박스/라디오
    style.configure('TCheckbutton',
                    background=BG,
                    foreground=TEXT)
    
    # Progressbar
    style.configure('TProgressbar',
                    background=BLUE,
                    troughcolor=BORDER)
    
    # Notebook (탭)
    style.configure('TNotebook',
                    background=CARD,
                    borderwidth=0)
    style.configure('TNotebook.Tab',
                    background=BG,
                    foreground=TEXT_MUTED,
                    font=('SF Pro Display', 11),
                    padding=[20, 10])
    style.map('TNotebook.Tab',
              background=[('selected', CARD)],
              foreground=[('selected', BLUE)])
    
    # Listbox
    style.configure('TListbox',
                    background=CARD,
                    foreground=TEXT,
                    bordercolor=BORDER,
                    selectbackground=BLUE_LT,
                    selectforeground=BLUE_DK)
    
    # Scrollbar
    style.configure('TScrollbar',
                    background=BORDER,
                    troughcolor=BG)
    
    # Separator
    style.configure('TSeparator',
                    background=BORDER)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Confluence MD 수집 + 보고서 생성 v2")
        self.geometry("900x700")
        self.resizable(True, True)
        
        # Apple 스타일 적용
        setup_apple_style()

        nb = ttk.Notebook(self, style='TNotebook')
        nb.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        self.tab_crawl    = CrawlTab(nb)
        self.tab_report   = ReportTab(nb)
        self.tab_settings = SettingsTab(nb)

        nb.add(self.tab_crawl,    text="📥 수집")
        nb.add(self.tab_report,   text="📊 보고서 생성")
        nb.add(self.tab_settings, text="⚙️ LLM 설정")


# ── 수집 탭 (Apple 스타일) ──────────────────────────────────
class CrawlTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, style='Card.TFrame', padding=20)
        self._build()

    def _build(self):
        # URL
        ttk.Label(self, text="Confluence URL:", style="Title.TLabel").grid(row=0, column=0, sticky=tk.W, pady=8)
        self.url_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.url_var, width=65, style='TEntry').grid(row=0, column=1, columnspan=2, sticky=tk.EW, pady=8)
        self.url_var.trace_add("write", self._on_url)

        self.pid_lbl = ttk.Label(self, text="페이지 ID: —", style="Muted.TLabel")
        self.pid_lbl.grid(row=1, column=1, sticky=tk.W, pady=4)

        # 깊이
        ttk.Label(self, text="재귀 깊이:", style="TLabel").grid(row=2, column=0, sticky=tk.W, pady=8)
        self.depth_var = tk.IntVar(value=3)
        ttk.Spinbox(self, from_=0, to=10, textvariable=self.depth_var, width=8).grid(row=2, column=1, sticky=tk.W)

        # 옵션
        self.vision_var  = tk.BooleanVar(value=True)
        self.summary_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self, text="이미지 Vision AI 분석", variable=self.vision_var, style='TCheckbutton').grid(row=3, column=1, sticky=tk.W, pady=4)
        ttk.Checkbutton(self, text="페이지별 LLM 요약 생성", variable=self.summary_var, style='TCheckbutton').grid(row=4, column=1, sticky=tk.W, pady=4)

        # 출력 폴더
        ttk.Label(self, text="출력 폴더:", style="TLabel").grid(row=5, column=0, sticky=tk.W, pady=8)
        self.out_var = tk.StringVar(value="./confluence_output")
        frm = ttk.Frame(self, style='Card.TFrame')
        frm.grid(row=5, column=1, columnspan=2, sticky=tk.EW)
        ttk.Entry(frm, textvariable=self.out_var, width=50, style='TEntry').pack(side=tk.LEFT)
        ttk.Button(frm, text="찾아보기", command=self._browse, style='TButton').pack(side=tk.LEFT, padx=8)

        # 실행
        self.btn = ttk.Button(self, text="▶ 수집 시작", command=self._run, style='TButton')
        self.btn.grid(row=6, column=0, pady=12)

        self.prog = ttk.Progressbar(self, mode="indeterminate", style='TProgressbar', length=700)
        self.prog.grid(row=7, column=0, columnspan=3, sticky=tk.EW)

        # 로그
        ttk.Label(self, text="로그:", style="Title.TLabel").grid(row=8, column=0, sticky=tk.W, pady=8)
        lf = ttk.Frame(self, style='Card.TFrame')
        lf.grid(row=9, column=0, columnspan=3, sticky=tk.NSEW)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(9, weight=1)
        sb = ttk.Scrollbar(lf, style='TScrollbar')
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_box = tk.Text(lf, height=16, yscrollcommand=sb.set, wrap=tk.WORD, bg=CARD, fg=TEXT, font=('SF Pro Display', 9))
        self.log_box.pack(fill=tk.BOTH, expand=True)
        sb.config(command=self.log_box.yview)

    def _on_url(self, *_):
        pid = extract_page_id_from_url(self.url_var.get())
        if pid:
            self.pid_lbl.config(text=f"페이지 ID: {pid}", foreground="green")
        else:
            self.pid_lbl.config(text="페이지 ID: 찾을 수 없음", foreground="red")

    def _browse(self):
        d = filedialog.askdirectory()
        if d: self.out_var.set(d)

    def log(self, msg):
        try:
            self.log_box.insert(tk.END, msg + "\n")
            self.log_box.see(tk.END)
            self.update_idletasks()
        except: pass

    def _run(self):
        pid = extract_page_id_from_url(self.url_var.get())
        if not pid:
            messagebox.showerror("오류", "URL에서 페이지 ID를 찾을 수 없습니다.")
            return
        self.btn.config(state="disabled")
        self.prog.start()
        threading.Thread(target=self._worker, args=(pid,), daemon=True).start()

    def _worker(self, page_id):
        from playwright.sync_api import sync_playwright
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

                if not messagebox.askokcancel("로그인 확인",
                        "Confluence에 로그인되어 있나요?\n[확인] 진행 / [취소] 중단"):
                    browser.close()
                    return

                page.reload(wait_until="networkidle")
                self.log("=" * 50)
                self.log(f"수집 시작 | 페이지 ID: {page_id}")
                self.log(f"Vision AI: {self.vision_var.get()} | LLM 요약: {self.summary_var.get()}")
                self.log("=" * 50)

                process_page(
                    page_session  = page,
                    page_id       = page_id,
                    save_dir      = save_dir,
                    depth         = self.depth_var.get(),
                    use_vision    = self.vision_var.get(),
                    use_llm_summary = self.summary_var.get(),
                    callback      = self.log,
                )
                browser.close()
                self.log("\n✅ 수집 완료!")
                self.log(f"저장 위치: {os.path.abspath(save_dir)}")
                messagebox.showinfo("완료", f"수집 완료!\n{os.path.abspath(save_dir)}")
        except Exception as e:
            self.log(f"[오류] {e}\n{traceback.format_exc()}")
            messagebox.showerror("오류", str(e))
        finally:
            self.after(0, lambda: (self.prog.stop(), self.btn.config(state="normal")))


# ── 보고서 생성 탭 (Apple 스타일) ────────────────────────────
class ReportTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, style='Card.TFrame', padding=20)
        self.md_files: List[str] = []
        self._build()

    def _build(self):
        ttk.Label(self, text="MD 파일 선택 (여러 개 가능):", style="Title.TLabel").grid(row=0, column=0, sticky=tk.W, pady=8)

        btn_frm = ttk.Frame(self, style='Card.TFrame')
        btn_frm.grid(row=0, column=1, sticky=tk.W, pady=8)
        ttk.Button(btn_frm, text="📂 폴더에서 MD 불러오기", command=self._load_folder, style='TButton').pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frm, text="📄 파일 직접 선택", command=self._load_files, style='TButton').pack(side=tk.LEFT, padx=4)

        # 파일 목록
        lf = ttk.LabelFrame(self, text="선택된 파일 목록", style='Card.TFrame')
        lf.grid(row=1, column=0, columnspan=2, sticky=tk.NSEW, pady=10)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        sb = ttk.Scrollbar(lf, style='TScrollbar')
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_list = tk.Listbox(lf, selectmode=tk.EXTENDED, height=10,
                                    yscrollcommand=sb.set, bg=CARD, fg=TEXT,
                                    selectbackground=BLUE_LT, selectforeground=BLUE_DK,
                                    font=('SF Pro Display', 10))
        self.file_list.pack(fill=tk.BOTH, expand=True)
        sb.config(command=self.file_list.yview)
        ttk.Button(lf, text="선택 항목 제거", command=self._remove_selected, style='TButton').pack(anchor=tk.E, padx=4, pady=4)

        # 보고서 옵션
        opt_frm = ttk.LabelFrame(self, text="보고서 옵션", style='Card.TFrame')
        opt_frm.grid(row=2, column=0, columnspan=2, sticky=tk.EW, pady=10)
        self.report_title_var = tk.StringVar(value="주간 보고서")
        ttk.Label(opt_frm, text="보고서 제목:", style="TLabel").grid(row=0, column=0, padx=6, pady=6, sticky=tk.W)
        ttk.Entry(opt_frm, textvariable=self.report_title_var, width=40, style='TEntry').grid(row=0, column=1, sticky=tk.W)

        self.out_var = tk.StringVar(value="./reports")
        ttk.Label(opt_frm, text="저장 폴더:", style="TLabel").grid(row=1, column=0, padx=6, pady=6, sticky=tk.W)
        frm2 = ttk.Frame(opt_frm, style='Card.TFrame')
        frm2.grid(row=1, column=1, sticky=tk.W)
        ttk.Entry(frm2, textvariable=self.out_var, width=40, style='TEntry').pack(side=tk.LEFT)
        ttk.Button(frm2, text="찾아보기",
                   command=lambda: self.out_var.set(filedialog.askdirectory() or self.out_var.get()),
                   style='TButton').pack(side=tk.LEFT, padx=8)

        # 실행
        self.btn = ttk.Button(self, text="📊 보고서 생성 (LLM)", command=self._generate, style='TButton')
        self.btn.grid(row=3, column=0, pady=12)

        self.prog = ttk.Progressbar(self, mode="indeterminate", style='TProgressbar', length=700)
        self.prog.grid(row=4, column=0, columnspan=2, sticky=tk.EW)

        # 로그
        ttk.Label(self, text="로그:", style="Title.TLabel").grid(row=5, column=0, sticky=tk.W, pady=8)
        lf2 = ttk.Frame(self, style='Card.TFrame')
        lf2.grid(row=6, column=0, columnspan=2, sticky=tk.NSEW, pady=8)
        self.rowconfigure(6, weight=1)
        sb2 = ttk.Scrollbar(lf2, style='TScrollbar')
        sb2.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_box = tk.Text(lf2, height=10, yscrollcommand=sb2.set, wrap=tk.WORD,
                               bg=CARD, fg=TEXT, font=('SF Pro Display', 9))
        self.log_box.pack(fill=tk.BOTH, expand=True)
        sb2.config(command=self.log_box.yview)

    def log(self, msg):
        """로그 출력"""
        try:
            self.log_box.insert(tk.END, msg + "\n")
            self.log_box.see(tk.END)
            self.update_idletasks()
        except:
            pass

    def _load_folder(self):
        """폴더에서 모든 .md 파일 로드"""
        folder = filedialog.askdirectory(title="MD 파일이 있는 폴더 선택")
        if not folder:
            return

        count = 0
        for f in os.listdir(folder):
            if f.endswith(".md"):
                path = os.path.join(folder, f)
                if path not in self.md_files:
                    self.md_files.append(path)
                    self.file_list.insert(tk.END, f)
                    count += 1
        self.log(f"✅ {count}개 파일을 로드했습니다.")

    def _load_files(self):
        """여러 MD 파일을 직접 선택"""
        files = filedialog.askopenfilenames(
            title="MD 파일 선택",
            filetypes=[("Markdown files", "*.md"), ("All files", "*.*")]
        )
        if not files:
            return

        count = 0
        for path in files:
            if path not in self.md_files:
                self.md_files.append(path)
                self.file_list.insert(tk.END, os.path.basename(path))
                count += 1
        self.log(f"✅ {count}개 파일을 로드했습니다.")

    def _remove_selected(self):
        """선택된 파일 목록 제거"""
        selected = self.file_list.curselection()
        if not selected:
            messagebox.showinfo("알림", "제거할 파일을 선택해주세요.")
            return

        # 역순으로 제거 (인덱스 변경 방지)
        for i in reversed(selected):
            self.file_list.delete(i)
            if i < len(self.md_files):
                self.md_files.pop(i)
        self.log(f"🗑️ {len(selected)}개 파일을 제거했습니다.")

    def _generate(self):
        """선택된 MD 파일들을 종합해 보고서 생성"""
        if not self.md_files:
            messagebox.showwarning("경고", "생성할 MD 파일이 없습니다.\n'폴더에서 MD 불러오기' 또는 '파일 직접 선택'을 이용해주세요.")
            return

        # 파일 내용 읽기
        selected_contents = []
        for path in self.md_files:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    content = f.read()
                    # 제목 추출 (첫 번째 줄)
                    title = os.path.basename(path)
                    if content.startswith("# "):
                        first_line = content.split("\n")[0]
                        if first_line.startswith("# "):
                            title = first_line[2:].strip()
                    selected_contents.append({"title": title, "content": content})
            except Exception as e:
                self.log(f"⚠️ 파일 읽기 실패: {path} - {e}")

        if not selected_contents:
            messagebox.showerror("오류", "모든 파일 읽기에 실패했습니다.")
            return

        self.btn.config(state="disabled")
        self.prog.start()
        self.log(f"📊 {len(selected_contents)}개 파일을 종합 중...")

        def _worker():
            try:
                report = llm_generate_report(selected_contents)
                self.after(0, lambda: self._save_report(report))
            except Exception as e:
                self.after(0, lambda: (
                    self.log(f"[오류] {e}"),
                    messagebox.showerror("오류", str(e)),
                    self.prog.stop(),
                    self.btn.config(state="normal")
                ))

        threading.Thread(target=_worker, daemon=True).start()

    def _save_report(self, report: str):
        """보고서 저장"""
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

            self.prog.stop()
            self.btn.config(state="normal")
            self.log(f"✅ 보고서 생성 완료: {path}")
            messagebox.showinfo("완료", f"보고서 생성 완료!\n\n{path}")
        except Exception as e:
            self.prog.stop()
            self.btn.config(state="normal")
            self.log(f"[오류] 보고서 저장 실패: {e}")
            messagebox.showerror("오류", f"보고서 저장 실패:\n{e}")


# ── LLM 설정 탭 (Apple 스타일) ────────────────────────────────
class SettingsTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, style='Card.TFrame', padding=20)
        self._build()

    def _build(self):
        self.columnconfigure(1, weight=1)

        ttk.Label(self, text="루코드 LLM 연결 설정", style="Title.TLabel").grid(
            row=0, column=0, columnspan=2, pady=(0, 16), sticky="w")

        # API Key
        ttk.Label(self, text="API Key:", style="TLabel").grid(row=1, column=0, sticky="w", pady=8)
        key_frm = ttk.Frame(self, style='Card.TFrame')
        key_frm.grid(row=1, column=1, sticky="ew", pady=8)
        self.llm_key_var = tk.StringVar(value=LLM_API_KEY)
        self.key_entry = ttk.Entry(key_frm, textvariable=self.llm_key_var, width=55, show="*", style='TEntry')
        self.key_entry.pack(side="left", fill="x", expand=True)
        self._show_key = False
        def toggle_key():
            self._show_key = not self._show_key
            self.key_entry.config(style='TEntry' if self._show_key else 'TEntry')
            self.key_entry.config(show="" if self._show_key else "*")
            show_btn.config(text="숨기기" if self._show_key else "보기")
        show_btn = ttk.Button(key_frm, text="보기", width=6, command=toggle_key, style='TButton')
        show_btn.pack(side="left", padx=6)

        # Base URL
        ttk.Label(self, text="Base URL:", style="TLabel").grid(row=2, column=0, sticky="w", pady=8)
        self.llm_url_var = tk.StringVar(value=LLM_BASE_URL)
        url_entry = ttk.Entry(self, textvariable=self.llm_url_var, width=62, style='TEntry')
        url_entry.grid(row=2, column=1, sticky="ew", pady=8)
        ttk.Label(self, text="  예) http://10.240.246.158:8000/v1 (기본값)", style="Muted.TLabel").grid(
            row=3, column=1, sticky="w")

        # 모델명
        ttk.Label(self, text="모델명:", style="TLabel").grid(row=4, column=0, sticky="w", pady=8)
        self.llm_model_var = tk.StringVar(value=LLM_MODEL)
        ttk.Entry(self, textvariable=self.llm_model_var, width=30, style='TEntry').grid(row=4, column=1, sticky="w", pady=8)
        ttk.Label(self, text="  예) Qwen3.5-122B", style="Muted.TLabel").grid(row=5, column=1, sticky="w")

        # Vision 모델명
        ttk.Label(self, text="Vision 모델:", style="TLabel").grid(row=6, column=0, sticky="w", pady=8)
        self.llm_vision_var = tk.StringVar(value=LLM_VISION_MODEL)
        ttk.Entry(self, textvariable=self.llm_vision_var, width=30, style='TEntry').grid(row=6, column=1, sticky="w", pady=8)

        # 최대 토큰
        ttk.Label(self, text="최대 토큰:", style="TLabel").grid(row=7, column=0, sticky="w", pady=8)
        self.llm_tokens_var = tk.StringVar(value=str(LLM_MAX_TOKENS))
        ttk.Entry(self, textvariable=self.llm_tokens_var, width=10, style='TEntry').grid(row=7, column=1, sticky="w", pady=8)

        # 버튼
        btn_frm = ttk.Frame(self, style='Card.TFrame')
        btn_frm.grid(row=8, column=0, columnspan=2, pady=16, sticky="w")
        ttk.Button(btn_frm, text="💾 설정 저장", command=self._save_settings, style='TButton').pack(side="left", padx=6)
        ttk.Button(btn_frm, text="🔗 연결 테스트", command=self._test_llm, style='TButton').pack(side="left", padx=6)

        self.settings_status = ttk.Label(self, text="", style="Muted.TLabel")
        self.settings_status.grid(row=9, column=0, columnspan=2, sticky="w", pady=8)

        # 도움말
        ttk.Separator(self, orient="horizontal").grid(row=10, column=0, columnspan=2, sticky="ew", pady=12)
        info = ("📌 설정 방법\n"
                "1. 루코드 서버 주소를 Base URL 에 입력하세요.\n"
                "2. API Key 는 루코드에서 발급받은 키를 입력하세요 (로컬 서버는 생략 가능).\n"
                "3. 모델명은 서버에서 제공하는 모델명으로 입력하세요.\n"
                "4. '설정 저장' 버튼을 누르면 다음 실행 시 자동으로 불러옵니다.")
        ttk.Label(self, text=info, style="Muted.TLabel", justify="left").grid(
            row=11, column=0, columnspan=2, sticky="w")

    def _save_settings(self):
        key    = self.llm_key_var.get().strip()
        url    = self.llm_url_var.get().strip()
        model  = self.llm_model_var.get().strip()
        vision = self.llm_vision_var.get().strip()
        tokens = self.llm_tokens_var.get().strip()
        if not url:
            self.settings_status.config(text="⚠️ Base URL 은 필수입니다.", foreground="red")
            return
        try:
            save_llm_config(key, url, model, vision, int(tokens))
            self.settings_status.config(text=f"✅ 저장 완료 → {CONFIG_FILE}", foreground="green")
            messagebox.showinfo("완료", "설정이 저장되었습니다.\n다음 실행 시 자동으로 적용됩니다.")
        except Exception as e:
            self.settings_status.config(text=f"❌ 저장 실패: {e}", foreground="red")

    def _test_llm(self):
        # 현재 설정으로 테스트
        self.settings_status.config(text="🔄 연결 테스트 중...", foreground="blue")
        self.update_idletasks()
        def _test():
            try:
                client = OpenAI(api_key=LLM_API_KEY, base_url=LLM_BASE_URL)
                resp = client.chat.completions.create(
                    model=LLM_MODEL,
                    messages=[{"role": "user", "content": "연결 테스트입니다. 한 문장으로 답해주세요."}],
                    max_tokens=100, temperature=0.1,
                )
                answer = resp.choices[0].message.content.strip()[:80]
                self.after(0, lambda: self.settings_status.config(
                    text=f"✅ 연결 성공! 응답: {answer}", foreground="green"))
            except Exception as e:
                self.after(0, lambda: self.settings_status.config(
                    text=f"❌ 연결 실패: {e}", foreground="red"))
        threading.Thread(target=_test, daemon=True).start()


# ─────────────────────────────────────────────
if __name__ == "__main__":
    App().mainloop()
