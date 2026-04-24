"""
Confluence 차일드 페이지 요약 GUI - v4
변경사항: 보고서 생성 시 Confluence 재호출 제거 → 저장된 MD 파일 직접 읽기
"""

# ── 패키지 자동 설치 ──────────────────────────
import subprocess, sys

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

# ── 기본 임포트 ───────────────────────────────
import os, re, threading, traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, List, Optional
from datetime import datetime
from bs4 import BeautifulSoup

# ─────────────────────────────────────────────
# 설정
# ─────────────────────────────────────────────
BASE_URL      = "https://confluence.sec.samsung.net"
USER_DATA_DIR = "./chrome_profile_confluence_md"

LLM_API_KEY   = os.getenv("LLM_API_KEY",   "")
LLM_BASE_URL  = os.getenv("LLM_BASE_URL",  "")
LLM_MODEL     = os.getenv("LLM_MODEL",     "qwen-plus")
LLM_MAX_TOKENS= int(os.getenv("LLM_MAX_TOKENS", "2000"))

# 설정 파일 (스크립트 폴더 옆에 저장)
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "llm_config.json")

def load_llm_config():
    """llm_config.json 에서 설정 불러오기"""
    global LLM_API_KEY, LLM_BASE_URL, LLM_MODEL, LLM_MAX_TOKENS
    if os.path.exists(CONFIG_FILE):
        try:
            import json
            cfg = json.loads(open(CONFIG_FILE, encoding="utf-8").read())
            LLM_API_KEY    = cfg.get("api_key",    LLM_API_KEY)
            LLM_BASE_URL   = cfg.get("base_url",   LLM_BASE_URL)
            LLM_MODEL      = cfg.get("model",      LLM_MODEL)
            LLM_MAX_TOKENS = int(cfg.get("max_tokens", LLM_MAX_TOKENS))
        except Exception as e:
            print(f"[설정 불러오기 실패] {e}")

def save_llm_config(api_key, base_url, model, max_tokens):
    """llm_config.json 에 설정 저장"""
    global LLM_API_KEY, LLM_BASE_URL, LLM_MODEL, LLM_MAX_TOKENS
    import json
    LLM_API_KEY    = api_key
    LLM_BASE_URL   = base_url
    LLM_MODEL      = model
    LLM_MAX_TOKENS = int(max_tokens)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump({"api_key": api_key, "base_url": base_url,
                   "model": model, "max_tokens": max_tokens}, f, ensure_ascii=False, indent=2)

# 시작 시 저장된 설정 불러오기
load_llm_config()

def _llm_ready():
    return LLM_API_KEY not in ("", "YOUR_API_KEY") and LLM_BASE_URL not in ("", "https://your-endpoint/v1")

# ─────────────────────────────────────────────
# LLM 함수
# ─────────────────────────────────────────────
def llm_summarize_page(title: str, text: str) -> str:
    try:
        client = OpenAI(api_key=LLM_API_KEY, base_url=LLM_BASE_URL)
        resp = client.chat.completions.create(
            model=LLM_MODEL,
            messages=[
                {"role": "system", "content":
                    "당신은 400명 규모 조직의 팀장을 위한 업무 보고서 작성 전문가입니다.\n\n"
                    "【절대 규칙】\n"
                    "- 제공된 문서 내용만 사용하세요. 없는 내용은 절대 만들지 마세요.\n"
                    "- 문서에 없는 항목은 '해당 내용이 문서에 명시되지 않았습니다.'라고 쓰세요.\n\n"
                    "【작성 방식】\n"
                    "- 전문 용어는 반드시 괄호로 쉽게 설명하세요.\n"
                    "  예) API 연동(서로 다른 시스템이 데이터를 주고받는 연결 작업)\n"
                    "- 격식체 사용. (~하였습니다, ~진행 중에 있습니다)\n"
                    "- 수치, 날짜, 완료 여부는 문서 그대로 포함하세요.\n\n"
                    "【보고서 형식】\n"
                    "## 개요\n"
                    "이 업무가 무엇인지 2~3문장으로 쉽게 설명\n\n"
                    "## 이번 주요 내용\n"
                    "문서에 기록된 사실만, 불릿 포인트로\n\n"
                    "## 완료된 사항\n"
                    "없으면 생략\n\n"
                    "## 진행 중인 사항\n"
                    "없으면 생략\n\n"
                    "## 이슈 및 특이사항\n"
                    "없으면 생략"},
                {"role": "user", "content":
                    f"페이지 제목: {title}\n\n=== 문서 내용 ===\n{text[:5000]}"}
            ],
            max_tokens=LLM_MAX_TOKENS,
            temperature=0.1,
        )
        return resp.choices[0].message.content
    except Exception as e:
        return f"[LLM 요약 실패: {e}]"


def llm_generate_report(pages: list) -> str:
    combined = ""
    for p in pages:
        combined += f"\n\n### {p['title']}\n{p['content'][:2000]}"
    try:
        client = OpenAI(api_key=LLM_API_KEY, base_url=LLM_BASE_URL)
        resp = client.chat.completions.create(
            model=LLM_MODEL,
            messages=[
                {"role": "system", "content":
                    "당신은 400명 규모 조직의 팀장을 위한 종합 보고서 작성 전문가입니다.\n\n"
                    "【절대 규칙】\n"
                    "- 제공된 내용만 사용. 없는 내용은 절대 만들지 마세요.\n"
                    "- 추측이나 창작 금지. 불확실하면 '문서에 명시되지 않음'으로 표기.\n\n"
                    "【작성 방식】\n"
                    "- 각 팀/담당자가 무슨 일을 하는지 처음 보는 사람도 이해하도록 설명\n"
                    "- 전문 용어는 괄호로 쉬운 말 추가\n"
                    "- 격식 있는 문어체 (~하였습니다, ~검토 중에 있습니다)\n"
                    "- 수치, 날짜, 완료율 등 구체적 수치 반드시 포함\n\n"
                    "【보고서 구성】\n"
                    "# 1. 전체 요약\n"
                    "처음 보는 사람도 이해할 수 있게 3~5문장\n\n"
                    "# 2. 항목별 상세 현황\n"
                    "각 페이지/팀별 업무 내용과 현재 상태\n\n"
                    "# 3. 완료된 주요 사항\n"
                    "없으면 생략\n\n"
                    "# 4. 진행 중인 주요 과제\n"
                    "없으면 생략\n\n"
                    "# 5. 이슈 및 리스크\n"
                    "없으면 생략\n\n"
                    "# 6. 팀장 조치 필요 사항\n"
                    "없으면 생략"},
                {"role": "user", "content":
                    f"{len(pages)}개 페이지 내용으로 팀장 보고용 종합 보고서를 작성해주세요.\n"
                    f"제공된 내용만 사용하세요.\n\n=== 내용 ===\n{combined}"}
            ],
            max_tokens=LLM_MAX_TOKENS * 2,
            temperature=0.1,
        )
        return resp.choices[0].message.content
    except Exception as e:
        return f"[보고서 생성 실패: {e}]"


# ─────────────────────────────────────────────
# 유틸
# ─────────────────────────────────────────────
def clean_filename(title: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "_", title).strip()

def extract_page_id_from_url(url: str) -> Optional[str]:
    m = re.search(r'/pages/(\d+)', url)
    if m: return m.group(1)
    m = re.search(r'pageId=(\d+)', url)
    if m: return m.group(1)
    m = re.search(r'pages/(\d{8,})', url)
    if m: return m.group(1)
    return None

def analyze_confluence_content(html_content: str) -> Dict:
    soup = BeautifulSoup(html_content, 'html.parser')
    text_content = []
    for i in range(1, 7):
        for h in soup.find_all(f'h{i}'):
            text_content.append({'type': f'heading_{i}', 'content': h.get_text(strip=True)})
    for table in soup.find_all('table'):
        headers = [th.get_text(strip=True) for th in table.find_all('th')]
        rows = [[td.get_text(strip=True) for td in tr.find_all('td')]
                for tr in table.find_all('tr') if tr.find_all('td')]
        if headers or rows:
            text_content.append({'type': 'table', 'headers': headers, 'rows': rows})
    for ul in soup.find_all(['ul', 'ol']):
        items = [li.get_text(strip=True) for li in ul.find_all('li', recursive=False)]
        if items: text_content.append({'type': 'list', 'items': items})
    for p in soup.find_all('p'):
        t = p.get_text(strip=True)
        if t and len(t) > 10:
            text_content.append({'type': 'paragraph', 'content': t})
    images = [{'src': img.get('src',''), 'alt': img.get('alt','')}
              for img in soup.find_all('img')]
    return {'text_content': text_content, 'images': images}

def convert_to_markdown(analysis: Dict, title: str) -> str:
    md = [f"# {title}", "",
          f"---", f"Generated at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", "---", ""]
    if analysis['images']:
        md += ["## 이미지", ""]
        for i, img in enumerate(analysis['images'], 1):
            md.append(f"{i}. **{img.get('alt','이미지')}**: {img.get('src','')}")
        md.append("")
    md += ["## 내용", ""]
    for item in analysis['text_content']:
        t = item['type']
        if t.startswith('heading_'):
            md += [f"{'#'*int(t[-1])} {item['content']}", ""]
        elif t == 'table':
            h = item.get('headers', [])
            if h:
                md.append("| " + " | ".join(h) + " |")
                md.append("| " + " | ".join(["---"]*len(h)) + " |")
                for row in item.get('rows', []):
                    while len(row) < len(h): row.append("")
                    md.append("| " + " | ".join(row) + " |")
                md.append("")
        elif t == 'list':
            for it in item.get('items', []): md.append(f"- {it}")
            md.append("")
        elif t == 'paragraph':
            md += [item['content'], ""]
    return "\n".join(md)

def get_page_children(page, parent_id: str) -> List[Dict]:
    all_pages, start = [], 0
    while True:
        resp = page.request.get(
            f"{BASE_URL}/rest/api/content/search",
            params={"cql": f"ancestor={parent_id} and type=page",
                    "start": str(start), "limit": "50", "expand": "version"}
        )
        if resp.status != 200: break
        data = resp.json()
        results = data.get("results", [])
        if not results: break
        for doc in results:
            all_pages.append({"id": doc["id"], "title": doc["title"]})
        if len(all_pages) >= data.get("size", 0): break
        start += 50
        if start > 500: break
    return all_pages

def get_page_content(page, page_id: str) -> Optional[Dict]:
    try:
        resp = page.request.get(
            f"{BASE_URL}/rest/api/content/{page_id}",
            params={"expand": "body.storage"}
        )
        if resp.status != 200: return None
        d = resp.json()
        html = d.get("body", {}).get("storage", {}).get("value", "")
        if not html: return None
        return {"id": d.get("id"), "title": d.get("title", ""), "html": html}
    except:
        return None


# ─────────────────────────────────────────────
# 핵심: 저장된 MD 파일들로 보고서 생성 (Confluence 재호출 없음)
# ─────────────────────────────────────────────
def generate_report_from_md_files(md_dir: str, use_llm: bool = True,
                                   callback=None) -> str:
    """저장된 MD 파일을 읽어서 보고서 생성 (Confluence 재접속 불필요)"""

    def log(msg):
        if callback: callback(msg)

    # 폴더 내 모든 MD 파일 재귀 탐색
    pages = []
    for root, _, files in os.walk(md_dir):
        for fname in sorted(files):
            if fname.endswith(".md") and "보고서" not in fname and "report" not in fname.lower():
                path = os.path.join(root, fname)
                try:
                    with open(path, encoding="utf-8") as f:
                        content = f.read()
                    title = fname.replace(".md", "")
                    pages.append({"title": title, "content": content, "path": path})
                    log(f"  읽기: {fname}")
                except Exception as e:
                    log(f"  [경고] 읽기 실패: {fname} → {e}")

    if not pages:
        return "보고서 생성 실패: MD 파일을 찾을 수 없습니다."

    log(f"\n총 {len(pages)}개 파일 로드 완료")

    # 헤더
    lines = [
        "# 종합 보고서",
        "",
        f"생성일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"참조 파일: {len(pages)}개",
        "",
        "---",
        "",
    ]

    # LLM 종합 보고서
    if use_llm and _llm_ready():
        log("LLM 종합 보고서 생성 중...")
        report = llm_generate_report(pages)
        lines += ["## 🤖 AI 종합 보고서 (팀장 보고용)", "", report, "", "---", ""]
    else:
        if not _llm_ready():
            lines += ["> ⚠️ LLM API 키가 설정되지 않아 AI 요약 없이 생성됩니다.", "", "---", ""]

    # 개별 페이지 요약 목록
    lines += ["## 📋 페이지별 내용", ""]
    for p in pages:
        lines.append(f"### {p['title']}")
        lines.append("")
        # 앞 부분 500자만 미리보기
        preview = p['content'][:500].strip()
        lines.append(preview)
        if len(p['content']) > 500:
            lines.append("...")
        lines.append("")
        lines.append("---")
        lines.append("")

    return "\n".join(lines)


# ─────────────────────────────────────────────
# Confluence 수집 (원본 유지)
# ─────────────────────────────────────────────
def process_child_pages_recursive(page, page_id: str, save_dir: str,
                                   depth: int = 2, use_llm: bool = True,
                                   callback=None):
    def log(msg):
        if callback: callback(msg)

    log(f"페이지 {page_id} 처리 중...")
    try:
        data = get_page_content(page, page_id)
        if not data:
            log(f"  [실패] {page_id}"); return

        analysis = analyze_confluence_content(data['html'])
        markdown  = convert_to_markdown(analysis, data['title'])

        # LLM 요약
        llm_block = ""
        if use_llm and _llm_ready():
            log(f"  LLM 요약: {data['title']}")
            summary   = llm_summarize_page(data['title'], markdown)
            llm_block = f"\n\n---\n## 🤖 AI 요약 (팀장 보고용)\n\n{summary}\n\n---\n"

        safe = clean_filename(data['title'])
        path = os.path.join(save_dir, f"{page_id}_{safe}.md")
        with open(path, "w", encoding="utf-8") as f:
            f.write(markdown + llm_block)
        log(f"  [저장] {data['title']}")

        children = get_page_children(page, page_id)
        if children and depth > 0:
            child_dir = os.path.join(save_dir, f"sub_{safe}")
            os.makedirs(child_dir, exist_ok=True)
            for child in children:
                process_child_pages_recursive(
                    page, child['id'], child_dir, depth-1, use_llm, callback)
    except Exception as e:
        if callback: callback(f"[오류] {e}\n{traceback.format_exc()}")


# ─────────────────────────────────────────────
# GUI
# ─────────────────────────────────────────────
class ConfluenceChildSummaryGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Confluence 차일드 페이지 요약 v4")
        self.root.geometry("820x680")

        self.url_var    = tk.StringVar()
        self.depth_var  = tk.IntVar(value=7)
        self.out_var    = tk.StringVar(value="./confluence_output")
        self.llm_var    = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="준비됨")

        # LLM 설정 변수 (저장된 값으로 초기화)
        self.llm_key_var    = tk.StringVar(value=LLM_API_KEY)
        self.llm_url_var    = tk.StringVar(value=LLM_BASE_URL)
        self.llm_model_var  = tk.StringVar(value=LLM_MODEL)
        self.llm_tokens_var = tk.StringVar(value=str(LLM_MAX_TOKENS))

        self._build()

    def _build(self):
        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True, padx=8, pady=8)

        tab1 = ttk.Frame(nb, padding=10)
        tab2 = ttk.Frame(nb, padding=10)
        tab3 = ttk.Frame(nb, padding=10)
        nb.add(tab1, text="📥 수집")
        nb.add(tab2, text="📊 보고서 생성")
        nb.add(tab3, text="⚙️ LLM 설정")

        self._build_crawl_tab(tab1)
        self._build_report_tab(tab2)
        self._build_settings_tab(tab3)

    # ── 수집 탭 ──────────────────────────────
    def _build_crawl_tab(self, f):
        ttk.Label(f, text="Confluence URL:").grid(row=0, column=0, sticky="w", pady=4)
        self.url_entry = ttk.Entry(f, textvariable=self.url_var, width=68)
        self.url_entry.grid(row=0, column=1, columnspan=2, sticky="ew", pady=4)
        self.url_entry.insert(0, "https://confluence.sec.samsung.net/pages/...")

        self.pid_lbl = ttk.Label(f, text="페이지 ID: ", foreground="gray")
        self.pid_lbl.grid(row=1, column=1, sticky="w")

        ttk.Label(f, text="재귀 깊이:").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Spinbox(f, from_=0, to=15, textvariable=self.depth_var, width=8).grid(row=2, column=1, sticky="w")

        ttk.Checkbutton(f, text="🤖 LLM 요약 생성 (팀장 보고용 / .env 설정 필요)",
                        variable=self.llm_var).grid(row=3, column=1, sticky="w", pady=3)

        ttk.Label(f, text="출력 폴더:").grid(row=4, column=0, sticky="w", pady=4)
        frm = ttk.Frame(f)
        frm.grid(row=4, column=1, columnspan=2, sticky="ew")
        ttk.Entry(frm, textvariable=self.out_var, width=50).pack(side="left")
        ttk.Button(frm, text="찾아보기", command=self._browse).pack(side="left", padx=5)

        self.run_btn = ttk.Button(f, text="▶ 수집 시작", command=self._run)
        self.run_btn.grid(row=5, column=0, pady=10)

        self.progress = ttk.Progressbar(f, mode='indeterminate', length=660)
        self.progress.grid(row=6, column=0, columnspan=3, sticky="ew")

        ttk.Label(f, text="로그:").grid(row=7, column=0, sticky="w")
        lf = ttk.Frame(f)
        lf.grid(row=8, column=0, columnspan=3, sticky="nsew")
        f.rowconfigure(8, weight=1); f.columnconfigure(1, weight=1)
        sb = ttk.Scrollbar(lf); sb.pack(side="right", fill="y")
        self.log_box = tk.Text(lf, height=14, yscrollcommand=sb.set)
        self.log_box.pack(fill="both", expand=True)
        sb.config(command=self.log_box.yview)

        ttk.Label(f, textvariable=self.status_var, foreground="blue").grid(
            row=9, column=0, columnspan=3, pady=4)

        self.url_var.trace_add("write", self._on_url)

    # ── 보고서 탭 ────────────────────────────
    def _build_report_tab(self, f):
        ttk.Label(f, text="수집된 MD 폴더:").grid(row=0, column=0, sticky="w", pady=6)
        self.rpt_dir_var = tk.StringVar(value="./confluence_output")
        frm = ttk.Frame(f)
        frm.grid(row=0, column=1, sticky="ew")
        ttk.Entry(frm, textvariable=self.rpt_dir_var, width=52).pack(side="left")
        ttk.Button(frm, text="폴더 선택",
                   command=lambda: self.rpt_dir_var.set(
                       filedialog.askdirectory() or self.rpt_dir_var.get())
                   ).pack(side="left", padx=4)

        self.rpt_llm_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(f, text="🤖 LLM 종합 보고서 생성",
                        variable=self.rpt_llm_var).grid(row=1, column=1, sticky="w", pady=3)

        self.rpt_btn = ttk.Button(f, text="📊 보고서 생성", command=self._gen_report)
        self.rpt_btn.grid(row=2, column=0, pady=10)

        self.rpt_prog = ttk.Progressbar(f, mode='indeterminate', length=660)
        self.rpt_prog.grid(row=3, column=0, columnspan=2, sticky="ew")

        ttk.Label(f, text="로그:").grid(row=4, column=0, sticky="w")
        lf = ttk.Frame(f)
        lf.grid(row=5, column=0, columnspan=2, sticky="nsew")
        f.rowconfigure(5, weight=1); f.columnconfigure(1, weight=1)
        sb = ttk.Scrollbar(lf); sb.pack(side="right", fill="y")
        self.rpt_log = tk.Text(lf, height=16, yscrollcommand=sb.set)
        self.rpt_log.pack(fill="both", expand=True)
        sb.config(command=self.rpt_log.yview)

    # ── LLM 설정 탭 ─────────────────────────
    def _build_settings_tab(self, f):
        f.columnconfigure(1, weight=1)

        ttk.Label(f, text="루코드 LLM 연결 설정", font=("", 11, "bold")).grid(
            row=0, column=0, columnspan=2, pady=(0, 12), sticky="w")

        ttk.Label(f, text="API Key:").grid(row=1, column=0, sticky="w", pady=6)
        key_frm = ttk.Frame(f)
        key_frm.grid(row=1, column=1, sticky="ew", pady=6)
        self.key_entry = ttk.Entry(key_frm, textvariable=self.llm_key_var, width=55, show="*")
        self.key_entry.pack(side="left", fill="x", expand=True)
        self._show_key = False
        def toggle_key():
            self._show_key = not self._show_key
            self.key_entry.config(show="" if self._show_key else "*")
            show_btn.config(text="숨기기" if self._show_key else "보기")
        show_btn = ttk.Button(key_frm, text="보기", width=5, command=toggle_key)
        show_btn.pack(side="left", padx=4)

        ttk.Label(f, text="Base URL:").grid(row=2, column=0, sticky="w", pady=6)
        url_entry = ttk.Entry(f, textvariable=self.llm_url_var, width=62)
        url_entry.grid(row=2, column=1, sticky="ew", pady=6)
        ttk.Label(f, text="  예) http://localhost:8000/v1  또는  https://루코드주소/v1",
                  foreground="gray").grid(row=3, column=1, sticky="w")

        ttk.Label(f, text="모델명:").grid(row=4, column=0, sticky="w", pady=6)
        ttk.Entry(f, textvariable=self.llm_model_var, width=30).grid(row=4, column=1, sticky="w", pady=6)
        ttk.Label(f, text="  예) qwen-plus  /  qwen2.5-72b-instruct",
                  foreground="gray").grid(row=5, column=1, sticky="w")

        ttk.Label(f, text="최대 토큰:").grid(row=6, column=0, sticky="w", pady=6)
        ttk.Entry(f, textvariable=self.llm_tokens_var, width=10).grid(row=6, column=1, sticky="w", pady=6)

        btn_frm = ttk.Frame(f)
        btn_frm.grid(row=7, column=0, columnspan=2, pady=14, sticky="w")
        ttk.Button(btn_frm, text="💾 설정 저장", command=self._save_settings).pack(side="left", padx=5)
        ttk.Button(btn_frm, text="🔗 연결 테스트", command=self._test_llm).pack(side="left", padx=5)

        self.settings_status = ttk.Label(f, text="", foreground="gray")
        self.settings_status.grid(row=8, column=0, columnspan=2, sticky="w", pady=4)

        ttk.Separator(f, orient="horizontal").grid(row=9, column=0, columnspan=2, sticky="ew", pady=10)
        info = ("📌 설정 방법\n"
                "1. 루코드 서버 주소를 Base URL에 입력하세요.\n"
                "2. API Key는 루코드에서 발급받은 키를 입력하세요.\n"
                "3. 모델명은 루코드 서버에서 제공하는 Qwen 모델명으로 입력하세요.\n"
                "4. '설정 저장' 버튼을 누르면 다음에 실행해도 자동으로 불러옵니다.")
        ttk.Label(f, text=info, foreground="#555", justify="left").grid(
            row=10, column=0, columnspan=2, sticky="w")

    def _save_settings(self):
        key    = self.llm_key_var.get().strip()
        url    = self.llm_url_var.get().strip()
        model  = self.llm_model_var.get().strip()
        tokens = self.llm_tokens_var.get().strip()
        if not key or not url:
            self.settings_status.config(text="⚠️ API Key와 Base URL은 필수입니다.", foreground="red")
            return
        try:
            save_llm_config(key, url, model, int(tokens))
            self.settings_status.config(
                text=f"✅ 저장 완료 → {CONFIG_FILE}", foreground="green")
        except Exception as e:
            self.settings_status.config(text=f"❌ 저장 실패: {e}", foreground="red")

    def _test_llm(self):
        self._save_settings()
        self.settings_status.config(text="🔄 연결 테스트 중...", foreground="blue")
        self.root.update_idletasks()
        def _test():
            try:
                client = OpenAI(api_key=LLM_API_KEY, base_url=LLM_BASE_URL)
                resp = client.chat.completions.create(
                    model=LLM_MODEL,
                    messages=[{"role": "user", "content": "안녕하세요. 연결 테스트입니다. 한 문장으로 답해주세요."}],
                    max_tokens=100, temperature=0.1,
                )
                answer = resp.choices[0].message.content.strip()[:80]
                self.root.after(0, lambda: self.settings_status.config(
                    text=f"✅ 연결 성공! 응답: {answer}", foreground="green"))
            except Exception as e:
                self.root.after(0, lambda: self.settings_status.config(
                    text=f"❌ 연결 실패: {e}", foreground="red"))
        threading.Thread(target=_test, daemon=True).start()

    def _on_url(self, *_):
        pid = extract_page_id_from_url(self.url_var.get())
        if pid:
            self.pid_lbl.config(text=f"페이지 ID: {pid}", foreground="green")
        else:
            self.pid_lbl.config(text="페이지 ID: (찾을 수 없음)", foreground="red")

    def _browse(self):
        d = filedialog.askdirectory()
        if d: self.out_var.set(d)

    def log(self, msg):
        try:
            self.log_box.insert("end", msg + "\n")
            self.log_box.see("end")
            self.root.update_idletasks()
        except: pass

    def rlog(self, msg):
        try:
            self.rpt_log.insert("end", msg + "\n")
            self.rpt_log.see("end")
            self.root.update_idletasks()
        except: pass

    # ── 수집 실행 ────────────────────────────
    def _run(self):
        pid = extract_page_id_from_url(self.url_var.get())
        if not pid:
            messagebox.showerror("오류", "URL에서 페이지 ID를 찾을 수 없습니다.")
            return
        self.run_btn.config(state="disabled")
        self.progress.start()
        self.status_var.set("실행 중...")
        threading.Thread(target=self._crawl_worker, args=(pid,), daemon=True).start()

    def _crawl_worker(self, page_id):
        from playwright.sync_api import sync_playwright
        save_dir = os.path.join(self.out_var.get(), f"page_{page_id}")
        os.makedirs(save_dir, exist_ok=True)
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch_persistent_context(
                    user_data_dir=USER_DATA_DIR, headless=False,
                    viewport={"width": 1280, "height": 720},
                )
                page = browser.new_page()
                page.goto(f"{BASE_URL}/pages/viewpage.action?pageId={page_id}")
                try: page.wait_for_load_state("networkidle", timeout=30000)
                except: pass

                if not messagebox.askokcancel("로그인 확인",
                        "Confluence에 로그인되어 있나요?\n[확인] 진행  [취소] 중단"):
                    browser.close(); return

                page.reload(wait_until="networkidle")
                self.log("=" * 55)
                self.log(f"수집 시작 | 페이지 ID: {page_id}")
                self.log(f"LLM 요약: {'켜짐' if self.llm_var.get() else '꺼짐'}")
                self.log("=" * 55)

                process_child_pages_recursive(
                    page, page_id, save_dir,
                    depth=self.depth_var.get(),
                    use_llm=self.llm_var.get(),
                    callback=self.log,
                )
                browser.close()

                # ★ 보고서는 저장된 MD 파일에서 생성 (Confluence 재호출 없음)
                self.log("\n보고서 생성 중 (저장된 MD 파일 기반)...")
                summary = generate_report_from_md_files(
                    save_dir,
                    use_llm=self.llm_var.get(),
                    callback=self.log,
                )
                rpt = os.path.join(save_dir, f"{page_id}_종합보고서.md")
                with open(rpt, "w", encoding="utf-8") as f:
                    f.write(summary)
                self.log(f"보고서 저장: {rpt}")
                self.log("\n✅ 완료!")
                self.status_var.set("완료")
                messagebox.showinfo("완료", f"완료!\n{os.path.abspath(save_dir)}")

        except Exception as e:
            self.log(f"[오류] {e}\n{traceback.format_exc()}")
            self.status_var.set("오류")
            messagebox.showerror("오류", str(e))
        finally:
            self.root.after(0, lambda: (
                self.progress.stop(), self.run_btn.config(state="normal")))

    # ── 보고서 탭 실행 ───────────────────────
    def _gen_report(self):
        d = self.rpt_dir_var.get()
        if not os.path.isdir(d):
            messagebox.showerror("오류", "폴더가 존재하지 않습니다.")
            return
        self.rpt_btn.config(state="disabled")
        self.rpt_prog.start()
        threading.Thread(target=self._report_worker, args=(d,), daemon=True).start()

    def _report_worker(self, md_dir):
        try:
            self.rlog("=" * 55)
            self.rlog(f"보고서 생성 시작: {md_dir}")
            self.rlog("=" * 55)
            summary = generate_report_from_md_files(
                md_dir, use_llm=self.rpt_llm_var.get(), callback=self.rlog)
            ts  = datetime.now().strftime("%Y%m%d_%H%M")
            rpt = os.path.join(md_dir, f"종합보고서_{ts}.md")
            with open(rpt, "w", encoding="utf-8") as f:
                f.write(summary)
            self.rlog(f"\n✅ 보고서 저장: {rpt}")
            messagebox.showinfo("완료", f"보고서 생성 완료!\n{rpt}")
        except Exception as e:
            self.rlog(f"[오류] {e}\n{traceback.format_exc()}")
            messagebox.showerror("오류", str(e))
        finally:
            self.root.after(0, lambda: (
                self.rpt_prog.stop(), self.rpt_btn.config(state="normal")))


def main():
    root = tk.Tk()
    ConfluenceChildSummaryGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
