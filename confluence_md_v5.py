"""
Confluence 차일드 페이지 요약 GUI - v5
디자인 개선: 모던 Samsung 스타일 GUI + 색상 로그 + LLM 설정 패널
"""

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

import os, re, threading, traceback, json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, List, Optional
from datetime import datetime
from bs4 import BeautifulSoup

# ─── 색상 팔레트 (Samsung 스타일) ──────────────
BLUE       = "#1428A0"
BLUE_DK    = "#0D1E7A"
BLUE_LT    = "#E8EAFF"
CYAN       = "#00B0FF"
BG         = "#F0F2F7"
CARD       = "#FFFFFF"
BORDER     = "#D0D5E0"
TEXT       = "#1A1A2E"
TEXT_MUTED = "#6B7280"
GREEN      = "#16A34A"
RED        = "#DC2626"
ORANGE     = "#EA580C"
LOG_BG     = "#0D1117"
LOG_FG     = "#C9D1D9"
LOG_GREEN  = "#3FB950"
LOG_RED    = "#FF7B72"
LOG_BLUE   = "#79C0FF"
LOG_YELLOW = "#E3B341"

# ─── 설정 ──────────────────────────────────────
BASE_URL      = "https://confluence.sec.samsung.net"
USER_DATA_DIR = "./chrome_profile_confluence_md"
CONFIG_FILE   = os.path.join(os.path.dirname(os.path.abspath(__file__)), "llm_config.json")

LLM_API_KEY    = ""
LLM_BASE_URL   = ""
LLM_MODEL      = "qwen-plus"
LLM_MAX_TOKENS = 2000

def load_llm_config():
    global LLM_API_KEY, LLM_BASE_URL, LLM_MODEL, LLM_MAX_TOKENS
    if os.path.exists(CONFIG_FILE):
        try:
            cfg = json.loads(open(CONFIG_FILE, encoding="utf-8").read())
            LLM_API_KEY    = cfg.get("api_key",    "")
            LLM_BASE_URL   = cfg.get("base_url",   "")
            LLM_MODEL      = cfg.get("model",      "qwen-plus")
            LLM_MAX_TOKENS = int(cfg.get("max_tokens", 2000))
        except: pass

def save_llm_config(api_key, base_url, model, max_tokens):
    global LLM_API_KEY, LLM_BASE_URL, LLM_MODEL, LLM_MAX_TOKENS
    LLM_API_KEY = api_key; LLM_BASE_URL = base_url
    LLM_MODEL = model; LLM_MAX_TOKENS = int(max_tokens)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump({"api_key": api_key, "base_url": base_url,
                   "model": model, "max_tokens": max_tokens}, f, ensure_ascii=False, indent=2)

load_llm_config()

def _llm_ready():
    return bool(LLM_API_KEY and LLM_BASE_URL)

# ─── LLM 함수 ──────────────────────────────────
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
                    "## 개요\n이 업무가 무엇인지 2~3문장으로 쉽게 설명\n\n"
                    "## 이번 주요 내용\n문서에 기록된 사실만, 불릿 포인트로\n\n"
                    "## 완료된 사항\n없으면 생략\n\n"
                    "## 진행 중인 사항\n없으면 생략\n\n"
                    "## 이슈 및 특이사항\n없으면 생략"},
                {"role": "user", "content":
                    f"페이지 제목: {title}\n\n=== 문서 내용 ===\n{text[:5000]}"}
            ],
            max_tokens=LLM_MAX_TOKENS, temperature=0.1,
        )
        return resp.choices[0].message.content
    except Exception as e:
        return f"[LLM 요약 실패: {e}]"

def llm_generate_report(pages: list) -> str:
    combined = "".join(f"\n\n### {p['title']}\n{p['content'][:2000]}" for p in pages)
    try:
        client = OpenAI(api_key=LLM_API_KEY, base_url=LLM_BASE_URL)
        resp = client.chat.completions.create(
            model=LLM_MODEL,
            messages=[
                {"role": "system", "content":
                    "당신은 400명 규모 조직의 팀장을 위한 종합 보고서 작성 전문가입니다.\n\n"
                    "【절대 규칙】\n- 제공된 내용만 사용. 없는 내용은 절대 만들지 마세요.\n"
                    "- 추측이나 창작 금지. 불확실하면 '문서에 명시되지 않음'으로 표기.\n\n"
                    "【작성 방식】\n- 각 팀/담당자가 무슨 일을 하는지 처음 보는 사람도 이해하도록 설명\n"
                    "- 전문 용어는 괄호로 쉬운 말 추가\n"
                    "- 격식 있는 문어체 (~하였습니다, ~검토 중에 있습니다)\n"
                    "- 수치, 날짜, 완료율 등 구체적 수치 반드시 포함\n\n"
                    "【보고서 구성】\n# 1. 전체 요약\n처음 보는 사람도 이해할 수 있게 3~5문장\n\n"
                    "# 2. 항목별 상세 현황\n각 팀별 업무 내용과 현재 상태\n\n"
                    "# 3. 완료된 주요 사항\n없으면 생략\n\n"
                    "# 4. 진행 중인 주요 과제\n없으면 생략\n\n"
                    "# 5. 이슈 및 리스크\n없으면 생략\n\n"
                    "# 6. 팀장 조치 필요 사항\n없으면 생략"},
                {"role": "user", "content":
                    f"{len(pages)}개 페이지 내용으로 팀장 보고용 종합 보고서를 작성해주세요.\n"
                    f"제공된 내용만 사용하세요.\n\n=== 내용 ===\n{combined}"}
            ],
            max_tokens=LLM_MAX_TOKENS * 2, temperature=0.1,
        )
        return resp.choices[0].message.content
    except Exception as e:
        return f"[보고서 생성 실패: {e}]"

# ─── 유틸 ──────────────────────────────────────
def clean_filename(title: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "_", title).strip()

def extract_page_id_from_url(url: str) -> Optional[str]:
    for pat in [r'/pages/(\d+)', r'pageId=(\d+)', r'pages/(\d{8,})']:
        m = re.search(pat, url)
        if m: return m.group(1)
    return None

def analyze_confluence_content(html_content: str) -> Dict:
    soup = BeautifulSoup(html_content, 'html.parser')
    items = []
    for i in range(1, 7):
        for h in soup.find_all(f'h{i}'):
            items.append({'type': f'heading_{i}', 'content': h.get_text(strip=True)})
    for table in soup.find_all('table'):
        headers = [th.get_text(strip=True) for th in table.find_all('th')]
        rows = [[td.get_text(strip=True) for td in tr.find_all('td')]
                for tr in table.find_all('tr') if tr.find_all('td')]
        if headers or rows:
            items.append({'type': 'table', 'headers': headers, 'rows': rows})
    for ul in soup.find_all(['ul', 'ol']):
        li = [l.get_text(strip=True) for l in ul.find_all('li', recursive=False)]
        if li: items.append({'type': 'list', 'items': li})
    for p in soup.find_all('p'):
        t = p.get_text(strip=True)
        if t and len(t) > 10:
            items.append({'type': 'paragraph', 'content': t})
    images = [{'src': img.get('src',''), 'alt': img.get('alt','')}
              for img in soup.find_all('img')]
    return {'text_content': items, 'images': images}

def convert_to_markdown(analysis: Dict, title: str) -> str:
    md = [f"# {title}", "", "---",
          f"Generated at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", "---", ""]
    if analysis['images']:
        md += ["## 이미지", ""]
        for i, img in enumerate(analysis['images'], 1):
            md.append(f"{i}. {img.get('alt','이미지')}: {img.get('src','')}")
        md.append("")
    md += ["## 내용", ""]
    for item in analysis['text_content']:
        t = item['type']
        if t.startswith('heading_'):
            md += [f"{'#'*int(t[-1])} {item['content']}", ""]
        elif t == 'table':
            h = item.get('headers', [])
            if h:
                md += ["| " + " | ".join(h) + " |",
                       "| " + " | ".join(["---"]*len(h)) + " |"]
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
                    "start": str(start), "limit": "50", "expand": "version"})
        if resp.status != 200: break
        data = resp.json(); results = data.get("results", [])
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
            params={"expand": "body.storage"})
        if resp.status != 200: return None
        d = resp.json()
        html = d.get("body", {}).get("storage", {}).get("value", "")
        if not html: return None
        return {"id": d.get("id"), "title": d.get("title"), "html": html}
    except: return None

def process_child_pages_recursive(page, page_id: str, save_dir: str,
                                   depth: int = 2, use_llm: bool = True,
                                   callback=None):
    if callback: callback(f"INFO: 페이지 {page_id} 처리 중...")
    try:
        data = get_page_content(page, page_id)
        if not data:
            if callback: callback(f"WARN: 페이지 {page_id} 내용 없음")
            return
        analysis = analyze_confluence_content(data['html'])
        markdown  = convert_to_markdown(analysis, data['title'])
        llm_block = ""
        if use_llm and _llm_ready():
            if callback: callback(f"INFO: LLM 요약 중: {data['title']}")
            summary   = llm_summarize_page(data['title'], markdown)
            llm_block = f"\n\n---\n## AI 요약 (팀장 보고용)\n\n{summary}\n\n---\n"
        safe = clean_filename(data['title'])
        path = os.path.join(save_dir, f"{page_id}_{safe}.md")
        with open(path, "w", encoding="utf-8") as f:
            f.write(markdown + llm_block)
        if callback: callback(f"OK: {data['title']}")
        children = get_page_children(page, page_id)
        if children and depth > 0:
            child_dir = os.path.join(save_dir, f"children_{page_id}")
            os.makedirs(child_dir, exist_ok=True)
            for child in children:
                process_child_pages_recursive(page, child['id'], child_dir,
                                               depth-1, use_llm, callback)
    except Exception as e:
        if callback: callback(f"ERR: {e}\n{traceback.format_exc()}")

def generate_summary_report(page, page_id: str, save_dir: str,
                             max_depth: int = 2, use_llm: bool = True,
                             callback=None) -> str:
    lines = ["# 차일드 페이지 요약 보고서", "",
             f"생성일: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
             f"최대 깊이: {max_depth}", "", "="*80, ""]
    page_data = get_page_content(page, page_id)
    if page_data:
        lines.append(f"선택 페이지: {page_data['title']} (ID: {page_id})")
        lines.append("")

    def collect(pid, depth=0):
        result = []
        for child in get_page_children(page, pid):
            cd = get_page_content(page, child['id'])
            if cd:
                analysis = analyze_confluence_content(cd['html'])
                md_text  = convert_to_markdown(analysis, cd['title'])
                result.append({"id": child['id'], "title": cd['title'],
                                "depth": depth+1, "analysis": analysis, "md_text": md_text})
                if depth+1 < max_depth:
                    result.extend(collect(child['id'], depth+1))
        return result

    all_pages = collect(page_id)
    if not all_pages:
        lines.append("차일드 페이지가 없습니다.")
    else:
        if use_llm and _llm_ready():
            if callback: callback("INFO: LLM 종합 보고서 생성 중...")
            report = llm_generate_report([
                {"title": p["title"], "content": p["md_text"]} for p in all_pages])
            lines += ["## AI 종합 보고서 (팀장 보고용)", "", report, "", "="*80, ""]
        lines += ["## 페이지 목록", ""]
        for p in all_pages:
            indent = "  " * (p['depth']-1)
            lines.append(f"{indent}  {p['title']} (ID: {p['id']})")
    return "\n".join(lines)


# ─── 커스텀 버튼 위젯 ──────────────────────────
class FlatButton(tk.Canvas):
    def __init__(self, parent, text, command=None, bg=BLUE, fg="white",
                 hover=BLUE_DK, width=140, height=36, font_size=10, **kw):
        super().__init__(parent, width=width, height=height,
                         highlightthickness=0, bd=0, cursor="hand2", **kw)
        self._bg = bg; self._hover = hover; self._fg = fg
        self._text = text; self._cmd = command
        self._font = ("Malgun Gothic", font_size, "bold")
        self._draw(bg)
        self.bind("<Enter>",    lambda e: self._draw(hover))
        self.bind("<Leave>",    lambda e: self._draw(bg))
        self.bind("<Button-1>", lambda e: self._click())
        self.bind("<ButtonRelease-1>", lambda e: self._draw(hover if self.winfo_containing(
            e.x_root, e.y_root) == self else bg))

    def _draw(self, color):
        self.delete("all")
        w, h = int(self["width"]), int(self["height"])
        r = 6
        self.create_arc(0, 0, r*2, r*2, start=90, extent=90, fill=color, outline=color)
        self.create_arc(w-r*2, 0, w, r*2, start=0, extent=90, fill=color, outline=color)
        self.create_arc(0, h-r*2, r*2, h, start=180, extent=90, fill=color, outline=color)
        self.create_arc(w-r*2, h-r*2, w, h, start=270, extent=90, fill=color, outline=color)
        self.create_rectangle(r, 0, w-r, h, fill=color, outline=color)
        self.create_rectangle(0, r, w, h-r, fill=color, outline=color)
        self.create_text(w//2, h//2, text=self._text, fill=self._fg, font=self._font)

    def _click(self):
        self._draw(self._bg)
        if self._cmd: self._cmd()

    def config_state(self, state):
        if state == "disabled":
            self._bg_orig = self._bg; self._bg = "#9CA3AF"
            self._draw("#9CA3AF"); self.unbind("<Button-1>")
        else:
            self._bg = getattr(self, '_bg_orig', BLUE)
            self._draw(self._bg)
            self.bind("<Button-1>", lambda e: self._click())


# ─── 메인 GUI ──────────────────────────────────
class ConfluenceGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Confluence 수집기  v5")
        self.root.geometry("1020x740")
        self.root.configure(bg=BG)
        self.root.resizable(True, True)

        self.url_var       = tk.StringVar()
        self.depth_var     = tk.IntVar(value=7)
        self.out_var       = tk.StringVar(value="./confluence_output")
        self.llm_var       = tk.BooleanVar(value=True)
        self.llm_key_var   = tk.StringVar(value=LLM_API_KEY)
        self.llm_url_var   = tk.StringVar(value=LLM_BASE_URL)
        self.llm_model_var = tk.StringVar(value=LLM_MODEL)
        self.llm_tokens_var= tk.StringVar(value=str(LLM_MAX_TOKENS))
        self._llm_panel_open = False
        self._running = False

        self._build()

    # ─── 빌드 ────────────────────────────────────
    def _build(self):
        self._build_header()
        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=14, pady=(10,0))
        body.columnconfigure(0, weight=0, minsize=340)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)
        self._build_settings(body)
        self._build_log(body)
        self._build_statusbar()

    def _build_header(self):
        hdr = tk.Frame(self.root, bg=BLUE, height=56)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        tk.Label(hdr, text="  Confluence 수집 & 보고서 생성",
                 bg=BLUE, fg="white",
                 font=("Malgun Gothic", 14, "bold")).pack(side="left", padx=16)

        badge = tk.Label(hdr, text=" v5 ", bg=CYAN, fg=BLUE_DK,
                         font=("Malgun Gothic", 9, "bold"), relief="flat", padx=4)
        badge.pack(side="left", pady=16)

        self._status_dot = tk.Label(hdr, text="●  준비됨", bg=BLUE, fg="#86EFAC",
                                    font=("Malgun Gothic", 9))
        self._status_dot.pack(side="right", padx=20)

    def _build_settings(self, parent):
        outer = tk.Frame(parent, bg=CARD, bd=0,
                         highlightthickness=1, highlightbackground=BORDER)
        outer.grid(row=0, column=0, sticky="nsew", padx=(0,10), pady=(0,10))

        tk.Label(outer, text="⚙  설정", bg=CARD, fg=BLUE,
                 font=("Malgun Gothic", 10, "bold")).pack(anchor="w", padx=16, pady=(12,0))
        tk.Frame(outer, bg=BLUE_LT, height=1).pack(fill="x", padx=16, pady=4)

        form = tk.Frame(outer, bg=CARD)
        form.pack(fill="x", padx=16)

        # URL
        self._lbl(form, "Confluence URL")
        url_frm = tk.Frame(form, bg=CARD)
        url_frm.pack(fill="x", pady=(0,6))
        self.url_entry = self._entry(url_frm, self.url_var, width=30)
        self.url_entry.pack(side="left", fill="x", expand=True)
        self.url_var.trace_add("write", self._on_url)

        self.pid_lbl = tk.Label(form, text="페이지 ID: —", bg=CARD,
                                fg=TEXT_MUTED, font=("Malgun Gothic", 8))
        self.pid_lbl.pack(anchor="w", pady=(0,6))

        # 깊이
        self._lbl(form, "재귀 깊이")
        depth_frm = tk.Frame(form, bg=CARD)
        depth_frm.pack(fill="x", pady=(0,6))
        sb = tk.Spinbox(depth_frm, from_=0, to=15, textvariable=self.depth_var,
                        width=5, font=("Malgun Gothic", 10),
                        bd=1, relief="solid")
        sb.pack(side="left")

        # 출력 폴더
        self._lbl(form, "출력 폴더")
        out_frm = tk.Frame(form, bg=CARD)
        out_frm.pack(fill="x", pady=(0,6))
        self._entry(out_frm, self.out_var, width=22).pack(side="left", fill="x", expand=True)
        tk.Button(out_frm, text="찾기", command=self._browse,
                  bg=BG, fg=TEXT, relief="flat", cursor="hand2",
                  font=("Malgun Gothic", 9), padx=6).pack(side="left", padx=(4,0))

        # LLM 체크
        chk_frm = tk.Frame(form, bg=CARD)
        chk_frm.pack(fill="x", pady=(4,0))
        tk.Checkbutton(chk_frm, text="LLM 요약 생성 (팀장 보고용)",
                       variable=self.llm_var, bg=CARD, fg=TEXT,
                       activebackground=CARD, font=("Malgun Gothic", 9),
                       selectcolor=BLUE_LT).pack(side="left")

        # LLM 설정 토글 버튼
        tk.Frame(form, bg=BORDER, height=1).pack(fill="x", pady=8)
        self._llm_toggle_btn = tk.Label(form,
            text="▶  LLM 연결 설정  (클릭하여 펼치기)",
            bg=BLUE_LT, fg=BLUE, font=("Malgun Gothic", 9, "bold"),
            cursor="hand2", padx=8, pady=4, anchor="w")
        self._llm_toggle_btn.pack(fill="x")
        self._llm_toggle_btn.bind("<Button-1>", lambda e: self._toggle_llm_panel())

        # LLM 패널 (접힘)
        self._llm_panel = tk.Frame(form, bg="#F5F7FF",
                                   highlightthickness=1, highlightbackground=BLUE_LT)
        self._build_llm_panel(self._llm_panel)

        # 진행바
        tk.Frame(outer, bg=BORDER, height=1).pack(fill="x", padx=16, pady=6)
        prog_frm = tk.Frame(outer, bg=CARD)
        prog_frm.pack(fill="x", padx=16, pady=(0,6))
        self.progress = ttk.Progressbar(prog_frm, mode="indeterminate", length=280)
        self.progress.pack(fill="x")

        # 실행 버튼
        btn_frm = tk.Frame(outer, bg=CARD)
        btn_frm.pack(pady=(4,16), padx=16, fill="x")
        self.run_btn = FlatButton(btn_frm, text="▶  수집 시작",
                                  command=self._run, bg=BLUE, fg="white",
                                  width=160, height=38, font_size=10)
        self.run_btn.pack(side="left")
        self.stop_lbl = tk.Label(btn_frm, text="", bg=CARD, fg=RED,
                                 font=("Malgun Gothic", 8))
        self.stop_lbl.pack(side="left", padx=10)

    def _build_llm_panel(self, f):
        def row(label, var, show=""):
            tk.Label(f, text=label, bg="#F5F7FF", fg=TEXT_MUTED,
                     font=("Malgun Gothic", 8)).pack(anchor="w", padx=8, pady=(4,0))
            entry = tk.Entry(f, textvariable=var, show=show,
                             font=("Malgun Gothic", 9), bd=1, relief="solid", bg="white")
            entry.pack(fill="x", padx=8, pady=(0,2))
            return entry
        row("API Key", self.llm_key_var, show="*")
        row("Base URL  (예: http://루코드주소/v1)", self.llm_url_var)
        row("모델명  (예: qwen-plus)", self.llm_model_var)
        row("최대 토큰", self.llm_tokens_var)
        btn_row = tk.Frame(f, bg="#F5F7FF")
        btn_row.pack(fill="x", padx=8, pady=6)
        FlatButton(btn_row, text="💾 저장", command=self._save_llm,
                   bg=GREEN, fg="white", width=80, height=30, font_size=9).pack(side="left")
        FlatButton(btn_row, text="🔗 연결 테스트", command=self._test_llm,
                   bg=BLUE, fg="white", width=110, height=30, font_size=9).pack(side="left", padx=6)
        self._llm_status = tk.Label(f, text="", bg="#F5F7FF",
                                    font=("Malgun Gothic", 8), wraplength=280, justify="left")
        self._llm_status.pack(anchor="w", padx=8, pady=(0,6))

    def _build_log(self, parent):
        log_outer = tk.Frame(parent, bg=LOG_BG,
                             highlightthickness=1, highlightbackground="#30363D")
        log_outer.grid(row=0, column=1, sticky="nsew", pady=(0,10))

        hdr = tk.Frame(log_outer, bg="#161B22")
        hdr.pack(fill="x")
        tk.Label(hdr, text="  로그", bg="#161B22", fg=LOG_BLUE,
                 font=("Consolas", 9, "bold")).pack(side="left", pady=6, padx=8)
        tk.Button(hdr, text="지우기", command=self._clear_log,
                  bg="#161B22", fg="#6E7681", relief="flat", cursor="hand2",
                  font=("Consolas", 8), padx=6).pack(side="right", padx=8)

        txt_frm = tk.Frame(log_outer, bg=LOG_BG)
        txt_frm.pack(fill="both", expand=True, padx=4, pady=4)
        sb = tk.Scrollbar(txt_frm, bg=LOG_BG, troughcolor=LOG_BG)
        sb.pack(side="right", fill="y")
        self.log_box = tk.Text(txt_frm, bg=LOG_BG, fg=LOG_FG,
                               font=("Consolas", 9), wrap="word",
                               yscrollcommand=sb.set,
                               insertbackground=LOG_FG, bd=0, relief="flat",
                               selectbackground="#264F78")
        self.log_box.pack(fill="both", expand=True)
        sb.config(command=self.log_box.yview)

        # 색상 태그
        self.log_box.tag_config("ok",   foreground=LOG_GREEN)
        self.log_box.tag_config("err",  foreground=LOG_RED)
        self.log_box.tag_config("info", foreground=LOG_BLUE)
        self.log_box.tag_config("warn", foreground=LOG_YELLOW)
        self.log_box.tag_config("ts",   foreground="#484F58")
        self.log_box.tag_config("done", foreground=LOG_GREEN, font=("Consolas", 9, "bold"))

    def _build_statusbar(self):
        bar = tk.Frame(self.root, bg="#E5E7EB", height=24)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        self._status_lbl = tk.Label(bar, text="  준비됨", bg="#E5E7EB", fg=TEXT_MUTED,
                                    font=("Malgun Gothic", 8), anchor="w")
        self._status_lbl.pack(side="left", fill="both", expand=True)
        tk.Label(bar, text=f"  Confluence Collector v5  ",
                 bg="#E5E7EB", fg=TEXT_MUTED, font=("Malgun Gothic", 8)).pack(side="right")

    # ─── 헬퍼 ────────────────────────────────────
    def _lbl(self, parent, text):
        tk.Label(parent, text=text, bg=CARD, fg=TEXT_MUTED,
                 font=("Malgun Gothic", 8)).pack(anchor="w", pady=(4,1))

    def _entry(self, parent, var, width=30):
        e = tk.Entry(parent, textvariable=var, width=width,
                     font=("Malgun Gothic", 9), bd=1, relief="solid",
                     bg="white", fg=TEXT)
        return e

    def _toggle_llm_panel(self):
        self._llm_panel_open = not self._llm_panel_open
        if self._llm_panel_open:
            self._llm_panel.pack(fill="x", pady=(2,0))
            self._llm_toggle_btn.config(text="▼  LLM 연결 설정  (클릭하여 접기)")
        else:
            self._llm_panel.pack_forget()
            self._llm_toggle_btn.config(text="▶  LLM 연결 설정  (클릭하여 펼치기)")

    def _on_url(self, *_):
        pid = extract_page_id_from_url(self.url_var.get())
        if pid:
            self.pid_lbl.config(text=f"페이지 ID: {pid}", fg=GREEN)
        else:
            self.pid_lbl.config(text="페이지 ID: 인식되지 않음", fg=RED)

    def _browse(self):
        d = filedialog.askdirectory()
        if d: self.out_var.set(d)

    def _clear_log(self):
        self.log_box.delete("1.0", "end")

    def _save_llm(self):
        k = self.llm_key_var.get().strip()
        u = self.llm_url_var.get().strip()
        m = self.llm_model_var.get().strip()
        t = self.llm_tokens_var.get().strip()
        if not k or not u:
            self._llm_status.config(text="API Key와 Base URL은 필수입니다.", fg=RED)
            return
        try:
            save_llm_config(k, u, m, int(t))
            self._llm_status.config(text="저장 완료!", fg=GREEN)
        except Exception as e:
            self._llm_status.config(text=f"저장 실패: {e}", fg=RED)

    def _test_llm(self):
        self._save_llm()
        self._llm_status.config(text="연결 테스트 중...", fg=LOG_BLUE)
        def _test():
            try:
                client = OpenAI(api_key=LLM_API_KEY, base_url=LLM_BASE_URL)
                resp = client.chat.completions.create(
                    model=LLM_MODEL,
                    messages=[{"role": "user", "content": "안녕하세요. 한 문장으로 답해주세요."}],
                    max_tokens=50, temperature=0.1)
                ans = resp.choices[0].message.content.strip()[:60]
                self.root.after(0, lambda: self._llm_status.config(
                    text=f"연결 성공!  응답: {ans}", fg=GREEN))
            except Exception as e:
                err = str(e)
                self.root.after(0, lambda: self._llm_status.config(
                    text=f"연결 실패: {err}", fg=RED))
        threading.Thread(target=_test, daemon=True).start()

    # ─── 로그 출력 ───────────────────────────────
    def log(self, msg: str):
        try:
            ts = datetime.now().strftime("%H:%M:%S")
            msg = msg.strip()
            if not msg: return
            self.log_box.insert("end", f"[{ts}] ", "ts")
            if msg.startswith("OK:"):
                self.log_box.insert("end", "✓ " + msg[3:].strip() + "\n", "ok")
            elif msg.startswith("ERR:"):
                self.log_box.insert("end", "✗ " + msg[4:].strip() + "\n", "err")
            elif msg.startswith("WARN:"):
                self.log_box.insert("end", "⚠ " + msg[5:].strip() + "\n", "warn")
            elif msg.startswith("DONE:"):
                self.log_box.insert("end", "★ " + msg[5:].strip() + "\n", "done")
            else:
                txt = msg[5:].strip() if msg.startswith("INFO:") else msg
                self.log_box.insert("end", txt + "\n", "info")
            self.log_box.see("end")
            self.root.update_idletasks()
        except: pass

    def _set_status(self, text, color=TEXT_MUTED, dot_color=None):
        self._status_lbl.config(text=f"  {text}", fg=color)
        self._status_dot.config(text=f"●  {text}",
                                fg=dot_color or ("#86EFAC" if color==GREEN else
                                                  "#FCA5A5" if color==RED else "#93C5FD"))

    # ─── 실행 ────────────────────────────────────
    def _run(self):
        pid = extract_page_id_from_url(self.url_var.get())
        if not pid:
            messagebox.showerror("오류", "URL에서 페이지 ID를 찾을 수 없습니다.")
            return
        if self.llm_var.get() and not _llm_ready():
            if not messagebox.askyesno("LLM 미설정",
                    "LLM 설정이 없습니다. LLM 없이 MD만 수집할까요?"):
                return
            self.llm_var.set(False)
        self.run_btn.config_state("disabled")
        self.progress.start(12)
        self._set_status("수집 중...", BLUE)
        self._running = True
        self._clear_log()
        threading.Thread(target=self._worker, args=(pid,), daemon=True).start()

    def _worker(self, page_id):
        from playwright.sync_api import sync_playwright
        save_dir = os.path.join(self.out_var.get(), f"page_{page_id}")
        os.makedirs(save_dir, exist_ok=True)
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch_persistent_context(
                    user_data_dir=USER_DATA_DIR, headless=False,
                    viewport={"width": 1280, "height": 720})
                page = browser.new_page()
                page.goto(f"{BASE_URL}/pages/viewpage.action?pageId={page_id}")
                try: page.wait_for_load_state("networkidle", timeout=30000)
                except: pass

                if not messagebox.askokcancel("로그인 확인",
                        "Confluence에 로그인되어 있나요?\n[확인] 진행  [취소] 중단"):
                    browser.close(); return

                page.reload(wait_until="networkidle")
                self.log("=" * 52)
                self.log(f"INFO: 수집 시작  |  페이지 ID: {page_id}")
                self.log(f"INFO: LLM 요약: {'켜짐' if self.llm_var.get() else '꺼짐'}")
                self.log("=" * 52)

                process_child_pages_recursive(
                    page, page_id, save_dir,
                    depth=self.depth_var.get(),
                    use_llm=self.llm_var.get(),
                    callback=self.log)

                self.log("INFO: 보고서 생성 중...")
                summary = generate_summary_report(
                    page, page_id, save_dir,
                    max_depth=self.depth_var.get(),
                    use_llm=self.llm_var.get(),
                    callback=self.log)
                rpt = os.path.join(save_dir, f"{page_id}_요약보고서.md")
                with open(rpt, "w", encoding="utf-8") as f:
                    f.write(summary)
                self.log(f"OK: 보고서 저장 완료")
                browser.close()
                self.log("DONE: 모든 작업 완료!")
                self._set_status("완료", GREEN)
                messagebox.showinfo("완료", f"완료!\n저장 위치: {os.path.abspath(save_dir)}")

        except Exception as e:
            self.log(f"ERR: {e}\n{traceback.format_exc()}")
            self._set_status("오류 발생", RED)
            messagebox.showerror("오류", str(e))
        finally:
            self._running = False
            self.root.after(0, lambda: (
                self.progress.stop(),
                self.run_btn.config_state("normal")))


def main():
    root = tk.Tk()

    style = ttk.Style()
    try: style.theme_use("clam")
    except: pass
    style.configure("TProgressbar", troughcolor=BORDER, background=BLUE,
                    bordercolor=BORDER, lightcolor=BLUE, darkcolor=BLUE_DK)

    ConfluenceGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
