"""
confluence_md_v2.py - Confluence 차일드 페이지 수집 + LLM 분석 (개선판)

개선사항:
  1. html2text 로 문서 순서 유지한 정확한 MD 변환
  2. 이미지 다운로드 → base64 → Vision LLM 으로 실제 내용 분석
  3. 페이지별 LLM 요약 (단순 텍스트 자르기 X)
  4. MD 파일 선택 → 보고서 생성 탭 추가
"""

import os, re, base64, threading, traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, List, Optional
from datetime import datetime
from bs4 import BeautifulSoup
import html2text
from openai import OpenAI

# ─────────────────────────────────────────────
# 설정 (회사 환경에 맞게 수정)
# ─────────────────────────────────────────────
BASE_URL        = "https://confluence.sec.samsung.net"
USER_DATA_DIR   = "./chrome_profile_confluence_md"

LLM_API_KEY     = os.getenv("LLM_API_KEY", "YOUR_API_KEY")
LLM_BASE_URL    = os.getenv("LLM_BASE_URL", "https://your-lucode-endpoint/v1")
LLM_MODEL       = os.getenv("LLM_MODEL",    "gpt-4o")
LLM_VISION_MODEL= os.getenv("LLM_VISION_MODEL", "gpt-4o")   # 이미지 분석용 (같아도 됨)
LLM_MAX_TOKENS  = int(os.getenv("LLM_MAX_TOKENS", "2000"))

# ─────────────────────────────────────────────
# LLM 클라이언트
# ─────────────────────────────────────────────
def _llm():
    return OpenAI(api_key=LLM_API_KEY, base_url=LLM_BASE_URL)


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
        combined += f"\n\n### {item['title']}\n{item['content'][:2000]}"

    try:
        resp = _llm().chat.completions.create(
            model=LLM_MODEL,
            messages=[
                {"role": "system", "content":
                    "당신은 기업 업무 보고서 작성 전문가입니다.\n"
                    "여러 Confluence 페이지를 종합해 경영진 보고서를 작성하세요.\n\n"
                    "보고서 형식:\n"
                    "1. 전체 요약\n"
                    "2. 주요 성과 및 완료 항목\n"
                    "3. 진행 중인 주요 과제\n"
                    "4. 공통 이슈 및 리스크\n"
                    "5. 다음 단계 / 액션 아이템\n\n"
                    "한국어, 명확하고 간결하게, 불릿 포인트 활용."},
                {"role": "user", "content":
                    f"다음 페이지들을 종합 분석해 보고서를 작성해주세요:\n{combined}"}
            ],
            max_tokens=LLM_MAX_TOKENS * 2,
            temperature=0.3,
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
# GUI
# ─────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Confluence MD 수집 + 보고서 생성 v2")
        self.geometry("900x700")
        self.resizable(True, True)

        nb = ttk.Notebook(self)
        nb.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tab_crawl  = CrawlTab(nb)
        self.tab_report = ReportTab(nb)

        nb.add(self.tab_crawl,  text="📥 수집")
        nb.add(self.tab_report, text="📊 보고서 생성")


# ── 수집 탭 ──────────────────────────────────
class CrawlTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, padding=10)
        self._build()

    def _build(self):
        # URL
        ttk.Label(self, text="Confluence URL:").grid(row=0, column=0, sticky=tk.W, pady=4)
        self.url_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.url_var, width=65).grid(row=0, column=1, columnspan=2, sticky=tk.EW, pady=4)
        self.url_var.trace_add("write", self._on_url)

        self.pid_lbl = ttk.Label(self, text="페이지 ID: —", foreground="gray")
        self.pid_lbl.grid(row=1, column=1, sticky=tk.W)

        # 깊이
        ttk.Label(self, text="재귀 깊이:").grid(row=2, column=0, sticky=tk.W, pady=4)
        self.depth_var = tk.IntVar(value=3)
        ttk.Spinbox(self, from_=0, to=10, textvariable=self.depth_var, width=8).grid(row=2, column=1, sticky=tk.W)

        # 옵션
        self.vision_var  = tk.BooleanVar(value=True)
        self.summary_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self, text="이미지 Vision AI 분석", variable=self.vision_var).grid(row=3, column=1, sticky=tk.W)
        ttk.Checkbutton(self, text="페이지별 LLM 요약 생성", variable=self.summary_var).grid(row=4, column=1, sticky=tk.W)

        # 출력 폴더
        ttk.Label(self, text="출력 폴더:").grid(row=5, column=0, sticky=tk.W, pady=4)
        self.out_var = tk.StringVar(value="./confluence_output")
        frm = ttk.Frame(self)
        frm.grid(row=5, column=1, columnspan=2, sticky=tk.EW)
        ttk.Entry(frm, textvariable=self.out_var, width=50).pack(side=tk.LEFT)
        ttk.Button(frm, text="찾아보기", command=self._browse).pack(side=tk.LEFT, padx=4)

        # 실행
        self.btn = ttk.Button(self, text="▶ 수집 시작", command=self._run)
        self.btn.grid(row=6, column=0, pady=8)

        self.prog = ttk.Progressbar(self, mode="indeterminate", length=700)
        self.prog.grid(row=7, column=0, columnspan=3, sticky=tk.EW)

        # 로그
        ttk.Label(self, text="로그:").grid(row=8, column=0, sticky=tk.W)
        lf = ttk.Frame(self)
        lf.grid(row=9, column=0, columnspan=3, sticky=tk.NSEW)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(9, weight=1)
        sb = ttk.Scrollbar(lf)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_box = tk.Text(lf, height=16, yscrollcommand=sb.set, wrap=tk.WORD)
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


# ── 보고서 생성 탭 ────────────────────────────
class ReportTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, padding=10)
        self.md_files: List[str] = []
        self._build()

    def _build(self):
        ttk.Label(self, text="MD 파일 선택 (여러 개 가능):").grid(row=0, column=0, sticky=tk.W, pady=4)

        btn_frm = ttk.Frame(self)
        btn_frm.grid(row=0, column=1, sticky=tk.W, pady=4)
        ttk.Button(btn_frm, text="📂 폴더에서 MD 불러오기", command=self._load_folder).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frm, text="📄 파일 직접 선택",         command=self._load_files).pack(side=tk.LEFT)

        # 파일 목록
        lf = ttk.LabelFrame(self, text="선택된 파일 목록")
        lf.grid(row=1, column=0, columnspan=2, sticky=tk.NSEW, pady=6)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        sb = ttk.Scrollbar(lf)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_list = tk.Listbox(lf, selectmode=tk.EXTENDED, height=10,
                                    yscrollcommand=sb.set)
        self.file_list.pack(fill=tk.BOTH, expand=True)
        sb.config(command=self.file_list.yview)
        ttk.Button(lf, text="선택 항목 제거", command=self._remove_selected).pack(anchor=tk.E, padx=4, pady=2)

        # 보고서 옵션
        opt_frm = ttk.LabelFrame(self, text="보고서 옵션")
        opt_frm.grid(row=2, column=0, columnspan=2, sticky=tk.EW, pady=6)
        self.report_title_var = tk.StringVar(value="주간 보고서")
        ttk.Label(opt_frm, text="보고서 제목:").grid(row=0, column=0, padx=6, pady=4, sticky=tk.W)
        ttk.Entry(opt_frm, textvariable=self.report_title_var, width=40).grid(row=0, column=1, sticky=tk.W)

        self.out_var = tk.StringVar(value="./reports")
        ttk.Label(opt_frm, text="저장 폴더:").grid(row=1, column=0, padx=6, pady=4, sticky=tk.W)
        frm2 = ttk.Frame(opt_frm)
        frm2.grid(row=1, column=1, sticky=tk.W)
        ttk.Entry(frm2, textvariable=self.out_var, width=40).pack(side=tk.LEFT)
        ttk.Button(frm2, text="찾아보기",
                   command=lambda: self.out_var.set(filedialog.askdirectory() or self.out_var.get())
                   ).pack(side=tk.LEFT, padx=4)

        # 실행
        self.btn = ttk.Button(self, text="📊 보고서 생성 (LLM)", command=self._generate)
        self.btn.grid(row=3, column=0, pady=8)

        self.prog = ttk.Progressbar(self, mode="indeterminate", length=700)
        self.prog.grid(row=4, column=0, columnspan=2, sticky=tk.EW)

        # 로그
        ttk.Label(self, text="로그:").grid(row=5, column=0, sticky=tk.W)
        lf2 = ttk.Frame(self)
        lf2.grid(row=6, column=0, columnspan=2, sticky=tk.NSEW, pady=4)
        self.rowconfigure(6, weight=1)
        sb2 = ttk.Scrollbar(lf2)
        sb2.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_box = tk.Text(lf2, height=10, yscrollcommand=sb2.set, wrap=tk.WORD)
        self.log_box.pack(fill=tk.BOTH, expand=True)
        sb2.config(command=self.log_box.yview)

    def _load_folder(self):
        d = filedialog.askdirectory()
        if not d: return
        found = []
        for root, _, files in os.walk(d):
            for f in files:
                if f.endswith(".md"):
                    found.append(os.path.join(root, f))
        self._add_files(sorted(found))

    def _load_files(self):
        files = filedialog.askopenfilenames(
            filetypes=[("Markdown", "*.md"), ("All", "*.*")])
        self._add_files(list(files))

    def _add_files(self, paths):
        for p in paths:
            if p not in self.md_files:
                self.md_files.append(p)
                self.file_list.insert(tk.END, os.path.basename(p))

    def _remove_selected(self):
        for i in reversed(self.file_list.curselection()):
            self.file_list.delete(i)
            self.md_files.pop(i)

    def log(self, msg):
        try:
            self.log_box.insert(tk.END, msg + "\n")
            self.log_box.see(tk.END)
            self.update_idletasks()
        except: pass

    def _generate(self):
        if not self.md_files:
            messagebox.showwarning("경고", "MD 파일을 먼저 선택하세요.")
            return
        self.btn.config(state="disabled")
        self.prog.start()
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            self.log(f"선택된 파일: {len(self.md_files)}개")
            contents = []
            for path in self.md_files:
                with open(path, encoding="utf-8") as f:
                    text = f.read()
                title = os.path.basename(path).replace(".md", "")
                contents.append({"title": title, "content": text})
                self.log(f"  읽기 완료: {title}")

            self.log("\nLLM 보고서 생성 중...")
            report_md = llm_generate_report(contents)

            os.makedirs(self.out_var.get(), exist_ok=True)
            ts    = datetime.now().strftime("%Y%m%d_%H%M")
            title = self.report_title_var.get()
            fname = os.path.join(self.out_var.get(), f"{title}_{ts}.md")

            header = (
                f"# {title}\n\n"
                f"---\n"
                f"생성일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"참고 파일: {len(self.md_files)}개\n"
                f"---\n\n"
            )
            with open(fname, "w", encoding="utf-8") as f:
                f.write(header + report_md)

            self.log(f"\n✅ 보고서 저장 완료: {fname}")
            messagebox.showinfo("완료", f"보고서 생성 완료!\n{fname}")

        except Exception as e:
            self.log(f"[오류] {e}\n{traceback.format_exc()}")
            messagebox.showerror("오류", str(e))
        finally:
            self.after(0, lambda: (self.prog.stop(), self.btn.config(state="normal")))


# ─────────────────────────────────────────────
if __name__ == "__main__":
    App().mainloop()
