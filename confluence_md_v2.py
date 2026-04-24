"""
Confluence 차일드 페이지 요약 GUI - v2
원본 구조 유지 + LLM 요약 추가 (팀장 보고용)
"""

# ── 패키지 자동 설치 ──────────────────────────
import subprocess, sys

def _install(pkg):
    subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

try:
    import html2text
except ImportError:
    print("html2text 설치 중..."); _install("html2text"); import html2text

try:
    from openai import OpenAI
    OPENAI_OK = True
except ImportError:
    print("openai 설치 중..."); _install("openai")
    from openai import OpenAI
    OPENAI_OK = True

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

LLM_API_KEY   = os.getenv("LLM_API_KEY",   "YOUR_API_KEY")
LLM_BASE_URL  = os.getenv("LLM_BASE_URL",  "https://your-endpoint/v1")
LLM_MODEL     = os.getenv("LLM_MODEL",     "qwen-plus")
LLM_MAX_TOKENS= int(os.getenv("LLM_MAX_TOKENS", "2000"))

# ─────────────────────────────────────────────
# LLM 요약 함수
# ─────────────────────────────────────────────
def llm_summarize_page(title: str, text: str) -> str:
    """페이지 내용 → 팀장 보고용 요약"""
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
                    "- 팀장은 각 담당자의 전문 기술을 잘 모릅니다. "
                    "전문 용어가 나오면 반드시 괄호로 쉽게 설명하세요.\n"
                    "  예) API 연동(서로 다른 시스템이 데이터를 주고받는 연결 작업)\n"
                    "- 격식체를 사용하세요. (~하였습니다, ~진행 중에 있습니다)\n"
                    "- 수치, 날짜, 완료 여부는 문서 그대로 포함하세요.\n\n"
                    "【보고서 형식】\n"
                    "## 개요\n"
                    "이 업무가 무엇인지, 왜 하는지 2~3문장으로 쉽게 설명\n\n"
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
    """선택된 MD 파일들 → 종합 보고서"""
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
                    "- 추측이나 창작 금지. 불확실한 내용은 '문서에 명시되지 않음'으로 표기.\n\n"
                    "【작성 방식】\n"
                    "- 각 팀/담당자가 무슨 일을 하는지 처음 보는 사람도 이해하도록 구체적으로 설명\n"
                    "- 전문 용어는 괄호로 쉬운 말 추가 (예: CI/CD(코드 변경 시 자동 테스트·배포 시스템))\n"
                    "- 격식 있는 문어체 사용 (~하였습니다, ~검토 중에 있습니다)\n"
                    "- 수치, 날짜, 완료율 등 구체적 수치는 반드시 포함\n\n"
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
    m = re.search(r'pages/(\d{8,})', url)
    if m: return m.group(1)
    return None

def extract_image_info(html_content: str) -> List[Dict]:
    images = []
    soup = BeautifulSoup(html_content, 'html.parser')
    for img in soup.find_all('img'):
        images.append({
            'src':   img.get('src', ''),
            'alt':   img.get('alt', ''),
            'title': img.get('title', ''),
        })
    return images

def analyze_confluence_content(html_content: str) -> Dict:
    soup = BeautifulSoup(html_content, 'html.parser')
    text_content = []
    for i in range(1, 7):
        for h in soup.find_all(f'h{i}'):
            text_content.append({'type': f'heading_{i}', 'content': h.get_text(strip=True)})
    for table in soup.find_all('table'):
        headers = [th.get_text(strip=True) for th in table.find_all('th')]
        rows = []
        for tr in table.find_all('tr'):
            row = [td.get_text(strip=True) for td in tr.find_all('td')]
            if row: rows.append(row)
        if headers or rows:
            text_content.append({'type': 'table', 'headers': headers, 'rows': rows})
    for ul in soup.find_all(['ul', 'ol']):
        items = [li.get_text(strip=True) for li in ul.find_all('li', recursive=False)]
        if items: text_content.append({'type': 'list', 'items': items})
    for p in soup.find_all('p'):
        t = p.get_text(strip=True)
        if t and len(t) > 10:
            text_content.append({'type': 'paragraph', 'content': t})
    return {'text_content': text_content, 'images': extract_image_info(html_content)}

def convert_to_markdown(analysis: Dict, title: str) -> str:
    md = [f"# {title}", "", "---",
          f"Generated at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", "---", ""]
    if analysis['images']:
        md += ["## 이미지", ""]
        for i, img in enumerate(analysis['images'], 1):
            md.append(f"{i}. **{img.get('alt','이미지')}**: {img.get('src','')}")
        md.append("")
    md += ["## 내용", ""]
    for item in analysis['text_content']:
        t = item['type']
        if t.startswith('heading_'):
            md.append(f"{'#' * int(t[-1])} {item['content']}")
            md.append("")
        elif t == 'table':
            h = item.get('headers', [])
            if h:
                md.append("| " + " | ".join(h) + " |")
                md.append("| " + " | ".join(["---"] * len(h)) + " |")
                for row in item.get('rows', []):
                    while len(row) < len(h): row.append("")
                    md.append("| " + " | ".join(row) + " |")
                md.append("")
        elif t == 'list':
            for it in item.get('items', []): md.append(f"- {it}")
            md.append("")
        elif t == 'paragraph':
            md.append(item['content'])
            md.append("")
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
            all_pages.append({"id": doc["id"], "title": doc["title"],
                               "url": doc.get("_links", {}).get("webui", "")})
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
        return {"id": d.get("id"), "title": d.get("title"), "html": html}
    except:
        return None

def process_child_pages_recursive(page, page_id: str, save_dir: str,
                                   depth: int = 2, use_llm: bool = True,
                                   callback=None):
    if callback: callback(f"페이지 {page_id} 처리 중...")
    try:
        data = get_page_content(page, page_id)
        if not data:
            if callback: callback(f"  [실패] {page_id}")
            return
        analysis = analyze_confluence_content(data['html'])
        markdown  = convert_to_markdown(analysis, data['title'])

        # LLM 요약 추가
        llm_block = ""
        if use_llm and LLM_API_KEY != "YOUR_API_KEY":
            if callback: callback(f"    LLM 요약 중: {data['title']}")
            summary  = llm_summarize_page(data['title'], markdown)
            llm_block = f"\n\n---\n## 🤖 AI 요약 (팀장 보고용)\n\n{summary}\n\n---\n"

        safe = clean_filename(data['title'])
        path = os.path.join(save_dir, f"{page_id}_{safe}.md")
        with open(path, "w", encoding="utf-8") as f:
            f.write(markdown + llm_block)
        if callback: callback(f"  [완료] {data['title']} → {path}")

        children = get_page_children(page, page_id)
        if children and depth > 0:
            child_dir = os.path.join(save_dir, f"children_{page_id}")
            os.makedirs(child_dir, exist_ok=True)
            for child in children:
                process_child_pages_recursive(page, child['id'], child_dir,
                                               depth - 1, use_llm, callback)
    except Exception as e:
        if callback: callback(f"[오류] {e}\n{traceback.format_exc()}")

def generate_summary_report(page, page_id: str, save_dir: str,
                             max_depth: int = 2, use_llm: bool = True,
                             callback=None) -> str:
    lines = ["# 차일드 페이지 요약 보고서", "",
             f"생성일: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
             f"최대 깊이: {max_depth}", "", "=" * 80, ""]

    page_data = get_page_content(page, page_id)
    if page_data:
        lines.append(f"📁 선택 페이지: {page_data['title']} (ID: {page_id})")
        lines.append("")

    def collect(pid, depth=0):
        result = []
        children = get_page_children(page, pid)
        for child in children:
            cd = get_page_content(page, child['id'])
            if cd:
                analysis = analyze_confluence_content(cd['html'])
                md_text  = convert_to_markdown(analysis, cd['title'])
                result.append({"id": child['id'], "title": cd['title'],
                                "depth": depth + 1, "analysis": analysis,
                                "md_text": md_text})
                if depth + 1 < max_depth:
                    result.extend(collect(child['id'], depth + 1))
        return result

    all_pages = collect(page_id)

    if not all_pages:
        lines.append("차일드 페이지가 없습니다.")
    else:
        # LLM 종합 보고서
        if use_llm and LLM_API_KEY != "YOUR_API_KEY":
            if callback: callback("LLM 종합 보고서 생성 중...")
            report = llm_generate_report([
                {"title": p["title"], "content": p["md_text"]} for p in all_pages
            ])
            lines += ["## 🤖 AI 종합 보고서 (팀장 보고용)", "", report, "", "=" * 80, ""]

        # 목차
        lines += ["## 페이지 목록", ""]
        for p in all_pages:
            indent = "  " * (p['depth'] - 1)
            lines.append(f"{indent}📄 {p['title']} (ID: {p['id']})")
            a = p['analysis']
            for item in a['text_content'][:3]:
                t = item['type']
                if t.startswith('heading_'):
                    lines.append(f"{indent}   {'#'*int(t[-1])} {item['content'][:80]}")
                elif t == 'paragraph':
                    lines.append(f"{indent}   {item['content'][:100]}...")
            if a['images']:
                lines.append(f"{indent}   🖼️ 이미지 {len(a['images'])}개")
            lines.append("")

    return "\n".join(lines)


# ─────────────────────────────────────────────
# GUI (원본 구조 유지)
# ─────────────────────────────────────────────
class ConfluenceChildSummaryGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Confluence 차일드 페이지 요약 v2")
        self.root.geometry("800x650")

        self.url_var    = tk.StringVar()
        self.depth_var  = tk.IntVar(value=7)
        self.out_var    = tk.StringVar(value="./confluence_output")
        self.llm_var    = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="준비됨")
        self.create_widgets()

    def create_widgets(self):
        f = ttk.Frame(self.root, padding="10")
        f.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        ttk.Label(f, text="Confluence URL:").grid(row=0, column=0, sticky="w", pady=5)
        self.url_entry = ttk.Entry(f, textvariable=self.url_var, width=70)
        self.url_entry.grid(row=0, column=1, columnspan=2, sticky="ew", pady=5)
        self.url_entry.insert(0, "https://confluence.sec.samsung.net/pages/...")

        self.pid_lbl = ttk.Label(f, text="페이지 ID: ", foreground="gray")
        self.pid_lbl.grid(row=1, column=1, sticky="w")

        ttk.Label(f, text="재귀 깊이:").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Spinbox(f, from_=0, to=15, textvariable=self.depth_var, width=10).grid(row=2, column=1, sticky="w")

        ttk.Checkbutton(f, text="🤖 LLM 요약 생성 (팀장 보고용)", variable=self.llm_var).grid(
            row=3, column=1, sticky="w", pady=3)

        ttk.Label(f, text="출력 폴더:").grid(row=4, column=0, sticky="w", pady=5)
        frm = ttk.Frame(f)
        frm.grid(row=4, column=1, columnspan=2, sticky="ew")
        ttk.Entry(frm, textvariable=self.out_var, width=50).pack(side="left")
        ttk.Button(frm, text="찾아보기", command=self.browse).pack(side="left", padx=5)

        self.run_btn = ttk.Button(f, text="▶ 실행", command=self.run)
        self.run_btn.grid(row=5, column=0, pady=10)

        self.progress = ttk.Progressbar(f, mode='indeterminate', length=650)
        self.progress.grid(row=6, column=0, columnspan=3, sticky="ew")

        ttk.Label(f, text="로그:").grid(row=7, column=0, sticky="w")
        lf = ttk.Frame(f)
        lf.grid(row=8, column=0, columnspan=3, sticky="nsew")
        f.rowconfigure(8, weight=1)
        f.columnconfigure(1, weight=1)
        sb = ttk.Scrollbar(lf)
        sb.pack(side="right", fill="y")
        self.log_box = tk.Text(lf, height=16, yscrollcommand=sb.set)
        self.log_box.pack(fill="both", expand=True)
        sb.config(command=self.log_box.yview)

        ttk.Label(f, textvariable=self.status_var, foreground="blue").grid(
            row=9, column=0, columnspan=3, pady=5)

        self.url_var.trace_add("write", self.on_url)

    def on_url(self, *_):
        pid = extract_page_id_from_url(self.url_var.get())
        if pid:
            self.pid_lbl.config(text=f"페이지 ID: {pid}", foreground="green")
        else:
            self.pid_lbl.config(text="페이지 ID: (찾을 수 없음)", foreground="red")

    def browse(self):
        d = filedialog.askdirectory()
        if d: self.out_var.set(d)

    def log(self, msg):
        try:
            self.log_box.insert("end", msg + "\n")
            self.log_box.see("end")
            self.root.update_idletasks()
        except: pass

    def run(self):
        pid = extract_page_id_from_url(self.url_var.get())
        if not pid:
            messagebox.showerror("오류", "URL에서 페이지 ID를 찾을 수 없습니다.")
            return
        self.run_btn.config(state="disabled")
        self.progress.start()
        self.status_var.set("실행 중...")
        threading.Thread(target=self._worker, args=(pid,), daemon=True).start()

    def _worker(self, page_id):
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
                self.log("=" * 50)
                self.log(f"수집 시작 | 페이지 ID: {page_id}")
                self.log(f"LLM 요약: {'켜짐' if self.llm_var.get() else '꺼짐'}")
                self.log("=" * 50)

                process_child_pages_recursive(
                    page, page_id, save_dir,
                    depth=self.depth_var.get(),
                    use_llm=self.llm_var.get(),
                    callback=self.log,
                )

                self.log("\n요약 보고서 생성 중...")
                summary = generate_summary_report(
                    page, page_id, save_dir,
                    max_depth=self.depth_var.get(),
                    use_llm=self.llm_var.get(),
                    callback=self.log,
                )
                rpt_path = os.path.join(save_dir, f"{page_id}_요약보고서.md")
                with open(rpt_path, "w", encoding="utf-8") as f:
                    f.write(summary)
                self.log(f"보고서 저장: {rpt_path}")

                browser.close()
                self.log("\n✅ 완료!")
                self.status_var.set("완료")
                messagebox.showinfo("완료", f"완료!\n저장 위치: {os.path.abspath(save_dir)}")
        except Exception as e:
            self.log(f"[오류] {e}\n{traceback.format_exc()}")
            self.status_var.set("오류")
            messagebox.showerror("오류", str(e))
        finally:
            self.root.after(0, lambda: (self.progress.stop(),
                                        self.run_btn.config(state="normal")))


def main():
    root = tk.Tk()
    ConfluenceChildSummaryGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
