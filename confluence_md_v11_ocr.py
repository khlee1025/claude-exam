
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

def llm_analyze_image(image_bytes: bytes, context: str = "", mime_type: str = "image/png") -> str:
    """
    Vision API path:
      image bytes -> base64 -> Qwen Vision direct analysis
    """
    try:
        b64 = base64.b64encode(image_bytes).decode()
        data_url = f"data:{mime_type};base64,{b64}"
        resp = _llm().chat.completions.create(
            model=LLM_VISION_MODEL,
            temperature=0.2,
            max_tokens=min(1000, LLM_MAX_TOKENS),
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": (
                                "당신은 기업 문서의 이미지, 표, 차트, 화면 캡처를 분석하는 전문가입니다.\n"
                                f"이미지 참고 정보: {context or '없음'}\n\n"
                                "이 이미지를 업무 보고서 관점에서 분석해주세요.\n\n"
                                "다음 항목을 최우선으로 추출하세요:\n"
                                "- 그래프/차트: 축 레이블, 수치, 단위, 추세\n"
                                "- 표: 행/열 데이터를 마크다운 표 형식으로 재현\n"
                                "- 수치/퍼센트/날짜: 정확히 그대로 기록\n"
                                "- 상태/색상 의미: 빨강=위험, 노랑=주의 등\n\n"
                                "읽기 어려운 글자는 '(판독불가)'로 표시하세요.\n"
                                "이미지가 단순 로고·아이콘이면 '단순 그래픽, 수치 없음'으로만 답하세요."
                            ),
                        },
                        {
                            "type": "image_url",
                            "image_url": {"url": data_url},
                        },
                    ],
                }
            ],
        )
        return resp.choices[0].message.content
    except Exception as e:
        return f"[Vision 이미지 분석 실패: {type(e).__name__}: {e}]"


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

def html_to_markdown(html: str, page_session=None, base_url: str = "", callback=None, img_save_dir: str = "") -> str:
    def _log(msg):
        if callback: callback(msg)
    _img_counter = [0]
    BeautifulSoup = _need_beautifulsoup()
    html2text = _need_html2text()
    soup = BeautifulSoup(html, "html.parser")

    imgs = soup.find_all("img")
    _log(f"[이미지 탐색] <img> 태그 발견: {len(imgs)}개")
    for img in imgs:
        src = img.get("src", "")
        alt = img.get("alt", "이미지")
        if not src:
            _log(f"  → src 없음, 스킵")
            continue

        img_desc = None
        if page_session and src:
            try:
                full_src = urljoin(base_url + "/", src) if not src.startswith("http") else src
                _log(f"  [이미지] src={src[:80]}")
                _log(f"  [이미지] full_url={full_src[:100]}")
                # page.request.get() 은 SSO/HTTPOnly 쿠키를 제대로 전달 못함
                # → 브라우저 컨텍스트의 fetch() 사용 (모든 쿠키 자동 포함)
                result = page_session.evaluate("""
                    async (url) => {
                        try {
                            const resp = await fetch(url, {credentials: 'include'});
                            if (!resp.ok) return {status: resp.status, data: null, contentType: ''};
                            const buf = await resp.arrayBuffer();
                            const bytes = Array.from(new Uint8Array(buf));
                            return {
                                status: resp.status,
                                data: bytes,
                                contentType: resp.headers.get('content-type') || ''
                            };
                        } catch(e) {
                            return {status: 0, data: null, contentType: '', error: e.message};
                        }
                    }
                """, full_src)

                status_code = result.get("status", 0)
                _log(f"  [이미지] HTTP {status_code}, contentType={result.get('contentType','')}, size={len(result.get('data') or [])} bytes")
                if result.get("status") == 200 and result.get("data"):
                    img_bytes = bytes(result["data"])
                    content_type = result.get("contentType", "image/png").split(";")[0].strip().lower()
                    if content_type == "image/jpg":
                        content_type = "image/jpeg"

                    if content_type in ("image/png", "image/jpeg", "image/webp", "image/gif") and len(img_bytes) > 300:
                        img_desc = llm_analyze_image(img_bytes, context=alt, mime_type=content_type)
                        # 이미지 파일 저장 (Word 삽입용)
                        if img_save_dir:
                            try:
                                os.makedirs(img_save_dir, exist_ok=True)
                                _img_counter[0] += 1
                                ext = content_type.split("/")[-1].replace("jpeg", "jpg")
                                img_filename = f"img_{_img_counter[0]:03d}.{ext}"
                                img_path = os.path.join(img_save_dir, img_filename)
                                with open(img_path, "wb") as _f:
                                    _f.write(img_bytes)
                                # MD에 상대 경로 마커 기록 (나중에 Word 삽입용)
                                rel_path = os.path.join("images", img_filename).replace("\\", "/")
                                img_desc = img_desc + f"\n\n[EMBED_IMAGE:{rel_path}]"
                                _log(f"  [이미지] 저장: {img_filename}")
                            except Exception as _e:
                                _log(f"  [이미지] 파일 저장 실패: {_e}")
                    else:
                        img_desc = f"[이미지 분석 스킵: content-type={content_type}, size={len(img_bytes)} bytes]"
                else:
                    status = result.get("status", 0)
                    err = result.get("error", "")
                    img_desc = f"[이미지 다운로드 실패: HTTP {status}{(' - ' + err) if err else ''}]"
            except Exception as e:
                img_desc = f"[이미지 다운로드 실패: {type(e).__name__}: {e}]"

        replacement = soup.new_tag("p")
        if img_desc:
            # [EMBED_IMAGE:path] 마커가 있으면 분리해서 별도 태그로
            if "[EMBED_IMAGE:" in img_desc:
                parts = img_desc.split("[EMBED_IMAGE:", 1)
                desc_text = parts[0].strip()
                embed_path = parts[1].rstrip("]").strip()
                replacement.string = f"**[이미지 분석: {alt}]** {desc_text}"
                img.replace_with(replacement)
                # EMBED_IMAGE 마커를 별도 p 태그로 삽입
                embed_tag = soup.new_tag("p")
                embed_tag.string = f"[EMBED_IMAGE:{embed_path}]"
                replacement.insert_after(embed_tag)
                continue
            else:
                replacement.string = f"**[이미지 분석: {alt}]** {img_desc}"
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
            params={"expand": "body.view"},
        )
        if resp.status != 200:
            return None
        d = resp.json()
        return {
            "id": d.get("id"),
            "title": d.get("title", ""),
            "html": d.get("body", {}).get("view", {}).get("value", ""),
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
    img_dir = os.path.join(save_dir, "images") if use_vision else ""
    md_body = html_to_markdown(data["html"], page_session=ps, base_url=BASE_URL, callback=callback, img_save_dir=img_dir)

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


# ──────────────────────────────────────────────────────────────────────────────
# Word 보고서 생성 (docx-js 기반, node.js 필요)
# ──────────────────────────────────────────────────────────────────────────────
import shutil as _shutil
import subprocess as _subprocess
import base64 as _base64

_MAKE_REPORT_JS_B64 = "LyoqCiAqIENvbmZsdWVuY2UgTUQg4oaSIFdvcmQg67O06rOg7IScIOyDneyEseq4sCAoU2Ftc3VuZyBTdHlsZSkKICogVXNhZ2U6CiAqICAgbm9kZSBtYWtlX3JlcG9ydC5qcyA8aW5wdXQubWR8Zm9sZGVyPiA8b3V0cHV0LmRvY3g+IFt0aXRsZV0KICovCiJ1c2Ugc3RyaWN0IjsKY29uc3QgZnMgICA9IHJlcXVpcmUoImZzIik7CmNvbnN0IHBhdGggPSByZXF1aXJlKCJwYXRoIik7CmNvbnN0IHsKICBEb2N1bWVudCwgUGFja2VyLCBQYXJhZ3JhcGgsIFRleHRSdW4sIFRhYmxlLCBUYWJsZVJvdywgVGFibGVDZWxsLAogIEhlYWRlciwgRm9vdGVyLCBBbGlnbm1lbnRUeXBlLCBIZWFkaW5nTGV2ZWwsIEJvcmRlclN0eWxlLCBXaWR0aFR5cGUsCiAgU2hhZGluZ1R5cGUsIFZlcnRpY2FsQWxpZ24sIFBhZ2VOdW1iZXIsIFBhZ2VCcmVhaywgVGFibGVPZkNvbnRlbnRzLAogIExldmVsRm9ybWF0LAp9ID0gcmVxdWlyZSgiZG9jeCIpOwoKLy8g4pSA4pSA4pSAIOyDieyDgSDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKY29uc3QgQyA9IHsKICBibHVlOiAgICAgICIxNDI4QTAiLAogIGxpZ2h0Qmx1ZTogIkU4RUJGOCIsCiAgbWlkQmx1ZTogICAiQzdDREU4IiwKICByb3dBbHQ6ICAgICJGNEY2RkQiLAogIGdyYXk6ICAgICAgIjZCNzI4MCIsCiAgcmVkOiAgICAgICAiREMyNjI2IiwKICBncmVlbjogICAgICIxNkEzNEEiLAogIHdoaXRlOiAgICAgIkZGRkZGRiIsCiAgZGFyazogICAgICAiMUYyOTM3IiwKfTsKCi8vIOKUgOKUgOKUgCDsnKDti7gg4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSACmZ1bmN0aW9uIGJvcmRlcihjb2xvciA9IEMubWlkQmx1ZSwgc2l6ZSA9IDQpIHsKICBjb25zdCBiID0geyBzdHlsZTogQm9yZGVyU3R5bGUuU0lOR0xFLCBzaXplLCBjb2xvciB9OwogIHJldHVybiB7IHRvcDogYiwgYm90dG9tOiBiLCBsZWZ0OiBiLCByaWdodDogYiB9Owp9CgpmdW5jdGlvbiBwYXJzZUlubGluZSh0ZXh0LCBzaXplID0gMjIsIGRlZmF1bHRDb2xvciA9IEMuZGFyaykgewogIGNvbnN0IHJ1bnMgPSBbXTsKICBjb25zdCByZSAgID0gLyhcKlwqKC4rPylcKlwqfFwqKC4rPylcKnxgKFteYF0rKWApL2c7CiAgbGV0IGxhc3QgPSAwLCBtOwogIHdoaWxlICgobSA9IHJlLmV4ZWModGV4dCkpICE9PSBudWxsKSB7CiAgICBpZiAobS5pbmRleCA+IGxhc3QpCiAgICAgIHJ1bnMucHVzaChuZXcgVGV4dFJ1bih7IHRleHQ6IHRleHQuc2xpY2UobGFzdCwgbS5pbmRleCksIHNpemUsIGZvbnQ6ICLrp5HsnYAg6rOg65SVIiwgY29sb3I6IGRlZmF1bHRDb2xvciB9KSk7CiAgICBpZiAgICAgIChtWzJdKSBydW5zLnB1c2gobmV3IFRleHRSdW4oeyB0ZXh0OiBtWzJdLCBib2xkOiB0cnVlLCAgICBzaXplLCBmb250OiAi66eR7J2AIOqzoOuUlSIsIGNvbG9yOiBkZWZhdWx0Q29sb3IgfSkpOwogICAgZWxzZSBpZiAobVszXSkgcnVucy5wdXNoKG5ldyBUZXh0UnVuKHsgdGV4dDogbVszXSwgaXRhbGljczogdHJ1ZSwgc2l6ZSwgZm9udDogIuunkeydgCDqs6DrlJUiLCBjb2xvcjogZGVmYXVsdENvbG9yIH0pKTsKICAgIGVsc2UgaWYgKG1bNF0pIHJ1bnMucHVzaChuZXcgVGV4dFJ1bih7IHRleHQ6IG1bNF0sIGZvbnQ6ICJDb25zb2xhcyIsIHNpemU6IHNpemUgLSAyLAogICAgICBzaGFkaW5nOiB7IHR5cGU6IFNoYWRpbmdUeXBlLkNMRUFSLCBmaWxsOiAiRjNGNEY2IiB9LCBjb2xvcjogIkI5MUMxQyIgfSkpOwogICAgbGFzdCA9IG0uaW5kZXggKyBtWzBdLmxlbmd0aDsKICB9CiAgaWYgKGxhc3QgPCB0ZXh0Lmxlbmd0aCkKICAgIHJ1bnMucHVzaChuZXcgVGV4dFJ1bih7IHRleHQ6IHRleHQuc2xpY2UobGFzdCksIHNpemUsIGZvbnQ6ICLrp5HsnYAg6rOg65SVIiwgY29sb3I6IGRlZmF1bHRDb2xvciB9KSk7CiAgcmV0dXJuIHJ1bnMubGVuZ3RoID8gcnVucyA6IFtuZXcgVGV4dFJ1bih7IHRleHQsIHNpemUsIGZvbnQ6ICLrp5HsnYAg6rOg65SVIiwgY29sb3I6IGRlZmF1bHRDb2xvciB9KV07Cn0KCi8vIOKUgOKUgOKUgCDtl6TrlKkg4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSACmNvbnN0IEhFQURJTkdfQ0ZHID0gewogIDE6IHsgbGV2ZWw6IEhlYWRpbmdMZXZlbC5IRUFESU5HXzEsIHNpemU6IDM2LCBjb2xvcjogQy5ibHVlLCAgYmVmb3JlOiA0MDAsIGFmdGVyOiAyMDAsIG91dGxpbmU6IDAgfSwKICAyOiB7IGxldmVsOiBIZWFkaW5nTGV2ZWwuSEVBRElOR18yLCBzaXplOiAzMCwgY29sb3I6IEMuYmx1ZSwgIGJlZm9yZTogMzAwLCBhZnRlcjogMTYwLCBvdXRsaW5lOiAxIH0sCiAgMzogeyBsZXZlbDogSGVhZGluZ0xldmVsLkhFQURJTkdfMywgc2l6ZTogMjYsIGNvbG9yOiAiMzMzMzMzIixiZWZvcmU6IDI0MCwgYWZ0ZXI6IDEyMCwgb3V0bGluZTogMiB9LAogIDQ6IHsgbGV2ZWw6IEhlYWRpbmdMZXZlbC5IRUFESU5HXzQsIHNpemU6IDI0LCBjb2xvcjogIjU1NTU1NSIsYmVmb3JlOiAxODAsIGFmdGVyOiAgODAsIG91dGxpbmU6IDMgfSwKfTsKZnVuY3Rpb24gbWFrZUhlYWRpbmcodGV4dCwgZGVwdGgpIHsKICBjb25zdCBjID0gSEVBRElOR19DRkdbTWF0aC5taW4oZGVwdGgsIDQpXSB8fCBIRUFESU5HX0NGR1s0XTsKICByZXR1cm4gbmV3IFBhcmFncmFwaCh7CiAgICBoZWFkaW5nOiBjLmxldmVsLAogICAgc3BhY2luZzogeyBiZWZvcmU6IGMuYmVmb3JlLCBhZnRlcjogYy5hZnRlciB9LAogICAgY2hpbGRyZW46IFtuZXcgVGV4dFJ1bih7IHRleHQsIGJvbGQ6IHRydWUsIHNpemU6IGMuc2l6ZSwgY29sb3I6IGMuY29sb3IsIGZvbnQ6ICLrp5HsnYAg6rOg65SVIiB9KV0sCiAgfSk7Cn0KCi8vIOKUgOKUgOKUgCDtkZwg4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSACmZ1bmN0aW9uIG1ha2VUYWJsZShoZWFkZXJzLCByb3dzKSB7CiAgY29uc3QgY29scyAgID0gTWF0aC5tYXgoaGVhZGVycy5sZW5ndGgsIC4uLnJvd3MubWFwKHIgPT4gci5sZW5ndGgpLCAxKTsKICBjb25zdCB0b3RhbFcgPSA5MDI2OwogIGNvbnN0IGNvbFcgICA9IE1hdGguZmxvb3IodG90YWxXIC8gY29scyk7CiAgY29uc3QgY29sV3MgID0gQXJyYXkoY29scykuZmlsbChjb2xXKTsKCiAgY29uc3QgaGRyUm93ID0gbmV3IFRhYmxlUm93KHsKICAgIHRhYmxlSGVhZGVyOiB0cnVlLAogICAgY2hpbGRyZW46IGhlYWRlcnMubWFwKChoLCBpKSA9PiBuZXcgVGFibGVDZWxsKHsKICAgICAgd2lkdGg6ICAgeyBzaXplOiBjb2xXc1tpXSwgdHlwZTogV2lkdGhUeXBlLkRYQSB9LAogICAgICBib3JkZXJzOiBib3JkZXIoQy5ibHVlKSwKICAgICAgc2hhZGluZzogeyB0eXBlOiBTaGFkaW5nVHlwZS5DTEVBUiwgZmlsbDogQy5ibHVlIH0sCiAgICAgIG1hcmdpbnM6IHsgdG9wOiAxMDAsIGJvdHRvbTogMTAwLCBsZWZ0OiAxNjAsIHJpZ2h0OiAxNjAgfSwKICAgICAgdmVydGljYWxBbGlnbjogVmVydGljYWxBbGlnbi5DRU5URVIsCiAgICAgIGNoaWxkcmVuOiBbbmV3IFBhcmFncmFwaCh7IGFsaWdubWVudDogQWxpZ25tZW50VHlwZS5DRU5URVIsCiAgICAgICAgY2hpbGRyZW46IFtuZXcgVGV4dFJ1bih7IHRleHQ6IGgsIGJvbGQ6IHRydWUsIGNvbG9yOiBDLndoaXRlLCBzaXplOiAyMCwgZm9udDogIuunkeydgCDqs6DrlJUiIH0pXSB9KV0sCiAgICB9KSksCiAgfSk7CgogIGNvbnN0IGRhdGFSb3dzID0gcm93cy5tYXAoKHJvdywgcmkpID0+CiAgICBuZXcgVGFibGVSb3coewogICAgICBjaGlsZHJlbjogQXJyYXkoY29scykuZmlsbChudWxsKS5tYXAoKF8sIGNpKSA9PiB7CiAgICAgICAgY29uc3QgdmFsICAgPSAocm93W2NpXSB8fCAiIikudHJpbSgpOwogICAgICAgIGNvbnN0IGlzUmVkID0gL+yngOyXsHzstIjqs7x87Iuk7YyofOyYpOulmHzqsr3qs6AvLnRlc3QodmFsKTsKICAgICAgICBjb25zdCBpc0dybiA9IC/soJXsg4F87JmE66OMfOyEseqztS8udGVzdCh2YWwpOwogICAgICAgIHJldHVybiBuZXcgVGFibGVDZWxsKHsKICAgICAgICAgIHdpZHRoOiAgIHsgc2l6ZTogY29sV3NbY2ldLCB0eXBlOiBXaWR0aFR5cGUuRFhBIH0sCiAgICAgICAgICBib3JkZXJzOiBib3JkZXIoQy5taWRCbHVlKSwKICAgICAgICAgIHNoYWRpbmc6IHsgdHlwZTogU2hhZGluZ1R5cGUuQ0xFQVIsIGZpbGw6IHJpICUgMiA9PT0gMCA/IEMud2hpdGUgOiBDLnJvd0FsdCB9LAogICAgICAgICAgbWFyZ2luczogeyB0b3A6IDgwLCBib3R0b206IDgwLCBsZWZ0OiAxNjAsIHJpZ2h0OiAxNjAgfSwKICAgICAgICAgIHZlcnRpY2FsQWxpZ246IFZlcnRpY2FsQWxpZ24uQ0VOVEVSLAogICAgICAgICAgY2hpbGRyZW46IFtuZXcgUGFyYWdyYXBoKHsgY2hpbGRyZW46IFtuZXcgVGV4dFJ1bih7CiAgICAgICAgICAgIHRleHQ6IHZhbCwgc2l6ZTogMjAsIGZvbnQ6ICLrp5HsnYAg6rOg65SVIiwKICAgICAgICAgICAgY29sb3I6IGlzUmVkID8gQy5yZWQgOiBpc0dybiA/IEMuZ3JlZW4gOiBDLmRhcmssCiAgICAgICAgICAgIGJvbGQ6IGlzUmVkIHx8IGlzR3JuLAogICAgICAgICAgfSldIH0pXSwKICAgICAgICB9KTsKICAgICAgfSksCiAgICB9KQogICk7CgogIHJldHVybiBuZXcgVGFibGUoeyB3aWR0aDogeyBzaXplOiB0b3RhbFcsIHR5cGU6IFdpZHRoVHlwZS5EWEEgfSwgY29sdW1uV2lkdGhzOiBjb2xXcywgcm93czogW2hkclJvdywgLi4uZGF0YVJvd3NdIH0pOwp9CgovLyDilIDilIDilIAg7L2c7JWE7JuDICjsnbTrr7jsp4Ag67aE7ISdKSDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKZnVuY3Rpb24gbWFrZUNhbGxvdXQodGV4dCkgewogIGNvbnN0IGNsZWFuID0gdGV4dC5yZXBsYWNlKC9cKlwqXFvsnbTrr7jsp4BbXlxdXSpcXVwqXCpccyovZywgIiIpLnJlcGxhY2UoL1wqXCovZywgIiIpLnRyaW0oKTsKICByZXR1cm4gbmV3IFBhcmFncmFwaCh7CiAgICBzcGFjaW5nOiB7IGJlZm9yZTogMTIwLCBhZnRlcjogMTIwIH0sCiAgICBpbmRlbnQ6ICB7IGxlZnQ6IDM2MCB9LAogICAgYm9yZGVyOiAgeyBsZWZ0OiB7IHN0eWxlOiBCb3JkZXJTdHlsZS5USElDSywgc2l6ZTogMjAsIGNvbG9yOiBDLmJsdWUsIHNwYWNlOiA4IH0gfSwKICAgIHNoYWRpbmc6IHsgdHlwZTogU2hhZGluZ1R5cGUuQ0xFQVIsIGZpbGw6IEMubGlnaHRCbHVlIH0sCiAgICBjaGlsZHJlbjogWwogICAgICBuZXcgVGV4dFJ1bih7IHRleHQ6ICLwn5SNIOydtOuvuOyngCDrtoTshJ0gICIsIGJvbGQ6IHRydWUsIHNpemU6IDIwLCBmb250OiAi66eR7J2AIOqzoOuUlSIsIGNvbG9yOiBDLmJsdWUgfSksCiAgICAgIG5ldyBUZXh0UnVuKHsgdGV4dDogY2xlYW4sIHNpemU6IDIwLCBmb250OiAi66eR7J2AIOqzoOuUlSIsIGNvbG9yOiAiMkQzQThBIiwgaXRhbGljczogdHJ1ZSB9KSwKICAgIF0sCiAgfSk7Cn0KCi8vIOKUgOKUgOKUgCDtjpjsnbTsp4Ag7Lac7LKYIOuwsOuEiCDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKZnVuY3Rpb24gbWFrZVNvdXJjZUJhbm5lcihwYWdlVGl0bGUpIHsKICByZXR1cm4gbmV3IFBhcmFncmFwaCh7CiAgICBzcGFjaW5nOiB7IGJlZm9yZTogMCwgYWZ0ZXI6IDE2MCB9LAogICAgc2hhZGluZzogeyB0eXBlOiBTaGFkaW5nVHlwZS5DTEVBUiwgZmlsbDogQy5saWdodEJsdWUgfSwKICAgIGJvcmRlcjogIHsgbGVmdDogeyBzdHlsZTogQm9yZGVyU3R5bGUuVEhJQ0ssIHNpemU6IDE2LCBjb2xvcjogQy5ibHVlLCBzcGFjZTogNiB9IH0sCiAgICBpbmRlbnQ6ICB7IGxlZnQ6IDE2MCB9LAogICAgY2hpbGRyZW46IFtuZXcgVGV4dFJ1bih7IHRleHQ6IGDwn5OEICAke3BhZ2VUaXRsZX1gLCBzaXplOiAxOSwgZm9udDogIuunkeydgCDqs6DrlJUiLCBjb2xvcjogQy5ncmF5LCBpdGFsaWNzOiB0cnVlIH0pXSwKICB9KTsKfQoKLy8g4pSA4pSA4pSAIE1EIOKGkiBjaGlsZHJlbiDrs4DtmZgg4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSACmZ1bmN0aW9uIHBhcnNlTUQobWQsIF9zb3VyY2VQYXRoKSB7CiAgY29uc3QgY2hpbGRyZW4gPSBbXTsKICBjb25zdCBsaW5lcyAgICA9IG1kLnNwbGl0KCJcbiIpOwogIGxldCBpID0gMCwgaW5Db2RlID0gZmFsc2UsIGNvZGVMaW5lcyA9IFtdOwogIGxldCB0aGRycyA9IG51bGwsIHRyb3dzID0gW107CgogIGZ1bmN0aW9uIGZsdXNoVGFibGUoKSB7CiAgICBpZiAoIXRoZHJzKSByZXR1cm47CiAgICBjaGlsZHJlbi5wdXNoKG1ha2VUYWJsZSh0aGRycywgdHJvd3MpKTsKICAgIGNoaWxkcmVuLnB1c2gobmV3IFBhcmFncmFwaCh7IHNwYWNpbmc6IHsgYWZ0ZXI6IDEyMCB9LCBjaGlsZHJlbjogW10gfSkpOwogICAgdGhkcnMgPSBudWxsOyB0cm93cyA9IFtdOwogIH0KCiAgd2hpbGUgKGkgPCBsaW5lcy5sZW5ndGgpIHsKICAgIGNvbnN0IGxpbmUgPSBsaW5lc1tpXSwgcyA9IGxpbmUudHJpbSgpOwoKICAgIC8vIOy9lOuTnCDruJTroZ0KICAgIGlmIChzLnN0YXJ0c1dpdGgoImBgYCIpKSB7CiAgICAgIGlmICghaW5Db2RlKSB7IGluQ29kZSA9IHRydWU7IGNvZGVMaW5lcyA9IFtdOyBpKys7IGNvbnRpbnVlOyB9CiAgICAgIGluQ29kZSA9IGZhbHNlOwogICAgICBjaGlsZHJlbi5wdXNoKG5ldyBQYXJhZ3JhcGgoewogICAgICAgIHNwYWNpbmc6IHsgYmVmb3JlOiA4MCwgYWZ0ZXI6IDgwIH0sCiAgICAgICAgc2hhZGluZzogeyB0eXBlOiBTaGFkaW5nVHlwZS5DTEVBUiwgZmlsbDogIkYzRjRGNiIgfSwKICAgICAgICBib3JkZXI6ICBib3JkZXIoQy5taWRCbHVlLCAyKSwKICAgICAgICBpbmRlbnQ6ICB7IGxlZnQ6IDI0MCB9LAogICAgICAgIGNoaWxkcmVuOiBbbmV3IFRleHRSdW4oeyB0ZXh0OiBjb2RlTGluZXMuam9pbigiXG4iKSwgZm9udDogIkNvbnNvbGFzIiwgc2l6ZTogMTgsIGNvbG9yOiAiMUYyOTM3IiB9KV0sCiAgICAgIH0pKTsKICAgICAgaSsrOyBjb250aW51ZTsKICAgIH0KICAgIGlmIChpbkNvZGUpIHsgY29kZUxpbmVzLnB1c2gobGluZSk7IGkrKzsgY29udGludWU7IH0KCiAgICAvLyDtl6TrlKkKICAgIGNvbnN0IGhtID0gcy5tYXRjaCgvXigjezEsNH0pXHMrKC4rKS8pOwogICAgaWYgKGhtKSB7IGZsdXNoVGFibGUoKTsgY2hpbGRyZW4ucHVzaChtYWtlSGVhZGluZyhobVsyXSwgaG1bMV0ubGVuZ3RoKSk7IGkrKzsgY29udGludWU7IH0KCiAgICAvLyDtkZwKICAgIGlmIChzLnN0YXJ0c1dpdGgoInwiKSkgewogICAgICBjb25zdCBjZWxscyA9IHMuc3BsaXQoInwiKS5zbGljZSgxLCAtMSkubWFwKGMgPT4gYy50cmltKCkpOwogICAgICBpZiAoY2VsbHMuZXZlcnkoYyA9PiAvXlstOiBdKyQvLnRlc3QoYykpKSB7IGkrKzsgY29udGludWU7IH0KICAgICAgaWYgKCF0aGRycykgdGhkcnMgPSBjZWxsczsgZWxzZSB0cm93cy5wdXNoKGNlbGxzKTsKICAgICAgaSsrOyBjb250aW51ZTsKICAgIH0gZWxzZSB7IGZsdXNoVGFibGUoKTsgfQoKICAgIC8vIOydtOuvuOyngCDrtoTshJ0g7L2c7JWE7JuDCiAgICBpZiAocy5zdGFydHNXaXRoKCIqKlvsnbTrr7jsp4AiKSB8fCBzLnN0YXJ0c1dpdGgoIioqW1Zpc2lvbiIpKSB7CiAgICAgIGNoaWxkcmVuLnB1c2gobWFrZUNhbGxvdXQocykpOyBpKys7IGNvbnRpbnVlOwogICAgfQoKICAgIC8vIOq4gOuouOumrAogICAgY29uc3QgYm0gPSBzLm1hdGNoKC9eWy0qK11ccysoLispLyk7CiAgICBpZiAoYm0pIHsKICAgICAgY2hpbGRyZW4ucHVzaChuZXcgUGFyYWdyYXBoKHsKICAgICAgICBzcGFjaW5nOiB7IGJlZm9yZTogNDAsIGFmdGVyOiA0MCB9LAogICAgICAgIGluZGVudDogIHsgbGVmdDogNDgwLCBoYW5naW5nOiAyNDAgfSwKICAgICAgICBjaGlsZHJlbjogW25ldyBUZXh0UnVuKHsgdGV4dDogIuKAoiAgIiwgc2l6ZTogMjIsIGZvbnQ6ICLrp5HsnYAg6rOg65SVIiwgY29sb3I6IEMuYmx1ZSwgYm9sZDogdHJ1ZSB9KSwgLi4ucGFyc2VJbmxpbmUoYm1bMV0pXSwKICAgICAgfSkpOwogICAgICBpKys7IGNvbnRpbnVlOwogICAgfQoKICAgIC8vIOuyiO2YuCDrqqnroZ0KICAgIGNvbnN0IG5tID0gcy5tYXRjaCgvXihcZCspXC5ccysoLispLyk7CiAgICBpZiAobm0pIHsKICAgICAgY2hpbGRyZW4ucHVzaChuZXcgUGFyYWdyYXBoKHsKICAgICAgICBzcGFjaW5nOiB7IGJlZm9yZTogNDAsIGFmdGVyOiA0MCB9LAogICAgICAgIGluZGVudDogIHsgbGVmdDogNDgwLCBoYW5naW5nOiAyODAgfSwKICAgICAgICBjaGlsZHJlbjogW25ldyBUZXh0UnVuKHsgdGV4dDogYCR7bm1bMV19LiAgYCwgc2l6ZTogMjIsIGZvbnQ6ICLrp5HsnYAg6rOg65SVIiwgY29sb3I6IEMuYmx1ZSwgYm9sZDogdHJ1ZSB9KSwgLi4ucGFyc2VJbmxpbmUobm1bMl0pXSwKICAgICAgfSkpOwogICAgICBpKys7IGNvbnRpbnVlOwogICAgfQoKICAgIC8vIOyduOyaqQogICAgY29uc3QgcW0gPSBzLm1hdGNoKC9ePlxzKiguKikvKTsKICAgIGlmIChxbSkgewogICAgICBjaGlsZHJlbi5wdXNoKG5ldyBQYXJhZ3JhcGgoewogICAgICAgIHNwYWNpbmc6IHsgYmVmb3JlOiA4MCwgYWZ0ZXI6IDgwIH0sIGluZGVudDogeyBsZWZ0OiA0ODAgfSwKICAgICAgICBib3JkZXI6ICB7IGxlZnQ6IHsgc3R5bGU6IEJvcmRlclN0eWxlLlNJTkdMRSwgc2l6ZTogMTYsIGNvbG9yOiBDLmJsdWUsIHNwYWNlOiA4IH0gfSwKICAgICAgICBzaGFkaW5nOiB7IHR5cGU6IFNoYWRpbmdUeXBlLkNMRUFSLCBmaWxsOiBDLmxpZ2h0Qmx1ZSB9LAogICAgICAgIGNoaWxkcmVuOiBwYXJzZUlubGluZShxbVsxXSwgMjApLAogICAgICB9KSk7CiAgICAgIGkrKzsgY29udGludWU7CiAgICB9CgogICAgLy8gSFIKICAgIGlmICgvXi17Myx9JC8udGVzdChzKSkgewogICAgICBjaGlsZHJlbi5wdXNoKG5ldyBQYXJhZ3JhcGgoewogICAgICAgIGJvcmRlcjogIHsgYm90dG9tOiB7IHN0eWxlOiBCb3JkZXJTdHlsZS5TSU5HTEUsIHNpemU6IDYsIGNvbG9yOiBDLm1pZEJsdWUgfSB9LAogICAgICAgIHNwYWNpbmc6IHsgYmVmb3JlOiAxMjAsIGFmdGVyOiAxMjAgfSwKICAgICAgICBjaGlsZHJlbjogW10sCiAgICAgIH0pKTsKICAgICAgaSsrOyBjb250aW51ZTsKICAgIH0KCiAgICAvLyDruYgg7KSECiAgICBpZiAoIXMpIHsKICAgICAgY2hpbGRyZW4ucHVzaChuZXcgUGFyYWdyYXBoKHsgc3BhY2luZzogeyBiZWZvcmU6IDIwLCBhZnRlcjogMjAgfSwgY2hpbGRyZW46IFtdIH0pKTsKICAgICAgaSsrOyBjb250aW51ZTsKICAgIH0KCiAgICAvLyDsnbTrr7jsp4Ag7J6E67Kg65OcIOuniOy7pAogICAgY29uc3QgZW1tID0gcy5tYXRjaCgvXlxbRU1CRURfSU1BR0U6KC4rKVxdJC8pOwogICAgaWYgKGVtbSkgewogICAgICBjb25zdCBpbWdQYXJhID0gbWFrZUltYWdlUGFyYWdyYXBoKGVtbVsxXSwgX3NvdXJjZVBhdGggfHwgIiIpOwogICAgICBpZiAoaW1nUGFyYSkgY2hpbGRyZW4ucHVzaChpbWdQYXJhKTsKICAgICAgaSsrOyBjb250aW51ZTsKICAgIH0KCiAgICAvLyDsnbzrsJgg66y464uoCiAgICBjaGlsZHJlbi5wdXNoKG5ldyBQYXJhZ3JhcGgoeyBzcGFjaW5nOiB7IGJlZm9yZTogNjAsIGFmdGVyOiA2MCB9LCBjaGlsZHJlbjogcGFyc2VJbmxpbmUocykgfSkpOwogICAgaSsrOwogIH0KICBmbHVzaFRhYmxlKCk7CiAgcmV0dXJuIGNoaWxkcmVuOwp9CgovLyDilIDilIDilIAg7Luk67KEIO2OmOydtOyngCDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKZnVuY3Rpb24gbWFrZUNvdmVyKHRpdGxlLCBzdWJ0aXRsZSwgZGF0ZSkgewogIHJldHVybiBbCiAgICBuZXcgUGFyYWdyYXBoKHsgc3BhY2luZzogeyBiZWZvcmU6IDE4MDAgfSwgY2hpbGRyZW46IFtdIH0pLAogICAgbmV3IFBhcmFncmFwaCh7CiAgICAgIGJvcmRlcjogeyB0b3A6IHsgc3R5bGU6IEJvcmRlclN0eWxlLlNJTkdMRSwgc2l6ZTogNDAsIGNvbG9yOiBDLmJsdWUgfSB9LAogICAgICBzcGFjaW5nOiB7IGFmdGVyOiA4MDAgfSwgY2hpbGRyZW46IFtdLAogICAgfSksCiAgICBuZXcgUGFyYWdyYXBoKHsKICAgICAgYWxpZ25tZW50OiBBbGlnbm1lbnRUeXBlLkNFTlRFUiwgc3BhY2luZzogeyBhZnRlcjogMjQwIH0sCiAgICAgIGNoaWxkcmVuOiBbbmV3IFRleHRSdW4oeyB0ZXh0OiB0aXRsZSwgYm9sZDogdHJ1ZSwgc2l6ZTogNzIsIGNvbG9yOiBDLmJsdWUsIGZvbnQ6ICLrp5HsnYAg6rOg65SVIiB9KV0sCiAgICB9KSwKICAgIG5ldyBQYXJhZ3JhcGgoewogICAgICBhbGlnbm1lbnQ6IEFsaWdubWVudFR5cGUuQ0VOVEVSLCBzcGFjaW5nOiB7IGFmdGVyOiA2MDAgfSwKICAgICAgY2hpbGRyZW46IFtuZXcgVGV4dFJ1bih7IHRleHQ6IHN1YnRpdGxlLCBzaXplOiAzMiwgY29sb3I6IEMuZ3JheSwgZm9udDogIuunkeydgCDqs6DrlJUiIH0pXSwKICAgIH0pLAogICAgbmV3IFBhcmFncmFwaCh7CiAgICAgIGJvcmRlcjogeyBib3R0b206IHsgc3R5bGU6IEJvcmRlclN0eWxlLlNJTkdMRSwgc2l6ZTogOCwgY29sb3I6IEMubWlkQmx1ZSB9IH0sCiAgICAgIHNwYWNpbmc6IHsgYWZ0ZXI6IDQwMCB9LCBjaGlsZHJlbjogW10sCiAgICB9KSwKICAgIG5ldyBQYXJhZ3JhcGgoewogICAgICBhbGlnbm1lbnQ6IEFsaWdubWVudFR5cGUuQ0VOVEVSLCBzcGFjaW5nOiB7IGFmdGVyOiAyMDAgfSwKICAgICAgY2hpbGRyZW46IFtuZXcgVGV4dFJ1bih7IHRleHQ6IGRhdGUsIHNpemU6IDI2LCBjb2xvcjogQy5ncmF5LCBmb250OiAi66eR7J2AIOqzoOuUlSIgfSldLAogICAgfSksCiAgICBuZXcgUGFyYWdyYXBoKHsKICAgICAgYWxpZ25tZW50OiBBbGlnbm1lbnRUeXBlLkNFTlRFUiwKICAgICAgY2hpbGRyZW46IFtuZXcgVGV4dFJ1bih7IHRleHQ6ICJTYW1zdW5nIENvbmZpZGVudGlhbCIsIHNpemU6IDIyLCBjb2xvcjogQy5yZWQsIGJvbGQ6IHRydWUsIGZvbnQ6ICLrp5HsnYAg6rOg65SVIiB9KV0sCiAgICB9KSwKICAgIG5ldyBQYXJhZ3JhcGgoeyBjaGlsZHJlbjogW25ldyBQYWdlQnJlYWsoKV0gfSksCiAgXTsKfQoKLy8g4pSA4pSA4pSAIO2XpOuNlCAvIO2RuO2EsCDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKZnVuY3Rpb24gbWFrZUhlYWRlcih0aXRsZSkgewogIHJldHVybiBuZXcgSGVhZGVyKHsgY2hpbGRyZW46IFtuZXcgUGFyYWdyYXBoKHsKICAgIGJvcmRlcjogIHsgYm90dG9tOiB7IHN0eWxlOiBCb3JkZXJTdHlsZS5TSU5HTEUsIHNpemU6IDQsIGNvbG9yOiBDLmJsdWUsIHNwYWNlOiAxIH0gfSwKICAgIHNwYWNpbmc6IHsgYWZ0ZXI6IDEwMCB9LAogICAgdGFiU3RvcHM6IFt7IHR5cGU6ICJyaWdodCIsIHBvc2l0aW9uOiA5MDI2IH1dLAogICAgY2hpbGRyZW46IFsKICAgICAgbmV3IFRleHRSdW4oeyB0ZXh0OiB0aXRsZSwgICAgICAgICAgICAgICAgICAgc2l6ZTogMTgsIGZvbnQ6ICLrp5HsnYAg6rOg65SVIiwgY29sb3I6IEMuYmx1ZSwgYm9sZDogdHJ1ZSB9KSwKICAgICAgbmV3IFRleHRSdW4oeyB0ZXh0OiAiXHRTYW1zdW5nIENvbmZpZGVudGlhbCIsIHNpemU6IDE4LCBmb250OiAi66eR7J2AIOqzoOuUlSIsIGNvbG9yOiBDLmdyYXkgfSksCiAgICBdLAogIH0pXSB9KTsKfQoKZnVuY3Rpb24gbWFrZUZvb3RlcihkYXRlKSB7CiAgcmV0dXJuIG5ldyBGb290ZXIoeyBjaGlsZHJlbjogW25ldyBQYXJhZ3JhcGgoewogICAgYm9yZGVyOiAgICB7IHRvcDogeyBzdHlsZTogQm9yZGVyU3R5bGUuU0lOR0xFLCBzaXplOiA0LCBjb2xvcjogQy5taWRCbHVlLCBzcGFjZTogMSB9IH0sCiAgICBzcGFjaW5nOiAgIHsgYmVmb3JlOiA4MCB9LAogICAgYWxpZ25tZW50OiBBbGlnbm1lbnRUeXBlLkNFTlRFUiwKICAgIGNoaWxkcmVuOiBbCiAgICAgIG5ldyBUZXh0UnVuKHsgdGV4dDogYCR7ZGF0ZX0gIMK3ICBgLCBzaXplOiAxOCwgZm9udDogIuunkeydgCDqs6DrlJUiLCBjb2xvcjogQy5ncmF5IH0pLAogICAgICBuZXcgVGV4dFJ1bih7IGNoaWxkcmVuOiBbUGFnZU51bWJlci5DVVJSRU5UXSwgICAgIHNpemU6IDE4LCBmb250OiAi66eR7J2AIOqzoOuUlSIsIGNvbG9yOiBDLmdyYXkgfSksCiAgICAgIG5ldyBUZXh0UnVuKHsgdGV4dDogIiAvICIsICAgICAgICAgICAgICAgICAgICAgICAgIHNpemU6IDE4LCBmb250OiAi66eR7J2AIOqzoOuUlSIsIGNvbG9yOiBDLmdyYXkgfSksCiAgICAgIG5ldyBUZXh0UnVuKHsgY2hpbGRyZW46IFtQYWdlTnVtYmVyLlRPVEFMX1BBR0VTXSwgc2l6ZTogMTgsIGZvbnQ6ICLrp5HsnYAg6rOg65SVIiwgY29sb3I6IEMuZ3JheSB9KSwKICAgIF0sCiAgfSldIH0pOwp9CgovLyDilIDilIDilIAg7Y6Y7J207KeAIOyEpOyglSDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKY29uc3QgUEFHRV9QUk9QUyA9IHsKICBwYWdlOiB7CiAgICBzaXplOiAgIHsgd2lkdGg6IDExOTA2LCBoZWlnaHQ6IDE2ODM4IH0sICAgICAgICAgIC8vIEE0CiAgICBtYXJnaW46IHsgdG9wOiAxNDQwLCByaWdodDogMTQ0MCwgYm90dG9tOiAxNDQwLCBsZWZ0OiAxODAwIH0sCiAgfSwKfTsKCi8vIOKUgOKUgOKUgCDrrLjshJwg7Iqk7YOA7J28IOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgApmdW5jdGlvbiBkb2NTdHlsZXMoKSB7CiAgcmV0dXJuIHsKICAgIGRlZmF1bHQ6IHsgZG9jdW1lbnQ6IHsgcnVuOiB7IGZvbnQ6ICLrp5HsnYAg6rOg65SVIiwgc2l6ZTogMjIgfSB9IH0sCiAgICBwYXJhZ3JhcGhTdHlsZXM6IFsKICAgICAgeyBpZDogIkhlYWRpbmcxIiwgbmFtZTogIkhlYWRpbmcgMSIsIGJhc2VkT246ICJOb3JtYWwiLCBuZXh0OiAiTm9ybWFsIiwgcXVpY2tGb3JtYXQ6IHRydWUsCiAgICAgICAgcnVuOiB7IHNpemU6IDM2LCBib2xkOiB0cnVlLCBmb250OiAi66eR7J2AIOqzoOuUlSIsIGNvbG9yOiBDLmJsdWUgfSwKICAgICAgICBwYXJhZ3JhcGg6IHsgc3BhY2luZzogeyBiZWZvcmU6IDQwMCwgYWZ0ZXI6IDIwMCB9LCBvdXRsaW5lTGV2ZWw6IDAgfSB9LAogICAgICB7IGlkOiAiSGVhZGluZzIiLCBuYW1lOiAiSGVhZGluZyAyIiwgYmFzZWRPbjogIk5vcm1hbCIsIG5leHQ6ICJOb3JtYWwiLCBxdWlja0Zvcm1hdDogdHJ1ZSwKICAgICAgICBydW46IHsgc2l6ZTogMzAsIGJvbGQ6IHRydWUsIGZvbnQ6ICLrp5HsnYAg6rOg65SVIiwgY29sb3I6IEMuYmx1ZSB9LAogICAgICAgIHBhcmFncmFwaDogeyBzcGFjaW5nOiB7IGJlZm9yZTogMzAwLCBhZnRlcjogMTYwIH0sIG91dGxpbmVMZXZlbDogMSB9IH0sCiAgICAgIHsgaWQ6ICJIZWFkaW5nMyIsIG5hbWU6ICJIZWFkaW5nIDMiLCBiYXNlZE9uOiAiTm9ybWFsIiwgbmV4dDogIk5vcm1hbCIsIHF1aWNrRm9ybWF0OiB0cnVlLAogICAgICAgIHJ1bjogeyBzaXplOiAyNiwgYm9sZDogdHJ1ZSwgZm9udDogIuunkeydgCDqs6DrlJUiLCBjb2xvcjogIjMzMzMzMyIgfSwKICAgICAgICBwYXJhZ3JhcGg6IHsgc3BhY2luZzogeyBiZWZvcmU6IDI0MCwgYWZ0ZXI6IDEyMCB9LCBvdXRsaW5lTGV2ZWw6IDIgfSB9LAogICAgXSwKICB9Owp9CgovLyDilIDilIDilIAg7J2066+47KeAIOyCveyehSDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKZnVuY3Rpb24gbWFrZUltYWdlUGFyYWdyYXBoKGltZ1BhdGgsIGJhc2VQYXRoKSB7CiAgdHJ5IHsKICAgIC8vIGltZ1BhdGg6ICJpbWFnZXMvaW1nXzAwMS5wbmciICjsg4HrjIDqsr3roZwpCiAgICBjb25zdCBmdWxsUGF0aCA9IHBhdGguaXNBYnNvbHV0ZShpbWdQYXRoKQogICAgICA/IGltZ1BhdGgKICAgICAgOiBwYXRoLmpvaW4ocGF0aC5kaXJuYW1lKGJhc2VQYXRoKSwgaW1nUGF0aCk7CiAgICBpZiAoIWZzLmV4aXN0c1N5bmMoZnVsbFBhdGgpKSByZXR1cm4gbnVsbDsKCiAgICBjb25zdCBpbWdEYXRhICA9IGZzLnJlYWRGaWxlU3luYyhmdWxsUGF0aCk7CiAgICBjb25zdCBleHQgICAgICA9IHBhdGguZXh0bmFtZShmdWxsUGF0aCkudG9Mb3dlckNhc2UoKS5yZXBsYWNlKCIuIiwgIiIpOwogICAgY29uc3QgdHlwZU1hcCAgPSB7IGpwZzogImpwZyIsIGpwZWc6ICJqcGciLCBwbmc6ICJwbmciLCBnaWY6ICJnaWYiLCBibXA6ICJibXAiLCB3ZWJwOiAicG5nIiB9OwogICAgY29uc3QgaW1nVHlwZSAgPSB0eXBlTWFwW2V4dF0gfHwgInBuZyI7CgogICAgLy8g7JuQ67O4IO2BrOq4sCDstpTsoJUg7JeG7J20IOy1nOuMgCDtj60g6riw7KSA7Jy866GcIOyCveyehSAoQTQgY29udGVudCDtj60g6riw7KSAKQogICAgY29uc3QgbWF4VyA9IDUwMCwgbWF4SCA9IDQwMDsKICAgIHJldHVybiBuZXcgUGFyYWdyYXBoKHsKICAgICAgYWxpZ25tZW50OiBBbGlnbm1lbnRUeXBlLkNFTlRFUiwKICAgICAgc3BhY2luZzogeyBiZWZvcmU6IDEyMCwgYWZ0ZXI6IDEyMCB9LAogICAgICBjaGlsZHJlbjogW25ldyBJbWFnZVJ1bih7CiAgICAgICAgdHlwZTogaW1nVHlwZSwKICAgICAgICBkYXRhOiBpbWdEYXRhLAogICAgICAgIHRyYW5zZm9ybWF0aW9uOiB7IHdpZHRoOiBtYXhXLCBoZWlnaHQ6IG1heEggfSwKICAgICAgICBhbHRUZXh0OiB7IHRpdGxlOiAi7J2066+47KeAIiwgZGVzY3JpcHRpb246IGltZ1BhdGgsIG5hbWU6IGltZ1BhdGggfSwKICAgICAgfSldLAogICAgfSk7CiAgfSBjYXRjaCAoZSkgewogICAgcmV0dXJuIG51bGw7CiAgfQp9CgovLyDilIDilIDilIAg7IaM7IqkIOuhnOuTnCAo7YyM7J28IG9yIO2PtOuNlCkg4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSACmZ1bmN0aW9uIGxvYWRTb3VyY2VzKHNyYykgewogIGNvbnN0IHN0YXQgPSBmcy5zdGF0U3luYyhzcmMpOwogIGlmIChzdGF0LmlzRmlsZSgpKSB7CiAgICByZXR1cm4gW3sgdGl0bGU6IHBhdGguYmFzZW5hbWUoc3JjLCAiLm1kIiksIG1kOiBmcy5yZWFkRmlsZVN5bmMoc3JjLCAidXRmOCIpLCBwYXRoOiBzcmMgfV07CiAgfQogIC8vIO2PtOuNlDogLm1kIO2MjOydvCDsoJXroKwg66Gc65OcICjrs7Tqs6DshJwvcmVwb3J0IO2MjOydvCDsoJzsmbgpCiAgcmV0dXJuIGZzLnJlYWRkaXJTeW5jKHNyYykKICAgIC5maWx0ZXIoZiA9PiBmLmVuZHNXaXRoKCIubWQiKSAmJiAhL+uztOqzoOyEnHxyZXBvcnQvaS50ZXN0KGYpKQogICAgLnNvcnQoKQogICAgLm1hcChmID0+ICh7CiAgICAgIHRpdGxlOiBmLnJlcGxhY2UoIi5tZCIsICIiKSwKICAgICAgbWQ6ICAgIGZzLnJlYWRGaWxlU3luYyhwYXRoLmpvaW4oc3JjLCBmKSwgInV0ZjgiKSwKICAgICAgcGF0aDogIHBhdGguam9pbihzcmMsIGYpLAogICAgfSkpOwp9CgovLyDilIDilIDilIAg66mU7J24IOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgAphc3luYyBmdW5jdGlvbiBtYWluKCkgewogIGNvbnN0IFssLCBzcmMsIG91dFBhdGgsIHRpdGxlQXJnXSA9IHByb2Nlc3MuYXJndjsKICBpZiAoIXNyYyB8fCAhb3V0UGF0aCkgewogICAgY29uc29sZS5lcnJvcigiVXNhZ2U6IG5vZGUgbWFrZV9yZXBvcnQuanMgPGZpbGUubWR8Zm9sZGVyPiA8b3V0cHV0LmRvY3g+IFt0aXRsZV0iKTsKICAgIHByb2Nlc3MuZXhpdCgxKTsKICB9CgogIGNvbnN0IHNvdXJjZXMgID0gbG9hZFNvdXJjZXMoc3JjKTsKICBjb25zdCB0aXRsZSAgICA9IHRpdGxlQXJnIHx8IHBhdGguYmFzZW5hbWUoc3JjLCAiLm1kIikgfHwgIkNvbmZsdWVuY2Ug67O06rOg7IScIjsKICBjb25zdCBkYXRlICAgICA9IG5ldyBEYXRlKCkudG9Mb2NhbGVEYXRlU3RyaW5nKCJrby1LUiIsIHsgeWVhcjogIm51bWVyaWMiLCBtb250aDogImxvbmciLCBkYXk6ICJudW1lcmljIiB9KTsKICBjb25zdCBzdWJ0aXRsZSA9IGBDb25mbHVlbmNlIOyekOuPmSDsiJjsp5Eg67O06rOg7IScICDCtyAg7LSdICR7c291cmNlcy5sZW5ndGh96rCcIO2OmOydtOyngGA7CgogIGNvbnNvbGUubG9nKGDwn5OEIOyGjOyKpCAke3NvdXJjZXMubGVuZ3RofeqwnCDroZzrk5wg7JmE66OMYCk7CgogIC8vIOuzuOusuCBjaGlsZHJlbiDqtazshLEKICBjb25zdCBib2R5Q2hpbGRyZW4gPSBbXTsKICBzb3VyY2VzLmZvckVhY2goKHMsIGlkeCkgPT4gewogICAgaWYgKGlkeCA+IDApIGJvZHlDaGlsZHJlbi5wdXNoKG5ldyBQYXJhZ3JhcGgoeyBjaGlsZHJlbjogW25ldyBQYWdlQnJlYWsoKV0gfSkpOwogICAgaWYgKHNvdXJjZXMubGVuZ3RoID4gMSkgYm9keUNoaWxkcmVuLnB1c2gobWFrZVNvdXJjZUJhbm5lcihzLnRpdGxlKSk7CiAgICBib2R5Q2hpbGRyZW4ucHVzaCguLi5wYXJzZU1EKHMubWQsIHMucGF0aCkpOwogIH0pOwoKICBjb25zdCBkb2MgPSBuZXcgRG9jdW1lbnQoewogICAgc3R5bGVzOiAgIGRvY1N0eWxlcygpLAogICAgc2VjdGlvbnM6IFsKICAgICAgLy8g4pSA4pSAIOyEueyFmDE6IOy7pOuyhCArIOuqqeywqCDilIDilIAKICAgICAgewogICAgICAgIHByb3BlcnRpZXM6IHsgcGFnZTogUEFHRV9QUk9QUy5wYWdlIH0sCiAgICAgICAgY2hpbGRyZW46IFsKICAgICAgICAgIC4uLm1ha2VDb3Zlcih0aXRsZSwgc3VidGl0bGUsIGRhdGUpLAogICAgICAgICAgbWFrZUhlYWRpbmcoIuuqqSAg7LCoIiwgMSksCiAgICAgICAgICBuZXcgVGFibGVPZkNvbnRlbnRzKCLrqqnssKgiLCB7IGh5cGVybGluazogdHJ1ZSwgaGVhZGluZ1N0eWxlUmFuZ2U6ICIxLTMiIH0pLAogICAgICAgICAgbmV3IFBhcmFncmFwaCh7IGNoaWxkcmVuOiBbbmV3IFBhZ2VCcmVhaygpXSB9KSwKICAgICAgICBdLAogICAgICB9LAogICAgICAvLyDshLnshZgyOiDrs7jrrLgKICAgICAgewogICAgICAgIHByb3BlcnRpZXM6IHsgcGFnZTogUEFHRV9QUk9QUy5wYWdlIH0sCiAgICAgICAgaGVhZGVyczogICAgeyBkZWZhdWx0OiBtYWtlSGVhZGVyKHRpdGxlKSAgfSwKICAgICAgICBmb290ZXJzOiAgICB7IGRlZmF1bHQ6IG1ha2VGb290ZXIoZGF0ZSkgICB9LAogICAgICAgIGNoaWxkcmVuOiAgIGJvZHlDaGlsZHJlbiwKICAgICAgfSwKICAgIF0sCiAgfSk7CgogIGNvbnN0IGJ1ZiA9IGF3YWl0IFBhY2tlci50b0J1ZmZlcihkb2MpOwogIGZzLndyaXRlRmlsZVN5bmMob3V0UGF0aCwgYnVmKTsKICBjb25zb2xlLmxvZyhg4pyFIFdvcmQg67O06rOg7IScIOyDneyEsSDsmYTro4w6ICR7b3V0UGF0aH1gKTsKfQoKbWFpbigpLmNhdGNoKGUgPT4geyBjb25zb2xlLmVycm9yKCLinYwiLCBlLm1lc3NhZ2UpOyBwcm9jZXNzLmV4aXQoMSk7IH0pOwo="

def _node_available():
    return _shutil.which("node") is not None

def _ensure_docx_npm(work_dir, callback=None):
    nm = os.path.join(work_dir, "node_modules", "docx")
    if os.path.isdir(nm):
        return True
    if callback: callback("[Word] npm install docx 실행 중 (최초 1회)...")
    try:
        r = _subprocess.run(["npm", "install", "docx"], cwd=work_dir,
                             capture_output=True, timeout=120)
        ok = os.path.isdir(nm)
        if not ok and callback: callback(f"[Word] npm install 실패: {r.stderr[:200]}")
        return ok
    except Exception as e:
        if callback: callback(f"[Word] npm install 오류: {e}")
        return False

_WORD_WORK_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_word_engine")

def generate_word_report(src, out_docx, title="", callback=None):
    """src: MD 파일 경로 or MD 폴더 경로. 실패시 빈 문자열 반환."""
    if not _node_available():
        if callback: callback("[Word] node.js 미설치. https://nodejs.org 에서 설치 후 재시도.")
        return ""
    os.makedirs(_WORD_WORK_DIR, exist_ok=True)
    js_path = os.path.join(_WORD_WORK_DIR, "make_report.js")
    js_txt  = _base64.b64decode(_MAKE_REPORT_JS_B64).decode("utf-8")
    try:
        existing = open(js_path, encoding="utf-8").read() if os.path.exists(js_path) else ""
    except Exception:
        existing = ""
    if existing != js_txt:
        with open(js_path, "w", encoding="utf-8") as fjs:
            fjs.write(js_txt)
    if not _ensure_docx_npm(_WORD_WORK_DIR, callback):
        return ""
    try:
        cmd = ["node", js_path, src, out_docx] + ([title] if title else [])
        if callback: callback(f"[Word] 생성 중... → {os.path.basename(out_docx)}")
        r = _subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        if r.stdout.strip() and callback: callback(r.stdout.strip())
        if r.returncode != 0:
            if callback: callback(f"[Word] 오류: {r.stderr.strip()[:300]}")
            return ""
        if callback: callback(f"[Word] ✅ 완료: {out_docx}")
        return out_docx
    except Exception as e:
        if callback: callback(f"[Word] 생성 실패: {e}")
        return ""


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
        self.rowconfigure(10, weight=1)

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
        ttk.Checkbutton(self, text="이미지 Vision 분석", variable=self.vision_var).grid(row=3, column=1, sticky=tk.W, pady=4)
        ttk.Checkbutton(self, text="페이지별 LLM 요약 생성", variable=self.summary_var).grid(row=4, column=1, sticky=tk.W, pady=4)
        self.word_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self, text="수집 완료 후 Word 보고서 자동 생성 (.docx)", variable=self.word_var).grid(row=5, column=1, sticky=tk.W, pady=4)

        ttk.Label(self, text="출력 폴더:").grid(row=6, column=0, sticky=tk.W, pady=8)
        self.out_var = tk.StringVar(value="./confluence_output")
        frm = ttk.Frame(self, style="Card.TFrame")
        frm.grid(row=5, column=1, columnspan=2, sticky=tk.EW)
        frm.columnconfigure(0, weight=1)
        ttk.Entry(frm, textvariable=self.out_var).grid(row=0, column=0, sticky=tk.EW)
        ttk.Button(frm, text="찾아보기", command=self._browse).grid(row=0, column=1, padx=8)

        self.btn = ttk.Button(self, text="수집 시작", command=self._run)
        self.btn.grid(row=7, column=0, pady=12)
        self.prog = ttk.Progressbar(self, mode="indeterminate", length=700)
        self.prog.grid(row=8, column=0, columnspan=3, sticky=tk.EW)

        ttk.Label(self, text="로그:", style="Title.TLabel").grid(row=9, column=0, sticky=tk.W, pady=8)
        lf = ttk.Frame(self, style="Card.TFrame")
        lf.grid(row=10, column=0, columnspan=3, sticky=tk.NSEW)
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

                # Word 보고서 자동 생성
                if self.word_var.get():
                    page_title_raw = self.url_var.get().split("/")[-1].replace("+", " ")
                    docx_path = os.path.join(self.out_var.get(), f"보고서_{page_id}.docx")
                    generate_word_report(save_dir, docx_path,
                                         title=page_title_raw or f"Confluence 보고서",
                                         callback=self.log)
                    if os.path.exists(docx_path):
                        messagebox.showinfo("완료", f"수집 및 Word 보고서 생성 완료!\n\n📁 MD: {os.path.abspath(save_dir)}\n📄 Word: {os.path.abspath(docx_path)}")
                    else:
                        messagebox.showinfo("완료", f"수집 완료!\n{os.path.abspath(save_dir)}")
                else:
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
        self.rowconfigure(7, weight=1)

        # ── 상단 버튼 바 ──────────────────────────────────
        top = ttk.Frame(self, style="Card.TFrame")
        top.grid(row=0, column=0, columnspan=2, sticky=tk.EW, pady=(0, 6))
        ttk.Label(top, text="파일 관리:", style="Title.TLabel").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(top, text="📂 폴더 추가",    command=self._load_folder).pack(side=tk.LEFT, padx=3)
        ttk.Button(top, text="📄 파일 추가",    command=self._load_files).pack(side=tk.LEFT, padx=3)
        ttk.Separator(top, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=6, pady=3)
        ttk.Button(top, text="✅ 전체 선택",    command=self._check_all).pack(side=tk.LEFT, padx=3)
        ttk.Button(top, text="☐ 전체 해제",    command=self._uncheck_all).pack(side=tk.LEFT, padx=3)
        ttk.Separator(top, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=6, pady=3)
        ttk.Button(top, text="🗑 선택 제거",    command=self._remove_selected).pack(side=tk.LEFT, padx=3)
        ttk.Button(top, text="🧹 목록 초기화",  command=self._clear_all).pack(side=tk.LEFT, padx=3)
        self.file_count_lbl = ttk.Label(top, text="0개 파일", style="Muted.TLabel")
        self.file_count_lbl.pack(side=tk.RIGHT, padx=8)

        # ── 파일 목록 (Treeview + 체크박스) ───────────────
        lf = ttk.LabelFrame(self, text="파일 목록  (행 클릭 = 체크 토글 / 체크된 파일만 보고서에 포함)")
        lf.grid(row=1, column=0, columnspan=2, sticky=tk.NSEW, pady=4)
        lf.columnconfigure(0, weight=1)
        lf.rowconfigure(0, weight=1)

        cols = ("check", "filename", "folder")
        self.file_list = ttk.Treeview(lf, columns=cols, show="headings", height=12,
                                       selectmode="extended")
        self.file_list.heading("check",    text="포함")
        self.file_list.heading("filename", text="파일명")
        self.file_list.heading("folder",   text="경로")
        self.file_list.column("check",    width=50,  minwidth=50,  stretch=False, anchor="center")
        self.file_list.column("filename", width=260, minwidth=120, stretch=True)
        self.file_list.column("folder",   width=320, minwidth=120, stretch=True)
        sb_y = ttk.Scrollbar(lf, orient="vertical",   command=self.file_list.yview)
        sb_x = ttk.Scrollbar(lf, orient="horizontal", command=self.file_list.xview)
        self.file_list.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        self.file_list.grid(row=0, column=0, sticky=tk.NSEW)
        sb_y.grid(row=0, column=1, sticky=tk.NS)
        sb_x.grid(row=1, column=0, sticky=tk.EW)
        self.file_list.bind("<ButtonRelease-1>", self._on_tree_click)
        self.file_list.tag_configure("checked",   background="#EFF6FF")
        self.file_list.tag_configure("unchecked", background="#FFFFFF", foreground="#9CA3AF")

        # Word 변환 버튼 (파일 목록 바로 아래)
        word_bar = ttk.Frame(self, style="Card.TFrame")
        word_bar.grid(row=2, column=0, columnspan=2, sticky=tk.EW, pady=(2, 6))
        ttk.Button(word_bar, text="📝 체크된 파일 → Word 변환", command=self._export_to_word).pack(side=tk.LEFT, padx=4)
        ttk.Label(word_bar, text="(체크된 파일만 하나의 Word 보고서로 합칩니다)", style="Muted.TLabel").pack(side=tk.LEFT)

        opt = ttk.LabelFrame(self, text="보고서 옵션")
        opt.grid(row=3, column=0, columnspan=2, sticky=tk.EW, pady=10)
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
        self.btn.grid(row=4, column=0, pady=12)
        self.prog = ttk.Progressbar(self, mode="indeterminate", length=700)
        self.prog.grid(row=5, column=0, columnspan=2, sticky=tk.EW)

        ttk.Label(self, text="로그:", style="Title.TLabel").grid(row=6, column=0, sticky=tk.W, pady=8)
        lf2 = ttk.Frame(self, style="Card.TFrame")
        lf2.grid(row=7, column=0, columnspan=2, sticky=tk.NSEW, pady=8)
        sb2 = ttk.Scrollbar(lf2)
        sb2.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_box = tk.Text(lf2, height=10, yscrollcommand=sb2.set, wrap=tk.WORD, bg=LOG_BG, fg=LOG_FG, font=("Consolas", 9))
        self.log_box.pack(fill=tk.BOTH, expand=True)
        sb2.config(command=self.log_box.yview)

    def log(self, msg):
        self.log_box.insert(tk.END, msg + "\n")
        self.log_box.see(tk.END)
        self.update_idletasks()

    # ── 파일 목록 관리 ────────────────────────────────────
    def _refresh_count(self):
        total   = len(self.md_files)
        checked = sum(1 for iid in self.file_list.get_children()
                      if self.file_list.item(iid, "values")[0] == "✅")
        self.file_count_lbl.config(text=f"총 {total}개  |  체크 {checked}개")

    def _insert_file(self, path):
        path = os.path.normpath(path)
        if path in self.md_files:
            return False
        self.md_files.append(path)
        fname  = os.path.basename(path)
        folder = os.path.basename(os.path.dirname(path)) or os.path.dirname(path)
        iid = self.file_list.insert("", tk.END,
                                    values=("✅", fname, folder),
                                    tags=("checked",))
        return True

    def _on_tree_click(self, event):
        """행 클릭 시 체크 토글"""
        iid = self.file_list.identify_row(event.y)
        if not iid:
            return
        vals = list(self.file_list.item(iid, "values"))
        if vals[0] == "✅":
            vals[0] = "☐"
            self.file_list.item(iid, values=vals, tags=("unchecked",))
        else:
            vals[0] = "✅"
            self.file_list.item(iid, values=vals, tags=("checked",))
        self._refresh_count()

    def _check_all(self):
        all_iids = self.file_list.get_children()
        for iid in all_iids:
            vals = list(self.file_list.item(iid, "values"))
            vals[0] = "✅"
            self.file_list.item(iid, values=vals, tags=("checked",))
        # Treeview 행도 하이라이트 → 선택 제거 버튼과 연동
        self.file_list.selection_set(all_iids)
        self._refresh_count()

    def _uncheck_all(self):
        all_iids = self.file_list.get_children()
        for iid in all_iids:
            vals = list(self.file_list.item(iid, "values"))
            vals[0] = "☐"
            self.file_list.item(iid, values=vals, tags=("unchecked",))
        self.file_list.selection_remove(all_iids)
        self._refresh_count()

    def _remove_selected(self):
        """Treeview 에서 하이라이트된 항목 제거"""
        sel = self.file_list.selection()
        if not sel:
            messagebox.showinfo("알림", "제거할 파일을 클릭하거나\n[전체 선택] 후 제거하세요.")
            return
        for iid in sel:
            vals  = self.file_list.item(iid, "values")
            fname = vals[1]
            folder_name = vals[2]
            # md_files 에서도 제거
            self.md_files = [p for p in self.md_files
                             if not (os.path.basename(p) == fname and
                                     os.path.basename(os.path.dirname(p)) == folder_name)]
            self.file_list.delete(iid)
        self.log(f"{len(sel)}개 파일을 제거했습니다.")
        self._refresh_count()

    def _clear_all(self):
        if not self.md_files:
            return
        if messagebox.askyesno("확인", "목록을 전부 비울까요?"):
            self.file_list.delete(*self.file_list.get_children())
            self.md_files.clear()
            self.log("목록을 초기화했습니다.")
            self._refresh_count()

    def _load_folder(self):
        folder = filedialog.askdirectory(title="MD 파일이 있는 폴더 선택")
        if not folder:
            return
        try:
            # 재귀 탐색 + 루트 직접 탐색 (Windows 경로 대응)
            files = set()
            files.update(glob.glob(os.path.join(folder, "*.md")))
            files.update(glob.glob(os.path.join(folder, "**", "*.md"), recursive=True))
            files = sorted(files)
            self.log(f"폴더 탐색: '{os.path.basename(folder)}' → {len(files)}개 MD 파일 발견")
            if not files:
                messagebox.showinfo("알림", f"폴더에 MD 파일이 없습니다.\n{folder}")
                return
            count = 0
            for f in files:
                try:
                    if self._insert_file(f):
                        count += 1
                except Exception as e:
                    self.log(f"  [오류] {os.path.basename(f)}: {e}")
            self.log(f"{count}개 파일 추가 완료.")
            self._refresh_count()
            self.update_idletasks()
        except Exception as e:
            self.log(f"[오류] 폴더 로드 실패: {e}")
            messagebox.showerror("오류", str(e))

    def _load_files(self):
        files = filedialog.askopenfilenames(
            title="MD 파일 선택",
            filetypes=[("Markdown files", "*.md"), ("All files", "*.*")])
        count = sum(1 for f in files if self._insert_file(f))
        self.log(f"{count}개 파일 추가.")
        self._refresh_count()

    def _get_active_files(self):
        """체크(✅)된 파일 경로 반환. 없으면 전체 반환."""
        checked = []
        for iid in self.file_list.get_children():
            vals = self.file_list.item(iid, "values")
            if vals[0] == "✅":
                fname  = vals[1]
                folder = vals[2]
                for p in self.md_files:
                    if (os.path.basename(p) == fname and
                            os.path.basename(os.path.dirname(p)) == folder):
                        checked.append(p)
                        break
        return checked if checked else list(self.md_files)

    def _generate(self):
        if not self.md_files:
            messagebox.showwarning("경고", "생성할 MD 파일이 없습니다.")
            return
        active = self._get_active_files()
        selected_contents = []
        for path in active:
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
        checked_n = sum(1 for iid in self.file_list.get_children()
                         if self.file_list.item(iid, "values")[0] == "✅")
        hint = f"체크 {checked_n}개 / 전체 {len(self.md_files)}개"
        self.log(f"파일 종합 중... ({hint})")

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
        active = self._get_active_files()
        checked_n = sum(1 for iid in self.file_list.get_children()
                         if self.file_list.item(iid, "values")[0] == "✅")
        sel_hint = f"체크 {checked_n}개 / 전체 {len(self.md_files)}개"
        self.btn.config(state="disabled")
        self.prog.start()
        self.log(f"Word 변환 중 ({sel_hint})...")

        def _worker():
            try:
                conv = MDConverter()
                out_dir = self.out_var.get().strip() or "./word_output"
                os.makedirs(out_dir, exist_ok=True)
                results = []
                for path in active:
                    try:
                        fname = os.path.basename(path).replace(".md", "")
                        out_path = os.path.join(out_dir, f"{fname}.docx")
                        result = generate_word_report(path, out_path,
                                                       title=fname, callback=self.log)
                        if not result:
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
                if not generate_word_report(path, out_docx,
                                             title=self.report_title_var.get(),
                                             callback=self.log):
                    # node 없으면 python-docx 폴백
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
            "Vision API 이미지 분석 안내\n"
            "이미지를 base64로 인코딩하여 Qwen Vision 모델에 직접 전달합니다.\n"
            "별도 OCR 라이브러리 설치 없이 동작합니다.\n"
            "Vision 모델 설정: LLM 설정 탭의 Vision 모델 항목을 확인하세요."
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
