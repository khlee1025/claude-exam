"""
collector.py - Confluence 페이지 수집 모듈
그룹 구조: 그룹 > 파트(5개) > 과제(년도별/주차별) + 회의록
"""
import sqlite3
import os
import re
import time
import requests
import urllib3
from dataclasses import dataclass, field
from typing import Optional
from bs4 import BeautifulSoup
import config

# SSL 경고 무시 (내부 서버 자체 서명 인증서 대응)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


@dataclass
class PageNode:
    """Confluence 페이지 트리 노드"""
    page_id: str
    title: str
    page_type: str = "unknown"   # "과제" | "회의록" | "unknown"
    year: str = ""
    week_label: str = ""
    part_name: str = ""
    children: list = field(default_factory=list)


class ConfluenceCollector:
    """Confluence 내부 서버 수집기 (Basic Auth / AD)"""

    def __init__(self):
        config.validate_confluence_config()
        self.base_url = config.CONFLUENCE_URL
        self.session = requests.Session()
        self.session.auth = (config.CONFLUENCE_USERNAME, config.CONFLUENCE_PASSWORD)
        self.session.verify = False  # 내부 서버 SSL 무시
        self.session.headers.update({"Content-Type": "application/json"})
        self._init_db()

    # ── DB 초기화 ────────────────────────────────────────────
    def _init_db(self):
        os.makedirs(config.IMAGES_DIR, exist_ok=True)
        conn = sqlite3.connect(config.DB_PATH)
        c = conn.cursor()
        c.execute("""
            CREATE TABLE IF NOT EXISTS pages (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                page_id      TEXT NOT NULL,
                title        TEXT,
                content_html TEXT,
                content_text TEXT,
                part_name    TEXT DEFAULT '',
                year         TEXT DEFAULT '',
                week_label   TEXT DEFAULT '',
                page_type    TEXT DEFAULT '',
                collected_at TEXT DEFAULT (datetime('now','localtime')),
                UNIQUE(page_id)
            )
        """)
        # 기존 DB 마이그레이션 (컬럼 없으면 추가)
        existing = {row[1] for row in c.execute("PRAGMA table_info(pages)")}
        for col, typedef in [
            ("part_name", "TEXT DEFAULT ''"),
            ("year",      "TEXT DEFAULT ''"),
            ("week_label","TEXT DEFAULT ''"),
            ("page_type", "TEXT DEFAULT ''"),
        ]:
            if col not in existing:
                c.execute(f"ALTER TABLE pages ADD COLUMN {col} {typedef}")
        conn.commit()
        conn.close()

    # ── Confluence REST API ───────────────────────────────────
    def _get(self, path: str, params: dict = None):
        url = f"{self.base_url}/rest/api{path}"
        r = self.session.get(url, params=params, timeout=30)
        r.raise_for_status()
        return r.json()

    def get_page(self, page_id: str) -> dict:
        """단일 페이지 내용 조회"""
        return self._get(f"/content/{page_id}", {"expand": "body.storage,metadata,space"})

    def get_child_pages(self, page_id: str) -> list:
        """하위 페이지 목록 전체 조회 (페이지네이션 처리)"""
        children = []
        start = 0
        limit = 50
        while True:
            data = self._get(f"/content/{page_id}/child/page",
                             {"limit": limit, "start": start, "expand": "metadata"})
            results = data.get("results", [])
            children.extend(results)
            if len(results) < limit:
                break
            start += limit
        return children

    def test_connection(self) -> bool:
        """연결 테스트"""
        try:
            self._get("/space", {"limit": 1})
            print("✅ Confluence 연결 성공")
            return True
        except Exception as e:
            print(f"❌ 연결 실패: {e}")
            return False

    # ── 주차 추출 유틸 ────────────────────────────────────────
    @staticmethod
    def _extract_week(title: str) -> str:
        m = re.search(r"(\d+)\s*주\s*차?", title)
        if m:
            return f"{m.group(1)}주차"
        m2 = re.search(r"week\s*(\d+)", title, re.IGNORECASE)
        if m2:
            return f"{m2.group(1)}주차"
        return ""

    @staticmethod
    def _extract_year(title: str) -> str:
        m = re.search(r"(20\d{2})", title)
        return m.group(1) if m else ""

    @staticmethod
    def _classify_type(title: str) -> str:
        t = title.lower()
        if any(k in t for k in ["회의록", "minutes", "meeting"]):
            return "회의록"
        if any(k in t for k in ["과제", "task", "project", "업무"]):
            return "과제"
        return "unknown"

    # ── 그룹 전체 자동 탐색 ──────────────────────────────────
    def explore_group_structure(self, root_page_id: str) -> list:
        """
        그룹 최상위 페이지 ID 하나로 전체 구조 탐색
        반환: [PageNode, ...]  (파트 레벨)
        """
        print(f"\n🔍 그룹 구조 탐색 시작 (root: {root_page_id})")
        parts = []

        # 1단계: 최상위 → 파트들
        part_pages = self.get_child_pages(root_page_id)
        for part in part_pages:
            part_node = PageNode(
                page_id=part["id"],
                title=part["title"],
                part_name=part["title"],
            )
            print(f"  📁 파트: {part['title']}")

            # 2단계: 파트 → 과제/회의록
            sub_pages = self.get_child_pages(part["id"])
            for sub in sub_pages:
                sub_type = self._classify_type(sub["title"])
                sub_node = PageNode(
                    page_id=sub["id"],
                    title=sub["title"],
                    page_type=sub_type,
                    part_name=part["title"],
                )

                if sub_type == "과제":
                    # 3단계: 과제 → 년도
                    year_pages = self.get_child_pages(sub["id"])
                    for yp in year_pages:
                        year = self._extract_year(yp["title"]) or yp["title"]
                        year_node = PageNode(
                            page_id=yp["id"],
                            title=yp["title"],
                            page_type="과제",
                            part_name=part["title"],
                            year=year,
                        )
                        # 4단계: 년도 → 주차
                        week_pages = self.get_child_pages(yp["id"])
                        for wp in week_pages:
                            week = self._extract_week(wp["title"]) or wp["title"]
                            week_node = PageNode(
                                page_id=wp["id"],
                                title=wp["title"],
                                page_type="과제",
                                part_name=part["title"],
                                year=year,
                                week_label=week,
                            )
                            year_node.children.append(week_node)
                        sub_node.children.append(year_node)

                elif sub_type == "회의록":
                    # 회의록 하위 페이지 수집
                    meeting_pages = self.get_child_pages(sub["id"])
                    for mp in meeting_pages:
                        m_node = PageNode(
                            page_id=mp["id"],
                            title=mp["title"],
                            page_type="회의록",
                            part_name=part["title"],
                            week_label=mp["title"],
                        )
                        sub_node.children.append(m_node)

                part_node.children.append(sub_node)
            parts.append(part_node)

        print(f"✅ 구조 탐색 완료: {len(parts)}개 파트 발견")
        return parts

    # ── HTML → 텍스트 변환 ───────────────────────────────────
    @staticmethod
    def _html_to_text(html: str) -> str:
        soup = BeautifulSoup(html, "html.parser")
        return soup.get_text(separator="\n", strip=True)

    # ── 페이지 내용 수집 및 DB 저장 ──────────────────────────
    def collect_page(self, node: PageNode) -> bool:
        """단일 페이지 수집 후 DB 저장"""
        try:
            data = self.get_page(node.page_id)
            html = data.get("body", {}).get("storage", {}).get("value", "")
            text = self._html_to_text(html)
            conn = sqlite3.connect(config.DB_PATH)
            conn.execute("""
                INSERT OR REPLACE INTO pages
                    (page_id, title, content_html, content_text,
                     part_name, year, week_label, page_type)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (node.page_id, node.title, html, text,
                  node.part_name, node.year, node.week_label, node.page_type))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"    ⚠️ 수집 실패 ({node.title}): {e}")
            return False

    def collect_all_from_structure(self, parts: list):
        """구조 탐색 결과로 전체 수집"""
        total, success = 0, 0

        def _collect_recursive(node: PageNode, depth=0):
            nonlocal total, success
            if node.week_label or node.page_type == "회의록":
                label = f"{node.part_name} > {node.year or ''} > [{node.page_type}] {node.week_label or node.title}"
                print(f"  {'  '*depth}📥 {label.strip(' >')} 수집 중...")
                total += 1
                if self.collect_page(node):
                    success += 1
            for child in node.children:
                _collect_recursive(child, depth+1)

        for part in parts:
            print(f"\n📂 {part.part_name}")
            for sub in part.children:
                _collect_recursive(sub)
            time.sleep(0.5)  # 서버 부하 방지

        print(f"\n✅ 수집 완료: {success}/{total}개 성공")

    # ── 직접 page_id 수집 (레거시) ───────────────────────────
    def collect_pages(self, page_ids: list):
        """page_id 목록으로 직접 수집"""
        for pid in page_ids:
            try:
                data = self.get_page(pid)
                html = data.get("body", {}).get("storage", {}).get("value", "")
                text = self._html_to_text(html)
                node = PageNode(page_id=pid, title=data.get("title",""))
                conn = sqlite3.connect(config.DB_PATH)
                conn.execute("""
                    INSERT OR REPLACE INTO pages
                        (page_id, title, content_html, content_text)
                    VALUES (?, ?, ?, ?)
                """, (pid, node.title, html, text))
                conn.commit()
                conn.close()
                print(f"  ✅ {node.title} ({pid})")
            except Exception as e:
                print(f"  ❌ 실패 ({pid}): {e}")

    # ── DB 조회 함수 ──────────────────────────────────────────
    def get_pages(self, part_name=None, year=None, week_label=None, page_type=None) -> list:
        """조건으로 DB에서 페이지 조회"""
        conn = sqlite3.connect(config.DB_PATH)
        query = "SELECT page_id, title, content_text, part_name, year, week_label, page_type FROM pages WHERE 1=1"
        params = []
        if part_name:
            query += " AND part_name LIKE ?"
            params.append(f"%{part_name}%")
        if year:
            query += " AND year = ?"
            params.append(year)
        if week_label:
            query += " AND week_label LIKE ?"
            params.append(f"%{week_label}%")
        if page_type:
            query += " AND page_type = ?"
            params.append(page_type)
        rows = conn.execute(query, params).fetchall()
        conn.close()
        return rows

    def get_available_options(self) -> dict:
        """DB에 저장된 파트/년도/주차 목록 반환"""
        conn = sqlite3.connect(config.DB_PATH)
        parts = [r[0] for r in conn.execute("SELECT DISTINCT part_name FROM pages WHERE part_name != '' ORDER BY part_name").fetchall()]
        years = [r[0] for r in conn.execute("SELECT DISTINCT year FROM pages WHERE year != '' ORDER BY year DESC").fetchall()]
        weeks = [r[0] for r in conn.execute("SELECT DISTINCT week_label FROM pages WHERE week_label != '' ORDER BY week_label").fetchall()]
        types = [r[0] for r in conn.execute("SELECT DISTINCT page_type FROM pages WHERE page_type != '' ORDER BY page_type").fetchall()]
        conn.close()
        return {"parts": parts, "years": years, "weeks": weeks, "types": types}

    def get_db_summary(self) -> str:
        """수집 현황 트리뷰 반환"""
        conn = sqlite3.connect(config.DB_PATH)
        rows = conn.execute("""
            SELECT part_name, year, page_type, week_label
            FROM pages WHERE part_name != ''
            ORDER BY part_name, year, page_type, week_label
        """).fetchall()
        conn.close()
        if not rows:
            return "  (수집된 데이터 없음)"
        lines = []
        current = {}
        for part, year, ptype, week in rows:
            if part not in current:
                lines.append(f"📂 {part}")
                current = {part: {}}
            key = f"{year}|{ptype}"
            if key not in current[part]:
                lines.append(f"  ├─ {year or '?'} [{ptype}]")
                current[part][key] = True
            lines.append(f"  │   └─ {week or '(제목없음)'}")
        return "\n".join(lines)
