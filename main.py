"""
main.py - Confluence 보고서 자동화 도구 메인 메뉴
"""
import sys
from collector import ConfluenceCollector
from report_generator import analyze_multiple_pages, analyze_with_llm, create_word_report


def print_banner():
    print("=" * 55)
    print("   📊 Confluence 보고서 자동화 도구")
    print("=" * 55)


def menu_collect_group():
    """[1] 그룹 전체 자동 수집"""
    print("\n── 그룹 전체 자동 수집 ──────────────────────────")
    root_id = input("그룹 최상위 페이지 ID 입력: ").strip()
    if not root_id:
        print("❌ 취소됨")
        return
    collector = ConfluenceCollector()
    if not collector.test_connection():
        return
    parts = collector.explore_group_structure(root_id)
    collector.collect_all_from_structure(parts)


def menu_collect_manual():
    """[2] page_id 직접 입력 수집"""
    print("\n── page_id 직접 입력 수집 ──────────────────────")
    ids_input = input("page_id 입력 (콤마로 구분): ").strip()
    page_ids = [pid.strip() for pid in ids_input.split(",") if pid.strip()]
    if not page_ids:
        print("❌ 취소됨")
        return
    collector = ConfluenceCollector()
    if not collector.test_connection():
        return
    collector.collect_pages(page_ids)


def menu_generate_report():
    """[3] 보고서 생성"""
    print("\n── 보고서 생성 ──────────────────────────────────")
    collector = ConfluenceCollector()
    opts = collector.get_available_options()

    if not any(opts.values()):
        print("❌ 수집된 데이터가 없습니다. 먼저 [1] 또는 [2]로 수집하세요.")
        return

    # 파트 선택
    part = None
    if opts["parts"]:
        print("\n파트 선택 (0=전체):")
        for i, p in enumerate(opts["parts"], 1):
            print(f"  {i}. {p}")
        sel = input("선택 (엔터=전체): ").strip()
        if sel.isdigit() and 1 <= int(sel) <= len(opts["parts"]):
            part = opts["parts"][int(sel)-1]

    # 년도 선택
    year = None
    if opts["years"]:
        print(f"\n년도 선택 (0=전체): {opts['years']}")
        sel = input("년도 입력 (엔터=전체): ").strip()
        if sel in opts["years"]:
            year = sel

    # 주차 선택
    week = None
    if opts["weeks"]:
        print("\n주차 선택 (0=전체):")
        for i, w in enumerate(opts["weeks"][:20], 1):
            print(f"  {i}. {w}")
        sel = input("선택 (엔터=전체): ").strip()
        if sel.isdigit() and 1 <= int(sel) <= len(opts["weeks"]):
            week = opts["weeks"][int(sel)-1]

    # 유형 선택
    ptype = None
    if opts["types"]:
        print(f"\n유형 선택: {opts['types']} (엔터=전체)")
        sel = input("유형 입력 (엔터=전체): ").strip()
        if sel in opts["types"]:
            ptype = sel

    # 데이터 조회
    pages = collector.get_pages(part_name=part, year=year, week_label=week, page_type=ptype)
    if not pages:
        print("❌ 조건에 맞는 데이터가 없습니다.")
        return

    print(f"\n📄 {len(pages)}개 페이지 분석 중...")

    # 제목 생성
    title_parts = ["보고서"]
    if part:  title_parts.insert(0, part)
    if year:  title_parts.append(year)
    if week:  title_parts.append(week)
    report_title = " | ".join(title_parts)

    # 분석
    if len(pages) == 1:
        _, title, text, p_name, p_year, p_week, p_type = pages[0]
        analysis = analyze_with_llm(text, title=title)
    else:
        analysis = analyze_multiple_pages(pages)

    # 저장
    output = create_word_report(
        analysis_text=analysis,
        title=report_title,
        meta={"part": part, "year": year, "week": week},
    )
    print(f"\n🎉 완료! 파일 위치: {output}")


def menu_status():
    """[4] 수집 현황"""
    print("\n── 수집 현황 ──────────────────────────────────")
    collector = ConfluenceCollector()
    print(collector.get_db_summary())


def main():
    print_banner()
    while True:
        print("\n메뉴:")
        print("  [1] 그룹 전체 자동 수집  (최상위 page_id 입력)")
        print("  [2] page_id 직접 입력 수집")
        print("  [3] 보고서 생성")
        print("  [4] 수집 현황 보기")
        print("  [0] 종료")
        choice = input("\n선택: ").strip()

        if choice == "1":
            menu_collect_group()
        elif choice == "2":
            menu_collect_manual()
        elif choice == "3":
            menu_generate_report()
        elif choice == "4":
            menu_status()
        elif choice == "0":
            print("👋 종료합니다.")
            sys.exit(0)
        else:
            print("❌ 잘못된 입력입니다.")


if __name__ == "__main__":
    main()
