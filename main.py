"""
도장 생산계획 시스템 — 메인 실행 v3.1

안전재고 3일치 + 지그 템플릿 선행 + 컬러 블록 스케줄링
"""
from config import (
    HANGER_COUNT, JIGS_PER_HANGER, ROTATIONS_PER_DAY,
    MAX_JIG_CHANGES_PER_DAY, DAILY_CAPACITY,
    SAFETY_STOCK_DAYS, PLANNING_DAYS,
)
from sample_data import get_all_sample_data, MAJOR_PRODUCTS
from bom_explode import explode_demand, summarize_daily_paint
from production_planner import calculate_production_plan, print_production_summary
from paint_scheduler import (
    schedule_painting,
    print_schedule_summary,
    print_day_rotations,
    print_jig_type_analysis,
)
from injection_scheduler import calculate_injection_schedule, print_injection_schedule


def main():
    print("═" * 100)
    print("도장 생산계획 시스템 v3.1 — 안전재고 기반")
    print("═" * 100)
    print(f"\n시스템 변수:")
    print(f"  행어: {HANGER_COUNT}개 | 지그/행어: {JIGS_PER_HANGER}개 | "
          f"회전: {ROTATIONS_PER_DAY}회/일")
    print(f"  일일 용량: {DAILY_CAPACITY:,}개 | "
          f"지그교환 한도: {MAX_JIG_CHANGES_PER_DAY}건/일")
    print(f"  안전재고: {SAFETY_STOCK_DAYS}일치 | "
          f"계획기간: {PLANNING_DAYS}일 + 참조 {SAFETY_STOCK_DAYS}일")
    print(f"  주력 아이템: {', '.join(MAJOR_PRODUCTS)} (~40%)")

    # ── Step 1: 샘플 데이터 ──
    print(f"\n{'─' * 100}")
    print("[1] 샘플 데이터 생성...")
    data = get_all_sample_data()
    bom = data["bom"]

    # 주력/비주력 완성품 수
    major_fps = sum(1 for v in bom.values()
                    if v["injection_product"] in MAJOR_PRODUCTS)
    minor_fps = len(bom) - major_fps
    print(f"  사출품: {len(data['injection_products'])}종 | "
          f"컬러: {len(data['colors'])}색 | "
          f"완성품: {len(bom)}종 (주력 {major_fps} + 비주력 {minor_fps})")
    print(f"  수요: {len(data['demand'])}건 "
          f"({PLANNING_DAYS}+{SAFETY_STOCK_DAYS}={PLANNING_DAYS + SAFETY_STOCK_DAYS}일)")

    # ── Step 2: BOM 전개 ──
    print(f"\n{'─' * 100}")
    print("[2] BOM 전개...")
    paint_req, inj_req = explode_demand(bom, data["demand"])
    daily_paint = summarize_daily_paint(paint_req)

    all_dates = sorted(daily_paint.keys())
    plan_dates = all_dates[:PLANNING_DAYS]

    print(f"  전체 기간: {all_dates[0]} ~ {all_dates[-1]} "
          f"({len(all_dates)}일)")
    print(f"  생산 기간: {plan_dates[0]} ~ {plan_dates[-1]} ({len(plan_dates)}일)")

    for dt in all_dates:
        items = daily_paint[dt]
        total = sum(items.values())
        is_plan = "◆" if dt in plan_dates else "  (참조)"
        # 주력 비율
        major_qty = sum(q for (p, c), q in items.items()
                        if p in MAJOR_PRODUCTS)
        pct = major_qty / total * 100 if total > 0 else 0
        print(f"    {dt}: {total:,}개 (주력 {pct:.0f}%) {is_plan}")

    # ── Step 3: 순수요 계산 (안전재고 반영) ──
    print(f"\n{'─' * 100}")
    print(f"[3] 순수요 계산 (안전재고 {SAFETY_STOCK_DAYS}일)...")
    production_plan, inv_report = calculate_production_plan(
        daily_paint, planning_days=PLANNING_DAYS
    )
    print_production_summary(production_plan, inv_report, daily_paint)

    # ── Step 4: 도장 스케줄링 ──
    print(f"\n{'─' * 100}")
    print("[4] 도장 스케줄링 (지그 선행 + 컬러 블록)...")

    # 생산계획(순수요)을 도장 스케줄러에 전달
    results = schedule_painting(production_plan, data["color_transition_matrix"])

    print_schedule_summary(results)

    # 첫날 회전별 상세
    print_day_rotations(results[0])

    # 지그 유형 분석
    print_jig_type_analysis(results)

    # ── Step 5: 사출 역산 ──
    print(f"\n{'─' * 100}")
    print("[5] 사출 일정 역산 (도장 D일 → 사출 D-1일)...")
    inj_schedule = calculate_injection_schedule(production_plan, lead_time_days=1)
    print_injection_schedule(inj_schedule)


if __name__ == "__main__":
    main()
