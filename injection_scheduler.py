"""
도장 생산계획 시스템 — 사출 일정 역산

도장 소요 기반으로 사출품의 사전 생산 일정 계산
"""
from collections import defaultdict


def calculate_injection_schedule(daily_paint_summary, lead_time_days=1):
    """
    도장 소요량 기반 사출 일정 역산

    도장 D일 소요 → 사출은 D - lead_time_days 까지 완료 필요

    Args:
        daily_paint_summary: {date: {(injection_product, color): qty}}
        lead_time_days: 사출→도장 리드타임 (일)

    Returns:
        injection_schedule: {date: {injection_product: qty}}
        (사출 완료 필요일 기준)
    """
    from datetime import datetime, timedelta

    schedule = defaultdict(lambda: defaultdict(int))

    for date_str in sorted(daily_paint_summary.keys()):
        paint_date = datetime.strptime(date_str, "%Y-%m-%d").date()
        injection_date = paint_date - timedelta(days=lead_time_days)
        inj_date_str = injection_date.isoformat()

        for (inj_product, color), qty in daily_paint_summary[date_str].items():
            schedule[inj_date_str][inj_product] += qty

    return {dt: dict(items) for dt, items in schedule.items()}


def print_injection_schedule(injection_schedule):
    """사출 일정 출력"""
    print("\n" + "=" * 80)
    print("사출 생산 일정 (도장 소요 기반 역산, 리드타임 1일)")
    print("=" * 80)

    for dt in sorted(injection_schedule.keys()):
        items = injection_schedule[dt]
        total = sum(items.values())
        print(f"\n{dt} (사출 완료 필요): 총 {total:,}개")
        for inj in sorted(items, key=lambda x: -items[x]):
            print(f"  {inj}: {items[inj]:,}개")
