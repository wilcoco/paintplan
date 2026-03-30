"""
도장 생산계획 시스템 — 순수요 계산 (안전재고 반영)

우선순위 (단계별):
  1순위: 당일(D+0) 수요 충족
  2순위: D+1 수요분 확보
  3순위: D+2 수요분 확보
  4순위: D+3 수요분 확보

각 단계에서 용량이 남으면 다음 단계로 넘어감
"""
import math
from collections import defaultdict
from config import (
    HANGER_COUNT, JIGS_PER_HANGER, ROTATIONS_PER_DAY,
    SAFETY_STOCK_DAYS, PLANNING_DAYS,
)

DAILY_CAPACITY = HANGER_COUNT * JIGS_PER_HANGER * ROTATIONS_PER_DAY


def calculate_production_plan(daily_paint_demand, planning_days=PLANNING_DAYS,
                              initial_inventory=None):
    """
    단계별 우선순위 기반 도장 생산계획

    Args:
        daily_paint_demand: {date: {(product, color): qty}}
        planning_days: 생산 일수 (12)
        initial_inventory: {(product, color): qty} 초기 재고 (없으면 0)

    Returns:
        production_plan: {date: {(product, color): qty}}
        inventory_report: [...]
    """
    all_dates = sorted(daily_paint_demand.keys())
    plan_dates = all_dates[:planning_days]

    all_items = set()
    for items in daily_paint_demand.values():
        all_items.update(items.keys())

    # 초기 재고 설정
    inventory = defaultdict(int)
    if initial_inventory:
        for item, qty in initial_inventory.items():
            inventory[item] = qty
    production_plan = {}
    inventory_report = []

    for day_idx, date_str in enumerate(plan_dates):
        day_demand = daily_paint_demand.get(date_str, {})

        # ── 향후 N일 수요 조회 ──
        future_demands = []  # [D+1 수요, D+2 수요, D+3 수요]
        for offset in range(1, SAFETY_STOCK_DAYS + 1):
            future_idx = day_idx + offset
            if future_idx < len(all_dates):
                future_demands.append(
                    daily_paint_demand.get(all_dates[future_idx], {})
                )
            else:
                future_demands.append({})

        # ── 단계별 용량 배분 ──
        day_production = {}
        remaining_cap = DAILY_CAPACITY

        # --- 1순위: 당일 수요 충족 ---
        remaining_cap = _allocate_priority(
            day_production, inventory, day_demand,
            remaining_cap, "D+0 당일"
        )

        # --- 2순위: D+1 수요분 확보 ---
        # 현재 재고(생산 포함) - 당일 소비 후, D+1 수요를 커버할 재고 확보
        if remaining_cap > 0 and len(future_demands) >= 1:
            cumulative_demand = {}  # D+1까지의 누적 수요
            for item in all_items:
                d0 = day_demand.get(item, 0)
                d1 = future_demands[0].get(item, 0)
                cumulative_demand[item] = d0 + d1
            remaining_cap = _allocate_priority(
                day_production, inventory, cumulative_demand,
                remaining_cap, "D+1"
            )

        # --- 3순위: D+2 수요분 확보 ---
        if remaining_cap > 0 and len(future_demands) >= 2:
            cumulative_demand = {}
            for item in all_items:
                d0 = day_demand.get(item, 0)
                d1 = future_demands[0].get(item, 0)
                d2 = future_demands[1].get(item, 0)
                cumulative_demand[item] = d0 + d1 + d2
            remaining_cap = _allocate_priority(
                day_production, inventory, cumulative_demand,
                remaining_cap, "D+2"
            )

        # --- 4순위: D+3 수요분 확보 ---
        if remaining_cap > 0 and len(future_demands) >= 3:
            cumulative_demand = {}
            for item in all_items:
                d0 = day_demand.get(item, 0)
                d1 = future_demands[0].get(item, 0)
                d2 = future_demands[1].get(item, 0)
                d3 = future_demands[2].get(item, 0)
                cumulative_demand[item] = d0 + d1 + d2 + d3
            remaining_cap = _allocate_priority(
                day_production, inventory, cumulative_demand,
                remaining_cap, "D+3"
            )

        # --- 5순위: 잔여 용량 채우기 (컨베이어 10회전 무조건 가동) ---
        # D+4 이후 수요까지 확장하여 빈 회전 없이 풀 가동
        if remaining_cap > 0:
            extended_offset = SAFETY_STOCK_DAYS + 1
            while remaining_cap > 0:
                future_idx = day_idx + extended_offset
                if future_idx >= len(all_dates):
                    break
                future_date = all_dates[future_idx]
                ext_demand = daily_paint_demand.get(future_date, {})
                if not ext_demand:
                    extended_offset += 1
                    continue

                cumulative_demand = {}
                for item in all_items:
                    cum = 0
                    for off in range(extended_offset + 1):
                        idx = day_idx + off
                        if idx < len(all_dates):
                            cum += daily_paint_demand.get(
                                all_dates[idx], {}
                            ).get(item, 0)
                    cumulative_demand[item] = cum

                prev_cap = remaining_cap
                remaining_cap = _allocate_priority(
                    day_production, inventory, cumulative_demand,
                    remaining_cap, f"D+{extended_offset}"
                )
                if remaining_cap == prev_cap:
                    break  # 더 이상 배분할 수요 없음
                extended_offset += 1

        # 잔여 용량이 있어도 미래 수요가 모두 커버되면 생산 중단
        # (컨베이어는 가동하되, 생산할 것이 없으면 빈 회전)

        production_plan[date_str] = day_production

        # ── 재고 업데이트 + 리포트 ──
        safety_target_total = {}
        for item in all_items:
            st = 0
            for fd in future_demands:
                st += fd.get(item, 0)
            safety_target_total[item] = st

        day_report = {"date": date_str, "items": {}}

        for item in all_items:
            stock_start = inventory[item]
            demand = day_demand.get(item, 0)
            produced = day_production.get(item, 0)
            stock_end = stock_start + produced - demand
            safety = safety_target_total.get(item, 0)

            if stock_end < 0:
                status = "SHORTAGE"
            elif stock_end < safety:
                status = "BELOW_SAFETY"
            else:
                status = "OK"

            day_report["items"][item] = {
                "demand": demand,
                "production": produced,
                "stock_start": stock_start,
                "stock_end": stock_end,
                "safety_target": safety,
                "status": status,
            }
            inventory[item] = stock_end

        inventory_report.append(day_report)

    return production_plan, inventory_report


def _allocate_priority(day_production, inventory, target_demand,
                       remaining_cap, label):
    """
    특정 우선순위 단계의 수요를 충족하도록 생산 배분

    target_demand: 이 단계까지 누적 수요 {item: cumulative_qty}
    이미 생산된 것(day_production) + 현재 재고(inventory)로 부족한 만큼 추가 생산

    Returns:
        remaining_cap: 남은 용량
    """
    needs = {}
    for item, cum_demand in target_demand.items():
        if cum_demand <= 0:
            continue
        stock = inventory[item]
        already_produced = day_production.get(item, 0)
        available = stock + already_produced
        shortfall = cum_demand - available
        if shortfall > 0:
            needs[item] = shortfall

    if not needs or remaining_cap <= 0:
        return remaining_cap

    total_needed = sum(needs.values())

    if total_needed <= remaining_cap:
        # 전량 충족 가능
        for item, qty in needs.items():
            day_production[item] = day_production.get(item, 0) + qty
            remaining_cap -= qty
    else:
        # 비례 배분 (큰 수요 우선)
        for item, qty in sorted(needs.items(), key=lambda x: -x[1]):
            if remaining_cap <= 0:
                break
            alloc = min(
                qty,
                max(1, round(remaining_cap * qty / max(1, total_needed)))
            )
            alloc = min(alloc, remaining_cap)
            if alloc > 0:
                day_production[item] = day_production.get(item, 0) + alloc
                remaining_cap -= alloc
                total_needed -= qty

    return remaining_cap


def print_production_summary(production_plan, inventory_report, daily_paint_demand):
    """생산계획 및 재고 현황 요약"""
    print(f"\n{'═' * 100}")
    print("도장 생산계획 — 단계별 우선순위")
    print(f"  1순위: 당일(D+0) → 2순위: D+1 → 3순위: D+2 → 4순위: D+3")
    print(f"  일일 용량: {DAILY_CAPACITY:,}개")
    print(f"{'═' * 100}")

    for report in inventory_report:
        date_str = report["date"]
        items = report["items"]

        demand_total = sum(v["demand"] for v in items.values())
        prod_total = sum(v["production"] for v in items.values())
        raw_demand = sum(daily_paint_demand.get(date_str, {}).values())
        cap_pct = prod_total / DAILY_CAPACITY * 100

        n_ok = sum(1 for v in items.values() if v["status"] == "OK")
        n_below = sum(1 for v in items.values() if v["status"] == "BELOW_SAFETY")
        n_short = sum(1 for v in items.values() if v["status"] == "SHORTAGE")

        status_str = f"OK:{n_ok}"
        if n_below > 0:
            status_str += f" 안전재고미달:{n_below}"
        if n_short > 0:
            status_str += f" 부족:{n_short}"

        print(f"\n  {date_str} │ 수요 {raw_demand:,} │ "
              f"생산 {prod_total:,} ({cap_pct:.0f}%) │ {status_str}")

        # 주력 아이템 상세
        for major_prod in ["001", "002"]:
            prod_items = {k: v for k, v in items.items()
                         if k[0] == major_prod and (v["demand"] > 0 or v["production"] > 0)}
            if prod_items:
                d_sum = sum(v["demand"] for v in prod_items.values())
                p_sum = sum(v["production"] for v in prod_items.values())
                s_end = sum(v["stock_end"] for v in prod_items.values())
                s_tgt = sum(v["safety_target"] for v in prod_items.values())
                delta = s_end - s_tgt
                mark = "✓" if delta >= 0 else f"▼{-delta}"
                print(f"    {major_prod}: 수요 {d_sum:3d} → 생산 {p_sum:3d} │ "
                      f"기말 {s_end:,} / 안전 {s_tgt:,} [{mark}]")

    # 합계
    print(f"\n{'─' * 100}")
    total_demand = sum(sum(v["demand"] for v in r["items"].values())
                       for r in inventory_report)
    total_prod = sum(sum(v["production"] for v in r["items"].values())
                     for r in inventory_report)
    n_short_days = sum(
        1 for r in inventory_report
        if any(v["status"] == "SHORTAGE" for v in r["items"].values())
    )
    n_ok_days = sum(
        1 for r in inventory_report
        if all(v["status"] == "OK" for v in r["items"].values())
    )
    print(f"  12일: 수요 {total_demand:,} │ 생산 {total_prod:,}")
    print(f"  부족일: {n_short_days}/12 │ 안전재고 완전달성: {n_ok_days}/12")
    print(f"{'═' * 100}")
