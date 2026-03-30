"""
도장 생산계획 시스템 — BOM 전개

완성품 수요 → (사출품, 컬러, 날짜) 소요량 계산
"""
from collections import defaultdict


def explode_demand(bom, demand):
    """
    완성품 수요를 BOM 전개하여 도장 소요량으로 변환

    Args:
        bom: {finished_product_id: {injection_product, color, name}}
        demand: [{date, product_id, qty}, ...]

    Returns:
        paint_requirements: {(date, injection_product, color): qty}
        injection_requirements: {(date, injection_product): qty}
    """
    paint_req = defaultdict(int)
    injection_req = defaultdict(int)
    missing = []

    for d in demand:
        fp_id = d["product_id"]
        if fp_id not in bom:
            missing.append(fp_id)
            continue

        entry = bom[fp_id]
        inj = entry["injection_product"]
        color = entry["color"]
        qty = d["qty"]
        dt = d["date"]

        # 도장 소요: (날짜, 사출품, 컬러) → 수량
        paint_req[(dt, inj, color)] += qty

        # 사출 소요: (날짜, 사출품) → 수량 (컬러 무관, 합산)
        injection_req[(dt, inj)] += qty

    if missing:
        print(f"  [경고] BOM 없는 완성품 {len(missing)}건 스킵")

    return dict(paint_req), dict(injection_req)


def summarize_daily_paint(paint_requirements):
    """
    일별 도장 소요량 요약

    Returns:
        {date: {(injection_product, color): qty}}
    """
    daily = defaultdict(lambda: defaultdict(int))
    for (dt, inj, color), qty in paint_requirements.items():
        daily[dt][(inj, color)] += qty
    return {dt: dict(items) for dt, items in daily.items()}


def summarize_daily_injection(injection_requirements):
    """
    일별 사출 소요량 요약

    Returns:
        {date: {injection_product: qty}}
    """
    daily = defaultdict(lambda: defaultdict(int))
    for (dt, inj), qty in injection_requirements.items():
        daily[dt][inj] += qty
    return {dt: dict(items) for dt, items in daily.items()}


if __name__ == "__main__":
    from sample_data import get_all_sample_data

    data = get_all_sample_data()
    paint_req, inj_req = explode_demand(data["bom"], data["demand"])

    daily_paint = summarize_daily_paint(paint_req)
    daily_inj = summarize_daily_injection(inj_req)

    print("=" * 60)
    print("일별 도장 소요량")
    print("=" * 60)
    for dt in sorted(daily_paint):
        items = daily_paint[dt]
        total = sum(items.values())
        n_combos = len(items)
        colors_used = len(set(c for _, c in items))
        products_used = len(set(p for p, _ in items))
        print(f"\n{dt}: 총 {total:,}개, {n_combos}조합 "
              f"(사출품 {products_used}종, 컬러 {colors_used}색)")
        # 상위 5개 조합
        top5 = sorted(items.items(), key=lambda x: -x[1])[:5]
        for (inj, color), qty in top5:
            print(f"  {inj} + {color}: {qty}개")

    print("\n" + "=" * 60)
    print("일별 사출 소요량")
    print("=" * 60)
    for dt in sorted(daily_inj):
        items = daily_inj[dt]
        total = sum(items.values())
        print(f"\n{dt}: 총 {total:,}개")
        for inj in sorted(items, key=lambda x: -items[x]):
            print(f"  {inj}: {items[inj]}개")
