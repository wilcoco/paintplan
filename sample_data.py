"""
도장 생산계획 시스템 — 샘플 데이터 생성

사출품 10종(001~010), 컬러 20색, 12일 수요
품번 체계: 001-RED, 001-BLUE, 002-YELLOW ...
  앞 번호(001) = 사출품 = 지그 결정
  뒤 컬러(RED) = 도장 색상

핵심: 001, 002가 전체 소요량의 ~40% 차지
"""
import random
from datetime import date, timedelta

random.seed(42)

# ─────────────────────────────────────────────
# 사출품 10종 (번호 = 지그 타입)
# ─────────────────────────────────────────────
INJECTION_PRODUCTS = [
    {"id": "001", "name": "프론트범퍼 상단"},  # 주력1 (~20%)
    {"id": "002", "name": "프론트범퍼 하단"},  # 주력2 (~20%)
    {"id": "003", "name": "프론트범퍼 좌측"},
    {"id": "004", "name": "프론트범퍼 우측"},
    {"id": "005", "name": "프론트범퍼 센터"},
    {"id": "006", "name": "리어범퍼 상단"},
    {"id": "007", "name": "리어범퍼 하단"},
    {"id": "008", "name": "리어범퍼 좌측"},
    {"id": "009", "name": "리어범퍼 우측"},
    {"id": "010", "name": "리어범퍼 센터"},
]

MAJOR_PRODUCTS = {"001", "002"}

# ─────────────────────────────────────────────
# 컬러 20색
# ─────────────────────────────────────────────
COLORS = [
    {"id": "WHITE",      "name": "퓨어화이트",   "group": "white"},
    {"id": "IVORY",      "name": "아이보리",     "group": "white"},
    {"id": "BLACK",      "name": "제트블랙",     "group": "black"},
    {"id": "MIDNIGHT",   "name": "미드나잇블랙", "group": "black"},
    {"id": "RED",        "name": "레이싱레드",   "group": "red"},
    {"id": "BURGUNDY",   "name": "버건디",       "group": "red"},
    {"id": "BLUE",       "name": "오션블루",     "group": "blue"},
    {"id": "SKYBLUE",    "name": "스카이블루",   "group": "blue"},
    {"id": "SILVER",     "name": "실버메탈릭",   "group": "silver"},
    {"id": "GRAY",       "name": "티타늄그레이", "group": "silver"},
    {"id": "GREEN",      "name": "포레스트그린", "group": "green"},
    {"id": "OLIVE",      "name": "올리브그린",   "group": "green"},
    {"id": "ORANGE",     "name": "선셋오렌지",   "group": "orange"},
    {"id": "YELLOW",     "name": "펄옐로우",     "group": "yellow"},
    {"id": "PURPLE",     "name": "딥퍼플",       "group": "purple"},
    {"id": "LAVENDER",   "name": "라벤더",       "group": "purple"},
    {"id": "GOLD",       "name": "샴페인골드",   "group": "gold"},
    {"id": "BRONZE",     "name": "브론즈",       "group": "gold"},
    {"id": "BEIGE",      "name": "세라믹베이지", "group": "beige"},
    {"id": "BROWN",      "name": "모카브라운",   "group": "brown"},
]

COLOR_IDS = [c["id"] for c in COLORS]
_color_group = {c["id"]: c["group"] for c in COLORS}

# ─────────────────────────────────────────────
# 컬러 전환 매트릭스
# ─────────────────────────────────────────────
_SIMILAR_GROUPS = {
    ("white", "beige"), ("beige", "white"),
    ("silver", "white"), ("white", "silver"),
    ("silver", "gold"),  ("gold", "silver"),
    ("red", "orange"),   ("orange", "red"),
    ("blue", "purple"),  ("purple", "blue"),
    ("green", "yellow"), ("yellow", "green"),
    ("gold", "brown"),   ("brown", "gold"),
    ("beige", "gold"),   ("gold", "beige"),
    ("orange", "yellow"),("yellow", "orange"),
}

_HARD_GROUPS = {
    ("white", "black"), ("black", "white"),
    ("white", "red"),   ("red", "white"),
    ("white", "blue"),  ("blue", "white"),
    ("black", "red"),   ("red", "black"),
}


def build_color_transition_matrix():
    """20×20 컬러 전환 매트릭스 (비워야 할 행어 수)"""
    matrix = {}
    for c_from in COLOR_IDS:
        for c_to in COLOR_IDS:
            if c_from == c_to:
                matrix[(c_from, c_to)] = 0
                continue
            g_from = _color_group[c_from]
            g_to = _color_group[c_to]
            if g_from == g_to:
                matrix[(c_from, c_to)] = 2
            elif (g_from, g_to) in _SIMILAR_GROUPS:
                matrix[(c_from, c_to)] = 4
            elif (g_from, g_to) in _HARD_GROUPS:
                matrix[(c_from, c_to)] = 8
            else:
                matrix[(c_from, c_to)] = 6
    return matrix


# ─────────────────────────────────────────────
# BOM: 완성품 = 사출품번호 + 컬러
# 품번: 001-RED, 001-BLUE, 002-WHITE ...
# ─────────────────────────────────────────────
def build_bom():
    """
    key: "001-RED" (사출품번호-컬러)
    value: {"injection_product": "001", "color": "RED", "name": "프론트범퍼 상단 레이싱레드"}
    """
    bom = {}
    for inj in INJECTION_PRODUCTS:
        is_major = inj["id"] in MAJOR_PRODUCTS
        n_colors = random.randint(10, 12) if is_major else random.randint(4, 6)

        must_have = ["WHITE", "BLACK"]
        others = [c for c in COLOR_IDS if c not in must_have]
        selected = must_have + random.sample(others, n_colors - len(must_have))

        for color_id in selected:
            fp_id = f"{inj['id']}-{color_id}"
            color_name = next(c["name"] for c in COLORS if c["id"] == color_id)
            bom[fp_id] = {
                "injection_product": inj["id"],
                "color": color_id,
                "name": f"{inj['name']} {color_name}",
            }
    return bom


# ─────────────────────────────────────────────
# 12일 + 3일(안전재고 참조) 수요
# 001, 002 = ~40%
# ─────────────────────────────────────────────
def generate_demand(bom, start_date=None, days=12, extra_days=3):
    if start_date is None:
        start_date = date(2026, 3, 18)

    total_days = days + extra_days
    demand = []

    major_fps = [fp for fp, info in bom.items()
                 if info["injection_product"] in MAJOR_PRODUCTS]
    minor_fps = [fp for fp, info in bom.items()
                 if info["injection_product"] not in MAJOR_PRODUCTS]

    for d in range(total_days):
        current_date = start_date + timedelta(days=d)
        target_total = random.randint(2000, 2400)
        major_target = int(target_total * 0.40)
        minor_target = target_total - major_target

        # 주력
        n_major = min(len(major_fps), random.randint(8, 12))
        active_major = random.sample(major_fps, n_major)
        major_total = 0
        for fp_id in active_major:
            remaining = major_target - major_total
            if remaining <= 0:
                break
            qty = random.randint(50, 150)
            qty = min(qty, remaining)
            if qty > 0:
                demand.append({"date": current_date.isoformat(),
                               "product_id": fp_id, "qty": qty})
                major_total += qty

        # 비주력
        n_minor = min(len(minor_fps), random.randint(10, 18))
        active_minor = random.sample(minor_fps, n_minor)
        minor_total = 0
        for fp_id in active_minor:
            remaining = minor_target - minor_total
            if remaining <= 0:
                break
            qty = random.randint(30, 120)
            qty = min(qty, remaining)
            if qty > 0:
                demand.append({"date": current_date.isoformat(),
                               "product_id": fp_id, "qty": qty})
                minor_total += qty

    return demand


def get_all_sample_data():
    bom = build_bom()
    demand = generate_demand(bom)
    color_matrix = build_color_transition_matrix()
    return {
        "injection_products": INJECTION_PRODUCTS,
        "colors": COLORS,
        "bom": bom,
        "demand": demand,
        "color_transition_matrix": color_matrix,
    }


if __name__ == "__main__":
    data = get_all_sample_data()
    bom = data["bom"]
    print(f"사출품: {len(data['injection_products'])}종")
    print(f"컬러: {len(data['colors'])}색")
    print(f"완성품: {len(bom)}종")

    # 품번 예시
    print("\n품번 예시:")
    for fp_id in sorted(bom.keys())[:10]:
        info = bom[fp_id]
        print(f"  {fp_id:15s} → 지그:{info['injection_product']} "
              f"컬러:{info['color']:10s} {info['name']}")

    # 일별 수요
    from collections import defaultdict
    daily = defaultdict(int)
    daily_major = defaultdict(int)
    for d in data["demand"]:
        daily[d["date"]] += d["qty"]
        if bom[d["product_id"]]["injection_product"] in MAJOR_PRODUCTS:
            daily_major[d["date"]] += d["qty"]
    print("\n일별 수요:")
    for dt in sorted(daily):
        t = daily[dt]
        m = daily_major.get(dt, 0)
        print(f"  {dt}: {t:,}개 (주력 {m:,} = {m/t*100:.0f}%)")
