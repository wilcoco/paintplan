"""
도장 생산계획 시스템 — 핵심 스케줄링 알고리즘 v3

컨베이어 모델:
  140 행어가 원형 순환, 행어당 지그 2개
  행어 사이클: 적재 → 도장 → 하차 → [지그교환?] → 적재 → ...
  1회전 = 같은 140행어가 도장 로봇 앞을 한 바퀴 통과
  10회전/일, 일일 최대 2,800개 (140×2×10)

알고리즘 순서:
  1단계: 지그 템플릿 설계 (140행어에 제품 배분)
  2단계: 세그먼트별 컬러×회전 수요 계산
  3단계: 컬러 순서 TSP (전환 비용 최소)
  4단계: 회전별 컬러 배정 (컬러 블록 방식)
  5단계: 회전 간 지그 부분 교체 (280건 예산 내)
  6단계: 전환 비용 및 생산량 계산
"""
import math
from collections import defaultdict
from config import (
    HANGER_COUNT, JIGS_PER_HANGER, ROTATIONS_PER_DAY,
    MAX_JIG_CHANGES_PER_DAY,
)


# ═══════════════════════════════════════════════════════════════
# 데이터 클래스
# ═══════════════════════════════════════════════════════════════

class Segment:
    """컨베이어 위의 한 구간 (같은 지그가 연속된 행어 묶음)"""
    def __init__(self, product, n_hangers, position_start):
        self.product = product
        self.n_hangers = n_hangers
        self.position_start = position_start  # 0-based 시작 위치
        self.pieces_per_rotation = n_hangers * JIGS_PER_HANGER


class Template:
    """140행어의 지그 배치"""
    def __init__(self, segments):
        self.segments = segments  # [Segment, ...]
        self.allocation = {s.product: s.n_hangers for s in segments}

    def describe(self):
        return " | ".join(
            f"{s.product}({s.n_hangers}h/{s.pieces_per_rotation}pcs)"
            for s in self.segments
        )

    def total_hangers(self):
        return sum(s.n_hangers for s in self.segments)


class RotationResult:
    """1회전 결과"""
    def __init__(self, rotation_num):
        self.rotation = rotation_num
        self.cells = []  # [(segment_idx, product, color, n_hangers, pieces)]
        self.color_transitions = []  # [(from_color, to_color, empty_hangers)]
        self.total_produced = 0
        self.total_empty = 0
        self.jig_changes_before = 0  # 이 회전 시작 전 지그 교체 건수


class DayResult:
    """하루 스케줄 결과"""
    def __init__(self, date_str):
        self.date = date_str
        self.template = None  # Template (시작 시점)
        self.template_changes = []  # [(rotation, old_template, new_template)]
        self.rotations = []  # [RotationResult, ...]
        self.color_grid = []  # [rotation][segment_idx] = color
        self.total_produced = 0
        self.total_empty = 0
        self.total_color_changes = 0
        self.total_jig_changes = 0
        self.demand = {}  # {(product, color): qty}
        self.fulfilled = {}  # {(product, color): qty}
        self.shortfall = {}  # {(product, color): qty}


# ═══════════════════════════════════════════════════════════════
# 1단계: 지그 템플릿 설계
# ═══════════════════════════════════════════════════════════════

def design_template(product_demand, prev_template=None, jig_budget=MAX_JIG_CHANGES_PER_DAY,
                    jig_limits=None):
    """
    140행어에 제품 지그를 배분 — 타입 기반

    지그 보유량으로 유효 타입(합=140) 열거 → 수요 커버율 최고 타입 선택

    Args:
        product_demand: {product: total_qty} (전 컬러 합산)
        prev_template: 전날 Template 또는 None
        jig_budget: 지그 교환 가용 예산
        jig_limits: {product: max_hangers} 지그 보유량 기반 행어 상한

    Returns:
        Template, jig_changes_used
    """
    total_qty = sum(product_demand.values())
    if total_qty == 0:
        return Template([]), 0

    if jig_limits is None:
        jig_limits = {}

    from type_generator import enumerate_valid_types

    # 유효 타입 열거 (합=140, 최소 세그먼트 20행어)
    valid_types = enumerate_valid_types(jig_limits, HANGER_COUNT, min_segment=20)

    if valid_types:
        # 각 타입 점수 계산: 수요 커버율 + 전날 유사도 보너스
        best_type = None
        best_score = -1

        for conv_type in valid_types:
            covered = sum(product_demand.get(pc, 0)
                          for pc in conv_type["products"])
            score = covered / max(1, total_qty)

            # 전날 타입과 유사하면 보너스 (지그 교환 절약)
            if prev_template:
                overlap = sum(
                    min(conv_type["products"].get(pc, 0),
                        prev_template.allocation.get(pc, 0))
                    for pc in set(list(conv_type["products"]) +
                                  list(prev_template.allocation))
                )
                score += overlap / HANGER_COUNT * 0.3

            if score > best_score:
                best_score = score
                best_type = conv_type

        raw_alloc = dict(best_type["products"])
    else:
        # 폴백: 수요 비례 배분 (타입 열거 불가 시)
        products_sorted = sorted(product_demand.items(), key=lambda x: -x[1])
        raw_alloc = {}
        remaining = HANGER_COUNT
        for i, (product, qty) in enumerate(products_sorted[:3]):
            max_h = jig_limits.get(product, HANGER_COUNT)
            if i == 2:
                h = remaining
            else:
                h = min(max_h, round(HANGER_COUNT * qty / total_qty))
                h = min(h, remaining)
            raw_alloc[product] = h
            remaining -= h

    # 전날 템플릿과 비교하여 지그 교환 계산
    jig_changes = 0
    final_alloc = raw_alloc

    if prev_template:
        prev_alloc = prev_template.allocation
        changes = _calc_jig_changes(prev_alloc, raw_alloc)

        if changes > jig_budget:
            # 예산 초과 시: 전날 배치를 최대한 유지하면서 조정
            final_alloc = _adjust_within_budget(prev_alloc, raw_alloc, jig_budget)
            changes = _calc_jig_changes(prev_alloc, final_alloc)

        jig_changes = changes

    # 세그먼트 생성 (컬러 유사성 기반 순서는 Step3 이후 적용)
    segments = []
    pos = 0
    for product, n_h in sorted(final_alloc.items(), key=lambda x: -x[1]):
        segments.append(Segment(product, n_h, pos))
        pos += n_h

    return Template(segments), jig_changes


def _calc_jig_changes(alloc_a, alloc_b):
    """두 배분 간 지그 교환 건수 (지그 단위)"""
    all_products = set(list(alloc_a.keys()) + list(alloc_b.keys()))
    # 늘어난 제품 = 새 지그 설치, 줄어든 제품 = 지그 제거
    # 총 교환 = 변경된 행어 수 × JIGS_PER_HANGER
    added = 0
    for p in all_products:
        a = alloc_a.get(p, 0)
        b = alloc_b.get(p, 0)
        if b > a:
            added += (b - a)
    return added * JIGS_PER_HANGER


def _adjust_within_budget(prev_alloc, ideal_alloc, budget):
    """예산 내에서 전날→이상적 배분으로 최대한 이동"""
    result = dict(prev_alloc)
    all_products = set(list(prev_alloc.keys()) + list(ideal_alloc.keys()))

    # 변경 필요량 계산
    changes_needed = {}
    for p in all_products:
        prev = prev_alloc.get(p, 0)
        ideal = ideal_alloc.get(p, 0)
        if ideal != prev:
            changes_needed[p] = ideal - prev  # +면 늘려야, -면 줄여야

    # 우선순위: 수요가 큰 제품 먼저
    budget_remaining = budget // JIGS_PER_HANGER  # 행어 단위로 변환

    # 늘려야 할 제품과 줄여야 할 제품 매칭
    to_increase = sorted(
        [(p, d) for p, d in changes_needed.items() if d > 0],
        key=lambda x: -x[1]
    )
    to_decrease = sorted(
        [(p, -d) for p, d in changes_needed.items() if d < 0],
        key=lambda x: -x[1]
    )

    inc_idx, dec_idx = 0, 0
    while inc_idx < len(to_increase) and dec_idx < len(to_decrease) and budget_remaining > 0:
        inc_p, inc_need = to_increase[inc_idx]
        dec_p, dec_avail = to_decrease[dec_idx]

        move = min(inc_need, dec_avail, budget_remaining)
        result[inc_p] = result.get(inc_p, 0) + move
        result[dec_p] = result.get(dec_p, 0) - move
        budget_remaining -= move

        to_increase[inc_idx] = (inc_p, inc_need - move)
        to_decrease[dec_idx] = (dec_p, dec_avail - move)

        if to_increase[inc_idx][1] == 0:
            inc_idx += 1
        if to_decrease[dec_idx][1] == 0:
            dec_idx += 1

    # 0인 제품 제거
    result = {p: h for p, h in result.items() if h > 0}
    return result


# ═══════════════════════════════════════════════════════════════
# 2단계: 세그먼트별 컬러 수요 → 회전 수 변환
# ═══════════════════════════════════════════════════════════════

def compute_segment_color_rotations(template, product_color_demand):
    """
    각 세그먼트가 각 컬러로 몇 회전 필요한지 계산

    Returns:
        {seg_idx: {color: n_rotations}}
        {seg_idx: {color: exact_qty}}
    """
    rotation_needs = {}
    exact_needs = {}

    for seg_idx, seg in enumerate(template.segments):
        product = seg.product
        colors = product_color_demand.get(product, {})
        seg_rotations = {}
        seg_exact = {}

        for color, qty in colors.items():
            if qty > 0:
                rotations = math.ceil(qty / seg.pieces_per_rotation)
                seg_rotations[color] = rotations
                seg_exact[color] = qty

        rotation_needs[seg_idx] = seg_rotations
        exact_needs[seg_idx] = seg_exact

    return rotation_needs, exact_needs


# ═══════════════════════════════════════════════════════════════
# 3단계: 컬러 순서 TSP
# ═══════════════════════════════════════════════════════════════

def solve_color_tsp(colors, color_matrix, start_color=None):
    """nearest-neighbor TSP로 컬러 전환 최소 순서 결정"""
    if len(colors) <= 2:
        return list(colors)

    best_order = None
    best_cost = float('inf')

    starts = [start_color] if start_color and start_color in colors else colors

    for start in starts:
        order = [start]
        remaining = set(colors) - {start}
        total_cost = 0
        current = start

        while remaining:
            nearest = min(remaining,
                          key=lambda c: color_matrix.get((current, c), 6))
            total_cost += color_matrix.get((current, nearest), 6)
            order.append(nearest)
            remaining.remove(nearest)
            current = nearest

        if total_cost < best_cost:
            best_cost = total_cost
            best_order = order

    return best_order


# ═══════════════════════════════════════════════════════════════
# 4단계: 회전별 컬러 배정 (컬러 블록 방식)
# ═══════════════════════════════════════════════════════════════

def order_segments_by_color_similarity(template, product_color_demand):
    """
    컬러 프로필이 비슷한 세그먼트를 인접 배치

    같은 컬러를 공유하는 제품끼리 가까이 놓으면
    회전 내 컬러 전환이 줄어듦
    """
    segments = template.segments
    n = len(segments)
    if n <= 2:
        return segments

    # 각 세그먼트의 컬러 집합
    seg_colors = {}
    for i, seg in enumerate(segments):
        colors = set(product_color_demand.get(seg.product, {}).keys())
        seg_colors[i] = colors

    # 유사도 매트릭스 (자카드 유사도)
    def similarity(i, j):
        ci, cj = seg_colors[i], seg_colors[j]
        if not ci or not cj:
            return 0
        return len(ci & cj) / len(ci | cj)

    # Nearest-neighbor로 순서 결정
    remaining = set(range(n))
    # 가장 많은 컬러를 가진 세그먼트부터 시작
    start = max(remaining, key=lambda i: len(seg_colors[i]))
    order = [start]
    remaining.remove(start)

    while remaining:
        current = order[-1]
        nearest = max(remaining, key=lambda i: similarity(current, i))
        order.append(nearest)
        remaining.remove(nearest)

    # 재배치된 세그먼트 리스트 (위치 재계산)
    new_segments = []
    pos = 0
    for idx in order:
        old = segments[idx]
        new_segments.append(Segment(old.product, old.n_hangers, pos))
        pos += old.n_hangers

    return new_segments


def assign_color_grid(template, rotation_needs, color_matrix,
                      prev_last_colors=None, n_rotations=None):
    """
    grid[rotation][segment_idx] = color 배정

    Args:
        n_rotations: 회전 수 (None이면 ROTATIONS_PER_DAY 사용)
    """
    n_seg = len(template.segments)
    n_rot = n_rotations or ROTATIONS_PER_DAY
    grid = [[None] * n_seg for _ in range(n_rot)]

    if n_seg == 0:
        return grid

    # 각 세그먼트별 잔여 수요 (회전 수)
    remaining = {}
    for seg_idx in range(n_seg):
        remaining[seg_idx] = dict(rotation_needs.get(seg_idx, {}))

    # 전체 컬러 수집
    all_colors = set()
    for seg_idx in range(n_seg):
        all_colors.update(remaining[seg_idx].keys())

    if not all_colors:
        return grid

    # TSP로 컬러 순서 결정 (컬러 전환 비용 최소)
    color_order = solve_color_tsp(list(all_colors), color_matrix)

    # 각 회전마다 지배 컬러 방식으로 배정
    for r in range(n_rot):
        # 이 회전에서 각 컬러를 사용할 수 있는 세그먼트 수 + 총 수요 계산
        color_scores = {}
        for color in all_colors:
            segs_available = []
            total_demand = 0
            for seg_idx in range(n_seg):
                need = remaining[seg_idx].get(color, 0)
                if need > 0:
                    segs_available.append(seg_idx)
                    total_demand += need
            if segs_available:
                # 점수 = (커버하는 세그먼트 수) × 100 + 총 잔여 수요
                # 세그먼트 수를 우선, 동점이면 수요가 큰 컬러
                color_scores[color] = (len(segs_available), total_demand)

        if not color_scores:
            break

        # 가장 많은 세그먼트를 커버하는 컬러부터 배정
        sorted_colors = sorted(
            color_scores.keys(),
            key=lambda c: (-color_scores[c][0], -color_scores[c][1])
        )

        assigned = [False] * n_seg

        for color in sorted_colors:
            for seg_idx in range(n_seg):
                if assigned[seg_idx]:
                    continue
                need = remaining[seg_idx].get(color, 0)
                if need > 0:
                    grid[r][seg_idx] = color
                    remaining[seg_idx][color] -= 1
                    assigned[seg_idx] = True

        # 미배정 세그먼트: 인접 세그먼트 컬러와 동일하게 (수요가 있으면)
        for seg_idx in range(n_seg):
            if assigned[seg_idx]:
                continue
            # 인접 세그먼트의 컬러 확인
            neighbor_color = None
            if seg_idx > 0 and grid[r][seg_idx - 1]:
                neighbor_color = grid[r][seg_idx - 1]
            elif seg_idx < n_seg - 1 and grid[r][seg_idx + 1]:
                neighbor_color = grid[r][seg_idx + 1]

            # 이전 회전의 같은 세그먼트 컬러
            prev_color = grid[r - 1][seg_idx] if r > 0 else None

            # 잔여 수요 중 선택
            seg_remaining = {c: n for c, n in remaining[seg_idx].items() if n > 0}
            if seg_remaining:
                # 인접 컬러와 같은 것 우선
                if neighbor_color and neighbor_color in seg_remaining:
                    chosen = neighbor_color
                elif prev_color and prev_color in seg_remaining:
                    chosen = prev_color
                else:
                    chosen = max(seg_remaining, key=seg_remaining.get)
                grid[r][seg_idx] = chosen
                remaining[seg_idx][chosen] -= 1
            elif neighbor_color:
                grid[r][seg_idx] = neighbor_color
            elif prev_color:
                grid[r][seg_idx] = prev_color

    return grid


# ═══════════════════════════════════════════════════════════════
# 5단계: 회전 간 지그 부분 교체
# ═══════════════════════════════════════════════════════════════

def _calc_swap_cost(alloc_from, alloc_to):
    """두 타입 간 지그 교체 비용 (제거+장착 행어의 지그 수)"""
    cost = 0
    all_pcs = set(list(alloc_from) + list(alloc_to))
    for pc in all_pcs:
        a = alloc_from.get(pc, 0)
        b = alloc_to.get(pc, 0)
        diff = abs(b - a)
        cost += diff * JIGS_PER_HANGER
    return cost


def plan_multi_type_day(product_demand, product_color_demand, jig_limits,
                        color_matrix, jig_budget, prev_template, prev_last_colors):
    """
    하루를 여러 타입 phase로 나눠 지그 교체하며 더 많은 차종 커버.
    오버랩 최대화로 swap cost를 낮춰 3~4 phase 가능.

    예: 회전1~5는 타입A, 회전6~10은 타입B
    타입A→B 전환 시 지그 교체 비용 발생 (280건 한도 내)

    Returns:
        phases: [(template, grid_rows, rotations_start, rotations_end)]
        total_jig_changes
    """
    from type_generator import enumerate_valid_types

    valid_types = enumerate_valid_types(jig_limits, HANGER_COUNT, min_segment=20)
    if not valid_types:
        return None, 0

    total_demand = sum(product_demand.values())
    if total_demand == 0:
        return None, 0

    # 수요가 있는 차종
    demanded_products = {pc for pc, qty in product_demand.items() if qty > 0}

    # 1타입으로 전 차종 커버 가능한지 확인
    for t in valid_types:
        if demanded_products <= set(t["products"].keys()):
            return None, 0  # 단일 타입 충분

    # ── Greedy multi-phase: 오버랩 최대화, swap cost 최소화 ──
    # 1단계: 수요 커버 가장 높은 타입을 첫 phase로
    type_by_coverage = sorted(
        valid_types,
        key=lambda t: -sum(product_demand.get(pc, 0) for pc in t["products"])
    )
    phases = []
    covered_products = set()
    remaining_budget = jig_budget
    remaining_rotations = ROTATIONS_PER_DAY
    prev_alloc = None  # 직전 phase의 product allocation

    while remaining_rotations >= 2 and type_by_coverage:
        # 다음 phase 타입 선택
        best_type = None
        best_score = -1

        for t in type_by_coverage:
            # 이 타입이 새로 커버하는 차종의 수요
            new_covered = sum(
                product_demand.get(pc, 0)
                for pc in t["products"]
                if pc not in covered_products
            )
            if new_covered == 0 and covered_products:
                continue  # 새로운 차종 없음 → 스킵

            # swap cost 계산
            if prev_alloc is not None:
                swap = _calc_swap_cost(prev_alloc, t["products"])
            else:
                swap = 0

            if swap > remaining_budget:
                continue  # 예산 초과

            # 점수: 새 커버 수요 - swap 패널티
            score = new_covered * 10 - swap
            if score > best_score:
                best_score = score
                best_type = t

        if best_type is None:
            break

        # swap cost 차감
        if prev_alloc is not None:
            swap = _calc_swap_cost(prev_alloc, best_type["products"])
            remaining_budget -= swap
        else:
            swap = 0

        # 회전 수 결정: 이 phase가 커버하는 수요 비례
        phase_demand = sum(product_demand.get(pc, 0) for pc in best_type["products"])
        if total_demand > 0:
            ideal_rots = max(2, round(remaining_rotations * phase_demand / max(1, total_demand)))
        else:
            ideal_rots = remaining_rotations
        ideal_rots = min(ideal_rots, remaining_rotations)

        # 마지막 phase면 남은 회전 전부
        phases.append({"type": best_type, "rotations": ideal_rots})
        covered_products.update(best_type["products"].keys())
        remaining_rotations -= ideal_rots
        prev_alloc = best_type["products"]

        # 커버된 수요 제거 (다음 phase 점수 계산용)
        total_demand -= phase_demand

        # 전 차종 커버되면 종료
        if demanded_products <= covered_products:
            break

    if len(phases) <= 1:
        return None, 0  # single phase → 기존 로직 사용

    # 남은 회전이 있으면 마지막 phase에 추가
    if remaining_rotations > 0 and phases:
        phases[-1]["rotations"] += remaining_rotations

    total_swap = jig_budget - remaining_budget
    return {"phases": phases, "swap_cost": total_swap}, total_swap


# ═══════════════════════════════════════════════════════════════
# 6단계: 결과 계산
# ═══════════════════════════════════════════════════════════════

def calculate_day_result(date_str, template, grid, daily_paint_items,
                         color_matrix, jig_changes_for_template):
    """전환 비용, 생산량, 미충족 등 계산"""
    day = DayResult(date_str)
    day.template = template
    day.color_grid = grid
    day.demand = dict(daily_paint_items)
    day.total_jig_changes = jig_changes_for_template

    n_seg = len(template.segments)
    n_rot = ROTATIONS_PER_DAY
    produced = defaultdict(int)  # (product, color) → qty

    prev_rotation_last_color = None  # 이전 회전의 마지막 컬러

    for r in range(n_rot):
        rot = RotationResult(r + 1)

        # 이 회전의 각 세그먼트 처리
        for seg_idx in range(n_seg):
            seg = template.segments[seg_idx]
            color = grid[r][seg_idx]
            if color is None:
                continue

            pieces = seg.pieces_per_rotation
            # 컨베이어 10회전 무조건 가동 — 수요 초과 시에도 풀 생산 (안전재고)
            key = (seg.product, color)
            actual = pieces  # 항상 풀 생산

            rot.cells.append((seg_idx, seg.product, color, seg.n_hangers, actual))
            rot.total_produced += actual
            produced[key] += actual

        # 컬러 전환 계산 (이 회전 내 세그먼트 간)
        colors_in_rotation = [grid[r][s] for s in range(n_seg) if grid[r][s]]
        for i in range(1, len(colors_in_rotation)):
            c_from = colors_in_rotation[i - 1]
            c_to = colors_in_rotation[i]
            if c_from != c_to:
                empty = color_matrix.get((c_from, c_to), 6)
                rot.color_transitions.append((c_from, c_to, empty))
                rot.total_empty += empty

        # 회전 간 전환 (이전 회전 마지막 → 이 회전 첫 번째)
        if colors_in_rotation and prev_rotation_last_color is not None:
            first_color = colors_in_rotation[0]
            if first_color != prev_rotation_last_color:
                empty = color_matrix.get((prev_rotation_last_color, first_color), 6)
                rot.color_transitions.insert(0, (prev_rotation_last_color, first_color, empty))
                rot.total_empty += empty

        if colors_in_rotation:
            prev_rotation_last_color = colors_in_rotation[-1]

        day.rotations.append(rot)

    # 집계
    day.total_produced = sum(r.total_produced for r in day.rotations)
    day.total_empty = sum(r.total_empty for r in day.rotations)
    day.total_color_changes = sum(len(r.color_transitions) for r in day.rotations)
    day.fulfilled = dict(produced)

    for key, demand_qty in daily_paint_items.items():
        fulfilled = produced.get(key, 0)
        if fulfilled < demand_qty:
            day.shortfall[key] = demand_qty - fulfilled

    return day


# ═══════════════════════════════════════════════════════════════
# 전체 12일 스케줄링
# ═══════════════════════════════════════════════════════════════

def schedule_painting(daily_paint_summary, color_matrix, jig_limits=None):
    """
    12일 도장 스케줄 생성

    Args:
        daily_paint_summary: {date: {(product, color): qty}}
        color_matrix: {(c_from, c_to): empty_hangers}
        jig_limits: {product: max_hangers} 지그 보유량 기반 행어 상한

    Returns:
        [DayResult, ...]
    """
    results = []
    prev_template = None
    prev_last_colors = None
    jig_budget = MAX_JIG_CHANGES_PER_DAY

    for date_str in sorted(daily_paint_summary.keys()):
        items = daily_paint_summary[date_str]

        # 제품별 총 수요 (전 컬러 합산)
        product_total = defaultdict(int)
        product_color = defaultdict(lambda: defaultdict(int))
        for (product, color), qty in items.items():
            product_total[product] += qty
            product_color[product][color] += qty

        # 멀티타입 시도: 하루를 2 phase로 나눌 수 있는지
        multi, swap_cost = plan_multi_type_day(
            dict(product_total), dict(product_color), jig_limits or {},
            color_matrix, jig_budget, prev_template, prev_last_colors,
        )

        if multi and multi.get("phases"):
            # 멀티타입: phase별로 독립 스케줄 → 결과만 합산
            phases = multi["phases"]
            day = DayResult(date_str)
            day.demand = dict(items)
            day.total_jig_changes = swap_cost
            produced = defaultdict(int)

            for phase in phases:
                p_type = phase["type"]
                p_rots = phase["rotations"]
                p_alloc = p_type["products"]

                # phase 템플릿
                segments = []
                pos = 0
                for pc, nh in sorted(p_alloc.items(), key=lambda x: -x[1]):
                    segments.append(Segment(pc, nh, pos))
                    pos += nh
                p_template = Template(segments)

                new_segs = order_segments_by_color_similarity(p_template, dict(product_color))
                p_template = Template(new_segs)

                p_rotation_needs, _ = compute_segment_color_rotations(
                    p_template, dict(product_color)
                )

                p_grid = assign_color_grid(
                    p_template, p_rotation_needs, color_matrix,
                    prev_last_colors, n_rotations=p_rots
                )

                # phase 결과를 회전별로 추가
                for r_idx, grid_row in enumerate(p_grid):
                    rot = RotationResult(len(day.rotations) + 1)
                    n_seg = len(p_template.segments)
                    colors_in_rot = []

                    for seg_idx in range(n_seg):
                        seg = p_template.segments[seg_idx]
                        color = grid_row[seg_idx] if seg_idx < len(grid_row) else None
                        if not color:
                            continue
                        key = (seg.product, color)
                        actual = seg.pieces_per_rotation  # 풀 생산
                        rot.cells.append((seg_idx, seg.product, color, seg.n_hangers, actual))
                        rot.total_produced += actual
                        produced[key] += actual
                        colors_in_rot.append(color)

                    # 컬러 전환
                    for i in range(1, len(colors_in_rot)):
                        if colors_in_rot[i] != colors_in_rot[i-1]:
                            empty = color_matrix.get((colors_in_rot[i-1], colors_in_rot[i]), 6)
                            rot.color_transitions.append((colors_in_rot[i-1], colors_in_rot[i], empty))
                            rot.total_empty += empty

                    day.rotations.append(rot)

            day.template = Template([Segment(pc, nh, 0) for pc, nh in
                                     sorted(phases[0]["type"]["products"].items(),
                                            key=lambda x: -x[1])])
            day.color_grid = []  # multi에서는 phase별로 다르므로 비워둠
            day.total_produced = sum(r.total_produced for r in day.rotations)
            day.total_empty = sum(r.total_empty for r in day.rotations)
            day.total_color_changes = sum(len(r.color_transitions) for r in day.rotations)
            day.fulfilled = dict(produced)
            for key, dq in items.items():
                if produced.get(key, 0) < dq:
                    day.shortfall[key] = dq - produced.get(key, 0)

            results.append(day)
            prev_template = day.template
            jig_changes = swap_cost
            # skip to next day
            continue

        else:
            # 단일 타입
            template, jig_changes = design_template(
                dict(product_total), prev_template, jig_budget,
                jig_limits=jig_limits,
            )

            # 세그먼트별 컬러 회전 수요
            rotation_needs, exact_needs = compute_segment_color_rotations(
                template, dict(product_color)
            )

            # 세그먼트 순서 최적화
            new_segments = order_segments_by_color_similarity(template, dict(product_color))
            template = Template(new_segments)
            rotation_needs, exact_needs = compute_segment_color_rotations(
                template, dict(product_color)
            )

            # 컬러 그리드
            grid = assign_color_grid(
                template, rotation_needs, color_matrix, prev_last_colors
            )

        # 6단계: 결과 계산
        day = calculate_day_result(
            date_str, template, grid, items,
            color_matrix, jig_changes
        )

        results.append(day)

        # 다음 날 참조용
        prev_template = template
        if grid and len(grid) > 0 and len(template.segments) > 0:
            prev_last_colors = {
                seg_idx: grid[-1][seg_idx]
                for seg_idx in range(len(template.segments))
                if grid[-1][seg_idx] is not None
            }

    return results


# ═══════════════════════════════════════════════════════════════
# 출력 함수
# ═══════════════════════════════════════════════════════════════

def print_schedule_summary(results):
    """12일 요약 출력"""
    print(f"\n{'═' * 95}")
    print("도장 생산계획 v3 — 12일 요약")
    print(f"{'═' * 95}")

    for day in results:
        demand_total = sum(day.demand.values())
        short_total = sum(day.shortfall.values()) if day.shortfall else 0
        cap_pct = day.total_produced / (HANGER_COUNT * JIGS_PER_HANGER * ROTATIONS_PER_DAY) * 100

        print(f"\n  {day.date} │ 수요 {demand_total:,} │ "
              f"생산 {day.total_produced:,} ({cap_pct:.0f}%) │ "
              f"미충족 {short_total:,} │ "
              f"컬러전환 {day.total_color_changes}회 (빈행어 {day.total_empty}) │ "
              f"지그교환 {day.total_jig_changes}건")

        # 지그 템플릿
        print(f"    지그: {day.template.describe()}")

    # 합계
    print(f"\n{'─' * 95}")
    n = len(results)
    total_demand = sum(sum(d.demand.values()) for d in results)
    total_prod = sum(d.total_produced for d in results)
    total_short = sum(sum(d.shortfall.values()) for d in results)
    total_color = sum(d.total_color_changes for d in results)
    total_empty = sum(d.total_empty for d in results)
    total_jig = sum(d.total_jig_changes for d in results)

    print(f"  12일 합계: 수요 {total_demand:,} │ 생산 {total_prod:,} │ "
          f"미충족 {total_short:,}")
    print(f"  컬러전환: {total_color}회 (일평균 {total_color/n:.1f}) │ "
          f"빈행어 손실: {total_empty}개")
    print(f"  지그교환: {total_jig}건 (일평균 {total_jig/n:.1f}, "
          f"한도 {MAX_JIG_CHANGES_PER_DAY}/일)")
    print(f"{'═' * 95}")


def print_day_rotations(day, max_rotations=None):
    """
    하루 회전별 140행어 상세 출력

    각 회전에서 140행어가 어떤 지그+컬러로 구성되는지 보여줌
    """
    print(f"\n{'━' * 95}")
    print(f"  {day.date} — 회전별 140행어 상세")
    print(f"{'━' * 95}")
    print(f"  지그 템플릿: {day.template.describe()}")

    n_rot = max_rotations or len(day.rotations)

    for r_idx in range(min(n_rot, len(day.rotations))):
        rot = day.rotations[r_idx]
        grid_row = day.color_grid[r_idx] if r_idx < len(day.color_grid) else []

        print(f"\n  ┌─ 회전 {rot.rotation:2d} ── 생산 {rot.total_produced:,}개 "
              f"── 컬러전환 {len(rot.color_transitions)}회 "
              f"(빈행어 {rot.total_empty}) ──┐")

        # 140행어 시각화: 세그먼트별로 표시
        if not rot.cells:
            print(f"  │  (생산 없음)")
        else:
            for seg_idx, product, color, n_hangers, pieces in rot.cells:
                seg = day.template.segments[seg_idx]
                pos_start = seg.position_start + 1  # 1-based
                pos_end = seg.position_start + seg.n_hangers
                bar_len = max(1, seg.n_hangers // 3)
                bar = "█" * bar_len
                print(f"  │  행어 {pos_start:3d}~{pos_end:3d} │ "
                      f"지그:{product:12s} │ 컬러:{color} │ "
                      f"{n_hangers:3d}행어 → {pieces:3d}개 {bar}")

        # 컬러 전환 표시
        if rot.color_transitions:
            transitions_str = ", ".join(
                f"{f}→{t}({e}행어)" for f, t, e in rot.color_transitions
            )
            print(f"  │  전환: {transitions_str}")

        print(f"  └{'─' * 88}┘")


def print_jig_type_analysis(results):
    """지그 유형 분석 — 12일간 사용된 지그 템플릿 패턴"""
    print(f"\n{'═' * 95}")
    print("지그 유형 분석")
    print(f"{'═' * 95}")

    templates = []
    for day in results:
        key = tuple(
            (s.product, s.n_hangers)
            for s in sorted(day.template.segments, key=lambda x: x.product)
        )
        templates.append((day.date, key, day.template))

    # 유형별 그룹핑
    from collections import Counter
    type_counter = Counter(t[1] for t in templates)
    unique_types = list(type_counter.keys())

    print(f"\n  12일간 {len(unique_types)}개 고유 지그 유형 사용\n")

    for i, (tkey, count) in enumerate(type_counter.most_common(), 1):
        dates = [t[0] for t in templates if t[1] == tkey]
        tmpl = next(t[2] for t in templates if t[1] == tkey)

        print(f"  유형 {i}: {count}일 사용 ({', '.join(dates[:4])}"
              f"{'...' if len(dates) > 4 else ''})")

        for seg in sorted(tmpl.segments, key=lambda s: -s.n_hangers):
            pct = seg.n_hangers / HANGER_COUNT * 100
            daily_cap = seg.pieces_per_rotation * ROTATIONS_PER_DAY
            print(f"    {seg.product}: {seg.n_hangers}행어 ({pct:.0f}%) "
                  f"→ 최대 {daily_cap}개/일")
        print()

    # 일별 지그 교환
    print(f"  일별 지그 교환:")
    for day in results:
        bar = "█" * (day.total_jig_changes // 10) if day.total_jig_changes > 0 else "·"
        over = " ⚠ 초과!" if day.total_jig_changes > MAX_JIG_CHANGES_PER_DAY else ""
        print(f"    {day.date}: {day.total_jig_changes:3d}건 "
              f"(한도 {MAX_JIG_CHANGES_PER_DAY}) {bar}{over}")
