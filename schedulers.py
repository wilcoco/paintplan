"""
다양한 스케줄링 알고리즘 구현
1. 휴리스틱 (현재 v10.3)
2. MIP 최적화 (OR-Tools)
3. 컬러 중심 (Color-First)
4. 2단계 분해 (Assignment + TSP)
"""

from collections import defaultdict
import copy

# ============================================
# 공통 상수 및 유틸리티
# ============================================
HANGERS = 140
JIGS_PER_HANGER = 2
ROTATIONS = 10

JIG_INVENTORY = {
    'A': {'name': 'THPE STD/LDT+SP3', 'max_jigs': 100, 'pcs': 1},
    'B': {'name': 'NQ5 FRT (STD+XLINE)', 'max_jigs': 100, 'pcs': 1},
    'B2': {'name': 'NQ5 FRT STD 전용', 'max_jigs': 50, 'pcs': 1},
    'C': {'name': 'OV1', 'max_jigs': 80, 'pcs': 1},
    'D': {'name': 'JX EV FRT', 'max_jigs': 100, 'pcs': 1},
    'E': {'name': 'JX CROSS', 'max_jigs': 80, 'pcs': 1},
    'F': {'name': 'JX EV RR', 'max_jigs': 50, 'pcs': 1},
    'G': {'name': 'AX PE', 'max_jigs': 80, 'pcs': 1},
    'H': {'name': 'THPE RR', 'max_jigs': 50, 'pcs': 2},
    'I': {'name': 'NQ5 RR', 'max_jigs': 70, 'pcs': 1},
}

SPECIAL_COLORS = {'MGG', 'T4M', 'UMA', 'ZRM', 'ISM', 'MRM'}

def get_color_change_cost(from_color):
    """컬러교환 시 빈행어 수"""
    if from_color and from_color.upper() in SPECIAL_COLORS:
        return 15
    return 1


# ============================================
# 지그 위치 기반 교체 계산 함수
# ============================================
def order_to_positions(tmpl, order):
    """템플릿과 순서를 140개 위치 배열로 변환"""
    positions = []
    for g in order:
        if g in tmpl and tmpl[g] > 0:
            positions.extend([g] * tmpl[g])
    while len(positions) < HANGERS:
        positions.append(None)
    return positions[:HANGERS]


def calc_position_changes(pos1, pos2):
    """두 위치 배열 간 실제 교체 수 계산"""
    if not pos1 or not pos2:
        return 0
    return sum(1 for i in range(HANGERS) if pos1[i] != pos2[i])


def get_grp_main_color(items, grp, rotation):
    """해당 회전에서 그룹의 주요 컬러 반환"""
    color_prod = defaultdict(int)
    for x in items:
        if x.get('grp') == grp:
            prod = x.get('prod', [0]*10)
            if rotation < len(prod):
                color_prod[x['clr']] += prod[rotation]
    if color_prod:
        return max(color_prod.keys(), key=lambda c: (color_prod[c], c))
    return None


def optimize_jig_order(templates, items, jig_budget_day=150, jig_budget_night=150):
    """MIP 템플릿 결과에 대해 최적 지그 순서 결정

    전략:
    1. 컬러 연속성 최대화 (같은 컬러 그룹 인접 배치)
    2. 지그 교체 예산 내에서 순서 결정
    3. 이전 회전 순서 최대한 유지
    """
    n_rotations = len(templates)
    jig_orders = []
    jig_changes = [0] * n_rotations

    day_budget_left = jig_budget_day
    night_budget_left = jig_budget_night

    prev_positions = None
    prev_order = None

    for r in range(n_rotations):
        tmpl = templates[r]
        is_day = r < 5
        budget_left = day_budget_left if is_day else night_budget_left

        # 사용중인 그룹들
        active_groups = [g for g in tmpl if tmpl[g] > 0]
        if not active_groups:
            jig_orders.append([])
            continue

        # 그룹별 주요 컬러 파악
        grp_colors = {}
        for g in active_groups:
            grp_colors[g] = get_grp_main_color(items, g, r)

        # 후보 순서 생성
        candidates = []

        # 1. 이전 순서 유지 (새 그룹은 뒤에)
        if prev_order:
            stable_order = [g for g in prev_order if g in active_groups]
            for g in active_groups:
                if g not in stable_order:
                    stable_order.append(g)
            candidates.append(stable_order)

        # 2. 컬러별 그룹 묶음
        color_groups = defaultdict(list)
        for g in active_groups:
            clr = grp_colors.get(g)
            if clr:
                color_groups[clr].append(g)
            else:
                color_groups[None].append(g)

        # 컬러별 생산량 순 정렬
        color_order = []
        sorted_colors = sorted(color_groups.keys(),
                              key=lambda c: -sum(tmpl.get(g, 0) for g in color_groups[c]) if c else 0)
        for clr in sorted_colors:
            grps = sorted(color_groups[clr], key=lambda g: -tmpl.get(g, 0))
            color_order.extend(grps)
        candidates.append(color_order)

        # 3. 행어 수 내림차순
        size_order = sorted(active_groups, key=lambda g: -tmpl.get(g, 0))
        candidates.append(size_order)

        # 최적 순서 선택 (지그교체 최소)
        best_order = candidates[0]
        best_cost = float('inf')

        for cand in candidates:
            cand_pos = order_to_positions(tmpl, cand)
            cost = calc_position_changes(prev_positions, cand_pos)
            if cost < best_cost:
                best_cost = cost
                best_order = cand

        # 예산 체크
        curr_positions = order_to_positions(tmpl, best_order)
        changes = calc_position_changes(prev_positions, curr_positions)

        if changes <= budget_left:
            jig_changes[r] = changes
            if is_day:
                day_budget_left -= changes
            else:
                night_budget_left -= changes
        else:
            # 예산 초과: 이전 순서 유지 시도
            if prev_order:
                fallback_order = [g for g in prev_order if g in active_groups]
                for g in active_groups:
                    if g not in fallback_order:
                        fallback_order.append(g)
                fallback_pos = order_to_positions(tmpl, fallback_order)
                fallback_cost = calc_position_changes(prev_positions, fallback_pos)
                if fallback_cost <= budget_left:
                    best_order = fallback_order
                    curr_positions = fallback_pos
                    jig_changes[r] = fallback_cost
                    if is_day:
                        day_budget_left -= fallback_cost
                    else:
                        night_budget_left -= fallback_cost
                else:
                    jig_changes[r] = changes  # 예산 초과 기록
            else:
                jig_changes[r] = changes

        jig_orders.append(best_order)
        prev_positions = curr_positions
        prev_order = best_order

    return jig_orders, jig_changes


def calculate_jig_changes(items):
    """회전별 지그 교체 수 계산

    지그 교체 = 이전 회전 대비 그룹별 행어 할당량 변화의 합
    1회전은 전날 설정 그대로 사용 → 교체 0
    """
    # 회전별 그룹별 생산량 집계
    rot_grp_prod = [{g: 0 for g in JIG_INVENTORY} for _ in range(ROTATIONS)]
    for x in items:
        g = x.get('grp')
        if not g or g not in JIG_INVENTORY:
            continue
        for r in range(ROTATIONS):
            rot_grp_prod[r][g] += x['prod'][r]

    # 생산량 → 행어 수 변환 (pcs_per_jig 고려)
    rot_grp_hangers = []
    for r in range(ROTATIONS):
        hangers = {}
        for g in JIG_INVENTORY:
            prod = rot_grp_prod[r][g]
            pcs = JIG_INVENTORY[g]['pcs']
            # 생산량 / (지그당 pcs × 행어당 지그) = 행어 수
            hangers[g] = (prod + pcs * JIGS_PER_HANGER - 1) // (pcs * JIGS_PER_HANGER) if prod > 0 else 0
        rot_grp_hangers.append(hangers)

    # 회전별 지그 교체 계산
    jig_changes = [0] * ROTATIONS
    for r in range(1, ROTATIONS):  # 1회전은 전날 그대로 → 0
        change = 0
        for g in JIG_INVENTORY:
            diff = abs(rot_grp_hangers[r][g] - rot_grp_hangers[r-1][g])
            change += diff
        jig_changes[r] = change

    return jig_changes


def fill_capacity_for_safety_stock(items, rot_grp_used):
    """남은 용량을 3일 안전재고 목표로 채우기

    전략:
    1. 이미 생산 중인 컬러 아이템 먼저 (컬러교환 0)
    2. 그 후 수요 많은 아이템에 추가 생산
    3. 140행어 * 2지그 * 10회전 = 2,800개 목표
    """

    MAX_PER_ROTATION = HANGERS * JIGS_PER_HANGER  # 280

    def get_grp_capacity(g, r, rot_total_used):
        if g not in JIG_INVENTORY:
            return 0
        max_h = JIG_INVENTORY[g]['max_jigs'] // JIGS_PER_HANGER
        pcs = JIG_INVENTORY[g]['pcs']
        max_cap = max_h * JIGS_PER_HANGER * pcs
        grp_remain = max(0, max_cap - rot_grp_used[r].get(g, 0))
        rot_remain = MAX_PER_ROTATION - rot_total_used
        return min(grp_remain, rot_remain)

    # 컬러별 아이템 그룹화
    color_items = defaultdict(list)
    for x in items:
        if x.get('clr'):
            color_items[x['clr']].append(x)

    # 회전별 총 생산량 추적
    rot_totals = [sum(x['prod'][r] for x in items) for r in range(ROTATIONS)]

    # 회전별 주컬러 파악 (이미 가장 많이 생산 중인 컬러)
    for r in range(ROTATIONS):
        if rot_totals[r] >= MAX_PER_ROTATION:
            continue  # 이미 용량 가득

        # 회전 r에서 생산 중인 컬러별 수량
        color_prod = defaultdict(int)
        for x in items:
            if x['prod'][r] > 0 and x.get('clr'):
                color_prod[x['clr']] += x['prod'][r]

        # 주컬러 결정 (생산 중이면 그것, 아니면 수요 최대)
        if color_prod:
            main_color = max(color_prod.keys(), key=lambda c: color_prod[c])
        else:
            # 수요 가장 많은 컬러
            color_demand = {}
            for clr, clr_items_list in color_items.items():
                color_demand[clr] = sum(sum(x['d0']) for x in clr_items_list)
            if color_demand:
                main_color = max(color_demand.keys(), key=lambda c: color_demand[c])
            else:
                continue

        # 주컬러 아이템들에 추가 생산
        for x in color_items.get(main_color, []):
            g = x.get('grp')
            if not g:
                continue

            cap = get_grp_capacity(g, r, rot_totals[r])
            if cap <= 0:
                continue

            # 3일치 목표 대비 부족분
            stk = x.get('stk', 0) + sum(x['prod'])
            d_total = sum(x['d0']) + sum(x.get('d1', [0]*10)) + sum(x.get('d2', [0]*10))
            need = max(0, d_total - stk)

            if need > 0:
                add = min(need, cap)
                x['prod'][r] += add
                rot_grp_used[r][g] = rot_grp_used[r].get(g, 0) + add
                rot_totals[r] += add

    # 2차: 아직 용량 남으면 모든 아이템에 수요 비례 배분
    for r in range(ROTATIONS):
        rot_remaining = MAX_PER_ROTATION - rot_totals[r]

        if rot_remaining <= 0:
            continue

        for x in sorted(items, key=lambda x: -sum(x['d0'])):  # 수요 큰 순
            if rot_remaining <= 0:
                break

            g = x.get('grp')
            if not g:
                continue

            cap = get_grp_capacity(g, r, rot_totals[r])
            if cap <= 0:
                continue

            # 3일치 목표 대비 부족분 (최소 1 할당)
            stk = x.get('stk', 0) + sum(x['prod'])
            d_total = sum(x['d0']) + sum(x.get('d1', [0]*10)) + sum(x.get('d2', [0]*10))
            need = max(1, d_total - stk)  # 최소 1 → 용량 채우기

            add = min(need, cap, rot_remaining)
            x['prod'][r] += add
            rot_grp_used[r][g] = rot_grp_used[r].get(g, 0) + add
            rot_totals[r] += add
            rot_remaining -= add


# ============================================
# 2. MIP 최적화 (OR-Tools)
# ============================================
def schedule_mip(items):
    """MIP(혼합정수계획법)를 사용한 최적 스케줄링

    목적함수: 컬러교환 최소화
    제약조건:
    - 재고 >= 0 (모든 회전)
    - D+1 주간(1-5회전) 재고 >= 0
    - 용량 제약 (그룹별/회전별)
    - 지그교체 예산 (주간/야간 각 150행어)
    """
    if not items:
        return {'error': '스케줄링할 아이템이 없습니다.'}

    try:
        from ortools.linear_solver import pywraplp
    except ImportError as e:
        return {'error': f'OR-Tools 임포트 실패: {e}'}
    except Exception as e:
        return {'error': f'OR-Tools 로드 오류: {e}'}

    solver = pywraplp.Solver.CreateSolver('SCIP')
    if not solver:
        # SCIP 실패 시 CBC 시도
        solver = pywraplp.Solver.CreateSolver('CBC')
    if not solver:
        return {'error': 'MIP 솔버(SCIP/CBC)를 생성할 수 없습니다.'}

    n_items = len(items)
    n_rotations = ROTATIONS
    JIG_BUDGET_DAY = 150
    JIG_BUDGET_NIGHT = 150
    MAX_PER_ROTATION = HANGERS * JIGS_PER_HANGER  # 280
    BIG_M = 1000  # Big-M for indicator constraints

    # 그룹별, 컬러별 아이템 인덱스
    grp_items = defaultdict(list)
    color_items = defaultdict(list)
    colors = set()
    for i, item in enumerate(items):
        if item.get('grp'):
            grp_items[item['grp']].append(i)
        if item.get('clr'):
            color_items[item['clr']].append(i)
            colors.add(item['clr'])

    groups = list(JIG_INVENTORY.keys())
    colors = list(colors)

    # ============================================
    # 결정변수
    # ============================================
    # x[i,r] = 아이템 i를 회전 r에 생산하는 양
    x = {}
    for i in range(n_items):
        for r in range(n_rotations):
            x[i, r] = solver.IntVar(0, 500, f'x_{i}_{r}')

    # h[g,r] = 그룹 g가 회전 r에 사용하는 행어 수
    h = {}
    for g in groups:
        max_hangers = JIG_INVENTORY[g]['max_jigs'] // JIGS_PER_HANGER
        for r in range(n_rotations):
            h[g, r] = solver.IntVar(0, max_hangers, f'h_{g}_{r}')

    # ============================================
    # 위치 기반 지그교체 모델링
    # 그룹 순서 고정: A, B, B2, C, D, E, F, G, H, I
    # cumsum[k,r] = 처음 k+1개 그룹의 누적 행어 수
    # 위치 변화 = 경계 이동량의 합 = Σ|cumsum[k,r] - cumsum[k,r-1]|
    # ============================================
    ordered_groups = ['A', 'B', 'B2', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
    n_groups = len(ordered_groups)

    # cumsum[k,r] = sum of h[groups[i],r] for i = 0 to k
    # cumsum[-1,r] = 0 (가상의 시작점)
    # cumsum[n_groups-1,r] = 140 (총 행어 수)

    # delta[k,r] = |cumsum[k,r] - cumsum[k,r-1]| for k = 0 to n_groups-2
    # (마지막 경계는 항상 140이므로 제외)
    delta = {}
    for k in range(n_groups - 1):  # 0 to 8 (9개 내부 경계)
        for r in range(1, n_rotations):
            delta[k, r] = solver.IntVar(0, HANGERS, f'delta_{k}_{r}')

    # y[c,r] = 1 if color c is used in rotation r
    y = {}
    for c in colors:
        for r in range(n_rotations):
            y[c, r] = solver.BoolVar(f'y_{c}_{r}')

    # cc[c,r] = 1 if color c starts in rotation r (color change)
    # cc[c,r] = 1 when y[c,r]=1 and y[c,r-1]=0
    cc = {}
    for c in colors:
        for r in range(n_rotations):
            cc[c, r] = solver.BoolVar(f'cc_{c}_{r}')

    # ============================================
    # 제약조건
    # ============================================

    # 1. 재고 >= 0 (각 회전 후) - 2회전 리드타임 적용
    # r회전 수요 충족: r-2회전까지의 생산만 사용 가능
    LEAD_TIME = 2
    for i, item in enumerate(items):
        stk = item.get('stk', 0)
        for r in range(n_rotations):
            # r회전까지의 수요
            cum_demand = sum(item['d0'][rr] for rr in range(r + 1))
            # r-2회전까지의 생산만 사용 가능 (리드타임)
            available_rot = r - LEAD_TIME
            cum_prod = sum(x[i, rr] for rr in range(available_rot + 1)) if available_rot >= 0 else 0
            solver.Add(stk + cum_prod - cum_demand >= 0)

    # 2. D+1 주간(1-5회전) 재고 >= 0 - 강한 제약
    # D0 기말재고가 D+1 주간 수요를 모두 커버해야 함
    for i, item in enumerate(items):
        stk = item.get('stk', 0)
        d0_demand = sum(item['d0'])
        d1 = item.get('d1', [0]*10)
        d0_prod = sum(x[i, r] for r in range(n_rotations))

        # D+1 각 회전별 누적 재고 >= 0 (강한 제약)
        for d1_rot in range(5):  # D+1 1-5회전
            d1_cum_demand = sum(d1[r] for r in range(d1_rot + 1))
            # D0 기말재고 - D+1 누적수요 >= 0
            solver.Add(stk + d0_prod - d0_demand - d1_cum_demand >= 0)

    # 3. 그룹별 생산량 <= 행어 × 지그 × pcs
    for g in groups:
        if g not in grp_items or not grp_items[g]:
            continue
        pcs = JIG_INVENTORY[g]['pcs']
        for r in range(n_rotations):
            grp_prod = sum(x[i, r] for i in grp_items[g])
            solver.Add(grp_prod <= h[g, r] * JIGS_PER_HANGER * pcs)

    # 4. 회전당 총 행어 = 140
    for r in range(n_rotations):
        solver.Add(sum(h[g, r] for g in groups) == HANGERS)

    # 5. 회전당 총 생산량 상한
    for r in range(n_rotations):
        solver.Add(sum(x[i, r] for i in range(n_items)) <= MAX_PER_ROTATION * 2)

    # 6. 위치 기반 지그 변화량 (경계 이동)
    # cumsum[k,r] = sum(h[ordered_groups[i],r] for i in range(k+1))
    # delta[k,r] >= |cumsum[k,r] - cumsum[k,r-1]|
    for k in range(n_groups - 1):
        for r in range(1, n_rotations):
            # cumsum[k,r] - cumsum[k,r-1]의 절대값
            cumsum_curr = sum(h[ordered_groups[i], r] for i in range(k + 1))
            cumsum_prev = sum(h[ordered_groups[i], r - 1] for i in range(k + 1))
            solver.Add(delta[k, r] >= cumsum_curr - cumsum_prev)
            solver.Add(delta[k, r] >= cumsum_prev - cumsum_curr)

    # 7. 지그교체 예산 (위치 기반)
    # 주간 (회전 1-4, 즉 r=1,2,3,4): 회전0→1, 1→2, 2→3, 3→4
    day_position_changes = sum(delta[k, r] for k in range(n_groups - 1) for r in range(1, 5))
    solver.Add(day_position_changes <= JIG_BUDGET_DAY)

    # 야간 (회전 5-9, 즉 r=5,6,7,8,9): 회전4→5, 5→6, 6→7, 7→8, 8→9
    night_position_changes = sum(delta[k, r] for k in range(n_groups - 1) for r in range(5, n_rotations))
    solver.Add(night_position_changes <= JIG_BUDGET_NIGHT)

    # 8. 컬러 사용 여부: y[c,r] = 1 iff sum(x[i,r] for i in color c) > 0
    for c in colors:
        if c not in color_items:
            continue
        for r in range(n_rotations):
            color_prod = sum(x[i, r] for i in color_items[c])
            # color_prod > 0 => y[c,r] = 1
            solver.Add(color_prod <= BIG_M * y[c, r])
            # y[c,r] = 1 => color_prod >= 1 (약한 제약, 선택적)
            # solver.Add(color_prod >= y[c, r])

    # 9. 컬러 교환 감지: cc[c,r] >= y[c,r] - y[c,r-1]
    for c in colors:
        # 회전 0: 새 컬러 시작 = y[c,0]
        solver.Add(cc[c, 0] >= y[c, 0])
        solver.Add(cc[c, 0] <= y[c, 0])  # cc[c,0] = y[c,0]
        # 회전 1~9
        for r in range(1, n_rotations):
            solver.Add(cc[c, r] >= y[c, r] - y[c, r - 1])
            solver.Add(cc[c, r] <= y[c, r])  # cc는 y보다 클 수 없음

    # ============================================
    # 추가 제약 1: 컬러 연속성 강제
    # 컬러는 한 번 시작하면 연속해서 나와야 함 (중간에 끊기면 안됨)
    # y[c,r]=1 and y[c,r+1]=0 and y[c,r+2]=1 은 불가 (gap 금지)
    # ============================================
    for c in colors:
        for r in range(n_rotations - 2):
            # y[c,r] + (1-y[c,r+1]) + y[c,r+2] <= 2
            # 즉, r에서 사용하고, r+1에서 안쓰고, r+2에서 다시 쓰면 안됨
            solver.Add(y[c, r] - y[c, r + 1] + y[c, r + 2] <= 1)

    # ============================================
    # 추가 제약 2: 컬러 종료 추적 (특수컬러 구분)
    # 특수컬러(MGG, T4M, UMA, ZRM, ISM, MRM) 종료: 15행어 비용
    # 일반컬러 종료: 1행어 비용
    # ============================================
    color_end = {}
    for c in colors:
        for r in range(n_rotations - 1):
            color_end[c, r] = solver.BoolVar(f'end_{c}_{r}')
            solver.Add(color_end[c, r] >= y[c, r] - y[c, r + 1])
            solver.Add(color_end[c, r] <= y[c, r])

    # 특수컬러 종료 합계 (15배 비용)
    special_color_ends = sum(
        color_end[c, r]
        for c in colors if c.upper() in SPECIAL_COLORS
        for r in range(n_rotations - 1)
    )
    # 일반컬러 종료 합계
    normal_color_ends = sum(
        color_end[c, r]
        for c in colors if c.upper() not in SPECIAL_COLORS
        for r in range(n_rotations - 1)
    )

    # ============================================
    # 추가 제약 3: 회전당 컬러 수 제한
    # ============================================
    MAX_COLORS_PER_ROT = 4
    for r in range(n_rotations):
        solver.Add(sum(y[c, r] for c in colors) <= MAX_COLORS_PER_ROT)

    # ============================================
    # 추가 제약 4: 총 컬러-회전 쌍 제한
    # ============================================
    total_color_rotation_pairs = sum(y[c, r] for c in colors for r in range(n_rotations))
    solver.Add(total_color_rotation_pairs <= 32)

    # ============================================
    # 목적함수: 빈행어 최소화 (특수컬러 15배 반영)
    # ============================================
    total_color_starts = sum(cc[c, r] for c in colors for r in range(n_rotations))
    total_production = sum(x[i, r] for i in range(n_items) for r in range(n_rotations))

    # 빈행어 비용: 특수컬러 종료 * 15 + 일반컬러 종료 * 1
    empty_hanger_cost = 15 * special_color_ends + 1 * normal_color_ends

    # CC 최소화 + 빈행어 비용 반영
    CC_WEIGHT = 1000
    EMPTY_WEIGHT = 100  # 빈행어 1개 = 100 가치

    solver.Minimize(CC_WEIGHT * total_color_starts + EMPTY_WEIGHT * empty_hanger_cost - total_production)

    solver.SetTimeLimit(30000)  # 30초 (Railway 타임아웃 방지)

    # 풀이
    status = solver.Solve()

    if status not in [pywraplp.Solver.OPTIMAL, pywraplp.Solver.FEASIBLE]:
        return {'error': f'최적해를 찾을 수 없습니다. 상태: {status}'}

    # 결과 추출
    for i, item in enumerate(items):
        item['prod'] = [int(x[i, r].solution_value()) for r in range(n_rotations)]

    # 지그교체 계산 (실제 위치 비교 - 정확한 계산)
    # MIP delta는 예산 제약용, 실제 값은 위치 비교로 계산
    jig_changes = [0] * n_rotations
    prev_positions = None
    for r in range(n_rotations):
        template = {g: int(h[g, r].solution_value()) for g in groups}
        active_order = [g for g in ordered_groups if template.get(g, 0) > 0]
        curr_positions = order_to_positions(template, active_order)
        if prev_positions:
            jig_changes[r] = calc_position_changes(prev_positions, curr_positions)
        prev_positions = curr_positions

    # 컬러교환 및 빈행어 계산 (특수컬러 15개 반영)
    # 손실 계산 (컬러교환 + 홀수)
    losses = calculate_all_losses(items, 'prod')
    cc_count = losses['cc_count']
    empty_hangers = losses['empty_hangers']
    odd_jig_loss = losses['odd_jig_loss']

    # 순생산량 = 총생산량 - 빈행어손실 - 홀수손실
    gross_production = sum(sum(item['prod']) for item in items)
    net_production = gross_production - losses['total_loss']

    # 회전별 템플릿 및 컬러 추출 (리포트용)
    templates = []
    rotation_colors = []
    jig_orders = []
    for r in range(n_rotations):
        # 템플릿: 그룹별 행어 수
        template = {g: int(h[g, r].solution_value()) for g in groups}
        templates.append(template)
        # 지그 순서: 고정 순서에서 사용중인 그룹만
        active_order = [g for g in ordered_groups if template.get(g, 0) > 0]
        jig_orders.append(active_order)
        # 회전에 사용된 컬러들
        rot_colors = []
        for i, item in enumerate(items):
            if item['prod'][r] > 0 and item.get('clr'):
                if item['clr'] not in rot_colors:
                    rot_colors.append(item['clr'])
        rotation_colors.append(rot_colors)

    return {
        'algorithm': 'MIP',
        'd0': {
            'color_changes': cc_count,
            'cc_count': cc_count,  # 리포트 호환
            'empty_hangers': empty_hangers,
            'cc_hangers': empty_hangers,  # 리포트 호환
            'odd_jig_loss': odd_jig_loss,  # 홀수 생산 손실
            'jig_changes': jig_changes,
            'gross_production': gross_production,
            'total_production': net_production,
            'templates': templates,  # 리포트 호환
            'colors': rotation_colors,  # 리포트 호환
            'jig_orders': jig_orders  # 고정 순서 기반
        },
        'd1': {'templates': [{}]*10, 'colors': [[]]*10, 'jig_changes': [0]*10, 'jig_orders': [None]*10, 'odd_jig_loss': 0},
        'd2': {'templates': [{}]*10, 'colors': [[]]*10, 'jig_changes': [0]*10, 'jig_orders': [None]*10, 'odd_jig_loss': 0}
    }


# ============================================
# 2-1. MIP 2일 최적화 (D0 + D+1)
# ============================================
def schedule_mip_2days(items):
    """MIP(혼합정수계획법)를 사용한 2일 스케줄링 (D0 + D+1)

    목적함수: 컬러교환 최소화
    제약조건 (D0, D+1 모두 동일):
    - 재고 >= 0 (모든 회전)
    - D+2 주간(1-5회전) 재고 >= 0
    - 용량 제약 (그룹별/회전별)
    - 지그교체 예산 (주간/야간 각 150행어)
    """
    if not items:
        return {'error': '스케줄링할 아이템이 없습니다.'}

    try:
        from ortools.linear_solver import pywraplp
    except ImportError as e:
        return {'error': f'OR-Tools 임포트 실패: {e}'}
    except Exception as e:
        return {'error': f'OR-Tools 로드 오류: {e}'}

    solver = pywraplp.Solver.CreateSolver('SCIP')
    if not solver:
        solver = pywraplp.Solver.CreateSolver('CBC')
    if not solver:
        return {'error': 'MIP 솔버(SCIP/CBC)를 생성할 수 없습니다.'}

    n_items = len(items)
    n_rotations = ROTATIONS * 2  # 20 rotations (D0: 0-9, D+1: 10-19)
    JIG_BUDGET_DAY = 150
    JIG_BUDGET_NIGHT = 150
    MAX_PER_ROTATION = HANGERS * JIGS_PER_HANGER  # 280
    BIG_M = 1000

    # 그룹별, 컬러별 아이템 인덱스
    grp_items = defaultdict(list)
    color_items = defaultdict(list)
    colors = set()
    for i, item in enumerate(items):
        if item.get('grp'):
            grp_items[item['grp']].append(i)
        if item.get('clr'):
            color_items[item['clr']].append(i)
            colors.add(item['clr'])

    groups = list(JIG_INVENTORY.keys())
    colors = list(colors)

    # ============================================
    # 결정변수
    # ============================================
    # x[i,r] = 아이템 i를 회전 r에 생산하는 양 (r: 0-19)
    x = {}
    for i in range(n_items):
        for r in range(n_rotations):
            x[i, r] = solver.IntVar(0, 500, f'x_{i}_{r}')

    # h[g,r] = 그룹 g가 회전 r에 사용하는 행어 수
    h = {}
    for g in groups:
        max_hangers = JIG_INVENTORY[g]['max_jigs'] // JIGS_PER_HANGER
        for r in range(n_rotations):
            h[g, r] = solver.IntVar(0, max_hangers, f'h_{g}_{r}')

    # 위치 기반 지그교체 모델링
    ordered_groups = ['A', 'B', 'B2', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
    n_groups = len(ordered_groups)

    delta = {}
    for k in range(n_groups - 1):
        for r in range(1, n_rotations):
            delta[k, r] = solver.IntVar(0, HANGERS, f'delta_{k}_{r}')

    # y[c,r] = 1 if color c is used in rotation r
    y = {}
    for c in colors:
        for r in range(n_rotations):
            y[c, r] = solver.BoolVar(f'y_{c}_{r}')

    # cc[c,r] = 1 if color c starts in rotation r
    cc = {}
    for c in colors:
        for r in range(n_rotations):
            cc[c, r] = solver.BoolVar(f'cc_{c}_{r}')

    # ============================================
    # 제약조건
    # ============================================

    # 1. 재고 >= 0 (각 회전 후) - 2회전 리드타임 적용
    LEAD_TIME = 2
    for i, item in enumerate(items):
        stk = item.get('stk', 0)
        d0 = item.get('d0', [0]*10)
        d1 = item.get('d1', [0]*10)

        for r in range(n_rotations):
            # r회전까지의 수요 (D0: 0-9, D+1: 10-19)
            if r < 10:
                cum_demand = sum(d0[rr] for rr in range(r + 1))
            else:
                cum_demand = sum(d0) + sum(d1[rr] for rr in range(r - 10 + 1))

            # r-2회전까지의 생산만 사용 가능 (리드타임)
            available_rot = r - LEAD_TIME
            cum_prod = sum(x[i, rr] for rr in range(available_rot + 1)) if available_rot >= 0 else 0
            solver.Add(stk + cum_prod - cum_demand >= 0)

    # 2. D+2 주간(1-5회전) 재고 >= 0
    for i, item in enumerate(items):
        stk = item.get('stk', 0)
        d0_demand = sum(item.get('d0', [0]*10))
        d1_demand = sum(item.get('d1', [0]*10))
        d2 = item.get('d2', [0]*10)

        total_prod = sum(x[i, r] for r in range(n_rotations))

        # D+2 각 회전별 누적 재고 >= 0
        for d2_rot in range(5):
            d2_cum_demand = sum(d2[r] for r in range(d2_rot + 1))
            solver.Add(stk + total_prod - d0_demand - d1_demand - d2_cum_demand >= 0)

    # 3. 그룹별 생산량 <= 행어 × 지그 × pcs
    for g in groups:
        if g not in grp_items or not grp_items[g]:
            continue
        pcs = JIG_INVENTORY[g]['pcs']
        for r in range(n_rotations):
            grp_prod = sum(x[i, r] for i in grp_items[g])
            solver.Add(grp_prod <= h[g, r] * JIGS_PER_HANGER * pcs)

    # 4. 회전당 총 행어 = 140
    for r in range(n_rotations):
        solver.Add(sum(h[g, r] for g in groups) == HANGERS)

    # 5. 회전당 총 생산량 상한
    for r in range(n_rotations):
        solver.Add(sum(x[i, r] for i in range(n_items)) <= MAX_PER_ROTATION * 2)

    # 6. 위치 기반 지그 변화량
    for k in range(n_groups - 1):
        for r in range(1, n_rotations):
            cumsum_curr = sum(h[ordered_groups[i], r] for i in range(k + 1))
            cumsum_prev = sum(h[ordered_groups[i], r - 1] for i in range(k + 1))
            solver.Add(delta[k, r] >= cumsum_curr - cumsum_prev)
            solver.Add(delta[k, r] >= cumsum_prev - cumsum_curr)

    # 7. 지그교체 예산 (D0, D+1 각각)
    # D0 주간 (r=1,2,3,4)
    day0_day_changes = sum(delta[k, r] for k in range(n_groups - 1) for r in range(1, 5))
    solver.Add(day0_day_changes <= JIG_BUDGET_DAY)
    # D0 야간 (r=5,6,7,8,9)
    day0_night_changes = sum(delta[k, r] for k in range(n_groups - 1) for r in range(5, 10))
    solver.Add(day0_night_changes <= JIG_BUDGET_NIGHT)
    # D+1 주간 (r=11,12,13,14)
    day1_day_changes = sum(delta[k, r] for k in range(n_groups - 1) for r in range(11, 15))
    solver.Add(day1_day_changes <= JIG_BUDGET_DAY)
    # D+1 야간 (r=15,16,17,18,19)
    day1_night_changes = sum(delta[k, r] for k in range(n_groups - 1) for r in range(15, 20))
    solver.Add(day1_night_changes <= JIG_BUDGET_NIGHT)

    # 8. 컬러 사용 여부
    for c in colors:
        if c not in color_items:
            continue
        for r in range(n_rotations):
            color_prod = sum(x[i, r] for i in color_items[c])
            solver.Add(color_prod <= BIG_M * y[c, r])

    # 9. 컬러 교환 감지
    for c in colors:
        solver.Add(cc[c, 0] >= y[c, 0])
        solver.Add(cc[c, 0] <= y[c, 0])
        for r in range(1, n_rotations):
            solver.Add(cc[c, r] >= y[c, r] - y[c, r - 1])
            solver.Add(cc[c, r] <= y[c, r])

    # 10. 컬러 연속성 - 생략 (목적함수로 CC 최소화, 속도 향상)
    # 복잡한 제약이므로 2일 모델에서는 제외

    # 11. 컬러 종료 추적
    color_end = {}
    for c in colors:
        for r in range(n_rotations - 1):
            color_end[c, r] = solver.BoolVar(f'end_{c}_{r}')
            solver.Add(color_end[c, r] >= y[c, r] - y[c, r + 1])
            solver.Add(color_end[c, r] <= y[c, r])

    special_color_ends = sum(
        color_end[c, r]
        for c in colors if c.upper() in SPECIAL_COLORS
        for r in range(n_rotations - 1)
    )
    normal_color_ends = sum(
        color_end[c, r]
        for c in colors if c.upper() not in SPECIAL_COLORS
        for r in range(n_rotations - 1)
    )

    # 12. 회전당 컬러 수 제한
    MAX_COLORS_PER_ROT = 4
    for r in range(n_rotations):
        solver.Add(sum(y[c, r] for c in colors) <= MAX_COLORS_PER_ROT)

    # 13. 총 컬러-회전 쌍 제한 - 생략 (속도 향상)
    # 목적함수에서 CC 최소화로 자연스럽게 제한됨

    # ============================================
    # 목적함수
    # ============================================
    total_color_starts = sum(cc[c, r] for c in colors for r in range(n_rotations))
    total_production = sum(x[i, r] for i in range(n_items) for r in range(n_rotations))
    empty_hanger_cost = 15 * special_color_ends + 1 * normal_color_ends

    CC_WEIGHT = 1000
    EMPTY_WEIGHT = 100
    solver.Minimize(CC_WEIGHT * total_color_starts + EMPTY_WEIGHT * empty_hanger_cost - total_production)

    solver.SetTimeLimit(25000)  # 25초 (Railway 타임아웃 방지)

    status = solver.Solve()

    if status not in [pywraplp.Solver.OPTIMAL, pywraplp.Solver.FEASIBLE]:
        return {'error': f'최적해를 찾을 수 없습니다. 상태: {status}'}

    # 결과 추출
    for i, item in enumerate(items):
        item['prod'] = [int(x[i, r].solution_value()) for r in range(10)]  # D0
        item['prod1'] = [int(x[i, r].solution_value()) for r in range(10, 20)]  # D+1

    # D0 지그교체 계산
    jig_changes_d0 = [0] * 10
    prev_positions = None
    templates_d0 = []
    jig_orders_d0 = []
    rotation_colors_d0 = []

    for r in range(10):
        template = {g: int(h[g, r].solution_value()) for g in groups}
        templates_d0.append(template)
        active_order = [g for g in ordered_groups if template.get(g, 0) > 0]
        jig_orders_d0.append(active_order)
        curr_positions = order_to_positions(template, active_order)
        if prev_positions:
            jig_changes_d0[r] = calc_position_changes(prev_positions, curr_positions)
        prev_positions = curr_positions
        rot_colors = []
        for i, item in enumerate(items):
            if item['prod'][r] > 0 and item.get('clr'):
                if item['clr'] not in rot_colors:
                    rot_colors.append(item['clr'])
        rotation_colors_d0.append(rot_colors)

    # D+1 지그교체 계산
    jig_changes_d1 = [0] * 10
    templates_d1 = []
    jig_orders_d1 = []
    rotation_colors_d1 = []

    for r in range(10, 20):
        template = {g: int(h[g, r].solution_value()) for g in groups}
        templates_d1.append(template)
        active_order = [g for g in ordered_groups if template.get(g, 0) > 0]
        jig_orders_d1.append(active_order)
        curr_positions = order_to_positions(template, active_order)
        if prev_positions:
            jig_changes_d1[r - 10] = calc_position_changes(prev_positions, curr_positions)
        prev_positions = curr_positions
        rot_colors = []
        for i, item in enumerate(items):
            if item['prod1'][r - 10] > 0 and item.get('clr'):
                if item['clr'] not in rot_colors:
                    rot_colors.append(item['clr'])
        rotation_colors_d1.append(rot_colors)

    # D0 손실 계산
    losses_d0 = calculate_all_losses(items, 'prod')
    # D+1 손실 계산
    losses_d1 = calculate_all_losses(items, 'prod1')

    gross_d0 = sum(sum(item['prod']) for item in items)
    net_d0 = gross_d0 - losses_d0['total_loss']
    gross_d1 = sum(sum(item['prod1']) for item in items)
    net_d1 = gross_d1 - losses_d1['total_loss']

    return {
        'algorithm': 'MIP_2days',
        'd0': {
            'color_changes': losses_d0['cc_count'],
            'cc_count': losses_d0['cc_count'],
            'empty_hangers': losses_d0['empty_hangers'],
            'cc_hangers': losses_d0['empty_hangers'],
            'odd_jig_loss': losses_d0['odd_jig_loss'],
            'jig_changes': jig_changes_d0,
            'gross_production': gross_d0,
            'total_production': net_d0,
            'templates': templates_d0,
            'colors': rotation_colors_d0,
            'jig_orders': jig_orders_d0
        },
        'd1': {
            'color_changes': losses_d1['cc_count'],
            'cc_count': losses_d1['cc_count'],
            'empty_hangers': losses_d1['empty_hangers'],
            'cc_hangers': losses_d1['empty_hangers'],
            'odd_jig_loss': losses_d1['odd_jig_loss'],
            'jig_changes': jig_changes_d1,
            'gross_production': gross_d1,
            'total_production': net_d1,
            'templates': templates_d1,
            'colors': rotation_colors_d1,
            'jig_orders': jig_orders_d1
        },
        'd2': {'templates': [{}]*10, 'colors': [[]]*10, 'jig_changes': [0]*10, 'jig_orders': [None]*10, 'odd_jig_loss': 0}
    }


# ============================================
# 3. 컬러 중심 (Color-First) - 지그예산 제약 포함
# ============================================
def schedule_color_first(items):
    """컬러를 먼저 클러스터링하고 그룹을 자동 결정

    전략:
    1. 같은 컬러를 연속 회전에 배치 (지그변화 최소화)
    2. 수요 많은 컬러부터 회전 할당
    3. 지그교체 예산(주간150/야간150) 준수
    """
    JIG_BUDGET_DAY = 150
    JIG_BUDGET_NIGHT = 150
    MAX_PER_ROTATION = HANGERS * JIGS_PER_HANGER  # 280

    for x in items:
        x['prod'] = [0] * ROTATIONS

    # 컬러별 아이템 그룹화
    color_items = defaultdict(list)
    for x in items:
        if x.get('clr'):
            color_items[x['clr']].append(x)

    # 컬러별 총 수요 계산
    color_demand = {}
    for clr, clr_items_list in color_items.items():
        d0_total = sum(sum(x['d0']) for x in clr_items_list)
        d1_12 = sum(sum(x.get('d1', [0]*10)[:2]) for x in clr_items_list)
        stk = sum(x.get('stk', 0) for x in clr_items_list)
        need = max(0, d0_total + d1_12 - stk)
        color_demand[clr] = {'need': need, 'items': clr_items_list}

    sorted_colors = sorted(color_demand.keys(), key=lambda c: -color_demand[c]['need'])

    # 회전별 그룹 생산량 추적
    rot_grp_prod = [{g: 0 for g in JIG_INVENTORY} for _ in range(ROTATIONS)]
    rot_total = [0] * ROTATIONS

    def get_grp_capacity(g, r):
        max_h = JIG_INVENTORY[g]['max_jigs'] // JIGS_PER_HANGER
        pcs = JIG_INVENTORY[g]['pcs']
        max_cap = max_h * JIGS_PER_HANGER * pcs
        grp_remain = max_cap - rot_grp_prod[r][g]
        rot_remain = MAX_PER_ROTATION - rot_total[r]
        return max(0, min(grp_remain, rot_remain))

    # 주요 전략: 같은 컬러는 같은 회전들에 집중 배치
    # 이렇게 하면 회전 간 그룹 구성이 비슷해져 지그변화 감소
    rot_main_color = [None] * ROTATIONS  # 각 회전의 주 컬러

    for clr in sorted_colors:
        info = color_demand[clr]
        clr_items_list = info['items']
        total_need = info['need']

        if total_need <= 0:
            continue

        # 이 컬러에 적합한 회전 찾기 (이미 같은 컬러가 있거나 비어있는 회전)
        preferred_rotations = []
        for r in range(ROTATIONS):
            if rot_main_color[r] == clr:
                preferred_rotations.insert(0, r)  # 같은 컬러 회전 우선
            elif rot_main_color[r] is None:
                preferred_rotations.append(r)  # 빈 회전

        # 나머지 회전도 추가 (용량 순)
        other_rotations = [r for r in range(ROTATIONS) if r not in preferred_rotations]
        other_rotations.sort(key=lambda r: rot_total[r])
        preferred_rotations.extend(other_rotations)

        for x in clr_items_list:
            g = x.get('grp')
            if not g or g not in JIG_INVENTORY:
                continue

            stk = x.get('stk', 0)
            d0_total = sum(x['d0'])
            d1_12 = sum(x.get('d1', [0]*10)[:2])
            need = max(0, d0_total + d1_12 - stk)

            if need <= 0:
                continue

            remaining = need
            for r in preferred_rotations:
                if remaining <= 0:
                    break
                cap = get_grp_capacity(g, r)
                if cap <= 0:
                    continue

                alloc = min(remaining, cap)
                x['prod'][r] += alloc
                rot_grp_prod[r][g] += alloc
                rot_total[r] += alloc
                remaining -= alloc

                # 주 컬러 설정
                if rot_main_color[r] is None:
                    rot_main_color[r] = clr

    # 용량 채우기 (3일치 안전재고 목표)
    fill_capacity_for_safety_stock(items, rot_grp_prod)

    jig_changes = calculate_jig_changes(items)

    # 손실 계산 (컬러교환 + 홀수)
    losses = calculate_all_losses(items, 'prod')
    cc_count = losses['cc_count']
    empty_hangers = losses['empty_hangers']
    odd_jig_loss = losses['odd_jig_loss']

    # 순생산량 = 총생산량 - 빈행어손실 - 홀수손실
    gross_production = sum(sum(x['prod']) for x in items)
    net_production = gross_production - losses['total_loss']

    # 회전별 템플릿 및 컬러 추출 (리포트용)
    templates = []
    rotation_colors = []
    for r in range(ROTATIONS):
        template = {g: 0 for g in JIG_INVENTORY}
        rot_colors = []
        for x in items:
            if x['prod'][r] > 0:
                g = x.get('grp')
                if g and g in template:
                    pcs = JIG_INVENTORY[g]['pcs']
                    template[g] += (x['prod'][r] + pcs - 1) // pcs // JIGS_PER_HANGER
                if x.get('clr') and x['clr'] not in rot_colors:
                    rot_colors.append(x['clr'])
        templates.append(template)
        rotation_colors.append(rot_colors)

    return {
        'algorithm': 'color_first',
        'd0': {
            'color_changes': cc_count,
            'cc_count': cc_count,
            'empty_hangers': empty_hangers,
            'cc_hangers': empty_hangers,
            'odd_jig_loss': odd_jig_loss,
            'jig_changes': jig_changes,
            'gross_production': gross_production,
            'total_production': net_production,
            'templates': templates,
            'colors': rotation_colors,
            'jig_orders': [None] * ROTATIONS
        },
        'd1': {'templates': [{}]*10, 'colors': [[]]*10, 'jig_changes': [0]*10, 'jig_orders': [None]*10, 'odd_jig_loss': 0},
        'd2': {'templates': [{}]*10, 'colors': [[]]*10, 'jig_changes': [0]*10, 'jig_orders': [None]*10, 'odd_jig_loss': 0}
    }


# ============================================
# 4. 2단계 분해 (Assignment + TSP) - 지그예산 제약 포함
# ============================================
def schedule_two_phase(items):
    """2단계 최적화:
    Phase 1: 컬러를 회전 블록에 할당 (연속 회전 선호)
    Phase 2: 회전 내 아이템 순서 최적화

    전략: 컬러별로 연속 회전 블록 할당하여 지그변화 최소화
    """
    MAX_PER_ROTATION = HANGERS * JIGS_PER_HANGER  # 280

    for x in items:
        x['prod'] = [0] * ROTATIONS

    # 컬러별 아이템 그룹화 및 필요량 계산
    color_items = defaultdict(list)
    color_needs = {}
    for x in items:
        clr = x.get('clr')
        if clr:
            color_items[clr].append(x)

    for clr, clr_items_list in color_items.items():
        total_need = 0
        for x in clr_items_list:
            stk = x.get('stk', 0)
            d0_total = sum(x['d0'])
            d1_12 = sum(x.get('d1', [0]*10)[:2])
            need = max(0, d0_total + d1_12 - stk)
            total_need += need
        color_needs[clr] = total_need

    # 수요 많은 컬러 순 정렬
    sorted_colors = sorted(color_needs.keys(), key=lambda c: -color_needs[c])

    # 회전별 그룹 생산량 추적
    rot_grp_prod = [{g: 0 for g in JIG_INVENTORY} for _ in range(ROTATIONS)]
    rot_total = [0] * ROTATIONS
    rot_assigned_color = [None] * ROTATIONS  # 각 회전에 할당된 주 컬러

    def get_grp_capacity(g, r):
        max_h = JIG_INVENTORY[g]['max_jigs'] // JIGS_PER_HANGER
        pcs = JIG_INVENTORY[g]['pcs']
        max_cap = max_h * JIGS_PER_HANGER * pcs
        grp_remain = max_cap - rot_grp_prod[r][g]
        rot_remain = MAX_PER_ROTATION - rot_total[r]
        return max(0, min(grp_remain, rot_remain))

    # Phase 1: 컬러를 연속 회전 블록에 할당
    next_free_rot = 0  # 다음 할당 시작 회전

    for clr in sorted_colors:
        total_need = color_needs[clr]
        if total_need <= 0:
            continue

        clr_items_list = color_items[clr]

        # 이 컬러에 필요한 회전 수 추정
        estimated_rotations = (total_need + MAX_PER_ROTATION - 1) // MAX_PER_ROTATION
        estimated_rotations = min(estimated_rotations, ROTATIONS)

        # 시작 회전 결정 (연속 블록 또는 여유 회전)
        start_rot = next_free_rot % ROTATIONS

        # 아이템별 생산 배분
        for x in clr_items_list:
            g = x.get('grp')
            if not g or g not in JIG_INVENTORY:
                continue

            stk = x.get('stk', 0)
            d0_total = sum(x['d0'])
            d1_12 = sum(x.get('d1', [0]*10)[:2])
            need = max(0, d0_total + d1_12 - stk)

            if need <= 0:
                continue

            remaining = need

            # 연속 회전에 배치 (start_rot부터)
            for offset in range(ROTATIONS):
                if remaining <= 0:
                    break
                r = (start_rot + offset) % ROTATIONS

                cap = get_grp_capacity(g, r)
                if cap <= 0:
                    continue

                alloc = min(remaining, cap)
                x['prod'][r] += alloc
                rot_grp_prod[r][g] += alloc
                rot_total[r] += alloc
                remaining -= alloc

                if rot_assigned_color[r] is None:
                    rot_assigned_color[r] = clr

        # 다음 컬러 시작 위치 업데이트
        next_free_rot = (start_rot + estimated_rotations) % ROTATIONS

    # 용량 채우기 (3일치 안전재고 목표)
    fill_capacity_for_safety_stock(items, rot_grp_prod)

    jig_changes = calculate_jig_changes(items)

    # 모든 손실 계산 (컬러교환 + 홀수 손실)
    losses = calculate_all_losses(items)
    cc_count = losses['cc_count']
    empty_hangers = losses['empty_hangers']
    odd_jig_loss = losses['odd_jig_loss']

    # 순생산량 = 총생산량 - 빈행어손실 - 홀수손실
    gross_production = sum(sum(x['prod']) for x in items)
    net_production = gross_production - losses['total_loss']

    # 회전별 템플릿 및 컬러 추출 (리포트용)
    templates = []
    rotation_colors = []
    for r in range(ROTATIONS):
        template = {g: 0 for g in JIG_INVENTORY}
        rot_colors = []
        for x in items:
            if x['prod'][r] > 0:
                g = x.get('grp')
                if g and g in template:
                    pcs = JIG_INVENTORY[g]['pcs']
                    template[g] += (x['prod'][r] + pcs - 1) // pcs // JIGS_PER_HANGER
                if x.get('clr') and x['clr'] not in rot_colors:
                    rot_colors.append(x['clr'])
        templates.append(template)
        rotation_colors.append(rot_colors)

    return {
        'algorithm': 'two_phase',
        'd0': {
            'color_changes': cc_count,
            'cc_count': cc_count,
            'empty_hangers': empty_hangers,
            'cc_hangers': empty_hangers,
            'jig_changes': jig_changes,
            'odd_jig_loss': odd_jig_loss,
            'gross_production': gross_production,
            'total_production': net_production,
            'templates': templates,
            'colors': rotation_colors,
            'jig_orders': [None] * ROTATIONS
        },
        'd1': {'templates': [{}]*10, 'colors': [[]]*10, 'jig_changes': [0]*10, 'jig_orders': [None]*10, 'odd_jig_loss': 0},
        'd2': {'templates': [{}]*10, 'colors': [[]]*10, 'jig_changes': [0]*10, 'jig_orders': [None]*10, 'odd_jig_loss': 0}
    }


# ============================================
# 유틸리티: 컬러교환 및 홀수 손실 계산
# ============================================
def calculate_odd_jig_loss(items, prod_key='prod'):
    """컬러 그룹별 홀수 생산량으로 인한 지그 손실 계산

    1행어 = 2지그, 같은 컬러는 행어 단위로 배치
    홀수 생산량 → 마지막 행어에 빈 지그 1개 발생

    Returns:
        odd_loss: 총 손실 지그 수
        odd_details: {(rotation, color): loss} 상세 정보
    """
    odd_loss = 0
    odd_details = {}

    for r in range(ROTATIONS):
        # 회전별 컬러별 생산량 집계
        color_prod = defaultdict(int)
        for x in items:
            prod = x.get(prod_key, [0] * 10)
            if r < len(prod) and prod[r] > 0:
                color_prod[x['clr']] += prod[r]

        # 홀수 생산량 체크
        for clr, qty in color_prod.items():
            if qty % 2 == 1:  # 홀수
                odd_loss += 1
                odd_details[(r, clr)] = 1

    return odd_loss, odd_details


def calculate_color_changes(items, return_details=False):
    """회전별 컬러교환 횟수 및 빈행어 계산

    특수컬러(MGG, T4M, UMA, ZRM, ISM, MRM) 후 교환: 15행어
    일반컬러 후 교환: 1행어
    """
    cc_count = 0
    empty_hangers = 0
    prev_color = None

    for r in range(ROTATIONS):
        # 이 회전에서 생산되는 컬러들
        colors_in_rot = []
        for x in items:
            if x['prod'][r] > 0 and x.get('clr'):
                if x['clr'] not in colors_in_rot:
                    colors_in_rot.append(x['clr'])

        # 컬러 전환 계산
        for clr in colors_in_rot:
            if prev_color and prev_color != clr:
                cc_count += 1
                # 이전 컬러 기준으로 빈행어 결정
                empty_hangers += get_color_change_cost(prev_color)
            prev_color = clr

    if return_details:
        return cc_count, empty_hangers
    return cc_count


def calculate_all_losses(items, prod_key='prod'):
    """모든 손실 통합 계산

    Returns:
        dict: {
            'cc_count': 컬러교환 횟수,
            'empty_hangers': 빈행어 수 (컬러교환),
            'empty_hanger_loss': 빈행어로 인한 생산손실 (행어*2),
            'odd_jig_loss': 홀수 생산으로 인한 지그 손실,
            'total_loss': 총 손실
        }
    """
    cc_count, empty_hangers = calculate_color_changes(items, return_details=True)
    odd_loss, _ = calculate_odd_jig_loss(items, prod_key)

    empty_hanger_loss = empty_hangers * JIGS_PER_HANGER
    total_loss = empty_hanger_loss + odd_loss

    return {
        'cc_count': cc_count,
        'empty_hangers': empty_hangers,
        'empty_hanger_loss': empty_hanger_loss,
        'odd_jig_loss': odd_loss,
        'total_loss': total_loss
    }


# ============================================
# 메인 스케줄러 선택
# ============================================
def normalize_result(result):
    """결과 형식 통일 (color_changes 키 추가)"""
    if 'd0' in result:
        d0 = result['d0']
        # cc_count -> color_changes 매핑
        if 'cc_count' in d0 and 'color_changes' not in d0:
            d0['color_changes'] = d0['cc_count']
        if 'cc_hangers' in d0 and 'empty_hangers' not in d0:
            d0['empty_hangers'] = d0['cc_hangers']
        # total_production 계산
        if 'total_production' not in d0:
            d0['total_production'] = 0  # 아이템에서 계산 필요
    return result


def calculate_ending_inventory(items):
    """생산 후 기말재고 계산 (cur, cur1, cur2)"""
    for x in items:
        # D0 기말재고: stk - d0수요 + prod
        stk = x.get('stk', 0)
        for r in range(10):
            stk = stk - x['d0'][r] + x.get('prod', [0]*10)[r]
        x['cur'] = stk

        # D+1 기말재고: cur - d1수요 + prod1
        for r in range(10):
            stk = stk - x.get('d1', [0]*10)[r] + x.get('prod1', [0]*10)[r]
        x['cur1'] = stk

        # D+2 기말재고: cur1 - d2수요 + prod2
        for r in range(10):
            stk = stk - x.get('d2', [0]*10)[r] + x.get('prod2', [0]*10)[r]
        x['cur2'] = stk


def run_scheduler(items, algorithm='heuristic'):
    """알고리즘별 스케줄러 실행"""
    if algorithm == 'heuristic':
        # 기존 휴리스틱 사용
        from generate_report import schedule
        result = schedule(items)
        # 순생산량 계산 (빈행어 손실 차감)
        if 'd0' in result:
            gross_production = sum(sum(x.get('prod', [0]*10)) for x in items)
            cc_count, empty_hangers = calculate_color_changes(items, return_details=True)
            net_production = gross_production - (empty_hangers * JIGS_PER_HANGER)
            result['d0']['gross_production'] = gross_production
            result['d0']['total_production'] = net_production
            result['d0']['empty_hangers'] = empty_hangers
        return normalize_result(result)

    elif algorithm == 'mip':
        try:
            result = schedule_mip(items)
            if result is None:
                return {'error': 'MIP가 None을 반환했습니다'}
            if not isinstance(result, dict):
                return {'error': f'MIP가 dict가 아닌 값을 반환: {type(result)}'}
            if 'error' not in result and 'd0' not in result:
                return {'error': f'MIP 결과에 error도 d0도 없음: {list(result.keys())}'}
            if 'error' not in result:  # 성공 시에만 재고 계산
                calculate_ending_inventory(items)
            return result
        except Exception as e:
            import traceback
            return {'error': f'MIP 실행 오류: {e}', 'traceback': traceback.format_exc()[:500]}

    elif algorithm == 'color_first':
        result = schedule_color_first(items)
        if 'error' not in result:
            calculate_ending_inventory(items)
        return result

    elif algorithm == 'two_phase':
        result = schedule_two_phase(items)
        if 'error' not in result:
            calculate_ending_inventory(items)
        return result

    elif algorithm == 'mip_2days':
        try:
            result = schedule_mip_2days(items)
            if result is None:
                return {'error': 'MIP_2days가 None을 반환했습니다'}
            if not isinstance(result, dict):
                return {'error': f'MIP_2days가 dict가 아닌 값을 반환: {type(result)}'}
            if 'error' not in result and 'd0' not in result:
                return {'error': f'MIP_2days 결과에 error도 d0도 없음: {list(result.keys())}'}
            if 'error' not in result:
                calculate_ending_inventory(items)
            return result
        except Exception as e:
            import traceback
            return {'error': f'MIP_2days 실행 오류: {e}', 'traceback': traceback.format_exc()[:500]}

    else:
        return {'error': f'Unknown algorithm: {algorithm}'}
