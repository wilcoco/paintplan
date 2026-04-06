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

    # d[g,r] = |h[g,r] - h[g,r-1]| (지그 변화량)
    d = {}
    for g in groups:
        for r in range(1, n_rotations):
            max_hangers = JIG_INVENTORY[g]['max_jigs'] // JIGS_PER_HANGER
            d[g, r] = solver.IntVar(0, max_hangers, f'd_{g}_{r}')

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

    # 6. 지그 변화량 절대값
    for g in groups:
        for r in range(1, n_rotations):
            solver.Add(d[g, r] >= h[g, r] - h[g, r - 1])
            solver.Add(d[g, r] >= h[g, r - 1] - h[g, r])

    # 7. 지그교체 예산
    solver.Add(sum(d[g, r] for g in groups for r in range(1, 5)) <= JIG_BUDGET_DAY)
    solver.Add(sum(d[g, r] for g in groups for r in range(5, n_rotations)) <= JIG_BUDGET_NIGHT)

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

    solver.SetTimeLimit(60000)  # 60초 (Railway 타임아웃 방지)

    # 풀이
    status = solver.Solve()

    if status not in [pywraplp.Solver.OPTIMAL, pywraplp.Solver.FEASIBLE]:
        return {'error': f'최적해를 찾을 수 없습니다. 상태: {status}'}

    # 결과 추출
    for i, item in enumerate(items):
        item['prod'] = [int(x[i, r].solution_value()) for r in range(n_rotations)]

    # 지그교체 계산 (실제 결과에서)
    jig_changes = [0] * n_rotations
    for r in range(1, n_rotations):
        jig_changes[r] = sum(int(d[g, r].solution_value()) for g in groups)

    # 컬러교환 및 빈행어 계산 (특수컬러 15개 반영)
    cc_count, empty_hangers = calculate_color_changes(items, return_details=True)

    # 순생산량 = 총생산량 - 빈행어손실 (빈행어 1개 = 2지그 = 2개 손실)
    gross_production = sum(sum(item['prod']) for item in items)
    net_production = gross_production - (empty_hangers * JIGS_PER_HANGER)

    return {
        'algorithm': 'MIP',
        'd0': {
            'color_changes': cc_count,
            'empty_hangers': empty_hangers,
            'jig_changes': jig_changes,
            'gross_production': gross_production,
            'total_production': net_production  # 빈행어 손실 차감된 순생산량
        }
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
    cc_count, empty_hangers = calculate_color_changes(items, return_details=True)

    # 순생산량 = 총생산량 - 빈행어손실
    gross_production = sum(sum(x['prod']) for x in items)
    net_production = gross_production - (empty_hangers * JIGS_PER_HANGER)

    return {
        'algorithm': 'color_first',
        'd0': {
            'color_changes': cc_count,
            'empty_hangers': empty_hangers,
            'jig_changes': jig_changes,
            'gross_production': gross_production,
            'total_production': net_production
        }
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
    cc_count, empty_hangers = calculate_color_changes(items, return_details=True)

    # 순생산량 = 총생산량 - 빈행어손실
    gross_production = sum(sum(x['prod']) for x in items)
    net_production = gross_production - (empty_hangers * JIGS_PER_HANGER)

    return {
        'algorithm': 'two_phase',
        'd0': {
            'color_changes': cc_count,
            'empty_hangers': empty_hangers,
            'jig_changes': jig_changes,
            'gross_production': gross_production,
            'total_production': net_production
        }
    }


# ============================================
# 유틸리티: 컬러교환 계산
# ============================================
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
            return schedule_mip(items)
        except Exception as e:
            import traceback
            return {'error': f'MIP 실행 오류: {e}', 'traceback': traceback.format_exc()[:500]}

    elif algorithm == 'color_first':
        return schedule_color_first(items)

    elif algorithm == 'two_phase':
        return schedule_two_phase(items)

    else:
        return {'error': f'Unknown algorithm: {algorithm}'}
