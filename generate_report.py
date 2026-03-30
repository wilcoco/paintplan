#!/usr/bin/env python3
"""
D0 생산계획 웹 리포트 생성
- 시스템 변수 표시
- CSV 형식 테이블
- 부족분 빨간색 표시
"""
import openpyxl
from collections import defaultdict
from datetime import datetime, timedelta

# ============================================
# 시스템 변수
# ============================================
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

HANGERS = 140              # 총 행어 수
JIGS_PER_HANGER = 2        # 행어당 지그 수
ROTATIONS_PER_DAY = 10     # 일일 회전 수
DAY_SHIFT_ROTATIONS = 5    # 주간 회전 수 (1-5)
NIGHT_SHIFT_ROTATIONS = 5  # 야간 회전 수 (6-10)
JIG_BUDGET_DAY = 150       # 주간 지그 교체 예산
JIG_BUDGET_NIGHT = 150     # 야간 지그 교체 예산
COLOR_CHANGE_LOSS = 6      # 컬러 교환 시 손실 (3빈행어 × 2)
SAFETY_STOCK_DAYS = 3      # 안전재고 일수

def get_grp(ct, it, det=''):
    """아이템을 지그그룹에 배정
    ct: 차종, it: 아이템, det: 세부아이템
    """
    ct = ct.upper().replace(' ','').replace('\n','')
    it = it.upper().replace(' ','').replace('\n','')
    det = det.upper().replace(' ','').replace('\n','') if det else ''

    if 'TH' in ct:
        if 'STD' in it or 'LDT' in it: return 'A'
        if 'RR' in it: return 'H'
    if 'OV' in ct: return 'C'
    if 'NQ5' in ct:
        if 'FRT' in it:
            # NQ5 FRT: STD는 B2 전용, XLINE은 B
            if 'STD' in it or 'STD' in det:
                return 'B2'  # STD 전용 지그
            else:
                return 'B'   # XLINE 등 기타
        return 'I'  # NQ5 RR
    if 'SP3' in ct: return 'A'
    if 'JX' in ct:
        if 'CROSS' in it: return 'E'
        if 'RR' in it: return 'F'
        return 'D'
    if 'AX' in ct or 'PE' in ct: return 'G'
    return None

def load_data():
    wb = openpyxl.load_workbook('paint2.xlsx', data_only=True)
    ws = wb.active
    items, ct, it = [], None, None
    for r in range(8, 138):  # 137행까지 읽음
        a,b,c,d,e,f = [ws.cell(r,i).value for i in range(1,7)]
        if any(v and ('합계' in str(v) or '소계' in str(v)) for v in [b,d,e]): continue
        if a: ct = str(a).replace('\n',' ')
        if b: it = str(b).replace('\n',' ')
        if not ct or not it: continue
        clr = str(f).replace('\n','') if f else ''
        if not clr or clr=='None': continue
        det = str(c).replace('\n',' ').strip() if c else '-'
        if not det or det=='None': det = '-'
        g,h = ws.cell(r,7).value, ws.cell(r,8).value
        stk = (int(g) if isinstance(g,(int,float)) else 0) + (int(h) if isinstance(h,(int,float)) else 0)
        # D0: 10회전별 수요
        d0 = [int(ws.cell(r,c).value) if isinstance(ws.cell(r,c).value,(int,float)) else 0 for c in [22,24,26,28,30,32,34,36,38,40]]
        # D+1: 10회전별 수요
        d1 = [int(ws.cell(r,c).value) if isinstance(ws.cell(r,c).value,(int,float)) else 0 for c in [43,45,47,49,51,53,55,57,59,61]]
        # D+2: 10회전별 수요
        d2 = [int(ws.cell(r,c).value) if isinstance(ws.cell(r,c).value,(int,float)) else 0 for c in [67,69,71,73,75,77,79,81,83,85]]
        items.append({
            'ct':ct,'it':it,'det':det,'clr':clr,'stk':stk,
            'd0':d0,'d0t':sum(d0),
            'd1':d1,'d1t':sum(d1),
            'd2':d2,'d2t':sum(d2),
            'grp':get_grp(ct,it,det),'cur':stk,
            'prod':[0]*10,      # D0 생산
            'prod1':[0]*10,     # D+1 생산 (예정)
            'prod2':[0]*10      # D+2 생산 (예정)
        })
    wb.close()
    return items

def calc_jig_change(tmpl1, tmpl2, prev_order=None):
    """
    위치 기반 지그 교체 수 계산
    - 수요 많은 지그그룹이 앞쪽 배치
    - 이전 회전 순서 유지하여 교체 최소화
    """
    if not tmpl1 or not tmpl2:
        changes = sum(tmpl2.values()) if tmpl2 else sum(tmpl1.values()) if tmpl1 else 0
        return changes, None

    def template_to_positions(tmpl, order=None):
        """템플릿을 위치 배열로 변환 (수요순 또는 지정순)"""
        if order:
            # 이전 순서 유지 + 새 그룹은 뒤에
            sorted_grps = [g for g in order if g in tmpl and tmpl[g] > 0]
            new_grps = [g for g in tmpl if g not in sorted_grps and tmpl[g] > 0]
            sorted_grps.extend(sorted(new_grps, key=lambda g: -tmpl[g]))
        else:
            # 수량 많은 순 (수요 = 행어수)
            sorted_grps = sorted([g for g in tmpl if tmpl[g] > 0], key=lambda g: -tmpl[g])

        positions = []
        for g in sorted_grps:
            positions.extend([g] * tmpl[g])
        return positions, sorted_grps

    pos1, order1 = template_to_positions(tmpl1, prev_order)
    pos2, order2 = template_to_positions(tmpl2, order1)  # 이전 순서 유지

    # 길이 맞추기 (140행어)
    while len(pos1) < HANGERS:
        pos1.append(None)
    while len(pos2) < HANGERS:
        pos2.append(None)

    # 위치별 비교
    changes = sum(1 for i in range(HANGERS) if pos1[i] != pos2[i])
    return changes, order2


def calc_jig_change_simple(tmpl1, tmpl2):
    """간단한 지그 교체 계산 (순서 정보 없이)"""
    result, _ = calc_jig_change(tmpl1, tmpl2, None)
    return result

def optimal_template_for_color(items, clr, prev_tmpl, budget_left, day='d0', prev_order=None):
    """day: 'd0', 'd1', 'd2' 중 하나
    Returns: (template, change, new_order)
    """
    demand_key = day + 't'  # 'd0t', 'd1t', 'd2t'
    clr_items = [x for x in items if x['clr'] == clr and x['grp']]
    grp_demand = defaultdict(int)
    for x in clr_items:
        grp_demand[x['grp']] += x.get(demand_key, x['d0t'])
    all_grp_demand = defaultdict(int)
    for x in items:
        if x['grp']:
            all_grp_demand[x['grp']] += x.get(demand_key, x['d0t'])

    if grp_demand:
        tot = sum(grp_demand.values())
        ideal = {}
        for g in grp_demand:
            mx = JIG_INVENTORY[g]['max_jigs'] // JIGS_PER_HANGER
            ideal[g] = min(mx, max(1, int(HANGERS * grp_demand[g] / tot))) if tot else 1
    else:
        ideal = prev_tmpl.copy() if prev_tmpl else {}

    s = sum(ideal.values())
    if s < HANGERS:
        remaining_grps = sorted(all_grp_demand, key=lambda g: -all_grp_demand[g])
        for g in remaining_grps:
            if s >= HANGERS: break
            mx = JIG_INVENTORY[g]['max_jigs'] // JIGS_PER_HANGER
            current = ideal.get(g, 0)
            add = min(HANGERS - s, mx - current)
            if add > 0:
                ideal[g] = current + add
                s += add

    if s < HANGERS:
        for g in JIG_INVENTORY:
            if s >= HANGERS: break
            mx = JIG_INVENTORY[g]['max_jigs'] // JIGS_PER_HANGER
            current = ideal.get(g, 0)
            add = min(HANGERS - s, mx - current)
            if add > 0:
                ideal[g] = current + add
                s += add

    while s > HANGERS:
        for g in sorted(ideal, key=lambda x: -ideal[x]):
            if s <= HANGERS: break
            if ideal[g] > 1:
                rm = min(s - HANGERS, ideal[g] - 1)
                ideal[g] -= rm
                s -= rm

    ideal = {g: h for g, h in ideal.items() if h > 0}

    if not prev_tmpl:
        # 첫 템플릿: 수요 많은 순으로 정렬
        new_order = sorted([g for g in ideal if ideal[g] > 0], key=lambda g: -ideal[g])
        return ideal, 0, new_order

    change, new_order = calc_jig_change(prev_tmpl, ideal, prev_order)
    if change <= budget_left:
        return ideal, change, new_order

    new_tmpl = prev_tmpl.copy()
    for g in grp_demand:
        if g not in new_tmpl or new_tmpl[g] == 0:
            biggest = max(new_tmpl, key=lambda x: new_tmpl[x])
            if new_tmpl[biggest] > 1:
                mx = JIG_INVENTORY[g]['max_jigs'] // JIGS_PER_HANGER
                alloc = min(10, mx, new_tmpl[biggest]-1)
                new_tmpl[biggest] -= alloc
                new_tmpl[g] = new_tmpl.get(g, 0) + alloc

    change, new_order = calc_jig_change(prev_tmpl, new_tmpl, prev_order)
    return new_tmpl, min(change, budget_left), new_order

def schedule_day(items, day_key, demand_key, prod_key, start_stock_key,
                 prev_day_template=None, prev_day_color=None, prev_day_order=None):
    """하루치 스케줄링 (전날 마지막 상태 연속 고려)

    Args:
        prev_day_template: 전날 10회전 마지막 지그 템플릿 (D+1, D+2용)
        prev_day_color: 전날 10회전 마지막 컬러 (D+1, D+2용)
        prev_day_order: 전날 10회전 마지막 지그 순서 (D+1, D+2용)
    """
    clr_demand = defaultdict(int)
    for x in items:
        if x['grp']:
            clr_demand[x['clr']] += x.get(demand_key, 0)

    sorted_colors = sorted(clr_demand, key=lambda c: -clr_demand[c])

    # 전날 마지막 컬러가 있으면 그 컬러를 우선 배치 (컬러 교환 최소화)
    if prev_day_color and prev_day_color in clr_demand:
        # 전날 컬러를 맨 앞으로
        sorted_colors = [prev_day_color] + [c for c in sorted_colors if c != prev_day_color]

    avg_cap = 400
    color_rotations = []
    for clr in sorted_colors:
        need = clr_demand[clr]
        rots = max(1, int(need / avg_cap + 0.5))
        color_rotations.append((clr, rots, need))

    rotation_color = [None] * 10
    rot = 0
    for clr, rots, need in color_rotations:
        if rot >= 10: break
        actual_rots = min(rots, 10 - rot)
        for r in range(rot, rot + actual_rots):
            rotation_color[r] = clr
        rot += actual_rots

    if rot < 10:
        top_clr = sorted_colors[0] if sorted_colors else None
        for r in range(rot, 10):
            rotation_color[r] = top_clr

    templates = [None] * 10
    jig_changes = [0] * 10
    jig_orders = [None] * 10  # 회전별 지그 순서 추적

    # 주간 (1-5회전): 전날 템플릿에서 시작
    day_budget = JIG_BUDGET_DAY
    prev = prev_day_template  # 전날 마지막 템플릿에서 시작!
    prev_order = prev_day_order  # 전날 마지막 순서에서 시작!
    for r in range(5):
        clr = rotation_color[r]
        tmpl, change, new_order = optimal_template_for_color(items, clr, prev, day_budget, day_key, prev_order)
        templates[r] = tmpl
        jig_changes[r] = change
        jig_orders[r] = new_order
        day_budget -= change
        prev = tmpl
        prev_order = new_order

    # 야간 (6-10회전)
    night_budget = JIG_BUDGET_NIGHT
    prev = templates[4]
    prev_order = jig_orders[4]  # 5회전 마지막 순서에서 시작
    for r in range(5, 10):
        clr = rotation_color[r]
        tmpl, change, new_order = optimal_template_for_color(items, clr, prev, night_budget, day_key, prev_order)
        templates[r] = tmpl
        jig_changes[r] = change
        jig_orders[r] = new_order
        night_budget -= change
        prev = tmpl
        prev_order = new_order

    # 생산량 배정 - 컬러 블록 최적화 (주간 15회, 야간 15회 이하 엄격 적용)
    MAX_CC_PER_SHIFT = 15  # 시프트당 최대 컬러교환

    color_changes_in_rotation = [0] * 10
    rotation_color_detail = [{} for _ in range(10)]

    # 1단계: 컬러별 총 수요 계산
    color_total_demand = defaultdict(int)
    for x in items:
        if x['grp'] and x.get(demand_key, 0) > 0:
            color_total_demand[x['clr']] += x.get(demand_key, 0)

    sorted_colors = sorted(color_total_demand.keys(), key=lambda c: -color_total_demand[c])

    # 2단계: 시프트별 컬러 배정 (엄격하게 제한)
    # 15회 컬러교환 = 최대 16컬러 사용 가능
    # 5회전에 16컬러 = 회전당 약 3.2컬러
    MAX_COLORS_PER_ROTATION = 3  # 회전당 최대 3컬러로 제한

    # 시프트별 사용할 컬러 (수요 상위)
    day_colors = sorted_colors[:16]    # 주간용
    night_colors = sorted_colors[:16]  # 야간용

    # 3단계: 생산량 배정 (시프트 내 컬러교환 추적)
    day_shift_used_colors = set()   # 주간 전체 사용 컬러
    night_shift_used_colors = set() # 야간 전체 사용 컬러

    for r in range(10):
        is_day_shift = r < 5
        shift_colors = day_colors if is_day_shift else night_colors
        shift_used = day_shift_used_colors if is_day_shift else night_shift_used_colors

        tmpl = templates[r]
        rotation_used_colors = set()

        for g, h in tmpl.items():
            pcs = JIG_INVENTORY[g]['pcs']
            cap = h * JIGS_PER_HANGER * pcs

            # 시프트 컬러 내의 아이템만
            grp_items = [x for x in items if x['grp'] == g and x['clr'] in shift_colors]

            if not grp_items:
                grp_items = [x for x in items if x['grp'] == g and x.get(demand_key, 0) > 0]

            if not grp_items:
                grp_items = [x for x in items if x['grp'] == g]

            if not grp_items:
                rotation_color_detail[r][g] = {}
                continue

            # 지그그룹 내 수요 가장 많은 컬러 1개 선택
            # 단, 이미 시프트에서 사용 중인 컬러 우선
            clr_demand = defaultdict(int)
            for x in grp_items:
                clr_demand[x['clr']] += x.get(demand_key, 0)

            # 이미 사용 중인 컬러 중 수요 있는 것 우선
            used_with_demand = [c for c in shift_used if c in clr_demand and clr_demand[c] > 0]
            if used_with_demand:
                best_color = max(used_with_demand, key=lambda c: clr_demand[c])
            else:
                # 시프트 교환 한도 체크
                current_cc = len(shift_used)
                if current_cc < MAX_CC_PER_SHIFT + 1:  # 아직 여유 있음
                    best_color = max(clr_demand.keys(), key=lambda c: clr_demand[c])
                else:
                    # 한도 초과 - 기존 컬러만 사용
                    available = [c for c in shift_used if c in clr_demand]
                    if available:
                        best_color = max(available, key=lambda c: clr_demand[c])
                    else:
                        best_color = max(clr_demand.keys(), key=lambda c: clr_demand[c])

            selected_items = [x for x in grp_items if x['clr'] == best_color]

            if not selected_items:
                selected_items = grp_items[:1]
                best_color = selected_items[0]['clr']

            rotation_used_colors.add(best_color)
            shift_used.add(best_color)

            # 생산량 배분
            item_demand = {id(x): max(1, x.get(demand_key, 0)) for x in selected_items}
            tot = sum(item_demand.values())

            clr_production = defaultdict(int)
            for x in selected_items:
                ratio = item_demand[id(x)] / tot if tot else 1/len(selected_items)
                prod_qty = max(1, int(cap * ratio))
                x[prod_key][r] += prod_qty
                clr_production[x['clr']] += prod_qty

            clr_hangers = {}
            total_prod = sum(clr_production.values())
            for clr, prod in clr_production.items():
                hanger_count = int(h * prod / total_prod) if total_prod > 0 else 0
                if hanger_count > 0:
                    clr_hangers[clr] = hanger_count

            rotation_color_detail[r][g] = clr_hangers

        # 임시로 0 설정 (아래에서 재계산)
        color_changes_in_rotation[r] = 0

    # 부족분 보정
    for x in items:
        stk = x.get(start_stock_key, x['stk'])
        demand_arr = x.get(day_key, [0]*10)
        for r in range(10):
            stk = stk - demand_arr[r] + x[prod_key][r]
            if stk < 0:
                deficit = -stk
                for pr in range(r, -1, -1):
                    if x[prod_key][pr] > 0:
                        x[prod_key][pr] += deficit
                        stk += deficit
                        break
                else:
                    x[prod_key][r] += deficit
                    stk += deficit

    # 컬러 교환 정확히 계산 (지그 순서 기반)
    def get_grp_main_color(g, r):
        """지그그룹의 주 생산 컬러"""
        color_prod = defaultdict(int)
        for x in items:
            if x['grp'] == g and x[prod_key][r] > 0:
                color_prod[x['clr']] += x[prod_key][r]
        if color_prod:
            return max(color_prod.keys(), key=lambda c: color_prod[c])
        return None

    prev_rot_last_color = prev_day_color  # 전날 마지막 컬러
    total_cc = 0

    for r in range(10):
        tmpl = templates[r]
        order = jig_orders[r] if jig_orders[r] else sorted(tmpl.keys())

        # 순서대로 컬러 나열
        colors_in_order = []
        for g in order:
            if g in tmpl and tmpl[g] > 0:
                color = get_grp_main_color(g, r)
                if color:
                    colors_in_order.append(color)

        first_color = colors_in_order[0] if colors_in_order else None
        last_color = colors_in_order[-1] if colors_in_order else None

        # 회전 간 컬러 교환
        between_cc = 0
        if prev_rot_last_color and first_color and prev_rot_last_color != first_color:
            between_cc = 1

        # 회전 내 컬러 교환
        within_cc = 0
        for i in range(1, len(colors_in_order)):
            if colors_in_order[i] != colors_in_order[i-1]:
                within_cc += 1

        color_changes_in_rotation[r] = within_cc + between_cc
        total_cc += within_cc + between_cc
        prev_rot_last_color = last_color

    cc = total_cc

    # 마지막 순서 (다음 날 시작용)
    last_order = jig_orders[9]

    return templates, rotation_color, jig_changes, cc, color_changes_in_rotation, rotation_color_detail, jig_orders, last_order

def schedule(items):
    """D0, D+1, D+2 전체 스케줄링 (일간 연속성 고려)"""

    # D0 스케줄링 (첫날이므로 이전 상태 없음)
    (templates_d0, rotation_color_d0, jig_changes_d0, cc_d0, cc_per_rot_d0,
     color_detail_d0, jig_orders_d0, d0_last_order) = schedule_day(
        items, 'd0', 'd0t', 'prod', 'stk',
        prev_day_template=None, prev_day_color=None, prev_day_order=None)

    # ★ D+1 부족분 선반영: D0 생산만으로 D+1 수요까지 커버
    # D0 기말재고 계산 후 D+1 회전별 부족분 체크
    for x in items:
        # D0 기말재고 (임시 계산)
        d0_end = x['stk']
        for r in range(10):
            d0_end = d0_end - x['d0'][r] + x['prod'][r]

        # D+1 회전별 부족분 체크 (D0 생산만, D+1 생산 없음)
        running = d0_end
        d1_deficit = 0
        for r in range(10):
            running = running - x['d1'][r]
            if running < 0:
                d1_deficit = max(d1_deficit, -running)

        # D+1 부족분이 있으면 D0 마지막 회전에 추가 생산
        if d1_deficit > 0:
            # D0 생산 가능한 회전 찾기 (뒤에서부터)
            for pr in range(9, -1, -1):
                if x['prod'][pr] > 0:
                    x['prod'][pr] += d1_deficit
                    break
            else:
                # 생산 중인 회전이 없으면 마지막 회전에 추가
                x['prod'][9] += d1_deficit

    # D0 기말재고 계산
    for x in items:
        stk = x['stk']
        for r in range(10):
            stk = stk - x['d0'][r] + x['prod'][r]
        x['cur'] = stk  # D0 기말 = D+1 기초

    # ★ D0 컬러 교환 재계산 (D+1 부족분 선반영 후)
    def calc_color_changes(items, prod_key, templates, jig_orders, prev_day_color=None):
        """컬러 교환 횟수 계산 (지그 순서 기반)"""
        def get_grp_main_color(g, r):
            color_prod = defaultdict(int)
            for x in items:
                if x['grp'] == g and x[prod_key][r] > 0:
                    color_prod[x['clr']] += x[prod_key][r]
            if color_prod:
                return max(color_prod.keys(), key=lambda c: color_prod[c])
            return None

        cc_per_rot = [0] * 10
        prev_color = prev_day_color
        total = 0

        for r in range(10):
            tmpl = templates[r]
            order = jig_orders[r] if jig_orders and jig_orders[r] else sorted(tmpl.keys())

            colors = []
            for g in order:
                if g in tmpl and tmpl[g] > 0:
                    c = get_grp_main_color(g, r)
                    if c:
                        colors.append(c)

            # 회전 간 교환
            between_cc = 0
            if prev_color and colors and colors[0] != prev_color:
                between_cc = 1

            # 회전 내 교환
            within_cc = 0
            for i in range(1, len(colors)):
                if colors[i] != colors[i-1]:
                    within_cc += 1

            cc_per_rot[r] = within_cc + between_cc
            total += within_cc + between_cc

            if colors:
                prev_color = colors[-1]

        return total, cc_per_rot

    cc_d0, cc_per_rot_d0 = calc_color_changes(items, 'prod', templates_d0, jig_orders_d0, None)

    # D0 마지막 상태 → D+1 시작 조건
    d0_last_template = templates_d0[9]  # D0 10회전 템플릿
    d0_last_color = rotation_color_d0[9]  # D0 10회전 컬러

    # D+1 스케줄링 (D0 마지막 상태에서 시작)
    (templates_d1, rotation_color_d1, jig_changes_d1, cc_d1, cc_per_rot_d1,
     color_detail_d1, jig_orders_d1, d1_last_order) = schedule_day(
        items, 'd1', 'd1t', 'prod1', 'cur',
        prev_day_template=d0_last_template, prev_day_color=d0_last_color, prev_day_order=d0_last_order)

    # D+1 기말재고 계산
    for x in items:
        stk = x['cur']
        for r in range(10):
            stk = stk - x['d1'][r] + x['prod1'][r]
        x['cur1'] = stk  # D+1 기말 = D+2 기초

    # D+1 마지막 상태 → D+2 시작 조건
    d1_last_template = templates_d1[9]  # D+1 10회전 템플릿
    d1_last_color = rotation_color_d1[9]  # D+1 10회전 컬러

    # D+2 스케줄링 (D+1 마지막 상태에서 시작)
    (templates_d2, rotation_color_d2, jig_changes_d2, cc_d2, cc_per_rot_d2,
     color_detail_d2, jig_orders_d2, d2_last_order) = schedule_day(
        items, 'd2', 'd2t', 'prod2', 'cur1',
        prev_day_template=d1_last_template, prev_day_color=d1_last_color, prev_day_order=d1_last_order)

    # D+2 기말재고 계산
    for x in items:
        stk = x['cur1']
        for r in range(10):
            stk = stk - x['d2'][r] + x['prod2'][r]
        x['cur2'] = stk

    # 일간 전환 지그교체 계산 (D0→D+1, D+1→D+2)
    d0_to_d1_jig = calc_jig_change_simple(d0_last_template, templates_d1[0])
    d1_to_d2_jig = calc_jig_change_simple(d1_last_template, templates_d2[0])

    return {
        'd0': {
            'templates': templates_d0,
            'colors': rotation_color_d0,
            'jig_changes': jig_changes_d0,
            'cc': cc_d0,
            'cc_per_rotation': cc_per_rot_d0,
            'color_detail': color_detail_d0,
            'jig_orders': jig_orders_d0,
            'last_template': d0_last_template,
            'last_color': d0_last_color
        },
        'd1': {
            'templates': templates_d1,
            'colors': rotation_color_d1,
            'jig_changes': jig_changes_d1,
            'cc': cc_d1,
            'cc_per_rotation': cc_per_rot_d1,
            'color_detail': color_detail_d1,
            'jig_orders': jig_orders_d1,
            'last_template': d1_last_template,
            'last_color': d1_last_color,
            'start_jig_change': d0_to_d1_jig
        },
        'd2': {
            'templates': templates_d2,
            'colors': rotation_color_d2,
            'jig_changes': jig_changes_d2,
            'cc': cc_d2,
            'cc_per_rotation': cc_per_rot_d2,
            'color_detail': color_detail_d2,
            'jig_orders': jig_orders_d2,
            'start_jig_change': d1_to_d2_jig
        },
    }

def get_rotation_items_detail(items, rotation, prod_key, templates, jig_orders):
    """회전별 생산 아이템 상세 (지그그룹 순서대로)
    Returns: list of (차종, 아이템, 세부, 컬러, 수량, 지그그룹) tuples
    """
    tmpl = templates[rotation]
    order = jig_orders[rotation] if jig_orders and jig_orders[rotation] else sorted(tmpl.keys())

    result = []
    for g in order:
        if g not in tmpl or tmpl[g] == 0:
            continue
        # 이 회전에서 이 지그그룹으로 생산되는 아이템들
        grp_items = [(x, x[prod_key][rotation]) for x in items
                     if x['grp'] == g and x[prod_key][rotation] > 0]
        # 생산량 내림차순
        grp_items.sort(key=lambda x: -x[1])
        for x, prod in grp_items:
            ct = x['ct'].replace('\n', ' ').strip()
            it = x['it'].replace('\n', ' ').strip()
            det = x['det'].replace('\n', ' ').strip() if x.get('det') else '-'
            result.append((ct, it, det, x['clr'], prod, g))
    return result

def format_rotation_items_html(items, rotation, prod_key, templates, jig_orders):
    """회전별 생산 아이템을 HTML 박스 형태로 포맷"""
    details = get_rotation_items_detail(items, rotation, prod_key, templates, jig_orders)
    if not details:
        return "<span style='color:#999;'>-</span>"

    # 지그그룹 색상
    grp_colors = {
        'A': '#E3F2FD', 'B': '#E8F5E9', 'B2': '#C8E6C9', 'C': '#FFF3E0', 'D': '#FCE4EC',
        'E': '#F3E5F5', 'F': '#E0F7FA', 'G': '#FFF8E1', 'H': '#EFEBE9', 'I': '#ECEFF1'
    }
    grp_border = {
        'A': '#1976D2', 'B': '#388E3C', 'B2': '#2E7D32', 'C': '#F57C00', 'D': '#C2185B',
        'E': '#7B1FA2', 'F': '#0097A7', 'G': '#FFA000', 'H': '#5D4037', 'I': '#455A64'
    }

    # 지그그룹별로 묶어서 박스 표시
    result_parts = []
    current_grp = None
    grp_items = []

    tmpl = templates[rotation]

    for ct, it, det, clr, prod, g in details:
        if g != current_grp:
            if grp_items and current_grp:
                bg = grp_colors.get(current_grp, '#F5F5F5')
                border = grp_border.get(current_grp, '#9E9E9E')
                h_count = tmpl.get(current_grp, 0)
                items_html = "<br>".join([f"&nbsp;{x}" for x in grp_items])
                result_parts.append(
                    f'<div style="display:inline-block;vertical-align:top;margin:2px;padding:4px 8px;'
                    f'background:{bg};border:2px solid {border};border-radius:6px;min-width:150px;">'
                    f'<div style="font-weight:bold;color:{border};border-bottom:1px solid {border};margin-bottom:3px;padding-bottom:2px;">'
                    f'{current_grp} ({h_count}H)</div>'
                    f'<div style="font-size:0.7em;line-height:1.5;">{items_html}</div>'
                    f'</div>'
                )
            current_grp = g
            grp_items = []

        # 차종 / 아이템 / 세부 / 컬러 / 수량
        det_str = f" {det}" if det and det != '-' else ""
        grp_items.append(f"<b>{ct}</b> {it}{det_str} <span style='color:#1565C0;'>{clr}</span> <span style='color:#D32F2F;font-weight:bold;'>{prod}</span>")

    # 마지막 그룹
    if grp_items and current_grp:
        bg = grp_colors.get(current_grp, '#F5F5F5')
        border = grp_border.get(current_grp, '#9E9E9E')
        h_count = tmpl.get(current_grp, 0)
        items_html = "<br>".join([f"&nbsp;{x}" for x in grp_items])
        result_parts.append(
            f'<div style="display:inline-block;vertical-align:top;margin:2px;padding:4px 8px;'
            f'background:{bg};border:2px solid {border};border-radius:6px;min-width:150px;">'
            f'<div style="font-weight:bold;color:{border};border-bottom:1px solid {border};margin-bottom:3px;padding-bottom:2px;">'
            f'{current_grp} ({h_count}H)</div>'
            f'<div style="font-size:0.7em;line-height:1.5;">{items_html}</div>'
            f'</div>'
        )

    return "".join(result_parts)

def generate_html_report(items, schedule_result):
    today = datetime.now()
    tomorrow = today + timedelta(days=1)
    day_after = today + timedelta(days=2)

    d0 = schedule_result['d0']
    d1 = schedule_result['d1']
    d2 = schedule_result['d2']

    templates = d0['templates']
    rotation_color = d0['colors']
    jig_changes = d0['jig_changes']
    cc = d0['cc']
    jig_orders_d0 = d0.get('jig_orders', [None]*10)
    jig_orders_d1 = d1.get('jig_orders', [None]*10)
    jig_orders_d2 = d2.get('jig_orders', [None]*10)

    day_jig = sum(jig_changes[:5])
    night_jig = sum(jig_changes[5:])
    total_prod_d0 = sum(sum(x['prod']) for x in items)
    total_prod_d1 = sum(sum(x['prod1']) for x in items)
    total_prod_d2 = sum(sum(x['prod2']) for x in items)

    html = f'''<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>D0 생산계획 리포트 - {today.strftime("%Y-%m-%d")}</title>
    <style>
        * {{ box-sizing: border-box; }}
        body {{
            font-family: 'Malgun Gothic', sans-serif;
            margin: 20px;
            background: #f5f5f5;
        }}
        h1 {{ color: #333; border-bottom: 3px solid #2196F3; padding-bottom: 10px; }}
        h2 {{ color: #1976D2; margin-top: 30px; }}
        h3 {{ color: #424242; }}

        .container {{ max-width: 1800px; margin: 0 auto; }}

        .card {{
            background: white;
            border-radius: 8px;
            padding: 20px;
            margin: 15px 0;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}

        .params-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
        }}

        .param-item {{
            background: #E3F2FD;
            padding: 12px;
            border-radius: 6px;
            border-left: 4px solid #2196F3;
        }}

        .param-label {{ font-weight: bold; color: #1565C0; }}
        .param-value {{ font-size: 1.2em; color: #333; margin-top: 5px; }}

        .jig-table {{
            width: 100%;
            border-collapse: collapse;
            margin: 10px 0;
        }}

        .jig-table th, .jig-table td {{
            border: 1px solid #ddd;
            padding: 8px;
            text-align: center;
        }}

        .jig-table th {{
            background: #2196F3;
            color: white;
        }}

        .jig-table tr:nth-child(even) {{ background: #f9f9f9; }}

        .rotation-table {{
            width: 100%;
            border-collapse: collapse;
            margin: 10px 0;
            font-size: 0.9em;
        }}

        .rotation-table th, .rotation-table td {{
            border: 1px solid #ddd;
            padding: 6px;
            text-align: center;
        }}

        .rotation-table th {{
            background: #1976D2;
            color: white;
            position: sticky;
            top: 0;
        }}

        .rotation-table .day {{ background: #FFF3E0; }}
        .rotation-table .night {{ background: #E8EAF6; }}

        .main-table-container {{
            overflow-x: auto;
            max-height: 800px;
            overflow-y: auto;
        }}

        .main-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 0.85em;
        }}

        .main-table th {{
            background: #37474F;
            color: white;
            padding: 8px 4px;
            position: sticky;
            top: 0;
            z-index: 10;
        }}

        .main-table td {{
            border: 1px solid #ddd;
            padding: 4px;
            text-align: center;
        }}

        .main-table tr:nth-child(even) {{ background: #fafafa; }}
        .main-table tr:hover {{ background: #e3f2fd; }}

        .shortage {{
            background: #FFCDD2 !important;
            color: #B71C1C;
            font-weight: bold;
        }}

        .header-group {{
            background: #546E7A !important;
        }}

        .summary-box {{
            display: flex;
            gap: 20px;
            flex-wrap: wrap;
        }}

        .summary-item {{
            flex: 1;
            min-width: 150px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 10px;
            text-align: center;
        }}

        .summary-item.warning {{
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        }}

        .summary-item.success {{
            background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
        }}

        .summary-number {{ font-size: 2em; font-weight: bold; }}
        .summary-label {{ font-size: 0.9em; opacity: 0.9; }}

        .legend {{
            display: flex;
            gap: 20px;
            margin: 10px 0;
            font-size: 0.9em;
        }}

        .legend-item {{
            display: flex;
            align-items: center;
            gap: 5px;
        }}

        .legend-color {{
            width: 20px;
            height: 20px;
            border-radius: 4px;
        }}

        .legend-shortage {{ background: #FFCDD2; border: 1px solid #EF9A9A; }}

        @media print {{
            body {{ margin: 0; }}
            .card {{ box-shadow: none; border: 1px solid #ddd; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>D0 생산계획 리포트</h1>
        <p>생성일시: {today.strftime("%Y-%m-%d %H:%M:%S")}</p>

        <div class="card">
            <h2>시스템 변수</h2>
            <div class="params-grid">
                <div class="param-item">
                    <div class="param-label">총 행어 수</div>
                    <div class="param-value">{HANGERS}개</div>
                </div>
                <div class="param-item">
                    <div class="param-label">행어당 지그 수</div>
                    <div class="param-value">{JIGS_PER_HANGER}개</div>
                </div>
                <div class="param-item">
                    <div class="param-label">일일 회전 수</div>
                    <div class="param-value">{ROTATIONS_PER_DAY}회전</div>
                </div>
                <div class="param-item">
                    <div class="param-label">주간 회전</div>
                    <div class="param-value">1~{DAY_SHIFT_ROTATIONS}회전</div>
                </div>
                <div class="param-item">
                    <div class="param-label">야간 회전</div>
                    <div class="param-value">{DAY_SHIFT_ROTATIONS+1}~{ROTATIONS_PER_DAY}회전</div>
                </div>
                <div class="param-item">
                    <div class="param-label">주간 지그교체 예산</div>
                    <div class="param-value">{JIG_BUDGET_DAY}개</div>
                </div>
                <div class="param-item">
                    <div class="param-label">야간 지그교체 예산</div>
                    <div class="param-value">{JIG_BUDGET_NIGHT}개</div>
                </div>
                <div class="param-item">
                    <div class="param-label">컬러교환 손실</div>
                    <div class="param-value">{COLOR_CHANGE_LOSS}개/회</div>
                </div>
                <div class="param-item">
                    <div class="param-label">안전재고 일수</div>
                    <div class="param-value">{SAFETY_STOCK_DAYS}일</div>
                </div>
                <div class="param-item">
                    <div class="param-label">일일 최대 용량</div>
                    <div class="param-value">{HANGERS * JIGS_PER_HANGER * ROTATIONS_PER_DAY:,}개</div>
                </div>
            </div>
        </div>

        <div class="card">
            <h2>지그 그룹 정보</h2>
            <table class="jig-table">
                <tr>
                    <th>그룹</th>
                    <th>명칭</th>
                    <th>최대 지그수</th>
                    <th>최대 행어</th>
                    <th>PCS/지그</th>
                    <th>최대 용량/회전</th>
                </tr>'''

    for g in sorted(JIG_INVENTORY.keys()):
        info = JIG_INVENTORY[g]
        max_h = info['max_jigs'] // JIGS_PER_HANGER
        max_cap = max_h * JIGS_PER_HANGER * info['pcs']
        html += f'''
                <tr>
                    <td><strong>{g}</strong></td>
                    <td>{info['name']}</td>
                    <td>{info['max_jigs']}</td>
                    <td>{max_h}</td>
                    <td>{info['pcs']}</td>
                    <td>{max_cap}</td>
                </tr>'''

    html += '''
            </table>
        </div>

        <div class="card">
            <h2>실행 결과 요약</h2>
            <div class="summary-box">
                <div class="summary-item success">
                    <div class="summary-number">{:,}</div>
                    <div class="summary-label">D0 생산량</div>
                </div>
                <div class="summary-item success">
                    <div class="summary-number">{:,}</div>
                    <div class="summary-label">D+1 생산량</div>
                </div>
                <div class="summary-item success">
                    <div class="summary-number">{:,}</div>
                    <div class="summary-label">D+2 생산량</div>
                </div>
                <div class="summary-item">
                    <div class="summary-number">{}</div>
                    <div class="summary-label">D0 컬러교환</div>
                </div>
                <div class="summary-item warning">
                    <div class="summary-number">{}</div>
                    <div class="summary-label">D0 컬러손실</div>
                </div>
                <div class="summary-item">
                    <div class="summary-number">{}/{}</div>
                    <div class="summary-label">D0 주간지그</div>
                </div>
                <div class="summary-item">
                    <div class="summary-number">{}/{}</div>
                    <div class="summary-label">D0 야간지그</div>
                </div>
            </div>
        </div>
'''.format(total_prod_d0, total_prod_d1, total_prod_d2, cc, cc * COLOR_CHANGE_LOSS, day_jig, JIG_BUDGET_DAY, night_jig, JIG_BUDGET_NIGHT)

    # 재고/수요/생산 합계 테이블
    init_stock = sum(x['stk'] for x in items)
    d0_demand = sum(x['d0t'] for x in items)
    d0_prod = sum(sum(x['prod']) for x in items)
    d0_end_stock = sum(x['cur'] for x in items)
    d1_demand = sum(x['d1t'] for x in items)
    d1_prod = sum(sum(x['prod1']) for x in items)
    d1_end_stock = sum(x['cur1'] for x in items)
    d2_demand = sum(x['d2t'] for x in items)
    d2_prod = sum(sum(x['prod2']) for x in items)
    d2_end_stock = sum(x['cur2'] for x in items)

    html += f'''
        <div class="card">
            <h2>재고/수요/생산 합계</h2>
            <table class="summary-table" style="width:100%;border-collapse:collapse;text-align:center;">
                <tr style="background:#37474F;color:white;">
                    <th style="padding:10px;border:1px solid #ccc;">구분</th>
                    <th style="padding:10px;border:1px solid #ccc;">기초재고</th>
                    <th style="padding:10px;border:1px solid #ccc;background:#1565C0;">D0 수요</th>
                    <th style="padding:10px;border:1px solid #ccc;background:#1565C0;">D0 생산</th>
                    <th style="padding:10px;border:1px solid #ccc;background:#1565C0;">D0 기말재고</th>
                    <th style="padding:10px;border:1px solid #ccc;background:#2E7D32;">D+1 수요</th>
                    <th style="padding:10px;border:1px solid #ccc;background:#2E7D32;">D+1 생산</th>
                    <th style="padding:10px;border:1px solid #ccc;background:#2E7D32;">D+1 기말재고</th>
                    <th style="padding:10px;border:1px solid #ccc;background:#E65100;">D+2 수요</th>
                    <th style="padding:10px;border:1px solid #ccc;background:#E65100;">D+2 생산</th>
                    <th style="padding:10px;border:1px solid #ccc;background:#E65100;">D+2 기말재고</th>
                </tr>
                <tr style="font-size:1.2em;font-weight:bold;">
                    <td style="padding:12px;border:1px solid #ccc;background:#ECEFF1;">합계</td>
                    <td style="padding:12px;border:1px solid #ccc;">{init_stock:,}</td>
                    <td style="padding:12px;border:1px solid #ccc;background:#E3F2FD;">{d0_demand:,}</td>
                    <td style="padding:12px;border:1px solid #ccc;background:#E3F2FD;color:#1565C0;">{d0_prod:,}</td>
                    <td style="padding:12px;border:1px solid #ccc;background:#E3F2FD;">{d0_end_stock:,}</td>
                    <td style="padding:12px;border:1px solid #ccc;background:#E8F5E9;">{d1_demand:,}</td>
                    <td style="padding:12px;border:1px solid #ccc;background:#E8F5E9;color:#2E7D32;">{d1_prod:,}</td>
                    <td style="padding:12px;border:1px solid #ccc;background:#E8F5E9;">{d1_end_stock:,}</td>
                    <td style="padding:12px;border:1px solid #ccc;background:#FFF3E0;">{d2_demand:,}</td>
                    <td style="padding:12px;border:1px solid #ccc;background:#FFF3E0;color:#E65100;">{d2_prod:,}</td>
                    <td style="padding:12px;border:1px solid #ccc;background:#FFF3E0;">{d2_end_stock:,}</td>
                </tr>
                <tr style="font-size:0.9em;color:#666;">
                    <td style="padding:8px;border:1px solid #ccc;background:#ECEFF1;">변동</td>
                    <td style="padding:8px;border:1px solid #ccc;">-</td>
                    <td style="padding:8px;border:1px solid #ccc;background:#E3F2FD;">-{d0_demand:,}</td>
                    <td style="padding:8px;border:1px solid #ccc;background:#E3F2FD;">+{d0_prod:,}</td>
                    <td style="padding:8px;border:1px solid #ccc;background:#E3F2FD;">{d0_end_stock - init_stock:+,}</td>
                    <td style="padding:8px;border:1px solid #ccc;background:#E8F5E9;">-{d1_demand:,}</td>
                    <td style="padding:8px;border:1px solid #ccc;background:#E8F5E9;">+{d1_prod:,}</td>
                    <td style="padding:8px;border:1px solid #ccc;background:#E8F5E9;">{d1_end_stock - d0_end_stock:+,}</td>
                    <td style="padding:8px;border:1px solid #ccc;background:#FFF3E0;">-{d2_demand:,}</td>
                    <td style="padding:8px;border:1px solid #ccc;background:#FFF3E0;">+{d2_prod:,}</td>
                    <td style="padding:8px;border:1px solid #ccc;background:#FFF3E0;">{d2_end_stock - d1_end_stock:+,}</td>
                </tr>
            </table>
        </div>
'''

    # 회전별 생산 상세 (지그 순서대로 아이템 표시) - 박스 스타일
    templates_d1 = d1['templates']
    jig_changes_d1 = d1['jig_changes']
    templates_d2_box = d2['templates']
    jig_changes_d2_box = d2['jig_changes']

    html += '''
        <div class="card">
            <h2>회전별 생산 상세 (지그 순서)</h2>
            <p>각 회전에서 생산되는 아이템을 컨베이어 지그 순서대로 표시 (박스 = 지그그룹, 숫자 = 행어수)</p>
'''

    # D0 회전별 상세
    html += '<h3 style="margin-top:20px;color:#1565C0;">D0 생산계획</h3>'
    for r in range(10):
        shift_name = '주간' if r < 5 else '야간'
        shift_bg = '#E3F2FD' if r < 5 else '#E8EAF6'
        detail_html = format_rotation_items_html(items, r, 'prod', templates, jig_orders_d0)
        html += f'''
            <div style="margin:8px 0;padding:10px;background:{shift_bg};border-radius:8px;">
                <div style="display:flex;align-items:center;gap:15px;margin-bottom:8px;">
                    <span style="font-weight:bold;font-size:1.1em;color:#1565C0;">D0-{r+1}</span>
                    <span style="background:#1976D2;color:white;padding:2px 8px;border-radius:4px;font-size:0.8em;">{shift_name}</span>
                    <span style="color:#666;font-size:0.85em;">지그교체: <b>{jig_changes[r]}</b></span>
                </div>
                <div style="display:flex;flex-wrap:wrap;gap:4px;">{detail_html}</div>
            </div>'''

    # D0→D+1 전환
    html += f'''
            <div style="margin:15px 0;padding:10px;background:#FFE082;border-radius:8px;text-align:center;">
                <strong>▶ D0→D+1 전환</strong> | 지그교체: <b>{d1.get('start_jig_change', 0)}</b>
            </div>'''

    # D+1 회전별 상세
    html += '<h3 style="margin-top:20px;color:#2E7D32;">D+1 생산계획</h3>'
    for r in range(10):
        shift_name = '주간' if r < 5 else '야간'
        shift_bg = '#E8F5E9' if r < 5 else '#F1F8E9'
        detail_html = format_rotation_items_html(items, r, 'prod1', templates_d1, jig_orders_d1)
        html += f'''
            <div style="margin:8px 0;padding:10px;background:{shift_bg};border-radius:8px;">
                <div style="display:flex;align-items:center;gap:15px;margin-bottom:8px;">
                    <span style="font-weight:bold;font-size:1.1em;color:#2E7D32;">D+1-{r+1}</span>
                    <span style="background:#388E3C;color:white;padding:2px 8px;border-radius:4px;font-size:0.8em;">{shift_name}</span>
                    <span style="color:#666;font-size:0.85em;">지그교체: <b>{jig_changes_d1[r]}</b></span>
                </div>
                <div style="display:flex;flex-wrap:wrap;gap:4px;">{detail_html}</div>
            </div>'''

    # D+1→D+2 전환
    html += f'''
            <div style="margin:15px 0;padding:10px;background:#FFE082;border-radius:8px;text-align:center;">
                <strong>▶ D+1→D+2 전환</strong> | 지그교체: <b>{d2.get('start_jig_change', 0)}</b>
            </div>'''

    # D+2 회전별 상세
    html += '<h3 style="margin-top:20px;color:#E65100;">D+2 생산계획</h3>'
    for r in range(10):
        shift_name = '주간' if r < 5 else '야간'
        shift_bg = '#FFF3E0' if r < 5 else '#FBE9E7'
        detail_html = format_rotation_items_html(items, r, 'prod2', templates_d2_box, jig_orders_d2)
        html += f'''
            <div style="margin:8px 0;padding:10px;background:{shift_bg};border-radius:8px;">
                <div style="display:flex;align-items:center;gap:15px;margin-bottom:8px;">
                    <span style="font-weight:bold;font-size:1.1em;color:#E65100;">D+2-{r+1}</span>
                    <span style="background:#F57C00;color:white;padding:2px 8px;border-radius:4px;font-size:0.8em;">{shift_name}</span>
                    <span style="color:#666;font-size:0.85em;">지그교체: <b>{jig_changes_d2_box[r]}</b></span>
                </div>
                <div style="display:flex;flex-wrap:wrap;gap:4px;">{detail_html}</div>
            </div>'''

    html += '''
        </div>
'''

    # 메인 생산계획 테이블 - D0, D+1, D+2 모두 회전별 표시
    html += f'''
        <div class="card">
            <h2>생산계획 상세 (D0: {today.strftime("%m/%d")}, D+1: {tomorrow.strftime("%m/%d")}, D+2: {day_after.strftime("%m/%d")})</h2>
            <div class="legend">
                <div class="legend-item">
                    <div class="legend-color legend-shortage"></div>
                    <span>재고 부족 (재고 < 0)</span>
                </div>
            </div>
            <div class="main-table-container">
                <table class="main-table">
                    <tr>
                        <th rowspan="2">차종</th>
                        <th rowspan="2">아이템</th>
                        <th rowspan="2">세부</th>
                        <th rowspan="2">컬러</th>
                        <th rowspan="2">지그</th>
                        <th rowspan="2">기초<br>재고</th>'''

    # D0 헤더
    html += '<th colspan="30" style="background:#1565C0;">D0 ({}) - 10회전</th>'.format(today.strftime("%m/%d"))
    html += '<th rowspan="2">D0<br>기말</th>'

    # D+1 헤더
    html += '<th colspan="30" style="background:#2E7D32;">D+1 ({}) - 10회전</th>'.format(tomorrow.strftime("%m/%d"))
    html += '<th rowspan="2">D+1<br>기말</th>'

    # D+2 헤더
    html += '<th colspan="30" style="background:#E65100;">D+2 ({}) - 10회전</th>'.format(day_after.strftime("%m/%d"))
    html += '<th rowspan="2">D+2<br>기말</th>'

    html += '</tr><tr>'

    # D0 회전별 헤더
    for i in range(1, 11):
        html += f'<th style="background:#1976D2;font-size:0.7em;">{i}H<br>수</th><th style="background:#1976D2;font-size:0.7em;">{i}H<br>생</th><th style="background:#1976D2;font-size:0.7em;">{i}H<br>재</th>'

    # D+1 회전별 헤더
    for i in range(1, 11):
        html += f'<th style="background:#388E3C;font-size:0.7em;">{i}H<br>수</th><th style="background:#388E3C;font-size:0.7em;">{i}H<br>생</th><th style="background:#388E3C;font-size:0.7em;">{i}H<br>재</th>'

    # D+2 회전별 헤더
    for i in range(1, 11):
        html += f'<th style="background:#F57C00;font-size:0.7em;">{i}H<br>수</th><th style="background:#F57C00;font-size:0.7em;">{i}H<br>생</th><th style="background:#F57C00;font-size:0.7em;">{i}H<br>재</th>'

    html += '</tr>'

    # 데이터 행
    for x in sorted(items, key=lambda x: (x['grp'] or 'Z', x['ct'], x['it'], x['clr'])):
        html += f'''
                    <tr>
                        <td style="font-size:0.75em;">{x['ct'][:10]}</td>
                        <td style="font-size:0.75em;">{x['it'][:12]}</td>
                        <td style="font-size:0.75em;">{x['det'][:6]}</td>
                        <td>{x['clr']}</td>
                        <td>{x['grp'] or '-'}</td>
                        <td>{x['stk']}</td>'''

        # D0 회전별
        running_stk = x['stk']
        for r in range(10):
            d = x['d0'][r]
            p = x['prod'][r]
            running_stk = running_stk - d + p
            shortage = running_stk < 0
            html += f'<td>{d}</td><td>{p}</td><td class="{"shortage" if shortage else ""}">{running_stk}</td>'

        html += f'<td><strong>{x["cur"]}</strong></td>'

        # D+1 회전별
        running_stk = x['cur']
        for r in range(10):
            d = x['d1'][r]
            p = x['prod1'][r]
            running_stk = running_stk - d + p
            shortage = running_stk < 0
            html += f'<td>{d}</td><td>{p}</td><td class="{"shortage" if shortage else ""}">{running_stk}</td>'

        html += f'<td><strong>{x["cur1"]}</strong></td>'

        # D+2 회전별
        running_stk = x['cur1']
        for r in range(10):
            d = x['d2'][r]
            p = x['prod2'][r]
            running_stk = running_stk - d + p
            shortage = running_stk < 0
            html += f'<td>{d}</td><td>{p}</td><td class="{"shortage" if shortage else ""}">{running_stk}</td>'

        html += f'<td><strong>{x["cur2"]}</strong></td></tr>'

    html += '''
                </table>
            </div>
        </div>
'''

    # D0 생산만으로 D0/D+1/D+2 재고 체크 테이블 (회전별 세분화 - 모든 일차)
    html += f'''
        <div class="card">
            <h2>D0 생산만 고려 시 재고 전망 (D0/D+1/D+2 회전별)</h2>
            <p>D0 생산만 있다고 가정할 때, D0/D+1/D+2 모든 회전별 재고 전망 (D+1/D+2 생산 없음)</p>
            <div class="legend">
                <div class="legend-item">
                    <div class="legend-color legend-shortage"></div>
                    <span>재고 부족 (재고 < 0)</span>
                </div>
            </div>
            <div class="main-table-container" style="max-height:700px;">
                <table class="main-table">
                    <tr>
                        <th rowspan="2">차종</th>
                        <th rowspan="2">아이템</th>
                        <th rowspan="2">컬러</th>
                        <th rowspan="2">지그</th>
                        <th rowspan="2">기초</th>
                        <th colspan="20" style="background:#1565C0;">D0 (수요/재고)</th>
                        <th colspan="20" style="background:#2E7D32;">D+1 (수요/재고) - D0생산만</th>
                        <th colspan="20" style="background:#E65100;">D+2 (수요/재고) - D0생산만</th>
                        <th rowspan="2">상태</th>
                    </tr>
                    <tr>'''

    # D0 10회전 헤더 (수요/재고만)
    for i in range(1, 11):
        html += f'<th style="background:#1976D2;font-size:0.6em;">{i}수</th><th style="background:#1976D2;font-size:0.6em;">{i}재</th>'

    # D+1 10회전 헤더
    for i in range(1, 11):
        html += f'<th style="background:#388E3C;font-size:0.6em;">{i}수</th><th style="background:#388E3C;font-size:0.6em;">{i}재</th>'

    # D+2 10회전 헤더
    for i in range(1, 11):
        html += f'<th style="background:#F57C00;font-size:0.6em;">{i}수</th><th style="background:#F57C00;font-size:0.6em;">{i}재</th>'

    html += '</tr>'

    # D0 생산만으로 D0/D+1/D+2 재고 계산
    shortage_items = []
    for x in sorted(items, key=lambda x: (x['grp'] or 'Z', x['ct'], x['it'], x['clr'])):
        d0_prod = sum(x['prod'])

        # 회전별 부족 체크
        d0_shortages = []
        d1_shortages = []
        d2_shortages = []

        # D0 회전별 재고 (생산 포함)
        running = x['stk']
        d0_stocks = []
        for r in range(10):
            running = running - x['d0'][r] + x['prod'][r]
            d0_stocks.append(running)
            d0_shortages.append(running < 0)
        d0_end = running

        # D+1 회전별 재고 (D0 생산만, D+1 생산 없음)
        running = d0_end
        d1_stocks = []
        for r in range(10):
            running = running - x['d1'][r]  # 생산 없음
            d1_stocks.append(running)
            d1_shortages.append(running < 0)

        # D+2 회전별 재고 (D0 생산만, D+2 생산 없음)
        d2_stocks = []
        for r in range(10):
            running = running - x['d2'][r]  # 생산 없음
            d2_stocks.append(running)
            d2_shortages.append(running < 0)

        # 상태 결정
        status = "OK"
        if any(d0_shortages):
            for r, short in enumerate(d0_shortages):
                if short:
                    status = f"D0-{r+1}H"
                    break
        elif any(d1_shortages):
            for r, short in enumerate(d1_shortages):
                if short:
                    status = f"D1-{r+1}H"
                    break
        elif any(d2_shortages):
            for r, short in enumerate(d2_shortages):
                if short:
                    status = f"D2-{r+1}H"
                    break

        if any(d0_shortages) or any(d1_shortages) or any(d2_shortages):
            shortage_items.append(x)

        html += f'''
                    <tr>
                        <td style="font-size:0.65em;">{x['ct'][:8]}</td>
                        <td style="font-size:0.65em;">{x['it'][:10]}</td>
                        <td style="font-size:0.7em;">{x['clr']}</td>
                        <td>{x['grp'] or '-'}</td>
                        <td>{x['stk']}</td>'''

        # D0 회전별 (수요/재고)
        for r in range(10):
            d = x['d0'][r]
            stk = d0_stocks[r]
            shortage = d0_shortages[r]
            html += f'<td style="font-size:0.65em;">{d if d else ""}</td><td class="{"shortage" if shortage else ""}" style="font-size:0.65em;">{stk}</td>'

        # D+1 회전별 (수요/재고 - D0 생산만)
        for r in range(10):
            d = x['d1'][r]
            stk = d1_stocks[r]
            shortage = d1_shortages[r]
            html += f'<td style="font-size:0.65em;">{d if d else ""}</td><td class="{"shortage" if shortage else ""}" style="font-size:0.65em;">{stk}</td>'

        # D+2 회전별 (수요/재고 - D0 생산만)
        for r in range(10):
            d = x['d2'][r]
            stk = d2_stocks[r]
            shortage = d2_shortages[r]
            html += f'<td style="font-size:0.65em;">{d if d else ""}</td><td class="{"shortage" if shortage else ""}" style="font-size:0.65em;">{stk}</td>'

        html += f'<td style="{"background:#FFCDD2;font-weight:bold;font-size:0.7em;" if status != "OK" else "font-size:0.7em;"}">{status}</td></tr>'''

    html += '''
                </table>
            </div>
        </div>
'''

    # 부족 요약
    d0_short = len([x for x in items if x['stk'] - x['d0t'] + sum(x['prod']) < 0])
    d1_short = len([x for x in items if x['stk'] - x['d0t'] + sum(x['prod']) - x['d1t'] < 0])
    d2_short = len([x for x in items if x['stk'] - x['d0t'] + sum(x['prod']) - x['d1t'] - x['d2t'] < 0])

    html += f'''
        <div class="card">
            <h2>D0 생산만 고려 시 부족 요약</h2>
            <div class="summary-box">
                <div class="summary-item {'warning' if d0_short > 0 else 'success'}">
                    <div class="summary-number">{d0_short}</div>
                    <div class="summary-label">D0 부족 아이템</div>
                </div>
                <div class="summary-item {'warning' if d1_short > 0 else 'success'}">
                    <div class="summary-number">{d1_short}</div>
                    <div class="summary-label">D+1 부족 아이템</div>
                </div>
                <div class="summary-item {'warning' if d2_short > 0 else 'success'}">
                    <div class="summary-number">{d2_short}</div>
                    <div class="summary-label">D+2 부족 아이템</div>
                </div>
            </div>
            <p style="margin-top:15px;color:#666;">
                * D0 생산만 있다고 가정했을 때의 부족 아이템 수<br>
                * D+1, D+2 생산을 추가하면 부족이 해소될 수 있음
            </p>
        </div>

        <div class="card">
            <h2>CSV 다운로드</h2>
            <p>production_plan_v8.csv 파일이 자동 생성됩니다.</p>
        </div>
    </div>
</body>
</html>'''

    return html

def save_csv(items):
    """D0, D+1, D+2 회전별 CSV 저장"""
    header = '차종,아이템,세부,컬러,지그,기초재고'
    # D0
    for i in range(1, 11):
        header += f',D0_{i}H수요,D0_{i}H생산,D0_{i}H재고'
    header += ',D0기말'
    # D+1
    for i in range(1, 11):
        header += f',D1_{i}H수요,D1_{i}H생산,D1_{i}H재고'
    header += ',D1기말'
    # D+2
    for i in range(1, 11):
        header += f',D2_{i}H수요,D2_{i}H생산,D2_{i}H재고'
    header += ',D2기말'

    lines = [header]
    for x in sorted(items, key=lambda x: (x['grp'] or 'Z', x['ct'], x['it'], x['clr'])):
        row = [
            x['ct'].replace(',',''),
            x['it'].replace(',',''),
            x.get('det','-').replace(',',''),
            x['clr'],
            x['grp'] or '-',
            str(x['stk'])
        ]
        # D0
        c = x['stk']
        for r in range(10):
            d, p = x['d0'][r], x['prod'][r]
            c = c - d + p
            row += [str(d), str(p), str(c)]
        row.append(str(x['cur']))
        # D+1
        c = x['cur']
        for r in range(10):
            d, p = x['d1'][r], x['prod1'][r]
            c = c - d + p
            row += [str(d), str(p), str(c)]
        row.append(str(x['cur1']))
        # D+2
        c = x['cur1']
        for r in range(10):
            d, p = x['d2'][r], x['prod2'][r]
            c = c - d + p
            row += [str(d), str(p), str(c)]
        row.append(str(x['cur2']))

        lines.append(','.join(row))
    with open('production_plan_v8.csv','w',encoding='utf-8-sig') as f:
        f.write('\n'.join(lines))

if __name__ == '__main__':
    print("데이터 로드...")
    items = load_data()
    print(f"{len(items)}개 아이템")

    print("스케줄링...")
    schedule_result = schedule(items)

    print("HTML 리포트 생성...")
    html = generate_html_report(items, schedule_result)

    with open('production_report.html', 'w', encoding='utf-8') as f:
        f.write(html)

    print("CSV 저장...")
    save_csv(items)

    print("\n완료!")
    print("  => production_report.html")
    print("  => production_plan_v8.csv")
