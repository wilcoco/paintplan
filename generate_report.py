#!/usr/bin/env python3
"""
D0 생산계획 웹 리포트 생성 v8.16
- 제약조건: 지그교체 150/시프트, 회전별 수요 충족
- 목적: 컬러교환 최소화
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
    """아이템을 지그그룹에 배정"""
    ct = ct.upper().replace(' ','').replace('\n','')
    it = it.upper().replace(' ','').replace('\n','')
    det = det.upper().replace(' ','').replace('\n','') if det else ''

    if 'TH' in ct:
        if 'STD' in it or 'LDT' in it: return 'A'
        if 'RR' in it: return 'H'
    if 'OV' in ct: return 'C'
    if 'NQ5' in ct:
        if 'FRT' in it:
            if 'STD' in it or 'STD' in det:
                return 'B2'
            else:
                return 'B'
        return 'I'
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
    for r in range(8, 138):
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
        d0 = [int(ws.cell(r,c).value) if isinstance(ws.cell(r,c).value,(int,float)) else 0 for c in [22,24,26,28,30,32,34,36,38,40]]
        d1 = [int(ws.cell(r,c).value) if isinstance(ws.cell(r,c).value,(int,float)) else 0 for c in [43,45,47,49,51,53,55,57,59,61]]
        d2 = [int(ws.cell(r,c).value) if isinstance(ws.cell(r,c).value,(int,float)) else 0 for c in [67,69,71,73,75,77,79,81,83,85]]
        items.append({
            'ct':ct,'it':it,'det':det,'clr':clr,'stk':stk,
            'd0':d0,'d0t':sum(d0),
            'd1':d1,'d1t':sum(d1),
            'd2':d2,'d2t':sum(d2),
            'grp':get_grp(ct,it,det),'cur':stk,
            'prod':[0]*10,
            'prod1':[0]*10,
            'prod2':[0]*10
        })
    wb.close()
    return items


# ============================================
# 핵심 스케줄링 함수
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
    """두 위치 배열 간 교체 수 계산"""
    if not pos1:
        return sum(1 for p in pos2 if p is not None) if pos2 else 0
    return sum(1 for i in range(HANGERS) if pos1[i] != pos2[i])


def get_optimal_order_for_colors(tmpl, grp_colors, prev_last_color):
    """컬러 교환 최소화하는 지그 순서 결정
    - 이전 마지막 컬러와 같은 컬러 먼저
    - 같은 컬러끼리 묶음
    """
    # 컬러별 지그그룹 묶기
    color_groups = defaultdict(list)
    for g, clr in grp_colors.items():
        if g in tmpl and tmpl[g] > 0:
            color_groups[clr].append(g)

    order = []
    used_colors = set()

    # 1. 이전 컬러와 같은 컬러 먼저
    if prev_last_color and prev_last_color in color_groups:
        # 행어 수 많은 순으로 정렬
        grps = sorted(color_groups[prev_last_color], key=lambda g: -tmpl.get(g, 0))
        order.extend(grps)
        used_colors.add(prev_last_color)

    # 2. 나머지 컬러들 (행어 합계 많은 순)
    remaining = [(clr, sum(tmpl.get(g, 0) for g in grps))
                 for clr, grps in color_groups.items() if clr not in used_colors]
    remaining.sort(key=lambda x: -x[1])

    for clr, _ in remaining:
        grps = sorted(color_groups[clr], key=lambda g: -tmpl.get(g, 0))
        order.extend(grps)

    return order


def calculate_template_for_demands(items, demand_arrays):
    """수요에 맞는 최적 템플릿 계산
    demand_arrays: list of (demand_key, weight) - 예: [('d0', 1.0), ('d1', 0.5)]
    """
    # 지그그룹별 가중 수요 계산
    grp_demand = defaultdict(float)
    for x in items:
        if x['grp']:
            for demand_key, weight in demand_arrays:
                grp_demand[x['grp']] += sum(x.get(demand_key, [0]*10)) * weight

    total_demand = sum(grp_demand.values())
    if total_demand == 0:
        return {}

    # 수요 비율로 행어 배분
    tmpl = {}
    remaining = HANGERS

    # 수요 많은 순으로 배분
    sorted_grps = sorted(grp_demand.keys(), key=lambda g: -grp_demand[g])

    for g in sorted_grps:
        if remaining <= 0:
            break
        max_h = JIG_INVENTORY[g]['max_jigs'] // JIGS_PER_HANGER
        ideal = int(HANGERS * grp_demand[g] / total_demand)
        alloc = min(max_h, max(1 if grp_demand[g] > 0 else 0, ideal), remaining)
        if alloc > 0:
            tmpl[g] = alloc
            remaining -= alloc

    # 남은 행어 배분 (pcs=2 우선)
    if remaining > 0:
        for g in sorted_grps:
            if remaining <= 0:
                break
            if grp_demand[g] > 0:
                max_h = JIG_INVENTORY[g]['max_jigs'] // JIGS_PER_HANGER
                current = tmpl.get(g, 0)
                add = min(remaining, max_h - current)
                if add > 0:
                    # pcs=2 지그그룹 우선
                    if JIG_INVENTORY[g]['pcs'] >= 2:
                        tmpl[g] = current + add
                        remaining -= add

    # 아직 남으면 아무 그룹에나
    if remaining > 0:
        for g in sorted_grps:
            if remaining <= 0:
                break
            max_h = JIG_INVENTORY[g]['max_jigs'] // JIGS_PER_HANGER
            current = tmpl.get(g, 0)
            add = min(remaining, max_h - current)
            if add > 0:
                tmpl[g] = current + add
                remaining -= add

    return tmpl


def try_adjust_template(base_tmpl, base_order, target_grp, delta, budget):
    """템플릿 조정 시도 (예산 내에서)
    target_grp에 delta만큼 행어 추가/제거
    Returns: (new_tmpl, new_order, changes_used) or None if not possible
    """
    new_tmpl = base_tmpl.copy()

    if delta > 0:
        # 행어 추가
        max_h = JIG_INVENTORY[target_grp]['max_jigs'] // JIGS_PER_HANGER
        current = new_tmpl.get(target_grp, 0)
        actual_add = min(delta, max_h - current)
        if actual_add <= 0:
            return None

        # 다른 그룹에서 빼기
        other_grps = [g for g in new_tmpl if g != target_grp and new_tmpl[g] > 1]
        other_grps.sort(key=lambda g: new_tmpl[g], reverse=True)

        removed = 0
        for g in other_grps:
            if removed >= actual_add:
                break
            can_remove = new_tmpl[g] - 1
            remove = min(can_remove, actual_add - removed)
            new_tmpl[g] -= remove
            removed += remove

        if removed < actual_add:
            return None

        new_tmpl[target_grp] = current + actual_add

    elif delta < 0:
        # 행어 제거
        current = new_tmpl.get(target_grp, 0)
        actual_remove = min(-delta, current - 1) if current > 1 else 0
        if actual_remove <= 0:
            return None

        new_tmpl[target_grp] = current - actual_remove

        # 다른 그룹에 추가
        other_grps = sorted(new_tmpl.keys(), key=lambda g: -new_tmpl.get(g, 0))
        added = 0
        for g in other_grps:
            if added >= actual_remove:
                break
            if g != target_grp:
                max_h = JIG_INVENTORY[g]['max_jigs'] // JIGS_PER_HANGER
                can_add = max_h - new_tmpl.get(g, 0)
                add = min(can_add, actual_remove - added)
                new_tmpl[g] = new_tmpl.get(g, 0) + add
                added += add

    # 순서 결정 (기존 순서 유지, 새 그룹은 뒤에)
    new_order = [g for g in base_order if g in new_tmpl and new_tmpl[g] > 0]
    for g in new_tmpl:
        if g not in new_order and new_tmpl[g] > 0:
            new_order.append(g)

    # 교체 수 계산
    old_pos = order_to_positions(base_tmpl, base_order)
    new_pos = order_to_positions(new_tmpl, new_order)
    changes = calc_position_changes(old_pos, new_pos)

    if changes > budget:
        return None

    return new_tmpl, new_order, changes


def schedule_day_v2(items, day_key, demand_key, prod_key, start_stock_key,
                    prev_template=None, prev_order=None, prev_color=None):
    """
    제약조건 기반 스케줄링 v3
    - 제약1: 지그교체 150/시프트 (적극 활용)
    - 제약2: 회전별 수요 충족
    - 목적: 컬러교환 최소화 (목표 ~30회/일)
    - 첫 회전 전환비용 = 0 (계산 안 함)
    """

    def get_grp_main_color_for_day(grp):
        """하루 전체에서 지그그룹의 주 컬러"""
        color_demand = defaultdict(int)
        for x in items:
            if x['grp'] == grp:
                color_demand[x['clr']] += sum(x.get(day_key, [0]*10))
        if color_demand:
            return max(color_demand.keys(), key=lambda c: color_demand[c])
        return None

    def get_grp_main_color_for_rotation(grp, rotation):
        """특정 회전에서 지그그룹의 주 컬러"""
        color_demand = defaultdict(int)
        for x in items:
            if x['grp'] == grp:
                color_demand[x['clr']] += x.get(day_key, [0]*10)[rotation]
        if color_demand:
            return max(color_demand.keys(), key=lambda c: color_demand[c])
        return None

    # 1. 기본 템플릿 계산 (전체 수요 기반)
    base_tmpl = calculate_template_for_demands(items, [(day_key, 1.0)])

    # 2. 컬러 정보 수집 및 그룹핑
    grp_colors = {}
    for g in base_tmpl:
        grp_colors[g] = get_grp_main_color_for_day(g)

    # 3. 컬러별로 지그그룹 분류
    color_to_grps = defaultdict(list)
    for g, clr in grp_colors.items():
        if clr and g in base_tmpl and base_tmpl[g] > 0:
            color_to_grps[clr].append(g)

    # 4. 최적 순서 결정 - 컬러 블록으로 그룹핑
    # 같은 컬러끼리 연속 배치하여 컬러교환 최소화
    def get_color_block_order(tmpl, grp_colors, prev_last_color=None):
        """컬러 블록 단위로 순서 결정"""
        # 컬러별 총 행어수 계산
        color_hangers = defaultdict(int)
        for g, clr in grp_colors.items():
            if g in tmpl and tmpl[g] > 0:
                color_hangers[clr] += tmpl[g]

        # 컬러 순서 결정 (이전 마지막 컬러 → 큰 컬러 순)
        sorted_colors = sorted(color_hangers.keys(), key=lambda c: -color_hangers[c])

        if prev_last_color and prev_last_color in sorted_colors:
            sorted_colors.remove(prev_last_color)
            sorted_colors.insert(0, prev_last_color)

        # 순서 생성
        order = []
        for clr in sorted_colors:
            # 해당 컬러의 지그그룹들 (행어 수 내림차순)
            clr_grps = [g for g in grp_colors if grp_colors[g] == clr and g in tmpl and tmpl[g] > 0]
            clr_grps.sort(key=lambda g: -tmpl[g])
            order.extend(clr_grps)

        return order

    # 시작 순서
    if prev_color:
        base_order = get_color_block_order(base_tmpl, grp_colors, prev_color)
    else:
        base_order = get_color_block_order(base_tmpl, grp_colors, None)

    # 결과 저장
    templates = []
    jig_orders = []
    jig_changes = [0] * 10

    # 시프트별 예산
    day_budget_left = JIG_BUDGET_DAY
    night_budget_left = JIG_BUDGET_NIGHT

    # 이전 상태 (첫 회전은 비용 0)
    prev_positions = None  # 첫 회전은 전환비용 계산 안 함

    for r in range(10):
        is_day_shift = r < 5
        budget_left = day_budget_left if is_day_shift else night_budget_left

        # 이 회전의 수요 기반 템플릿 계산
        rotation_demand = defaultdict(int)
        for x in items:
            if x['grp']:
                rotation_demand[x['grp']] += x.get(day_key, [0]*10)[r]

        # 수요가 있는 그룹만 포함
        active_grps = [g for g in base_tmpl if rotation_demand[g] > 0]

        if not active_grps:
            # 수요 없으면 기본 템플릿 사용
            curr_tmpl = base_tmpl.copy()
            curr_order = list(base_order)
        else:
            # 수요 비율로 템플릿 조정
            total_demand = sum(rotation_demand[g] for g in active_grps)
            curr_tmpl = {}
            remaining = HANGERS

            # 수요 순으로 배분
            sorted_grps = sorted(active_grps, key=lambda g: -rotation_demand[g])
            for g in sorted_grps:
                if remaining <= 0:
                    break
                max_h = JIG_INVENTORY[g]['max_jigs'] // JIGS_PER_HANGER
                ideal = max(1, int(HANGERS * rotation_demand[g] / total_demand))
                alloc = min(max_h, ideal, remaining)
                curr_tmpl[g] = alloc
                remaining -= alloc

            # 남은 행어 배분 (pcs=2 우선)
            if remaining > 0:
                for g in sorted_grps:
                    if remaining <= 0:
                        break
                    if JIG_INVENTORY[g]['pcs'] >= 2:
                        max_h = JIG_INVENTORY[g]['max_jigs'] // JIGS_PER_HANGER
                        add = min(remaining, max_h - curr_tmpl.get(g, 0))
                        if add > 0:
                            curr_tmpl[g] = curr_tmpl.get(g, 0) + add
                            remaining -= add

            # 아직 남으면 일반 배분
            if remaining > 0:
                for g in sorted_grps:
                    if remaining <= 0:
                        break
                    max_h = JIG_INVENTORY[g]['max_jigs'] // JIGS_PER_HANGER
                    add = min(remaining, max_h - curr_tmpl.get(g, 0))
                    if add > 0:
                        curr_tmpl[g] = curr_tmpl.get(g, 0) + add
                        remaining -= add

            # 하루 기준 컬러 정보 사용 (안정적 순서 유지)
            # 회전별 컬러가 아닌 하루 전체 주 컬러로 순서 결정
            rot_grp_colors = {}
            for g in curr_tmpl:
                rot_grp_colors[g] = grp_colors.get(g)  # 하루 기준 주 컬러

            # 이전 회전 마지막 컬러 확인
            if r > 0 and jig_orders:
                prev_order_list = jig_orders[-1]
                if prev_order_list:
                    last_g = prev_order_list[-1]
                    prev_last_color = grp_colors.get(last_g) if last_g else None
                else:
                    prev_last_color = prev_color
            else:
                prev_last_color = prev_color

            # 이전 회전 순서 최대한 유지 (안정성)
            if r > 0 and jig_orders:
                prev_order_list = jig_orders[-1]
                # 이전 순서에서 현재 템플릿에 있는 그룹만 유지
                curr_order = [g for g in prev_order_list if g in curr_tmpl and curr_tmpl[g] > 0]
                # 새로 추가된 그룹은 컬러 기준으로 적절한 위치에 삽입
                new_grps = [g for g in curr_tmpl if curr_tmpl[g] > 0 and g not in curr_order]
                for ng in new_grps:
                    ng_color = rot_grp_colors.get(ng)
                    # 같은 컬러 그룹 뒤에 삽입
                    inserted = False
                    for i, og in enumerate(curr_order):
                        if rot_grp_colors.get(og) == ng_color:
                            curr_order.insert(i+1, ng)
                            inserted = True
                            break
                    if not inserted:
                        curr_order.append(ng)
            else:
                # 첫 회전: 컬러 블록 순서
                curr_order = get_color_block_order(curr_tmpl, rot_grp_colors, prev_last_color)

        # 지그 교체 계산 (첫 회전은 0)
        curr_positions = order_to_positions(curr_tmpl, curr_order)

        if prev_positions is not None:
            changes = calc_position_changes(prev_positions, curr_positions)

            # 예산 체크
            if changes <= budget_left:
                jig_changes[r] = changes
                if is_day_shift:
                    day_budget_left -= changes
                else:
                    night_budget_left -= changes
            else:
                # 예산 초과시 이전 템플릿/순서 유지
                if templates:
                    curr_tmpl = templates[-1].copy()
                    curr_order = list(jig_orders[-1])
                    curr_positions = order_to_positions(curr_tmpl, curr_order)
                jig_changes[r] = 0
        else:
            # 첫 회전: 전환비용 0
            jig_changes[r] = 0

        templates.append(curr_tmpl)
        jig_orders.append(curr_order)
        prev_positions = curr_positions

    # 생산량 배정
    rotation_color_detail = [{} for _ in range(10)]

    for r in range(10):
        tmpl = templates[r]
        order = jig_orders[r]

        for g in order:
            if g not in tmpl or tmpl[g] == 0:
                continue

            h = tmpl[g]
            cap = h * JIGS_PER_HANGER * JIG_INVENTORY[g]['pcs']

            # 이 지그그룹의 아이템들
            grp_items = [x for x in items if x['grp'] == g]
            if not grp_items:
                continue

            # 수요 비율로 용량 배분
            grp_total = sum(max(1, x[day_key][r]) for x in grp_items)

            clr_hangers = defaultdict(int)
            for x in grp_items:
                ratio = max(1, x[day_key][r]) / grp_total if grp_total > 0 else 1/len(grp_items)
                prod = int(cap * ratio)
                x[prod_key][r] += prod
                clr_hangers[x['clr']] += int(h * ratio)

            rotation_color_detail[r][g] = dict(clr_hangers)

    # 부족분 보정
    for x in items:
        stk = x.get(start_stock_key, x['stk'])
        for r in range(10):
            stk = stk - x[day_key][r] + x[prod_key][r]
            if stk < 0:
                deficit = -stk
                # 뒤에서부터 추가 생산
                for pr in range(r, -1, -1):
                    if x[prod_key][pr] > 0:
                        x[prod_key][pr] += deficit
                        stk += deficit
                        break
                else:
                    x[prod_key][r] += deficit
                    stk += deficit

    # 컬러 교환 계산
    def get_grp_main_color(grp, rotation):
        """해당 회전에서 지그그룹의 주 생산 컬러"""
        color_demand = defaultdict(int)
        for x in items:
            if x['grp'] == grp:
                color_demand[x['clr']] += x[day_key][rotation]
        if color_demand:
            return max(color_demand.keys(), key=lambda c: color_demand[c])
        return None

    def count_color_changes():
        total_cc = 0
        cc_per_rotation = [0] * 10
        p_color = prev_color  # 외부 변수 참조

        for r in range(10):
            colors_in_order = []
            for g in jig_orders[r]:
                if g in templates[r] and templates[r][g] > 0:
                    clr = get_grp_main_color(g, r)
                    if clr:
                        colors_in_order.append(clr)

            if not colors_in_order:
                continue

            # 회전 간 교환
            if p_color and colors_in_order[0] != p_color:
                cc_per_rotation[r] += 1

            # 회전 내 교환
            for i in range(1, len(colors_in_order)):
                if colors_in_order[i] != colors_in_order[i-1]:
                    cc_per_rotation[r] += 1

            total_cc += cc_per_rotation[r]
            p_color = colors_in_order[-1] if colors_in_order else p_color

        return total_cc, cc_per_rotation

    cc, cc_per_rotation = count_color_changes()

    # 회전별 주 컬러
    rotation_color = []
    for r in range(10):
        colors = defaultdict(int)
        for g in jig_orders[r]:
            if g in templates[r] and templates[r][g] > 0:
                clr = get_grp_main_color(g, r)
                if clr:
                    colors[clr] += templates[r][g]
        rotation_color.append(max(colors.keys(), key=lambda c: colors[c]) if colors else None)

    last_order = jig_orders[9]
    last_color = rotation_color[9]

    return (templates, rotation_color, jig_changes, cc, cc_per_rotation,
            rotation_color_detail, jig_orders, last_order, last_color)


def schedule(items):
    """D0, D+1, D+2 전체 스케줄링"""

    # D0 스케줄링
    (templates_d0, rotation_color_d0, jig_changes_d0, cc_d0, cc_per_rot_d0,
     color_detail_d0, jig_orders_d0, d0_last_order, d0_last_color) = schedule_day_v2(
        items, 'd0', 'd0', 'prod', 'stk',
        prev_template=None, prev_order=None, prev_color=None)

    # D+1 부족분 선반영
    for x in items:
        d0_end = x['stk']
        for r in range(10):
            d0_end = d0_end - x['d0'][r] + x['prod'][r]

        running = d0_end
        d1_deficit = 0
        for r in range(10):
            running = running - x['d1'][r]
            if running < 0:
                d1_deficit = max(d1_deficit, -running)

        if d1_deficit > 0:
            for pr in range(9, -1, -1):
                if x['prod'][pr] > 0:
                    x['prod'][pr] += d1_deficit
                    break
            else:
                x['prod'][9] += d1_deficit

    # D0 기말재고 계산
    for x in items:
        stk = x['stk']
        for r in range(10):
            stk = stk - x['d0'][r] + x['prod'][r]
        x['cur'] = stk

    # D0→D+1 전환 지그교체
    d0_last_tmpl = templates_d0[9]
    d0_last_pos = order_to_positions(d0_last_tmpl, d0_last_order)

    # D+1 스케줄링
    (templates_d1, rotation_color_d1, jig_changes_d1, cc_d1, cc_per_rot_d1,
     color_detail_d1, jig_orders_d1, d1_last_order, d1_last_color) = schedule_day_v2(
        items, 'd1', 'd1', 'prod1', 'cur',
        prev_template=d0_last_tmpl, prev_order=d0_last_order, prev_color=d0_last_color)

    d1_first_pos = order_to_positions(templates_d1[0], jig_orders_d1[0])
    d0_to_d1_jig = calc_position_changes(d0_last_pos, d1_first_pos)

    # D+1 기말재고
    for x in items:
        stk = x['cur']
        for r in range(10):
            stk = stk - x['d1'][r] + x['prod1'][r]
        x['cur1'] = stk

    # D+2 스케줄링
    d1_last_tmpl = templates_d1[9]
    d1_last_pos = order_to_positions(d1_last_tmpl, d1_last_order)

    (templates_d2, rotation_color_d2, jig_changes_d2, cc_d2, cc_per_rot_d2,
     color_detail_d2, jig_orders_d2, d2_last_order, d2_last_color) = schedule_day_v2(
        items, 'd2', 'd2', 'prod2', 'cur1',
        prev_template=d1_last_tmpl, prev_order=d1_last_order, prev_color=d1_last_color)

    d2_first_pos = order_to_positions(templates_d2[0], jig_orders_d2[0])
    d1_to_d2_jig = calc_position_changes(d1_last_pos, d2_first_pos)

    # D+2 기말재고
    for x in items:
        stk = x['cur1']
        for r in range(10):
            stk = stk - x['d2'][r] + x['prod2'][r]
        x['cur2'] = stk

    return {
        'd0': {
            'templates': templates_d0,
            'colors': rotation_color_d0,
            'jig_changes': jig_changes_d0,
            'cc': cc_d0,
            'cc_per_rotation': cc_per_rot_d0,
            'color_detail': color_detail_d0,
            'jig_orders': jig_orders_d0,
        },
        'd1': {
            'templates': templates_d1,
            'colors': rotation_color_d1,
            'jig_changes': jig_changes_d1,
            'cc': cc_d1,
            'cc_per_rotation': cc_per_rot_d1,
            'color_detail': color_detail_d1,
            'jig_orders': jig_orders_d1,
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


# ============================================
# 리포트 생성 함수
# ============================================

def get_rotation_items_detail(items, rotation, prod_key, templates, jig_orders):
    """회전별 생산 아이템 상세"""
    tmpl = templates[rotation]
    order = jig_orders[rotation] if jig_orders and jig_orders[rotation] else sorted(tmpl.keys())

    result = []
    for g in order:
        if g not in tmpl or tmpl[g] == 0:
            continue
        grp_items = [(x, x[prod_key][rotation]) for x in items
                     if x['grp'] == g and x[prod_key][rotation] > 0]
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

    grp_colors = {
        'A': '#E3F2FD', 'B': '#E8F5E9', 'B2': '#C8E6C9', 'C': '#FFF3E0', 'D': '#FCE4EC',
        'E': '#F3E5F5', 'F': '#E0F7FA', 'G': '#FFF8E1', 'H': '#EFEBE9', 'I': '#ECEFF1'
    }
    grp_border = {
        'A': '#1976D2', 'B': '#388E3C', 'B2': '#2E7D32', 'C': '#F57C00', 'D': '#C2185B',
        'E': '#7B1FA2', 'F': '#0097A7', 'G': '#FFA000', 'H': '#5D4037', 'I': '#455A64'
    }

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

        det_str = f" {det}" if det and det != '-' else ""
        grp_items.append(f"<b>{ct}</b> {it}{det_str} <span style='color:#1565C0;'>{clr}</span> <span style='color:#D32F2F;font-weight:bold;'>{prod}</span>")

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


def format_hanger_positions_html(templates, jig_orders, rotation, prev_positions=None):
    """140개 행어 위치를 시각적으로 표시"""
    tmpl = templates[rotation]
    order = jig_orders[rotation] if jig_orders and jig_orders[rotation] else sorted(tmpl.keys())

    curr_positions = order_to_positions(tmpl, order)

    grp_colors = {
        'A': '#1976D2', 'B': '#388E3C', 'B2': '#2E7D32', 'C': '#F57C00', 'D': '#C2185B',
        'E': '#7B1FA2', 'F': '#0097A7', 'G': '#FFA000', 'H': '#5D4037', 'I': '#455A64'
    }

    segments = []
    if curr_positions:
        start = 0
        current_grp = curr_positions[0]
        for i in range(1, len(curr_positions)):
            if curr_positions[i] != current_grp:
                segments.append({
                    'grp': current_grp,
                    'start': start,
                    'end': i - 1,
                    'count': i - start
                })
                start = i
                current_grp = curr_positions[i]
        segments.append({
            'grp': current_grp,
            'start': start,
            'end': len(curr_positions) - 1,
            'count': len(curr_positions) - start
        })

    change_positions = set()
    if prev_positions:
        for i in range(HANGERS):
            if i < len(prev_positions) and i < len(curr_positions):
                if prev_positions[i] != curr_positions[i]:
                    change_positions.add(i)

    html_parts = []
    html_parts.append('<div style="display:flex;align-items:center;gap:10px;margin:5px 0;">')
    html_parts.append('<span style="font-size:0.75em;color:#666;min-width:60px;">행어위치:</span>')
    html_parts.append('<div style="display:flex;flex-wrap:nowrap;border:1px solid #999;border-radius:4px;overflow:hidden;flex:1;">')

    for seg in segments:
        grp = seg['grp']
        count = seg['count']
        start = seg['start']
        end = seg['end']

        if grp is None:
            color = '#E0E0E0'
            label = '-'
        else:
            color = grp_colors.get(grp, '#9E9E9E')
            label = grp

        has_change = any(pos in change_positions for pos in range(start, end + 1))
        border_style = 'border-left:3px solid #F44336;' if has_change and start > 0 else ''

        width_pct = (count / HANGERS) * 100

        html_parts.append(
            f'<div style="width:{width_pct:.1f}%;background:{color};color:white;text-align:center;'
            f'font-size:0.7em;padding:2px 0;min-width:15px;{border_style}" '
            f'title="위치 {start+1}-{end+1} ({count}행어)">'
            f'{label}<span style="font-size:0.8em;opacity:0.8;">({count})</span></div>'
        )

    html_parts.append('</div>')

    if change_positions:
        html_parts.append(f'<span style="font-size:0.7em;color:#D32F2F;margin-left:5px;">교체:{len(change_positions)}개</span>')

    html_parts.append('</div>')

    return ''.join(html_parts), curr_positions


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
        .summary-item.over-budget {{
            background: linear-gradient(135deg, #ff416c 0%, #ff4b2b 100%);
        }}
        .summary-number {{ font-size: 2em; font-weight: bold; }}
        .summary-label {{ font-size: 0.9em; opacity: 0.9; }}
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
    </style>
</head>
<body>
    <div class="container">
        <h1>D0 생산계획 리포트 v8.16</h1>
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

    day_jig_class = 'over-budget' if day_jig > JIG_BUDGET_DAY else 'success'
    night_jig_class = 'over-budget' if night_jig > JIG_BUDGET_NIGHT else 'success'

    html += f'''
            </table>
        </div>

        <div class="card">
            <h2>실행 결과 요약</h2>
            <div class="summary-box">
                <div class="summary-item success">
                    <div class="summary-number">{total_prod_d0:,}</div>
                    <div class="summary-label">D0 생산량</div>
                </div>
                <div class="summary-item success">
                    <div class="summary-number">{total_prod_d1:,}</div>
                    <div class="summary-label">D+1 생산량</div>
                </div>
                <div class="summary-item success">
                    <div class="summary-number">{total_prod_d2:,}</div>
                    <div class="summary-label">D+2 생산량</div>
                </div>
                <div class="summary-item">
                    <div class="summary-number">{cc}</div>
                    <div class="summary-label">D0 컬러교환</div>
                </div>
                <div class="summary-item warning">
                    <div class="summary-number">{cc * COLOR_CHANGE_LOSS}</div>
                    <div class="summary-label">D0 컬러손실</div>
                </div>
                <div class="summary-item {day_jig_class}">
                    <div class="summary-number">{day_jig}/{JIG_BUDGET_DAY}</div>
                    <div class="summary-label">D0 주간지그</div>
                </div>
                <div class="summary-item {night_jig_class}">
                    <div class="summary-number">{night_jig}/{JIG_BUDGET_NIGHT}</div>
                    <div class="summary-label">D0 야간지그</div>
                </div>
            </div>
        </div>
'''

    # 재고/수요/생산 합계
    init_stock = sum(x['stk'] for x in items)
    d0_demand = sum(x['d0t'] for x in items)
    d0_prod = total_prod_d0
    d0_end_stock = sum(x['cur'] for x in items)
    d1_demand = sum(x['d1t'] for x in items)
    d1_prod = total_prod_d1
    d1_end_stock = sum(x['cur1'] for x in items)
    d2_demand = sum(x['d2t'] for x in items)
    d2_prod = total_prod_d2
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
            </table>
        </div>
'''

    # 회전별 생산 상세
    templates_d1 = d1['templates']
    jig_changes_d1 = d1['jig_changes']
    templates_d2_box = d2['templates']
    jig_changes_d2_box = d2['jig_changes']

    html += '''
        <div class="card">
            <h2>회전별 생산 상세 (지그 순서)</h2>
            <p>각 회전에서 생산되는 아이템을 컨베이어 지그 순서대로 표시 (박스 = 지그그룹, 숫자 = 행어수)</p>
'''

    # 지그그룹 색상 범례
    grp_colors_legend = {
        'A': '#1976D2', 'B': '#388E3C', 'B2': '#2E7D32', 'C': '#F57C00', 'D': '#C2185B',
        'E': '#7B1FA2', 'F': '#0097A7', 'G': '#FFA000', 'H': '#5D4037', 'I': '#455A64'
    }
    legend_html = '<div style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:10px;font-size:0.75em;">'
    for g, color in grp_colors_legend.items():
        legend_html += f'<span style="background:{color};color:white;padding:2px 6px;border-radius:3px;">{g}</span>'
    legend_html += '<span style="border-left:3px solid #F44336;padding-left:5px;margin-left:10px;">= 지그교체</span></div>'

    # D0
    html += '<h3 style="margin-top:20px;color:#1565C0;">D0 생산계획</h3>'
    html += legend_html

    prev_positions_d0 = None
    for r in range(10):
        shift_name = '주간' if r < 5 else '야간'
        shift_bg = '#E3F2FD' if r < 5 else '#E8EAF6'
        detail_html = format_rotation_items_html(items, r, 'prod', templates, jig_orders_d0)
        hanger_html, curr_positions = format_hanger_positions_html(templates, jig_orders_d0, r, prev_positions_d0)
        prev_positions_d0 = curr_positions

        html += f'''
            <div style="margin:8px 0;padding:10px;background:{shift_bg};border-radius:8px;">
                <div style="display:flex;align-items:center;gap:15px;margin-bottom:8px;">
                    <span style="font-weight:bold;font-size:1.1em;color:#1565C0;">D0-{r+1}</span>
                    <span style="background:#1976D2;color:white;padding:2px 8px;border-radius:4px;font-size:0.8em;">{shift_name}</span>
                    <span style="color:#666;font-size:0.85em;">지그교체: <b>{jig_changes[r]}</b></span>
                </div>
                {hanger_html}
                <div style="display:flex;flex-wrap:wrap;gap:4px;margin-top:5px;">{detail_html}</div>
            </div>'''

    # D0→D+1 전환
    html += f'''
            <div style="margin:15px 0;padding:10px;background:#FFE082;border-radius:8px;text-align:center;">
                <strong>▶ D0→D+1 전환</strong> | 지그교체: <b>{d1.get('start_jig_change', 0)}</b>
            </div>'''

    # D+1
    html += '<h3 style="margin-top:20px;color:#2E7D32;">D+1 생산계획</h3>'
    prev_positions_d1 = prev_positions_d0
    for r in range(10):
        shift_name = '주간' if r < 5 else '야간'
        shift_bg = '#E8F5E9' if r < 5 else '#F1F8E9'
        detail_html = format_rotation_items_html(items, r, 'prod1', templates_d1, jig_orders_d1)
        hanger_html, curr_positions = format_hanger_positions_html(templates_d1, jig_orders_d1, r, prev_positions_d1)
        prev_positions_d1 = curr_positions

        html += f'''
            <div style="margin:8px 0;padding:10px;background:{shift_bg};border-radius:8px;">
                <div style="display:flex;align-items:center;gap:15px;margin-bottom:8px;">
                    <span style="font-weight:bold;font-size:1.1em;color:#2E7D32;">D+1-{r+1}</span>
                    <span style="background:#388E3C;color:white;padding:2px 8px;border-radius:4px;font-size:0.8em;">{shift_name}</span>
                    <span style="color:#666;font-size:0.85em;">지그교체: <b>{jig_changes_d1[r]}</b></span>
                </div>
                {hanger_html}
                <div style="display:flex;flex-wrap:wrap;gap:4px;margin-top:5px;">{detail_html}</div>
            </div>'''

    # D+1→D+2 전환
    html += f'''
            <div style="margin:15px 0;padding:10px;background:#FFE082;border-radius:8px;text-align:center;">
                <strong>▶ D+1→D+2 전환</strong> | 지그교체: <b>{d2.get('start_jig_change', 0)}</b>
            </div>'''

    # D+2
    html += '<h3 style="margin-top:20px;color:#E65100;">D+2 생산계획</h3>'
    prev_positions_d2 = prev_positions_d1
    for r in range(10):
        shift_name = '주간' if r < 5 else '야간'
        shift_bg = '#FFF3E0' if r < 5 else '#FBE9E7'
        detail_html = format_rotation_items_html(items, r, 'prod2', templates_d2_box, jig_orders_d2)
        hanger_html, curr_positions = format_hanger_positions_html(templates_d2_box, jig_orders_d2, r, prev_positions_d2)
        prev_positions_d2 = curr_positions

        html += f'''
            <div style="margin:8px 0;padding:10px;background:{shift_bg};border-radius:8px;">
                <div style="display:flex;align-items:center;gap:15px;margin-bottom:8px;">
                    <span style="font-weight:bold;font-size:1.1em;color:#E65100;">D+2-{r+1}</span>
                    <span style="background:#F57C00;color:white;padding:2px 8px;border-radius:4px;font-size:0.8em;">{shift_name}</span>
                    <span style="color:#666;font-size:0.85em;">지그교체: <b>{jig_changes_d2_box[r]}</b></span>
                </div>
                {hanger_html}
                <div style="display:flex;flex-wrap:wrap;gap:4px;margin-top:5px;">{detail_html}</div>
            </div>'''

    html += '''
        </div>
    </div>
</body>
</html>'''

    return html


def save_csv(items, filename='production_plan_v8.csv'):
    """CSV 저장"""
    import csv
    with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        header = ['차종', '아이템', '세부', '컬러', '지그', '기초재고']
        for i in range(10):
            header.extend([f'D0_{i+1}수요', f'D0_{i+1}생산'])
        header.append('D0기말')
        for i in range(10):
            header.extend([f'D1_{i+1}수요', f'D1_{i+1}생산'])
        header.append('D1기말')
        for i in range(10):
            header.extend([f'D2_{i+1}수요', f'D2_{i+1}생산'])
        header.append('D2기말')
        writer.writerow(header)

        for x in items:
            row = [x['ct'], x['it'], x['det'], x['clr'], x['grp'], x['stk']]
            for i in range(10):
                row.extend([x['d0'][i], x['prod'][i]])
            row.append(x['cur'])
            for i in range(10):
                row.extend([x['d1'][i], x['prod1'][i]])
            row.append(x['cur1'])
            for i in range(10):
                row.extend([x['d2'][i], x['prod2'][i]])
            row.append(x['cur2'])
            writer.writerow(row)


if __name__ == '__main__':
    print("데이터 로드...")
    items = load_data()
    print(f"{len(items)}개 아이템")

    print("스케줄링...")
    result = schedule(items)

    print("HTML 리포트 생성...")
    html = generate_html_report(items, result)
    with open('production_report.html', 'w', encoding='utf-8') as f:
        f.write(html)

    print("CSV 저장...")
    save_csv(items)

    print("\n완료!")
    print("  => production_report.html")
    print("  => production_plan_v8.csv")
