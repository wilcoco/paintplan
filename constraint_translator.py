#!/usr/bin/env python3
"""
자연어 제약 → MIP 제약 코드 번역 + 검증 (C 방식)
====================================================
웹 입력창의 자연어 지시를 Claude API로 OR-Tools 제약 함수로 번역하고,
AST 화이트리스트로 안전성을 검증한 뒤, 사람 승인 후 저장한다.

생성 코드는 항상 다음 형태:
    def constraint(ctx):
        ...
        ctx.add(<선형식> <= / >= / == <선형식>)

ctx 객체가 노출하는 인터페이스는 SYSTEM_PROMPT의 변수 스펙과 동일하게 유지해야 한다.
스케줄러(schedulers.py)의 MIPContext와 1:1로 맞출 것.
"""
import ast
import json
import os

CONSTRAINTS_FILE = os.path.join(os.path.dirname(__file__), 'custom_constraints.json')

# ============================================
# LLM 시스템 프롬프트 — ctx 인터페이스 스펙
# ============================================
SYSTEM_PROMPT = """\
너는 자동차 범퍼 도장 컨베이어 생산계획 MIP(혼합정수계획) 모델에 제약을 추가하는 코드 생성기다.
사용자의 한국어 지시를 받아, OR-Tools 제약을 추가하는 파이썬 함수 하나만 생성한다.

반드시 다음 형태의 함수 **하나만** 출력한다 (설명/마크다운/주석 금지, 코드만):

def constraint(ctx):
    ...

함수 안에서 사용할 수 있는 것은 ctx 객체와 그 메서드/속성, 그리고 sum, range, len, min, max, abs 내장함수뿐이다.
import, 파일접근, while문, 람다, 밑줄(_)로 시작하는 이름은 절대 사용하지 마라.

ctx 인터페이스:
- ctx.n_rotations        : 전체 회전 수. 1일 모델=10, 2일 모델=20. 항상 range(ctx.n_rotations)로 전체를 순회하라.
- ctx.rotations_per_day  : 하루 회전 수 (10)
- ctx.n_days             : 날짜 수 (1 또는 2). r=0~9가 D0, r=10~19가 D+1
- ctx.is_night(r)        : 회전 r이 야간(6~10회전)이면 True. **시프트 기준 제약은 r<5 같은 하드코딩 대신 반드시 이 함수를 써라** (1일/2일 모델 모두에서 올바르게 동작)
- ctx.is_day_shift(r)    : 회전 r이 주간(1~5회전)이면 True
- ctx.day_of(r)          : 회전 r이 속한 날짜 (0=D0, 1=D+1)
- ctx.groups             : 지그 그룹 코드 리스트. ['A','B','B2','C','D','E','F','G','H','I'] 중 데이터에 존재하는 것들
- ctx.colors             : 컬러 코드 문자열 리스트
- ctx.items              : 아이템 dict 리스트. 각 item은 'grp'(그룹), 'clr'(컬러) 키를 가질 수 있음
- ctx.n_items            : 아이템 수
- ctx.SPECIAL_COLORS     : 특수컬러 set {'MGG','T4M','UMA','ZRM','ISM','MRM'}

변수 접근 (모두 ctx 메서드, 정수/이진 선형식 반환):
- ctx.prod(i, r)         : 아이템 i를 회전 r에 생산하는 수량 (정수변수)
- ctx.hangers(g, r)      : 그룹 g가 회전 r에서 쓰는 행어 수 (정수변수). g가 없으면 0
- ctx.uses_color(c, r)   : 컬러 c가 회전 r에서 쓰이면 1, 아니면 0 (이진변수). c가 없으면 0
- ctx.color_start(c, r)  : 컬러 c가 회전 r에서 새로 시작하면 1 (이진변수)
- ctx.group_prod(g, r)   : 그룹 g의 회전 r 총생산량 = 그룹 g 아이템들의 prod 합
- ctx.color_prod(c, r)   : 컬러 c의 회전 r 총생산량

제약 추가:
- ctx.add(expr)          : 선형 제약을 추가. 예: ctx.add(ctx.hangers('A', r) + ctx.hangers('H', r) <= 1)

목적함수 가중치 조절 (우선순위 변경 지시일 때 사용):
- ctx.set_weight('cc_weight', 값)        : 컬러교환 페널티 (기본 1000, 클수록 컬러교환을 강하게 억제)
- ctx.set_weight('production_weight', 값): 생산량 보상 (기본 1, 클수록 생산량 최대화를 우선)
- ctx.set_weight('empty_weight', 값)     : 빈행어 페널티 (기본 100, 1일 모델에만 영향)
  "A보다 B를 우선해" 같은 지시는 두 가중치의 상대크기로 표현하라.

추가로 사용할 수 있는 보조:
- ctx.group_active(g, r) : 그룹 g가 회전 r에서 행어를 1개 이상 쓰면 1 (이진변수)

예시 1) "A그룹과 H그룹은 같은 회전에 같이 쓰지 마라":
def constraint(ctx):
    for r in range(ctx.n_rotations):
        ctx.add(ctx.group_active('A', r) + ctx.group_active('H', r) <= 1)

예시 2) "특수컬러는 야간에만 배치" (주간 회전에서 사용 금지):
def constraint(ctx):
    for c in ctx.colors:
        if c.upper() in ctx.SPECIAL_COLORS:
            for r in range(ctx.n_rotations):
                if ctx.is_day_shift(r):
                    ctx.add(ctx.uses_color(c, r) == 0)

예시 3) "OV1(C그룹)는 하루에 최대 4회전까지만 사용":
def constraint(ctx):
    for d in range(ctx.n_days):
        base = d * ctx.rotations_per_day
        ctx.add(sum(ctx.group_active('C', base + r) for r in range(ctx.rotations_per_day)) <= 4)

예시 4) "컬러교환보다 생산량을 더 우선해" (목적함수 가중치 조절):
def constraint(ctx):
    ctx.set_weight('cc_weight', 200)
    ctx.set_weight('production_weight', 5)

지시가 모델 변수로 표현 불가능하거나 모호하면, 함수 본문 첫 줄에
    ctx.reject("이유")
만 호출하고 끝내라.

이제 사용자 지시에 맞는 constraint 함수 하나만 코드로 출력하라."""


# ============================================
# AST 화이트리스트 검증
# ============================================
ALLOWED_NODES = {
    ast.Module, ast.FunctionDef, ast.arguments, ast.arg,
    ast.For, ast.If, ast.Expr, ast.Assign, ast.AugAssign, ast.Return, ast.Pass,
    ast.Call, ast.Attribute, ast.Subscript, ast.Index,
    ast.BinOp, ast.UnaryOp, ast.BoolOp, ast.Compare,
    ast.Name, ast.Load, ast.Store, ast.Constant,
    ast.Tuple, ast.List, ast.ListComp, ast.GeneratorExp, ast.comprehension,
    # 연산자
    ast.Add, ast.Sub, ast.Mult, ast.Div, ast.FloorDiv, ast.Mod, ast.Pow,
    ast.USub, ast.UAdd, ast.Not,
    ast.And, ast.Or,
    ast.Eq, ast.NotEq, ast.Lt, ast.LtE, ast.Gt, ast.GtE,
    ast.In, ast.NotIn,
}
# 파이썬 버전별로 ast.Index가 없을 수 있음
ALLOWED_NODES = {n for n in ALLOWED_NODES if n is not None}

ALLOWED_BUILTINS = {'sum', 'range', 'len', 'min', 'max', 'abs', 'enumerate'}

# ctx 이외의 객체(주로 컬러/그룹 문자열)에 허용되는 안전한 메서드
SAFE_METHODS = {'upper', 'lower', 'strip', 'startswith', 'endswith'}


class ValidationError(Exception):
    pass


def validate_ast(code):
    """생성 코드를 AST 화이트리스트로 검증. 통과하면 None, 실패하면 ValidationError 발생."""
    try:
        tree = ast.parse(code)
    except SyntaxError as e:
        raise ValidationError(f'문법 오류: {e}')

    # 최상위는 def constraint(ctx) 하나뿐이어야 함
    if (len(tree.body) != 1 or not isinstance(tree.body[0], ast.FunctionDef)):
        raise ValidationError('최상위에 함수 정의 하나만 허용됩니다.')
    fn = tree.body[0]
    if fn.name != 'constraint':
        raise ValidationError("함수 이름은 'constraint' 여야 합니다.")
    args = fn.args
    if (len(args.args) != 1 or args.args[0].arg != 'ctx' or
            args.vararg or args.kwarg or args.kwonlyargs or args.defaults):
        raise ValidationError("함수 시그니처는 constraint(ctx) 여야 합니다.")

    for node in ast.walk(tree):
        if type(node) not in ALLOWED_NODES:
            raise ValidationError(f'허용되지 않은 구문: {type(node).__name__}')

        # 밑줄(_) 시작 이름 / 속성 차단 (던더, 내부 접근 방지)
        if isinstance(node, ast.Name) and node.id.startswith('_'):
            raise ValidationError(f"허용되지 않은 이름: {node.id}")
        if isinstance(node, ast.Attribute):
            if node.attr.startswith('_'):
                raise ValidationError(f"허용되지 않은 속성 접근: .{node.attr}")
            # 속성 접근은 ctx 변수에서만 (ctx.xxx). 체인은 허용하되 베이스가 결국 Name이어야 함
            base = node.value
            while isinstance(base, (ast.Attribute, ast.Subscript, ast.Call)):
                base = base.value if isinstance(base, ast.Attribute) else (
                    base.func if isinstance(base, ast.Call) else base.value)
            if not isinstance(base, ast.Name):
                raise ValidationError('속성 접근의 베이스가 변수여야 합니다.')

        # 호출 검증: ctx.* 메서드 또는 허용 내장함수만
        if isinstance(node, ast.Call):
            func = node.func
            if isinstance(func, ast.Attribute):
                # ctx.add(...), ctx.prod(...) 등 — 베이스가 ctx 인지 확인
                is_ctx = isinstance(func.value, ast.Name) and func.value.id == 'ctx'
                # 문자열 안전 메서드(c.upper() 등)는 베이스가 ctx가 아니어도 허용
                if not is_ctx and func.attr not in SAFE_METHODS:
                    raise ValidationError(
                        f'메서드 호출은 ctx.* 또는 {sorted(SAFE_METHODS)} 만 허용됩니다.')
            elif isinstance(func, ast.Name):
                if func.id not in ALLOWED_BUILTINS:
                    raise ValidationError(f'허용되지 않은 함수 호출: {func.id}')
            else:
                raise ValidationError('허용되지 않은 호출 형태입니다.')

    return None


# ============================================
# Claude API 번역
# ============================================
def translate(nl_text, model=None):
    """자연어 지시 → constraint 함수 소스코드. (API 키 필요)

    반환: 코드 문자열 (def constraint(ctx): ...)
    예외: RuntimeError (API 키 없음/SDK 없음/응답 파싱 실패)
    """
    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        raise RuntimeError('ANTHROPIC_API_KEY 환경변수가 설정되지 않았습니다.')
    try:
        import anthropic
    except ImportError:
        raise RuntimeError("anthropic SDK가 없습니다. 'pip install anthropic' 필요.")

    model = model or os.environ.get('CONSTRAINT_MODEL', 'claude-sonnet-4-6')
    client = anthropic.Anthropic(api_key=api_key)
    resp = client.messages.create(
        model=model,
        max_tokens=1500,
        system=SYSTEM_PROMPT,
        messages=[{'role': 'user', 'content': nl_text}],
    )
    text = ''.join(block.text for block in resp.content if block.type == 'text').strip()
    return _strip_code_fence(text)


def _strip_code_fence(text):
    """```python ... ``` 펜스 제거."""
    t = text.strip()
    if t.startswith('```'):
        lines = t.splitlines()
        # 첫 줄(```python)과 마지막 ``` 제거
        if lines[0].startswith('```'):
            lines = lines[1:]
        if lines and lines[-1].strip().startswith('```'):
            lines = lines[:-1]
        t = '\n'.join(lines).strip()
    return t


# ============================================
# 저장소 (custom_constraints.json)
# ============================================
def load_constraints():
    if not os.path.exists(CONSTRAINTS_FILE):
        return []
    try:
        with open(CONSTRAINTS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return []


def save_constraints(constraints):
    with open(CONSTRAINTS_FILE, 'w', encoding='utf-8') as f:
        json.dump(constraints, f, ensure_ascii=False, indent=2)


# 저장 전 실행가능성 테스트용 임시 제약 (스케줄러가 1회 주입)
_PREVIEW = None


def set_preview(code):
    global _PREVIEW
    _PREVIEW = code


def clear_preview():
    global _PREVIEW
    _PREVIEW = None


def active_constraints():
    """활성(enabled) 제약의 코드 리스트 반환 — 스케줄러가 주입할 대상.
    미리보기 제약이 설정돼 있으면 마지막에 포함(실행가능성 테스트용)."""
    items = [c for c in load_constraints() if c.get('enabled', True)]
    if _PREVIEW:
        items = items + [{'code': _PREVIEW, 'text': '[미리보기]', 'enabled': True}]
    return items


def compile_constraint(code):
    """검증된 코드를 컴파일해 constraint 함수 객체 반환.
    호출 전 반드시 validate_ast 통과시킬 것."""
    namespace = {}
    # 내장함수 제한: 화이트리스트만 제공
    safe_builtins = {name: __builtins__[name] if isinstance(__builtins__, dict)
                     else getattr(__builtins__, name)
                     for name in ALLOWED_BUILTINS}
    exec(compile(code, '<constraint>', 'exec'), {'__builtins__': safe_builtins}, namespace)
    fn = namespace.get('constraint')
    if not callable(fn):
        raise ValidationError('constraint 함수를 찾을 수 없습니다.')
    return fn
