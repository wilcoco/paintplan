"""
도장 생산계획 시스템 — 시스템 변수
"""

# === 컨베이어 설정 ===
HANGER_COUNT = 140           # 컨베이어 행어 수
JIGS_PER_HANGER = 2          # 행어당 지그 슬롯
ROTATIONS_PER_DAY = 10       # 일일 회전 수

# 일일 최대 생산 용량 (지그 슬롯 기준)
DAILY_CAPACITY = HANGER_COUNT * JIGS_PER_HANGER * ROTATIONS_PER_DAY  # 2,800

# === 지그 교환 제약 ===
MAX_JIG_CHANGES_PER_DAY = 280  # 지그 교환 한도 (지그 단위)

# === 안전재고 ===
SAFETY_STOCK_DAYS = 3        # 안전재고 일수 (향후 N 생산일 수요)

# === 계획 기간 ===
PLANNING_DAYS = 12           # 생산계획 일수

# === 도장 배치 설정 ===
MIN_BATCH_SIZE = 10          # 최소 배치 크기 (이하면 합산 고려)
