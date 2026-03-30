# 도장 생산계획 시스템 v8.11

자동차 범퍼 도장 공정의 생산계획 최적화 시스템

## 주요 기능

- **위치 기반 지그교체 계산**: 140행어 각 위치별로 이전 회전과 비교하여 실제 교체 공수 계산
- **컬러교환 제한**: 시프트별(주간/야간) 15회 이하로 제한
- **D+1 수요 선반영**: D0 생산만으로 D+1 수요까지 커버
- **일간 연속성**: 전날 마지막 상태(지그템플릿, 컬러, 순서)를 다음날 시작 조건으로 연결

## 시스템 변수

| 변수 | 값 | 설명 |
|------|-----|------|
| HANGERS | 140 | 컨베이어 총 행어 수 |
| JIGS_PER_HANGER | 2 | 행어당 지그 수 |
| ROTATIONS_PER_DAY | 10 | 일일 회전 수 |
| JIG_BUDGET_DAY | 150 | 주간 지그교체 예산 |
| JIG_BUDGET_NIGHT | 150 | 야간 지그교체 예산 |

## 지그 그룹

| 그룹 | 명칭 | 최대지그 | PCS/지그 |
|------|------|---------|----------|
| A | THPE STD/LDT+SP3 | 100 | 1 |
| B | NQ5 FRT (STD+XLINE) | 100 | 1 |
| B2 | NQ5 FRT STD 전용 | 50 | 1 |
| C | OV1 | 80 | 1 |
| D | JX EV FRT | 100 | 1 |
| E | JX CROSS | 80 | 1 |
| F | JX EV RR | 50 | 1 |
| G | AX PE | 80 | 1 |
| H | THPE RR | 50 | 2 |
| I | NQ5 RR | 70 | 1 |

## 설치 및 실행

### 로컬 실행
```bash
# 가상환경 생성
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# 패키지 설치
pip install -r requirements.txt

# 웹앱 실행
python app.py
```

### Railway 배포
1. GitHub 저장소 연결
2. Railway에서 PostgreSQL 추가
3. 환경변수 자동 설정됨 (`DATABASE_URL`)

## API 엔드포인트

| 메서드 | 경로 | 설명 |
|--------|------|------|
| GET | `/api/config` | 시스템 설정 조회 |
| POST | `/api/config` | 시스템 설정 수정 |
| GET | `/api/jig-groups` | 지그그룹 조회 |
| POST | `/api/demand/upload` | 엑셀 수요 업로드 |
| GET | `/api/demand?date=YYYY-MM-DD` | 수요 데이터 조회 |
| POST | `/api/schedule` | 스케줄링 실행 |
| GET | `/api/report?date=YYYY-MM-DD` | HTML 리포트 생성 |

## 파일 구조

```
paintplan/
├── app.py              # Flask 웹앱 (Railway 배포용)
├── generate_report.py  # 핵심 스케줄링 로직
├── models.py           # DB 모델
├── config.py           # 설정
├── templates/
│   └── index.html      # 웹 UI
├── requirements.txt    # Python 패키지
├── Procfile           # Railway/Heroku 실행 설정
└── PRODUCTION_LOGIC.md # 상세 로직 문서
```

## 상세 로직

[PRODUCTION_LOGIC.md](PRODUCTION_LOGIC.md) 참조

## 라이센스

Private - 무단 복제 및 배포 금지
