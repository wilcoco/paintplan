# 도장 생산계획 시스템

## 프로젝트 개요
자동차 범퍼 도장 생산계획 최적화 시스템
- 12일치 완성품 수요 → BOM 전개 → 도장 컨베이어 스케줄링
- 컬러/지그 교환 최소화 (TSP 기반 알고리즘)
- 안전재고 3일치 관리

## 기술 스택
- Backend: Python 3.14, Flask
- DB: SQLite (로컬), PostgreSQL (Railway 배포 시)
- Frontend: Bootstrap 5, Chart.js
- ORM: Flask-SQLAlchemy

## 핵심 파일
```
paint/
├── web_app.py           # Flask 웹앱 (API + 라우팅)
├── paint_scheduler.py   # 핵심 스케줄링 알고리즘
├── production_planner.py # 안전재고 반영 순수요 계산
├── bom_explode.py       # BOM 전개
├── injection_scheduler.py # 사출 역산
├── models.py            # DB 모델 (Product, Item, Demand, PlanResult 등)
├── config.py            # 시스템 변수 (행어, 회전, 지그한도)
├── sample_data.py       # 샘플 데이터 생성
├── main.py              # CLI 진입점
└── templates/index.html # 웹 UI (단일 페이지)
```

## 도메인 용어
- **행어(Hanger)**: 컨베이어에 걸린 도장 거치대, 140개 순환
- **지그(Jig)**: 행어에 부품을 고정하는 기구, 행어당 2개
- **회전(Rotation)**: 140행어가 도장 로봇 앞을 한 바퀴 통과, 10회전/일
- **일일 용량**: 140 × 2 × 10 = 2,800개
- **컬러 전환**: 색상 변경 시 빈 행어 삽입 필요 (세척 시간)

## 실행 방법
```bash
# 웹앱 실행
source venv/bin/activate
python web_app.py

# CLI 실행 (샘플 데이터로 테스트)
python main.py
```

## 웹 UI 탭 구성
1. 설정 - 컨베이어 파라미터, 사출품별 지그 설정
2. 아이템 - 완성품 마스터 (사출품 + 컬러)
3. 수요 - 일별 수요 입력/조회
4. 안전재고 - 재고 상태 확인
5. 생산계획 - 스케줄링 실행 및 결과 조회
6. 그래프 - 수요/생산/재고 시각화

## 개발 노트
- DB 파일: `paint_plan.db` (SQLite)
- Railway 배포 시 `DATABASE_URL` 환경변수로 PostgreSQL 연결
