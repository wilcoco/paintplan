"""DB 모델 — PostgreSQL + SQLAlchemy"""
from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


class Product(db.Model):
    """사출품 마스터 — 지그 타입 결정"""
    __tablename__ = "products"
    id = db.Column(db.Integer, primary_key=True)
    product_code = db.Column(db.String(10), unique=True, nullable=False)  # 001, 002 ...
    name = db.Column(db.String(100))
    jigs_per_hanger = db.Column(db.Integer, default=2)  # 2, 3, 4, 6 선택
    jig_count = db.Column(db.Integer, default=80)  # 보유 지그 수량


class Item(db.Model):
    """완성품 마스터: 사출품번호 + 컬러"""
    __tablename__ = "items"
    id = db.Column(db.Integer, primary_key=True)
    product_code = db.Column(db.String(10), nullable=False)   # 001, 002 ...
    color = db.Column(db.String(20), nullable=False)          # RED, BLUE ...
    item_code = db.Column(db.String(30), unique=True, nullable=False)  # 001-RED
    name = db.Column(db.String(100))
    initial_stock = db.Column(db.Integer, default=0)

    __table_args__ = (
        db.UniqueConstraint("product_code", "color", name="uq_product_color"),
    )


class Demand(db.Model):
    """일별 수요"""
    __tablename__ = "demand"
    id = db.Column(db.Integer, primary_key=True)
    item_code = db.Column(db.String(30), db.ForeignKey("items.item_code"), nullable=False)
    date = db.Column(db.Date, nullable=False)
    qty = db.Column(db.Integer, nullable=False, default=0)

    __table_args__ = (
        db.UniqueConstraint("item_code", "date", name="uq_item_date"),
    )


class PlanConfig(db.Model):
    """시스템 변수"""
    __tablename__ = "plan_config"
    id = db.Column(db.Integer, primary_key=True)
    hanger_count = db.Column(db.Integer, default=140)
    rotations_per_day = db.Column(db.Integer, default=10)
    max_jig_changes = db.Column(db.Integer, default=280)
    safety_stock_days = db.Column(db.Integer, default=3)
    planning_days = db.Column(db.Integer, default=12)


class PlanResult(db.Model):
    """생산계획 결과 (회전별)"""
    __tablename__ = "plan_results"
    id = db.Column(db.Integer, primary_key=True)
    run_id = db.Column(db.Integer, nullable=False)
    date = db.Column(db.Date, nullable=False)
    rotation = db.Column(db.Integer, nullable=False)
    segment_idx = db.Column(db.Integer, nullable=False)
    product_code = db.Column(db.String(10))
    color = db.Column(db.String(20))
    n_hangers = db.Column(db.Integer)
    pieces = db.Column(db.Integer)


class InventoryReport(db.Model):
    """일별 재고 리포트"""
    __tablename__ = "inventory_report"
    id = db.Column(db.Integer, primary_key=True)
    run_id = db.Column(db.Integer, nullable=False)
    date = db.Column(db.Date, nullable=False)
    item_code = db.Column(db.String(30), nullable=False)
    demand = db.Column(db.Integer, default=0)
    production = db.Column(db.Integer, default=0)
    stock_start = db.Column(db.Integer, default=0)
    stock_end = db.Column(db.Integer, default=0)
    safety_target = db.Column(db.Integer, default=0)
    status = db.Column(db.String(20))
