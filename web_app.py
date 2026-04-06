"""도장 생산계획 웹 애플리케이션"""
import os
import json
from datetime import datetime, date, timedelta
from collections import defaultdict

from flask import Flask, render_template, request, jsonify
from models import db, Product, Item, Demand, PlanConfig, PlanResult, InventoryReport

# planning engine imports
from production_planner import calculate_production_plan
from paint_scheduler import schedule_painting
from sample_data import build_color_transition_matrix, COLORS, INJECTION_PRODUCTS

app = Flask(__name__)

DB_URL = os.environ.get("DATABASE_URL", "")
if not DB_URL:
    # 로컬: SQLite 사용
    DB_URL = "sqlite:///" + os.path.join(os.path.dirname(__file__), "paint_plan.db")
elif DB_URL.startswith("postgres://"):
    # Railway: postgres:// → postgresql://
    DB_URL = DB_URL.replace("postgres://", "postgresql://", 1)

app.config["SQLALCHEMY_DATABASE_URI"] = DB_URL
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db.init_app(app)

COLOR_MATRIX = build_color_transition_matrix()
COLOR_LIST = [c["id"] for c in COLORS]
PRODUCT_LIST = [p["id"] for p in INJECTION_PRODUCTS]


# ── 페이지 ──────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html",
                           colors=COLORS,
                           products=INJECTION_PRODUCTS)


# ── API: 설정 ───────────────────────────────────────

@app.route("/api/config", methods=["GET"])
def get_config():
    cfg = PlanConfig.query.first()
    if not cfg:
        cfg = PlanConfig()
        db.session.add(cfg)
        db.session.commit()
    return jsonify({
        "hanger_count": cfg.hanger_count,
        "jigs_per_hanger": 2,  # 기본값 (행어당 지그 수는 Product별로 관리)
        "rotations_per_day": cfg.rotations_per_day,
        "max_jig_changes": cfg.max_jig_changes,
        "safety_stock_days": cfg.safety_stock_days,
        "planning_days": cfg.planning_days,
    })


@app.route("/api/config", methods=["POST"])
def save_config():
    data = request.json
    cfg = PlanConfig.query.first()
    if not cfg:
        cfg = PlanConfig()
        db.session.add(cfg)
    for k in ["hanger_count", "rotations_per_day",
              "max_jig_changes", "safety_stock_days", "planning_days"]:
        if k in data:
            setattr(cfg, k, int(data[k]))
    # jigs_per_hanger는 Product별로 관리되므로 여기서 무시
    db.session.commit()
    return jsonify({"ok": True})


# ── API: 사출품 (지그/행어 설정) ────────────────────

@app.route("/api/products", methods=["GET"])
def get_products():
    prods = Product.query.order_by(Product.product_code).all()
    return jsonify([{
        "product_code": p.product_code,
        "name": p.name,
        "jigs_per_hanger": p.jigs_per_hanger,
        "jig_count": p.jig_count,
        "max_hangers": p.jig_count // p.jigs_per_hanger if p.jigs_per_hanger else 0,
    } for p in prods])


@app.route("/api/products", methods=["POST"])
def save_products():
    """사출품별 지그/행어 설정 저장: [{product_code, name, jigs_per_hanger}, ...]"""
    data = request.json
    for item in data:
        pc = item["product_code"]
        prod = Product.query.filter_by(product_code=pc).first()
        if prod:
            prod.jigs_per_hanger = int(item.get("jigs_per_hanger", 2))
            if "jig_count" in item:
                prod.jig_count = int(item["jig_count"])
            if "name" in item:
                prod.name = item["name"]
        else:
            prod = Product(
                product_code=pc,
                name=item.get("name", pc),
                jigs_per_hanger=int(item.get("jigs_per_hanger", 2)),
                jig_count=int(item.get("jig_count", 80)),
            )
            db.session.add(prod)
    db.session.commit()
    return jsonify({"ok": True})


# ── API: 아이템 마스터 ──────────────────────────────

@app.route("/api/items", methods=["GET"])
def get_items():
    items = Item.query.order_by(Item.product_code, Item.color).all()
    prods = {p.product_code: p.jigs_per_hanger for p in Product.query.all()}
    return jsonify([{
        "item_code": i.item_code,
        "product_code": i.product_code,
        "color": i.color,
        "name": i.name,
        "initial_stock": i.initial_stock,
        "jigs_per_hanger": prods.get(i.product_code, 2),
    } for i in items])


@app.route("/api/items", methods=["POST"])
def add_item():
    data = request.json
    pc = data["product_code"].strip()
    color = data["color"].strip().upper()
    item_code = f"{pc}-{color}"
    existing = Item.query.filter_by(item_code=item_code).first()
    if existing:
        return jsonify({"error": f"{item_code} 이미 존재"}), 400
    item = Item(
        product_code=pc, color=color, item_code=item_code,
        name=data.get("name", ""), initial_stock=int(data.get("initial_stock", 0))
    )
    db.session.add(item)
    db.session.commit()
    return jsonify({"ok": True, "item_code": item_code})


@app.route("/api/items/bulk", methods=["POST"])
def bulk_add_items():
    """샘플 데이터 일괄 등록 (사출품 + 완성품)"""
    data = request.json
    products = data.get("products", [])
    colors = data.get("colors", [])
    jigs_map = data.get("jigs_map", {})  # {"001": 2, "002": 4, ...}

    jig_counts = data.get("jig_counts", {})  # {"001": 80, "002": 60, ...}

    # 사출품 등록
    for pc in products:
        if not Product.query.filter_by(product_code=pc).first():
            jph = int(jigs_map.get(pc, 2))
            jc = int(jig_counts.get(pc, 80))
            db.session.add(Product(product_code=pc, name=pc,
                                   jigs_per_hanger=jph, jig_count=jc))

    # 완성품 등록
    count = 0
    for pc in products:
        for color in colors:
            item_code = f"{pc}-{color}"
            if not Item.query.filter_by(item_code=item_code).first():
                item = Item(product_code=pc, color=color, item_code=item_code,
                            name=f"{pc} {color}", initial_stock=0)
                db.session.add(item)
                count += 1
    db.session.commit()
    return jsonify({"ok": True, "added": count})


@app.route("/api/items/stock", methods=["POST"])
def update_stock():
    """초기 재고 업데이트"""
    data = request.json  # {item_code: qty, ...}
    for item_code, qty in data.items():
        item = Item.query.filter_by(item_code=item_code).first()
        if item:
            item.initial_stock = int(qty)
    db.session.commit()
    return jsonify({"ok": True})


@app.route("/api/items/stock/auto", methods=["POST"])
def auto_stock():
    """초기재고 = 1일치 수요 (안전재고 부족으로 즉시 생산 시작)"""
    demands = Demand.query.order_by(Demand.date).all()
    if not demands:
        return jsonify({"error": "수요 먼저 등록"}), 400

    # 날짜별 아이템별 수요
    daily = defaultdict(lambda: defaultdict(int))
    for d in demands:
        daily[d.date][d.item_code] += d.qty

    dates = sorted(daily.keys())
    first_day = dates[0] if dates else None

    # 첫 1일 수요 = 1일치 재고
    item_stock = defaultdict(int)
    if first_day:
        for ic, qty in daily[first_day].items():
            item_stock[ic] = qty

    updated = 0
    for item in Item.query.all():
        stock = item_stock.get(item.item_code, 0)
        item.initial_stock = stock
        if stock > 0:
            updated += 1
    db.session.commit()
    return jsonify({"ok": True, "updated": updated, "stock_days": 1})


# ── API: 수요 ───────────────────────────────────────

@app.route("/api/demand", methods=["GET"])
def get_demand():
    rows = Demand.query.order_by(Demand.date, Demand.item_code).all()
    result = defaultdict(dict)
    for r in rows:
        result[r.date.isoformat()][r.item_code] = r.qty
    return jsonify(dict(result))


@app.route("/api/demand", methods=["POST"])
def save_demand():
    """수요 저장: {date: {item_code: qty, ...}, ...}"""
    data = request.json
    for date_str, items in data.items():
        dt = datetime.strptime(date_str, "%Y-%m-%d").date()
        for item_code, qty in items.items():
            qty = int(qty)
            row = Demand.query.filter_by(item_code=item_code, date=dt).first()
            if row:
                row.qty = qty
            else:
                row = Demand(item_code=item_code, date=dt, qty=qty)
                db.session.add(row)
    db.session.commit()
    return jsonify({"ok": True})


@app.route("/api/demand/sample", methods=["POST"])
def generate_sample_demand():
    """샘플 수요 생성"""
    import random
    random.seed(42)

    items = Item.query.all()
    if not items:
        return jsonify({"error": "아이템을 먼저 등록하세요"}), 400

    cfg = PlanConfig.query.first()
    if not cfg:
        cfg = PlanConfig(hanger_count=140, rotations_per_day=10,
                         max_jig_changes=280, safety_stock_days=3, planning_days=12)
        db.session.add(cfg)
        db.session.commit()
    days = (cfg.planning_days or 12) + (cfg.safety_stock_days or 3)
    start = date(2026, 3, 18)

    # 주력 차종: NQ5 (지그 가장 많음, ~40%)
    major_codes = set()
    products = Product.query.all()
    if products:
        # 지그 수 상위 2개 차종을 주력으로
        sorted_prods = sorted(products, key=lambda p: -(p.jig_count or 0))
        top2 = set(p.product_code for p in sorted_prods[:2])
        major_codes = top2
    if not major_codes:
        major_codes = {"NQ5_FRT", "NQ5_RR"}

    major_items = [i for i in items if i.product_code in major_codes]
    minor_items = [i for i in items if i.product_code not in major_codes]

    Demand.query.delete()

    # 일일 용량 계산
    product_jigs = {p.product_code: p.jigs_per_hanger for p in products}
    avg_jigs = sum(product_jigs.values()) / len(product_jigs) if product_jigs else 2
    daily_cap = int((cfg.hanger_count or 140) * avg_jigs * (cfg.rotations_per_day or 10))

    # 수요 목표 = 일일 용량의 90~100%
    for d in range(days):
        dt = start + timedelta(days=d)
        target = random.randint(int(daily_cap * 0.90), int(daily_cap * 1.00))
        major_target = int(target * 0.40)
        minor_target = target - major_target

        # 주력 (~40%)
        n_major = min(len(major_items), random.randint(12, 18))
        active = random.sample(major_items, n_major)
        total = 0
        for item in active:
            remaining = major_target - total
            if remaining <= 0:
                break
            qty = min(random.randint(50, 200), remaining)
            if qty > 0:
                db.session.add(Demand(item_code=item.item_code, date=dt, qty=qty))
                total += qty
        # 남은 major_target 배분
        if total < major_target and active:
            leftover = major_target - total
            item = random.choice(active)
            db.session.add(Demand(item_code=item.item_code, date=dt, qty=leftover))

        # 비주력 (~60%)
        n_minor = min(len(minor_items), random.randint(25, 45))
        active = random.sample(minor_items, n_minor)
        total = 0
        for item in active:
            remaining = minor_target - total
            if remaining <= 0:
                break
            qty = min(random.randint(20, 120), remaining)
            if qty > 0:
                db.session.add(Demand(item_code=item.item_code, date=dt, qty=qty))
                total += qty
        # 남은 minor_target 배분
        if total < minor_target and active:
            leftover = minor_target - total
            item = random.choice(active)
            db.session.add(Demand(item_code=item.item_code, date=dt, qty=leftover))

    db.session.commit()
    return jsonify({"ok": True, "days": days})


# ── API: 안전재고 상태 ──────────────────────────────

@app.route("/api/safety-status", methods=["GET"])
def safety_status():
    """각 아이템별 D+0, D+1, D+2, D+3 충족 여부"""
    items = {i.item_code: i.initial_stock for i in Item.query.all()}
    demands = Demand.query.order_by(Demand.date).all()

    # 날짜별 수요 정리
    daily = defaultdict(lambda: defaultdict(int))
    for d in demands:
        daily[d.date.isoformat()][d.item_code] += d.qty

    dates = sorted(daily.keys())
    if not dates:
        return jsonify([])

    result = []
    for item_code, stock in sorted(items.items()):
        row = {"item_code": item_code, "initial_stock": stock, "dates": {}}
        for i, dt in enumerate(dates):
            demand_d0 = daily[dt].get(item_code, 0)
            cum = [demand_d0]
            for offset in range(1, 4):
                if i + offset < len(dates):
                    cum.append(cum[-1] + daily[dates[i + offset]].get(item_code, 0))
                else:
                    cum.append(cum[-1])

            row["dates"][dt] = {
                "demand": demand_d0,
                "d0_ok": stock >= cum[0],
                "d1_ok": stock >= cum[1] if len(cum) > 1 else True,
                "d2_ok": stock >= cum[2] if len(cum) > 2 else True,
                "d3_ok": stock >= cum[3] if len(cum) > 3 else True,
            }
        result.append(row)
    return jsonify(result)


# ── API: 생산계획 실행 ──────────────────────────────

@app.route("/api/plan/run", methods=["POST"])
def run_plan():
    """생산계획 실행"""
    cfg = PlanConfig.query.first() or PlanConfig()

    # 사출품별 지그/행어 로드
    product_jigs = {p.product_code: p.jigs_per_hanger
                    for p in Product.query.all()}
    # 평균 지그/행어 (용량 계산용 대표값)
    avg_jigs = (sum(product_jigs.values()) / len(product_jigs)
                if product_jigs else 2)

    # config 모듈 값 오버라이드
    import config as cfg_mod
    cfg_mod.HANGER_COUNT = cfg.hanger_count
    cfg_mod.JIGS_PER_HANGER = round(avg_jigs)  # 대표값 (스케줄러용)
    cfg_mod.ROTATIONS_PER_DAY = cfg.rotations_per_day
    cfg_mod.MAX_JIG_CHANGES_PER_DAY = cfg.max_jig_changes
    cfg_mod.SAFETY_STOCK_DAYS = cfg.safety_stock_days
    cfg_mod.PLANNING_DAYS = cfg.planning_days
    # 용량은 실제 지그/행어 비율 고려
    cfg_mod.DAILY_CAPACITY = int(
        cfg.hanger_count * avg_jigs * cfg.rotations_per_day
    )

    import importlib
    importlib.reload(cfg_mod)
    import production_planner
    production_planner.DAILY_CAPACITY = cfg_mod.DAILY_CAPACITY

    # DB에서 수요 로드
    items_map = {i.item_code: i for i in Item.query.all()}
    demands = Demand.query.order_by(Demand.date).all()

    daily_paint = defaultdict(lambda: defaultdict(int))
    for d in demands:
        item = items_map.get(d.item_code)
        if item:
            daily_paint[d.date.isoformat()][(item.product_code, item.color)] += d.qty

    daily_paint = {dt: dict(items) for dt, items in daily_paint.items()}

    if not daily_paint:
        return jsonify({"error": "수요 데이터가 없습니다"}), 400

    # 초기 재고 설정 (production_planner에 전달)
    initial_stocks = {}
    for item in items_map.values():
        if item.initial_stock > 0:
            initial_stocks[(item.product_code, item.color)] = item.initial_stock

    # 지그 보유량 → 행어 상한
    jig_limits = {}
    for pc, jph in product_jigs.items():
        prod = Product.query.filter_by(product_code=pc).first()
        if prod and prod.jig_count:
            jig_limits[pc] = prod.jig_count // jph

    color_matrix = build_color_transition_matrix()

    # 1차: 수요 기반으로 도장 스케줄링 (풀 생산)
    schedules = schedule_painting(daily_paint, color_matrix, jig_limits=jig_limits)

    # 2차: 스케줄러 실제 생산량으로 재고 계산
    actual_production = {}
    plan_dates = sorted(daily_paint.keys())[:cfg.planning_days]
    for day_sched in schedules:
        day_prod = defaultdict(int)
        for rot in day_sched.rotations:
            for seg_idx, product, color, n_h, pieces in rot.cells:
                day_prod[(product, color)] += pieces
        actual_production[day_sched.date] = dict(day_prod)

    # 재고 리포트 생성 (실제 생산량 기반)
    all_items = set()
    for items in daily_paint.values():
        all_items.update(items.keys())
    for items in actual_production.values():
        all_items.update(items.keys())

    inventory = defaultdict(int)
    if initial_stocks:
        for item, qty in initial_stocks.items():
            inventory[item] = qty

    inv_report = []
    for date_str in plan_dates:
        day_demand = daily_paint.get(date_str, {})
        day_prod = actual_production.get(date_str, {})

        # 안전재고 목표: 향후 3일 수요 합
        safety_target = defaultdict(int)
        date_idx = plan_dates.index(date_str) if date_str in plan_dates else -1
        all_dates = sorted(daily_paint.keys())
        for offset in range(1, (cfg.safety_stock_days or 3) + 1):
            future_idx = all_dates.index(date_str) + offset if date_str in all_dates else -1
            if 0 <= future_idx < len(all_dates):
                for item, qty in daily_paint.get(all_dates[future_idx], {}).items():
                    safety_target[item] += qty

        day_report = {"date": date_str, "items": {}}
        for item in all_items:
            stock_start = inventory[item]
            demand = day_demand.get(item, 0)
            produced = day_prod.get(item, 0)
            stock_end = stock_start + produced - demand
            safety = safety_target.get(item, 0)

            if stock_end < 0:
                status = "SHORTAGE"
            elif stock_end < safety:
                status = "BELOW_SAFETY"
            else:
                status = "OK"

            day_report["items"][item] = {
                "demand": demand, "production": produced,
                "stock_start": stock_start, "stock_end": stock_end,
                "safety_target": safety, "status": status,
            }
            inventory[item] = stock_end

        inv_report.append(day_report)

    production_plan = actual_production

    # run_id
    import time
    run_id = int(time.time())

    # 결과 저장
    PlanResult.query.filter_by(run_id=run_id).delete()
    InventoryReport.query.delete()

    # 회전별 결과 저장
    for day_sched in schedules:
        dt = datetime.strptime(day_sched.date, "%Y-%m-%d").date()
        for rot in day_sched.rotations:
            for seg_idx, product, color, n_h, pieces in rot.cells:
                db.session.add(PlanResult(
                    run_id=run_id, date=dt, rotation=rot.rotation,
                    segment_idx=seg_idx, product_code=product, color=color,
                    n_hangers=n_h, pieces=pieces,
                ))

    # 재고 리포트 저장
    for report in inv_report:
        dt = datetime.strptime(report["date"], "%Y-%m-%d").date()
        for (pc, color), info in report["items"].items():
            # 재고가 있거나 활동이 있는 모든 아이템 저장
            if info["demand"] > 0 or info["production"] > 0 or info["stock_start"] > 0 or info["stock_end"] > 0:
                item_code = f"{pc}-{color}"
                db.session.add(InventoryReport(
                    run_id=run_id, date=dt, item_code=item_code,
                    demand=info["demand"], production=info["production"],
                    stock_start=info["stock_start"], stock_end=info["stock_end"],
                    safety_target=info["safety_target"], status=info["status"],
                ))

    db.session.commit()

    # 응답: 요약 데이터
    summary = []
    for day_sched in schedules:
        total_prod = sum(r.total_produced for r in day_sched.rotations)
        total_empty = sum(r.total_empty for r in day_sched.rotations)
        total_trans = sum(len(r.color_transitions) for r in day_sched.rotations)
        jig_str = " | ".join(
            f"{s.product}({s.n_hangers})" for s in day_sched.template.segments
        )

        rotations_data = []
        for rot in day_sched.rotations:
            cells = []
            for seg_idx, product, color, n_h, pieces in rot.cells:
                cells.append({
                    "seg": seg_idx, "product": product, "color": color,
                    "hangers": n_h, "pieces": pieces,
                    "item": f"{product}-{color}",
                })
            rotations_data.append({
                "rotation": rot.rotation,
                "produced": rot.total_produced,
                "transitions": len(rot.color_transitions),
                "empty": rot.total_empty,
                "cells": cells,
            })

        summary.append({
            "date": day_sched.date,
            "produced": total_prod,
            "empty": total_empty,
            "transitions": total_trans,
            "jig_changes": day_sched.total_jig_changes,
            "template": jig_str,
            "rotations": rotations_data,
        })

    # 재고 그래프 데이터
    chart_data = _build_chart_data(inv_report)

    return jsonify({
        "ok": True,
        "run_id": run_id,
        "summary": summary,
        "chart": chart_data,
    })


def _build_chart_data(inv_report):
    """Chart.js용 데이터 구성"""
    dates = [r["date"] for r in inv_report]

    # 전체 합산
    totals = {"dates": dates, "demand": [], "production": [],
              "stock_start": [], "stock_end": [], "safety": []}
    for r in inv_report:
        totals["demand"].append(sum(v["demand"] for v in r["items"].values()))
        totals["production"].append(sum(v["production"] for v in r["items"].values()))
        totals["stock_start"].append(sum(v["stock_start"] for v in r["items"].values()))
        totals["stock_end"].append(sum(v["stock_end"] for v in r["items"].values()))
        totals["safety"].append(sum(v["safety_target"] for v in r["items"].values()))

    # 아이템별
    all_items = set()
    for r in inv_report:
        all_items.update(r["items"].keys())

    per_item = {}
    for pc, color in sorted(all_items):
        item_code = f"{pc}-{color}"
        item_data = {"dates": dates, "demand": [], "production": [],
                     "stock_end": [], "safety": []}
        for r in inv_report:
            info = r["items"].get((pc, color), {})
            item_data["demand"].append(info.get("demand", 0))
            item_data["production"].append(info.get("production", 0))
            item_data["stock_end"].append(info.get("stock_end", 0))
            item_data["safety"].append(info.get("safety_target", 0))
        per_item[item_code] = item_data

    return {"total": totals, "items": per_item}


# ── API: 결과 조회 ──────────────────────────────────

@app.route("/api/plan/chart", methods=["GET"])
def get_chart_data():
    """저장된 재고 리포트에서 차트 데이터"""
    rows = InventoryReport.query.order_by(InventoryReport.date).all()
    if not rows:
        return jsonify({"error": "계획을 먼저 실행하세요"}), 404

    dates = sorted(set(r.date.isoformat() for r in rows))
    by_date = defaultdict(list)
    for r in rows:
        by_date[r.date.isoformat()].append(r)

    totals = {"dates": dates, "demand": [], "production": [],
              "stock_end": [], "safety": []}
    per_item = defaultdict(lambda: {"dates": dates, "demand": [], "production": [],
                                     "stock_end": [], "safety": []})

    for dt in dates:
        day_rows = by_date[dt]
        totals["demand"].append(sum(r.demand for r in day_rows))
        totals["production"].append(sum(r.production for r in day_rows))
        totals["stock_end"].append(sum(r.stock_end for r in day_rows))
        totals["safety"].append(sum(r.safety_target for r in day_rows))

        item_map = {r.item_code: r for r in day_rows}
        all_items = set(r.item_code for r in InventoryReport.query.all())
        for ic in all_items:
            r = item_map.get(ic)
            per_item[ic]["demand"].append(r.demand if r else 0)
            per_item[ic]["production"].append(r.production if r else 0)
            per_item[ic]["stock_end"].append(r.stock_end if r else 0)
            per_item[ic]["safety"].append(r.safety_target if r else 0)

    return jsonify({"total": totals, "items": dict(per_item)})


# ── 초기화 ──────────────────────────────────────────

@app.cli.command("init-db")
def init_db():
    """DB 테이블 생성"""
    db.create_all()
    print("DB 테이블 생성 완료")


if __name__ == "__main__":
    with app.app_context():
        db.create_all()
    app.run(debug=True, port=5001)
