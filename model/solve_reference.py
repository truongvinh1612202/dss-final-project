"""Lời giải đối chiếu (scipy) cho 4 bài toán Phần B.

Chạy: python -I model/solve_reference.py
Đầu ra: output/reference.json, dùng để kiểm tra kết quả Solver/Goal Seek trong Excel.
"""

import json
import sys
from pathlib import Path

import numpy as np
from scipy.optimize import linprog, minimize

sys.path.insert(0, str(Path(__file__).resolve().parent))
import params as P

OUT = P.ROOT / "output" / "reference.json"


def problem1(base):
    sub = base["subcat"]
    margin = (
        sum(sub.loc[s, "revenue"] * P.SUBCAT_MARGIN[s] for s in sub.index)
        / sub["revenue"].sum()
    )
    ship = base["ship"]
    ship_avg = (
        sum(ship.loc[m, "orders"] * P.SHIP_COST[m] for m in P.SHIP_MODES)
        / ship["orders"].sum()
    )
    fee, handling = (
        P.value(P.GENERAL, "payment_fee"),
        P.value(P.GENERAL, "handling_cost"),
    )
    fixed, target = (
        P.value(P.GENERAL, "fixed_cost"),
        P.value(P.GENERAL, "profit_target"),
    )
    aov = base["aov"]
    orders = P.orders_2019(base)
    contrib = aov * (margin - fee) - handling - ship_avg
    return {
        "gross_margin": margin,
        "ship_cost_avg": ship_avg,
        "contribution_per_order": contrib,
        "orders_2019": orders,
        "profit_2019": orders * contrib - fixed,
        "gs_breakeven_orders": fixed / contrib,
        "gs_target_orders": (fixed + target) / contrib,
        "gs_target_aov": ((fixed + target) / orders + handling + ship_avg)
        / (margin - fee),
    }


def mkt_uplift(b, rev, alpha, scale):
    return float(np.sum(rev * alpha * (1 - np.exp(-np.asarray(b) / scale))))


def problem2(base):
    regions = sorted(P.MKT_ALPHA)
    rev = base["region"].reindex(regions).to_numpy()
    alpha = np.array([P.MKT_ALPHA[r] for r in regions])
    scale = P.value(P.MKT, "scale")
    lo, hi = P.value(P.MKT, "min_region"), P.value(P.MKT, "max_region")

    def solve(budget):
        res = minimize(
            lambda b: -mkt_uplift(b, rev, alpha, scale),
            x0=np.full(4, budget / 4),
            bounds=[(lo, hi)] * 4,
            constraints=[{"type": "ineq", "fun": lambda b: budget - b.sum()}],
            method="SLSQP",
        )
        assert res.success, res.message
        return res.x

    out = {}
    for name, budget in P.MKT_SCENARIOS.items():
        b = solve(budget)
        out[name] = {
            "allocation": dict(zip(regions, b.round(0).tolist())),
            "uplift": mkt_uplift(b, rev, alpha, scale),
            "equal_split_uplift": mkt_uplift(np.full(4, budget / 4), rev, alpha, scale),
        }
    # Goal Seek: tổng ngân sách chia đều để đạt doanh thu tăng thêm mục tiêu
    goal = P.value(P.MKT, "revenue_goal")
    lo_b, hi_b = 1.0, 1e7
    for _ in range(200):
        mid = (lo_b + hi_b) / 2
        if mkt_uplift(np.full(4, mid / 4), rev, alpha, scale) < goal:
            lo_b = mid
        else:
            hi_b = mid
    out["gs_equal_split_budget"] = hi_b
    return out


def problem3(base, budget=None, capacity=None):
    sub = base["subcat"]
    names = list(sub.index)
    m = np.array([P.SUBCAT_MARGIN[s] for s in names])
    vol = np.array([P.SUBCAT_VOLUME[s] for s in names]) / 1000  # m³ cho mỗi USD
    demand = sub["revenue"].to_numpy() * (1 + P.value(P.INV, "demand_growth"))
    cost_cap = demand * (1 - m)
    budget = budget or P.value(P.INV, "budget")
    capacity = capacity or P.value(P.INV, "capacity")
    min_service = P.value(P.INV, "min_service")

    res = linprog(
        c=-(m / (1 - m)),
        A_ub=np.vstack([np.ones(len(names)), vol]),
        b_ub=[budget, capacity],
        bounds=list(zip(min_service * cost_cap, cost_cap)),
        method="highs",
    )
    assert res.success, res.message
    x = res.x
    return {
        "purchase": dict(zip(names, x.round(2).tolist())),
        "gross_profit": -res.fun,
        "budget_used": x.sum(),
        "space_used": float(vol @ x),
        "shadow_budget": -res.ineqlin.marginals[0],
        "shadow_capacity": -res.ineqlin.marginals[1],
        "total_cost_demand": cost_cap.sum(),
        "total_volume_demand": float(vol @ cost_cap),
    }


def problem4(base, fast_target=None):
    ship = base["ship"]
    n = P.orders_2019(base)
    cost = np.array([P.SHIP_COST[m] for m in P.SHIP_MODES])
    fast = ship["fast_rate"].to_numpy()
    days = ship["avg_days"].to_numpy()
    target = fast_target or P.value(P.SHIP, "fast_target")
    min_n = P.value(P.SHIP, "min_share") * n
    cap = P.value(P.SHIP, "sameday_cap")
    res = linprog(
        c=cost,
        A_ub=np.vstack([-fast, days]),
        b_ub=[-target * n, P.value(P.SHIP, "max_avg_days") * n],
        A_eq=np.ones((1, 4)),
        b_eq=[n],
        bounds=[(min_n, cap)] + [(min_n, None)] * 3,
        method="highs",
    )
    assert res.success, res.message
    x = res.x
    return {
        "orders": dict(zip(P.SHIP_MODES, x.round(1).tolist())),
        "total_cost": res.fun,
        "fast_rate": float(fast @ x / n),
        "avg_days": float(days @ x / n),
        "shadow_fast_orders": res.ineqlin.marginals[
            0
        ],  # USD cho mỗi đơn giao nhanh thêm
        "baseline_cost": float(
            cost @ (ship["orders"].to_numpy() / ship["orders"].sum() * n)
        ),
    }


def problem4_goal_seek(base):
    """Tỷ trọng First Class cần để đạt mục tiêu giao nhanh, giữ Same Day/Second Class như 2018."""
    ship = base["ship"]
    share = ship["orders"] / ship["orders"].sum()
    fast = ship["fast_rate"]
    target = P.value(P.SHIP, "fast_target")
    fixed = (
        share["Same Day"] * fast["Same Day"]
        + share["Second Class"] * fast["Second Class"]
    )
    # STD = 1 − SD − SC − FC, fast rate của STD = fast["Standard Class"]
    rest = 1 - share["Same Day"] - share["Second Class"]
    # fixed + fc·f_fc + (rest − fc)·f_std = target
    return (target - fixed - rest * fast["Standard Class"]) / (
        fast["First Class"] - fast["Standard Class"]
    )


def main():
    sys.stdout.reconfigure(
        encoding="utf-8"
    )  # console Windows mặc định không phải UTF-8
    base = P.load_base_year()
    ref = {
        "problem1": problem1(base),
        "problem2": problem2(base),
        "problem3": {
            name: problem3(base, b, c) for name, (b, c) in P.INV_SCENARIOS.items()
        },
        "problem4": {name: problem4(base, t) for name, t in P.SHIP_SCENARIOS.items()},
        "problem4_gs_first_class_share": problem4_goal_seek(base),
    }
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(
        json.dumps(ref, ensure_ascii=False, indent=2, default=float), encoding="utf-8"
    )
    print(json.dumps(ref, ensure_ascii=False, indent=1, default=float))


if __name__ == "__main__":
    main()
