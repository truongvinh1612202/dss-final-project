"""Số liệu đầu vào (thật, năm 2018) và giả định cho 4 bài toán Phần B.

Dùng chung cho model/build_model.py (tạo file Excel) và model/solve_reference.py
(lời giải đối chiếu bằng scipy) để hai bên luôn cùng tham số.
"""

from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
STAR = ROOT / "data" / "star"
BASE_YEAR = 2018

# ------------------------------------------------------------------ Giả định
# Mỗi giả định: (giá trị, đơn vị, lý do) — in ra sheet Gia_dinh và báo cáo.
GENERAL = {
    "order_growth": (
        0.20,
        "%",
        "Số đơn thực tế tăng 27–28%/năm giai đoạn 2017–2018; chọn 20% cho thận trọng",
    ),
    "payment_fee": (0.02, "% doanh thu", "Phí cổng thanh toán/thẻ phổ biến 1,5–3%"),
    "handling_cost": (10.0, "USD/đơn", "Chi phí xử lý, đóng gói một đơn hàng"),
    "fixed_cost": (
        70_000.0,
        "USD/năm",
        "Chi phí cố định 2019: thuê kho, lương bộ máy, hệ thống",
    ),
    "profit_target": (
        40_000.0,
        "USD",
        "Lợi nhuận mục tiêu năm 2019 do ban giám đốc đặt",
    ),
}

# Biên lợi nhuận gộp theo Sub-Category (giả định theo đặc thù ngành bán lẻ văn phòng:
# giấy/nhãn/bìa biên cao; nội thất cồng kềnh và máy móc cạnh tranh giá biên thấp)
SUBCAT_MARGIN = {
    "Accessories": 0.22, "Appliances": 0.22, "Art": 0.30, "Binders": 0.35,
    "Bookcases": 0.10, "Chairs": 0.18, "Copiers": 0.30, "Envelopes": 0.35,
    "Fasteners": 0.30, "Furnishings": 0.28, "Labels": 0.40, "Machines": 0.10,
    "Paper": 0.38, "Phones": 0.14, "Storage": 0.20, "Supplies": 0.18, "Tables": 0.08,
}  # fmt: skip

# Thể tích lưu kho: m³ cho mỗi 1.000 USD hàng (theo giá vốn) — hàng nội thất cồng kềnh
SUBCAT_VOLUME = {
    "Accessories": 0.3, "Appliances": 1.5, "Art": 0.4, "Binders": 0.6,
    "Bookcases": 3.0, "Chairs": 2.5, "Copiers": 1.0, "Envelopes": 0.4,
    "Fasteners": 0.2, "Furnishings": 1.2, "Labels": 0.2, "Machines": 1.2,
    "Paper": 0.8, "Phones": 0.2, "Storage": 2.0, "Supplies": 0.5, "Tables": 3.0,
}  # fmt: skip

# Bài 2: phản ứng doanh thu với marketing, lợi suất giảm dần
# DT tăng thêm_r = DT2018_r × alpha_r × (1 − exp(−ngân sách_r / scale))
# alpha_r = tiềm năng tăng trưởng tối đa nhờ marketing, căn cứ tăng trưởng thật 2017→2018:
# South +30,6% nhưng thị phần thấp nhất (17%) → cao nhất; Central −2,8% → thấp nhất
MKT_ALPHA = {"Central": 0.12, "East": 0.18, "South": 0.40, "West": 0.22}
MKT = {
    "scale": (
        25_000.0,
        "USD",
        "Mức chi mà tại đó đạt ~63% tiềm năng tăng trưởng của vùng",
    ),
    "budget": (100_000.0, "USD", "Tổng ngân sách marketing 2019"),
    "min_region": (
        10_000.0,
        "USD",
        "Mức tối thiểu để duy trì hiện diện thương hiệu mỗi vùng",
    ),
    "max_region": (
        40_000.0,
        "USD",
        "Giới hạn năng lực triển khai chiến dịch của một vùng",
    ),
    "revenue_goal": (110_000.0, "USD", "Mục tiêu doanh thu tăng thêm nhờ marketing"),
}
MKT_SCENARIOS = {
    "Ngân sách thấp": 60_000.0,
    "Ngân sách cơ sở": 100_000.0,
    "Ngân sách cao": 140_000.0,
}

# Bài 3: nhập hàng
INV = {
    "demand_growth": (0.20, "%", "Nhu cầu 2019 = doanh thu 2018 × (1 + tăng trưởng)"),
    "budget": (
        520_000.0,
        "USD",
        "Ngân sách nhập hàng 2019 (theo giá vốn), ~76% nhu cầu",
    ),
    "capacity": (650.0, "m³", "Sức chứa kho khả dụng, ~68% thể tích nhu cầu"),
    "min_service": (
        0.50,
        "% nhu cầu",
        "Mỗi Sub-Category phải đáp ứng tối thiểu 50% nhu cầu",
    ),
}
INV_SCENARIOS = {
    "Thắt chặt": (460_000.0, 600.0),
    "Cơ sở": (520_000.0, 650.0),
    "Mở rộng": (600_000.0, 700.0),
}

# Bài 4: giao hàng
SHIP_COST = {
    "Same Day": 30.0,
    "First Class": 18.0,
    "Second Class": 12.0,
    "Standard Class": 7.0,
}
SHIP = {
    "fast_target": (
        0.30,
        "% đơn",
        "Mục tiêu tỷ lệ đơn giao trong ≤ 2 ngày (hiện tại ~22%)",
    ),
    "sameday_cap": (150.0, "đơn/năm", "Năng lực đội giao trong ngày"),
    "max_avg_days": (4.0, "ngày", "Thời gian giao trung bình tối đa chấp nhận được"),
    "min_share": (
        0.05,
        "% đơn",
        "Mỗi phương thức giữ tối thiểu 5% để khách vẫn có lựa chọn",
    ),
}
SHIP_SCENARIOS = {"Mục tiêu 25%": 0.25, "Mục tiêu 30%": 0.30, "Mục tiêu 35%": 0.35}
SHIP_MODES = ["Same Day", "First Class", "Second Class", "Standard Class"]
FAST_DAYS = 2


def value(d, key):
    return d[key][0]


def orders_2019(base):
    """Số đơn dự báo 2019, làm tròn đến đơn — khớp ô B1_Orders/B4_N trong Excel."""
    return round(base["orders"] * (1 + value(GENERAL, "order_growth")))


# ------------------------------------------------------------------ Dữ liệu thật
def load_base_year():
    """Số liệu thật năm 2018 từ star schema."""
    t = {
        p.stem: pd.read_csv(p, dtype={"Postal Code": str}, encoding="utf-8-sig")
        for p in STAR.glob("*.csv")
    }
    f = (
        t["Fact_Sales"]
        .merge(
            t["Dim_Date"][["DateKey", "Year"]],
            left_on="OrderDateKey",
            right_on="DateKey",
        )
        .merge(t["Dim_Product"], on="ProductKey")
        .merge(t["Dim_Geography"], on="GeoKey")
        .merge(t["Dim_ShipMode"], on="ShipModeKey")
    )
    y = f[f["Year"] == BASE_YEAR]
    prev = f[f["Year"] == BASE_YEAR - 1]

    orders = y.groupby("Order ID").agg(
        mode=("Ship Mode", "first"), days=("Ship Days", "max")
    )
    ship = (
        orders.groupby("mode")
        .agg(orders=("days", "size"), avg_days=("days", "mean"))
        .assign(
            fast_rate=orders.groupby("mode")["days"].apply(
                lambda d: (d <= FAST_DAYS).mean()
            )
        )
        .reindex(SHIP_MODES)
    )
    subcat = (
        y.groupby("Sub-Category")
        .agg(category=("Category", "first"), revenue=("Sales", "sum"))
        .sort_values("revenue", ascending=False)
    )
    return {
        "revenue": y["Sales"].sum(),
        "orders": y["Order ID"].nunique(),
        "aov": y["Sales"].sum() / y["Order ID"].nunique(),
        "orders_prev": prev["Order ID"].nunique(),
        "region": y.groupby("Region")["Sales"].sum(),
        "region_prev": prev.groupby("Region")["Sales"].sum(),
        "category": y.groupby("Category")["Sales"].sum(),
        "subcat": subcat,
        "ship": ship,
    }
