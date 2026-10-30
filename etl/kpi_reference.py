"""Tính giá trị tham chiếu cho các KPI của dashboard (Phần A).

Chạy: python -I etl/kpi_reference.py (sau etl/build_star_schema.py)
Đầu vào : data/star/*.csv
Đầu ra  : docs/KPI_doi_chieu.md, dùng để so với số trên Power BI

Định nghĩa KPI khớp với các measure DAX trong docs/KPI_va_DAX.md.
"""

import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
STAR = ROOT / "data" / "star"
REPORT = ROOT / "docs" / "KPI_doi_chieu.md"

FAST_SHIP_DAYS = 2  # đơn "giao nhanh" = giao trong vòng 2 ngày


def md_table(rows, columns):
    lines = ["| " + " | ".join(columns) + " |", "|" + "---|" * len(columns)]
    for r in rows:
        lines.append("| " + " | ".join(str(r[c]) for c in columns) + " |")
    return "\n".join(lines)


def load():
    t = {
        p.stem: pd.read_csv(p, dtype={"Postal Code": str}, encoding="utf-8-sig")
        for p in STAR.glob("*.csv")
    }
    fact = (
        t["Fact_Sales"]
        .merge(
            t["Dim_Date"][["DateKey", "Year"]],
            left_on="OrderDateKey",
            right_on="DateKey",
        )
        .merge(t["Dim_Customer"], on="Customer ID")
        .merge(t["Dim_Product"], on="ProductKey")
        .merge(t["Dim_Geography"], on="GeoKey")
        .merge(t["Dim_ShipMode"], on="ShipModeKey")
    )
    assert len(fact) == len(t["Fact_Sales"])
    return fact


def kpis(df):
    """Bộ KPI cho một tập dòng Fact (tương đương một ngữ cảnh lọc trong Power BI)."""
    revenue = df["Sales"].sum()
    orders = df["Order ID"].nunique()
    customers = df["Customer ID"].nunique()
    orders_per_customer = df.groupby("Customer ID")["Order ID"].nunique()
    # Ship Days giống nhau trên mọi dòng của một đơn → tính ở mức đơn hàng
    ship_days = df.groupby("Order ID")["Ship Days"].max()
    return {
        "Doanh thu": revenue,
        "Số đơn": orders,
        "AOV": revenue / orders,
        "Số khách hàng": customers,
        "Tỷ lệ KH quay lại": (orders_per_customer >= 2).sum() / customers,
        "Ngày giao TB": ship_days.mean(),
        "Tỷ lệ giao nhanh": (ship_days <= FAST_SHIP_DAYS).mean(),
    }


def fmt(k, v):
    if k in ("Doanh thu", "AOV"):
        return f"{v:,.2f}"
    if k.startswith(("Tỷ lệ", "YoY")) or k == "Tỷ trọng DT":
        return "—" if pd.isna(v) else f"{v:.1%}"
    if k == "Ngày giao TB":
        return f"{v:.2f}"
    return f"{v:,}"


def breakdown(df, by, sort_by_revenue=True):
    """KPI theo từng giá trị của `by`; nếu không sắp theo doanh thu thì giữ thứ tự xuất hiện."""
    total = df["Sales"].sum()
    rows = [
        {by: key, **kpis(g), "Tỷ trọng DT": g["Sales"].sum() / total}
        for key, g in df.groupby(by, sort=False)
    ]
    if sort_by_revenue:
        rows.sort(key=lambda r: r["Doanh thu"], reverse=True)
    return rows


def render(rows, columns):
    """Cột đầu là nhãn nhóm, giữ nguyên; các cột còn lại định dạng theo loại KPI."""
    label, *metrics = columns
    return md_table(
        [{label: r[label], **{c: fmt(c, r[c]) for c in metrics}} for r in rows], columns
    )


def main():
    sys.stdout.reconfigure(
        encoding="utf-8"
    )  # console Windows mặc định không phải UTF-8
    fact = load()

    overall = kpis(fact)
    years = breakdown(fact.sort_values("Year"), "Year", sort_by_revenue=False)
    for prev, cur in zip([None] + years[:-1], years):
        cur["YoY doanh thu"] = (
            cur["Doanh thu"] / prev["Doanh thu"] - 1 if prev else float("nan")
        )
    assert abs(sum(r["Doanh thu"] for r in years) - overall["Doanh thu"]) < 1e-6

    kpi_cols = list(overall)
    seg_cols = ["Doanh thu", "Tỷ trọng DT", "Số đơn", "AOV", "Số khách hàng"]

    ship_rows = breakdown(
        fact.sort_values("ShipModeKey"), "Ship Mode", sort_by_revenue=False
    )

    text = f"""# Giá trị tham chiếu KPI

> File này do `etl/kpi_reference.py` tự sinh từ `data/star/`. Không sửa tay, hãy chạy lại script.
> Dùng để đối chiếu với các card và visual trên Power BI (định nghĩa xem `docs/KPI_va_DAX.md`).

## Toàn bộ giai đoạn 2015–2018

{md_table([{k: fmt(k, v) for k, v in overall.items()}], kpi_cols)}

## Theo năm (Order Date)

{render(years, ["Year", *kpi_cols, "YoY doanh thu"])}

## Theo Region

{render(breakdown(fact, "Region"), ["Region", *seg_cols])}

## Theo Category

{render(breakdown(fact, "Category"), ["Category", *seg_cols])}

## Theo Segment

{render(breakdown(fact, "Segment"), ["Segment", *seg_cols])}

## Theo Ship Mode

{render(ship_rows, ["Ship Mode", "Số đơn", "Tỷ trọng DT", "Ngày giao TB", "Tỷ lệ giao nhanh"])}

## Top 5 Sub-Category theo doanh thu

{render(breakdown(fact, "Sub-Category")[:5], ["Sub-Category", "Doanh thu", "Tỷ trọng DT", "Số đơn", "AOV"])}
"""
    REPORT.write_text(text, encoding="utf-8")
    for k, v in overall.items():
        print(f"{k:<20} {fmt(k, v)}")
    print(f"Đã ghi {REPORT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
