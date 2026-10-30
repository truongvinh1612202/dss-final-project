"""Tách star schema từ dữ liệu sạch Superstore.

Chạy: python -I etl/build_star_schema.py (sau etl/preprocess.py)
Đầu vào : data/clean/Superstore_clean.csv
Đầu ra  : data/star/*.csv (6 bảng), docs/Kiem_tra_star_schema.md
"""

import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
CLEAN = ROOT / "data" / "clean" / "Superstore_clean.csv"
STAR = ROOT / "data" / "star"
REPORT = ROOT / "docs" / "Kiem_tra_star_schema.md"

# Dim_Date phủ trọn các năm có Order Date lẫn Ship Date (Ship Date muộn nhất: 2019-01-05)
DATE_START, DATE_END = "2015-01-01", "2019-12-31"
# ShipModeKey xếp theo tốc độ giao hàng, nhanh nhất = 1
SHIP_MODE_ORDER = ["Same Day", "First Class", "Second Class", "Standard Class"]
WEEKDAY_VI = [
    "Thứ Hai",
    "Thứ Ba",
    "Thứ Tư",
    "Thứ Năm",
    "Thứ Sáu",
    "Thứ Bảy",
    "Chủ Nhật",
]


def md_table(rows, columns):
    lines = ["| " + " | ".join(columns) + " |", "|" + "---|" * len(columns)]
    for r in rows:
        lines.append(
            "| " + " | ".join(str(r[c]).replace("|", "\\|") for c in columns) + " |"
        )
    return "\n".join(lines)


def date_key(s):
    return s.dt.strftime("%Y%m%d").astype(int)


# ---------------------------------------------------------------- Dựng bảng
def build_dim_date():
    d = pd.DataFrame({"Date": pd.date_range(DATE_START, DATE_END, freq="D")})
    d.insert(0, "DateKey", date_key(d["Date"]))
    d["Year"] = d["Date"].dt.year
    d["Quarter"] = d["Date"].dt.quarter
    d["Quarter Name"] = "Q" + d["Quarter"].astype(str)
    d["Month"] = d["Date"].dt.month
    d["Month Name"] = "Tháng " + d["Month"].astype(str)
    d["YearMonth"] = d["Date"].dt.strftime("%Y-%m")
    d["Day"] = d["Date"].dt.day
    d["DayOfWeek"] = d["Date"].dt.dayofweek + 1  # 1 = Thứ Hai … 7 = Chủ Nhật
    d["Weekday Name"] = d["DayOfWeek"].map(lambda i: WEEKDAY_VI[i - 1])
    d["IsWeekend"] = (d["DayOfWeek"] >= 6).astype(int)
    return d


def build_tables(df):
    dim_customer = (
        df[["Customer ID", "Customer Name", "Segment"]]
        .drop_duplicates()
        .sort_values("Customer ID")
    )
    dim_product = (
        df[["ProductKey", "Product ID", "Product Name", "Sub-Category", "Category"]]
        .drop_duplicates()
        .sort_values("ProductKey")
    )
    dim_geography = (
        df[["GeoKey", "City", "State", "Postal Code", "Region"]]
        .drop_duplicates()
        .sort_values("GeoKey")
    )
    dim_shipmode = pd.DataFrame(
        {
            "ShipModeKey": range(1, len(SHIP_MODE_ORDER) + 1),
            "Ship Mode": SHIP_MODE_ORDER,
        }
    )
    dim_date = build_dim_date()

    fact = df.merge(dim_shipmode, on="Ship Mode", how="left")
    fact["OrderDateKey"] = date_key(fact["Order Date"])
    fact["ShipDateKey"] = date_key(fact["Ship Date"])
    fact = fact[
        [
            "Row ID",
            "Order ID",
            "OrderDateKey",
            "ShipDateKey",
            "Customer ID",
            "ProductKey",
            "GeoKey",
            "ShipModeKey",
            "Sales",
            "Ship Days",
            "Is Sales Outlier",
        ]
    ].sort_values("Row ID")

    return {
        "Fact_Sales": fact,
        "Dim_Customer": dim_customer,
        "Dim_Product": dim_product,
        "Dim_Geography": dim_geography,
        "Dim_Date": dim_date,
        "Dim_ShipMode": dim_shipmode,
    }


# ---------------------------------------------------------------- Kiểm tra
PRIMARY_KEYS = {
    "Fact_Sales": "Row ID",
    "Dim_Customer": "Customer ID",
    "Dim_Product": "ProductKey",
    "Dim_Geography": "GeoKey",
    "Dim_Date": "DateKey",
    "Dim_ShipMode": "ShipModeKey",
}
FOREIGN_KEYS = [  # (cột trong Fact, bảng Dim, khóa Dim)
    ("OrderDateKey", "Dim_Date", "DateKey"),
    ("ShipDateKey", "Dim_Date", "DateKey"),
    ("Customer ID", "Dim_Customer", "Customer ID"),
    ("ProductKey", "Dim_Product", "ProductKey"),
    ("GeoKey", "Dim_Geography", "GeoKey"),
    ("ShipModeKey", "Dim_ShipMode", "ShipModeKey"),
]


def run_checks(t, flat):
    results = []

    def check(name, ok, detail):
        results.append(
            {
                "#": len(results) + 1,
                "Kiểm tra": name,
                "Kết quả": "Đạt" if ok else "KHÔNG ĐẠT",
                "Chi tiết": detail,
            }
        )

    for table, pk in PRIMARY_KEYS.items():
        col = t[table][pk]
        check(
            f"Khóa chính {table}.{pk} duy nhất, không null",
            col.is_unique and col.notna().all(),
            f"{len(col):,} dòng, {col.nunique():,} giá trị phân biệt",
        )

    fact = t["Fact_Sales"]
    for fk, dim, dk in FOREIGN_KEYS:
        orphans = int((~fact[fk].isin(t[dim][dk])).sum())
        check(
            f"Fact_Sales.{fk} → {dim}.{dk} không có dòng mồ côi",
            orphans == 0,
            f"{orphans} dòng mồ côi",
        )

    check(
        "Số dòng Fact = số dòng bản sạch",
        len(fact) == len(flat),
        f"{len(fact):,} / {len(flat):,}",
    )
    fact_total, flat_total = fact["Sales"].sum(), flat["Sales"].sum()
    check(
        "Tổng Sales Fact = tổng Sales bản sạch",
        abs(fact_total - flat_total) < 1e-6,
        f"{fact_total:,.2f} / {flat_total:,.2f}",
    )

    joined = (
        fact.merge(t["Dim_Product"], on="ProductKey")
        .merge(t["Dim_Geography"], on="GeoKey")
        .merge(t["Dim_Customer"], on="Customer ID")
    )
    mismatched = []
    for attr in ["Category", "Region", "Segment"]:
        a = joined.groupby(attr)["Sales"].sum()
        b = flat.groupby(attr)["Sales"].sum()
        if not ((a - b).abs() < 1e-6).all():
            mismatched.append(attr)
    check(
        "Tổng Sales theo Category, Region, Segment sau JOIN = bảng phẳng",
        len(joined) == len(fact) and not mismatched,
        f"{len(joined):,} dòng sau JOIN; lệch: {mismatched or 'không'}",
    )

    multi_region = int(
        (t["Dim_Geography"].groupby("State")["Region"].nunique() > 1).sum()
    )
    check(
        "Mỗi State chỉ thuộc một Region",
        multi_region == 0,
        f"{multi_region} State thuộc nhiều Region",
    )
    return results


def write_report(t, results):
    counts = [
        {"Bảng": name, "Số dòng": f"{len(df):,}", "Số cột": df.shape[1]}
        for name, df in t.items()
    ]
    n_ok = sum(r["Kết quả"] == "Đạt" for r in results)
    text = f"""# Kiểm tra star schema

> File này do `etl/build_star_schema.py` tự sinh. Không sửa tay, hãy chạy lại script.

Nguồn: `data/clean/Superstore_clean.csv` (dữ liệu 2015–2018). Kết quả: `data/star/`. Sơ đồ: `docs/Star_schema.md`.

## Số dòng từng bảng

{md_table(counts, ["Bảng", "Số dòng", "Số cột"])}

## Kiểm tra toàn vẹn ({n_ok}/{len(results)} đạt)

{md_table(results, ["#", "Kiểm tra", "Kết quả", "Chi tiết"])}

## Lưu ý khi nạp vào Power BI

- **Postal Code** phải để kiểu **Text**. Power Query có thể tự đổi sang số và làm mất số 0 đầu (`06824` → `6824`). Xóa bước "Changed Type" tự động hoặc đặt lại kiểu Text.
- Không mở các file CSV bằng Excel rồi lưu lại, vì Excel cũng làm mất số 0 đầu.
- `Dim_Date` có hai quan hệ với `Fact_Sales`: `OrderDateKey` là quan hệ **active**, `ShipDateKey` là **inactive**. Khi cần tính theo ngày giao, dùng `CALCULATE(..., USERELATIONSHIP(Fact_Sales[ShipDateKey], Dim_Date[DateKey]))`.
- `Dim_Date` chạy từ {DATE_START} đến {DATE_END}, vì Ship Date muộn nhất là 2019-01-05. Đánh dấu bảng này là Date table (Mark as date table) theo cột `Date`.
- Sắp xếp `Month Name` theo `Month` và `Weekday Name` theo `DayOfWeek` (Sort by column) để tháng và thứ hiển thị đúng thứ tự.
- `ShipModeKey` xếp theo tốc độ giao hàng: Same Day = 1, First Class = 2, Second Class = 3, Standard Class = 4.
- Mọi quan hệ là 1–nhiều, lọc một chiều từ Dim sang Fact.
"""
    REPORT.write_text(text, encoding="utf-8")


def main():
    sys.stdout.reconfigure(
        encoding="utf-8"
    )  # console Windows mặc định không phải UTF-8
    flat = pd.read_csv(
        CLEAN,
        dtype={"Postal Code": str},
        parse_dates=["Order Date", "Ship Date"],
        encoding="utf-8-sig",
    )
    tables = build_tables(flat)
    results = run_checks(tables, flat)
    for r in results:
        print(f"[{r['#']:>2}] {r['Kết quả']:<9} {r['Kiểm tra']} ({r['Chi tiết']})")
    failed = [r for r in results if r["Kết quả"] != "Đạt"]
    assert not failed, f"{len(failed)} kiểm tra không đạt"

    STAR.mkdir(parents=True, exist_ok=True)
    for name, df in tables.items():
        df.to_csv(
            STAR / f"{name}.csv",
            index=False,
            encoding="utf-8-sig",
            date_format="%Y-%m-%d",
        )
    write_report(tables, results)
    print(
        f"Đã ghi {len(tables)} bảng vào {STAR.relative_to(ROOT)} và {REPORT.relative_to(ROOT)}"
    )


if __name__ == "__main__":
    main()
