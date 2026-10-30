"""Tiền xử lý dữ liệu Superstore (train.csv).

Chạy: python -I etl/preprocess.py
Đầu vào : data/raw/train.csv (bản gốc, chỉ đọc)
Đầu ra  : data/clean/Superstore_clean.csv, docs/Nhat_ky_tien_xu_ly.md
"""

import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw" / "train.csv"
CLEAN = ROOT / "data" / "clean" / "Superstore_clean.csv"
LOG = ROOT / "docs" / "Nhat_ky_tien_xu_ly.md"

DATE_COLS = ["Order Date", "Ship Date"]
CATEGORICAL_DOMAINS = {
    "Ship Mode": {"First Class", "Same Day", "Second Class", "Standard Class"},
    "Segment": {"Consumer", "Corporate", "Home Office"},
    "Region": {"Central", "East", "South", "West"},
    "Category": {"Furniture", "Office Supplies", "Technology"},
}
# Mã bưu chính USPS của Burlington, Vermont — dùng để điền Postal Code bị trống
BURLINGTON_VT_ZIP = "05401"

steps = []  # nhật ký: mỗi phần tử là một bước xử lý


def log_step(name, finding, action, before, after):
    steps.append(
        {
            "Bước": name,
            "Phát hiện": finding,
            "Cách xử lý": action,
            "Số dòng trước → sau": f"{before:,} → {after:,}",
        }
    )
    print(f"[{len(steps)}] {name}: {finding}")


def md_table(rows, columns):
    lines = ["| " + " | ".join(columns) + " |", "|" + "---|" * len(columns)]
    for r in rows:
        lines.append(
            "| " + " | ".join(str(r[c]).replace("|", "\\|") for c in columns) + " |"
        )
    return "\n".join(lines)


def check_one_to_one(df, key, attr):
    """Đếm số giá trị của `key` ứng với nhiều hơn một giá trị `attr`."""
    counts = df.groupby(key)[attr].nunique()
    return int((counts > 1).sum())


# ---------------------------------------------------------------- Giai đoạn 1
def profile(df):
    rows = []
    for c in df.columns:
        rows.append(
            {
                "Cột": c,
                "Kiểu": str(df[c].dtype),
                "Số ô trống": int(df[c].isna().sum()),
                "Số giá trị phân biệt": int(df[c].nunique()),
            }
        )
    return rows


def sales_stats(df):
    s = df["Sales"]
    return {
        "Min": f"{s.min():,.3f}",
        "Max": f"{s.max():,.2f}",
        "Mean": f"{s.mean():,.2f}",
        "Median": f"{s.median():,.2f}",
        "Tổng": f"{s.sum():,.2f}",
    }


# ---------------------------------------------------------------- Giai đoạn 2
def clean(raw):
    df = raw.copy()
    n = len(df)
    log_step(
        "Nạp dữ liệu",
        f"{n:,} dòng, {df.shape[1]} cột; mã hóa UTF-8; ngày dạng d/m/yyyy",
        "Đọc Postal Code dạng text để giữ số 0 đầu",
        n,
        n,
    )

    # Trùng lặp toàn bộ dòng (bỏ qua Row ID vì đây là khóa tự tăng)
    content_cols = [c for c in df.columns if c != "Row ID"]
    dup_mask = df.duplicated(content_cols, keep="first")
    dup_ids = df.loc[dup_mask, "Row ID"].tolist()
    removed_sales = df.loc[dup_mask, "Sales"].sum()
    df = df[~dup_mask]
    log_step(
        "Kiểm tra trùng toàn bộ dòng",
        f"{len(dup_ids)} dòng trùng hoàn toàn với dòng khác (Row ID {dup_ids})",
        "Xóa dòng trùng, giữ lại lần xuất hiện đầu tiên",
        n,
        len(df),
    )
    n = len(df)

    dup_rowid = int(df["Row ID"].duplicated().sum())
    log_step(
        "Kiểm tra trùng Row ID",
        f"{dup_rowid} Row ID bị trùng",
        "Không cần xử lý" if dup_rowid == 0 else "Cần xem xét",
        n,
        n,
    )

    pair_mask = df.duplicated(["Order ID", "Product ID"], keep=False)
    n_pairs = int(
        df.loc[pair_mask, ["Order ID", "Product ID"]].drop_duplicates().shape[0]
    )
    log_step(
        "Kiểm tra trùng Order ID + Product ID",
        f"{n_pairs} cặp (Order ID, Product ID) xuất hiện nhiều lần nhưng khác Sales",
        "Giữ nguyên: là các dòng hàng hợp lệ (cùng sản phẩm, mua thêm trong cùng đơn)",
        n,
        n,
    )

    # Sai định dạng: ngày
    for c in DATE_COLS:
        df[c] = pd.to_datetime(df[c], format="%d/%m/%Y")
    log_step(
        "Chuẩn hóa kiểu ngày",
        "Order Date, Ship Date ở dạng text d/m/yyyy",
        "Ép kiểu Date theo định dạng d/m/yyyy (Power Query: locale en-GB); xuất ra dạng yyyy-mm-dd",
        n,
        n,
    )

    # Sai định dạng: Postal Code mất số 0 đầu
    pc = df["Postal Code"]
    short = int((pc.notna() & (pc.str.len() < 5)).sum())
    df["Postal Code"] = pc.str.zfill(5)
    log_step(
        "Chuẩn hóa Postal Code",
        f"{short} dòng có mã bưu chính 4 chữ số do mất số 0 đầu (vd 6824)",
        "Đệm số 0 bên trái cho đủ 5 ký tự (6824 → 06824)",
        n,
        n,
    )

    # Dữ liệu trống
    na_mask = df["Postal Code"].isna()
    na_places = df.loc[na_mask, ["City", "State"]].drop_duplicates().values.tolist()
    assert na_places == [["Burlington", "Vermont"]], na_places
    df.loc[na_mask, "Postal Code"] = BURLINGTON_VT_ZIP
    log_step(
        "Xử lý ô trống",
        f"{int(na_mask.sum())} dòng trống Postal Code, tất cả thuộc Burlington, Vermont; "
        "các cột khác không có ô trống",
        f"Điền {BURLINGTON_VT_ZIP} (mã bưu chính USPS của Burlington, VT)",
        n,
        n,
    )

    # Ký tự lạ / khoảng trắng
    text_cols = df.select_dtypes("object").columns
    nbsp = int(df["Product Name"].str.contains(" ").sum())
    df["Product Name"] = df["Product Name"].str.replace(" ", " ", regex=False)
    untrimmed = 0
    for c in text_cols:
        stripped = df[c].str.strip().str.replace(r"\s{2,}", " ", regex=True)
        untrimmed += int((stripped != df[c]).sum())
        df[c] = stripped
    log_step(
        "Làm sạch text",
        f"{nbsp} dòng Product Name chứa ký tự NBSP (U+00A0); "
        f"{untrimmed} ô có khoảng trắng thừa",
        "Thay NBSP bằng dấu cách; Trim và gộp khoảng trắng liên tiếp ở mọi cột text",
        n,
        n,
    )

    # Không thống nhất / giá trị vô lý
    bad_domain = {
        c: sorted(set(df[c].unique()) - allowed)
        for c, allowed in CATEGORICAL_DOMAINS.items()
    }
    bad_domain = {c: v for c, v in bad_domain.items() if v}
    ship_before = int((df["Ship Date"] < df["Order Date"]).sum())
    sales_nonpos = int((df["Sales"] <= 0).sum())
    log_step(
        "Kiểm tra tính thống nhất",
        f"Giá trị lạ ở cột phân loại: {bad_domain or 'không có'}; "
        f"Ship Date < Order Date: {ship_before}; Sales ≤ 0: {sales_nonpos}",
        "Không phát hiện bản ghi vô lý, không cần loại bỏ"
        if not (bad_domain or ship_before or sales_nonpos)
        else "Cần xem xét",
        n,
        n,
    )
    return df, removed_sales


# ---------------------------------------------------------------- Giai đoạn 3
CONSISTENCY_CHECKS = [
    ("Customer ID", "Customer Name"),
    ("Customer ID", "Segment"),
    ("Product ID", "Product Name"),
    ("Product Name", "Product ID"),
    ("Product ID", "Sub-Category"),
    ("Postal Code", "City"),
    ("Postal Code", "State"),
    ("City", "State"),
    ("Order ID", "Order Date"),
    ("Order ID", "Ship Date"),
    ("Order ID", "Ship Mode"),
    ("Order ID", "Customer ID"),
]


def consistency(df):
    rows = []
    for key, attr in CONSISTENCY_CHECKS:
        bad = check_one_to_one(df, key, attr)
        rows.append(
            {
                "Kiểm tra": f"Mỗi {key} → một {attr}",
                "Số vi phạm": bad,
                "Kết luận": "Nhất quán"
                if bad == 0
                else "Không nhất quán → cần khóa thay thế",
            }
        )
    return rows


# ---------------------------------------------------------------- Giai đoạn 4
def surrogate_key(df, cols):
    keys = df[cols].drop_duplicates().sort_values(cols).reset_index(drop=True)
    keys["_key"] = range(1, len(keys) + 1)
    return df.merge(keys, on=cols, how="left")["_key"].to_numpy()


def enrich(df):
    n = len(df)
    df["Order Year"] = df["Order Date"].dt.year
    df["Order Quarter"] = df["Order Date"].dt.quarter
    df["Order Month"] = df["Order Date"].dt.month
    df["Order YearMonth"] = df["Order Date"].dt.strftime("%Y-%m")
    df["Ship Days"] = (df["Ship Date"] - df["Order Date"]).dt.days

    q1, q3 = df["Sales"].quantile([0.25, 0.75])
    upper = q3 + 1.5 * (q3 - q1)
    df["Is Sales Outlier"] = (df["Sales"] > upper).astype(int)

    df["ProductKey"] = surrogate_key(df, ["Product ID", "Product Name"])
    df["GeoKey"] = surrogate_key(df, ["City", "State", "Postal Code"])
    log_step(
        "Tạo cột bổ sung",
        "Cần thuộc tính thời gian, thời gian giao hàng, cờ ngoại lai và khóa thay thế cho star schema",
        "Thêm Order Year/Quarter/Month/YearMonth, Ship Days, "
        f"Is Sales Outlier (Sales > Q3 + 1.5·IQR = {upper:,.2f}; "
        f"{int(df['Is Sales Outlier'].sum()):,} dòng, giữ lại không xóa), "
        f"ProductKey ({df['ProductKey'].nunique():,} sản phẩm), "
        f"GeoKey ({df['GeoKey'].nunique():,} địa điểm)",
        n,
        n,
    )
    return df


# ---------------------------------------------------------------- Giai đoạn 5
def validate(raw, df, removed_sales):
    assert len(df) == len(raw) - 1, len(df)
    assert df.isna().sum().sum() == 0, "Còn ô trống"
    assert (df["Postal Code"].str.len() == 5).all(), "Postal Code chưa đủ 5 ký tự"
    assert df["Row ID"].is_unique
    assert df["Ship Days"].between(0, 7).all()
    assert abs(df["Sales"].sum() - (raw["Sales"].sum() - removed_sales)) < 1e-6
    assert not df["Product Name"].str.contains(" ").any()
    print("Tất cả kiểm tra hậu xử lý đều đạt.")


def write_log(raw, df, raw_profile, consistency_rows, removed_sales):
    raw_stats, clean_stats = sales_stats(raw), sales_stats(df)
    recon = [
        {"Chỉ số": "Số dòng", "Trước": f"{len(raw):,}", "Sau": f"{len(df):,}"},
        {"Chỉ số": "Số cột", "Trước": raw.shape[1], "Sau": df.shape[1]},
        {
            "Chỉ số": "Số đơn hàng (Order ID phân biệt)",
            "Trước": f"{raw['Order ID'].nunique():,}",
            "Sau": f"{df['Order ID'].nunique():,}",
        },
        {
            "Chỉ số": "Số ô trống",
            "Trước": int(raw.isna().sum().sum()),
            "Sau": int(df.isna().sum().sum()),
        },
    ] + [
        {"Chỉ số": f"Sales – {k}", "Trước": raw_stats[k], "Sau": clean_stats[k]}
        for k in raw_stats
    ]

    step_rows = [{"#": i, **s} for i, s in enumerate(steps, 1)]
    text = f"""# Nhật ký tiền xử lý dữ liệu Superstore

> File này do `etl/preprocess.py` tự sinh. Không sửa tay, hãy chạy lại script.

- Nguồn: `data/raw/train.csv` (Superstore Sales, Kaggle), giữ nguyên bản gốc ở chế độ chỉ đọc.
- Kết quả: `data/clean/Superstore_clean.csv`.
- Nguyên tắc: chỉ xử lý lỗi có thật trong dữ liệu, không chèn lỗi giả lập.

## 1. Khảo sát dữ liệu gốc

{md_table(raw_profile, ["Cột", "Kiểu", "Số ô trống", "Số giá trị phân biệt"])}

Dataset chỉ có một cột định lượng là **Sales**. Không có Quantity, Discount, Profit.

## 2. Các bước làm sạch và làm giàu dữ liệu

{md_table(step_rows, ["#", "Bước", "Phát hiện", "Cách xử lý", "Số dòng trước → sau"])}

## 3. Kiểm tra tính nhất quán cho star schema

{md_table(consistency_rows, ["Kiểm tra", "Số vi phạm", "Kết luận"])}

Quyết định thiết kế khóa:
- **DimCustomer**: khóa `Customer ID`.
- **DimProduct**: khóa thay thế `ProductKey` trên cặp (Product ID, Product Name), vì Product ID không xác định duy nhất tên sản phẩm.
- **DimGeography**: khóa thay thế `GeoKey` trên (City, State, Postal Code), vì một mã bưu chính có thể ứng với nhiều thành phố và một tên thành phố có thể thuộc nhiều bang. Bỏ cột Country vì chỉ có một giá trị (United States).
- **Order ID**: dùng làm degenerate dimension trong bảng fact, vì mỗi đơn có đúng một ngày đặt, ngày giao, Ship Mode và khách hàng.

## 4. Đối chiếu trước / sau

{md_table(recon, ["Chỉ số", "Trước", "Sau"])}

Tổng Sales giảm đúng **{removed_sales:,.3f}**, bằng giá trị của dòng trùng đã xóa.
"""
    LOG.write_text(text, encoding="utf-8")


def main():
    sys.stdout.reconfigure(
        encoding="utf-8"
    )  # console Windows mặc định không phải UTF-8
    raw = pd.read_csv(RAW, dtype={"Postal Code": str}, encoding="utf-8")
    raw_profile = profile(raw)
    df, removed_sales = clean(raw)
    consistency_rows = consistency(df)
    df = enrich(df)
    validate(raw, df, removed_sales)

    CLEAN.parent.mkdir(parents=True, exist_ok=True)
    LOG.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(CLEAN, index=False, encoding="utf-8-sig", date_format="%Y-%m-%d")
    write_log(raw, df, raw_profile, consistency_rows, removed_sales)
    print(
        f"Đã ghi {CLEAN.relative_to(ROOT)} ({len(df):,} dòng) và {LOG.relative_to(ROOT)}"
    )


if __name__ == "__main__":
    main()
