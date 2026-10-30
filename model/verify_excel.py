"""Đối chiếu kết quả Excel (sau apply_excel_tools.ps1) với lời giải scipy, xuất bảng kết quả.

Chạy: python -I model/verify_excel.py (sau solve_reference.py và apply_excel_tools.ps1)
Đầu ra: docs/PhanB_Ket_qua.md (tự sinh) — số liệu kết quả 4 bài toán cho báo cáo.
"""

import json
import sys
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
import params as P

XLSX = P.ROOT / "output" / "MaNhom_Model.xlsx"
REF = P.ROOT / "output" / "reference.json"
REPORT = P.ROOT / "docs" / "PhanB_Ket_qua.md"
TOL = 0.005  # sai số tương đối cho phép giữa Excel và scipy


def md_table(rows, columns):
    lines = ["| " + " | ".join(columns) + " |", "|" + "---|" * len(columns)]
    for r in rows:
        lines.append("| " + " | ".join(str(r[c]) for c in columns) + " |")
    return "\n".join(lines)


class Book:
    def __init__(self, path):
        self.wb = load_workbook(path, data_only=True)

    def range(self, name):
        dn = self.wb.defined_names[name]
        ((sheet, ref),) = list(dn.destinations)
        cells = self.wb[sheet][ref.replace("$", "")]
        if not isinstance(cells, tuple):
            return cells.value
        return [[c.value for c in row] for row in cells]

    def column(self, name):
        return [row[0] for row in self.range(name)]

    def sheet_rows(self, sheet):
        return [[c for c in row] for row in self.wb[sheet].iter_rows(values_only=True)]


checks = []


def check(label, excel, ref, tol=TOL):
    ok = excel is not None and abs(excel - ref) <= tol * max(1.0, abs(ref))
    checks.append(
        {
            "Kiểm tra": label,
            "Excel": f"{excel:,.4f}",
            "scipy": f"{ref:,.4f}",
            "Kết quả": "Khớp" if ok else "LỆCH",
        }
    )


def usd(v):
    return f"{v:,.0f}"


def pct(v):
    return f"{v:.1%}"


def main():
    sys.stdout.reconfigure(
        encoding="utf-8"
    )  # console Windows mặc định không phải UTF-8
    book = Book(XLSX)
    ref = json.loads(REF.read_text(encoding="utf-8"))
    base = P.load_base_year()

    # ---------------------------------------------------------- Bài 1
    r1 = ref["problem1"]
    check(
        "B1 lợi nhuận góp/đơn", book.range("B1_Contrib"), r1["contribution_per_order"]
    )
    check("B1 lợi nhuận 2019 (cơ sở)", book.range("B1_Profit"), r1["profit_2019"])
    check(
        "B1 Goal Seek – số đơn hòa vốn",
        book.range("B1_GS_BreakEven"),
        r1["gs_breakeven_orders"],
    )
    check(
        "B1 Goal Seek – số đơn đạt mục tiêu",
        book.range("B1_GS_TargetOrders"),
        r1["gs_target_orders"],
    )
    check(
        "B1 Goal Seek – AOV đạt mục tiêu",
        book.range("B1_GS_TargetAOV"),
        r1["gs_target_aov"],
    )

    # ---------------------------------------------------------- Bài 2
    r2 = ref["problem2"]["Ngân sách cơ sở"]
    check("B2 Solver – DT tăng thêm tối ưu", book.range("B2_Uplift"), r2["uplift"])
    for region, excel in zip(sorted(P.MKT_ALPHA), book.column("B2_Alloc")):
        check(
            f"B2 Solver – ngân sách {region}", excel, r2["allocation"][region], tol=0.02
        )
    check(
        "B2 Goal Seek – ngân sách chia đều đạt mục tiêu",
        book.range("B2_GS_EqBudget"),
        ref["problem2"]["gs_equal_split_budget"],
    )

    # ---------------------------------------------------------- Bài 3
    r3 = ref["problem3"]["Cơ sở"]
    check("B3 Solver – LN gộp tối ưu", book.range("B3_Profit"), r3["gross_profit"])
    check("B3 Solver – ngân sách đã dùng", book.range("B3_XSum"), r3["budget_used"])
    check("B3 Solver – thể tích đã dùng", book.range("B3_Space"), r3["space_used"])

    # ---------------------------------------------------------- Bài 4
    r4 = ref["problem4"]["Mục tiêu 30%"]
    check("B4 Solver – tổng chi phí tối thiểu", book.range("B4_Cost"), r4["total_cost"])
    for mode, excel in zip(P.SHIP_MODES, book.column("B4_X")):
        check(f"B4 Solver – số đơn {mode}", excel, r4["orders"][mode], tol=0.01)
    check(
        "B4 Goal Seek – tỷ trọng First Class",
        book.range("B4_GS_FCShare"),
        ref["problem4_gs_first_class_share"],
    )

    sheets = set(book.wb.sheetnames)
    expected = {
        "SC_B1",
        "SC_B2",
        "SC_B3",
        "SC_B4",
        "AR_B2",
        "SR_B2",
        "AR_B3",
        "SR_B3",
        "AR_B4",
        "SR_B4",
    }
    missing = expected - sheets
    failed = [c for c in checks if c["Kết quả"] != "Khớp"]
    for c in checks:
        print(
            f"{c['Kết quả']:<5} {c['Kiểm tra']}: Excel {c['Excel']} / scipy {c['scipy']}"
        )
    print("Sheet thiếu:", missing or "không")

    # ---------------------------------------------------------- Báo cáo
    b1_dt = book.range("B1_DT")
    dt1 = [
        {
            "Số đơn \\ AOV": f"{row[0]:,}",
            **{f"{h:,.0f}": usd(v) for h, v in zip(b1_dt[0][1:], row[1:])},
        }
        for row in b1_dt[1:]
    ]
    dt1_cols = ["Số đơn \\ AOV"] + [f"{h:,.0f}" for h in b1_dt[0][1:]]

    def dt_1d(name, labels, fmts):
        rows = book.range(name)[1:]
        return [{labels[i]: fmts[i](v) for i, v in enumerate(row)} for row in rows]

    sub_names = list(base["subcat"].index)
    b3_x = book.column("B3_X")
    b3_max = book.column("B3_Max")
    b3_min = book.column("B3_Min")
    b3_rows = [
        {
            "Sub-Category": s,
            "Biên LN": pct(P.SUBCAT_MARGIN[s]),
            "Nhập (giá vốn)": usd(x),
            "% nhu cầu": pct(x / mx),
            "Trạng thái": "Tối đa"
            if x >= mx - 1
            else ("Tối thiểu" if x <= mn + 1 else "Một phần"),
        }
        for s, x, mx, mn in zip(sub_names, b3_x, b3_max, b3_min)
    ]

    def scenario_rows(sheet):
        return [r for r in book.sheet_rows(sheet) if any(v is not None for v in r)]

    def scenario_md(sheet):
        rows = scenario_rows(sheet)
        out = []
        for r in rows:
            cells = [
                "" if v is None else (f"{v:,.2f}" if isinstance(v, float) else str(v))
                for v in r[1:]
            ]
            if any(cells):
                out.append("| " + " | ".join(cells) + " |")
        width = max(len(line.split("|")) for line in out) - 2
        out.insert(1, "|" + "---|" * width)
        return "\n".join(out)

    s4 = ref["problem4"]
    text = f"""# Phần B – Kết quả mô hình (tự sinh)

> File này do `model/verify_excel.py` sinh từ `output/MaNhom_Model.xlsx` sau khi Excel chạy Goal Seek, Data Table,
> Scenario Manager và Solver. Không sửa tay. Diễn giải và ý nghĩa quản lý xem `docs/PhanB_Mo_hinh.md`.

## Đối chiếu Excel ↔ scipy ({len(checks) - len(failed)}/{len(checks)} khớp, sai số ≤ {TOL:.1%})

{md_table(checks, ["Kiểm tra", "Excel", "scipy", "Kết quả"])}

## Bài 1 – Hòa vốn và lợi nhuận mục tiêu 2019

| Chỉ số | Giá trị |
|---|---|
| Biên LN gộp bình quân | {pct(book.range("B1_Margin"))} |
| Chi phí giao TB / đơn | {book.range("B1_Ship"):,.2f} USD |
| Lợi nhuận góp / đơn | {book.range("B1_Contrib"):,.2f} USD |
| Số đơn dự báo 2019 | {book.range("B1_Orders"):,} |
| Lợi nhuận 2019 (cơ sở) | {usd(book.range("B1_Profit"))} USD |
| Biên an toàn | {pct(book.range("B1_Safety"))} |
| **Goal Seek** – số đơn hòa vốn | **{book.range("B1_GS_BreakEven"):,.0f} đơn** |
| **Goal Seek** – số đơn đạt LN mục tiêu {usd(P.value(P.GENERAL, "profit_target"))} USD | **{book.range("B1_GS_TargetOrders"):,.0f} đơn** |
| **Goal Seek** – AOV tối thiểu đạt mục tiêu (với {book.range("B1_Orders"):,} đơn) | **{book.range("B1_GS_TargetAOV"):,.2f} USD** |

**Data Table 2 chiều – Lợi nhuận (USD) theo số đơn × AOV**

{md_table(dt1, dt1_cols)}

**Scenario Summary**

{scenario_md("SC_B1")}

## Bài 2 – Phân bổ ngân sách marketing

**Solver (GRG Nonlinear) – phương án cơ sở 100.000 USD**

{md_table([{"Region": r, "Ngân sách tối ưu": usd(b), "Ngân sách chia đều": usd(P.value(P.MKT, "budget") / 4)} for r, b in zip(sorted(P.MKT_ALPHA), book.column("B2_Alloc"))], ["Region", "Ngân sách tối ưu", "Ngân sách chia đều"])}

| Chỉ số | Giá trị |
|---|---|
| DT tăng thêm – phân bổ tối ưu | {usd(book.range("B2_Uplift"))} USD |
| DT tăng thêm – chia đều | {usd(ref["problem2"]["Ngân sách cơ sở"]["equal_split_uplift"])} USD |
| ROAS tối ưu | {book.range("B2_ROAS"):.2f} |
| **Goal Seek** – ngân sách chia đều cần để tăng {usd(P.value(P.MKT, "revenue_goal"))} USD DT | **{usd(book.range("B2_GS_EqBudget"))} USD** |

**Data Table – hiệu quả theo tổng ngân sách (chia đều)**

{md_table(dt_1d("B2_DT", ["Tổng ngân sách", "DT tăng thêm", "ROAS"], [usd, usd, lambda v: f"{v:.2f}"]), ["Tổng ngân sách", "DT tăng thêm", "ROAS"])}

**Scenario Summary (mỗi kịch bản là nghiệm Solver tối ưu)**

{scenario_md("SC_B2")}

## Bài 3 – Kế hoạch nhập hàng

| Chỉ số | Giá trị |
|---|---|
| LN gộp kế hoạch tối ưu | {usd(book.range("B3_Profit"))} USD |
| Ngân sách đã dùng | {usd(book.range("B3_XSum"))} / {usd(book.range("B3_Budget"))} USD ({pct(book.range("B3_BudgetUse"))}) |
| Kho đã dùng | {book.range("B3_Space"):,.1f} / {book.range("B3_Capacity"):,.0f} m³ ({pct(book.range("B3_SpaceUse"))}) |
| Tỷ lệ đáp ứng nhu cầu | {pct(book.range("B3_FillRate"))} |
| Giá trị biên (shadow price) của 1 USD ngân sách | {ref["problem3"]["Cơ sở"]["shadow_budget"]:.3f} USD LN |
| Giá trị biên của 1 m³ kho | {ref["problem3"]["Cơ sở"]["shadow_capacity"]:.1f} USD LN |

{md_table(b3_rows, ["Sub-Category", "Biên LN", "Nhập (giá vốn)", "% nhu cầu", "Trạng thái"])}

**Data Table – rủi ro khi nhu cầu thực tế lệch dự báo**

{md_table(dt_1d("B3_DT", ["Nhu cầu thực tế", "LN thực hiện", "Tồn kho dư"], [pct, usd, usd]), ["Nhu cầu thực tế", "LN thực hiện", "Tồn kho dư"])}

**Scenario Summary**

{scenario_md("SC_B3")}

## Bài 4 – Cơ cấu phương thức giao hàng

{md_table([{"Ship Mode": m, "Số đơn 2018": int(base["ship"].loc[m, "orders"]), "Số đơn 2019 tối ưu": f"{x:,.0f}", "Tỷ lệ giao nhanh của mode": pct(base["ship"].loc[m, "fast_rate"]), "Chi phí/đơn": P.SHIP_COST[m]} for m, x in zip(P.SHIP_MODES, book.column("B4_X"))], ["Ship Mode", "Số đơn 2018", "Số đơn 2019 tối ưu", "Tỷ lệ giao nhanh của mode", "Chi phí/đơn"])}

| Chỉ số | Giá trị |
|---|---|
| Tổng chi phí tối thiểu (mục tiêu 30%) | {usd(book.range("B4_Cost"))} USD |
| Chi phí nếu giữ cơ cấu 2018 (~22% giao nhanh) | {usd(book.range("B4_BaseCost"))} USD |
| Chi phí tăng thêm | {usd(book.range("B4_ExtraCost"))} USD |
| Tỷ lệ giao nhanh / ngày giao TB đạt được | {pct(book.range("B4_FastRate"))} / {book.range("B4_AvgDays"):.2f} ngày |
| Chi phí biên cho mỗi đơn giao nhanh thêm (shadow price) | {abs(s4["Mục tiêu 30%"]["shadow_fast_orders"]):.2f} USD |
| **Goal Seek** – tỷ trọng First Class cần (nếu chỉ tăng First Class) | **{pct(book.range("B4_GS_FCShare"))}** |

**Data Table – chỉ tăng First Class**

{md_table(dt_1d("B4_DT", ["Tỷ trọng First Class", "Tỷ lệ giao nhanh", "Tổng chi phí"], [pct, pct, usd]), ["Tỷ trọng First Class", "Tỷ lệ giao nhanh", "Tổng chi phí"])}

**Scenario Summary**

{scenario_md("SC_B4")}
"""
    REPORT.write_text(text, encoding="utf-8")
    print(f"Đã ghi {REPORT.relative_to(P.ROOT)}")
    assert not failed and not missing, "Có kiểm tra không đạt"


if __name__ == "__main__":
    main()
