"""Dựng file Excel Phần B (Model-Driven DSS): 4 bài toán ra quyết định.

Chạy: python -I model/build_model.py
Đầu ra:
  output/PhanB_Dap_an_tham_khao.xlsx  — workbook có công thức, named range, khung Data Table/Goal Seek
  output/excel_tools.json   — cấu hình cho model/apply_excel_tools.ps1 (chạy Goal Seek,
                              Data Table, Scenario Manager, Solver bằng Excel thật)
"""

import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

sys.path.insert(0, str(Path(__file__).resolve().parent))
import params as P

OUT_XLSX = P.ROOT / "output" / "PhanB_Dap_an_tham_khao.xlsx"
OUT_TOOLS = P.ROOT / "output" / "excel_tools.json"

# ------------------------------------------------------------------ Định dạng
FILL = {
    "input": PatternFill("solid", fgColor="FFF2CC"),  # tham số nhập (vàng)
    "decision": PatternFill("solid", fgColor="C6EFCE"),  # biến quyết định (xanh lá)
    "output": PatternFill("solid", fgColor="DDEBF7"),  # kết quả/mục tiêu (xanh dương)
    "header": PatternFill("solid", fgColor="1F4E78"),
    "section": PatternFill("solid", fgColor="D9E1F2"),
}
FMT = {"usd": "#,##0", "usd2": "#,##0.00", "pct": "0.0%", "num": "#,##0", "dec": "0.00"}
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


class SheetBuilder:
    """Ghi tuần tự theo dòng: tiêu đề, khối tham số, bảng; tạo named range."""

    def __init__(self, wb, name, widths):
        self.wb = wb
        self.ws = wb.create_sheet(name)
        self.name = name
        self.row = 1
        for i, w in enumerate(widths, 1):
            self.ws.column_dimensions[get_column_letter(i)].width = w

    def ref(self, col, row):
        # Luôn đặt tên sheet trong nháy: "B1_Hoa_von" bắt đầu giống địa chỉ ô B1
        return f"'{self.name}'!${get_column_letter(col)}${row}"

    def define(self, name, col, row, col2=None, row2=None):
        ref = self.ref(col, row)
        if col2:
            ref += f":${get_column_letter(col2)}${row2}"
        self.wb.defined_names[name] = DefinedName(name, attr_text=ref)
        return name

    def cell(self, col, row, value, fmt=None, kind=None, bold=False):
        c = self.ws.cell(row=row, column=col, value=value)
        if fmt:
            c.number_format = FMT[fmt]
        if kind:
            c.fill = FILL[kind]
            c.border = BORDER
        if bold:
            c.font = Font(bold=True)
        return c

    def title(self, text, subtitle):
        self.ws.cell(row=1, column=1, value=text).font = Font(
            bold=True, size=14, color="1F4E78"
        )
        self.ws.cell(row=2, column=1, value=subtitle).alignment = Alignment(
            wrap_text=False
        )
        self.ws.cell(row=2, column=1).font = Font(italic=True, color="595959")
        self.row = 4

    def section(self, text, ncols=6):
        for col in range(1, ncols + 1):
            self.ws.cell(row=self.row, column=col).fill = FILL["section"]
        self.ws.cell(row=self.row, column=1, value=text).font = Font(bold=True)
        self.row += 1

    def header(self, labels, col=1):
        for i, label in enumerate(labels):
            c = self.ws.cell(row=self.row, column=col + i, value=label)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = FILL["header"]
            c.alignment = Alignment(wrap_text=True, vertical="center")
            c.border = BORDER
        self.row += 1

    def param(self, label, value, fmt, name=None, kind="input", unit="", note=""):
        r = self.row
        self.cell(1, r, label)
        self.cell(2, r, value, fmt, kind)
        self.cell(3, r, unit)
        note_cell = self.cell(4, r, note)
        note_cell.data_type = "s"  # ghi chú luôn là chữ, kể cả khi bắt đầu bằng "="
        note_cell.font = Font(italic=True, color="595959")
        if name:
            self.define(name, 2, r)
        self.row += 1
        return r

    def skip(self, n=1):
        self.row += n


def data_table_block(
    sb,
    title,
    corner_formulas,
    col_values,
    fmt_in,
    fmt_out,
    labels,
    name,
    row_values=None,
):
    """Khung Data Table: cột giá trị đầu vào bên trái, công thức ở hàng đầu.

    1 chiều: corner_formulas = danh sách công thức kết quả (mỗi cột một kết quả).
    2 chiều: row_values khác None, corner_formulas = [một công thức].
    """
    sb.section(title)
    sb.header(labels)
    top = sb.row
    if row_values is None:
        sb.cell(1, top, "(công thức →)").font = Font(italic=True, color="808080")
        for j, f in enumerate(corner_formulas):
            sb.cell(2 + j, top, f, fmt_out[j] if isinstance(fmt_out, list) else fmt_out)
        width = len(corner_formulas)
    else:
        sb.cell(1, top, corner_formulas[0], fmt_out)
        for j, v in enumerate(row_values):
            sb.cell(2 + j, top, v, fmt_in[1], "input")
        width = len(row_values)
    for i, v in enumerate(col_values, 1):
        sb.cell(1, top + i, v, fmt_in if row_values is None else fmt_in[0], "input")
        for j in range(width):
            sb.cell(
                2 + j,
                top + i,
                None,
                fmt_out[j] if isinstance(fmt_out, list) else fmt_out,
                "output",
            )
    sb.define(name, 1, top, 1 + width, top + len(col_values))
    sb.row = top + len(col_values) + 2


def goal_seek_block(sb, rows):
    """Bảng kết quả Goal Seek; cột 'Kết quả' do apply_excel_tools.ps1 điền."""
    sb.section("KẾT QUẢ GOAL SEEK (Data → What-If Analysis → Goal Seek)")
    sb.header(["Câu hỏi", "Set cell", "To value", "By changing", "Kết quả", "Đơn vị"])
    out = []
    for q, set_name, target, change_name, unit, fmt, result_name in rows:
        r = sb.row
        sb.cell(1, r, q)
        sb.cell(2, r, set_name)
        sb.cell(3, r, target, fmt if unit != "%" else "pct")
        sb.cell(4, r, change_name)
        sb.cell(
            5,
            r,
            None,
            "usd2" if fmt == "usd2" else ("pct" if unit == "%" else "num"),
            "output",
        )
        sb.cell(6, r, unit)
        sb.define(result_name, 5, r)
        out.append(
            {
                "set": set_name,
                "target": target,
                "change": change_name,
                "result": result_name,
            }
        )
        sb.row += 1
    sb.skip()
    return out


# ------------------------------------------------------------------ Các sheet
def sheet_guide(wb):
    ws = wb.active
    ws.title = "Huong_dan"
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 110
    rows = [
        ("MÔ HÌNH RA QUYẾT ĐỊNH – SUPERSTORE 2019", None),
        ("", None),
        ("Sheet", "Nội dung"),
        (
            "Du_lieu_2018",
            "Số liệu thật năm 2018 trích từ star schema (data/star). Không sửa.",
        ),
        (
            "Gia_dinh",
            "Giả định dùng chung: tăng trưởng, phí, chi phí giao hàng, biên lợi nhuận và thể tích theo Sub-Category.",
        ),
        (
            "B1_Hoa_von",
            "Bài 1 – Hòa vốn & lợi nhuận mục tiêu 2019: Goal Seek, Data Table 2 chiều, Scenario.",
        ),
        (
            "B2_Marketing",
            "Bài 2 – Phân bổ ngân sách marketing theo Region: Solver (GRG Nonlinear), Goal Seek, Data Table, Scenario.",
        ),
        (
            "B3_Nhap_hang",
            "Bài 3 – Kế hoạch nhập hàng theo Sub-Category: Solver (Simplex LP) + Sensitivity, Data Table, Scenario.",
        ),
        (
            "B4_Giao_hang",
            "Bài 4 – Cơ cấu phương thức giao hàng: Solver (Simplex LP) + Sensitivity, Goal Seek, Data Table, Scenario.",
        ),
        (
            "SC_*, AR_*, SR_*",
            "Scenario Summary / Answer Report / Sensitivity Report do Excel sinh ra.",
        ),
        ("", None),
        ("Màu ô", "Ý nghĩa"),
        ("Vàng", "Tham số đầu vào – được phép thay đổi để phân tích What-if."),
        ("Xanh lá", "Biến quyết định – Solver thay đổi các ô này."),
        ("Xanh dương", "Kết quả / hàm mục tiêu – công thức, không nhập tay."),
        ("", None),
        (
            "Lưu ý",
            "Mọi chi phí và biên lợi nhuận là GIẢ ĐỊNH (dataset chỉ có doanh thu). Xem lý do ở cột Ghi chú.",
        ),
        (
            "",
            "Mô hình Solver của mỗi sheet đã được lưu sẵn: mở sheet → Data → Solver → Solve.",
        ),
        (
            "",
            "Cần bật add-in Solver: File → Options → Add-ins → Excel Add-ins → Go → chọn Solver Add-in.",
        ),
    ]
    for i, (a, b) in enumerate(rows, 1):
        ws.cell(row=i, column=1, value=a)
        ws.cell(row=i, column=2, value=b)
    ws["A1"].font = Font(bold=True, size=14, color="1F4E78")
    for r in (3, 12):
        for c in (1, 2):
            ws.cell(row=r, column=c).font = Font(bold=True, color="FFFFFF")
            ws.cell(row=r, column=c).fill = FILL["header"]
    for r, kind in ((13, "input"), (14, "decision"), (15, "output")):
        ws.cell(row=r, column=1).fill = FILL[kind]


def sheet_data(wb, base):
    sb = SheetBuilder(wb, "Du_lieu_2018", [24, 16, 16, 16, 14])
    sb.title(
        "SỐ LIỆU THỰC TẾ NĂM 2018",
        "Nguồn: data/star/*.csv (Superstore, sau tiền xử lý). Giá trị cố định, không sửa.",
    )
    sb.section("TỔNG QUAN")
    sb.param(
        "Doanh thu 2018",
        round(base["revenue"], 2),
        "usd2",
        "D_Revenue",
        kind="output",
        unit="USD",
    )
    sb.param(
        "Số đơn hàng 2018", base["orders"], "num", "D_Orders", kind="output", unit="đơn"
    )
    sb.param(
        "Số đơn hàng 2017",
        base["orders_prev"],
        "num",
        "D_OrdersPrev",
        kind="output",
        unit="đơn",
    )
    sb.param(
        "AOV 2018",
        round(base["aov"], 2),
        "usd2",
        "D_AOV",
        kind="output",
        unit="USD/đơn",
    )
    sb.skip()

    sb.section("DOANH THU THEO REGION")
    sb.header(["Region", "DT 2017", "DT 2018", "Tăng trưởng", "Tỷ trọng 2018"])
    top = sb.row
    regions = sorted(P.MKT_ALPHA)
    for i, r in enumerate(regions):
        row = top + i
        sb.cell(1, row, r)
        sb.cell(2, row, round(base["region_prev"][r], 2), "usd")
        sb.cell(3, row, round(base["region"][r], 2), "usd")
        sb.cell(4, row, f"=C{row}/B{row}-1", "pct")
        sb.cell(5, row, f"=C{row}/SUM($C${top}:$C${top + 3})", "pct")
    sb.define("D_RegionRevenue", 3, top, 3, top + 3)
    sb.define("D_RegionGrowth", 4, top, 4, top + 3)
    sb.row = top + 5

    sb.section("DOANH THU THEO SUB-CATEGORY (sắp theo doanh thu 2018)")
    sb.header(["Sub-Category", "Category", "DT 2018"])
    top = sb.row
    sub = base["subcat"]
    for i, (s, row_) in enumerate(sub.iterrows()):
        sb.cell(1, top + i, s)
        sb.cell(2, top + i, row_["category"])
        sb.cell(3, top + i, round(row_["revenue"], 2), "usd")
    sb.define("D_SubRevenue", 3, top, 3, top + len(sub) - 1)
    sb.row = top + len(sub) + 1

    sb.section("GIAO HÀNG THEO SHIP MODE (tính trên đơn hàng)")
    sb.header(["Ship Mode", "Số đơn", "Ngày giao TB", "Tỷ lệ giao ≤2 ngày"])
    top = sb.row
    for i, m in enumerate(P.SHIP_MODES):
        s = base["ship"].loc[m]
        sb.cell(1, top + i, m)
        sb.cell(2, top + i, int(s["orders"]), "num")
        sb.cell(3, top + i, round(s["avg_days"], 4), "dec")
        sb.cell(4, top + i, round(s["fast_rate"], 4), "pct")
    sb.define("D_ShipOrders", 2, top, 2, top + 3)
    sb.define("D_ShipDays", 3, top, 3, top + 3)
    sb.define("D_ShipFast", 4, top, 4, top + 3)


def sheet_assumptions(wb, base):
    sb = SheetBuilder(wb, "Gia_dinh", [34, 14, 14, 90])
    sb.title(
        "GIẢ ĐỊNH DÙNG CHUNG",
        "Dataset chỉ có doanh thu → các tham số chi phí/biên lợi nhuận dưới đây là giả định của nhóm.",
    )
    sb.section("THAM SỐ CHUNG")
    sb.header(["Tham số", "Giá trị", "Đơn vị", "Lý do"])
    g = P.GENERAL
    sb.param(
        "Tăng trưởng số đơn 2019",
        g["order_growth"][0],
        "pct",
        "G_OrderGrowth",
        unit=g["order_growth"][1],
        note=g["order_growth"][2],
    )
    sb.param(
        "Phí thanh toán",
        g["payment_fee"][0],
        "pct",
        "G_Fee",
        unit=g["payment_fee"][1],
        note=g["payment_fee"][2],
    )
    sb.param(
        "Chi phí xử lý đơn",
        g["handling_cost"][0],
        "usd2",
        "G_Handling",
        unit=g["handling_cost"][1],
        note=g["handling_cost"][2],
    )
    sb.skip()

    sb.section("CHI PHÍ GIAO HÀNG THEO SHIP MODE")
    sb.header(["Ship Mode", "USD/đơn", "", "Lý do"])
    top = sb.row
    for i, m in enumerate(P.SHIP_MODES):
        sb.cell(1, top + i, m)
        sb.cell(2, top + i, P.SHIP_COST[m], "usd2", "input")
        sb.cell(
            4,
            top + i,
            "Giả định theo mặt bằng phí vận chuyển nội địa Mỹ; giao càng nhanh càng đắt",
        ).font = Font(italic=True, color="595959")
    sb.define("G_ShipCost", 2, top, 2, top + 3)
    sb.row = top + 5

    sb.section("BIÊN LỢI NHUẬN GỘP VÀ THỂ TÍCH LƯU KHO THEO SUB-CATEGORY")
    sb.header(["Sub-Category", "Biên LN gộp", "m³ / 1.000 USD", "Lý do"])
    top = sb.row
    for i, s in enumerate(base["subcat"].index):
        sb.cell(1, top + i, s)
        sb.cell(2, top + i, P.SUBCAT_MARGIN[s], "pct", "input")
        sb.cell(3, top + i, P.SUBCAT_VOLUME[s], "dec", "input")
    notes = {
        top: "Biên: giấy, nhãn, bìa hồ sơ biên cao; bàn, kệ sách, máy móc cạnh tranh giá nên biên thấp.",
        top
        + 1: "Thể tích: nội thất (bàn, ghế, kệ) cồng kềnh; điện thoại, phụ kiện, nhãn nhỏ gọn.",
    }
    for r, n in notes.items():
        sb.cell(4, r, n).font = Font(italic=True, color="595959")
    sb.define("G_SubMargin", 2, top, 2, top + len(base["subcat"]) - 1)
    sb.define("G_SubVolume", 3, top, 3, top + len(base["subcat"]) - 1)


def sheet_b1(wb, base, tools):
    sb = SheetBuilder(wb, "B1_Hoa_von", [44, 16, 16, 16, 16, 16, 16, 16])
    sb.title(
        "BÀI 1 – HÒA VỐN VÀ LỢI NHUẬN MỤC TIÊU NĂM 2019",
        "Cần bán bao nhiêu đơn (hoặc AOV bao nhiêu) để hòa vốn và đạt lợi nhuận mục tiêu?",
    )
    g = P.GENERAL
    sb.section("THAM SỐ VÀ BIẾN QUYẾT ĐỊNH")
    sb.header(["Tham số", "Giá trị", "Đơn vị", "Ghi chú"])
    sb.param(
        "Số đơn hàng 2019 (biến quyết định)",
        P.orders_2019(base),
        "num",
        "B1_Orders",
        "decision",
        "đơn",
        "Số đơn 2018 × (1 + tăng trưởng 20%)",
    )
    sb.param(
        "Giá trị đơn trung bình – AOV (biến quyết định)",
        round(base["aov"], 2),
        "usd2",
        "B1_AOV",
        "decision",
        "USD/đơn",
        "AOV thực tế 2018",
    )
    sb.param(
        "Biên lợi nhuận gộp bình quân",
        "=SUMPRODUCT(D_SubRevenue,G_SubMargin)/SUM(D_SubRevenue)",
        "pct",
        "B1_Margin",
        "output",
        "%",
        "Bình quân gia quyền theo cơ cấu doanh thu 2018",
    )
    sb.param("Phí thanh toán", "=G_Fee", "pct", "B1_Fee", "output", "% DT")
    sb.param(
        "Chi phí xử lý / đơn", "=G_Handling", "usd2", "B1_Handling", "output", "USD/đơn"
    )
    sb.param(
        "Chi phí giao hàng TB / đơn",
        "=SUMPRODUCT(D_ShipOrders,G_ShipCost)/SUM(D_ShipOrders)",
        "usd2",
        "B1_Ship",
        "output",
        "USD/đơn",
        "Theo cơ cấu Ship Mode 2018 (xem Bài 4)",
    )
    sb.param(
        "Chi phí cố định năm",
        g["fixed_cost"][0],
        "usd",
        "B1_Fixed",
        "input",
        "USD",
        g["fixed_cost"][2],
    )
    sb.param(
        "Lợi nhuận mục tiêu",
        g["profit_target"][0],
        "usd",
        "B1_Target",
        "input",
        "USD",
        g["profit_target"][2],
    )
    sb.skip()

    sb.section("KẾT QUẢ")
    sb.param(
        "Doanh thu 2019", "=B1_Orders*B1_AOV", "usd", "B1_Revenue", "output", "USD"
    )
    sb.param(
        "Lợi nhuận góp / đơn",
        "=B1_AOV*(B1_Margin-B1_Fee)-B1_Handling-B1_Ship",
        "usd2",
        "B1_Contrib",
        "output",
        "USD/đơn",
        "AOV × (biên gộp − phí) − chi phí xử lý − chi phí giao",
    )
    sb.param(
        "Tổng lợi nhuận góp",
        "=B1_Orders*B1_Contrib",
        "usd",
        "B1_TotalContrib",
        "output",
        "USD",
    )
    sb.param(
        "LỢI NHUẬN 2019",
        "=B1_TotalContrib-B1_Fixed",
        "usd",
        "B1_Profit",
        "output",
        "USD",
        "Biến kết quả",
    )
    sb.param(
        "Số đơn hòa vốn (công thức)",
        "=B1_Fixed/B1_Contrib",
        "num",
        "B1_BreakEven",
        "output",
        "đơn",
        "Dùng để kiểm chứng kết quả Goal Seek",
    )
    sb.param(
        "Biên an toàn",
        "=(B1_Orders-B1_BreakEven)/B1_Orders",
        "pct",
        "B1_Safety",
        "output",
        "%",
        "Số đơn có thể giảm trước khi lỗ",
    )
    sb.skip()

    tools["goal_seek"] += goal_seek_block(
        sb,
        [
            (
                "Cần bao nhiêu đơn để hòa vốn?",
                "B1_Profit",
                0,
                "B1_Orders",
                "đơn",
                "num",
                "B1_GS_BreakEven",
            ),
            (
                "Cần bao nhiêu đơn để đạt lợi nhuận mục tiêu?",
                "B1_Profit",
                g["profit_target"][0],
                "B1_Orders",
                "đơn",
                "num",
                "B1_GS_TargetOrders",
            ),
            (
                "Với số đơn dự báo, AOV tối thiểu để đạt mục tiêu?",
                "B1_Profit",
                g["profit_target"][0],
                "B1_AOV",
                "USD/đơn",
                "usd2",
                "B1_GS_TargetAOV",
            ),
        ],
    )

    aov_values = list(range(360, 501, 20))
    order_values = list(range(1200, 2401, 200))
    data_table_block(
        sb,
        "DATA TABLE 2 CHIỀU – LỢI NHUẬN THEO SỐ ĐƠN (dọc) × AOV (ngang)",
        ["=B1_Profit"],
        order_values,
        ["num", "usd"],
        "usd",
        ["Số đơn \\ AOV"] + [""] * len(aov_values),
        "B1_DT",
        row_values=aov_values,
    )
    tools["data_tables"].append(
        {"range": "B1_DT", "row_input": "B1_AOV", "col_input": "B1_Orders"}
    )

    tools["scenarios"].append(
        {
            "sheet": "B1_Hoa_von",
            "summary_name": "SC_B1",
            "base": "Cơ sở",
            "changing": ["B1_Orders", "B1_AOV", "B1_Fixed"],
            "results": [
                "B1_Revenue",
                "B1_Contrib",
                "B1_Profit",
                "B1_BreakEven",
                "B1_Safety",
            ],
            "items": [
                {
                    "name": "Bi quan",
                    "values": [base["orders"], 400.0, 75_000.0],
                    "comment": "Số đơn không tăng, AOV giảm, chi phí cố định tăng",
                },
                {
                    "name": "Cơ sở",
                    "values": [
                        P.orders_2019(base),
                        round(base["aov"], 2),
                        g["fixed_cost"][0],
                    ],
                    "comment": "Số đơn +20%, AOV như 2018",
                },
                {
                    "name": "Lạc quan",
                    "values": [round(base["orders"] * 1.3), 450.0, 68_000.0],
                    "comment": "Số đơn +30%, AOV tăng, tối ưu chi phí cố định",
                },
            ],
        }
    )


def sheet_b2(wb, base, tools):
    sb = SheetBuilder(wb, "B2_Marketing", [36, 16, 14, 16, 18, 14, 18, 14])
    sb.title(
        "BÀI 2 – PHÂN BỔ NGÂN SÁCH MARKETING THEO REGION",
        "Chia ngân sách cho 4 vùng để tối đa doanh thu tăng thêm, với lợi suất giảm dần theo mức chi.",
    )
    m = P.MKT
    sb.section("THAM SỐ")
    sb.header(["Tham số", "Giá trị", "Đơn vị", "Ghi chú"])
    sb.param(
        "Tổng ngân sách marketing",
        m["budget"][0],
        "usd",
        "B2_Budget",
        unit="USD",
        note=m["budget"][2],
    )
    sb.param(
        "Ngân sách tối thiểu mỗi vùng",
        m["min_region"][0],
        "usd",
        "B2_Min",
        unit="USD",
        note=m["min_region"][2],
    )
    sb.param(
        "Ngân sách tối đa mỗi vùng",
        m["max_region"][0],
        "usd",
        "B2_Max",
        unit="USD",
        note=m["max_region"][2],
    )
    sb.param(
        "Hệ số bão hòa (scale)",
        m["scale"][0],
        "usd",
        "B2_Scale",
        unit="USD",
        note=m["scale"][2],
    )
    sb.param(
        "Mục tiêu DT tăng thêm",
        m["revenue_goal"][0],
        "usd",
        "B2_Goal",
        unit="USD",
        note=m["revenue_goal"][2],
    )
    sb.skip()

    sb.section("MÔ HÌNH: DT tăng thêm = DT2018 × α × (1 − EXP(−Ngân sách / scale))", 8)
    sb.header(
        [
            "Region",
            "DT 2018",
            "α (tiềm năng)",
            "Ngân sách",
            "DT tăng thêm",
            "ROAS",
            "DT biên / 1 USD",
            "Tăng trưởng 17→18",
        ]
    )
    top = sb.row
    regions = sorted(P.MKT_ALPHA)
    for i, r in enumerate(regions):
        row = top + i
        sb.cell(1, row, r)
        sb.cell(2, row, f"=INDEX(D_RegionRevenue,{i + 1})", "usd", "output")
        sb.cell(3, row, P.MKT_ALPHA[r], "pct", "input")
        sb.cell(4, row, m["budget"][0] / 4, "usd", "decision")
        sb.define(f"B2_NS_{r}", 4, row)  # tên từng ô để Scenario Summary dễ đọc
        sb.cell(5, row, f"=B{row}*C{row}*(1-EXP(-D{row}/B2_Scale))", "usd", "output")
        sb.cell(6, row, f"=E{row}/D{row}", "dec", "output")
        sb.cell(
            7, row, f"=B{row}*C{row}/B2_Scale*EXP(-D{row}/B2_Scale)", "dec", "output"
        )
        sb.cell(8, row, f"=INDEX(D_RegionGrowth,{i + 1})", "pct")
    bottom = top + 3
    sb.define("B2_Alloc", 4, top, 4, bottom)
    total = bottom + 1
    sb.cell(1, total, "TỔNG", bold=True)
    sb.cell(4, total, f"=SUM(D{top}:D{bottom})", "usd", "output")
    sb.cell(5, total, f"=SUM(E{top}:E{bottom})", "usd", "output")
    sb.cell(6, total, f"=E{total}/D{total}", "dec", "output")
    sb.define("B2_AllocSum", 4, total)
    sb.define("B2_Uplift", 5, total)
    sb.define("B2_ROAS", 6, total)
    sb.row = total + 1
    sb.cell(
        1,
        sb.row,
        "Hàm mục tiêu: MAX DT tăng thêm (ô E tổng). Ràng buộc: Σ ngân sách ≤ tổng; tối thiểu ≤ mỗi vùng ≤ tối đa.",
    ).font = Font(italic=True)
    sb.skip(2)

    sb.section(
        "PHƯƠNG ÁN SO SÁNH: CHIA ĐỀU NGÂN SÁCH (dùng cho Goal Seek và Data Table)"
    )
    sb.param(
        "Tổng ngân sách (chia đều 4 vùng)",
        m["budget"][0],
        "usd",
        "B2_EqBudget",
        unit="USD",
    )
    sb.param(
        "DT tăng thêm khi chia đều",
        f"=(1-EXP(-(B2_EqBudget/4)/B2_Scale))*SUMPRODUCT(D_RegionRevenue,C{top}:C{bottom})",
        "usd",
        "B2_EqUplift",
        "output",
        "USD",
    )
    sb.param(
        "ROAS khi chia đều",
        "=B2_EqUplift/B2_EqBudget",
        "dec",
        "B2_EqROAS",
        "output",
        "lần",
    )
    sb.param(
        "Lợi ích của phân bổ tối ưu so với chia đều",
        "=B2_Uplift-B2_EqUplift",
        "usd",
        "B2_Gain",
        "output",
        "USD",
        "Chỉ có ý nghĩa khi tổng ngân sách hai phương án bằng nhau",
    )
    sb.skip()

    tools["goal_seek"] += goal_seek_block(
        sb,
        [
            (
                "Chia đều thì cần tổng ngân sách bao nhiêu để DT tăng thêm đạt mục tiêu?",
                "B2_EqUplift",
                m["revenue_goal"][0],
                "B2_EqBudget",
                "USD",
                "usd",
                "B2_GS_EqBudget",
            )
        ],
    )
    data_table_block(
        sb,
        "DATA TABLE 1 CHIỀU – HIỆU QUẢ THEO TỔNG NGÂN SÁCH (chia đều)",
        ["=B2_EqUplift", "=B2_EqROAS"],
        list(range(40_000, 160_001, 20_000)),
        "usd",
        ["usd", "dec"],
        ["Tổng ngân sách", "DT tăng thêm", "ROAS"],
        "B2_DT",
    )
    tools["data_tables"].append(
        {"range": "B2_DT", "row_input": None, "col_input": "B2_EqBudget"}
    )

    tools["solver"].append(
        {
            "sheet": "B2_Marketing",
            "report_suffix": "B2",
            "objective": "B2_Uplift",
            "sense": 1,
            "vars": "B2_Alloc",
            "engine": 1,
            "constraints": [
                ["B2_AllocSum", 1, "B2_Budget"],
                ["B2_Alloc", 3, "B2_Min"],
                ["B2_Alloc", 1, "B2_Max"],
            ],
            "scenario": {
                "summary_name": "SC_B2",
                "param": "B2_Budget",
                "results": ["B2_AllocSum", "B2_Uplift", "B2_ROAS"],
                "items": [{"name": k, "value": v} for k, v in P.MKT_SCENARIOS.items()],
                "base": "Ngân sách cơ sở",
                "comment": "Nghiệm tối ưu của Solver",
            },
        }
    )


def sheet_b3(wb, base, tools):
    sub = base["subcat"]
    n = len(sub)
    sb = SheetBuilder(
        wb, "B3_Nhap_hang", [16, 15, 13, 10, 10, 13, 13, 13, 13, 13, 13, 11, 13, 13]
    )
    sb.title(
        "BÀI 3 – KẾ HOẠCH NHẬP HÀNG THEO SUB-CATEGORY NĂM 2019",
        "Phân bổ ngân sách nhập hàng và sức chứa kho có hạn cho 17 Sub-Category để tối đa lợi nhuận gộp.",
    )
    inv = P.INV
    sb.section("THAM SỐ")
    sb.header(["Tham số", "Giá trị", "Đơn vị", "Ghi chú"])
    sb.param(
        "Tăng trưởng nhu cầu 2019",
        inv["demand_growth"][0],
        "pct",
        "B3_Growth",
        unit="%",
        note=inv["demand_growth"][2],
    )
    sb.param(
        "Ngân sách nhập hàng (giá vốn)",
        inv["budget"][0],
        "usd",
        "B3_Budget",
        unit="USD",
        note=inv["budget"][2],
    )
    sb.param(
        "Sức chứa kho",
        inv["capacity"][0],
        "dec",
        "B3_Capacity",
        unit="m³",
        note=inv["capacity"][2],
    )
    sb.param(
        "Mức đáp ứng tối thiểu",
        inv["min_service"][0],
        "pct",
        "B3_MinService",
        unit="% nhu cầu",
        note=inv["min_service"][2],
    )
    sb.param(
        "Tỷ lệ nhu cầu thực tế so với dự báo",
        1.0,
        "pct",
        "B3_Realization",
        unit="%",
        note="Dùng cho Data Table: nhu cầu thực tế có thể lệch dự báo",
    )
    sb.skip()

    sb.section(
        "MÔ HÌNH (đơn vị tiền: USD; biến quyết định = giá trị nhập theo giá vốn)", 14
    )
    sb.header(
        [
            "Sub-Category", "Category", "DT 2018", "Biên LN", "m³/1.000$", "Nhu cầu 2019 (DT)",
            "Nhập tối đa (giá vốn)", "Nhập tối thiểu", "GIÁ TRỊ NHẬP", "DT kế hoạch",
            "LN gộp kế hoạch", "Thể tích (m³)", "DT thực hiện", "LN thực hiện",
        ]
    )  # fmt: skip
    top = sb.row
    for i, (s, r_) in enumerate(sub.iterrows()):
        row = top + i
        k = i + 1
        sb.cell(1, row, s)
        sb.cell(2, row, r_["category"])
        sb.cell(3, row, f"=INDEX(D_SubRevenue,{k})", "usd")
        sb.cell(4, row, f"=INDEX(G_SubMargin,{k})", "pct")
        sb.cell(5, row, f"=INDEX(G_SubVolume,{k})", "dec")
        sb.cell(6, row, f"=C{row}*(1+B3_Growth)", "usd")
        sb.cell(7, row, f"=F{row}*(1-D{row})", "usd")
        sb.cell(8, row, f"=G{row}*B3_MinService", "usd")
        sb.cell(9, row, round(inv["budget"][0] / n, 0), "usd", "decision")
        sb.define(f"B3_Nhap_{s}", 9, row)
        sb.cell(10, row, f"=I{row}/(1-D{row})", "usd", "output")
        sb.cell(11, row, f"=J{row}-I{row}", "usd", "output")
        sb.cell(12, row, f"=I{row}*E{row}/1000", "dec", "output")
        sb.cell(13, row, f"=MIN(J{row},F{row}*B3_Realization)", "usd", "output")
        sb.cell(14, row, f"=M{row}-I{row}", "usd", "output")
    bottom = top + n - 1
    total = bottom + 1
    sb.define("B3_X", 9, top, 9, bottom)
    sb.define("B3_Max", 7, top, 7, bottom)
    sb.define("B3_Min", 8, top, 8, bottom)
    sb.cell(1, total, "TỔNG", bold=True)
    for col, fmt in (
        (6, "usd"),
        (7, "usd"),
        (8, "usd"),
        (9, "usd"),
        (10, "usd"),
        (11, "usd"),
        (12, "dec"),
        (13, "usd"),
        (14, "usd"),
    ):
        L = get_column_letter(col)
        sb.cell(col, total, f"=SUM({L}{top}:{L}{bottom})", fmt, "output")
    sb.define("B3_XSum", 9, total)
    sb.define("B3_Profit", 11, total)
    sb.define("B3_Space", 12, total)
    sb.define("B3_RealProfit", 14, total)
    sb.row = total + 1
    sb.cell(
        1,
        sb.row,
        "Hàm mục tiêu: MAX LN gộp kế hoạch (K tổng). Ràng buộc: Σ nhập ≤ ngân sách; Σ thể tích ≤ sức chứa; nhập tối thiểu ≤ nhập ≤ nhập tối đa.",
    ).font = Font(italic=True)
    sb.skip(2)

    sb.section("CHỈ SỐ TỔNG HỢP")
    sb.param(
        "Tỷ lệ sử dụng ngân sách",
        "=B3_XSum/B3_Budget",
        "pct",
        "B3_BudgetUse",
        "output",
        "%",
    )
    sb.param(
        "Tỷ lệ sử dụng kho",
        "=B3_Space/B3_Capacity",
        "pct",
        "B3_SpaceUse",
        "output",
        "%",
    )
    sb.param(
        "Tỷ lệ đáp ứng nhu cầu (theo DT)",
        f"=J{total}/F{total}",
        "pct",
        "B3_FillRate",
        "output",
        "%",
    )
    sb.param(
        "Biên LN gộp của kế hoạch",
        f"=K{total}/J{total}",
        "pct",
        "B3_PlanMargin",
        "output",
        "%",
    )
    sb.param(
        "Hàng tồn không bán được (giá vốn)",
        f"=SUMPRODUCT(J{top}:J{bottom}-M{top}:M{bottom},1-D{top}:D{bottom})",
        "usd",
        "B3_Leftover",
        "output",
        "USD",
        "Khi nhu cầu thực tế thấp hơn kế hoạch",
    )
    sb.skip()

    data_table_block(
        sb,
        "DATA TABLE 1 CHIỀU – RỦI RO KHI NHU CẦU THỰC TẾ LỆCH DỰ BÁO (với kế hoạch nhập hiện tại)",
        ["=B3_RealProfit", "=B3_Leftover"],
        [0.8, 0.85, 0.9, 0.95, 1.0, 1.05, 1.1, 1.2],
        "pct",
        ["usd", "usd"],
        ["Nhu cầu thực tế", "LN thực hiện", "Tồn kho dư"],
        "B3_DT",
    )
    tools["data_tables"].append(
        {"range": "B3_DT", "row_input": None, "col_input": "B3_Realization"}
    )

    tools["solver"].append(
        {
            "sheet": "B3_Nhap_hang",
            "report_suffix": "B3",
            "objective": "B3_Profit",
            "sense": 1,
            "vars": "B3_X",
            "engine": 2,
            "constraints": [
                ["B3_XSum", 1, "B3_Budget"],
                ["B3_Space", 1, "B3_Capacity"],
                ["B3_X", 3, "B3_Min"],
                ["B3_X", 1, "B3_Max"],
            ],
            "scenario": {
                "summary_name": "SC_B3",
                "params": ["B3_Budget", "B3_Capacity"],
                "results": [
                    "B3_XSum",
                    "B3_Space",
                    "B3_Profit",
                    "B3_FillRate",
                    "B3_BudgetUse",
                    "B3_SpaceUse",
                ],
                "items": [
                    {"name": k, "values": list(v)} for k, v in P.INV_SCENARIOS.items()
                ],
                "base": "Cơ sở",
                "comment": "Nghiệm tối ưu của Solver",
            },
        }
    )


def sheet_b4(wb, base, tools):
    sb = SheetBuilder(wb, "B4_Giao_hang", [40, 14, 14, 14, 14, 14, 16, 14, 14, 14, 14])
    sb.title(
        "BÀI 4 – CƠ CẤU PHƯƠNG THỨC GIAO HÀNG NĂM 2019",
        "Phân bổ số đơn cho 4 Ship Mode để tối thiểu chi phí giao hàng mà vẫn đạt mục tiêu tỷ lệ giao nhanh.",
    )
    s = P.SHIP
    sb.section("THAM SỐ")
    sb.header(["Tham số", "Giá trị", "Đơn vị", "Ghi chú"])
    sb.param(
        "Số đơn dự báo 2019",
        "=ROUND(D_Orders*(1+G_OrderGrowth),0)",
        "num",
        "B4_N",
        "output",
        "đơn",
        "Như Bài 1",
    )
    sb.param(
        "Mục tiêu tỷ lệ giao nhanh (≤ 2 ngày)",
        s["fast_target"][0],
        "pct",
        "B4_FastTarget",
        unit="%",
        note=s["fast_target"][2],
    )
    sb.param(
        "Năng lực Same Day",
        s["sameday_cap"][0],
        "num",
        "B4_SameDayCap",
        unit="đơn",
        note=s["sameday_cap"][2],
    )
    sb.param(
        "Ngày giao TB tối đa",
        s["max_avg_days"][0],
        "dec",
        "B4_MaxDays",
        unit="ngày",
        note=s["max_avg_days"][2],
    )
    sb.param(
        "Tỷ trọng tối thiểu mỗi phương thức",
        s["min_share"][0],
        "pct",
        "B4_MinShare",
        unit="%",
        note=s["min_share"][2],
    )
    sb.skip()

    sb.section("MÔ HÌNH", 11)
    sb.header(
        ["Ship Mode", "Chi phí/đơn", "Tỷ lệ giao nhanh", "Ngày giao TB", "Số đơn 2018", "Tỷ trọng 2018",
         "SỐ ĐƠN 2019", "Chi phí", "Đơn giao nhanh", "Ngày × đơn", "Tỷ trọng mới"]
    )  # fmt: skip
    top = sb.row
    bottom = top + 3
    total = bottom + 1
    for i, mode in enumerate(P.SHIP_MODES):
        row = top + i
        k = i + 1
        share = base["ship"].loc[mode, "orders"] / base["ship"]["orders"].sum()
        sb.cell(1, row, mode)
        sb.cell(2, row, f"=INDEX(G_ShipCost,{k})", "usd2")
        sb.cell(3, row, f"=INDEX(D_ShipFast,{k})", "pct")
        sb.cell(4, row, f"=INDEX(D_ShipDays,{k})", "dec")
        sb.cell(5, row, f"=INDEX(D_ShipOrders,{k})", "num")
        sb.cell(6, row, f"=E{row}/SUM($E${top}:$E${bottom})", "pct")
        sb.cell(7, row, round(share * P.orders_2019(base), 1), "num", "decision")
        sb.define("B4_Don_" + mode.replace(" ", ""), 7, row)
        sb.cell(8, row, f"=G{row}*B{row}", "usd", "output")
        sb.cell(9, row, f"=G{row}*C{row}", "num", "output")
        sb.cell(10, row, f"=G{row}*D{row}", "num", "output")
        sb.cell(11, row, f"=G{row}/$G${total}", "pct", "output")
    sb.define("B4_X", 7, top, 7, bottom)
    sb.cell(1, total, "TỔNG", bold=True)
    for col, fmt in ((7, "num"), (8, "usd"), (9, "num"), (10, "num")):
        L = get_column_letter(col)
        sb.cell(col, total, f"=SUM({L}{top}:{L}{bottom})", fmt, "output")
    sb.define("B4_XSum", 7, total)
    sb.define("B4_Cost", 8, total)
    sb.define("B4_FastOrders", 9, total)
    sb.define("B4_DaySum", 10, total)
    sb.row = total + 1
    sb.cell(
        1,
        sb.row,
        "Hàm mục tiêu: MIN tổng chi phí (H tổng). Ràng buộc: Σ đơn = dự báo; đơn giao nhanh ≥ mục tiêu; ngày×đơn ≤ trần; Same Day ≤ năng lực; mỗi mode ≥ tối thiểu.",
    ).font = Font(italic=True)
    sb.skip(2)

    sb.section("VẾ PHẢI RÀNG BUỘC VÀ CHỈ SỐ")
    sb.param(
        "Số đơn giao nhanh tối thiểu",
        "=B4_FastTarget*B4_N",
        "num",
        "B4_FastReq",
        "output",
        "đơn",
    )
    sb.param(
        "Trần tổng ngày × đơn",
        "=B4_MaxDays*B4_N",
        "num",
        "B4_DayCap",
        "output",
        "ngày·đơn",
    )
    sb.param(
        "Số đơn tối thiểu mỗi phương thức",
        "=B4_MinShare*B4_N",
        "num",
        "B4_MinOrders",
        "output",
        "đơn",
    )
    sb.param(
        "Tỷ lệ giao nhanh đạt được",
        "=B4_FastOrders/B4_XSum",
        "pct",
        "B4_FastRate",
        "output",
        "%",
    )
    sb.param(
        "Ngày giao TB đạt được",
        "=B4_DaySum/B4_XSum",
        "dec",
        "B4_AvgDays",
        "output",
        "ngày",
    )
    sb.param(
        "Chi phí giao TB / đơn",
        "=B4_Cost/B4_XSum",
        "usd2",
        "B4_CostPerOrder",
        "output",
        "USD/đơn",
    )
    sb.param(
        "Chi phí nếu giữ cơ cấu 2018",
        f"=B4_N*SUMPRODUCT(F{top}:F{bottom},B{top}:B{bottom})",
        "usd",
        "B4_BaseCost",
        "output",
        "USD",
        "Tỷ lệ giao nhanh khi đó ≈ 22%",
    )
    sb.param(
        "Chi phí tăng thêm so với cơ cấu 2018",
        "=B4_Cost-B4_BaseCost",
        "usd",
        "B4_ExtraCost",
        "output",
        "USD",
    )
    sb.skip()

    sb.section(
        "PHƯƠNG ÁN ĐƠN GIẢN: CHỈ TĂNG FIRST CLASS (giữ Same Day, Second Class như 2018; Standard nhận phần còn lại)"
    )
    fc = base["ship"].loc["First Class", "orders"] / base["ship"]["orders"].sum()
    sb.param("Tỷ trọng First Class", round(fc, 4), "pct", "B4_FCShare", unit="%")
    sb.param(
        "Tỷ trọng Standard Class",
        f"=1-F{top}-F{top + 2}-B4_FCShare",
        "pct",
        "B4_STDShare",
        "output",
        "%",
    )
    sb.param(
        "Tỷ lệ giao nhanh",
        f"=F{top}*C{top}+B4_FCShare*C{top + 1}+F{top + 2}*C{top + 2}+B4_STDShare*C{top + 3}",
        "pct",
        "B4_SimpleFast",
        "output",
        "%",
    )
    sb.param(
        "Tổng chi phí giao",
        f"=B4_N*(F{top}*B{top}+B4_FCShare*B{top + 1}+F{top + 2}*B{top + 2}+B4_STDShare*B{top + 3})",
        "usd",
        "B4_SimpleCost",
        "output",
        "USD",
    )
    sb.skip()

    tools["goal_seek"] += goal_seek_block(
        sb,
        [
            (
                "Chỉ tăng First Class: cần tỷ trọng bao nhiêu để đạt mục tiêu giao nhanh?",
                "B4_SimpleFast",
                s["fast_target"][0],
                "B4_FCShare",
                "%",
                "pct",
                "B4_GS_FCShare",
            )
        ],
    )
    data_table_block(
        sb,
        "DATA TABLE 1 CHIỀU – CHI PHÍ VÀ TỶ LỆ GIAO NHANH THEO TỶ TRỌNG FIRST CLASS",
        ["=B4_SimpleFast", "=B4_SimpleCost"],
        [0.10, 0.15, 0.20, 0.25, 0.30, 0.35, 0.40, 0.45, 0.50],
        "pct",
        ["pct", "usd"],
        ["Tỷ trọng First Class", "Tỷ lệ giao nhanh", "Tổng chi phí"],
        "B4_DT",
    )
    tools["data_tables"].append(
        {"range": "B4_DT", "row_input": None, "col_input": "B4_FCShare"}
    )

    tools["solver"].append(
        {
            "sheet": "B4_Giao_hang",
            "report_suffix": "B4",
            "objective": "B4_Cost",
            "sense": 2,
            "vars": "B4_X",
            "engine": 2,
            "constraints": [
                ["B4_XSum", 2, "B4_N"],
                ["B4_FastOrders", 3, "B4_FastReq"],
                ["B4_DaySum", 1, "B4_DayCap"],
                ["B4_Don_SameDay", 1, "B4_SameDayCap"],
                ["B4_X", 3, "B4_MinOrders"],
            ],
            "scenario": {
                "summary_name": "SC_B4",
                "param": "B4_FastTarget",
                "results": [
                    "B4_Cost",
                    "B4_ExtraCost",
                    "B4_FastRate",
                    "B4_AvgDays",
                    "B4_CostPerOrder",
                ],
                "items": [{"name": k, "value": v} for k, v in P.SHIP_SCENARIOS.items()],
                "base": "Mục tiêu 30%",
                "comment": "Nghiệm tối ưu của Solver",
            },
        }
    )


def main():
    sys.stdout.reconfigure(
        encoding="utf-8"
    )  # console Windows mặc định không phải UTF-8
    base = P.load_base_year()
    wb = Workbook()
    tools = {"goal_seek": [], "data_tables": [], "scenarios": [], "solver": []}

    sheet_guide(wb)
    sheet_data(wb, base)
    sheet_assumptions(wb, base)
    sheet_b1(wb, base, tools)
    sheet_b2(wb, base, tools)
    sheet_b3(wb, base, tools)
    sheet_b4(wb, base, tools)

    OUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT_XLSX)
    OUT_TOOLS.write_text(
        json.dumps(tools, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"Đã ghi {OUT_XLSX.relative_to(P.ROOT)} và {OUT_TOOLS.relative_to(P.ROOT)}")


if __name__ == "__main__":
    main()
