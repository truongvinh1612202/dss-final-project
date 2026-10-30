"""Tạo file mẫu Phần B: chỉ có dữ liệu đầu vào, nhóm tự gõ công thức và chạy công cụ Excel.

Chạy: python -I model/build_template.py
Đầu ra: output/PhanB_Mau_nhap_lieu.xlsx

Giữ nguyên bố cục, nhãn, màu ô, định dạng số, số liệu 2018, giả định, tham số và giá trị
khởi đầu của biến quyết định như output/PhanB_Dap_an_tham_khao.xlsx. Bỏ toàn bộ:
công thức, Data Table, named range, kết quả Goal Seek/Solver/Scenario.
Hướng dẫn tự làm: docs/Huong_dan_PowerBI.md, PHẦN B.
"""

import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_model as B
import params as P

OUT = P.ROOT / "output" / "PhanB_Mau_nhap_lieu.xlsx"

GUIDE = [
    ("FILE MẪU PHẦN B – DỮ LIỆU ĐẦU VÀO ĐÃ ĐIỀN SẴN", None),
    ("", None),
    (
        "Cách dùng",
        "Làm theo docs/Huong_dan_PowerBI.md – PHẦN B. Lưu thành MaNhom_Model.xlsx khi xong.",
    ),
    (
        "Đã có sẵn",
        "Số liệu thật 2018, giả định, tham số, giá trị khởi đầu của biến quyết định, nhãn và bố cục.",
    ),
    (
        "Nhóm tự làm",
        "Gõ công thức vào các ô xanh dương còn trống; chạy Goal Seek, Data Table, Scenario Manager, Solver.",
    ),
    ("", None),
    ("Màu ô", "Ý nghĩa"),
    ("Vàng", "Tham số đầu vào – được phép thay đổi để phân tích What-if."),
    (
        "Xanh lá",
        "Biến quyết định – Solver thay đổi các ô này (đã có giá trị khởi đầu).",
    ),
    ("Xanh dương", "Ô cần gõ CÔNG THỨC (đang trống)."),
    ("", None),
    ("Lưu ý", "Mọi chi phí và biên lợi nhuận là GIẢ ĐỊNH (dataset chỉ có doanh thu)."),
    (
        "",
        "Bật Solver: File → Options → Add-ins → Excel Add-ins → Go → chọn Solver Add-in.",
    ),
]


def name_addresses(wb):
    """Tên ô → địa chỉ không có $ (để thay nhãn 'B1_Profit' trong bảng Goal Seek thành 'B19')."""
    out = {}
    for name, dn in wb.defined_names.items():
        dests = list(dn.destinations)
        if len(dests) == 1:
            out[name] = dests[0][1].replace("$", "")
    return out


def strip_calculations(wb):
    names = name_addresses(wb)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("=") and c.data_type == "f":
                    c.value = None
                elif v is not None and not isinstance(v, (str, int, float)):
                    c.value = None  # Data Table
                elif isinstance(v, str) and v in names:
                    c.value = names[v]
    for name in list(wb.defined_names):
        del wb.defined_names[name]


def rewrite_guide(ws):
    ws.delete_rows(1, ws.max_row)
    for i, (a, b) in enumerate(GUIDE, 1):
        ws.cell(row=i, column=1, value=a)
        ws.cell(row=i, column=2, value=b)
    ws["A1"].font = Font(bold=True, size=14, color="1F4E78")
    for c in (1, 2):
        ws.cell(row=7, column=c).font = Font(bold=True, color="FFFFFF")
        ws.cell(row=7, column=c).fill = B.FILL["header"]
    for r, kind in ((8, "input"), (9, "decision"), (10, "output")):
        ws.cell(row=r, column=1).fill = B.FILL[kind]
        ws.cell(row=r, column=1).font = Font()


def main():
    sys.stdout.reconfigure(
        encoding="utf-8"
    )  # console Windows mặc định không phải UTF-8
    base = P.load_base_year()
    wb = Workbook()
    tools = {"goal_seek": [], "data_tables": [], "scenarios": [], "solver": []}
    B.sheet_guide(wb)
    B.sheet_data(wb, base)
    B.sheet_assumptions(wb, base)
    B.sheet_b1(wb, base, tools)
    B.sheet_b2(wb, base, tools)
    B.sheet_b3(wb, base, tools)
    B.sheet_b4(wb, base, tools)
    strip_calculations(wb)
    rewrite_guide(wb["Huong_dan"])
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"Đã ghi {OUT.relative_to(P.ROOT)}")


if __name__ == "__main__":
    main()
