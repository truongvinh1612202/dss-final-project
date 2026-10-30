# Đồ án Hệ hỗ trợ ra quyết định (DSS): Superstore Sales 2015–2018

Đồ án môn Hệ hỗ trợ ra quyết định, Cao học K34 HCMUS. Đặc tả đề bài: [`dac_ta_do_an_DSS.md`](dac_ta_do_an_DSS.md).

- **Phần A – Data-Driven DSS (Power BI)**: làm sạch dữ liệu, star schema, KPI và dashboard.
- **Phần B – Model-Driven DSS (Excel)**: 4 bài toán ra quyết định, dùng Goal Seek, Data Table, Scenario Manager, Solver.

## Bắt đầu từ đâu

**Hướng dẫn cho thành viên nhóm (cả Phần A và Phần B):** [`docs/Huong_dan_PowerBI.md`](docs/Huong_dan_PowerBI.md)

| Cần làm | Dùng file |
|---|---|
| Dựng dashboard Power BI | `data/star/*.csv` (6 bảng) + Phần A của hướng dẫn |
| Dựng mô hình Excel | `output/PhanB_Mau_nhap_lieu.xlsx` (file mẫu có sẵn dữ liệu) + Phần B của hướng dẫn |
| So số liệu | `docs/KPI_doi_chieu.md` (Phần A), `docs/PhanB_Ket_qua.md` (Phần B) |
| Viết báo cáo | `docs/Nhat_ky_tien_xu_ly.md`, `docs/Star_schema.md`, `docs/KPI_va_DAX.md`, `docs/PhanB_Mo_hinh.md` |

> ⚠️ Không mở các file CSV bằng Excel rồi lưu lại, vì Excel sẽ xóa số 0 đầu của Postal Code.

## Cấu trúc thư mục

```
data/raw/     dữ liệu gốc tải từ Kaggle (train.csv), không sửa
data/clean/   dữ liệu sau tiền xử lý (Superstore_clean.csv)
data/star/    star schema: Fact_Sales + 5 bảng Dim
etl/          script tiền xử lý, tách star schema, tính KPI tham chiếu
model/        script dựng mô hình Excel Phần B, lời giải đối chiếu (scipy)
output/       file Excel mẫu (PhanB_Mau_nhap_lieu.xlsx), bản đáp án tham khảo (PhanB_Dap_an_tham_khao.xlsx, không nộp)
docs/         hướng dẫn, nhật ký xử lý, KPI/DAX, mô hình và kết quả Phần B
```

## Chạy lại pipeline (không bắt buộc)

Cần Python 3.10+ với `pandas`, `scipy`, `openpyxl`. Bước chạy công cụ Excel cần Microsoft Excel kèm Solver add-in.

```
python -I etl/preprocess.py            # làm sạch → data/clean + docs/Nhat_ky_tien_xu_ly.md
python -I etl/build_star_schema.py     # star schema → data/star + docs/Kiem_tra_star_schema.md
python -I etl/kpi_reference.py         # KPI tham chiếu → docs/KPI_doi_chieu.md

python -I model/solve_reference.py     # lời giải đối chiếu → output/reference.json
python -I model/build_model.py         # bản đáp án → output/PhanB_Dap_an_tham_khao.xlsx
powershell -NoProfile -ExecutionPolicy Bypass -File model/apply_excel_tools.ps1
python -I model/verify_excel.py        # đối chiếu Excel ↔ scipy → docs/PhanB_Ket_qua.md
python -I model/build_template.py      # file mẫu → output/PhanB_Mau_nhap_lieu.xlsx
```
