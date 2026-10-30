# Kiểm tra star schema

> File này do `etl/build_star_schema.py` tự sinh. Không sửa tay, hãy chạy lại script.

Nguồn: `data/clean/Superstore_clean.csv` (dữ liệu 2015–2018). Kết quả: `data/star/`. Sơ đồ: `docs/Star_schema.md`.

## Số dòng từng bảng

| Bảng | Số dòng | Số cột |
|---|---|---|
| Fact_Sales | 9,799 | 11 |
| Dim_Customer | 793 | 3 |
| Dim_Product | 1,893 | 5 |
| Dim_Geography | 628 | 5 |
| Dim_Date | 1,826 | 12 |
| Dim_ShipMode | 4 | 2 |

## Kiểm tra toàn vẹn (16/16 đạt)

| # | Kiểm tra | Kết quả | Chi tiết |
|---|---|---|---|
| 1 | Khóa chính Fact_Sales.Row ID duy nhất, không null | Đạt | 9,799 dòng, 9,799 giá trị phân biệt |
| 2 | Khóa chính Dim_Customer.Customer ID duy nhất, không null | Đạt | 793 dòng, 793 giá trị phân biệt |
| 3 | Khóa chính Dim_Product.ProductKey duy nhất, không null | Đạt | 1,893 dòng, 1,893 giá trị phân biệt |
| 4 | Khóa chính Dim_Geography.GeoKey duy nhất, không null | Đạt | 628 dòng, 628 giá trị phân biệt |
| 5 | Khóa chính Dim_Date.DateKey duy nhất, không null | Đạt | 1,826 dòng, 1,826 giá trị phân biệt |
| 6 | Khóa chính Dim_ShipMode.ShipModeKey duy nhất, không null | Đạt | 4 dòng, 4 giá trị phân biệt |
| 7 | Fact_Sales.OrderDateKey → Dim_Date.DateKey không có dòng mồ côi | Đạt | 0 dòng mồ côi |
| 8 | Fact_Sales.ShipDateKey → Dim_Date.DateKey không có dòng mồ côi | Đạt | 0 dòng mồ côi |
| 9 | Fact_Sales.Customer ID → Dim_Customer.Customer ID không có dòng mồ côi | Đạt | 0 dòng mồ côi |
| 10 | Fact_Sales.ProductKey → Dim_Product.ProductKey không có dòng mồ côi | Đạt | 0 dòng mồ côi |
| 11 | Fact_Sales.GeoKey → Dim_Geography.GeoKey không có dòng mồ côi | Đạt | 0 dòng mồ côi |
| 12 | Fact_Sales.ShipModeKey → Dim_ShipMode.ShipModeKey không có dòng mồ côi | Đạt | 0 dòng mồ côi |
| 13 | Số dòng Fact = số dòng bản sạch | Đạt | 9,799 / 9,799 |
| 14 | Tổng Sales Fact = tổng Sales bản sạch | Đạt | 2,261,255.41 / 2,261,255.41 |
| 15 | Tổng Sales theo Category, Region, Segment sau JOIN = bảng phẳng | Đạt | 9,799 dòng sau JOIN; lệch: không |
| 16 | Mỗi State chỉ thuộc một Region | Đạt | 0 State thuộc nhiều Region |

## Lưu ý khi nạp vào Power BI

- **Postal Code** phải để kiểu **Text**. Power Query có thể tự đổi sang số và làm mất số 0 đầu (`06824` → `6824`). Xóa bước "Changed Type" tự động hoặc đặt lại kiểu Text.
- Không mở các file CSV bằng Excel rồi lưu lại, vì Excel cũng làm mất số 0 đầu.
- `Dim_Date` có hai quan hệ với `Fact_Sales`: `OrderDateKey` là quan hệ **active**, `ShipDateKey` là **inactive**. Khi cần tính theo ngày giao, dùng `CALCULATE(..., USERELATIONSHIP(Fact_Sales[ShipDateKey], Dim_Date[DateKey]))`.
- `Dim_Date` chạy từ 2015-01-01 đến 2019-12-31, vì Ship Date muộn nhất là 2019-01-05. Đánh dấu bảng này là Date table (Mark as date table) theo cột `Date`.
- Sắp xếp `Month Name` theo `Month` và `Weekday Name` theo `DayOfWeek` (Sort by column) để tháng và thứ hiển thị đúng thứ tự.
- `ShipModeKey` xếp theo tốc độ giao hàng: Same Day = 1, First Class = 2, Second Class = 3, Standard Class = 4.
- Mọi quan hệ là 1–nhiều, lọc một chiều từ Dim sang Fact.
