# Nhật ký tiền xử lý dữ liệu Superstore

> File này do `etl/preprocess.py` tự sinh. Không sửa tay, hãy chạy lại script.

- Nguồn: `data/raw/train.csv` (Superstore Sales, Kaggle), giữ nguyên bản gốc ở chế độ chỉ đọc.
- Kết quả: `data/clean/Superstore_clean.csv`.
- Nguyên tắc: chỉ xử lý lỗi có thật trong dữ liệu, không chèn lỗi giả lập.

## 1. Khảo sát dữ liệu gốc

| Cột | Kiểu | Số ô trống | Số giá trị phân biệt |
|---|---|---|---|
| Row ID | int64 | 0 | 9800 |
| Order ID | object | 0 | 4922 |
| Order Date | object | 0 | 1230 |
| Ship Date | object | 0 | 1326 |
| Ship Mode | object | 0 | 4 |
| Customer ID | object | 0 | 793 |
| Customer Name | object | 0 | 793 |
| Segment | object | 0 | 3 |
| Country | object | 0 | 1 |
| City | object | 0 | 529 |
| State | object | 0 | 49 |
| Postal Code | object | 11 | 626 |
| Region | object | 0 | 4 |
| Product ID | object | 0 | 1861 |
| Category | object | 0 | 3 |
| Sub-Category | object | 0 | 17 |
| Product Name | object | 0 | 1849 |
| Sales | float64 | 0 | 5757 |

Dataset chỉ có một cột định lượng là **Sales**. Không có Quantity, Discount, Profit.

## 2. Các bước làm sạch và làm giàu dữ liệu

| # | Bước | Phát hiện | Cách xử lý | Số dòng trước → sau |
|---|---|---|---|---|
| 1 | Nạp dữ liệu | 9,800 dòng, 18 cột; mã hóa UTF-8; ngày dạng d/m/yyyy | Đọc Postal Code dạng text để giữ số 0 đầu | 9,800 → 9,800 |
| 2 | Kiểm tra trùng toàn bộ dòng | 1 dòng trùng hoàn toàn với dòng khác (Row ID [3407]) | Xóa dòng trùng, giữ lại lần xuất hiện đầu tiên | 9,800 → 9,799 |
| 3 | Kiểm tra trùng Row ID | 0 Row ID bị trùng | Không cần xử lý | 9,799 → 9,799 |
| 4 | Kiểm tra trùng Order ID + Product ID | 7 cặp (Order ID, Product ID) xuất hiện nhiều lần nhưng khác Sales | Giữ nguyên: là các dòng hàng hợp lệ (cùng sản phẩm, mua thêm trong cùng đơn) | 9,799 → 9,799 |
| 5 | Chuẩn hóa kiểu ngày | Order Date, Ship Date ở dạng text d/m/yyyy | Ép kiểu Date theo định dạng d/m/yyyy (Power Query: locale en-GB); xuất ra dạng yyyy-mm-dd | 9,799 → 9,799 |
| 6 | Chuẩn hóa Postal Code | 429 dòng có mã bưu chính 4 chữ số do mất số 0 đầu (vd 6824) | Đệm số 0 bên trái cho đủ 5 ký tự (6824 → 06824) | 9,799 → 9,799 |
| 7 | Xử lý ô trống | 11 dòng trống Postal Code, tất cả thuộc Burlington, Vermont; các cột khác không có ô trống | Điền 05401 (mã bưu chính USPS của Burlington, VT) | 9,799 → 9,799 |
| 8 | Làm sạch text | 220 dòng Product Name chứa ký tự NBSP (U+00A0); 31 ô có khoảng trắng thừa | Thay NBSP bằng dấu cách; Trim và gộp khoảng trắng liên tiếp ở mọi cột text | 9,799 → 9,799 |
| 9 | Kiểm tra tính thống nhất | Giá trị lạ ở cột phân loại: không có; Ship Date < Order Date: 0; Sales ≤ 0: 0 | Không phát hiện bản ghi vô lý, không cần loại bỏ | 9,799 → 9,799 |
| 10 | Tạo cột bổ sung | Cần thuộc tính thời gian, thời gian giao hàng, cờ ngoại lai và khóa thay thế cho star schema | Thêm Order Year/Quarter/Month/YearMonth, Ship Days, Is Sales Outlier (Sales > Q3 + 1.5·IQR = 500.56; 1,145 dòng, giữ lại không xóa), ProductKey (1,893 sản phẩm), GeoKey (628 địa điểm) | 9,799 → 9,799 |

## 3. Kiểm tra tính nhất quán cho star schema

| Kiểm tra | Số vi phạm | Kết luận |
|---|---|---|
| Mỗi Customer ID → một Customer Name | 0 | Nhất quán |
| Mỗi Customer ID → một Segment | 0 | Nhất quán |
| Mỗi Product ID → một Product Name | 32 | Không nhất quán → cần khóa thay thế |
| Mỗi Product Name → một Product ID | 16 | Không nhất quán → cần khóa thay thế |
| Mỗi Product ID → một Sub-Category | 0 | Nhất quán |
| Mỗi Postal Code → một City | 1 | Không nhất quán → cần khóa thay thế |
| Mỗi Postal Code → một State | 0 | Nhất quán |
| Mỗi City → một State | 55 | Không nhất quán → cần khóa thay thế |
| Mỗi Order ID → một Order Date | 0 | Nhất quán |
| Mỗi Order ID → một Ship Date | 0 | Nhất quán |
| Mỗi Order ID → một Ship Mode | 0 | Nhất quán |
| Mỗi Order ID → một Customer ID | 0 | Nhất quán |

Quyết định thiết kế khóa:
- **DimCustomer**: khóa `Customer ID`.
- **DimProduct**: khóa thay thế `ProductKey` trên cặp (Product ID, Product Name), vì Product ID không xác định duy nhất tên sản phẩm.
- **DimGeography**: khóa thay thế `GeoKey` trên (City, State, Postal Code), vì một mã bưu chính có thể ứng với nhiều thành phố và một tên thành phố có thể thuộc nhiều bang. Bỏ cột Country vì chỉ có một giá trị (United States).
- **Order ID**: dùng làm degenerate dimension trong bảng fact, vì mỗi đơn có đúng một ngày đặt, ngày giao, Ship Mode và khách hàng.

## 4. Đối chiếu trước / sau

| Chỉ số | Trước | Sau |
|---|---|---|
| Số dòng | 9,800 | 9,799 |
| Số cột | 18 | 26 |
| Số đơn hàng (Order ID phân biệt) | 4,922 | 4,922 |
| Số ô trống | 11 | 0 |
| Sales – Min | 0.444 | 0.444 |
| Sales – Max | 22,638.48 | 22,638.48 |
| Sales – Mean | 230.77 | 230.76 |
| Sales – Median | 54.49 | 54.48 |
| Sales – Tổng | 2,261,536.78 | 2,261,255.41 |

Tổng Sales giảm đúng **281.372**, bằng giá trị của dòng trùng đã xóa.
