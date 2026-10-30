# Hướng dẫn thực hiện đồ án DSS: Phần A (Power BI) và Phần B (Excel)

Tài liệu dành cho các thành viên trong nhóm, gồm hai phần:
- **PHẦN A: Data-Driven DSS (Power BI)**, mục 1–8. Mục 1 tóm tắt dữ liệu đã được xử lý thế nào. Mục 2–8 hướng dẫn từng bước dựng `MaNhom_Dashboard.pbix`.
- **PHẦN B: Model-Driven DSS (Excel)**, mục B0–B8. Hướng dẫn tự gõ công thức và chạy Goal Seek, Data Table, Scenario Manager, Solver cho 4 bài toán trong `MaNhom_Model.xlsx`.

**Cần nhận kèm tài liệu này:**
- Thư mục `data/star/` gồm 6 file CSV: `Fact_Sales`, `Dim_Customer`, `Dim_Product`, `Dim_Geography`, `Dim_Date`, `Dim_ShipMode`.
- `docs/KPI_doi_chieu.md`: số liệu chuẩn để kiểm tra dashboard.
- `output/PhanB_Mau_nhap_lieu.xlsx`: file Excel mẫu có sẵn dữ liệu cho Phần B.
- `docs/PhanB_Mo_hinh.md` và `docs/PhanB_Ket_qua.md`: ý nghĩa mô hình và đáp án đầy đủ của Phần B.
- Tham khảo thêm: `docs/Star_schema.md` (sơ đồ), `docs/Nhat_ky_tien_xu_ly.md` (nhật ký làm sạch), `docs/KPI_va_DAX.md` (bản tóm tắt KPI).

> ⚠️ **Không mở các file CSV bằng Excel rồi lưu lại.** Excel sẽ xóa số 0 đầu của Postal Code (`06824` → `6824`). Nếu chỉ cần xem thì mở bằng Notepad hoặc VS Code.

---

# PHẦN A: Data-Driven DSS với Power BI

## 1. Dữ liệu đã được xử lý như thế nào

### 1.1 Nguồn dữ liệu
- Bộ **Superstore Sales** trên Kaggle, file `train.csv`: 9.800 dòng × 18 cột. Mỗi dòng là một dòng hàng trong một đơn hàng của chuỗi bán lẻ tại Mỹ.
- Đơn hàng từ **03/01/2015 đến 30/12/2018**. Ngày giao muộn nhất là 05/01/2019.
- Chỉ có **một cột số đo là `Sales`** (doanh thu). Không có Profit, Quantity, Discount, nên dashboard không có KPI lợi nhuận.

### 1.2 Quy trình xử lý (bằng Python, có thể chạy lại)

```
train.csv (Kaggle) ──► data/raw/train.csv (bản gốc, chỉ đọc)
                         │  etl/preprocess.py        → làm sạch + thêm cột
                         ▼
               data/clean/Superstore_clean.csv (9.799 dòng)
                         │  etl/build_star_schema.py → tách star schema + 16 kiểm tra
                         ▼
               data/star/*.csv (6 bảng)  ──►  Power BI
```

### 1.3 Các bước làm sạch

| # | Vấn đề phát hiện | Cách xử lý | Số dòng |
|---|---|---|---|
| 1 | 1 dòng trùng hoàn toàn với dòng khác (Row ID 3407) | Xóa, giữ lần xuất hiện đầu | 9.800 → 9.799 |
| 2 | 7 cặp Order ID + Product ID lặp lại nhưng khác Sales | **Giữ lại**: đây là các dòng hàng hợp lệ (mua thêm cùng sản phẩm trong một đơn) | không đổi |
| 3 | Ngày lưu dạng text **ngày/tháng/năm** | Đổi sang kiểu Date, xuất ra dạng `yyyy-mm-dd` để Power BI không nhầm ngày và tháng | không đổi |
| 4 | 429 Postal Code mất số 0 đầu (`6824`) | Đệm đủ 5 chữ số (`06824`) | không đổi |
| 5 | 11 ô trống Postal Code, tất cả ở Burlington, Vermont | Điền `05401` (mã bưu chính USPS của Burlington, VT) | không đổi |
| 6 | 220 tên sản phẩm chứa ký tự khoảng trắng đặc biệt (NBSP); 31 ô có dấu cách kép | Thay bằng dấu cách thường, bỏ khoảng trắng thừa | không đổi |
| 7 | Kiểm tra giá trị vô lý: ngày giao trước ngày đặt, Sales ≤ 0, giá trị phân loại lạ | Không phát hiện trường hợp nào | không đổi |
| 8 | Một số dòng Sales rất lớn (> 500,56 theo quy tắc IQR, 1.145 dòng) | **Giữ lại**, chỉ gắn cờ `Is Sales Outlier` (đây là đơn hàng thật giá trị cao) | không đổi |

Sau xử lý, tổng Sales giảm đúng 281,37, bằng giá trị dòng trùng đã xóa: **2.261.536,78 → 2.261.255,41**.

### 1.4 Kiểm tra tính nhất quán, dùng để quyết định khóa
- Mỗi khách hàng chỉ có một tên và một phân khúc → `Customer ID` dùng làm khóa được.
- **32 Product ID ứng với nhiều tên sản phẩm** → tạo khóa thay thế `ProductKey`.
- **Mã bưu chính 92024 ứng với 2 thành phố, và 55 tên thành phố có mặt ở nhiều bang** → tạo khóa thay thế `GeoKey`.
- Mỗi đơn hàng chỉ có một ngày đặt, một ngày giao, một Ship Mode và một khách hàng → `Order ID` để lại trong bảng Fact.
- Mỗi bang (State) chỉ thuộc một vùng (Region).

### 1.5 Star schema (6 bảng)

| Bảng | Số dòng | Khóa | Các cột |
|---|---|---|---|
| **Fact_Sales** | 9.799 | Row ID | Order ID, OrderDateKey, ShipDateKey, Customer ID, ProductKey, GeoKey, ShipModeKey, **Sales**, **Ship Days**, Is Sales Outlier |
| Dim_Customer | 793 | Customer ID | Customer Name, Segment |
| Dim_Product | 1.893 | ProductKey | Product ID, Product Name, Sub-Category, Category |
| Dim_Geography | 628 | GeoKey | City, State, Postal Code, Region |
| Dim_Date | 1.826 | DateKey (yyyymmdd) | Date, Year, Quarter, Quarter Name, Month, Month Name, YearMonth, Day, DayOfWeek, Weekday Name, IsWeekend. Phạm vi 01/01/2015 → 31/12/2019 |
| Dim_ShipMode | 4 | ShipModeKey | Ship Mode (1 = Same Day, 2 = First Class, 3 = Second Class, 4 = Standard Class) |

Cả 16 kiểm tra toàn vẹn đều đạt: không có khóa trùng, không có dòng Fact mồ côi, tổng Sales sau khi JOIN khớp bảng gốc. Chi tiết xem `docs/Kiem_tra_star_schema.md`.

---

## 2. Nạp dữ liệu vào Power BI

### 2.1 Tắt Auto date/time (làm trước tiên)
**File → Options and settings → Options → CURRENT FILE → Data Load → bỏ chọn "Auto date/time"**.
Lý do: nhóm đã có bảng `Dim_Date` riêng. Nếu để Auto date/time bật, Power BI sẽ tự tạo bảng ngày ẩn, làm file nặng hơn và dễ nhầm khi drill-down.

### 2.2 Nạp 6 file CSV
1. **Home → Get data → Text/CSV**, chọn `Fact_Sales.csv`.
2. Trong cửa sổ xem trước: **File origin = 65001: Unicode (UTF-8)**, **Delimiter = Comma**. Bấm **Transform Data**, không bấm Load.
3. Lặp lại cho 5 file còn lại. Có thể dùng **New Source → Text/CSV** ngay trong Power Query Editor.
4. Đặt tên query **đúng tên file** (`Fact_Sales`, `Dim_Customer`…), vì các công thức DAX bên dưới dùng đúng các tên này.

### 2.3 Đặt kiểu dữ liệu trong Power Query
Power BI tự thêm bước **"Changed Type"** và có thể đoán sai kiểu. Với mỗi query, kiểm tra và sửa theo bảng sau:

| Bảng | Cột | Kiểu đúng |
|---|---|---|
| Fact_Sales | Row ID, OrderDateKey, ShipDateKey, ProductKey, GeoKey, ShipModeKey, Ship Days, Is Sales Outlier | Whole Number |
| Fact_Sales | Sales | Decimal Number (hoặc Fixed decimal) |
| Fact_Sales | Order ID, Customer ID | Text |
| Dim_Geography | **Postal Code** | **Text** ⚠️ (nếu để Whole Number sẽ mất số 0 đầu) |
| Dim_Geography | GeoKey | Whole Number |
| Dim_Product | ProductKey | Whole Number |
| Dim_Date | DateKey, Year, Quarter, Month, Day, DayOfWeek, IsWeekend | Whole Number |
| Dim_Date | **Date** | **Date** |
| Dim_ShipMode | ShipModeKey | Whole Number |
| Các cột còn lại | | Text |

Thêm một cột cho bản đồ: trong `Dim_Geography` chọn **Add Column → Custom Column**, tên `Country`, công thức `"United States"`. Cột này giúp bản đồ không nhầm bang Georgia của Mỹ với nước Georgia.

Bấm **Close & Apply**.

### 2.4 Kiểm tra nhanh sau khi nạp
Ở **Data view**, chọn từng bảng và xem số dòng ở góc dưới bên trái. Phải khớp bảng ở mục 1.5 (Fact_Sales = 9.799 dòng…). Mở `Dim_Geography` và tìm Postal Code bắt đầu bằng `0` (ví dụ `01810`) để chắc số 0 đầu còn nguyên.

---

## 3. Thiết lập mô hình (Model view)

### 3.1 Tạo quan hệ
Power BI có thể tự tạo một số quan hệ. Mở **Manage relationships**, xóa các quan hệ sai, rồi đảm bảo có đúng 6 quan hệ sau:

| # | Từ (bảng Dim, phía "1") | Đến (Fact_Sales, phía "nhiều") | Cardinality | Cross filter | Active |
|---|---|---|---|---|---|
| 1 | Dim_Date[DateKey] | Fact_Sales[OrderDateKey] | One to many | Single | ✅ Active |
| 2 | Dim_Date[DateKey] | Fact_Sales[ShipDateKey] | One to many | Single | ❌ **Inactive** (nét đứt) |
| 3 | Dim_Customer[Customer ID] | Fact_Sales[Customer ID] | One to many | Single | ✅ |
| 4 | Dim_Product[ProductKey] | Fact_Sales[ProductKey] | One to many | Single | ✅ |
| 5 | Dim_Geography[GeoKey] | Fact_Sales[GeoKey] | One to many | Single | ✅ |
| 6 | Dim_ShipMode[ShipModeKey] | Fact_Sales[ShipModeKey] | One to many | Single | ✅ |

Cách tạo quan hệ: kéo cột khóa từ bảng Dim thả vào cột tương ứng ở Fact_Sales. Riêng quan hệ số 2: tạo xong thì bỏ chọn **"Make this relationship active"**. Một cặp bảng chỉ được có một quan hệ active. Quan hệ inactive sẽ được bật trong DAX bằng `USERELATIONSHIP` (mục 4).

### 3.2 Đánh dấu bảng ngày
Chọn bảng `Dim_Date` → **Table tools → Mark as date table** → chọn cột `Date`. Nếu bỏ qua bước này, KPI tăng trưởng YoY sẽ tính sai.

### 3.3 Sắp xếp cột chữ theo cột số (Sort by column)
Nếu không làm bước này, tháng và thứ sẽ bị xếp theo bảng chữ cái (Tháng 1, Tháng 10, Tháng 11, Tháng 2…). Ở Data view, chọn cột rồi vào **Column tools → Sort by column**:

| Cột | Sắp theo |
|---|---|
| Dim_Date[Month Name] | Month |
| Dim_Date[Weekday Name] | DayOfWeek |
| Dim_Date[Quarter Name] | Quarter |
| Dim_ShipMode[Ship Mode] | ShipModeKey (nhanh → chậm) |

### 3.4 Data category cho bản đồ
Chọn cột rồi vào **Column tools → Data category**:
- Dim_Geography[Country] → **Country/Region**
- Dim_Geography[State] → **State or Province**
- Dim_Geography[City] → **City**
- Dim_Geography[Postal Code] → **Postal code**

### 3.5 Tạo phân cấp (hierarchy) để drill-down
Ở Data pane, chuột phải lên cột cấp cao nhất → **Create hierarchy**, rồi chuột phải các cột tiếp theo → **Add to hierarchy**:

| Bảng | Tên hierarchy | Các cấp (từ trên xuống) |
|---|---|---|
| Dim_Date | Lịch | Year → Quarter Name → Month Name → Date |
| Dim_Product | Sản phẩm | Category → Sub-Category → Product Name |
| Dim_Geography | Địa lý | Region → State → City |

### 3.6 Ẩn cột khóa
Ẩn các cột khóa để người xem không kéo nhầm (chuột phải → **Hide in report view**):
- Ở Fact_Sales: OrderDateKey, ShipDateKey, ProductKey, GeoKey, ShipModeKey, Customer ID, Row ID.
- Ở các bảng Dim: DateKey, ProductKey, GeoKey, ShipModeKey.

Lúc vẽ hình, luôn lấy thuộc tính từ bảng Dim (ví dụ Dim_Customer[Segment]), không lấy từ Fact.

---

## 4. Tạo measure DAX

### 4.1 Tạo bảng chứa measure
**Home → Enter data**, đặt tên bảng `_Measures` rồi bấm Load. Bảng sẽ có sẵn một cột `Column1`; ẩn cột này sau khi đã có measure. Để tạo measure: chọn bảng `_Measures` → **Home → New measure** → dán từng công thức bên dưới.

### 4.2 Các measure chính (8 KPI)

```dax
Doanh thu = SUM ( Fact_Sales[Sales] )
```
```dax
Số đơn hàng = DISTINCTCOUNT ( Fact_Sales[Order ID] )
```
```dax
AOV = DIVIDE ( [Doanh thu], [Số đơn hàng] )
```
```dax
Số khách hàng = DISTINCTCOUNT ( Fact_Sales[Customer ID] )
```
```dax
Doanh thu cùng kỳ năm trước =
CALCULATE ( [Doanh thu], SAMEPERIODLASTYEAR ( Dim_Date[Date] ) )
```
```dax
YoY Doanh thu % =
VAR NamTruoc = [Doanh thu cùng kỳ năm trước]
RETURN IF ( NOT ISBLANK ( NamTruoc ), DIVIDE ( [Doanh thu] - NamTruoc, NamTruoc ) )
```
```dax
Tỷ lệ KH quay lại =
VAR KhachHang = VALUES ( Fact_Sales[Customer ID] )
VAR KhachQuayLai =
    COUNTROWS (
        FILTER ( KhachHang, CALCULATE ( DISTINCTCOUNT ( Fact_Sales[Order ID] ) ) >= 2 )
    )
RETURN DIVIDE ( KhachQuayLai, COUNTROWS ( KhachHang ) )
```
```dax
Ngày giao TB =
AVERAGEX ( VALUES ( Fact_Sales[Order ID] ), CALCULATE ( MAX ( Fact_Sales[Ship Days] ) ) )
```
```dax
Tỷ lệ giao nhanh =
VAR DonGiaoNhanh =
    COUNTROWS (
        FILTER ( VALUES ( Fact_Sales[Order ID] ), CALCULATE ( MAX ( Fact_Sales[Ship Days] ) ) <= 2 )
    )
RETURN DIVIDE ( DonGiaoNhanh, [Số đơn hàng] )
```

### 4.3 Các measure phụ trợ

```dax
Tỷ trọng doanh thu = DIVIDE ( [Doanh thu], CALCULATE ( [Doanh thu], ALLSELECTED () ) )
```
```dax
Doanh thu theo ngày giao =
CALCULATE ( [Doanh thu], USERELATIONSHIP ( Fact_Sales[ShipDateKey], Dim_Date[DateKey] ) )
```
```dax
Màu YoY = IF ( [YoY Doanh thu %] < 0, "#D64550", "#2E9E6B" )
```

### 4.4 Định dạng measure
Chọn measure, chỉnh trong **Measure tools**:

| Measure | Format | Số thập phân |
|---|---|---|
| Doanh thu, Doanh thu cùng kỳ năm trước, Doanh thu theo ngày giao, AOV | Currency ($ English (United States)) | 0 cho Doanh thu, 2 cho AOV |
| Số đơn hàng, Số khách hàng | Whole number, có dấu phân cách hàng nghìn | 0 |
| YoY Doanh thu %, Tỷ lệ KH quay lại, Tỷ lệ giao nhanh, Tỷ trọng doanh thu | Percentage | 1 |
| Ngày giao TB | Decimal number | 2 |

### 4.5 Ý nghĩa và lưu ý khi dùng từng KPI

| KPI | Ý nghĩa quản lý | Lưu ý khi hiển thị |
|---|---|---|
| Doanh thu | Quy mô kinh doanh, dùng để đặt mục tiêu | — |
| Số đơn hàng | Lượng giao dịch | Đếm theo Order ID, không phải số dòng |
| AOV | Giá trị trung bình mỗi đơn. AOV giảm thì cần bán kèm hoặc nâng hạng sản phẩm | — |
| Số khách hàng | Độ phủ khách hàng | — |
| YoY Doanh thu % | Đà tăng trưởng | **Chỉ có nghĩa khi chọn 1 năm.** Năm 2015 để trống vì không có năm trước |
| Tỷ lệ KH quay lại | Mức độ trung thành | Phụ thuộc độ dài kỳ: cả 4 năm ra 98,4%, nên hãy xem theo từng năm |
| Ngày giao TB | Năng lực vận hành | Tính theo đơn hàng, không theo dòng |
| Tỷ lệ giao nhanh | Tỷ lệ đơn giao trong vòng 2 ngày | — |

---

## 5. Thiết kế từng trang dashboard

### 5.1 Quy ước chung
- **Kích thước trang**: 16:9 (1280 × 720). **Theme**: View → Themes, chọn một theme sáng và giữ thống nhất cho mọi trang.
- **Bảng màu cố định cho Category** (Format → Data colors): Technology `#2B6CB0`, Furniture `#DD8452`, Office Supplies `#55A868`. Region dùng một bảng màu khác và giữ giống nhau ở mọi trang.
- **Tiêu đề** visual viết tiếng Việt, rõ nghĩa, ví dụ "Doanh thu theo tháng" (không để "Sum of Sales by Date").
- **Thanh trên cùng mỗi trang**: tên trang và các slicer dùng chung.
- **Lọc năm**: `Dim_Date` có cả năm 2019, chỉ để chứa ngày giao. Ở mỗi trang, kéo `Dim_Date[Year]` vào **Filters on this page** và chọn 2015–2018, để không hiện năm 2019 trống.
- **Đồng bộ slicer giữa các trang**: View → Sync slicers, chọn sync slicer Year, Region, Segment, Category trên mọi trang.

### 5.2 Trang 1: Tổng quan (dành cho Giám đốc kinh doanh)

| Vị trí | Visual | Field wells | Ghi chú định dạng |
|---|---|---|---|
| Trên cùng | 4 **Slicer** | Dim_Date[Year] (dạng Tile, single select), Dim_Geography[Region] (Dropdown), Dim_Customer[Segment] (Dropdown), Dim_Product[Category] (Dropdown) | Year để single select cho YoY có nghĩa; thêm tùy chọn "Select all" |
| Hàng 1 | 8 **Card** (hoặc New card visual) | Doanh thu · Số đơn hàng · AOV · Số khách hàng · YoY Doanh thu % · Tỷ lệ KH quay lại · Ngày giao TB · Tỷ lệ giao nhanh | Card YoY: Callout value → Color → fx → Field value → measure `Màu YoY` |
| Giữa, trái | **Line chart**: Doanh thu theo thời gian | X-axis: hierarchy **Lịch** · Y-axis: Doanh thu | Bật drill-down để đi Year → Quarter → Month |
| Giữa, phải | **Line and clustered column chart**: Đơn hàng và AOV theo năm | X-axis: Dim_Date[Year] · Column y-axis: Số đơn hàng · Line y-axis: AOV | Bật Secondary y-axis cho AOV. Đây là chỗ thấy rõ "số đơn tăng, AOV giảm" |
| Dưới | **Clustered column**: Doanh thu và cùng kỳ năm trước | X-axis: Dim_Date[Year] · Y-axis: Doanh thu, Doanh thu cùng kỳ năm trước | Tooltip thêm YoY Doanh thu % |

### 5.3 Trang 2: Sản phẩm (dành cho Quản lý ngành hàng)

| Visual | Field wells | Ghi chú |
|---|---|---|
| **Clustered bar**: Doanh thu theo sản phẩm | Y-axis: hierarchy **Sản phẩm** · X-axis: Doanh thu · Tooltips: Số đơn hàng, AOV, Tỷ trọng doanh thu | Dùng để demo drill-down Category → Sub-Category → Product. Ở mức Product, thêm filter **Top N = 10** theo Doanh thu |
| **Treemap**: Cơ cấu doanh thu | Category: Dim_Product[Category] · Details: Dim_Product[Sub-Category] · Values: Doanh thu | Data labels hiển thị Tỷ trọng doanh thu |
| **Matrix**: Doanh thu theo Category × Năm | Rows: Category → Sub-Category · Columns: Dim_Date[Year] · Values: Doanh thu, YoY Doanh thu % | Conditional formatting → Background color cho Doanh thu; Font color cho YoY (đỏ/xanh). Dùng để demo **Pivot** |
| **Scatter chart**: Số đơn và AOV theo Sub-Category | Values: Dim_Product[Sub-Category] · X: Số đơn hàng · Y: AOV · Size: Doanh thu · Legend: Category | Thấy ngay nhóm "nhiều đơn, giá trị thấp" (Office Supplies) và "ít đơn, giá trị cao" (Technology) |

### 5.4 Trang 3: Khách hàng & Khu vực (dành cho Quản lý vùng và Marketing)

| Visual | Field wells | Ghi chú |
|---|---|---|
| **Filled map**: Doanh thu theo bang | Location: Dim_Geography[Country] rồi Dim_Geography[State] · Tooltips: Doanh thu, Số khách hàng | Fill colors → fx → Gradient theo Doanh thu |
| **Clustered bar**: Doanh thu theo địa lý | Y-axis: hierarchy **Địa lý** · X-axis: Doanh thu | Dùng để demo drill-down Region → State → City và roll-up ngược lại |
| **Donut chart**: Doanh thu theo phân khúc | Legend: Dim_Customer[Segment] · Values: Doanh thu | Detail labels: Category, percent of total |
| **Line chart**: Tỷ lệ KH quay lại theo năm | X-axis: Dim_Date[Year] · Y-axis: Tỷ lệ KH quay lại | Legend: Segment nếu muốn so sánh các phân khúc |
| **Table**: Top 10 khách hàng | Dim_Customer[Customer Name], Dim_Customer[Segment], Doanh thu, Số đơn hàng, AOV | Filter Customer Name → Top N = 10 theo Doanh thu |

### 5.5 Trang 4: Giao hàng (dành cho Quản lý vận hành)

| Visual | Field wells | Ghi chú |
|---|---|---|
| **Line and clustered column chart**: Hiệu quả theo Ship Mode | X-axis: Dim_ShipMode[Ship Mode] · Column: Số đơn hàng · Line: Ngày giao TB | Sort theo Ship Mode (đã sort theo ShipModeKey: nhanh → chậm) |
| **Gauge**: Tỷ lệ giao nhanh | Value: Tỷ lệ giao nhanh · Min 0 · Max 1 · Target tùy nhóm đặt (ví dụ 0,3) | — |
| **Matrix**: Tỷ lệ giao nhanh theo Ship Mode × Region | Rows: Ship Mode · Columns: Region · Values: Tỷ lệ giao nhanh, Ngày giao TB | Conditional formatting dạng heatmap |
| **Line chart**: Doanh thu theo ngày đặt và ngày giao | X-axis: Dim_Date[YearMonth] · Y-axis: Doanh thu, Doanh thu theo ngày giao | Minh họa quan hệ inactive (`USERELATIONSHIP`): hai đường lệch nhau vài ngày, nhìn rõ ở các tháng cuối năm |
| **100% stacked column**: Cơ cấu Ship Mode theo Segment | X-axis: Dim_Customer[Segment] · Y-axis: Số đơn hàng · Legend: Ship Mode | — |

---

## 6. Kịch bản demo OLAP khi thuyết trình

Đề yêu cầu trình bày các thao tác OLAP. Gợi ý kịch bản khoảng 2–3 phút:

| Thao tác | Làm trên visual nào | Cách thực hiện trong Power BI | Câu nói gợi ý |
|---|---|---|---|
| **Slice** (cắt theo 1 chiều) | Trang Tổng quan | Chọn Year = 2018 trên slicer | "Chỉ xem năm 2018: doanh thu 722 nghìn USD, tăng 20,3%" |
| **Dice** (cắt theo nhiều chiều) | Trang Khách hàng & Khu vực | Chọn Region = West và Segment = Corporate | "Thu hẹp vào khách doanh nghiệp ở vùng West" |
| **Drill-down** (đi sâu) | Line chart Lịch; bar Sản phẩm; bar Địa lý | Bật nút mũi tên xuống (Drill mode) rồi click vào một cột, hoặc dùng nút "Expand all down one level" (mũi tên đôi) | "Technology → Phones → sản phẩm bán chạy nhất" |
| **Roll-up** (đi lên) | Cùng các visual trên | Nút mũi tên lên (Drill up) | "Quay về mức vùng để so sánh tổng thể" |
| **Pivot** (xoay chiều) | Matrix trên trang Sản phẩm | Kéo Dim_Date[Year] từ Columns sang Rows, Category từ Rows sang Columns | "Đổi góc nhìn: từ năm theo ngành hàng sang ngành hàng theo năm" |

---

## 7. Kiểm tra số liệu trước khi nộp

Sau khi dựng xong, bỏ hết slicer (chọn tất cả năm 2015–2018) và so với bảng dưới. Bảng đầy đủ xem `docs/KPI_doi_chieu.md`.

**Toàn giai đoạn 2015–2018:**

| Doanh thu | Số đơn | AOV | Số khách hàng | Tỷ lệ KH quay lại | Ngày giao TB | Tỷ lệ giao nhanh |
|---|---|---|---|---|---|---|
| 2.261.255,41 | 4.922 | 459,42 | 793 | 98,4% | 3,96 | 22,1% |

**Theo năm:**

| Năm | Doanh thu | Số đơn | AOV | Số KH | Tỷ lệ KH quay lại | YoY |
|---|---|---|---|---|---|---|
| 2015 | 479.574,84 | 947 | 506,41 | 589 | 43,0% | — |
| 2016 | 459.436,01 | 1.019 | 450,87 | 567 | 54,0% | −4,2% |
| 2017 | 600.192,55 | 1.295 | 463,47 | 635 | 58,4% | 30,6% |
| 2018 | 722.052,02 | 1.661 | 434,71 | 690 | 70,7% | 20,3% |

**Theo Region:** West 710.219,68 · East 669.237,35 · Central 492.646,91 · South 389.151,46.
**Theo Category:** Technology 827.455,87 · Furniture 728.377,20 · Office Supplies 705.422,33.

Checklist:
- [ ] Số dòng 6 bảng khớp mục 1.5.
- [ ] Postal Code giữ số 0 đầu.
- [ ] Có 6 quan hệ; chỉ quan hệ ShipDateKey là inactive.
- [ ] Đã Mark as date table cho `Dim_Date`.
- [ ] Tháng, thứ và Ship Mode hiển thị đúng thứ tự.
- [ ] Các card khớp bảng ở trên.
- [ ] Không còn năm 2019 trống trên các biểu đồ.
- [ ] Mọi tiêu đề visual viết tiếng Việt và rõ nghĩa.
- [ ] Lưu file với tên `MaNhom_Dashboard.pbix`.

---

## 8. Lỗi thường gặp

| Hiện tượng | Nguyên nhân | Cách sửa |
|---|---|---|
| Postal Code hiện `6824` thay vì `06824` | Bước "Changed Type" đổi cột sang số | Power Query → đổi lại kiểu **Text** (chọn **Replace current**), hoặc xóa bước Changed Type rồi đặt kiểu thủ công |
| YoY trống hoặc sai | Chưa Mark as date table, hoặc đang xem năm 2015 | Làm mục 3.2; năm 2015 trống là đúng |
| Card YoY ra số lạ khi không chọn năm | Đang so cả giai đoạn 2015–2018 với 2014–2017 | Chỉ đọc YoY khi slicer đang chọn **1 năm** |
| Tháng hiện theo thứ tự Tháng 1, Tháng 10, Tháng 11… | Chưa sort by column | Làm mục 3.3 |
| Biểu đồ có năm 2019 trống | Dim_Date kéo dài đến 2019 | Thêm page filter Year 2015–2018 |
| Bản đồ đặt một bang sai chỗ (ví dụ Georgia) | Thiếu ngữ cảnh quốc gia | Thêm cột `Country` (mục 2.3) và đặt Country trước State trong ô Location |
| Số đơn hàng lớn hơn 4.922 | Đếm số dòng thay vì đếm Order ID | Dùng measure `Số đơn hàng` (DISTINCTCOUNT), không kéo cột Order ID rồi để Count |
| Kéo thuộc tính vào visual nhưng số không lọc theo | Lấy cột từ bảng Fact, hoặc quan hệ bị tắt hay sai chiều | Lấy thuộc tính từ bảng Dim; kiểm tra lại mục 3.1 |

---

# PHẦN B: Model-Driven DSS với Excel

Hướng dẫn này để nhóm **tự gõ công thức và tự chạy các công cụ Excel** cho 4 bài toán Phần B.

**Bắt đầu từ file mẫu `output/PhanB_Mau_nhap_lieu.xlsx`.** Mở file, Save As thành `MaNhom_Model.xlsx`, rồi làm theo từ mục B0.

| Trong file mẫu đã có sẵn | Nhóm tự làm |
|---|---|
| 6 sheet, tiêu đề, nhãn, bố cục, màu ô, định dạng số | Gõ **mọi công thức**: các ô 🟦 xanh dương đang trống, và cột D–E của bảng Region trong `Du_lieu_2018` |
| Số liệu thật 2018 (`Du_lieu_2018`), giả định (`Gia_dinh`) | Đặt tên ô (nếu muốn) |
| Tham số 🟨 và giá trị khởi đầu của biến quyết định 🟩 | Chạy **Goal Seek** và ghi kết quả vào cột E của bảng Goal Seek |
| Bảng Goal Seek đã ghi sẵn câu hỏi, Set cell, To value, By changing | Tạo **Data Table** (giá trị đầu vào của bảng đã có sẵn, chỉ cần gõ ô công thức ở góc rồi chạy công cụ) |
| | Tạo **Scenario** và Summary; chạy **Solver** và báo cáo |

Trong các bảng bên dưới, giá trị **không bắt đầu bằng `=`** là đã có sẵn trong file mẫu (chỉ cần kiểm tra lại). Giá trị **bắt đầu bằng `=`** là công thức nhóm phải gõ.

> Muốn làm hoàn toàn từ workbook trống thì vẫn được: gõ cả các giá trị trong bảng.

- Ý nghĩa của từng mô hình và cách viết báo cáo: xem `docs/PhanB_Mo_hinh.md`.
- Đáp án để tự kiểm tra: in đậm cuối mỗi bước, và đầy đủ trong `docs/PhanB_Ket_qua.md`.
- File `output/PhanB_Dap_an_tham_khao.xlsx` là bản mẫu để so khi bị lệch. **Nộp file nhóm tự làm**, không nộp bản mẫu.

Thời gian ước tính khoảng 1,5–2 giờ khi dùng file mẫu. Có thể chia mỗi thành viên một bài, nhưng **mọi người nên tự làm thử ít nhất một lần Solver và một lần Data Table**, vì khi thuyết trình sẽ phải demo.

---

## B0. Chuẩn bị

1. **Bật Solver**: File → Options → Add-ins → ô *Manage* chọn **Excel Add-ins** → Go → tick **Solver Add-in** → OK. Nút **Solver** sẽ xuất hiện ở tab Data, bên phải.
2. **Mở file mẫu** → File → Save As → `MaNhom_Model.xlsx`. Workbook gồm 6 sheet: `Du_lieu_2018`, `Gia_dinh`, `B1_Hoa_von`, `B2_Marketing`, `B3_Nhap_hang`, `B4_Giao_hang` (cùng sheet `Huong_dan`).
3. **Quy ước màu ô** (đã tô sẵn trong file mẫu), để giảng viên nhìn là hiểu mô hình:
   - 🟨 **Vàng**: tham số đầu vào, được phép thay đổi.
   - 🟩 **Xanh lá**: biến quyết định, Solver sẽ thay đổi các ô này.
   - 🟦 **Xanh dương**: kết quả, mục tiêu (là công thức).
4. **Định dạng số** đã cài sẵn: tiền `#,##0`; tỷ lệ dạng Percent. Ô công thức sẽ tự hiển thị đúng định dạng khi gõ xong.
5. Gõ công thức **đúng ô như hướng dẫn**. Khi cần tham chiếu sheet khác, có thể gõ `=` rồi click sang sheet đó chọn ô.

> Dòng 1–2 của mỗi sheet ghi tiêu đề và câu hỏi quyết định, ví dụ A1: `BÀI 1 – HÒA VỐN VÀ LỢI NHUẬN MỤC TIÊU NĂM 2019`. Mỗi khối có một dòng tiêu đề tô màu nhạt, ví dụ A4: `THAM SỐ`.

---

## B1. Sheet `Du_lieu_2018`: số liệu thật năm 2018

**Đã có sẵn trong file mẫu.** Đây là dữ liệu đầu vào, lấy từ `docs/KPI_doi_chieu.md` và star schema của Phần A. Nhóm chỉ cần gõ **2 công thức ở D12 và E12** (bảng Region) rồi kéo xuống dòng 15. Các bảng dưới đây dùng để đối chiếu:

| Ô | Nội dung (cột A) | Giá trị (cột B) |
|---|---|---|
| A5 / B5 | Doanh thu 2018 | 722052.02 |
| A6 / B6 | Số đơn hàng 2018 | 1661 |
| A7 / B7 | Số đơn hàng 2017 | 1295 |
| A8 / B8 | AOV 2018 | 434.71 |

**Doanh thu theo Region**: tiêu đề cột ở dòng 11 (Region, DT 2017, DT 2018, Tăng trưởng, Tỷ trọng 2018), dữ liệu ở dòng 12–15:

| Dòng | A | B (DT 2017) | C (DT 2018) | D | E |
|---|---|---|---|---|---|
| 12 | Central | 145673.88 | 141627.34 | `=C12/B12-1` | `=C12/SUM($C$12:$C$15)` |
| 13 | East | 178511.54 | 210129.19 | kéo công thức xuống | kéo xuống |
| 14 | South | 93535.90 | 122164.57 | | |
| 15 | West | 182471.23 | 248130.93 | | |

**Doanh thu theo Sub-Category**: tiêu đề ở dòng 18 (Sub-Category, Category, DT 2018), dữ liệu ở dòng 19–35. **Giữ đúng thứ tự này**, vì sheet Gia_dinh và Bài 3 dùng cùng thứ tự:

| Dòng | A | B | C |
|---|---|---|---|
| 19 | Phones | Technology | 104685.66 |
| 20 | Chairs | Furniture | 93191.27 |
| 21 | Binders | Office Supplies | 71721.26 |
| 22 | Storage | Office Supplies | 68072.70 |
| 23 | Copiers | Technology | 62899.39 |
| 24 | Tables | Furniture | 60893.54 |
| 25 | Accessories | Technology | 58240.97 |
| 26 | Machines | Technology | 43544.68 |
| 27 | Appliances | Office Supplies | 40212.72 |
| 28 | Bookcases | Furniture | 30024.28 |
| 29 | Furnishings | Furniture | 28204.69 |
| 30 | Paper | Office Supplies | 27504.78 |
| 31 | Supplies | Office Supplies | 15952.11 |
| 32 | Art | Office Supplies | 8825.63 |
| 33 | Labels | Office Supplies | 3844.70 |
| 34 | Envelopes | Office Supplies | 3378.57 |
| 35 | Fasteners | Office Supplies | 855.07 |

**Giao hàng theo Ship Mode**: tiêu đề ở dòng 38 (Ship Mode, Số đơn, Ngày giao TB, Tỷ lệ giao ≤2 ngày), dữ liệu ở dòng 39–42:

| Dòng | A | B (Số đơn) | C (Ngày giao TB) | D (Tỷ lệ giao ≤ 2 ngày) |
|---|---|---|---|---|
| 39 | Same Day | 88 | 0.0568 | 1 |
| 40 | First Class | 280 | 2.1893 | 0.5821 |
| 41 | Second Class | 318 | 3.2516 | 0.3679 |
| 42 | Standard Class | 975 | 5.0185 | 0 |

---

## B2. Sheet `Gia_dinh`: giả định dùng chung

**Đã có sẵn trong file mẫu, không cần gõ gì.** Lý do của từng giả định nằm ở cột D và trong `docs/PhanB_Mo_hinh.md` mục 0.3.

| Ô | Nội dung | Giá trị |
|---|---|---|
| A6 / B6 | Tăng trưởng số đơn 2019 | 0.2 |
| A7 / B7 | Phí thanh toán | 0.02 |
| A8 / B8 | Chi phí xử lý đơn (USD/đơn) | 10 |

**Chi phí giao hàng**: tiêu đề ở dòng 11, dữ liệu ở dòng 12–15:

| Ô | A | B (USD/đơn) |
|---|---|---|
| 12 | Same Day | 30 |
| 13 | First Class | 18 |
| 14 | Second Class | 12 |
| 15 | Standard Class | 7 |

**Biên lợi nhuận gộp và thể tích**: tiêu đề ở dòng 18 (Sub-Category, Biên LN gộp, m³ / 1.000 USD), dữ liệu ở dòng 19–35, **cùng thứ tự với Du_lieu_2018**:

| Dòng | A | B (Biên) | C (m³/1.000 USD) | | Dòng | A | B | C |
|---|---|---|---|---|---|---|---|---|
| 19 | Phones | 0.14 | 0.2 | | 28 | Bookcases | 0.10 | 3.0 |
| 20 | Chairs | 0.18 | 2.5 | | 29 | Furnishings | 0.28 | 1.2 |
| 21 | Binders | 0.35 | 0.6 | | 30 | Paper | 0.38 | 0.8 |
| 22 | Storage | 0.20 | 2.0 | | 31 | Supplies | 0.18 | 0.5 |
| 23 | Copiers | 0.30 | 1.0 | | 32 | Art | 0.30 | 0.4 |
| 24 | Tables | 0.08 | 3.0 | | 33 | Labels | 0.40 | 0.2 |
| 25 | Accessories | 0.22 | 0.3 | | 34 | Envelopes | 0.35 | 0.4 |
| 26 | Machines | 0.10 | 1.2 | | 35 | Fasteners | 0.30 | 0.2 |
| 27 | Appliances | 0.22 | 1.5 | | | | | |

---

## B3. Bài 1 (`B1_Hoa_von`): Hòa vốn và lợi nhuận mục tiêu

### B3.1 Dựng mô hình

| Ô | Nội dung (cột A) | Cột B: giá trị / công thức | Màu |
|---|---|---|---|
| A6 | Số đơn hàng 2019 (biến quyết định) | `1993` | 🟩 |
| A7 | AOV (biến quyết định) | `434.71` | 🟩 |
| A8 | Biên lợi nhuận gộp bình quân | `=SUMPRODUCT(Du_lieu_2018!C19:C35,Gia_dinh!B19:B35)/SUM(Du_lieu_2018!C19:C35)` | 🟦 |
| A9 | Phí thanh toán | `=Gia_dinh!B7` | 🟦 |
| A10 | Chi phí xử lý / đơn | `=Gia_dinh!B8` | 🟦 |
| A11 | Chi phí giao hàng TB / đơn | `=SUMPRODUCT(Du_lieu_2018!B39:B42,Gia_dinh!B12:B15)/SUM(Du_lieu_2018!B39:B42)` | 🟦 |
| A12 | Chi phí cố định năm | `70000` | 🟨 |
| A13 | Lợi nhuận mục tiêu | `40000` | 🟨 |
| A16 | Doanh thu 2019 | `=B6*B7` | 🟦 |
| A17 | Lợi nhuận góp / đơn | `=B7*(B8-B9)-B10-B11` | 🟦 |
| A18 | Tổng lợi nhuận góp | `=B6*B17` | 🟦 |
| A19 | **LỢI NHUẬN 2019** | `=B18-B12` | 🟦 |
| A20 | Số đơn hòa vốn (công thức) | `=B12/B17` | 🟦 |
| A21 | Biên an toàn | `=(B6-B20)/B6` | 🟦 |

Giải thích: 1.993 = 1.661 × 1,2. B8 là biên bình quân gia quyền theo doanh thu từng Sub-Category; B11 là chi phí giao bình quân theo cơ cấu Ship Mode năm 2018.

✅ **Kiểm tra**: B8 = 20,7%; B11 = 11,03; B17 = 60,43; **B19 = 50.442**; B20 = 1.158.

> **Nên đặt tên ô** (click ô → gõ tên vào Name Box bên trái thanh công thức → Enter), để bảng Scenario Summary hiện tên thay vì `$B$6`: B6 → `SoDon`, B7 → `AOV`, B12 → `ChiPhiCoDinh`, B16 → `DoanhThu`, B17 → `LNGopMoiDon`, B19 → `LoiNhuan`, B20 → `DiemHoaVon`, B21 → `BienAnToan`.

### B3.2 Goal Seek (3 câu hỏi)
Data → What-If Analysis → **Goal Seek**:

| Câu hỏi | Set cell | To value | By changing cell | Đáp án |
|---|---|---|---|---|
| Cần bao nhiêu đơn để hòa vốn? | `B19` | `0` | `B6` | **1.158 đơn** |
| Cần bao nhiêu đơn để lãi 40.000 USD? | `B19` | `40000` | `B6` | **1.820 đơn** |
| Với 1.993 đơn, AOV tối thiểu để lãi 40.000? | `B19` | `40000` | `B7` | **406,75 USD** |

Sau mỗi lần: chụp màn hình hộp thoại kết quả cho báo cáo. Ghi kết quả vào bảng Goal Seek ở dòng 24–27 (cột A câu hỏi, cột B Set cell, C To value, D By changing, E Kết quả), rồi bấm **Cancel** hoặc gõ lại giá trị gốc (B6 = 1993, B7 = 434.71).

### B3.3 Data Table 2 chiều (lợi nhuận theo số đơn × AOV)
1. A31: `=B19`, ô góc bắt buộc phải trỏ tới kết quả.
2. B31:I31 (hàng ngang, giá trị AOV, đã có sẵn): `360, 380, 400, 420, 440, 460, 480, 500`.
3. A32:A38 (cột dọc, số đơn, đã có sẵn): `1200, 1400, 1600, 1800, 2000, 2200, 2400`.
4. Chọn vùng **A31:I38** → Data → What-If Analysis → **Data Table** → **Row input cell: `B7`**, **Column input cell: `B6`** → OK.
5. Tô màu có điều kiện (Conditional Formatting → Color Scales) để thấy ngay vùng lỗ.

✅ **Kiểm tra**: ô (2000, 440) = 52.848; ô (1200, 360) = −14.281.

### B3.4 Scenario Manager
Data → What-If Analysis → **Scenario Manager** → **Add**. Changing cells: `B6,B7,B12` (giữ Ctrl để chọn nhiều ô):

| Tên kịch bản | B6 (số đơn) | B7 (AOV) | B12 (chi phí cố định) |
|---|---|---|---|
| Bi quan | 1661 | 400 | 75000 |
| Cơ sở | 1993 | 434.71 | 70000 |
| Lạc quan | 2159 | 450 | 68000 |

Sau đó bấm **Summary** → Scenario summary → Result cells: `B16,B17,B19,B20,B21` → OK. Excel tạo sheet "Scenario Summary". Cuối cùng chọn "Cơ sở" → **Show** để trả lại giá trị gốc.

✅ **Kiểm tra**: lợi nhuận Bi quan 14.575, Cơ sở 50.442, Lạc quan 68.660.

---

## B4. Bài 2 (`B2_Marketing`): Phân bổ ngân sách marketing (Solver GRG Nonlinear)

### B4.1 Dựng mô hình

**Tham số:**

| Ô | Nội dung | B | Màu |
|---|---|---|---|
| A6 | Tổng ngân sách marketing | `100000` | 🟨 |
| A7 | Ngân sách tối thiểu mỗi vùng | `10000` | 🟨 |
| A8 | Ngân sách tối đa mỗi vùng | `40000` | 🟨 |
| A9 | Hệ số bão hòa (scale) | `25000` | 🟨 |
| A10 | Mục tiêu DT tăng thêm | `110000` | 🟨 |

**Bảng mô hình**: tiêu đề ở dòng 13 (Region, DT 2018, α, Ngân sách, DT tăng thêm, ROAS, DT biên / 1 USD, Tăng trưởng 17→18). Gõ dòng 14, rồi kéo cột B, E, F, G, H xuống đến dòng 17:

| Cột | Dòng 14 (Central) | Ghi chú |
|---|---|---|
| A | Central (15: East, 16: South, 17: West) | |
| B | `=Du_lieu_2018!C12` | Kéo xuống thành C13, C14, C15 |
| C | `0.12` (East `0.18`, South `0.40`, West `0.22`) | 🟨 α: tiềm năng tăng trưởng tối đa |
| D | `25000` (cả 4 dòng) | 🟩 **biến quyết định**, giá trị khởi đầu |
| E | `=B14*C14*(1-EXP(-D14/$B$9))` | 🟦 DT tăng thêm, lợi suất giảm dần |
| F | `=E14/D14` | ROAS |
| G | `=B14*C14/$B$9*EXP(-D14/$B$9)` | Doanh thu biên của 1 USD chi thêm |
| H | `=Du_lieu_2018!D12` | Bằng chứng để chọn α |

Dòng 18 (TỔNG): D18 `=SUM(D14:D17)`, **E18 `=SUM(E14:E17)`** (🟦 hàm mục tiêu), F18 `=E18/D18`.

**Phương án so sánh: chia đều**

| Ô | Nội dung | B |
|---|---|---|
| A22 | Tổng ngân sách (chia đều 4 vùng) | `100000` 🟨 |
| A23 | DT tăng thêm khi chia đều | `=(1-EXP(-(B22/4)/B9))*SUMPRODUCT(B14:B17,C14:C17)` |
| A24 | ROAS khi chia đều | `=B23/B22` |
| A25 | Lợi ích của tối ưu so với chia đều | `=E18-B23` |

✅ **Kiểm tra** (trước khi chạy Solver, mỗi vùng 25.000): E18 = B23 = **100.048**.

### B4.2 Solver
Data → **Solver**:
- **Set Objective**: `$E$18`, chọn **Max**.
- **By Changing Variable Cells**: `$D$14:$D$17`.
- **Subject to the Constraints** (bấm Add cho từng dòng):
  - `$D$18 <= $B$6` (tổng không vượt ngân sách)
  - `$D$14:$D$17 >= $B$7` (mỗi vùng tối thiểu)
  - `$D$14:$D$17 <= $B$8` (mỗi vùng tối đa)
- Tick **Make Unconstrained Variables Non-Negative**.
- **Select a Solving Method: GRG Nonlinear**, vì hàm EXP là phi tuyến.
- Bấm **Solve**, chọn **Keep Solver Solution**, ở khung Reports chọn **Answer** và **Sensitivity** → OK.

✅ **Đáp án**: D14 (Central) = **10.000**, D15 (East) ≈ **24.808**, D16 (South) ≈ **31.212**, D17 (West) ≈ **33.980**; **E18 ≈ 104.816**. Kết quả GRG có thể lệch vài chục USD ở từng vùng, điều này bình thường.

### B4.3 Goal Seek
Set cell `B23`, To value `110000`, By changing `B22` → **118.745**. Ghi vào bảng kết quả (dòng 28–29), rồi trả B22 về 100000.

### B4.4 Data Table 1 chiều
1. A33 đã ghi sẵn chữ `(công thức →)`. Gõ B33: `=B23`; C33: `=B24`.
2. A34:A40 (đã có sẵn): `40000, 60000, 80000, 100000, 120000, 140000, 160000`.
3. Chọn **A33:C40** → Data Table → để trống Row input, **Column input cell: `B22`** → OK.

✅ **Kiểm tra**: ROAS giảm từ 1,30 (40.000) xuống 0,79 (160.000).

### B4.5 Scenario Manager (mỗi kịch bản là một nghiệm Solver)
Lặp lại cho 3 mức ngân sách, theo thứ tự:
1. Gõ B6 = `60000` → mở Solver → **Solve** → Keep Solver Solution.
2. Scenario Manager → Add → Name `Ngân sách thấp`, Changing cells: `B6,D14:D17` → OK. Excel tự điền giá trị hiện tại (chính là nghiệm vừa giải) → OK.
3. Lặp lại với B6 = `100000` (tên `Ngân sách cơ sở`) và B6 = `140000` (tên `Ngân sách cao`).
4. Summary → Result cells `D18,E18,F18`.
5. Show "Ngân sách cơ sở" để file trở về phương án chính.

✅ **Kiểm tra**: E18 lần lượt là 75.176 / 104.816 / 122.113.

---

## B5. Bài 3 (`B3_Nhap_hang`): Kế hoạch nhập hàng (Solver Simplex LP)

### B5.1 Dựng mô hình

**Tham số:**

| Ô | Nội dung | B | Màu |
|---|---|---|---|
| A6 | Tăng trưởng nhu cầu 2019 | `0.2` | 🟨 |
| A7 | Ngân sách nhập hàng (giá vốn) | `520000` | 🟨 |
| A8 | Sức chứa kho (m³) | `650` | 🟨 |
| A9 | Mức đáp ứng tối thiểu | `0.5` | 🟨 |
| A10 | Tỷ lệ nhu cầu thực tế so với dự báo | `1` | 🟨 (dùng cho Data Table) |

**Bảng mô hình**: tiêu đề ở dòng 13. A14:A30 là 17 Sub-Category, **cùng thứ tự** như Du_lieu_2018 (Phones … Fasteners); B14:B30 là Category. Cột A, B, I đã có sẵn. Gõ các công thức dòng 14 rồi **chọn C14:H14 kéo xuống dòng 30, và J14:N14 kéo xuống dòng 30**:

| Cột | Tiêu đề | Dòng 14 | Ý nghĩa |
|---|---|---|---|
| C | DT 2018 | `=Du_lieu_2018!C19` | |
| D | Biên LN | `=Gia_dinh!B19` | |
| E | m³/1.000$ | `=Gia_dinh!C19` | |
| F | Nhu cầu 2019 (DT) | `=C14*(1+$B$6)` | |
| G | Nhập tối đa (giá vốn) | `=F14*(1-D14)` | Giá vốn của toàn bộ nhu cầu |
| H | Nhập tối thiểu | `=G14*$B$9` | |
| I | **GIÁ TRỊ NHẬP** | đã có sẵn giá trị khởi đầu 30.588 (là số, không phải công thức) | 🟩 biến quyết định |
| J | DT kế hoạch | `=I14/(1-D14)` | Giá vốn → giá bán |
| K | LN gộp kế hoạch | `=J14-I14` | |
| L | Thể tích (m³) | `=I14*E14/1000` | |
| M | DT thực hiện | `=MIN(J14,F14*$B$10)` | Chỉ bán được tối đa bằng nhu cầu thực tế |
| N | LN thực hiện | `=M14-I14` | |

Dòng 31 (TỔNG): F31:N31, mỗi ô là `=SUM(…14:…30)` (F31 `=SUM(F14:F30)` rồi kéo sang phải đến N31). **K31** là hàm mục tiêu 🟦.

**Chỉ số tổng hợp**: B35 `=I31/B7` (dùng ngân sách); B36 `=L31/B8` (dùng kho); B37 `=J31/F31` (đáp ứng nhu cầu); B38 `=K31/J31` (biên kế hoạch); B39 `=SUMPRODUCT(J14:J30-M14:M30,1-D14:D30)` (tồn kho dư).

✅ **Kiểm tra**: F31 = 866.462 (doanh thu nhu cầu 2019); G31 = 686.762 (giá vốn toàn bộ nhu cầu).

### B5.2 Solver
- **Set Objective**: `$K$31`, **Max**.
- **By Changing**: `$I$14:$I$30`.
- **Constraints**:
  - `$I$31 <= $B$7` (ngân sách)
  - `$L$31 <= $B$8` (sức chứa kho)
  - `$I$14:$I$30 >= $H$14:$H$30` (đáp ứng tối thiểu)
  - `$I$14:$I$30 <= $G$14:$G$30` (không nhập quá nhu cầu)
- **Solving Method: Simplex LP**, vì mô hình tuyến tính. Hàm mục tiêu K31 không dùng MIN nên vẫn tuyến tính.
- Solve → Keep → Reports: **Answer + Sensitivity**.

✅ **Đáp án**: **K31 = 154.433**; I31 = 520.000; L31 = 650. Trong Sensitivity Report, *Shadow Price* của ràng buộc ngân sách là **0,153** và của kho là **48,4**. Binders, Copiers, Accessories… được nhập 100% nhu cầu; Chairs, Tables, Machines, Bookcases chỉ ở mức tối thiểu.

### B5.3 Data Table 1 chiều (rủi ro khi nhu cầu thấp hơn dự báo)
Làm **sau khi** đã có nghiệm Solver:
1. Gõ B43 `=N31`; C43 `=B39` (A43 đã có chữ).
2. A44:A51 (đã có sẵn): `0.8, 0.85, 0.9, 0.95, 1, 1.05, 1.1, 1.2`.
3. Chọn **A43:C51** → Data Table → **Column input cell: `B10`**.

✅ **Kiểm tra**: ở 90% lợi nhuận là 109.408 và tồn kho dư 32.642; từ 100% trở lên lợi nhuận là 154.433.

### B5.4 Scenario Manager
Như Bài 2 (giải Solver cho từng kịch bản rồi Add). Changing cells: `B7,B8,I14:I30`:

| Tên | B7 | B8 | ✅ K31 sau Solve |
|---|---|---|---|
| Thắt chặt | 460000 | 600 | 139.939 |
| Cơ sở | 520000 | 650 | 154.433 |
| Mở rộng | 600000 | 700 | 165.769 |

Summary → Result cells `I31,L31,K31,B37,B35,B36`. Cuối cùng Show "Cơ sở".

---

## B6. Bài 4 (`B4_Giao_hang`): Cơ cấu phương thức giao hàng (Solver Simplex LP)

### B6.1 Dựng mô hình

**Tham số:**

| Ô | Nội dung | B | Màu |
|---|---|---|---|
| A6 | Số đơn dự báo 2019 | `=ROUND(Du_lieu_2018!B6*(1+Gia_dinh!B6),0)` → 1993 | 🟦 |
| A7 | Mục tiêu tỷ lệ giao nhanh (≤ 2 ngày) | `0.3` | 🟨 |
| A8 | Năng lực Same Day (đơn) | `150` | 🟨 |
| A9 | Ngày giao TB tối đa | `4` | 🟨 |
| A10 | Tỷ trọng tối thiểu mỗi phương thức | `0.05` | 🟨 |

**Bảng mô hình**: tiêu đề ở dòng 13. A14:A17 và cột G đã có sẵn. Gõ công thức dòng 14 rồi kéo xuống dòng 17:

| Cột | Tiêu đề | Dòng 14 |
|---|---|---|
| B | Chi phí/đơn | `=Gia_dinh!B12` |
| C | Tỷ lệ giao nhanh | `=Du_lieu_2018!D39` |
| D | Ngày giao TB | `=Du_lieu_2018!C39` |
| E | Số đơn 2018 | `=Du_lieu_2018!B39` |
| F | Tỷ trọng 2018 | `=E14/SUM($E$14:$E$17)` |
| G | **SỐ ĐƠN 2019** | đã có sẵn giá trị khởi đầu (theo tỷ trọng 2018) 🟩 biến quyết định |
| H | Chi phí | `=G14*B14` |
| I | Đơn giao nhanh | `=G14*C14` |
| J | Ngày × đơn | `=G14*D14` |
| K | Tỷ trọng mới | `=G14/$G$18` |

Dòng 18 (TỔNG): G18 `=SUM(G14:G17)`, **H18 `=SUM(H14:H17)`** (🟦 hàm mục tiêu), I18 `=SUM(I14:I17)`, J18 `=SUM(J14:J17)`.

**Vế phải ràng buộc và chỉ số:**

| Ô | Nội dung | B |
|---|---|---|
| A22 | Số đơn giao nhanh tối thiểu | `=B7*B6` |
| A23 | Trần tổng ngày × đơn | `=B9*B6` |
| A24 | Số đơn tối thiểu mỗi phương thức | `=B10*B6` |
| A25 | Tỷ lệ giao nhanh đạt được | `=I18/G18` |
| A26 | Ngày giao TB đạt được | `=J18/G18` |
| A27 | Chi phí giao TB / đơn | `=H18/G18` |
| A28 | Chi phí nếu giữ cơ cấu 2018 | `=B6*SUMPRODUCT(F14:F17,B14:B17)` |
| A29 | Chi phí tăng thêm so với 2018 | `=H18-B28` |

**Phương án đơn giản: chỉ tăng First Class** (dùng cho Goal Seek và Data Table):

| Ô | Nội dung | B |
|---|---|---|
| A32 | Tỷ trọng First Class | `0.1686` (gõ số, bằng tỷ trọng 2018) 🟨 |
| A33 | Tỷ trọng Standard Class | `=1-F14-F16-B32` |
| A34 | Tỷ lệ giao nhanh | `=F14*C14+B32*C15+F16*C16+B33*C17` |
| A35 | Tổng chi phí giao | `=B6*(F14*B14+B32*B15+F16*B16+B33*B17)` |

✅ **Kiểm tra**: B28 = 21.983; B34 = 22,2%.

### B6.2 Solver
- **Set Objective**: `$H$18`, chọn **Min**.
- **By Changing**: `$G$14:$G$17`.
- **Constraints**:
  - `$G$18 = $B$6` (đủ số đơn dự báo)
  - `$I$18 >= $B$22` (đạt mục tiêu giao nhanh)
  - `$J$18 <= $B$23` (ngày giao TB ≤ 4)
  - `$G$14 <= $B$8` (năng lực Same Day)
  - `$G$14:$G$17 >= $B$24` (mỗi phương thức ≥ 5%)
- **Simplex LP** → Solve → Keep → Reports: Answer + Sensitivity.

✅ **Đáp án**: G14 ≈ 99,65; G15 ≈ 99,65; G16 ≈ **1.196,6**; G17 ≈ 597,1; **H18 ≈ 23.322**; B29 ≈ 1.339. Trong Sensitivity Report, *Shadow Price* của ràng buộc `$I$18` là **13,59**: mỗi đơn giao nhanh thêm tốn 13,59 USD.

### B6.3 Goal Seek
Set cell `B34`, To value `0.3`, By changing `B32` → **30,3%**. Ghi kết quả vào dòng 38–39, rồi trả B32 về 0.1686.

### B6.4 Data Table 1 chiều
1. Gõ B43 `=B34`; C43 `=B35` (A43 đã có chữ).
2. A44:A52 (đã có sẵn): `0.10, 0.15, 0.20, 0.25, 0.30, 0.35, 0.40, 0.45, 0.50`.
3. Chọn **A43:C52** → Data Table → **Column input cell: `B32`**.

✅ **Kiểm tra**: ở 30% First Class thì tỷ lệ giao nhanh 29,8% và chi phí 24.864.

### B6.5 Scenario Manager
Giải Solver cho từng mục tiêu rồi Add. Changing cells: `B7,G14:G17`:

| Tên | B7 | ✅ H18 sau Solve |
|---|---|---|
| Mục tiêu 25% | 0.25 | 21.968 |
| Mục tiêu 30% | 0.30 | 23.322 |
| Mục tiêu 35% | 0.35 | 24.677 |

Summary → Result cells `H18,B29,B25,B26,B27`. Show "Mục tiêu 30%".

---

## B7. Lỗi thường gặp

| Hiện tượng | Cách sửa |
|---|---|
| Data Table hiện cùng một số ở mọi ô | Formulas → Calculation Options → **Automatic** (không chọn "Automatic except for data tables"), hoặc bấm F9 |
| Data Table báo "Input cell reference is not valid" | Input cell phải nằm **cùng sheet** với bảng; Bài 1 phải điền đúng Row = B7, Column = B6 |
| Solver báo "The linearity conditions required by this LP Solver are not satisfied" | Bài 3 và 4 phải dùng Simplex với hàm mục tiêu tuyến tính. Kiểm tra K31 và H18 không phụ thuộc cột M/N (có hàm MIN). Bài 2 phải chọn **GRG Nonlinear** |
| Solver ra nghiệm khác đáp án | Kiểm tra lại từng ràng buộc (dấu ≤ hay ≥); ô biến phải là **số gõ tay**, không phải công thức; đã tick Non-Negative |
| Scenario cảnh báo sẽ thay công thức bằng hằng số | Changing cells chỉ được là ô **số gõ tay** (ô vàng, ô xanh lá), không chọn ô công thức |
| Scenario Summary hiện `$B$6` thay vì tên | Đặt tên ô bằng Name Box **trước** khi bấm Summary |
| Sai lệch nhỏ vài đơn vị so với đáp án | Bình thường, do làm tròn số liệu đầu vào hoặc sai số của GRG |

---

## B8. Câu hỏi vấn đáp có thể gặp

1. **Biến quyết định, biến kết quả và ràng buộc của bài X là gì?** Đọc đúng các ô xanh lá, ô mục tiêu xanh dương và danh sách ràng buộc trong Solver.
2. **Vì sao Bài 2 dùng GRG Nonlinear còn Bài 3, 4 dùng Simplex LP?** Bài 2 có hàm EXP (phi tuyến, lợi suất giảm dần). Bài 3 và 4 có mục tiêu và ràng buộc đều là tổng có trọng số của biến (tuyến tính); Simplex cho nghiệm tối ưu toàn cục và có shadow price.
3. **Shadow price 48,4 của ràng buộc kho nghĩa là gì?** Nới sức chứa thêm 1 m³ thì lợi nhuận gộp tối ưu tăng khoảng 48,4 USD, nên thuê thêm kho có lời nếu giá dưới mức này.
4. **Vì sao Data Table không tự chạy lại Solver?** Data Table chỉ tính lại công thức. Muốn so sánh nghiệm tối ưu ở các mức tham số khác nhau thì dùng Scenario (nhóm đã giải Solver cho từng kịch bản).
5. **Goal Seek khác Solver thế nào?** Goal Seek tìm **một** ô đầu vào để **một** ô kết quả đạt giá trị cho trước, không có ràng buộc. Solver tối ưu nhiều biến với nhiều ràng buộc.
6. **Hàm `1 − e^(−x/scale)` ở Bài 2 có ý nghĩa gì?** Mô tả hiệu quả marketing giảm dần: tăng nhanh lúc đầu rồi bão hòa. Khi chi bằng `scale` (25.000 USD) thì đạt khoảng 63% tiềm năng.
7. **Vì sao South có α cao nhất?** Dữ liệu thật năm 2018: South tăng 30,6% nhưng chỉ chiếm 17% doanh thu, nên còn nhiều dư địa.
8. **Giả định nào yếu nhất?** Biên lợi nhuận và chi phí, vì dataset không có giá vốn. Thứ hai là tỷ lệ giao nhanh của Second Class giữ ở 36,8% khi tăng số đơn.
9. **Vì sao Bài 4 tăng Second Class mà không tăng First Class?** Chi phí cho mỗi đơn giao nhanh thêm: Second (12−7)/0,368 ≈ 13,6 USD, rẻ hơn First (18−7)/0,582 ≈ 18,9 USD và Same Day 23 USD.
10. **Marketing 100.000 USD có lời không?** Theo mô hình thì doanh thu tăng thêm 104.816 USD, nhưng lợi nhuận gộp tăng chỉ khoảng 20.000 USD. Vì vậy cần biện minh bằng giá trị khách hàng dài hạn (tỷ lệ khách quay lại tăng từ 43% lên 71%).
