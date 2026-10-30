# Phần A: Mục tiêu dashboard, KPI và DAX

Giá trị tham chiếu để đối chiếu với Power BI nằm ở `docs/KPI_doi_chieu.md`, do `etl/kpi_reference.py` sinh ra. Mô hình dữ liệu xem `docs/Star_schema.md`.

## 1. Mục tiêu dashboard

Dashboard giúp ban quản lý một chuỗi bán lẻ văn phòng phẩm, nội thất và thiết bị công nghệ tại Mỹ **theo dõi hiệu quả kinh doanh giai đoạn 2015–2018**. Dashboard trả lời bốn câu hỏi:
- Doanh thu tăng hay giảm, đến từ đâu (thời gian, khu vực, sản phẩm, phân khúc khách hàng)?
- Doanh thu tăng nhờ có nhiều đơn hơn hay nhờ mỗi đơn có giá trị lớn hơn?
- Khách hàng có quay lại mua tiếp không?
- Năng lực giao hàng đang ở mức nào và phân bổ theo phương thức giao ra sao?

## 2. Đối tượng sử dụng

| Đối tượng | Nhu cầu | Trang dashboard chính |
|---|---|---|
| Giám đốc kinh doanh | Nắm tổng quan và xu hướng doanh thu để đặt mục tiêu năm | Tổng quan |
| Quản lý ngành hàng | Biết ngành hàng/sản phẩm nào đóng góp nhiều nhất để ưu tiên nhập hàng và khuyến mãi | Sản phẩm |
| Quản lý vùng và marketing | So sánh khu vực, phân khúc, mức độ trung thành của khách để phân bổ ngân sách | Khách hàng & Khu vực |
| Quản lý vận hành/logistics | Theo dõi thời gian giao và cơ cấu phương thức giao | Giao hàng |

## 3. Bộ KPI

| # | KPI | Cách tính | Ý nghĩa quản lý | Visual gợi ý |
|---|---|---|---|---|
| 1 | **Doanh thu** | Tổng Sales | Quy mô kinh doanh; cơ sở để đặt và theo dõi mục tiêu | Card, line theo tháng |
| 2 | **Số đơn hàng** | Đếm Order ID phân biệt | Lượng giao dịch; tách được tăng trưởng do số đơn hay do giá trị đơn | Card, cột theo năm |
| 3 | **Giá trị đơn hàng trung bình (AOV)** | Doanh thu / Số đơn hàng | AOV giảm báo hiệu cần bán kèm (cross-sell) hoặc nâng hạng sản phẩm (upsell) | Card, line theo năm |
| 4 | **Số khách hàng hoạt động** | Đếm Customer ID phân biệt có mua trong kỳ | Độ phủ khách hàng; so với số đơn để biết tần suất mua | Card, cột theo Segment |
| 5 | **Tăng trưởng doanh thu so với cùng kỳ (YoY %)** | (Doanh thu kỳ này − cùng kỳ năm trước) / cùng kỳ năm trước | Đánh giá đà tăng trưởng, phát hiện năm hoặc khu vực suy giảm | Card có màu tăng/giảm, cột theo năm |
| 6 | **Tỷ lệ khách hàng quay lại** | Số khách có ≥ 2 đơn trong kỳ / Số khách hàng hoạt động | Mức độ trung thành; căn cứ cho chương trình khách hàng thân thiết | Line theo năm |
| 7 | **Thời gian giao hàng trung bình** | Trung bình Ship Days, tính theo đơn hàng | Năng lực vận hành; so sánh giữa các phương thức giao | Card, cột theo Ship Mode |
| 8 | **Tỷ lệ đơn giao nhanh (≤ 2 ngày)** | Số đơn có Ship Days ≤ 2 / Số đơn hàng | Mức độ đáp ứng nhu cầu giao nhanh của khách | Gauge, cột theo Region |

Ghi chú:
- Dataset không có Profit, Quantity và Discount, nên không có KPI lợi nhuận. Điều này cần ghi vào phần hạn chế của báo cáo.
- **Tỷ lệ khách hàng quay lại phụ thuộc độ dài kỳ**: tính trên cả 4 năm thì ra 98,4%, nên hãy xem theo từng năm (43% → 71%).
- Ship Days giống nhau trên mọi dòng của cùng một đơn. KPI 7 và 8 vì vậy tính trên đơn hàng, để đơn có nhiều dòng không bị đếm nặng hơn.

## 4. Measure DAX

Tạo một bảng `_Measures` (Enter data → bảng rỗng) để chứa các measure. Tên bảng và cột lấy đúng theo `data/star/*.csv`.

```dax
Doanh thu = SUM ( Fact_Sales[Sales] )

Số đơn hàng = DISTINCTCOUNT ( Fact_Sales[Order ID] )

AOV = DIVIDE ( [Doanh thu], [Số đơn hàng] )

Số khách hàng = DISTINCTCOUNT ( Fact_Sales[Customer ID] )

Doanh thu cùng kỳ năm trước =
CALCULATE ( [Doanh thu], SAMEPERIODLASTYEAR ( Dim_Date[Date] ) )

YoY Doanh thu % =
VAR NamTruoc = [Doanh thu cùng kỳ năm trước]
RETURN IF ( NOT ISBLANK ( NamTruoc ), DIVIDE ( [Doanh thu] - NamTruoc, NamTruoc ) )

Tỷ lệ KH quay lại =
VAR KhachHang = VALUES ( Fact_Sales[Customer ID] )
VAR KhachQuayLai =
    COUNTROWS (
        FILTER ( KhachHang, CALCULATE ( DISTINCTCOUNT ( Fact_Sales[Order ID] ) ) >= 2 )
    )
RETURN DIVIDE ( KhachQuayLai, COUNTROWS ( KhachHang ) )

Ngày giao TB =
AVERAGEX ( VALUES ( Fact_Sales[Order ID] ), CALCULATE ( MAX ( Fact_Sales[Ship Days] ) ) )

Tỷ lệ giao nhanh =
VAR DonGiaoNhanh =
    COUNTROWS (
        FILTER ( VALUES ( Fact_Sales[Order ID] ), CALCULATE ( MAX ( Fact_Sales[Ship Days] ) ) <= 2 )
    )
RETURN DIVIDE ( DonGiaoNhanh, [Số đơn hàng] )

-- Measure phụ trợ
Tỷ trọng doanh thu = DIVIDE ( [Doanh thu], CALCULATE ( [Doanh thu], ALLSELECTED () ) )

Doanh thu theo ngày giao =
CALCULATE ( [Doanh thu], USERELATIONSHIP ( Fact_Sales[ShipDateKey], Dim_Date[DateKey] ) )
```

Thiết lập cần có để các measure chạy đúng:
- `Dim_Date` phải được **Mark as date table** theo cột `Date`, nếu không `SAMEPERIODLASTYEAR` sẽ sai.
- Định dạng: Doanh thu và AOV dạng tiền tệ, 2 chữ số thập phân; các tỷ lệ dạng %, 1 chữ số thập phân.
- `Dim_Date` kéo dài đến 2019 (để chứa Ship Date). Khi vẽ theo Order Date, lọc `Year ≤ 2018` ở mức trang để không hiện năm 2019 trống.

## 5. Bố cục dashboard và thao tác OLAP

| Trang | Visual chính | Thao tác OLAP minh họa |
|---|---|---|
| **Tổng quan** | 8 card KPI; line doanh thu theo Year → Quarter → Month; cột Số đơn và AOV theo năm; slicer Year, Region, Segment, Category | **Slice**: chọn Year = 2018. **Drill-down**: Year → Quarter → Month trên line |
| **Sản phẩm** | Bar Category → Sub-Category → Product; treemap tỷ trọng doanh thu; matrix Category × Year | **Drill-down** theo phân cấp sản phẩm. **Pivot**: đổi hàng/cột của matrix (Category × Year ↔ Year × Category) |
| **Khách hàng & Khu vực** | Bản đồ doanh thu theo State; bar Region → State → City; donut theo Segment; line Tỷ lệ KH quay lại theo năm | **Dice**: Region = West và Segment = Corporate. **Drill-down** Region → State → City. **Roll-up**: quay lại mức Region |
| **Giao hàng** | Cột Ngày giao TB và Tỷ lệ giao nhanh theo Ship Mode; matrix Ship Mode × Region; line Doanh thu theo ngày giao | Slice theo Ship Mode; so sánh Order Date với Ship Date qua `USERELATIONSHIP` |

## 6. Gợi ý insight từ số liệu (cần xác nhận lại trên dashboard)

Số liệu lấy từ `docs/KPI_doi_chieu.md`:
1. **Tăng trưởng mạnh sau năm 2016**: doanh thu giảm 4,2% năm 2016, rồi tăng 30,6% năm 2017 và 20,3% năm 2018.
2. **Tăng trưởng đến từ số đơn, không phải giá trị đơn**: số đơn tăng từ 947 (2015) lên 1.661 (2018), nhưng AOV giảm từ 506 xuống 435. Gợi ý cần chương trình bán kèm và combo để nâng giá trị đơn.
3. **Doanh thu tập trung ở hai vùng ven biển**: West và East chiếm 61% doanh thu, South chỉ 17%. Cần xem South là thị trường cần đầu tư thêm hay nên giữ chi phí thấp.
4. **Technology dẫn đầu nhờ giá trị đơn cao**: chiếm 36,6% doanh thu với AOV 545. Office Supplies có số đơn gấp hơn 2 lần nhưng AOV chỉ 192. Phones và Chairs là hai Sub-Category lớn nhất, mỗi loại khoảng 14%.
5. **Khách hàng ngày càng trung thành**: tỷ lệ khách quay lại trong năm tăng từ 43% (2015) lên 71% (2018).
6. **Giao hàng chủ yếu là Standard Class**: Standard Class chiếm 59% doanh thu, trung bình 5 ngày. Chỉ 22% đơn được giao trong vòng 2 ngày, và tỷ lệ này hầu như không đổi qua các năm.
