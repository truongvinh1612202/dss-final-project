# Giá trị tham chiếu KPI

> File này do `etl/kpi_reference.py` tự sinh từ `data/star/`. Không sửa tay, hãy chạy lại script.
> Dùng để đối chiếu với các card và visual trên Power BI (định nghĩa xem `docs/KPI_va_DAX.md`).

## Toàn bộ giai đoạn 2015–2018

| Doanh thu | Số đơn | AOV | Số khách hàng | Tỷ lệ KH quay lại | Ngày giao TB | Tỷ lệ giao nhanh |
|---|---|---|---|---|---|---|
| 2,261,255.41 | 4,922 | 459.42 | 793 | 98.4% | 3.96 | 22.1% |

## Theo năm (Order Date)

| Year | Doanh thu | Số đơn | AOV | Số khách hàng | Tỷ lệ KH quay lại | Ngày giao TB | Tỷ lệ giao nhanh | YoY doanh thu |
|---|---|---|---|---|---|---|---|---|
| 2015 | 479,574.84 | 947 | 506.41 | 589 | 43.0% | 4.00 | 21.6% | — |
| 2016 | 459,436.01 | 1,019 | 450.87 | 567 | 54.0% | 3.99 | 20.6% | -4.2% |
| 2017 | 600,192.55 | 1,295 | 463.47 | 635 | 58.4% | 3.92 | 23.6% | 30.6% |
| 2018 | 722,052.02 | 1,661 | 434.71 | 690 | 70.7% | 3.94 | 22.2% | 20.3% |

## Theo Region

| Region | Doanh thu | Tỷ trọng DT | Số đơn | AOV | Số khách hàng |
|---|---|---|---|---|---|
| West | 710,219.68 | 31.4% | 1,587 | 447.52 | 681 |
| East | 669,237.35 | 29.6% | 1,369 | 488.85 | 669 |
| Central | 492,646.91 | 21.8% | 1,156 | 426.17 | 626 |
| South | 389,151.46 | 17.2% | 810 | 480.43 | 509 |

## Theo Category

| Category | Doanh thu | Tỷ trọng DT | Số đơn | AOV | Số khách hàng |
|---|---|---|---|---|---|
| Technology | 827,455.87 | 36.6% | 1,519 | 544.74 | 684 |
| Furniture | 728,377.20 | 32.2% | 1,727 | 421.76 | 705 |
| Office Supplies | 705,422.33 | 31.2% | 3,676 | 191.90 | 787 |

## Theo Segment

| Segment | Doanh thu | Tỷ trọng DT | Số đơn | AOV | Số khách hàng |
|---|---|---|---|---|---|
| Consumer | 1,148,060.53 | 50.8% | 2,537 | 452.53 | 409 |
| Corporate | 688,494.07 | 30.4% | 1,491 | 461.77 | 236 |
| Home Office | 424,700.80 | 18.8% | 894 | 475.06 | 148 |

## Theo Ship Mode

| Ship Mode | Số đơn | Tỷ trọng DT | Ngày giao TB | Tỷ lệ giao nhanh |
|---|---|---|---|---|
| Same Day | 261 | 5.5% | 0.05 | 100.0% |
| First Class | 772 | 15.3% | 2.19 | 60.1% |
| Second Class | 944 | 19.9% | 3.24 | 38.6% |
| Standard Class | 2,945 | 59.3% | 5.00 | 0.0% |

## Top 5 Sub-Category theo doanh thu

| Sub-Category | Doanh thu | Tỷ trọng DT | Số đơn | AOV |
|---|---|---|---|---|
| Phones | 327,782.45 | 14.5% | 803 | 408.20 |
| Chairs | 322,541.36 | 14.3% | 566 | 569.86 |
| Storage | 219,343.39 | 9.7% | 764 | 287.10 |
| Tables | 202,810.63 | 9.0% | 302 | 671.56 |
| Binders | 200,028.79 | 8.8% | 1,291 | 154.94 |
