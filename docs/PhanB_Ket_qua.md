# Phần B – Kết quả mô hình (tự sinh)

> File này do `model/verify_excel.py` sinh từ `output/PhanB_Dap_an_tham_khao.xlsx` sau khi Excel chạy Goal Seek, Data Table,
> Scenario Manager và Solver. Không sửa tay. Diễn giải và ý nghĩa quản lý xem `docs/PhanB_Mo_hinh.md`.

## Đối chiếu Excel ↔ scipy (20/20 khớp, sai số ≤ 0.5%)

| Kiểm tra | Excel | scipy | Kết quả |
|---|---|---|---|
| B1 lợi nhuận góp/đơn | 60.4327 | 60.4326 | Khớp |
| B1 lợi nhuận 2019 (cơ sở) | 50,442.3657 | 50,442.0746 | Khớp |
| B1 Goal Seek – số đơn hòa vốn | 1,158.3133 | 1,158.3161 | Khớp |
| B1 Goal Seek – số đơn đạt mục tiêu | 1,820.2067 | 1,820.2111 | Khớp |
| B1 Goal Seek – AOV đạt mục tiêu | 406.7503 | 406.7503 | Khớp |
| B2 Solver – DT tăng thêm tối ưu | 104,815.6097 | 104,815.5244 | Khớp |
| B2 Solver – ngân sách Central | 10,000.0000 | 10,000.0000 | Khớp |
| B2 Solver – ngân sách East | 24,807.8452 | 24,842.0000 | Khớp |
| B2 Solver – ngân sách South | 31,211.6972 | 31,141.0000 | Khớp |
| B2 Solver – ngân sách West | 33,980.4576 | 34,017.0000 | Khớp |
| B2 Goal Seek – ngân sách chia đều đạt mục tiêu | 118,744.6579 | 118,744.6618 | Khớp |
| B3 Solver – LN gộp tối ưu | 154,433.3475 | 154,433.3458 | Khớp |
| B3 Solver – ngân sách đã dùng | 520,000.0000 | 520,000.0000 | Khớp |
| B3 Solver – thể tích đã dùng | 650.0000 | 650.0000 | Khớp |
| B4 Solver – tổng chi phí tối thiểu | 23,322.2984 | 23,321.8414 | Khớp |
| B4 Solver – số đơn Same Day | 99.6500 | 99.6000 | Khớp |
| B4 Solver – số đơn First Class | 99.6500 | 99.6000 | Khớp |
| B4 Solver – số đơn Second Class | 1,196.6397 | 1,196.5000 | Khớp |
| B4 Solver – số đơn Standard Class | 597.0603 | 597.2000 | Khớp |
| B4 Goal Seek – tỷ trọng First Class | 0.3034 | 0.3033 | Khớp |

## Bài 1 – Hòa vốn và lợi nhuận mục tiêu 2019

| Chỉ số | Giá trị |
|---|---|
| Biên LN gộp bình quân | 20.7% |
| Chi phí giao TB / đơn | 11.03 USD |
| Lợi nhuận góp / đơn | 60.43 USD |
| Số đơn dự báo 2019 | 1,993 |
| Lợi nhuận 2019 (cơ sở) | 50,442 USD |
| Biên an toàn | 41.9% |
| **Goal Seek** – số đơn hòa vốn | **1,158 đơn** |
| **Goal Seek** – số đơn đạt LN mục tiêu 40,000 USD | **1,820 đơn** |
| **Goal Seek** – AOV tối thiểu đạt mục tiêu (với 1,993 đơn) | **406.75 USD** |

**Data Table 2 chiều – Lợi nhuận (USD) theo số đơn × AOV**

| Số đơn \ AOV | 360 | 380 | 400 | 420 | 440 | 460 | 480 | 500 |
|---|---|---|---|---|---|---|---|---|
| 1,200 | -14,281 | -9,784 | -5,286 | -789 | 3,709 | 8,206 | 12,704 | 17,201 |
| 1,400 | -4,995 | 252 | 5,499 | 10,747 | 15,994 | 21,241 | 26,488 | 31,735 |
| 1,600 | 4,292 | 10,288 | 16,285 | 22,282 | 28,278 | 34,275 | 40,272 | 46,268 |
| 1,800 | 13,578 | 20,324 | 27,071 | 33,817 | 40,563 | 47,309 | 54,056 | 60,802 |
| 2,000 | 22,865 | 30,361 | 37,856 | 45,352 | 52,848 | 60,344 | 67,840 | 75,336 |
| 2,200 | 32,151 | 40,397 | 48,642 | 56,887 | 65,133 | 73,378 | 81,624 | 89,869 |
| 2,400 | 41,438 | 50,433 | 59,428 | 68,423 | 77,418 | 86,413 | 95,408 | 104,403 |

**Scenario Summary**

| Scenario Summary |  |  |  |  |  |
|---|---|---|---|---|---|
|  |  | Current Values: | Bi quan | Cơ sở | Lạc quan |
|  |  |  | Số đơn không tăng, AOV giảm, chi phí cố định tăng | Số đơn +20%, AOV như 2018 | Số đơn +30%, AOV tăng, tối ưu chi phí cố định |
| Changing Cells: |  |  |  |  |  |
|  | B1_Orders | 1993 | 1661 | 1993 | 2159 |
|  | B1_AOV | 434.71 | 400 | 434.71 | 450 |
|  | B1_Fixed | 70000 | 75000 | 70000 | 68000 |
| Result Cells: |  |  |  |  |  |
|  | B1_Revenue | 866,377.03 | 664400 | 866,377.03 | 971550 |
|  | B1_Contrib | 60.43 | 53.93 | 60.43 | 63.30 |
|  | B1_Profit | 50,442.37 | 14,574.73 | 50,442.37 | 68,660.33 |
|  | B1_BreakEven | 1,158.31 | 1,390.74 | 1,158.31 | 1,074.28 |
|  | B1_Safety | 0.42 | 0.16 | 0.42 | 0.50 |
| Notes:  Current Values column represents values of changing cells at |  |  |  |  |  |
| time Scenario Summary Report was created.  Changing cells for each |  |  |  |  |  |
| scenario are highlighted in gray. |  |  |  |  |  |

## Bài 2 – Phân bổ ngân sách marketing

**Solver (GRG Nonlinear) – phương án cơ sở 100.000 USD**

| Region | Ngân sách tối ưu | Ngân sách chia đều |
|---|---|---|
| Central | 10,000 | 25,000 |
| East | 24,808 | 25,000 |
| South | 31,212 | 25,000 |
| West | 33,980 | 25,000 |

| Chỉ số | Giá trị |
|---|---|
| DT tăng thêm – phân bổ tối ưu | 104,816 USD |
| DT tăng thêm – chia đều | 100,048 USD |
| ROAS tối ưu | 1.05 |
| **Goal Seek** – ngân sách chia đều cần để tăng 110,000 USD DT | **118,745 USD** |

**Data Table – hiệu quả theo tổng ngân sách (chia đều)**

| Tổng ngân sách | DT tăng thêm | ROAS |
|---|---|---|
| 40,000 | 52,179 | 1.30 |
| 60,000 | 71,411 | 1.19 |
| 80,000 | 87,156 | 1.09 |
| 100,000 | 100,048 | 1.00 |
| 120,000 | 110,602 | 0.92 |
| 140,000 | 119,243 | 0.85 |
| 160,000 | 126,318 | 0.79 |

**Scenario Summary (mỗi kịch bản là nghiệm Solver tối ưu)**

| Scenario Summary |  |  |  |  |  |
|---|---|---|---|---|---|
|  |  | Current Values: | Ngân sách thấp | Ngân sách cơ sở | Ngân sách cao |
|  |  |  | Nghiệm tối ưu của Solver | Nghiệm tối ưu của Solver | Nghiệm tối ưu của Solver |
| Changing Cells: |  |  |  |  |  |
|  | B2_Budget | 100000 | 60000 | 100000 | 140000 |
|  | B2_NS_Central | 10000 | 10000 | 10000 | 20,000.15 |
|  | B2_NS_East | 24,807.85 | 11,474.51 | 24,807.85 | 39,999.85 |
|  | B2_NS_South | 31,211.70 | 17,878.36 | 31,211.70 | 40000 |
|  | B2_NS_West | 33,980.46 | 20,647.12 | 33,980.46 | 40000 |
| Result Cells: |  |  |  |  |  |
|  | B2_AllocSum | 100,000.00 | 60000 | 100,000.00 | 140000 |
|  | B2_Uplift | 104,815.61 | 75,176.21 | 104,815.61 | 122,113.18 |
|  | B2_ROAS | 1.05 | 1.25 | 1.05 | 0.87 |
| Notes:  Current Values column represents values of changing cells at |  |  |  |  |  |
| time Scenario Summary Report was created.  Changing cells for each |  |  |  |  |  |
| scenario are highlighted in gray. |  |  |  |  |  |

## Bài 3 – Kế hoạch nhập hàng

| Chỉ số | Giá trị |
|---|---|
| LN gộp kế hoạch tối ưu | 154,433 USD |
| Ngân sách đã dùng | 520,000 / 520,000 USD (100.0%) |
| Kho đã dùng | 650.0 / 650 m³ (100.0%) |
| Tỷ lệ đáp ứng nhu cầu | 77.8% |
| Giá trị biên (shadow price) của 1 USD ngân sách | 0.153 USD LN |
| Giá trị biên của 1 m³ kho | 48.4 USD LN |

| Sub-Category | Biên LN | Nhập (giá vốn) | % nhu cầu | Trạng thái |
|---|---|---|---|---|
| Phones | 14.0% | 61,857 | 57.3% | Một phần |
| Chairs | 18.0% | 45,850 | 50.0% | Tối thiểu |
| Binders | 35.0% | 55,943 | 100.0% | Tối đa |
| Storage | 20.0% | 63,957 | 97.9% | Một phần |
| Copiers | 30.0% | 52,835 | 100.0% | Tối đa |
| Tables | 8.0% | 33,613 | 50.0% | Tối thiểu |
| Accessories | 22.0% | 54,514 | 100.0% | Tối đa |
| Machines | 10.0% | 23,514 | 50.0% | Tối thiểu |
| Appliances | 22.0% | 37,639 | 100.0% | Tối đa |
| Bookcases | 10.0% | 16,213 | 50.0% | Tối thiểu |
| Furnishings | 28.0% | 24,369 | 100.0% | Tối đa |
| Paper | 38.0% | 20,464 | 100.0% | Tối đa |
| Supplies | 18.0% | 15,697 | 100.0% | Tối đa |
| Art | 30.0% | 7,414 | 100.0% | Tối đa |
| Labels | 40.0% | 2,768 | 100.0% | Tối đa |
| Envelopes | 35.0% | 2,635 | 100.0% | Tối đa |
| Fasteners | 30.0% | 718 | 100.0% | Tối đa |

**Data Table – rủi ro khi nhu cầu thực tế lệch dự báo**

| Nhu cầu thực tế | LN thực hiện | Tồn kho dư |
|---|---|---|
| 80.0% | 62,643 | 66,677 |
| 85.0% | 86,025 | 49,659 |
| 90.0% | 109,408 | 32,642 |
| 95.0% | 132,791 | 15,625 |
| 100.0% | 154,433 | 0 |
| 105.0% | 154,433 | 0 |
| 110.0% | 154,433 | 0 |
| 120.0% | 154,433 | 0 |

**Scenario Summary**

| Scenario Summary |  |  |  |  |  |
|---|---|---|---|---|---|
|  |  | Current Values: | Thắt chặt | Cơ sở | Mở rộng |
|  |  |  | Nghiệm tối ưu của Solver | Nghiệm tối ưu của Solver | Nghiệm tối ưu của Solver |
| Changing Cells: |  |  |  |  |  |
|  | B3_Budget | 520000 | 460000 | 520000 | 600000 |
|  | B3_Capacity | 650 | 600 | 650 | 700 |
|  | B3_Nhap_Phones | 61,856.69 | 54,017.80 | 61,856.69 | 108,035.60 |
|  | B3_Nhap_Chairs | 45,850.10 | 45,850.10 | 45,850.10 | 49,755.15 |
|  | B3_Nhap_Binders | 55,942.58 | 55,942.58 | 55,942.58 | 55,942.58 |
|  | B3_Nhap_Storage | 63,957.47 | 32,674.90 | 63,957.47 | 65,349.79 |
|  | B3_Nhap_Copiers | 52,835.49 | 52,835.49 | 52,835.49 | 52,835.49 |
|  | B3_Nhap_Tables | 33,613.23 | 33,613.23 | 33,613.23 | 33,613.23 |
|  | B3_Nhap_Accessories | 54,513.55 | 54,513.55 | 54,513.55 | 54,513.55 |
|  | B3_Nhap_Machines | 23,514.13 | 23,514.13 | 23,514.13 | 47,028.25 |
|  | B3_Nhap_Appliances | 37,639.11 | 24,609.00 | 37,639.11 | 37,639.11 |
|  | B3_Nhap_Bookcases | 16,213.11 | 16,213.11 | 16,213.11 | 16,213.11 |
|  | B3_Nhap_Furnishings | 24,368.85 | 24,368.85 | 24,368.85 | 24,368.85 |
|  | B3_Nhap_Paper | 20,463.56 | 20,463.56 | 20,463.56 | 20,463.56 |
|  | B3_Nhap_Supplies | 15,696.88 | 7,848.44 | 15,696.88 | 15,696.88 |
|  | B3_Nhap_Art | 7,413.53 | 7,413.53 | 7,413.53 | 7,413.53 |
|  | B3_Nhap_Labels | 2,768.18 | 2,768.18 | 2,768.18 | 2,768.18 |
|  | B3_Nhap_Envelopes | 2,635.28 | 2,635.28 | 2,635.28 | 2,635.28 |
|  | B3_Nhap_Fasteners | 718.26 | 718.26 | 718.26 | 718.26 |
| Result Cells: |  |  |  |  |  |
|  | B3_XSum | 520000 | 460000 | 520000 | 594,990.41 |
|  | B3_Space | 650 | 562.40 | 650 | 700 |
|  | B3_Profit | 154,433.35 | 139,938.62 | 154,433.35 | 165,768.81 |
|  | B3_FillRate | 0.78 | 0.69 | 0.78 | 0.88 |
|  | B3_BudgetUse | 1 | 1 | 1 | 0.99 |
|  | B3_SpaceUse | 1 | 0.94 | 1 | 1 |
| Notes:  Current Values column represents values of changing cells at |  |  |  |  |  |
| time Scenario Summary Report was created.  Changing cells for each |  |  |  |  |  |
| scenario are highlighted in gray. |  |  |  |  |  |

## Bài 4 – Cơ cấu phương thức giao hàng

| Ship Mode | Số đơn 2018 | Số đơn 2019 tối ưu | Tỷ lệ giao nhanh của mode | Chi phí/đơn |
|---|---|---|---|---|
| Same Day | 88 | 100 | 100.0% | 30.0 |
| First Class | 280 | 100 | 58.2% | 18.0 |
| Second Class | 318 | 1,197 | 36.8% | 12.0 |
| Standard Class | 975 | 597 | 0.0% | 7.0 |

| Chỉ số | Giá trị |
|---|---|
| Tổng chi phí tối thiểu (mục tiêu 30%) | 23,322 USD |
| Chi phí nếu giữ cơ cấu 2018 (~22% giao nhanh) | 21,983 USD |
| Chi phí tăng thêm | 1,339 USD |
| Tỷ lệ giao nhanh / ngày giao TB đạt được | 30.0% / 3.57 ngày |
| Chi phí biên cho mỗi đơn giao nhanh thêm (shadow price) | 13.59 USD |
| **Goal Seek** – tỷ trọng First Class cần (nếu chỉ tăng First Class) | **30.3%** |

**Data Table – chỉ tăng First Class**

| Tỷ trọng First Class | Tỷ lệ giao nhanh | Tổng chi phí |
|---|---|---|
| 10.0% | 18.2% | 20,480 |
| 15.0% | 21.1% | 21,576 |
| 20.0% | 24.0% | 22,672 |
| 25.0% | 26.9% | 23,768 |
| 30.0% | 29.8% | 24,864 |
| 35.0% | 32.7% | 25,960 |
| 40.0% | 35.6% | 27,057 |
| 45.0% | 38.5% | 28,153 |
| 50.0% | 41.4% | 29,249 |

**Scenario Summary**

| Scenario Summary |  |  |  |  |  |
|---|---|---|---|---|---|
|  |  | Current Values: | Mục tiêu 25% | Mục tiêu 30% | Mục tiêu 35% |
|  |  |  | Nghiệm tối ưu của Solver | Nghiệm tối ưu của Solver | Nghiệm tối ưu của Solver |
| Changing Cells: |  |  |  |  |  |
|  | B4_FastTarget | 0.30 | 0.25 | 0.30 | 0.35 |
|  | B4_Don_SameDay | 99.65 | 99.65 | 99.65 | 99.65 |
|  | B4_Don_FirstClass | 99.65 | 99.65 | 99.65 | 99.65 |
|  | B4_Don_SecondClass | 1,196.64 | 925.78 | 1,196.64 | 1,467.50 |
|  | B4_Don_StandardClass | 597.06 | 867.92 | 597.06 | 326.20 |
| Result Cells: |  |  |  |  |  |
|  | B4_Cost | 23,322.30 | 21,967.99 | 23,322.30 | 24,676.61 |
|  | B4_ExtraCost | 1,339.30 | -15.00 | 1,339.30 | 2,693.61 |
|  | B4_FastRate | 0.30 | 0.25 | 0.30 | 0.35 |
|  | B4_AvgDays | 3.57 | 3.81 | 3.57 | 3.33 |
|  | B4_CostPerOrder | 11.70 | 11.02 | 11.70 | 12.38 |
| Notes:  Current Values column represents values of changing cells at |  |  |  |  |  |
| time Scenario Summary Report was created.  Changing cells for each |  |  |  |  |  |
| scenario are highlighted in gray. |  |  |  |  |  |
