# Phần B: Model-Driven DSS với Excel

File nộp: `output/MaNhom_Model.xlsx`. Khi nộp, đổi `MaNhom` thành mã nhóm.
Toàn bộ số liệu kết quả và bảng Scenario Summary đầy đủ nằm ở `docs/PhanB_Ket_qua.md`, file này do script tự sinh từ chính file Excel.

## 0. Tổng quan

### 0.1 Bốn bài toán và công cụ Excel

| # | Bài toán | Lĩnh vực | Goal Seek | Data Table | Scenario | Solver |
|---|---|---|---|---|---|---|
| 1 | Hòa vốn và lợi nhuận mục tiêu 2019 | Quản trị kinh doanh | ✅ 3 câu hỏi | ✅ 2 chiều | ✅ 3 kịch bản | — |
| 2 | Phân bổ ngân sách marketing theo Region | Marketing | ✅ | ✅ 1 chiều | ✅ 3 kịch bản (mỗi kịch bản là một nghiệm Solver) | ✅ GRG Nonlinear |
| 3 | Kế hoạch nhập hàng theo Sub-Category | Logistics / tồn kho | — | ✅ 1 chiều | ✅ 3 kịch bản (nghiệm Solver) | ✅ Simplex LP + Sensitivity |
| 4 | Cơ cấu phương thức giao hàng | Vận hành | ✅ | ✅ 1 chiều | ✅ 3 kịch bản (nghiệm Solver) | ✅ Simplex LP + Sensitivity |

### 0.2 Cấu trúc file Excel
- `Huong_dan`: mục lục và quy ước màu ô. Vàng là tham số, xanh lá là biến quyết định, xanh dương là kết quả.
- `Du_lieu_2018`: số liệu **thật** năm 2018 trích từ star schema của Phần A.
- `Gia_dinh`: giả định dùng chung, có ghi lý do cho từng con số.
- `B1_…` đến `B4_…`: mỗi sheet là một bài toán, chia khối cố định: Tham số → Biến quyết định → Kết quả → Goal Seek → Data Table.
- `SC_B1`–`SC_B4`: Scenario Summary. `AR_B2`–`AR_B4`: Answer Report. `SR_B2`–`SR_B4`: Sensitivity Report của Solver.
- Mọi công thức dùng **named range** (`B1_Profit`, `B3_X`…), nên công thức tự giải thích được ý nghĩa. Mô hình Solver đã lưu sẵn trong từng sheet.

### 0.3 Nguồn số liệu và giả định
Dataset chỉ có doanh thu, không có giá vốn hay lợi nhuận. Vì vậy:
- **Lấy từ dữ liệu thật năm 2018**: doanh thu 722.052 USD; 1.661 đơn; AOV 434,71 USD; doanh thu theo Region, Sub-Category; số đơn, ngày giao trung bình và tỷ lệ giao trong ≤ 2 ngày theo Ship Mode.
- **Giả định của nhóm** (sheet `Gia_dinh` và khối Tham số của từng bài):

| Giả định | Giá trị | Lý do |
|---|---|---|
| Tăng trưởng số đơn 2019 | 20% | Thực tế tăng 27–28%/năm trong 2017–2018; chọn mức thận trọng |
| Biên lợi nhuận gộp theo Sub-Category | 8% (Tables) đến 40% (Labels); bình quân 20,7% | Giấy, nhãn, bìa hồ sơ có biên cao; bàn, kệ sách, máy móc cạnh tranh giá nên biên thấp |
| Phí thanh toán | 2% doanh thu | Phí thẻ phổ biến 1,5–3% |
| Chi phí xử lý đơn | 10 USD/đơn | Đóng gói, xử lý |
| Chi phí giao hàng | Same Day 30, First 18, Second 12, Standard 7 USD/đơn | Giao càng nhanh càng đắt |
| Chi phí cố định 2019 | 70.000 USD | Thuê kho, lương bộ máy, hệ thống |
| Thể tích lưu kho | 0,2–3,0 m³ cho mỗi 1.000 USD giá vốn | Nội thất cồng kềnh, điện thoại và nhãn nhỏ gọn |
| Hệ số phản ứng marketing α | South 40%, West 22%, East 18%, Central 12% | Theo tăng trưởng thật 2017→2018: South +30,6% nhưng thị phần chỉ 17%; Central −2,8% |

### 0.4 Chạy lại toàn bộ
```
python -I model/solve_reference.py      # lời giải đối chiếu bằng scipy → output/reference.json
python -I model/build_model.py          # dựng workbook → output/MaNhom_Model.xlsx
powershell -NoProfile -ExecutionPolicy Bypass -File model/apply_excel_tools.ps1   # Excel thật chạy Goal Seek, Data Table, Scenario, Solver
python -I model/verify_excel.py         # đối chiếu Excel với scipy → docs/PhanB_Ket_qua.md
```
Kết quả Excel khớp với scipy ở cả 20/20 kiểm tra (sai số ≤ 0,5%).

---

## Bài 1: Hòa vốn và lợi nhuận mục tiêu năm 2019

**Bối cảnh.** Phần A cho thấy số đơn tăng đều qua các năm nhưng AOV giảm từ 506 (2015) xuống 435 USD (2018). Ban giám đốc cần biết công ty phải bán tối thiểu bao nhiêu đơn để không lỗ, và cần bao nhiêu để đạt lợi nhuận mục tiêu 40.000 USD.

**Mô hình** (sheet `B1_Hoa_von`):

| Thành phần | Nội dung |
|---|---|
| Biến quyết định | Số đơn hàng 2019 (`B1_Orders`, mặc định 1.993 = 1.661 × 1,2), AOV (`B1_AOV`, mặc định 434,71) |
| Biến kết quả | Lợi nhuận 2019 (`B1_Profit`) |
| Công thức | Lợi nhuận góp/đơn = AOV × (biên gộp − phí thanh toán) − chi phí xử lý − chi phí giao TB<br>Lợi nhuận = Số đơn × Lợi nhuận góp/đơn − Chi phí cố định<br>Điểm hòa vốn = Chi phí cố định / Lợi nhuận góp/đơn |
| Ràng buộc | Không có (bài toán phân tích độ nhạy và tìm ngược) |

**Kết quả.**
- Lợi nhuận góp mỗi đơn là 60,43 USD. Với 1.993 đơn, **lợi nhuận 2019 ước tính 50.442 USD**, biên an toàn 41,9%.
- **Goal Seek** (Set `B1_Profit`, By changing `B1_Orders` hoặc `B1_AOV`):
  - Hòa vốn ở **1.158 đơn**, bằng 70% số đơn năm 2018.
  - Đạt lợi nhuận mục tiêu 40.000 USD ở **1.820 đơn**, tức chỉ cần số đơn tăng 9,6% so với 2018.
  - Nếu số đơn đạt dự báo 1.993, AOV có thể giảm đến **406,75 USD** mà vẫn đạt mục tiêu.
- **Data Table 2 chiều** (số đơn 1.200–2.400 × AOV 360–500 USD): vùng lỗ chỉ xuất hiện khi số đơn ≤ 1.400 và AOV ≤ 420. Ở mức 2.000 đơn, mỗi 20 USD AOV tương đương khoảng 7.500 USD lợi nhuận, tức khoảng 160 đơn hàng.
- **Scenario Manager**:

| Kịch bản | Số đơn | AOV | Chi phí cố định | Lợi nhuận | Điểm hòa vốn |
|---|---|---|---|---|---|
| Bi quan | 1.661 | 400 | 75.000 | 14.575 | 1.391 đơn |
| Cơ sở | 1.993 | 434,71 | 70.000 | 50.442 | 1.158 đơn |
| Lạc quan | 2.159 | 450 | 68.000 | 68.660 | 1.074 đơn |

**Ý nghĩa quản lý.**
- Mục tiêu 40.000 USD **khả thi**: chỉ cần 1.820 đơn, thấp hơn đà tăng trưởng thực tế.
- Rủi ro lớn nhất là **AOV tiếp tục giảm**. Trong kịch bản bi quan (AOV 400, số đơn đứng yên), lợi nhuận giảm 71%. Vì vậy giữ AOV (bán kèm, combo, ngưỡng miễn phí giao hàng) quan trọng ngang với tăng số đơn. Điều này khớp với insight AOV giảm ở Phần A.
- Có thể dùng mức 1.158 đơn làm "ngưỡng cảnh báo" trên dashboard Power BI.

**Hạn chế.** Biên gộp và chi phí là giả định. Mô hình coi AOV và số đơn độc lập với nhau, trong khi thực tế giảm giá có thể tăng số đơn nhưng làm giảm AOV.

---

## Bài 2: Phân bổ ngân sách marketing theo Region

**Bối cảnh.** Công ty có 100.000 USD ngân sách marketing cho năm 2019. Phần A cho thấy South chỉ chiếm 17% doanh thu nhưng tăng 30,6% năm 2018, còn Central giảm 2,8%. Câu hỏi: nên chia ngân sách cho 4 vùng thế nào để doanh thu tăng thêm lớn nhất?

**Mô hình** (sheet `B2_Marketing`):

| Thành phần | Nội dung |
|---|---|
| Biến quyết định | Ngân sách cho Central, East, South, West (`B2_Alloc`) |
| Biến kết quả | Tổng doanh thu tăng thêm (`B2_Uplift`), cần **tối đa hóa** |
| Công thức | DT tăng thêmᵣ = DT2018ᵣ × αᵣ × (1 − e^(−Ngân sáchᵣ / 25.000)). Đây là hàm **lợi suất giảm dần**: càng chi nhiều thì mỗi USD thêm càng mang lại ít doanh thu |
| Ràng buộc | Σ ngân sách ≤ 100.000; 10.000 ≤ ngân sách mỗi vùng ≤ 40.000 |
| Công cụ | Solver, phương pháp **GRG Nonlinear** (hàm mục tiêu phi tuyến) |

**Kết quả.**
- **Solver**:

| Region | Ngân sách tối ưu | Chia đều |
|---|---|---|
| Central | 10.000 (chạm mức tối thiểu) | 25.000 |
| East | 24.808 | 25.000 |
| South | 31.212 | 25.000 |
| West | 33.980 | 25.000 |
| **DT tăng thêm** | **104.816** | 100.048 |

  Với cùng 100.000 USD, phân bổ tối ưu mang thêm **4.768 USD doanh thu (+4,8%)** so với chia đều.
- **Goal Seek**: nếu chia đều, cần **118.745 USD** mới đạt mục tiêu tăng thêm 110.000 USD doanh thu.
- **Data Table** (chia đều, 40k–160k): ROAS giảm từ 1,30 xuống 0,79, thể hiện rõ lợi suất giảm dần.
- **Scenario** (mỗi kịch bản là nghiệm Solver):
  - 60.000 USD: doanh thu tăng thêm 75.176, ROAS 1,25.
  - 100.000 USD: doanh thu tăng thêm 104.816, ROAS 1,05.
  - 140.000 USD: doanh thu tăng thêm 122.113, ROAS 0,87; East, South, West đều chạm trần 40.000.

**Ý nghĩa quản lý.**
- Chuyển ngân sách **khỏi Central** (giữ mức tối thiểu để duy trì hiện diện) sang **South và West**, nơi còn dư địa tăng trưởng.
- Khi ngân sách vượt khoảng 100.000 USD, mỗi USD marketing đem về chưa tới 1 USD doanh thu (ROAS < 1). **Không nên tăng ngân sách lên mức cao**, trừ khi mục tiêu là giành thị phần dài hạn.
- Lưu ý quan trọng: với biên gộp khoảng 20%, 104.816 USD doanh thu tăng thêm chỉ tạo khoảng 20.000 USD lợi nhuận gộp, thấp hơn khoản chi 100.000 USD. Vì vậy marketing phải được biện minh bằng giá trị khách hàng dài hạn: Phần A cho thấy tỷ lệ khách quay lại tăng từ 43% lên 71%.

**Hạn chế.** Hàm phản ứng và hệ số α là giả định, cần kiểm chứng bằng thử nghiệm A/B hoặc số liệu chiến dịch thật. GRG có thể dừng ở nghiệm cục bộ, nhưng ở đây hàm lõm nên nghiệm cục bộ cũng là nghiệm toàn cục, và đã được đối chiếu với scipy.

---

## Bài 3: Kế hoạch nhập hàng theo Sub-Category

**Bối cảnh.** Nhu cầu 2019 dự báo bằng doanh thu 2018 × 1,2. Đáp ứng toàn bộ nhu cầu cần 687.000 USD tiền hàng và 954 m³ kho, trong khi công ty chỉ có **520.000 USD** và **650 m³**. Cần chọn nhập bao nhiêu cho mỗi loại để lợi nhuận gộp lớn nhất.

**Mô hình** (sheet `B3_Nhap_hang`):

| Thành phần | Nội dung |
|---|---|
| Biến quyết định | Giá trị nhập (theo giá vốn) của 17 Sub-Category (`B3_X`) |
| Biến kết quả | Lợi nhuận gộp kế hoạch = Σ (Nhậpᵢ / (1 − biênᵢ) − Nhậpᵢ) (`B3_Profit`), cần **tối đa hóa** |
| Ràng buộc | Σ Nhập ≤ ngân sách 520.000<br>Σ Nhậpᵢ × thể tíchᵢ ≤ 650 m³<br>Nhậpᵢ ≤ nhu cầu dự báo × (1 − biênᵢ) (không nhập quá nhu cầu)<br>Nhậpᵢ ≥ 50% mức trên (đáp ứng tối thiểu mỗi loại) |
| Công cụ | Solver, phương pháp **Simplex LP** (mô hình tuyến tính), kèm **Sensitivity Report** |

**Kết quả.**
- **Solver**: lợi nhuận gộp tối ưu **154.433 USD**. **Cả ngân sách và kho đều dùng hết 100%**, nghĩa là cả hai ràng buộc đều chặn. Tỷ lệ đáp ứng nhu cầu là 77,8%.
  - Nhập **đủ 100% nhu cầu** cho các loại biên cao hoặc nhỏ gọn: Binders, Copiers, Accessories, Appliances, Furnishings, Paper, Supplies, Art, Labels, Envelopes, Fasteners.
  - Chỉ nhập **mức tối thiểu 50%** cho các loại biên thấp hoặc cồng kềnh: Chairs, Tables, Machines, Bookcases.
  - Nhập một phần: Phones (57%), Storage (98%).
- **Sensitivity Report (shadow price)**:
  - Thêm 1 USD ngân sách nhập hàng làm tăng **0,153 USD** lợi nhuận gộp.
  - Thêm 1 m³ kho làm tăng **48,4 USD** lợi nhuận gộp.
- **Data Table** (nhu cầu thực tế 80–120% so với dự báo):
  - Nếu nhu cầu chỉ đạt 90%, lợi nhuận giảm còn 109.408 USD và tồn kho dư 32.642 USD.
  - Nếu nhu cầu đạt 80%, lợi nhuận còn 62.643 USD và tồn kho dư 66.677 USD.
  - Nhu cầu cao hơn dự báo không làm tăng lợi nhuận, vì lượng hàng đã bị giới hạn.
- **Scenario** (nghiệm Solver):

| Kịch bản | Ngân sách / Kho | LN gộp | Ràng buộc chặn |
|---|---|---|---|
| Thắt chặt | 460.000 / 600 m³ | 139.939 | Ngân sách (kho chỉ dùng 94%) |
| Cơ sở | 520.000 / 650 m³ | 154.433 | Cả hai |
| Mở rộng | 600.000 / 700 m³ | 165.769 | Kho (ngân sách còn dư 1%) |

**Ý nghĩa quản lý.**
- Ưu tiên vốn và kho cho hàng **biên cao, nhỏ gọn** (văn phòng phẩm, phụ kiện công nghệ). Hàng nội thất cồng kềnh, biên thấp chỉ nên giữ ở mức duy trì.
- **Kho là điểm nghẽn khi mở rộng**: tăng ngân sách mà không mở rộng kho thì gần như không thêm lợi nhuận. Theo shadow price, nếu thuê thêm kho với giá dưới **48 USD/m³/năm** thì có lời.
- Dự báo nhu cầu sai 10% (nhu cầu thực tế thấp hơn) làm mất khoảng 29% lợi nhuận. Nên nhập theo đợt và theo dõi tồn kho.

**Hạn chế.** Biên lợi nhuận và thể tích là giả định. Nhu cầu được giả định tăng đều 20% cho mọi loại hàng. Mô hình một kỳ, chưa tính chi phí lưu kho theo thời gian hay chiết khấu khi nhập số lượng lớn.

---

## Bài 4: Cơ cấu phương thức giao hàng

**Bối cảnh.** Phần A cho thấy chỉ 22% đơn được giao trong vòng 2 ngày. Ban giám đốc muốn nâng tỷ lệ này lên **30%** với chi phí giao hàng thấp nhất.

**Mô hình** (sheet `B4_Giao_hang`):

| Thành phần | Nội dung |
|---|---|
| Biến quyết định | Số đơn 2019 theo 4 Ship Mode (`B4_X`) |
| Biến kết quả | Tổng chi phí giao hàng (`B4_Cost`), cần **tối thiểu hóa** |
| Công thức | Chi phí = Σ số đơnₘ × chi phí/đơnₘ<br>Đơn giao nhanh = Σ số đơnₘ × tỷ lệ giao ≤ 2 ngày của mode m (tỷ lệ **thật** năm 2018: Same Day 100%, First 58,2%, Second 36,8%, Standard 0%) |
| Ràng buộc | Σ số đơn = 1.993<br>Đơn giao nhanh ≥ 30% × 1.993<br>Ngày giao TB ≤ 4<br>Same Day ≤ 150 đơn<br>Mỗi mode ≥ 5% số đơn |
| Công cụ | Solver, phương pháp **Simplex LP**, kèm Sensitivity Report |

**Kết quả.**
- **Solver**:
  - Same Day 100 đơn và First Class 100 đơn, cả hai ở mức tối thiểu 5%.
  - **Second Class 1.197 đơn** (năm 2018 là 318).
  - Standard Class 597 đơn.
  - Tổng chi phí **23.322 USD**, chỉ hơn cơ cấu 2018 là **1.339 USD (+6%)**. Tỷ lệ giao nhanh đạt 30%, ngày giao trung bình giảm từ 3,94 xuống 3,57.
- **Sensitivity Report**: mỗi đơn giao nhanh thêm tốn **13,59 USD**. Đây là chi phí biên của việc nâng mục tiêu giao nhanh.
- **Goal Seek** (phương án đơn giản: chỉ tăng First Class): cần nâng tỷ trọng First Class từ 16,9% lên **30,3%**. Theo Data Table, khi đó chi phí khoảng 24.900 USD, **cao hơn phương án Solver khoảng 1.600 USD**.
- **Scenario** (nghiệm Solver):

| Mục tiêu giao nhanh | Tổng chi phí | So với cơ cấu 2018 | Ngày giao TB |
|---|---|---|---|
| 25% | 21.968 | **−15 USD** (rẻ hơn mà vẫn nhanh hơn) | 3,81 |
| 30% | 23.322 | +1.339 USD | 3,57 |
| 35% | 24.677 | +2.694 USD | 3,33 |

**Ý nghĩa quản lý.**
- **Second Class là cách rẻ nhất để có thêm đơn giao nhanh.** Chi phí cho mỗi đơn giao nhanh thêm (so với Standard) là khoảng 13,6 USD với Second Class, 18,9 USD với First Class và 23 USD với Same Day.
- Chỉ cần **sắp xếp lại cơ cấu** (giảm First Class, tăng Second Class), công ty đã có thể nâng tỷ lệ giao nhanh từ 22% lên 25% mà không tốn thêm chi phí.
- Mỗi 5 điểm % giao nhanh thêm tốn khoảng 1.350 USD/năm. Ban giám đốc có thể so con số này với lợi ích giữ chân khách hàng để chọn mục tiêu.

**Hạn chế.**
- Tỷ lệ giao nhanh của từng mode lấy từ lịch sử. Nếu dồn nhiều đơn vào Second Class, tỷ lệ 36,8% có thể không giữ được.
- Thực tế khách hàng tự chọn Ship Mode. Muốn đạt cơ cấu tối ưu, công ty phải dẫn dắt bằng giá hoặc ưu đãi.
- Số đơn là số thực, khi triển khai cần làm tròn.

---

## Thao tác lại khi demo (Excel)

Cần bật add-in Solver trước: **File → Options → Add-ins → Manage: Excel Add-ins → Go → chọn Solver Add-in**.

| Công cụ | Cách làm | Ví dụ trong file |
|---|---|---|
| **Goal Seek** | Data → What-If Analysis → Goal Seek. Điền **Set cell**, **To value**, **By changing cell** → OK | B1: Set cell `B19` (Lợi nhuận), To value `0`, By changing `B6` (Số đơn) → ra 1.158. Nhớ hoàn tác (Ctrl+Z) để trả số đơn về 1.993 |
| **Data Table** | Bảng đã có sẵn và tự cập nhật khi đổi tham số. Để tạo lại: chọn cả vùng bảng → Data → What-If Analysis → Data Table → điền Row/Column input cell | B1: chọn `A31:I38`, Row input `B7` (AOV), Column input `B6` (Số đơn). B2/B3/B4: chỉ điền Column input (`B22`, `B10`, `B32`) |
| **Scenario Manager** | Data → What-If Analysis → Scenario Manager → chọn kịch bản → **Show** để xem; **Summary** để tạo báo cáo | Thử: đổi tham số trong B1 bằng cách chọn "Bi quan" → Show → lợi nhuận còn 14.575. Chọn "Cơ sở" → Show để trả lại |
| **Solver** | Data → Solver: mô hình đã nạp sẵn (ô mục tiêu, biến, ràng buộc) → **Solve** → Keep Solver Solution, chọn thêm Answer và Sensitivity nếu cần | B3: Set Objective `K31` Max, By Changing `I14:I30`, các ràng buộc ngân sách, kho, min/max. Demo: đổi ngân sách `B7` thành 600.000 → Solve → xem lợi nhuận và tỷ lệ dùng kho |

**Gợi ý kịch bản demo (3–4 phút):**
1. Bài 1: Goal Seek tìm số đơn hòa vốn, rồi Show kịch bản "Bi quan" để thấy tác động của việc AOV giảm.
2. Bài 3: tăng ngân sách và bấm Solve, thấy lợi nhuận gần như không tăng vì kho đã đầy. Mở `SR_B3` chỉ shadow price của kho (48,4 USD/m³).
3. Bài 4: đổi mục tiêu giao nhanh 30% → 35% và bấm Solve. Nêu chi phí biên 13,59 USD/đơn.
