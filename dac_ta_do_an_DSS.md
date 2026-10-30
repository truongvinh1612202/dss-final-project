# CAO HỌC K34 – HỆ HỖ TRỢ RA QUYẾT ĐỊNH

# ĐẶC TẢ VÀ YÊU CẦU ĐỒ ÁN

## 1. MỤC TIÊU

- Hiểu và vận dụng kiến thức về Data-Driven DSS và Model-Driven DSS.
- Thực hành xây dựng hệ DSS sử dụng Power BI (phân tích mô tả) và Excel (mô hình dự đoán, mô phỏng, tối ưu).
- Phát triển kỹ năng tích hợp dữ liệu và mô hình để hỗ trợ ra quyết định trong bối cảnh thực tế.

## 2. YÊU CẦU ĐỒ ÁN

### Yêu cầu chung

- Học viên làm việc theo nhóm 2–3 người.
- Nhóm tự chọn một chủ đề (giáo dục, thương mại điện tử, y tế, …).
- Dataset:
  - Có thể sử dụng dataset công khai có sẵn (Kaggle, UCI, …) hoặc tự phát sinh (thu thập, giả lập dữ liệu).
  - Dataset phải có ít nhất 6 trường dữ liệu chính (columns/features).
  - Dataset phải có ít nhất 200 bản ghi (rows/records).
  - Dataset cần được thiết kế sao cho đáp ứng được cả hai phần yêu cầu:
    - **Phần A (Power BI – Data-Driven DSS):** phục vụ phân tích mô tả, trực quan hóa, dashboard.
    - **Phần B (Excel – Model-Driven DSS):** có dữ liệu đủ để mô phỏng/tối ưu hóa, thực hiện các công cụ What-if, Scenario, Goal Seek, Solver.

### Phần A – Data-Driven DSS với Power BI

#### Xử lý dữ liệu

- Làm sạch dữ liệu: xử lý dữ liệu trống, trùng lặp, định dạng sai.
- Đảm bảo dữ liệu thống nhất để đưa vào phân tích.

#### Phân tích mô tả với Power BI

- Xác định mục tiêu dashboard: giải thích dashboard phục vụ cho mục đích gì (quản lý kinh doanh, theo dõi học tập, quản lý bệnh nhân, …).
- Xác định đối tượng sử dụng: ví dụ nhà quản lý, giảng viên, chuyên viên y tế, quản lý marketing.

#### Xác định các KPI chính

- Tối thiểu 5 KPI phải được thể hiện trên dashboard.
- Nêu rõ:
  - Tên KPI (ví dụ: Doanh thu trung bình, Tỷ lệ khách hàng quay lại, …).
  - Cách tính (công thức hoặc cách tổng hợp dữ liệu).
  - Ý nghĩa (giúp nhà quản lý ra quyết định gì).

#### Thiết kế lược đồ chiều (Star Schema)

- Xác định bảng fact (chứa dữ liệu định lượng như doanh thu, số đơn hàng, …).
- Xác định bảng dimension (mô tả thuộc tính như khách hàng, sản phẩm, thời gian, …).
- Vẽ sơ đồ lược đồ sao (star schema) minh họa.

#### Truy vấn dữ liệu

- Viết truy vấn hoặc thiết kế ETL để đưa dữ liệu vào fact/dimension trong star schema.
- Chứng minh dữ liệu được tổ chức hợp lý để phân tích.

#### Thiết kế dashboard trên Power BI

- Trình bày các KPI đã chọn bằng biểu đồ, bảng, hoặc các visual khác phù hợp.
- Dashboard cần đảm bảo: trực quan, rõ ràng, dễ hiểu, phục vụ tốt cho việc ra quyết định.

#### Phân tích & Khuyến nghị

- Dựa trên dashboard, nhóm chỉ ra những insight quan trọng cho biết thông tin trên dashboard có thể giúp doanh nghiệp/tổ chức ra quyết định gì.

### Phần B – Model-Driven DSS với Excel

#### Xác định 4 bài toán cần ra quyết định

- Nhóm lựa chọn ít nhất 03 bài toán thực tế liên quan đến chủ đề đã chọn.
- Ví dụ:
  - **Marketing:** tối ưu chi phí quảng cáo để đạt doanh thu mục tiêu.
  - **Sản xuất:** lập kế hoạch sản xuất để giảm chi phí vận hành.
  - **Logistics:** tối ưu tồn kho hoặc phân bổ nguồn lực để giảm chi phí lưu kho.
  - **Tài chính:** phân bổ ngân sách cho các dự án khác nhau.
  - **Quản trị kinh doanh:** tính điểm hòa vốn để xác định mức sản lượng tối thiểu cần bán.

#### Xác định các thành phần của mô hình phân tích cho từng bài toán

- **Biến quyết định (Decision Variables):** những yếu tố mà nhà quản lý có thể điều chỉnh. Ví dụ: số tiền phân bổ cho từng kênh quảng cáo, số sản phẩm sản xuất mỗi ngày.
- **Biến kết quả (Objective/Outcome Variables):** mục tiêu cần đạt. Ví dụ: tổng lợi nhuận, tổng chi phí, doanh thu.
- **Công thức/mô hình (Formulas/Model):** các quan hệ toán học hoặc công thức Excel mô tả sự phụ thuộc giữa biến quyết định và biến kết quả. Ví dụ: Lợi nhuận = Doanh thu – Chi phí.
- **Ràng buộc (Constraints):** các giới hạn mà biến quyết định phải tuân theo. Ví dụ: ngân sách quảng cáo ≤ 100 triệu, sản lượng ≤ năng lực nhà máy, số giờ làm việc ≤ 160 giờ/tháng.

#### Sử dụng Excel để phân tích và tối ưu

- **What-if Analysis (Data Table):** phân tích ảnh hưởng khi thay đổi giá trị một hoặc nhiều biến đầu vào.
- **Scenario Analysis:** so sánh kết quả của các kịch bản khác nhau (ví dụ: ngân sách thấp, trung bình, cao).
- **Goal Seek:** tìm giá trị đầu vào để đạt được mục tiêu đầu ra cụ thể (ví dụ: doanh thu tối thiểu 1 tỷ).
- **Solver Optimization:** giải bài toán tối ưu hóa có ràng buộc (ví dụ: tối đa hóa lợi nhuận, tối thiểu hóa chi phí).

#### Diễn giải và rút ra ý nghĩa thực tiễn

Với mỗi bài toán, nhóm cần trình bày:

- Kết quả đầu ra từ Excel (bảng, biểu đồ, giải pháp tối ưu).
- Ý nghĩa quản lý:
  - Giải pháp này giúp nhà quản lý ra quyết định gì?
  - Nếu áp dụng vào thực tế, kết quả mang lại lợi ích gì (giảm chi phí, tăng lợi nhuận, tối ưu nguồn lực, …)?
  - Có hạn chế hay giả định nào cần lưu ý không?

## 3. QUY ĐỊNH NỘP BÀI

- Tạo thư mục nộp bài đặt tên `MaNhom_MSHV1_MSHV2_MSHV3`.
  - Ví dụ: `1_24C12001_24C12002_24C12003`
- Nội dung bên trong thư mục nộp bài bao gồm:
  - **Báo cáo nhóm** (đặt tên: `MaNhom_BaoCao.pdf`), bao gồm:
    - Danh sách thành viên và phân công.
    - Giới thiệu chủ đề lựa chọn và các nhu cầu phân tích trên chủ đề này.
    - Giới thiệu dataset và các kỹ thuật tiền xử lý dữ liệu sử dụng.
    - Kết quả phần A theo các nội dung ở phần yêu cầu đồ án.
    - Kết quả phần B theo các nội dung ở phần yêu cầu đồ án.
    - Cấu trúc chi tiết của báo cáo:
      - Trang bìa: tên môn học, tên đề tài, mã nhóm, danh sách thành viên.
      - Danh sách thành viên và phân công: ghi rõ ai làm phần nào.
      - Giới thiệu chủ đề lựa chọn và mô tả các nhu cầu phân tích đặt ra.
      - Giới thiệu dataset: nguồn gốc, cấu trúc, các bước tiền xử lý dữ liệu (cleaning, transform, integration).
      - Kết quả Phần A (Power BI): mô hình dữ liệu, KPI, dashboard, insight, khuyến nghị.
      - Kết quả Phần B (Excel – Prescriptive Models): mô tả bài toán, mô hình phân tích, công cụ Excel đã dùng (What-if, Scenario, Goal Seek, Solver), kết quả tối ưu, ý nghĩa quản lý.
      - Kết luận & hướng phát triển.
  - **File Power BI** (đặt tên: `MaNhom_Dashboard.pbix`) chứa dashboard hoàn chỉnh với các KPIs đã trình bày trong báo cáo.
  - **File Excel** (đặt tên: `MaNhom_Model.xlsx`) chứa các mô hình phân tích và tối ưu hóa.
  - **Slide thuyết trình** (đặt tên: `MaNhom_TrinhBay.pptx`), bao gồm:
    - Giới thiệu đề tài & dataset.
    - Kết quả chính từ Phần A (Power BI).
    - Kết quả chính từ Phần B (Excel).
    - Kết luận & khuyến nghị.

## 4. YÊU CẦU TRÌNH BÀY

### Giới thiệu

- Bối cảnh, lý do chọn đề tài, mục tiêu DSS.
- Mô tả dataset.

### Phần A – Data-Driven DSS

- Demo Power BI Dashboard.
- Trình bày OLAP thao tác (slice/dice, drill-down, pivot).
- Nêu ít nhất 3 insight quan trọng.

### Phần B – Model-Driven DSS

- Demo Excel.
- Giải thích ý nghĩa kết quả và ứng dụng quản lý.

### Kết luận

- Tóm tắt kết quả.
- Hạn chế của hệ DSS.
- Đề xuất hướng mở rộng.

## 5. TIÊU CHÍ ĐÁNH GIÁ (100%)

| Hạng mục | Tỷ trọng | Nội dung đánh giá |
|---|---|---|
| Báo cáo | 25% | Mô tả bài toán rõ ràng, có mục tiêu và bối cảnh thực tế; xử lý dữ liệu hợp lý, mô tả dataset đầy đủ; phân tích kết quả, lập luận logic, có khuyến nghị |
| Power BI – Data-Driven DSS | 25% | Dashboard trực quan, rõ KPI, có sử dụng OLAP; trình bày báo cáo đẹp, dễ hiểu, có insight hữu ích |
| Excel – Model-Driven DSS | 25% | Ứng dụng được Excel trong phân tích dữ liệu; phân tích & giải thích ý nghĩa kết quả |
| Thuyết trình | 25% | Slide rõ ràng, thuyết trình và trả lời câu hỏi mạch lạc |
