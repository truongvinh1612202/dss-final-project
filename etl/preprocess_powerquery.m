// Tiền xử lý Superstore trong Power Query: tương đương etl/preprocess.py
// Cách dùng: Power BI → Transform data → New Source → Blank Query → Advanced Editor → dán toàn bộ.
// Sửa DuongDan nếu thư mục dự án nằm ở chỗ khác.
// Đối chiếu kết quả với docs/Nhat_ky_tien_xu_ly.md: 9.799 dòng, tổng Sales 2,261,255.41.
let
    DuongDan = "D:\StudyHCMUS\DSS\Final Proj\data\raw\train.csv",

    // 1. Nạp dữ liệu: UTF-8 (65001); mọi cột đọc dạng text để giữ số 0 đầu của Postal Code
    Nguon = Csv.Document(
        File.Contents(DuongDan),
        [Delimiter = ",", Columns = 18, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]
    ),
    DatTieuDe = Table.PromoteHeaders(Nguon, [PromoteAllScalars = true]),

    // Kiểu số: dấu thập phân là "." → culture en-US
    DoiKieuSo = Table.TransformColumnTypes(
        DatTieuDe, {{"Row ID", Int64.Type}, {"Sales", type number}}, "en-US"
    ),

    // 2. Xóa dòng trùng toàn bộ (bỏ qua Row ID), giữ lần xuất hiện đầu → xóa Row ID 3407
    XoaDongTrung = Table.Distinct(
        DoiKieuSo, List.RemoveItems(Table.ColumnNames(DoiKieuSo), {"Row ID"})
    ),

    // 5. Ngày dạng d/m/yyyy → BẮT BUỘC culture en-GB (en-US sẽ đảo ngày/tháng)
    DoiKieuNgay = Table.TransformColumnTypes(
        XoaDongTrung, {{"Order Date", type date}, {"Ship Date", type date}}, "en-GB"
    ),

    // 6–7. Postal Code: đệm đủ 5 chữ số; ô trống (chỉ có ở Burlington, Vermont) → 05401
    ChuanHoaPostalCode = Table.ReplaceValue(
        DoiKieuNgay,
        each [Postal Code],
        each
            if ([Postal Code] = null or [Postal Code] = "")
                and [City] = "Burlington" and [State] = "Vermont"
            then "05401"
            else Text.PadStart([Postal Code], 5, "0"),
        Replacer.ReplaceValue,
        {"Postal Code"}
    ),

    // 8. Thay NBSP (U+00A0) trong Product Name; Text.Clean KHÔNG xóa ký tự này
    ThayNBSP = Table.TransformColumns(
        ChuanHoaPostalCode,
        {{"Product Name", each Text.Replace(_, Character.FromNumber(160), " "), type text}}
    ),

    // Trim và gộp khoảng trắng liên tiếp ở mọi cột text
    CotText = List.RemoveItems(
        Table.ColumnNames(ThayNBSP), {"Row ID", "Sales", "Order Date", "Ship Date"}
    ),
    TrimText = Table.TransformColumns(
        ThayNBSP,
        List.Transform(
            CotText,
            (c) => {c, each Text.Combine(List.Select(Text.Split(Text.Trim(_), " "), (w) => w <> ""), " "), type text}
        )
    ),

    // 10. Cột bổ sung
    ThemNam = Table.AddColumn(TrimText, "Order Year", each Date.Year([Order Date]), Int64.Type),
    ThemQuy = Table.AddColumn(ThemNam, "Order Quarter", each Date.QuarterOfYear([Order Date]), Int64.Type),
    ThemThang = Table.AddColumn(ThemQuy, "Order Month", each Date.Month([Order Date]), Int64.Type),
    ThemNamThang = Table.AddColumn(
        ThemThang, "Order YearMonth", each Date.ToText([Order Date], "yyyy-MM"), type text
    ),
    ThemShipDays = Table.AddColumn(
        ThemNamThang, "Ship Days", each Duration.Days([Ship Date] - [Order Date]), Int64.Type
    ),

    // Cờ ngoại lai Sales > Q3 + 1.5·IQR (ExcelInc = cách tính tứ phân vị mặc định của pandas)
    Q1 = List.Percentile(ThemShipDays[Sales], 0.25, [PercentileMode = PercentileMode.ExcelInc]),
    Q3 = List.Percentile(ThemShipDays[Sales], 0.75, [PercentileMode = PercentileMode.ExcelInc]),
    NguongNgoaiLai = Q3 + 1.5 * (Q3 - Q1),
    ThemNgoaiLai = Table.AddColumn(
        ThemShipDays, "Is Sales Outlier", each if [Sales] > NguongNgoaiLai then 1 else 0, Int64.Type
    ),

    // Khóa thay thế ProductKey trên (Product ID, Product Name)
    BangSanPham = Table.AddIndexColumn(
        Table.Sort(
            Table.Distinct(Table.SelectColumns(ThemNgoaiLai, {"Product ID", "Product Name"})),
            {"Product ID", "Product Name"}
        ),
        "ProductKey", 1, 1, Int64.Type
    ),
    GhepSanPham = Table.Join(ThemNgoaiLai, {"Product ID", "Product Name"}, BangSanPham, {"Product ID", "Product Name"}),

    // Khóa thay thế GeoKey trên (City, State, Postal Code)
    BangDiaLy = Table.AddIndexColumn(
        Table.Sort(
            Table.Distinct(Table.SelectColumns(GhepSanPham, {"City", "State", "Postal Code"})),
            {"City", "State", "Postal Code"}
        ),
        "GeoKey", 1, 1, Int64.Type
    ),
    GhepDiaLy = Table.Join(GhepSanPham, {"City", "State", "Postal Code"}, BangDiaLy, {"City", "State", "Postal Code"}),

    KetQua = Table.Sort(GhepDiaLy, {{"Row ID", Order.Ascending}})
in
    KetQua
