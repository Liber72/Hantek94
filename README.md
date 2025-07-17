# VinGroup Financial Analysis Tool

Công cụ phân tích báo cáo tài chính VinGroup dành cho sinh viên - Một bài tập hoàn chỉnh về phân tích tài chính doanh nghiệp.

## 🎯 Mục tiêu

Tạo ra một file Excel hoàn chỉnh để sinh viên thực hành phân tích báo cáo tài chính của Tập đoàn VinGroup (VIC) với dữ liệu thực tế cho năm 2023 và 2024.

## 📋 Tính năng chính

### 📊 Báo cáo tài chính đầy đủ
- **Bảng cân đối kế toán** (Balance Sheet) 2023-2024
- **Báo cáo kết quả kinh doanh** (Income Statement) 2023-2024  
- **Báo cáo lưu chuyển tiền tệ** (Cash Flow Statement) 2023-2024

### 🔍 Phân tích chỉ số tài chính
- **Chỉ số thanh khoản**: Current Ratio, Quick Ratio, Cash Ratio
- **Chỉ số sinh lời**: ROE, ROA, Net Profit Margin, Gross Profit Margin
- **Chỉ số hiệu quả**: Asset Turnover, Inventory Turnover
- **Chỉ số cơ cấu tài chính**: Debt-to-Equity, Debt-to-Assets

### 📚 Tài liệu hướng dẫn
- Hướng dẫn cách đọc báo cáo tài chính
- Bài tập thực hành cho sinh viên
- Gợi ý cách trình bày kết quả phân tích

## 🛠️ Cài đặt

### Yêu cầu hệ thống
- Python 3.8+
- pip (Python package manager)

### Cài đặt thư viện
```bash
pip install -r requirements.txt
```

Hoặc cài đặt từng thư viện:
```bash
pip install openpyxl pandas matplotlib numpy
```

## 🚀 Cách sử dụng

### 1. Tạo files CSV (không cần thư viện bổ sung)
```bash
python vingroup_financial_analyzer.py
```

Kết quả: Tạo thư mục `vingroup_analysis` với các file CSV:
- `balance_sheet.csv` - Bảng cân đối kế toán
- `income_statement.csv` - Báo cáo kết quả kinh doanh
- `cash_flow.csv` - Báo cáo lưu chuyển tiền tệ
- `financial_ratios.csv` - Phân tích chỉ số tài chính
- `guidelines_exercises.csv` - Hướng dẫn và bài tập

### 2. Tạo file Excel hoàn chỉnh (cần openpyxl)
```bash
python excel_generator.py
```

Kết quả: File `VinGroup_Financial_Analysis.xlsx` với 3 sheet:
1. **Báo cáo tài chính VinGroup** - Báo cáo tài chính đầy đủ
2. **Phân tích chỉ số tài chính** - Tính toán và phân tích chỉ số
3. **Hướng dẫn và Bài tập** - Tài liệu học tập

## 📁 Cấu trúc dự án

```
Hantek94/
├── README.md                           # Tài liệu hướng dẫn
├── requirements.txt                    # Danh sách thư viện cần thiết
├── vingroup_financial_analyzer.py     # Script tạo CSV và phân tích cơ bản
├── excel_generator.py                 # Script tạo file Excel hoàn chỉnh
└── vingroup_analysis/                 # Thư mục chứa kết quả
    ├── balance_sheet.csv
    ├── income_statement.csv
    ├── cash_flow.csv
    ├── financial_ratios.csv
    ├── guidelines_exercises.csv
    └── vingroup_data.json            # Dữ liệu gốc JSON
```

## 💡 Dữ liệu tài chính

### Thông tin công ty
- **Tên công ty**: Tập đoàn VinGroup
- **Mã chứng khoán**: VIC
- **Ngành**: Đa ngành (Bất động sản, Bán lẻ, Công nghiệp)
- **Đơn vị**: Tỷ VND

### Báo cáo bao gồm
- **Tài sản**: Tài sản ngắn hạn, dài hạn
- **Nợ phải trả**: Nợ ngắn hạn, dài hạn
- **Vốn chủ sở hữu**: Vốn góp, lợi nhuận chưa phân phối
- **Doanh thu và chi phí**: Doanh thu, giá vốn, chi phí hoạt động
- **Lưu chuyển tiền tệ**: Từ hoạt động kinh doanh, đầu tư, tài chính

## 🔢 Chỉ số tài chính được tính toán

### Chỉ số thanh khoản
```
Current Ratio = Tài sản ngắn hạn / Nợ ngắn hạn
Quick Ratio = (Tài sản ngắn hạn - Hàng tồn kho) / Nợ ngắn hạn
Cash Ratio = (Tiền mặt + Đầu tư ngắn hạn) / Nợ ngắn hạn
```

### Chỉ số sinh lời
```
ROE = Lợi nhuận sau thuế / Vốn chủ sở hữu × 100%
ROA = Lợi nhuận sau thuế / Tổng tài sản × 100%
Net Profit Margin = Lợi nhuận sau thuế / Doanh thu × 100%
Gross Profit Margin = Lợi nhuận gộp / Doanh thu × 100%
```

### Chỉ số hiệu quả
```
Asset Turnover = Doanh thu / Tổng tài sản
Inventory Turnover = Giá vốn hàng bán / Hàng tồn kho
```

### Chỉ số cơ cấu tài chính
```
Debt-to-Equity = Tổng nợ / Vốn chủ sở hữu
Debt-to-Assets = Tổng nợ / Tổng tài sản
```

## 📈 Kết quả mẫu

```
Hệ số thanh khoản hiện hành: 2023=2.60, 2024=2.66
ROE: 2023=11.69%, 2024=11.58%
ROA: 2023=4.52%, 2024=4.56%
Biên lợi nhuận ròng: 2023=11.28%, 2024=11.40%
```

## 🎓 Bài tập cho sinh viên

### Câu hỏi thực hành
1. **Phân tích cơ cấu tài sản**: Tính tỷ trọng tài sản ngắn hạn/dài hạn
2. **Đánh giá thanh khoản**: Phân tích khả năng thanh toán nợ
3. **Phân tích sinh lời**: Đánh giá hiệu quả kinh doanh
4. **Hiệu quả hoạt động**: Phân tích vòng quay tài sản
5. **Cơ cấu tài chính**: Đánh giá rủi ro tài chính

### Hướng dẫn trình bày
- Sử dụng bảng biểu và biểu đồ
- So sánh giữa các năm
- Phân tích xu hướng
- Đưa ra nhận xét và khuyến nghị

## ⚡ Tính năng Excel nâng cao

### Định dạng chuyên nghiệp
- Màu sắc phân loại theo nội dung
- Font chữ và viền đẹp mắt
- Định dạng số và phần trăm chuẩn

### Công thức tự động
- Tính toán chỉ số tài chính tự động
- Công thức Excel có thể chỉnh sửa
- Validation dữ liệu đầu vào

### Biểu đồ minh họa
- Biểu đồ so sánh các chỉ số
- Xu hướng thay đổi qua thời gian
- Biểu đồ cơ cấu tài sản

## 🐛 Xử lý lỗi

### Lỗi thiếu thư viện
```bash
pip install openpyxl pandas matplotlib
```

### Lỗi quyền ghi file
```bash
chmod +w VinGroup_Financial_Analysis.xlsx
```

### Lỗi encoding
Đảm bảo Python sử dụng UTF-8 encoding cho tiếng Việt.

## 📞 Hỗ trợ

Nếu gặp vấn đề trong quá trình sử dụng:
1. Kiểm tra requirements.txt đã cài đặt đầy đủ
2. Đảm bảo Python version 3.8+
3. Kiểm tra quyền ghi file trong thư mục

## 📜 Giấy phép

Dự án này được phát triển cho mục đích giáo dục, dữ liệu tài chính dựa trên cấu trúc thực tế của VinGroup nhưng đã được điều chỉnh phù hợp cho việc học tập.

## 🔄 Cập nhật

- **Version 1.0**: Tạo báo cáo tài chính cơ bản và phân tích chỉ số
- **Version 1.1**: Thêm biểu đồ và định dạng Excel nâng cao
- **Version 1.2**: Bổ sung bài tập thực hành và hướng dẫn chi tiết

---

**Tác giả**: Financial Analysis Tool  
**Ngày tạo**: 2024  
**Mục đích**: Giáo dục - Phân tích tài chính doanh nghiệp