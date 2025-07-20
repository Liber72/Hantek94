# Hệ thống Phân tích Tài chính - Financial Analysis System

Bộ code Python hoàn chỉnh để tạo file Excel bảng cân đối kế toán và phân tích tài chính theo chuẩn mực kế toán Việt Nam.

## 📋 Tính năng chính

### 1. **Bảng cân đối kế toán Excel**
- ✅ Dữ liệu mẫu thực tế với nguồn gốc rõ ràng
- ✅ Cấu trúc theo Thông tư 200/2014/TT-BTC
- ✅ Các khoản mục: Tài sản, Nợ phải trả, Vốn chủ sở hữu
- ✅ Định dạng chuyên nghiệp với style phù hợp
- ✅ Kiểm tra cân đối tự động

### 2. **File phân tích tài chính**
- ✅ Các chỉ số thanh khoản (Current Ratio, Quick Ratio, Cash Ratio)
- ✅ Chỉ số đòn bẩy tài chính (Debt-to-Equity, Debt Ratio)
- ✅ Chỉ số hiệu quả (ROA, ROE, Asset Turnover)
- ✅ Công thức Excel liên kết tự động
- ✅ Biểu đồ trực quan hóa dữ liệu
- ✅ Đánh giá và khuyến nghị cho từng chỉ số

## 🚀 Cách sử dụng

### Yêu cầu hệ thống
```bash
Python 3.7+
pip install openpyxl pandas numpy matplotlib
```

### Chạy nhanh
```bash
# Di chuyển vào thư mục
cd financial_analysis

# Cài đặt dependencies
pip install -r requirements.txt

# Chạy toàn bộ hệ thống
python main.py
```

### Sử dụng trong code
```python
from main import FinancialAnalysisSystem

# Tạo hệ thống
system = FinancialAnalysisSystem("my_output")

# Chạy phân tích hoàn chỉnh
result = system.run_complete_analysis()

print(f"Files created: {result['balance_sheet_file']}")
print(f"Analysis file: {result['financial_analysis_file']}")
```

## 📁 Cấu trúc dự án

```
financial_analysis/
├── main.py                           # File chính chạy toàn bộ hệ thống
├── balance_sheet_generator.py        # Tạo bảng cân đối kế toán
├── financial_analysis_generator.py   # Tạo file phân tích tài chính
├── data_source.py                    # Dữ liệu mẫu và truy xuất
├── requirements.txt                  # Dependencies
└── output/                          # Thư mục chứa file Excel được tạo
    ├── bang_can_doi_ke_toan_[timestamp].xlsx
    └── phan_tich_tai_chinh_[timestamp].xlsx
```

## 📊 Dữ liệu mẫu

**Nguồn dữ liệu:**
- Dựa trên cấu trúc kế toán chuẩn Việt Nam (Thông tư 200/2014/TT-BTC)
- Tham khảo các báo cáo tài chính công khai
- Dữ liệu mang tính chất minh họa, phù hợp cho học tập và phân tích

**Thông tin công ty mẫu:**
- Tên: CÔNG TY CỔ PHẦN MẪU XYZ
- Kỳ báo cáo: 31/12/2023
- Đơn vị tính: VND (triệu đồng)
- Tổng tài sản: 280,000 triệu VND

## 📈 Các chỉ số phân tích

### Chỉ số thanh khoản
- **Hệ số thanh khoản hiện tại**: Tài sản ngắn hạn / Nợ ngắn hạn
- **Hệ số thanh khoản nhanh**: (Tài sản ngắn hạn - Hàng tồn kho) / Nợ ngắn hạn
- **Hệ số tiền mặt**: Tiền và tương đương tiền / Nợ ngắn hạn

### Chỉ số đòn bẩy tài chính
- **Tỷ số nợ/Tài sản**: Tổng nợ / Tổng tài sản
- **Tỷ số nợ/Vốn CSH**: Tổng nợ / Vốn chủ sở hữu
- **Hệ số nhân vốn**: Tổng tài sản / Vốn chủ sở hữu

### Chỉ số hiệu quả
- **ROA**: Lợi nhuận sau thuế / Tổng tài sản × 100
- **ROE**: Lợi nhuận sau thuế / Vốn chủ sở hữu × 100
- **Vòng quay tài sản**: Doanh thu thuần / Tổng tài sản
- **Tỷ lệ lợi nhuận**: Lợi nhuận sau thuế / Doanh thu × 100

## 🎯 Tính năng nổi bật

### 1. **Định dạng chuyên nghiệp**
- Font Times New Roman cho bảng cân đối kế toán
- Font Calibri cho file phân tích
- Màu sắc phân loại theo mức độ đánh giá
- Border và alignment chuẩn

### 2. **Công thức Excel linh hoạt**
- Tự động tính toán các chỉ số
- Có thể chỉnh sửa dữ liệu để cập nhật
- Named ranges để dễ tham chiếu

### 3. **Biểu đồ trực quan**
- Biểu đồ tròn cơ cấu tài sản
- Biểu đồ cột so sánh chỉ số
- Màu sắc phân biệt rõ ràng

### 4. **Đánh giá tự động**
- Phân loại chỉ số: Tốt, Trung bình, Kém
- Màu nền tương ứng: Xanh, Vàng, Đỏ
- Ghi chú và khuyến nghị cụ thể

## 📖 Hướng dẫn sử dụng file Excel

1. **Mở file bảng cân đối kế toán**
   - Xem cấu trúc tài sản và nguồn vốn
   - Kiểm tra tính cân đối của bảng
   - Có thể chỉnh sửa số liệu

2. **Mở file phân tích tài chính**
   - Sheet "Tổng quan": Xem các chỉ số quan trọng
   - Sheet "Phân tích thanh khoản": Chi tiết về khả năng thanh toán
   - Sheet "Phân tích đòn bẩy": Đánh giá cơ cấu tài chính
   - Sheet "Phân tích hiệu quả": Hiệu quả sử dụng tài sản
   - Sheet "Biểu đồ phân tích": Trực quan hóa dữ liệu

## ⚠️ Lưu ý quan trọng

- **Dữ liệu mẫu**: Chỉ mang tính chất minh họa
- **Xác minh**: Cần kiểm tra với dữ liệu thực tế khi sử dụng
- **Tuân thủ**: Tuân thủ quy định pháp luật về kế toán
- **Backup**: Sao lưu file trước khi chỉnh sửa

## 🔧 Tùy chỉnh

### Thay đổi dữ liệu
```python
# Chỉnh sửa trong data_source.py
class FinancialDataSource:
    def get_balance_sheet_data(self):
        # Thay đổi số liệu ở đây
        assets_data = {
            # Cập nhật dữ liệu của bạn
        }
```

### Thay đổi công thức
```python
# Chỉnh sửa trong financial_analysis_generator.py
def _calculate_key_metrics(self):
    # Thêm hoặc sửa công thức tính toán
```

### Thay đổi style
```python
# Chỉnh sửa trong balance_sheet_generator.py
def _define_styles(self):
    # Tùy chỉnh font, màu sắc, border
```

## 🤝 Đóng góp

Chào mừng các đóng góp để cải thiện hệ thống:
1. Fork dự án
2. Tạo feature branch
3. Commit changes
4. Push to branch
5. Tạo Pull Request

## 📄 Giấy phép

Dự án này được phát hành dưới giấy phép MIT - xem file LICENSE để biết chi tiết.

---

**Phát triển bởi:** Financial Analysis System Team  
**Phiên bản:** 1.0  
**Ngày cập nhật:** 20/07/2025  

✨ *Tạo báo cáo tài chính chuyên nghiệp với Python và Excel* ✨