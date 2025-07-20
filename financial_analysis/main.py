"""
File chính để chạy toàn bộ hệ thống phân tích tài chính
Main file to run the complete financial analysis system

Chức năng:
- Tạo file Excel bảng cân đối kế toán với dữ liệu mẫu
- Tạo file Excel phân tích tài chính với các chỉ số và biểu đồ
- Liên kết giữa các file để tự động cập nhật
- Xuất báo cáo tổng hợp

Sử dụng:
    python main.py
    
Hoặc import và sử dụng:
    from main import FinancialAnalysisSystem
    system = FinancialAnalysisSystem()
    system.run_complete_analysis()
"""

import os
import sys
from datetime import datetime
import traceback

# Import các module tự tạo
from data_source import FinancialDataSource, get_sample_data
from balance_sheet_generator import BalanceSheetGenerator, create_balance_sheet_file
from financial_analysis_generator import FinancialAnalysisGenerator, create_financial_analysis_file

class FinancialAnalysisSystem:
    """Lớp chính điều phối toàn bộ hệ thống phân tích tài chính"""
    
    def __init__(self, output_directory="output"):
        """
        Khởi tạo hệ thống
        
        Args:
            output_directory (str): Thư mục lưu các file output
        """
        self.output_directory = output_directory
        self.data_source = FinancialDataSource()
        
        # Tạo thư mục output nếu chưa tồn tại
        if not os.path.exists(output_directory):
            os.makedirs(output_directory)
            print(f"✓ Đã tạo thư mục output: {output_directory}")
        
        # Biến lưu đường dẫn các file đã tạo
        self.balance_sheet_file = None
        self.financial_analysis_file = None
        
    def print_header(self):
        """In header cho chương trình"""
        print("=" * 80)
        print("HỆ THỐNG PHÂN TÍCH TÀI CHÍNH - FINANCIAL ANALYSIS SYSTEM")
        print("=" * 80)
        print("Phiên bản: 1.0")
        print("Tác giả: Financial Analysis System")
        print(f"Ngày chạy: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        print("-" * 80)
        print("Chức năng:")
        print("• Tạo bảng cân đối kế toán Excel với dữ liệu mẫu")
        print("• Tạo file phân tích tài chính với các chỉ số và biểu đồ")
        print("• Áp dụng chuẩn mực kế toán Việt Nam và quốc tế")
        print("• Xuất báo cáo với công thức Excel linh hoạt")
        print("=" * 80)
    
    def validate_system(self):
        """Kiểm tra tính sẵn sàng của hệ thống"""
        print("🔍 KIỂM TRA HỆ THỐNG...")
        
        # Kiểm tra dữ liệu mẫu
        try:
            sample_data = self.data_source.get_balance_sheet_data()
            print("✓ Dữ liệu mẫu: OK")
            
            # Hiển thị thông tin công ty mẫu
            company_info = sample_data['company_info']
            print(f"  - Công ty: {company_info['name']}")
            print(f"  - Kỳ báo cáo: {company_info['period']}")
            print(f"  - Đơn vị tính: {company_info['unit']}")
            
        except Exception as e:
            print(f"❌ Lỗi dữ liệu mẫu: {str(e)}")
            return False
        
        # Kiểm tra khả năng tạo file
        try:
            test_path = os.path.join(self.output_directory, "test.txt")
            with open(test_path, 'w') as f:
                f.write("test")
            os.remove(test_path)
            print("✓ Quyền ghi file: OK")
        except Exception as e:
            print(f"❌ Lỗi quyền ghi file: {str(e)}")
            return False
        
        # Kiểm tra thư viện
        try:
            import openpyxl
            print("✓ Thư viện openpyxl: OK")
        except ImportError:
            print("❌ Chưa cài đặt thư viện openpyxl")
            print("   Chạy: pip install openpyxl")
            return False
        
        print("✅ Hệ thống sẵn sàng!\n")
        return True
    
    def create_balance_sheet(self, filename=None):
        """
        Tạo file bảng cân đối kế toán
        
        Args:
            filename (str): Tên file tùy chọn
            
        Returns:
            str: Đường dẫn file đã tạo
        """
        print("📊 ĐANG TẠO BẢNG CÂN ĐỐI KẾ TOÁN...")
        
        try:
            generator = BalanceSheetGenerator(self.output_directory)
            
            if filename is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"bang_can_doi_ke_toan_{timestamp}.xlsx"
            
            filepath = generator.create_balance_sheet(filename)
            self.balance_sheet_file = filepath
            
            print(f"✅ Đã tạo thành công: {os.path.basename(filepath)}")
            print(f"   Đường dẫn: {filepath}")
            
            # Hiển thị thông tin file
            file_size = os.path.getsize(filepath)
            print(f"   Kích thước: {file_size:,} bytes")
            
            return filepath
            
        except Exception as e:
            print(f"❌ Lỗi khi tạo bảng cân đối kế toán: {str(e)}")
            print("Chi tiết lỗi:")
            traceback.print_exc()
            return None
    
    def create_financial_analysis(self, filename=None):
        """
        Tạo file phân tích tài chính
        
        Args:
            filename (str): Tên file tùy chọn
            
        Returns:
            str: Đường dẫn file đã tạo
        """
        print("\n📈 ĐANG TẠO FILE PHÂN TÍCH TÀI CHÍNH...")
        
        try:
            generator = FinancialAnalysisGenerator(
                self.output_directory, 
                self.balance_sheet_file
            )
            
            if filename is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"phan_tich_tai_chinh_{timestamp}.xlsx"
            
            filepath = generator.create_financial_analysis(filename)
            self.financial_analysis_file = filepath
            
            print(f"✅ Đã tạo thành công: {os.path.basename(filepath)}")
            print(f"   Đường dẫn: {filepath}")
            
            # Hiển thị thông tin file
            file_size = os.path.getsize(filepath)
            print(f"   Kích thước: {file_size:,} bytes")
            
            return filepath
            
        except Exception as e:
            print(f"❌ Lỗi khi tạo file phân tích tài chính: {str(e)}")
            print("Chi tiết lỗi:")
            traceback.print_exc()
            return None
    
    def print_summary_report(self):
        """In báo cáo tổng kết"""
        print("\n" + "=" * 80)
        print("📋 BÁO CÁO TỔNG KẾT")
        print("=" * 80)
        
        # Thông tin file đã tạo
        print("📁 CÁC FILE ĐÃ TẠO:")
        if self.balance_sheet_file:
            print(f"   • Bảng cân đối kế toán: {os.path.basename(self.balance_sheet_file)}")
            print(f"     Đường dẫn: {self.balance_sheet_file}")
        
        if self.financial_analysis_file:
            print(f"   • Phân tích tài chính: {os.path.basename(self.financial_analysis_file)}")
            print(f"     Đường dẫn: {self.financial_analysis_file}")
        
        # Thống kê dữ liệu
        print("\n📊 THỐNG KÊ DỮ LIỆU:")
        data = self.data_source.get_balance_sheet_data()
        
        # Tính toán tổng quan
        total_assets = sum(sum(item['value'] for item in section['items'].values()) 
                          for section in data['assets'].values())
        total_liabilities = sum(item['value'] for item in 
                               data['liabilities_equity']['C_NO_PHAI_TRA']['items'].values())
        total_equity = sum(item['value'] for item in 
                          data['liabilities_equity']['D_VON_CHU_SO_HUU']['items'].values())
        
        print(f"   • Tổng tài sản: {total_assets:,} triệu VND")
        print(f"   • Tổng nợ phải trả: {total_liabilities:,} triệu VND")
        print(f"   • Tổng vốn chủ sở hữu: {total_equity:,} triệu VND")
        print(f"   • Kiểm tra cân đối: {total_assets == (total_liabilities + total_equity)}")
        
        # Thông tin kỹ thuật
        print("\n🔧 THÔNG TIN KỸ THUẬT:")
        print("   • Chuẩn mực: Thông tư 200/2014/TT-BTC")
        print("   • Định dạng: Excel (.xlsx)")
        print("   • Công thức: Tự động tính toán")
        print("   • Biểu đồ: Có hỗ trợ trực quan hóa")
        
        # Hướng dẫn sử dụng
        print("\n📖 HƯỚNG DẪN SỬ DỤNG:")
        print("   1. Mở file Excel đã tạo")
        print("   2. File bảng cân đối kế toán chứa dữ liệu cơ bản")
        print("   3. File phân tích tài chính chứa:")
        print("      - Các sheet phân tích theo từng nhóm chỉ số")
        print("      - Biểu đồ trực quan")
        print("      - Đánh giá và khuyến nghị")
        print("   4. Có thể chỉnh sửa dữ liệu để cập nhật tự động")
        
        # Lưu ý quan trọng
        print("\n⚠️  LƯU Ý QUAN TRỌNG:")
        print("   • Dữ liệu mang tính chất minh họa")
        print("   • Cần xác minh với dữ liệu thực tế khi sử dụng")
        print("   • Tuân thủ quy định pháp luật về kế toán")
        print("   • Backup file trước khi chỉnh sửa")
        
        print("=" * 80)
    
    def run_complete_analysis(self, balance_sheet_filename=None, analysis_filename=None):
        """
        Chạy toàn bộ quy trình phân tích tài chính
        
        Args:
            balance_sheet_filename (str): Tên file bảng cân đối kế toán
            analysis_filename (str): Tên file phân tích tài chính
            
        Returns:
            dict: Thông tin về các file đã tạo
        """
        
        # In header
        self.print_header()
        
        # Kiểm tra hệ thống
        if not self.validate_system():
            print("❌ Hệ thống chưa sẵn sàng. Vui lòng khắc phục các lỗi trên.")
            return None
        
        # Tạo file bảng cân đối kế toán
        balance_sheet_path = self.create_balance_sheet(balance_sheet_filename)
        if not balance_sheet_path:
            print("❌ Không thể tiếp tục do lỗi tạo bảng cân đối kế toán")
            return None
        
        # Tạo file phân tích tài chính
        analysis_path = self.create_financial_analysis(analysis_filename)
        if not analysis_path:
            print("❌ Không thể tạo file phân tích tài chính")
            return None
        
        # In báo cáo tổng kết
        self.print_summary_report()
        
        return {
            'balance_sheet_file': balance_sheet_path,
            'financial_analysis_file': analysis_path,
            'output_directory': self.output_directory,
            'status': 'success'
        }
    
    def get_data_sources_info(self):
        """Lấy thông tin về nguồn dữ liệu"""
        return self.data_source.get_data_sources_info()

def main():
    """Hàm main để chạy chương trình"""
    
    # Tạo hệ thống
    system = FinancialAnalysisSystem()
    
    # Chạy phân tích hoàn chỉnh
    result = system.run_complete_analysis()
    
    if result and result['status'] == 'success':
        print("\n🎉 HOÀN THÀNH THÀNH CÔNG!")
        print("Bạn có thể tìm thấy các file Excel trong thư mục 'output'")
        
        # Hỏi người dùng có muốn mở file không (chỉ để tham khảo)
        print("\n💡 GỢI Ý:")
        print("- Có thể mở các file Excel để xem kết quả")
        print("- Sử dụng các file làm template cho dự án thực tế")
        print("- Tùy chỉnh dữ liệu và công thức theo nhu cầu")
        
    else:
        print("\n❌ QUY TRÌNH THẤT BẠI!")
        print("Vui lòng kiểm tra lại các lỗi và thử lại.")
    
    return result

def create_sample_files(output_dir="output"):
    """
    Hàm tiện ích để tạo nhanh các file mẫu
    
    Args:
        output_dir (str): Thư mục output
        
    Returns:
        dict: Thông tin file đã tạo
    """
    system = FinancialAnalysisSystem(output_dir)
    return system.run_complete_analysis()

def print_usage():
    """In hướng dẫn sử dụng"""
    print("""
HƯỚNG DẪN SỬ DỤNG HỆ THỐNG PHÂN TÍCH TÀI CHÍNH

1. Chạy trực tiếp:
   python main.py

2. Import và sử dụng:
   from main import FinancialAnalysisSystem
   system = FinancialAnalysisSystem()
   result = system.run_complete_analysis()

3. Tạo file nhanh:
   from main import create_sample_files
   result = create_sample_files("my_output")

CÁC FILE SẼ ĐƯỢC TẠO:
- bang_can_doi_ke_toan_[timestamp].xlsx
- phan_tich_tai_chinh_[timestamp].xlsx

YÊU CẦU HỆ THỐNG:
- Python 3.7+
- openpyxl
- pandas (tùy chọn)

CÀI ĐẶT THỦ CÔNG:
pip install openpyxl pandas

LIÊN HỆ HỖ TRỢ:
Nếu gặp lỗi, vui lòng kiểm tra:
1. Quyền ghi file trong thư mục
2. Phiên bản Python và thư viện
3. Dung lượng ổ đĩa
""")

if __name__ == "__main__":
    # Kiểm tra tham số dòng lệnh
    if len(sys.argv) > 1:
        if sys.argv[1] in ['-h', '--help', 'help']:
            print_usage()
            sys.exit(0)
        elif sys.argv[1] in ['-v', '--version', 'version']:
            print("Financial Analysis System v1.0")
            print("Tạo bảng cân đối kế toán và phân tích tài chính Excel")
            sys.exit(0)
    
    # Chạy chương trình chính
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️  Chương trình bị dừng bởi người dùng")
    except Exception as e:
        print(f"\n❌ Lỗi không mong muốn: {str(e)}")
        print("Chi tiết lỗi:")
        traceback.print_exc()
        print("\nVui lòng báo cáo lỗi này để được hỗ trợ.")