"""
Main Integration System - Hệ thống Tích hợp Chính
=================================================

Điều phối toàn bộ hệ thống báo cáo tài chính động
Tích hợp tất cả các module và cung cấp giao diện thân thiện

Tác giả: Hệ thống Phân tích Tài chính Động
Chuẩn: VAS/Circular 200/2014/TT-BTC
"""

import os
import sys
import datetime
import json
from pathlib import Path

# Import các module chính
from enhanced_balance_sheet_generator import EnhancedBalanceSheetGenerator
from dynamic_financial_analyzer import DynamicFinancialAnalyzer
from formula_validator import FormulaValidator
from multi_period_analyzer import MultiPeriodAnalyzer

class MainIntegrationSystem:
    def __init__(self):
        self.system_info = {
            'name': 'Hệ thống Báo cáo Tài chính Động',
            'version': '1.0.0',
            'author': 'Dynamic Financial Analysis System',
            'standard': 'VAS/Circular 200/2014/TT-BTC',
            'created': datetime.datetime.now().isoformat()
        }
        
        self.generated_files = []
        self.validation_results = []
        
    def print_banner(self):
        """In banner hệ thống"""
        print("=" * 80)
        print("🏢 HỆ THỐNG BÁO CÁO TÀI CHÍNH ĐỘNG")
        print("📊 Dynamic Financial Reporting System")
        print("=" * 80)
        print(f"📌 Phiên bản: {self.system_info['version']}")
        print(f"📅 Chuẩn kế toán: {self.system_info['standard']}")
        print(f"🕐 Thời gian: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        print("=" * 80)
        
    def print_menu(self):
        """In menu chính"""
        print("\n📋 MENU CHÍNH:")
        print("1. 🏗️  Tạo Bảng Cân Đối Kế Toán với Named Ranges")
        print("2. 📊 Tạo Hệ thống Phân tích Tài chính Động")
        print("3. 🔍 Kiểm tra và Validation Công thức")
        print("4. 📈 Phân tích Nhiều kỳ và Xu hướng")
        print("5. 🚀 Tạo Toàn bộ Hệ thống (Tự động)")
        print("6. 📁 Xem Danh sách File đã tạo")
        print("7. 📖 Hướng dẫn Sử dụng")
        print("8. ❌ Thoát")
        print("-" * 80)
        
    def create_output_folder(self):
        """Tạo thư mục output nếu chưa tồn tại"""
        folders = ['output', 'backups', 'reports', 'templates']
        for folder in folders:
            os.makedirs(folder, exist_ok=True)
            
    def option_1_balance_sheet(self):
        """Tùy chọn 1: Tạo bảng cân đối kế toán"""
        print("\n🏗️  ĐANG TẠO BẢNG CÂN ĐỐI KẾ TOÁN...")
        print("-" * 60)
        
        try:
            generator = EnhancedBalanceSheetGenerator()
            filename = generator.generate_complete_balance_sheet()
            
            if filename:
                self.generated_files.append({
                    'type': 'balance_sheet',
                    'filename': filename,
                    'timestamp': datetime.datetime.now().isoformat(),
                    'named_ranges': len(generator.named_ranges)
                })
                
                print(f"\n✅ THÀNH CÔNG!")
                print(f"📁 File đã tạo: {filename}")
                print(f"🏷️  Named Ranges: {len(generator.named_ranges)}")
                return filename
            else:
                print("❌ Lỗi tạo bảng cân đối kế toán")
                return None
                
        except Exception as e:
            print(f"❌ Lỗi: {e}")
            return None
            
    def option_2_financial_analysis(self):
        """Tùy chọn 2: Tạo hệ thống phân tích tài chính"""
        print("\n📊 ĐANG TẠO HỆ THỐNG PHÂN TÍCH TÀI CHÍNH...")
        print("-" * 60)
        
        try:
            analyzer = DynamicFinancialAnalyzer()
            filename = analyzer.generate_complete_analysis()
            
            if filename:
                self.generated_files.append({
                    'type': 'financial_analysis',
                    'filename': filename,
                    'timestamp': datetime.datetime.now().isoformat(),
                    'sheets': len(analyzer.sheets),
                    'formulas': len(analyzer.formulas)
                })
                
                print(f"\n✅ THÀNH CÔNG!")
                print(f"📁 File đã tạo: {filename}")
                print(f"📊 Sheets: {len(analyzer.sheets)}")
                print(f"🔢 Formulas: {len(analyzer.formulas)}")
                return filename
            else:
                print("❌ Lỗi tạo hệ thống phân tích tài chính")
                return None
                
        except Exception as e:
            print(f"❌ Lỗi: {e}")
            return None
            
    def option_3_validation(self):
        """Tùy chọn 3: Kiểm tra và validation"""
        print("\n🔍 ĐANG KIỂM TRA VÀ VALIDATION HỆ THỐNG...")
        print("-" * 60)
        
        # Tìm file balance sheet mới nhất
        balance_files = [f for f in os.listdir('.') if f.startswith('bang_can_doi_ke_toan_dynamic_') and f.endswith('.xlsx')]
        
        if not balance_files:
            print("❌ Không tìm thấy file bảng cân đối kế toán")
            print("💡 Vui lòng chọn tùy chọn 1 để tạo bảng cân đối trước")
            return None
            
        latest_file = max(balance_files, key=os.path.getctime)
        print(f"🔍 Đang kiểm tra file: {latest_file}")
        
        try:
            validator = FormulaValidator(latest_file)
            result = validator.run_complete_validation()
            
            self.validation_results.append({
                'filename': latest_file,
                'timestamp': datetime.datetime.now().isoformat(),
                'overall_status': validator.validation_results['overall_status'],
                'passed': result
            })
            
            if result:
                print(f"\n🎉 VALIDATION THÀNH CÔNG!")
                print(f"📊 Trạng thái: {validator.validation_results['overall_status'].upper()}")
            else:
                print(f"\n⚠️  VALIDATION CẦN CẢI THIỆN!")
                print(f"📊 Trạng thái: {validator.validation_results['overall_status'].upper()}")
                
            return result
            
        except Exception as e:
            print(f"❌ Lỗi validation: {e}")
            return False
            
    def option_4_multi_period(self):
        """Tùy chọn 4: Phân tích nhiều kỳ"""
        print("\n📈 ĐANG TẠO HỆ THỐNG PHÂN TÍCH NHIỀU KỲ...")
        print("-" * 60)
        
        try:
            analyzer = MultiPeriodAnalyzer(periods=3)
            filename = analyzer.generate_complete_analysis()
            
            if filename:
                self.generated_files.append({
                    'type': 'multi_period',
                    'filename': filename,
                    'timestamp': datetime.datetime.now().isoformat(),
                    'sheets': len(analyzer.sheets),
                    'periods': len(analyzer.periods_list)
                })
                
                print(f"\n✅ THÀNH CÔNG!")
                print(f"📁 File đã tạo: {filename}")
                print(f"📊 Sheets: {len(analyzer.sheets)}")
                print(f"📈 Periods: {len(analyzer.periods_list)}")
                return filename
            else:
                print("❌ Lỗi tạo hệ thống phân tích nhiều kỳ")
                return None
                
        except Exception as e:
            print(f"❌ Lỗi: {e}")
            return None
            
    def option_5_complete_system(self):
        """Tùy chọn 5: Tạo toàn bộ hệ thống tự động"""
        print("\n🚀 ĐANG TẠO TOÀN BỘ HỆ THỐNG TỰ ĐỘNG...")
        print("=" * 80)
        
        success_count = 0
        total_steps = 4
        
        # Bước 1: Tạo bảng cân đối
        print("\n📋 BƯỚC 1/4: Tạo Bảng Cân Đối Kế Toán")
        balance_file = self.option_1_balance_sheet()
        if balance_file:
            success_count += 1
            
        # Bước 2: Tạo phân tích tài chính
        print("\n📋 BƯỚC 2/4: Tạo Hệ thống Phân tích Tài chính")
        analysis_file = self.option_2_financial_analysis()
        if analysis_file:
            success_count += 1
            
        # Bước 3: Validation
        print("\n📋 BƯỚC 3/4: Kiểm tra và Validation")
        validation_result = self.option_3_validation()
        if validation_result:
            success_count += 1
            
        # Bước 4: Phân tích nhiều kỳ
        print("\n📋 BƯỚC 4/4: Tạo Phân tích Nhiều kỳ")
        multi_period_file = self.option_4_multi_period()
        if multi_period_file:
            success_count += 1
            
        # Tạo báo cáo tổng kết
        self.generate_summary_report()
        
        # Kết quả
        print("\n" + "=" * 80)
        print("🎉 KẾT QUẢ TẠO HỆ THỐNG HOÀN CHỈNH")
        print("=" * 80)
        print(f"✅ Hoàn thành: {success_count}/{total_steps} bước")
        
        if success_count == total_steps:
            print("🎊 THÀNH CÔNG HOÀN TOÀN!")
            print("🚀 Hệ thống báo cáo tài chính động đã sẵn sàng sử dụng!")
        elif success_count >= total_steps * 0.75:
            print("⚠️  THÀNH CÔNG PHẦN LỚN - Một số tính năng có thể cần xem lại")
        else:
            print("❌ CẦN KIỂM TRA LẠI - Nhiều bước gặp lỗi")
            
        self.show_file_summary()
        return success_count == total_steps
        
    def option_6_file_list(self):
        """Tùy chọn 6: Xem danh sách file đã tạo"""
        print("\n📁 DANH SÁCH FILE ĐÃ TẠO:")
        print("-" * 60)
        
        if not self.generated_files:
            print("📭 Chưa có file nào được tạo")
            print("💡 Sử dụng các tùy chọn 1-5 để tạo file")
            return
            
        for i, file_info in enumerate(self.generated_files, 1):
            print(f"\n📄 File {i}:")
            print(f"   📁 Tên: {file_info['filename']}")
            print(f"   🏷️  Loại: {file_info['type']}")
            print(f"   🕐 Thời gian: {file_info['timestamp']}")
            
            if file_info['type'] == 'balance_sheet':
                print(f"   🏷️  Named Ranges: {file_info.get('named_ranges', 'N/A')}")
            elif file_info['type'] == 'financial_analysis':
                print(f"   📊 Sheets: {file_info.get('sheets', 'N/A')}")
                print(f"   🔢 Formulas: {file_info.get('formulas', 'N/A')}")
            elif file_info['type'] == 'multi_period':
                print(f"   📊 Sheets: {file_info.get('sheets', 'N/A')}")
                print(f"   📈 Periods: {file_info.get('periods', 'N/A')}")
                
        # Hiển thị validation results
        if self.validation_results:
            print(f"\n🔍 KẾT QUẢ VALIDATION:")
            for result in self.validation_results:
                status_icon = "✅" if result['passed'] else "⚠️"
                print(f"   {status_icon} {result['filename']}: {result['overall_status'].upper()}")
                
    def option_7_help(self):
        """Tùy chọn 7: Hướng dẫn sử dụng"""
        print("\n📖 HƯỚNG DẪN SỬ DỤNG HỆ THỐNG")
        print("=" * 80)
        
        help_content = """
🎯 TỔNG QUAN:
Hệ thống tạo báo cáo tài chính với công thức Excel động, tuân thủ chuẩn kế toán Việt Nam.

🚀 CÁCH SỬ DỤNG NHANH:
1. Chọn tùy chọn 5 để tạo toàn bộ hệ thống tự động
2. Mở các file Excel được tạo
3. Cập nhật dữ liệu trong bảng cân đối kế toán
4. Tất cả báo cáo sẽ tự động cập nhật theo dữ liệu mới

📊 CÁC THÀNH PHẦN CHÍNH:

🏗️  1. BẢNG CÂN ĐỐI KẾ TOÁN:
   - Cấu trúc theo chuẩn VAS/Circular 200/2014/TT-BTC
   - 40+ named ranges tự động
   - Kiểm tra phương trình cân đối (Assets = Liabilities + Equity)
   - Sheet mapping với mã kế toán Việt Nam

📈 2. HỆ THỐNG PHÂN TÍCH TÀI CHÍNH:
   - 5 báo cáo chuyên sâu + Dashboard
   - Tất cả công thức Excel tham chiếu động
   - Đánh giá tự động theo tiêu chuẩn ngành
   - Phân tích: Thanh khoản, Sinh lời, Hiệu quả, Cơ cấu tài chính

🔍 3. VALIDATION HỆ THỐNG:
   - Kiểm tra phương trình cân đối
   - Xác thực named ranges
   - Phát hiện lỗi công thức Excel
   - Báo cáo chi tiết và khuyến nghị

📊 4. PHÂN TÍCH NHIỀU KỲ:
   - So sánh 3+ kỳ báo cáo
   - Phân tích xu hướng tăng trưởng
   - Dự báo tài chính tự động
   - Biểu đồ và visualization

💡 MẸO SỬ DỤNG:
✓ Luôn backup file trước khi thay đổi dữ liệu
✓ Sử dụng validation trước khi phân tích
✓ Cập nhật dữ liệu ở sheet gốc, các báo cáo sẽ tự động cập nhật
✓ Xem sheet "Mapping và Công thức" để hiểu cách hoạt động

🔧 YÊU CẦU HỆ THỐNG:
- Excel 2016+ (khuyến nghị Excel 365)
- Python 3.8+ với các thư viện: openpyxl, pandas
- Windows/Mac/Linux

📞 HỖ TRỢ:
- Xem file README.md để biết thêm chi tiết
- Kiểm tra các file JSON để xem thông tin named ranges
- Sử dụng chức năng validation để chẩn đoán vấn đề
"""
        print(help_content)
        
        print("\n" + "=" * 80)
        input("Nhấn Enter để tiếp tục...")
        
    def generate_summary_report(self):
        """Tạo báo cáo tổng kết"""
        try:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            report_filename = f"system_summary_report_{timestamp}.json"
            
            summary_data = {
                'system_info': self.system_info,
                'generation_time': datetime.datetime.now().isoformat(),
                'generated_files': self.generated_files,
                'validation_results': self.validation_results,
                'statistics': {
                    'total_files': len(self.generated_files),
                    'successful_validations': len([r for r in self.validation_results if r['passed']]),
                    'total_validations': len(self.validation_results)
                }
            }
            
            with open(report_filename, 'w', encoding='utf-8') as f:
                json.dump(summary_data, f, ensure_ascii=False, indent=2)
                
            print(f"✅ Đã tạo báo cáo tổng kết: {report_filename}")
            
        except Exception as e:
            print(f"⚠️  Lỗi tạo báo cáo tổng kết: {e}")
            
    def show_file_summary(self):
        """Hiển thị tóm tắt file đã tạo"""
        print(f"\n📋 TÓM TẮT FILE ĐÃ TẠO:")
        print("-" * 40)
        
        file_types = {}
        for file_info in self.generated_files:
            file_type = file_info['type']
            if file_type not in file_types:
                file_types[file_type] = []
            file_types[file_type].append(file_info['filename'])
            
        for file_type, files in file_types.items():
            icon_map = {
                'balance_sheet': '🏗️',
                'financial_analysis': '📊',
                'multi_period': '📈'
            }
            icon = icon_map.get(file_type, '📄')
            print(f"{icon} {file_type}: {len(files)} file(s)")
            for filename in files:
                print(f"   - {filename}")
                
    def run(self):
        """Chạy hệ thống chính"""
        self.create_output_folder()
        self.print_banner()
        
        while True:
            self.print_menu()
            
            try:
                choice = input("👉 Chọn tùy chọn (1-8): ").strip()
                
                if choice == '1':
                    self.option_1_balance_sheet()
                elif choice == '2':
                    self.option_2_financial_analysis()
                elif choice == '3':
                    self.option_3_validation()
                elif choice == '4':
                    self.option_4_multi_period()
                elif choice == '5':
                    self.option_5_complete_system()
                elif choice == '6':
                    self.option_6_file_list()
                elif choice == '7':
                    self.option_7_help()
                elif choice == '8':
                    print("\n👋 Cảm ơn bạn đã sử dụng hệ thống!")
                    print("🚀 Chúc bạn phân tích tài chính hiệu quả!")
                    break
                else:
                    print("❌ Tùy chọn không hợp lệ. Vui lòng chọn từ 1-8.")
                    
            except KeyboardInterrupt:
                print("\n\n👋 Đã dừng hệ thống. Tạm biệt!")
                break
            except Exception as e:
                print(f"❌ Lỗi: {e}")
                
            input("\nNhấn Enter để tiếp tục...")

if __name__ == "__main__":
    system = MainIntegrationSystem()
    system.run()