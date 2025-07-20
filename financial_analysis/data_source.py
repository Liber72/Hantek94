"""
Mô-đun cung cấp dữ liệu mẫu cho bảng cân đối kế toán
Data source module for sample balance sheet data

Dữ liệu mẫu dựa trên cấu trúc kế toán Việt Nam và quốc tế
Sample data based on Vietnamese and international accounting standards

Nguồn tham khảo: 
- Thông tư 200/2014/TT-BTC về chế độ kế toán doanh nghiệp
- Các báo cáo tài chính công khai của doanh nghiệp niêm yết
"""

from datetime import datetime
import pandas as pd

class FinancialDataSource:
    """Lớp cung cấp dữ liệu mẫu cho bảng cân đối kế toán"""
    
    def __init__(self):
        """Khởi tạo nguồn dữ liệu với thông tin công ty mẫu"""
        self.company_info = {
            'name': 'CÔNG TY CỔ PHẦN MẪU XYZ',
            'period': '31/12/2023',
            'unit': 'VND (triệu đồng)',
            'currency': 'VND',
            'data_source': 'Dữ liệu mẫu dựa trên cấu trúc kế toán chuẩn Việt Nam'
        }
    
    def get_balance_sheet_data(self):
        """
        Lấy dữ liệu bảng cân đối kế toán mẫu
        Returns sample balance sheet data following Vietnamese accounting standards
        
        Returns:
            dict: Dữ liệu bảng cân đối kế toán với cấu trúc chuẩn
        """
        
        # Dữ liệu mẫu cho TÀI SẢN (ASSETS)
        assets_data = {
            # A. TÀI SẢN NGẮN HẠN (Current Assets)
            'A_TAI_SAN_NGAN_HAN': {
                'description': 'A. TÀI SẢN NGẮN HẠN',
                'items': {
                    '111': {'name': 'Tiền và các khoản tương đương tiền', 'value': 15_000},
                    '112': {'name': 'Tiền gửi có kỳ hạn', 'value': 5_000},
                    '121': {'name': 'Đầu tư tài chính ngắn hạn', 'value': 8_000},
                    '131': {'name': 'Phải thu khách hàng', 'value': 25_000},
                    '132': {'name': 'Trả trước cho người bán', 'value': 3_000},
                    '133': {'name': 'Phải thu nội bộ ngắn hạn', 'value': 2_000},
                    '135': {'name': 'Phải thu khác', 'value': 4_000},
                    '141': {'name': 'Hàng tồn kho', 'value': 35_000},
                    '151': {'name': 'Tài sản ngắn hạn khác', 'value': 3_000}
                }
            },
            
            # B. TÀI SẢN DÀI HẠN (Non-current Assets)
            'B_TAI_SAN_DAI_HAN': {
                'description': 'B. TÀI SẢN DÀI HẠN',
                'items': {
                    '211': {'name': 'Phải thu dài hạn của khách hàng', 'value': 5_000},
                    '213': {'name': 'Tài sản thuế thu nhập hoãn lại', 'value': 2_000},
                    '221': {'name': 'Đầu tư tài chính dài hạn', 'value': 12_000},
                    '211_tangible': {'name': 'Tài sản cố định hữu hình', 'value': 180_000},
                    '213_depreciation': {'name': 'Hao mòn lũy kế TSCĐ hữu hình', 'value': -45_000},
                    '217': {'name': 'Tài sản cố định vô hình', 'value': 8_000},
                    '241': {'name': 'Tài sản dở dang dài hạn', 'value': 15_000},
                    '261': {'name': 'Tài sản dài hạn khác', 'value': 3_000}
                }
            }
        }
        
        # Dữ liệu mẫu cho NGUỒN VỐN (LIABILITIES & EQUITY)
        liabilities_equity_data = {
            # C. NỢ PHẢI TRẢ (Liabilities)
            'C_NO_PHAI_TRA': {
                'description': 'C. NỢ PHẢI TRẢ',
                'items': {
                    # Nợ ngắn hạn
                    '311': {'name': 'Phải trả người bán', 'value': 18_000},
                    '312': {'name': 'Người mua trả tiền trước', 'value': 5_000},
                    '313': {'name': 'Thuế và các khoản phải nộp', 'value': 7_000},
                    '314': {'name': 'Phải trả người lao động', 'value': 8_000},
                    '319': {'name': 'Phải trả ngắn hạn khác', 'value': 4_000},
                    '323': {'name': 'Vay và nợ thuê tài chính ngắn hạn', 'value': 15_000},
                    '327': {'name': 'Dự phòng phải trả ngắn hạn', 'value': 3_000},
                    
                    # Nợ dài hạn
                    '337': {'name': 'Vay và nợ thuê tài chính dài hạn', 'value': 60_000},
                    '341': {'name': 'Thuế thu nhập hoãn lại phải trả', 'value': 4_000},
                    '347': {'name': 'Dự phòng phải trả dài hạn', 'value': 6_000},
                    '349': {'name': 'Vay nội bộ dài hạn', 'value': 5_000},
                    '353': {'name': 'Phải trả dài hạn khác', 'value': 5_000}
                }
            },
            
            # D. VỐN CHỦ SỞ HỮU (Owner's Equity)
            'D_VON_CHU_SO_HUU': {
                'description': 'D. VỐN CHỦ SỞ HỮU',
                'items': {
                    '411': {'name': 'Vốn đầu tư của chủ sở hữu', 'value': 100_000},
                    '412': {'name': 'Thặng dư vốn cổ phần', 'value': 15_000},
                    '421': {'name': 'Lợi nhuận sau thuế chưa phân phối', 'value': 15_000},
                    '422': {'name': 'Quỹ đầu tư phát triển', 'value': 8_000},
                    '429': {'name': 'Quỹ khác thuộc vốn chủ sở hữu', 'value': 2_000}
                }
            }
        }
        
        return {
            'company_info': self.company_info,
            'assets': assets_data,
            'liabilities_equity': liabilities_equity_data
        }
    
    def get_income_statement_data(self):
        """
        Lấy dữ liệu báo cáo kết quả kinh doanh mẫu để tính toán các chỉ số
        Returns sample income statement data for ratio calculations
        """
        return {
            'revenue': 280_000,  # Doanh thu thuần
            'cost_of_goods_sold': 168_000,  # Giá vốn hàng bán
            'gross_profit': 112_000,  # Lợi nhuận gộp
            'operating_expenses': 75_000,  # Chi phí hoạt động
            'operating_income': 37_000,  # Lợi nhuận từ hoạt động kinh doanh
            'financial_expenses': 8_000,  # Chi phí tài chính
            'other_income': 2_000,  # Thu nhập khác
            'profit_before_tax': 31_000,  # Lợi nhuận trước thuế
            'tax_expense': 7_750,  # Chi phí thuế thu nhập doanh nghiệp
            'net_income': 23_250  # Lợi nhuận sau thuế
        }
    
    def calculate_totals(self, data_dict):
        """
        Tính tổng cho các nhóm tài sản và nguồn vốn
        Calculate totals for asset and liability groups
        """
        totals = {}
        
        for section_key, section_data in data_dict.items():
            if 'items' in section_data:
                total = sum(item['value'] for item in section_data['items'].values())
                totals[section_key] = total
        
        return totals
    
    def get_data_sources_info(self):
        """Trả về thông tin về nguồn dữ liệu được sử dụng"""
        return {
            'primary_source': 'Dữ liệu mẫu dựa trên cấu trúc kế toán Việt Nam',
            'standards_reference': [
                'Thông tư 200/2014/TT-BTC về chế độ kế toán doanh nghiệp',
                'Chuẩn mực kế toán Việt Nam (VAS)',
                'Chuẩn mực báo cáo tài chính quốc tế (IFRS) - tham khảo'
            ],
            'data_characteristics': [
                'Số liệu được làm tròn đến triệu đồng',
                'Cấu trúc tuân thủ mẫu bảng cân đối kế toán theo TT 200/2014/TT-BTC',
                'Dữ liệu mang tính chất minh họa, phù hợp cho mục đích học tập và phân tích'
            ],
            'last_updated': datetime.now().strftime('%d/%m/%Y'),
            'disclaimer': 'Dữ liệu này chỉ mang tính chất minh họa và không phản ánh tình hình tài chính thực tế của bất kỳ doanh nghiệp cụ thể nào.'
        }

# Hàm tiện ích để lấy dữ liệu nhanh
def get_sample_data():
    """Hàm tiện ích để lấy dữ liệu mẫu nhanh"""
    source = FinancialDataSource()
    return source.get_balance_sheet_data()

if __name__ == "__main__":
    # Test chức năng
    source = FinancialDataSource()
    data = source.get_balance_sheet_data()
    
    print("=== THÔNG TIN CÔNG TY ===")
    for key, value in data['company_info'].items():
        print(f"{key}: {value}")
    
    print("\n=== TỔNG QUAN DỮ LIỆU ===")
    print("Số lượng khoản mục tài sản:", 
          sum(len(section['items']) for section in data['assets'].values()))
    print("Số lượng khoản mục nợ và vốn:", 
          sum(len(section['items']) for section in data['liabilities_equity'].values()))