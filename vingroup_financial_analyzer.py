"""
VinGroup Financial Analysis Excel Generator
==========================================

This script generates a comprehensive Excel file for financial analysis
of VinGroup (VIC) for educational purposes.

Requirements:
- openpyxl (for Excel file generation)
- pandas (for data manipulation)
- matplotlib (for charts)

Author: Financial Analysis Tool
Date: 2024
"""

import csv
import json
import os
from datetime import datetime
from typing import Dict, List, Any

# Financial data for VinGroup (VIC) - Based on realistic financial structure
VINGROUP_DATA = {
    "company_info": {
        "name": "Tập đoàn VinGroup",
        "ticker": "VIC",
        "industry": "Đa ngành",
        "currency": "VND",
        "unit": "Tỷ VND"
    },
    "balance_sheet": {
        "2023": {
            "assets": {
                "current_assets": {
                    "cash_and_equivalents": 45000,
                    "short_term_investments": 12000,
                    "accounts_receivable": 8500,
                    "inventory": 85000,
                    "prepaid_expenses": 3500,
                    "other_current_assets": 2000
                },
                "non_current_assets": {
                    "long_term_investments": 25000,
                    "property_plant_equipment": 180000,
                    "intangible_assets": 15000,
                    "goodwill": 5000,
                    "other_non_current_assets": 8000
                }
            },
            "liabilities": {
                "current_liabilities": {
                    "short_term_debt": 35000,
                    "accounts_payable": 12000,
                    "accrued_expenses": 8000,
                    "other_current_liabilities": 5000
                },
                "non_current_liabilities": {
                    "long_term_debt": 120000,
                    "deferred_tax_liabilities": 8000,
                    "other_non_current_liabilities": 10000
                }
            },
            "equity": {
                "share_capital": 50000,
                "retained_earnings": 95000,
                "other_equity": 5500
            }
        },
        "2024": {
            "assets": {
                "current_assets": {
                    "cash_and_equivalents": 52000,
                    "short_term_investments": 15000,
                    "accounts_receivable": 9200,
                    "inventory": 92000,
                    "prepaid_expenses": 4000,
                    "other_current_assets": 2500
                },
                "non_current_assets": {
                    "long_term_investments": 28000,
                    "property_plant_equipment": 195000,
                    "intangible_assets": 18000,
                    "goodwill": 5000,
                    "other_non_current_assets": 9000
                }
            },
            "liabilities": {
                "current_liabilities": {
                    "short_term_debt": 38000,
                    "accounts_payable": 13500,
                    "accrued_expenses": 8800,
                    "other_current_liabilities": 5500
                },
                "non_current_liabilities": {
                    "long_term_debt": 125000,
                    "deferred_tax_liabilities": 9000,
                    "other_non_current_liabilities": 11000
                }
            },
            "equity": {
                "share_capital": 55000,
                "retained_earnings": 108000,
                "other_equity": 6200
            }
        }
    },
    "income_statement": {
        "2023": {
            "revenue": 156000,
            "cost_of_goods_sold": 98000,
            "gross_profit": 58000,
            "operating_expenses": {
                "selling_expenses": 15000,
                "administrative_expenses": 12000,
                "research_development": 3000
            },
            "operating_profit": 28000,
            "financial_income": 2000,
            "financial_expenses": 8000,
            "profit_before_tax": 22000,
            "tax_expense": 4400,
            "net_profit": 17600
        },
        "2024": {
            "revenue": 172000,
            "cost_of_goods_sold": 108000,
            "gross_profit": 64000,
            "operating_expenses": {
                "selling_expenses": 16500,
                "administrative_expenses": 13200,
                "research_development": 3500
            },
            "operating_profit": 30800,
            "financial_income": 2500,
            "financial_expenses": 8800,
            "profit_before_tax": 24500,
            "tax_expense": 4900,
            "net_profit": 19600
        }
    },
    "cash_flow": {
        "2023": {
            "operating_cash_flow": 25000,
            "investing_cash_flow": -18000,
            "financing_cash_flow": -5000,
            "net_cash_flow": 2000,
            "beginning_cash": 43000,
            "ending_cash": 45000
        },
        "2024": {
            "operating_cash_flow": 28000,
            "investing_cash_flow": -22000,
            "financing_cash_flow": 1000,
            "net_cash_flow": 7000,
            "beginning_cash": 45000,
            "ending_cash": 52000
        }
    }
}

class VinGroupFinancialAnalyzer:
    """
    A comprehensive financial analyzer for VinGroup data
    """
    
    def __init__(self, data: Dict[str, Any]):
        self.data = data
        self.company_info = data["company_info"]
        self.balance_sheet = data["balance_sheet"]
        self.income_statement = data["income_statement"]
        self.cash_flow = data["cash_flow"]
    
    def calculate_totals(self, year: str) -> Dict[str, float]:
        """Calculate total assets, liabilities, and equity for a given year"""
        bs = self.balance_sheet[year]
        
        # Calculate current assets total
        current_assets_total = sum(bs["assets"]["current_assets"].values())
        
        # Calculate non-current assets total
        non_current_assets_total = sum(bs["assets"]["non_current_assets"].values())
        
        # Calculate total assets
        total_assets = current_assets_total + non_current_assets_total
        
        # Calculate current liabilities total
        current_liabilities_total = sum(bs["liabilities"]["current_liabilities"].values())
        
        # Calculate non-current liabilities total
        non_current_liabilities_total = sum(bs["liabilities"]["non_current_liabilities"].values())
        
        # Calculate total liabilities
        total_liabilities = current_liabilities_total + non_current_liabilities_total
        
        # Calculate total equity
        total_equity = sum(bs["equity"].values())
        
        return {
            "current_assets_total": current_assets_total,
            "non_current_assets_total": non_current_assets_total,
            "total_assets": total_assets,
            "current_liabilities_total": current_liabilities_total,
            "non_current_liabilities_total": non_current_liabilities_total,
            "total_liabilities": total_liabilities,
            "total_equity": total_equity
        }
    
    def calculate_financial_ratios(self, year: str) -> Dict[str, float]:
        """Calculate financial ratios for a given year"""
        bs = self.balance_sheet[year]
        inc = self.income_statement[year]
        totals = self.calculate_totals(year)
        
        # Liquidity ratios
        current_ratio = totals["current_assets_total"] / totals["current_liabilities_total"]
        
        # Quick assets (current assets - inventory - prepaid expenses)
        quick_assets = (totals["current_assets_total"] - 
                       bs["assets"]["current_assets"]["inventory"] - 
                       bs["assets"]["current_assets"]["prepaid_expenses"])
        quick_ratio = quick_assets / totals["current_liabilities_total"]
        
        # Cash ratio
        cash_ratio = (bs["assets"]["current_assets"]["cash_and_equivalents"] + 
                     bs["assets"]["current_assets"]["short_term_investments"]) / totals["current_liabilities_total"]
        
        # Profitability ratios
        net_profit_margin = (inc["net_profit"] / inc["revenue"]) * 100
        gross_profit_margin = (inc["gross_profit"] / inc["revenue"]) * 100
        roa = (inc["net_profit"] / totals["total_assets"]) * 100
        roe = (inc["net_profit"] / totals["total_equity"]) * 100
        
        # Efficiency ratios
        asset_turnover = inc["revenue"] / totals["total_assets"]
        inventory_turnover = inc["cost_of_goods_sold"] / bs["assets"]["current_assets"]["inventory"]
        
        # Leverage ratios
        debt_to_equity = totals["total_liabilities"] / totals["total_equity"]
        debt_to_assets = totals["total_liabilities"] / totals["total_assets"]
        
        return {
            "current_ratio": current_ratio,
            "quick_ratio": quick_ratio,
            "cash_ratio": cash_ratio,
            "net_profit_margin": net_profit_margin,
            "gross_profit_margin": gross_profit_margin,
            "roa": roa,
            "roe": roe,
            "asset_turnover": asset_turnover,
            "inventory_turnover": inventory_turnover,
            "debt_to_equity": debt_to_equity,
            "debt_to_assets": debt_to_assets
        }
    
    def generate_csv_reports(self, output_dir: str = "."):
        """Generate CSV files for each financial statement"""
        os.makedirs(output_dir, exist_ok=True)
        
        # Balance Sheet CSV
        self._generate_balance_sheet_csv(os.path.join(output_dir, "balance_sheet.csv"))
        
        # Income Statement CSV
        self._generate_income_statement_csv(os.path.join(output_dir, "income_statement.csv"))
        
        # Cash Flow CSV
        self._generate_cash_flow_csv(os.path.join(output_dir, "cash_flow.csv"))
        
        # Financial Ratios CSV
        self._generate_ratios_csv(os.path.join(output_dir, "financial_ratios.csv"))
        
        # Guidelines and Exercises CSV
        self._generate_guidelines_csv(os.path.join(output_dir, "guidelines_exercises.csv"))
    
    def _generate_balance_sheet_csv(self, filename: str):
        """Generate balance sheet CSV"""
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            
            # Header
            writer.writerow(["BẢNG CÂN ĐỐI KẾ TOÁN - VINGROUP", "", ""])
            writer.writerow(["Đơn vị: Tỷ VND", "", ""])
            writer.writerow(["", "", ""])
            writer.writerow(["Khoản mục", "2023", "2024"])
            writer.writerow(["", "", ""])
            
            # Assets
            writer.writerow(["TÀI SẢN", "", ""])
            writer.writerow(["A. TÀI SẢN NGẮN HẠN", "", ""])
            
            for year in ["2023", "2024"]:
                if year == "2023":
                    writer.writerow(["I. Tiền và tương đương tiền", 
                                   self.balance_sheet[year]["assets"]["current_assets"]["cash_and_equivalents"], 
                                   self.balance_sheet["2024"]["assets"]["current_assets"]["cash_and_equivalents"]])
                    break
            
            for item, key in [
                ("I. Tiền và tương đương tiền", "cash_and_equivalents"),
                ("II. Đầu tư tài chính ngắn hạn", "short_term_investments"),
                ("III. Phải thu ngắn hạn", "accounts_receivable"),
                ("IV. Hàng tồn kho", "inventory"),
                ("V. Tài sản ngắn hạn khác", "prepaid_expenses")
            ]:
                writer.writerow([item, 
                               self.balance_sheet["2023"]["assets"]["current_assets"][key],
                               self.balance_sheet["2024"]["assets"]["current_assets"][key]])
            
            # Calculate totals
            totals_2023 = self.calculate_totals("2023")
            totals_2024 = self.calculate_totals("2024")
            
            writer.writerow(["TỔNG TÀI SẢN NGẮN HẠN", 
                           totals_2023["current_assets_total"],
                           totals_2024["current_assets_total"]])
            
            writer.writerow(["", "", ""])
            writer.writerow(["B. TÀI SẢN DÀI HẠN", "", ""])
            
            for item, key in [
                ("I. Đầu tư tài chính dài hạn", "long_term_investments"),
                ("II. Tài sản cố định", "property_plant_equipment"),
                ("III. Tài sản vô hình", "intangible_assets"),
                ("IV. Lợi thế thương mại", "goodwill")
            ]:
                writer.writerow([item, 
                               self.balance_sheet["2023"]["assets"]["non_current_assets"][key],
                               self.balance_sheet["2024"]["assets"]["non_current_assets"][key]])
            
            writer.writerow(["TỔNG TÀI SẢN DÀI HẠN", 
                           totals_2023["non_current_assets_total"],
                           totals_2024["non_current_assets_total"]])
            
            writer.writerow(["", "", ""])
            writer.writerow(["TỔNG CỘNG TÀI SẢN", 
                           totals_2023["total_assets"],
                           totals_2024["total_assets"]])
    
    def _generate_income_statement_csv(self, filename: str):
        """Generate income statement CSV"""
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            
            # Header
            writer.writerow(["BÁO CÁO KẾT QUẢ KINH DOANH - VINGROUP", "", ""])
            writer.writerow(["Đơn vị: Tỷ VND", "", ""])
            writer.writerow(["", "", ""])
            writer.writerow(["Khoản mục", "2023", "2024"])
            writer.writerow(["", "", ""])
            
            # Income statement items
            for item, key in [
                ("1. Doanh thu bán hàng", "revenue"),
                ("2. Giá vốn hàng bán", "cost_of_goods_sold"),
                ("3. Lợi nhuận gộp", "gross_profit"),
                ("4. Chi phí bán hàng", None),
                ("5. Chi phí quản lý", None),
                ("6. Chi phí R&D", None),
                ("7. Lợi nhuận từ hoạt động kinh doanh", "operating_profit"),
                ("8. Thu nhập tài chính", "financial_income"),
                ("9. Chi phí tài chính", "financial_expenses"),
                ("10. Lợi nhuận trước thuế", "profit_before_tax"),
                ("11. Chi phí thuế", "tax_expense"),
                ("12. Lợi nhuận sau thuế", "net_profit")
            ]:
                if key is None:
                    # Handle operating expenses
                    if "bán hàng" in item:
                        writer.writerow([item, 
                                       self.income_statement["2023"]["operating_expenses"]["selling_expenses"],
                                       self.income_statement["2024"]["operating_expenses"]["selling_expenses"]])
                    elif "quản lý" in item:
                        writer.writerow([item, 
                                       self.income_statement["2023"]["operating_expenses"]["administrative_expenses"],
                                       self.income_statement["2024"]["operating_expenses"]["administrative_expenses"]])
                    elif "R&D" in item:
                        writer.writerow([item, 
                                       self.income_statement["2023"]["operating_expenses"]["research_development"],
                                       self.income_statement["2024"]["operating_expenses"]["research_development"]])
                else:
                    writer.writerow([item, 
                                   self.income_statement["2023"][key],
                                   self.income_statement["2024"][key]])
    
    def _generate_cash_flow_csv(self, filename: str):
        """Generate cash flow CSV"""
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            
            # Header
            writer.writerow(["BÁO CÁO LƯU CHUYỂN TIỀN TỆ - VINGROUP", "", ""])
            writer.writerow(["Đơn vị: Tỷ VND", "", ""])
            writer.writerow(["", "", ""])
            writer.writerow(["Khoản mục", "2023", "2024"])
            writer.writerow(["", "", ""])
            
            # Cash flow items
            for item, key in [
                ("1. Lưu chuyển tiền từ hoạt động kinh doanh", "operating_cash_flow"),
                ("2. Lưu chuyển tiền từ hoạt động đầu tư", "investing_cash_flow"),
                ("3. Lưu chuyển tiền từ hoạt động tài chính", "financing_cash_flow"),
                ("4. Lưu chuyển tiền thuần trong kỳ", "net_cash_flow"),
                ("5. Tiền đầu kỳ", "beginning_cash"),
                ("6. Tiền cuối kỳ", "ending_cash")
            ]:
                writer.writerow([item, 
                               self.cash_flow["2023"][key],
                               self.cash_flow["2024"][key]])
    
    def _generate_ratios_csv(self, filename: str):
        """Generate financial ratios CSV"""
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            
            # Header
            writer.writerow(["PHÂN TÍCH CHỈ SỐ TÀI CHÍNH - VINGROUP", "", "", ""])
            writer.writerow(["", "", "", ""])
            writer.writerow(["Chỉ số", "2023", "2024", "Ý nghĩa"])
            writer.writerow(["", "", "", ""])
            
            # Calculate ratios
            ratios_2023 = self.calculate_financial_ratios("2023")
            ratios_2024 = self.calculate_financial_ratios("2024")
            
            # Ratio data with meanings
            ratio_data = [
                ("CHỈ SỐ THANH KHOẢN", "", "", ""),
                ("Hệ số thanh khoản hiện hành", f"{ratios_2023['current_ratio']:.2f}", f"{ratios_2024['current_ratio']:.2f}", "Khả năng thanh toán nợ ngắn hạn"),
                ("Hệ số thanh khoản nhanh", f"{ratios_2023['quick_ratio']:.2f}", f"{ratios_2024['quick_ratio']:.2f}", "Khả năng thanh toán nhanh"),
                ("Hệ số thanh khoản tuyệt đối", f"{ratios_2023['cash_ratio']:.2f}", f"{ratios_2024['cash_ratio']:.2f}", "Khả năng thanh toán bằng tiền mặt"),
                ("", "", "", ""),
                ("CHỈ SỐ SINH LỜI", "", "", ""),
                ("Biên lợi nhuận gộp (%)", f"{ratios_2023['gross_profit_margin']:.2f}%", f"{ratios_2024['gross_profit_margin']:.2f}%", "Hiệu quả kiểm soát chi phí"),
                ("Biên lợi nhuận ròng (%)", f"{ratios_2023['net_profit_margin']:.2f}%", f"{ratios_2024['net_profit_margin']:.2f}%", "Hiệu quả kinh doanh tổng thể"),
                ("ROA (%)", f"{ratios_2023['roa']:.2f}%", f"{ratios_2024['roa']:.2f}%", "Hiệu quả sử dụng tài sản"),
                ("ROE (%)", f"{ratios_2023['roe']:.2f}%", f"{ratios_2024['roe']:.2f}%", "Hiệu quả sử dụng vốn chủ sở hữu"),
                ("", "", "", ""),
                ("CHỈ SỐ HIỆU QUẢ", "", "", ""),
                ("Vòng quay tài sản", f"{ratios_2023['asset_turnover']:.2f}", f"{ratios_2024['asset_turnover']:.2f}", "Hiệu quả sử dụng tài sản"),
                ("Vòng quay hàng tồn kho", f"{ratios_2023['inventory_turnover']:.2f}", f"{ratios_2024['inventory_turnover']:.2f}", "Hiệu quả quản lý hàng tồn kho"),
                ("", "", "", ""),
                ("CHỈ SỐ CƠ CẤU TÀI CHÍNH", "", "", ""),
                ("Nợ/Vốn chủ sở hữu", f"{ratios_2023['debt_to_equity']:.2f}", f"{ratios_2024['debt_to_equity']:.2f}", "Cơ cấu nợ và vốn"),
                ("Nợ/Tổng tài sản", f"{ratios_2023['debt_to_assets']:.2f}", f"{ratios_2024['debt_to_assets']:.2f}", "Mức độ sử dụng nợ")
            ]
            
            for row in ratio_data:
                writer.writerow(row)
    
    def _generate_guidelines_csv(self, filename: str):
        """Generate guidelines and exercises CSV"""
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            
            # Header
            writer.writerow(["HƯỚNG DẪN VÀ BÀI TẬP PHÂN TÍCH TÀI CHÍNH", ""])
            writer.writerow(["", ""])
            
            # Guidelines
            writer.writerow(["A. HƯỚNG DẪN PHÂN TÍCH", ""])
            writer.writerow(["", ""])
            writer.writerow(["1. Cách đọc báo cáo tài chính:", ""])
            writer.writerow(["- Bảng cân đối kế toán: Phản ánh tình hình tài chính tại thời điểm cụ thể", ""])
            writer.writerow(["- Báo cáo KQKD: Phản ánh kết quả hoạt động kinh doanh trong kỳ", ""])
            writer.writerow(["- Báo cáo lưu chuyển tiền tệ: Phản ánh luồng tiền vào/ra", ""])
            writer.writerow(["", ""])
            writer.writerow(["2. Các bước phân tích:", ""])
            writer.writerow(["Bước 1: Phân tích cơ cấu tài sản và nguồn vốn", ""])
            writer.writerow(["Bước 2: Tính toán các chỉ số tài chính", ""])
            writer.writerow(["Bước 3: So sánh với năm trước và ngành", ""])
            writer.writerow(["Bước 4: Đánh giá xu hướng và đưa ra nhận xét", ""])
            writer.writerow(["", ""])
            
            # Exercises
            writer.writerow(["B. BÀI TẬP THỰC HÀNH", ""])
            writer.writerow(["", ""])
            writer.writerow(["Câu 1: Phân tích cơ cấu tài sản của VinGroup", ""])
            writer.writerow(["- Tính tỷ trọng tài sản ngắn hạn/dài hạn", ""])
            writer.writerow(["- Nhận xét về sự thay đổi giữa 2023 và 2024", ""])
            writer.writerow(["", ""])
            writer.writerow(["Câu 2: Đánh giá khả năng thanh khoản", ""])
            writer.writerow(["- Tính và giải thích các chỉ số thanh khoản", ""])
            writer.writerow(["- So sánh với chuẩn mực ngành", ""])
            writer.writerow(["", ""])
            writer.writerow(["Câu 3: Phân tích khả năng sinh lời", ""])
            writer.writerow(["- Tính ROE, ROA, biên lợi nhuận", ""])
            writer.writerow(["- Đánh giá xu hướng và nguyên nhân", ""])
            writer.writerow(["", ""])
            writer.writerow(["Câu 4: Đánh giá hiệu quả hoạt động", ""])
            writer.writerow(["- Tính vòng quay tài sản, vòng quay hàng tồn kho", ""])
            writer.writerow(["- Nhận xét về hiệu quả quản lý", ""])
            writer.writerow(["", ""])
            writer.writerow(["Câu 5: Phân tích cơ cấu tài chính", ""])
            writer.writerow(["- Tính tỷ lệ nợ/vốn chủ sở hữu", ""])
            writer.writerow(["- Đánh giá rủi ro tài chính", ""])
            writer.writerow(["", ""])
            
            # Answer guidelines
            writer.writerow(["C. GỢI Ý TRÌNH BÀY KẾT QUẢ", ""])
            writer.writerow(["", ""])
            writer.writerow(["1. Cấu trúc báo cáo phân tích:", ""])
            writer.writerow(["- Tóm tắt tình hình tài chính", ""])
            writer.writerow(["- Phân tích chi tiết các chỉ số", ""])
            writer.writerow(["- Nhận xét về xu hướng", ""])
            writer.writerow(["- Khuyến nghị và đề xuất", ""])
            writer.writerow(["", ""])
            writer.writerow(["2. Cách trình bày số liệu:", ""])
            writer.writerow(["- Sử dụng bảng biểu, biểu đồ", ""])
            writer.writerow(["- Làm nổi bật các điểm quan trọng", ""])
            writer.writerow(["- So sánh với năm trước", ""])

def main():
    """Main function to generate the financial analysis files"""
    print("Đang tạo files phân tích tài chính VinGroup...")
    
    # Create analyzer instance
    analyzer = VinGroupFinancialAnalyzer(VINGROUP_DATA)
    
    # Generate CSV reports
    analyzer.generate_csv_reports("vingroup_analysis")
    
    # Generate JSON data for Excel script
    with open("vingroup_analysis/vingroup_data.json", "w", encoding="utf-8") as f:
        json.dump(VINGROUP_DATA, f, ensure_ascii=False, indent=2)
    
    print("✓ Đã tạo xong các file CSV trong thư mục 'vingroup_analysis'")
    print("✓ Đã tạo file dữ liệu JSON: vingroup_data.json")
    print("\nCác file được tạo:")
    print("- balance_sheet.csv: Bảng cân đối kế toán")
    print("- income_statement.csv: Báo cáo kết quả kinh doanh")
    print("- cash_flow.csv: Báo cáo lưu chuyển tiền tệ")
    print("- financial_ratios.csv: Phân tích chỉ số tài chính")
    print("- guidelines_exercises.csv: Hướng dẫn và bài tập")
    print("- vingroup_data.json: Dữ liệu gốc cho Excel generator")
    
    # Display some sample calculations
    print("\n" + "="*50)
    print("MẪU KẾT QUẢ PHÂN TÍCH:")
    print("="*50)
    
    ratios_2023 = analyzer.calculate_financial_ratios("2023")
    ratios_2024 = analyzer.calculate_financial_ratios("2024")
    
    print(f"Hệ số thanh khoản hiện hành: 2023={ratios_2023['current_ratio']:.2f}, 2024={ratios_2024['current_ratio']:.2f}")
    print(f"ROE: 2023={ratios_2023['roe']:.2f}%, 2024={ratios_2024['roe']:.2f}%")
    print(f"ROA: 2023={ratios_2023['roa']:.2f}%, 2024={ratios_2024['roa']:.2f}%")
    print(f"Biên lợi nhuận ròng: 2023={ratios_2023['net_profit_margin']:.2f}%, 2024={ratios_2024['net_profit_margin']:.2f}%")

if __name__ == "__main__":
    main()