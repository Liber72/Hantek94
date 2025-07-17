#!/usr/bin/env python3
"""
Demo script for VinGroup Financial Analysis Tool
===============================================

This script demonstrates how to use the VinGroup Financial Analysis Tool
to create both CSV and Excel files.

Usage:
    python demo.py
"""

import os
import sys
from datetime import datetime

def main():
    print("="*60)
    print("VINGROUP FINANCIAL ANALYSIS TOOL - DEMO")
    print("="*60)
    print(f"Demo started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    # Step 1: Generate CSV files
    print("🔄 Step 1: Generating CSV files...")
    print("Running: python vingroup_financial_analyzer.py")
    print("-" * 40)
    
    try:
        os.system("python3 vingroup_financial_analyzer.py")
        print("✅ CSV files generated successfully!")
    except Exception as e:
        print(f"❌ Error generating CSV files: {e}")
        return
    
    print()
    
    # Step 2: Try to generate Excel file
    print("🔄 Step 2: Attempting to generate Excel file...")
    print("Running: python excel_generator.py")
    print("-" * 40)
    
    try:
        os.system("python3 excel_generator.py")
        print("✅ Excel generation attempt completed!")
    except Exception as e:
        print(f"❌ Error running Excel generator: {e}")
    
    print()
    
    # Step 3: Show generated files
    print("📁 Step 3: Generated files overview:")
    print("-" * 40)
    
    # Check CSV files
    csv_dir = "vingroup_analysis"
    if os.path.exists(csv_dir):
        print(f"📂 CSV files in '{csv_dir}':")
        for file in os.listdir(csv_dir):
            file_path = os.path.join(csv_dir, file)
            if os.path.isfile(file_path):
                size = os.path.getsize(file_path)
                print(f"  ✓ {file} ({size} bytes)")
    else:
        print("❌ CSV directory not found")
    
    # Check Excel file
    excel_file = "VinGroup_Financial_Analysis.xlsx"
    if os.path.exists(excel_file):
        size = os.path.getsize(excel_file)
        print(f"📊 Excel file: {excel_file} ({size} bytes)")
    else:
        print("ℹ️  Excel file not generated (likely due to missing openpyxl)")
    
    print()
    
    # Step 4: Show usage instructions
    print("📋 Step 4: Usage instructions:")
    print("-" * 40)
    print("To use the generated files:")
    print("1. Open CSV files in Excel or LibreOffice Calc")
    print("2. Use the financial_ratios.csv for analysis")
    print("3. Follow guidelines_exercises.csv for student exercises")
    print()
    print("To generate Excel file with full features:")
    print("1. Install openpyxl: pip install openpyxl")
    print("2. Run: python excel_generator.py")
    print()
    
    # Step 5: Show sample data
    print("📊 Step 5: Sample financial data:")
    print("-" * 40)
    
    try:
        from vingroup_financial_analyzer import VinGroupFinancialAnalyzer, VINGROUP_DATA
        analyzer = VinGroupFinancialAnalyzer(VINGROUP_DATA)
        
        # Show sample ratios
        ratios_2023 = analyzer.calculate_financial_ratios("2023")
        ratios_2024 = analyzer.calculate_financial_ratios("2024")
        
        print("Key Financial Ratios (VinGroup):")
        print(f"  Current Ratio:      2023={ratios_2023['current_ratio']:.2f}    2024={ratios_2024['current_ratio']:.2f}")
        print(f"  ROE:                2023={ratios_2023['roe']:.2f}%      2024={ratios_2024['roe']:.2f}%")
        print(f"  ROA:                2023={ratios_2023['roa']:.2f}%      2024={ratios_2024['roa']:.2f}%")
        print(f"  Net Profit Margin:  2023={ratios_2023['net_profit_margin']:.2f}%      2024={ratios_2024['net_profit_margin']:.2f}%")
        print(f"  Debt-to-Equity:     2023={ratios_2023['debt_to_equity']:.2f}      2024={ratios_2024['debt_to_equity']:.2f}")
        
    except Exception as e:
        print(f"❌ Error displaying sample data: {e}")
    
    print()
    print("="*60)
    print("✅ DEMO COMPLETED SUCCESSFULLY!")
    print("="*60)
    print("Next steps:")
    print("1. Review generated CSV files in the 'vingroup_analysis' folder")
    print("2. Install openpyxl if you want to generate Excel files")
    print("3. Use the files for financial analysis exercises")
    print("4. Follow the README.md for detailed instructions")
    print()
    print("Thank you for using VinGroup Financial Analysis Tool!")

if __name__ == "__main__":
    main()