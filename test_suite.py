#!/usr/bin/env python3
"""
Test script for VinGroup Financial Analysis Tool
===============================================

This script tests all components of the VinGroup Financial Analysis Tool.
"""

import os
import sys
import json
import csv
from datetime import datetime

def test_data_integrity():
    """Test data integrity and structure"""
    print("🔍 Testing data integrity...")
    
    try:
        from vingroup_financial_analyzer import VINGROUP_DATA
        
        # Test data structure
        required_keys = ['company_info', 'balance_sheet', 'income_statement', 'cash_flow']
        for key in required_keys:
            if key not in VINGROUP_DATA:
                print(f"❌ Missing key: {key}")
                return False
        
        # Test years
        years = ['2023', '2024']
        for year in years:
            if year not in VINGROUP_DATA['balance_sheet']:
                print(f"❌ Missing year in balance_sheet: {year}")
                return False
        
        print("✅ Data integrity check passed")
        return True
        
    except Exception as e:
        print(f"❌ Data integrity test failed: {e}")
        return False

def test_csv_generation():
    """Test CSV file generation"""
    print("🔍 Testing CSV generation...")
    
    try:
        from vingroup_financial_analyzer import VinGroupFinancialAnalyzer, VINGROUP_DATA
        
        analyzer = VinGroupFinancialAnalyzer(VINGROUP_DATA)
        
        # Generate CSV files
        analyzer.generate_csv_reports("test_output")
        
        # Create JSON file for Excel generator test
        json_file = "test_output/vingroup_data.json"
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(VINGROUP_DATA, f, ensure_ascii=False, indent=2)
        
        # Check if files were created
        expected_files = [
            'balance_sheet.csv',
            'income_statement.csv', 
            'cash_flow.csv',
            'financial_ratios.csv',
            'guidelines_exercises.csv'
        ]
        
        for file in expected_files:
            file_path = os.path.join("test_output", file)
            if not os.path.exists(file_path):
                print(f"❌ Missing file: {file}")
                return False
            
            # Check file size
            if os.path.getsize(file_path) == 0:
                print(f"❌ Empty file: {file}")
                return False
        
        print("✅ CSV generation test passed")
        return True
        
    except Exception as e:
        print(f"❌ CSV generation test failed: {e}")
        return False

def test_financial_calculations():
    """Test financial ratio calculations"""
    print("🔍 Testing financial calculations...")
    
    try:
        from vingroup_financial_analyzer import VinGroupFinancialAnalyzer, VINGROUP_DATA
        
        analyzer = VinGroupFinancialAnalyzer(VINGROUP_DATA)
        
        # Test ratio calculations
        ratios_2023 = analyzer.calculate_financial_ratios("2023")
        ratios_2024 = analyzer.calculate_financial_ratios("2024")
        
        # Check key ratios
        key_ratios = ['current_ratio', 'quick_ratio', 'roa', 'roe', 'debt_to_equity']
        
        for ratio in key_ratios:
            if ratio not in ratios_2023 or ratio not in ratios_2024:
                print(f"❌ Missing ratio: {ratio}")
                return False
            
            # Check for reasonable values
            if ratios_2023[ratio] < 0 or ratios_2024[ratio] < 0:
                # Only some ratios should be negative
                if ratio not in ['debt_to_equity']:
                    print(f"❌ Negative ratio: {ratio}")
                    return False
        
        print("✅ Financial calculations test passed")
        return True
        
    except Exception as e:
        print(f"❌ Financial calculations test failed: {e}")
        return False

def test_csv_content():
    """Test CSV content format"""
    print("🔍 Testing CSV content format...")
    
    try:
        # Test financial ratios CSV
        ratios_file = "test_output/financial_ratios.csv"
        if not os.path.exists(ratios_file):
            print(f"❌ File not found: {ratios_file}")
            return False
        
        with open(ratios_file, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = list(reader)
            
            # Check header
            if len(rows) < 5:
                print("❌ CSV too short")
                return False
            
            # Check for Vietnamese content
            found_vietnamese = False
            for row in rows:
                if any('CHỈ SỐ' in cell for cell in row):
                    found_vietnamese = True
                    break
            
            if not found_vietnamese:
                print("❌ Vietnamese content not found")
                return False
        
        print("✅ CSV content test passed")
        return True
        
    except Exception as e:
        print(f"❌ CSV content test failed: {e}")
        return False

def test_excel_generator_structure():
    """Test Excel generator structure (without openpyxl)"""
    print("🔍 Testing Excel generator structure...")
    
    try:
        from excel_generator import VinGroupExcelGenerator, OPENPYXL_AVAILABLE
        
        if not OPENPYXL_AVAILABLE:
            print("ℹ️  openpyxl not available, skipping Excel generator test")
            return True
        
        # Test initialization
        generator = VinGroupExcelGenerator("test_output/vingroup_data.json")
        
        # Test data loading
        if not generator.data:
            print("❌ Data not loaded")
            return False
        
        # Test style definitions
        if not hasattr(generator, 'header_font'):
            print("❌ Styles not defined")
            return False
        
        print("✅ Excel generator structure test passed")
        return True
        
    except Exception as e:
        print(f"❌ Excel generator structure test failed: {e}")
        return False

def main():
    """Run all tests"""
    print("🧪 VinGroup Financial Analysis Tool - Test Suite")
    print("=" * 55)
    print(f"Test started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    tests = [
        test_data_integrity,
        test_csv_generation,
        test_financial_calculations,
        test_csv_content,
        test_excel_generator_structure
    ]
    
    passed = 0
    total = len(tests)
    
    for test in tests:
        try:
            if test():
                passed += 1
        except Exception as e:
            print(f"❌ Test {test.__name__} crashed: {e}")
    
    print()
    print("=" * 55)
    print(f"📊 Test Results: {passed}/{total} tests passed")
    
    if passed == total:
        print("✅ All tests passed!")
        print("🎉 VinGroup Financial Analysis Tool is ready to use!")
    else:
        print("❌ Some tests failed!")
        print("🔧 Please check the implementation")
    
    # Clean up test files
    try:
        import shutil
        if os.path.exists("test_output"):
            shutil.rmtree("test_output")
        print("🧹 Test files cleaned up")
    except:
        pass
    
    return passed == total

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)