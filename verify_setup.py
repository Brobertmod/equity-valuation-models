#!/usr/bin/env python3
"""
Verification script to test the equity valuation setup
Usage: python verify_setup.py
"""
import os
import sys
sys.path.append(os.path.join(os.path.dirname(__file__), '.'))

from src.data_fetcher import FinancialDataFetcher

def verify_setup():
    print("=== Equity Valuation Models - Setup Verification ===")
    print()
    
    print("1. Testing data fetching...")
    fetcher = FinancialDataFetcher()
    test_data = fetcher.get_company_profile("AAPL")
    
    if test_data and test_data.get('companyName'):
        print(f"   ✓ Data fetching works: {test_data.get('companyName')}")
    else:
        print("   ✗ Data fetching failed - check internet connection")
        return False
    
    print("2. Testing output directory...")
    if os.path.exists("output"):
        print("   ✓ Output directory exists")
    else:
        print("   ✗ Output directory missing")
        return False
    
    print("3. Testing file generation...")
    from src.excel_generator import ExcelTemplateGenerator
    generator = ExcelTemplateGenerator()
    
    try:
        test_file = "output/TEST_verification.xlsx"
        generator.create_ev_sales_template("TEST", test_data, test_file)
        if os.path.exists(test_file):
            print("   ✓ Excel file generation works")
            os.remove(test_file)  # Clean up test file
        else:
            print("   ✗ Excel file generation failed")
            return False
    except Exception as e:
        print(f"   ✗ Excel file generation failed: {e}")
        return False
    
    print()
    print("✓ Setup verification complete - system is working correctly")
    print()
    print("If you're still having issues:")
    print("1. Make sure you're looking for files with the NEW symbol name")
    print("2. Check the output/ folder for recently created files")
    print("3. The scripts create NEW files, they don't update existing ones")
    return True

if __name__ == "__main__":
    verify_setup()
