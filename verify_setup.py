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
    
    print("3. Testing enhanced API features...")
    try:
        comprehensive_data = fetcher.get_comprehensive_data("AAPL")
        
        if comprehensive_data.get('analyst_data', {}).get('estimates'):
            print("   ✓ Analyst estimates available (premium feature)")
        else:
            print("   ⚠ Analyst estimates not available (may require premium subscription)")
            
        if comprehensive_data.get('analyst_data', {}).get('price_targets'):
            print("   ✓ Price targets available (premium feature)")
        else:
            print("   ⚠ Price targets not available (may require premium subscription)")
            
        key_metrics = comprehensive_data.get('key_metrics', {})
        if key_metrics.get('enterpriseValueMultipleTTM'):
            print("   ✓ Enhanced key metrics available")
        else:
            print("   ⚠ Limited key metrics data")
            
    except Exception as e:
        print(f"   ✗ Enhanced API features test failed: {e}")
    
    print("4. Testing file generation...")
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
    print("API Configuration Status:")
    print(f"  - FMP API Key: {'✓ Configured' if fetcher.fmp_api_key else '✗ Missing'}")
    print(f"  - Alpha Vantage API Key: {'✓ Configured' if fetcher.alpha_vantage_key else '✗ Missing'}")
    print()
    print("Next Steps:")
    print("1. Install FMP Excel add-in: See EXCEL_ADDON_SETUP_GUIDE.md")
    print("2. Generate models: python examples/analyze_stock.py AAPL")
    print("3. Test enhanced data quality with different symbols")
    print()
    print("If you're still having issues:")
    print("1. Make sure you're looking for files with the NEW symbol name")
    print("2. Check the output/ folder for recently created files")
    print("3. The scripts create NEW files, they don't update existing ones")
    return True

if __name__ == "__main__":
    verify_setup()
