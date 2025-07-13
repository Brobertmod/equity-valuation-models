#!/usr/bin/env python3
"""
Test script to verify API keys are working correctly
"""
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '.'))

from src.data_fetcher import FinancialDataFetcher

def test_api_keys():
    print("=== Testing API Keys Configuration ===")
    print()
    
    fetcher = FinancialDataFetcher()
    
    print(f"FMP API Key configured: {'Yes' if fetcher.fmp_api_key else 'No'}")
    print(f"Alpha Vantage API Key configured: {'Yes' if fetcher.alpha_vantage_key else 'No'}")
    print()
    
    print("1. Testing Financial Modeling Prep API...")
    try:
        profile_data = fetcher.get_company_profile("AAPL")
        if profile_data and profile_data.get('companyName'):
            print(f"   ✓ FMP API working: {profile_data.get('companyName')}")
            print(f"   ✓ Current Price: ${profile_data.get('price', 0):.2f}")
            print(f"   ✓ Market Cap: ${profile_data.get('marketCap', 0):,.0f}")
        else:
            print("   ✗ FMP API failed or returned no data")
    except Exception as e:
        print(f"   ✗ FMP API error: {e}")
    
    print()
    
    print("2. Testing enhanced financial data with FMP...")
    try:
        financial_data = fetcher.get_financial_statements("AAPL")
        if financial_data.get('income_statement'):
            print("   ✓ Income statement data retrieved")
        if financial_data.get('balance_sheet'):
            print("   ✓ Balance sheet data retrieved")
        if financial_data.get('cash_flow'):
            print("   ✓ Cash flow data retrieved")
        
        key_metrics = fetcher.get_key_metrics("AAPL")
        if key_metrics:
            print("   ✓ Key metrics data retrieved")
            print(f"   ✓ P/E Ratio: {key_metrics.get('peRatioTTM', 'N/A')}")
            print(f"   ✓ EV/EBITDA: {key_metrics.get('enterpriseValueMultipleTTM', 'N/A')}")
        
    except Exception as e:
        print(f"   ✗ Enhanced data error: {e}")
    
    print()
    
    print("3. Testing analyst estimates and price targets...")
    try:
        analyst_data = fetcher.get_analyst_estimates("AAPL")
        if analyst_data.get('estimates'):
            print("   ✓ Analyst estimates retrieved")
        if analyst_data.get('price_targets'):
            print("   ✓ Price targets retrieved")
        if not analyst_data.get('estimates') and not analyst_data.get('price_targets'):
            print("   ⚠ No analyst data (may require premium FMP subscription)")
    except Exception as e:
        print(f"   ✗ Analyst data error: {e}")
    
    print()
    print("=== API Configuration Complete ===")
    print("Your API keys have been configured and tested!")
    print()
    print("Next steps:")
    print("1. Generate models with enhanced data: python examples/analyze_stock.py AAPL")
    print("2. Try different symbols to test data quality")
    print("3. Check if premium features work with your FMP subscription level")

if __name__ == "__main__":
    test_api_keys()
