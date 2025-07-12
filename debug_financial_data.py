#!/usr/bin/env python3
"""
Debug script to examine financial data and identify calculation issues
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '.'))

from src.data_fetcher import FinancialDataFetcher

def debug_financial_data(symbol):
    """Debug financial data for a given symbol"""
    fetcher = FinancialDataFetcher()
    data = fetcher.get_comprehensive_data(symbol)

    profile = data.get('profile', {})
    ev_data = data.get('enterprise_value', {})
    financial_statements = data.get('financial_statements', {})
    income_statement = financial_statements.get('income_statement', [])

    print(f'=== {symbol} Financial Data Debug ===')
    print(f'Current Price: {profile.get("price", 0)}')
    print(f'Market Cap: {ev_data.get("marketCapitalization", 0):,}')
    print(f'Total Debt: {ev_data.get("totalDebt", 0):,}')
    print(f'Cash: {ev_data.get("cashAndCashEquivalents", 0):,}')
    print(f'Enterprise Value: {ev_data.get("enterpriseValue", 0):,}')

    if income_statement:
        print(f'Income statement structure: {type(income_statement)}')
        print(f'Income statement length: {len(income_statement)}')
        if income_statement:
            latest = income_statement[0]
            print(f'Latest statement keys: {list(latest.keys())[:10]}')  # Show first 10 keys
            
            ebitda_keys = ['ebitda', 'operatingIncome', 'EBITDA', 'Operating Income', 'EBIT']
            revenue_keys = ['revenue', 'totalRevenue', 'Total Revenue', 'Net Sales', 'Sales']
            
            ebitda = 0
            for key in ebitda_keys:
                if key in latest and latest[key] is not None:
                    ebitda = latest[key]
                    print(f'Found EBITDA using key: {key}')
                    break
            
            revenue = 0
            for key in revenue_keys:
                if key in latest and latest[key] is not None:
                    revenue = latest[key]
                    print(f'Found Revenue using key: {key}')
                    break
            
            print(f'EBITDA: {ebitda:,}')
            print(f'Revenue: {revenue:,}')
        else:
            print('Income statement is empty')
            ebitda = 0
            revenue = 0
        
        net_debt = ev_data.get('totalDebt', 0) - ev_data.get('cashAndCashEquivalents', 0)
        shares_outstanding = ev_data.get('marketCapitalization', 0) / profile.get('price', 1) if profile.get('price', 0) > 0 else 0
        
        print(f'Net Debt: {net_debt:,}')
        print(f'Shares Outstanding: {shares_outstanding:,}')
        
        test_multiples = [15.0, 18.0, 22.0]
        print(f'\n=== Test Calculations ===')
        for multiple in test_multiples:
            target_ev = ebitda * multiple
            target_equity = target_ev - net_debt
            target_price = target_equity / shares_outstanding if shares_outstanding > 0 else 0
            upside = (target_price - profile.get('price', 0)) / profile.get('price', 1) if profile.get('price', 0) > 0 else 0
            
            print(f'{multiple}x Multiple:')
            print(f'  Target EV: {target_ev:,}')
            print(f'  Target Equity: {target_equity:,}')
            print(f'  Target Price: ${target_price:.2f}')
            print(f'  Upside: {upside:.1%}')
            print()
    else:
        print('No income statement data available')

if __name__ == "__main__":
    symbols = ['AAPL', 'MSFT']
    for symbol in symbols:
        debug_financial_data(symbol)
        print('\n' + '='*50 + '\n')
