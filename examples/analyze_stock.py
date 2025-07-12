#!/usr/bin/env python3
"""
Example script demonstrating how to use the equity valuation models
Usage: python examples/analyze_stock.py AAPL
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from src.data_fetcher import FinancialDataFetcher
from src.excel_generator import ExcelTemplateGenerator
import argparse

def main():
    parser = argparse.ArgumentParser(description='Generate equity valuation models for a stock')
    parser.add_argument('symbol', help='Stock symbol (e.g., AAPL, MSFT, GOOGL)')
    parser.add_argument('--models', nargs='+', 
                       choices=['ev_ebitda', 'ev_sales', 'dcf', 'all'],
                       default=['all'],
                       help='Which models to generate')
    parser.add_argument('--output-dir', default='output',
                       help='Output directory for Excel files')
    
    args = parser.parse_args()
    
    os.makedirs(args.output_dir, exist_ok=True)
    
    fetcher = FinancialDataFetcher()
    generator = ExcelTemplateGenerator()
    
    print(f"Analyzing {args.symbol}...")
    
    data = fetcher.get_comprehensive_data(args.symbol)
    
    if not data.get('profile'):
        print(f"Error: Could not fetch data for {args.symbol}")
        return
    
    models_to_generate = args.models if 'all' not in args.models else ['ev_ebitda', 'ev_sales', 'dcf']
    
    for model in models_to_generate:
        output_path = os.path.join(args.output_dir, f"{args.symbol}_{model}_model.xlsx")
        
        if model == 'ev_ebitda':
            generator.create_ev_ebitda_template(args.symbol, data, output_path)
        elif model == 'ev_sales':
            generator.create_ev_sales_template(args.symbol, data, output_path)
        elif model == 'dcf':
            generator.create_dcf_template(args.symbol, data, output_path)
    
    print(f"\nValuation models generated for {args.symbol}")
    print(f"Files saved to: {args.output_dir}/")
    print(f"\nNew files created:")
    for model in models_to_generate:
        output_path = os.path.join(args.output_dir, f"{args.symbol}_{model}_model.xlsx")
        if os.path.exists(output_path):
            stat = os.stat(output_path)
            size_kb = stat.st_size / 1024
            print(f"  - {os.path.basename(output_path)} ({size_kb:.1f} KB)")
    
    profile = data.get('profile', {})
    key_metrics = data.get('key_metrics', {})
    
    print(f"\n--- {args.symbol} Summary ---")
    print(f"Company: {profile.get('companyName', 'N/A')}")
    print(f"Current Price: ${profile.get('price', 0):.2f}")
    print(f"Market Cap: ${profile.get('marketCap', 0):,.0f}")
    print(f"EV/EBITDA: {key_metrics.get('enterpriseValueMultipleTTM', 'N/A')}")
    print(f"EV/Sales: {key_metrics.get('evToSalesTTM', 'N/A')}")
    print(f"P/E Ratio: {key_metrics.get('peRatioTTM', 'N/A')}")

if __name__ == "__main__":
    main()
