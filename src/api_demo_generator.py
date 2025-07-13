#!/usr/bin/env python3
"""
Generate Excel demonstration files showing FMP and Alpha Vantage API usage
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from src.data_fetcher import FinancialDataFetcher
import requests

class APIDemoGenerator:
    def __init__(self):
        self.fetcher = FinancialDataFetcher()
        
    def create_api_demo_spreadsheet(self, symbol, output_path):
        """Create Excel demo showing FMP and Alpha Vantage API usage"""
        wb = Workbook()
        ws = wb.active
        ws.title = f"{symbol} API Demo"
        
        header_font = Font(bold=True, size=12, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        data_font = Font(size=11)
        formula_font = Font(size=10, italic=True, color="0066CC")
        border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                       top=Side(style='thin'), bottom=Side(style='thin'))
        
        ws.merge_cells('A1:F1')
        ws['A1'] = f"{symbol} - API Integration Demo"
        ws['A1'].font = Font(bold=True, size=16)
        ws['A1'].alignment = Alignment(horizontal='center')
        
        ws.merge_cells('A3:F3')
        ws['A3'] = "This spreadsheet demonstrates FMP and Alpha Vantage API integration"
        ws['A3'].font = Font(size=11, italic=True)
        ws['A3'].alignment = Alignment(horizontal='center')
        
        row = 5
        ws[f'A{row}'] = "Financial Modeling Prep (FMP) Data"
        ws[f'A{row}'].font = header_font
        ws[f'A{row}'].fill = header_fill
        ws.merge_cells(f'A{row}:D{row}')
        
        row += 2
        
        try:
            profile_data = self.fetcher.get_company_profile(symbol)
            financial_data = self.fetcher.get_financial_statements(symbol)
            key_metrics = self.fetcher.get_key_metrics(symbol)
            
            ws[f'A{row}'] = "Company Profile:"
            ws[f'A{row}'].font = Font(bold=True)
            row += 1
            
            ws[f'A{row}'] = "Company Name:"
            ws[f'B{row}'] = profile_data.get('companyName', 'N/A')
            ws[f'D{row}'] = "Excel Add-in Formula:"
            ws[f'E{row}'] = f'=FMPProfile("{symbol}")'
            ws[f'E{row}'].font = formula_font
            row += 1
            
            ws[f'A{row}'] = "Current Price:"
            ws[f'B{row}'] = f"${profile_data.get('price', 0):.2f}"
            ws[f'E{row}'] = f'=FMPPrice("{symbol}")'
            ws[f'E{row}'].font = formula_font
            row += 1
            
            ws[f'A{row}'] = "Market Cap:"
            ws[f'B{row}'] = f"${profile_data.get('marketCap', 0):,.0f}"
            ws[f'E{row}'] = f'=FMPMarketCap("{symbol}")'
            ws[f'E{row}'].font = formula_font
            row += 2
            
            ws[f'A{row}'] = "LTM Financial Metrics:"
            ws[f'A{row}'].font = Font(bold=True)
            row += 1
            
            income_stmt = financial_data.get('income_statement', [])
            if income_stmt:
                latest_data = income_stmt[0]  # Most recent year
                eps = latest_data.get('eps', 0)
                revenue = latest_data.get('revenue', 0)
                net_income = latest_data.get('netIncome', 0)
                
                ws[f'A{row}'] = "LTM EPS:"
                ws[f'B{row}'] = f"${eps:.2f}"
                ws[f'D{row}'] = "Excel Add-in Formula:"
                ws[f'E{row}'] = f'=FMPIncomeStatement("{symbol}", "TTM", "eps")'
                ws[f'E{row}'].font = formula_font
                row += 1
                
                ws[f'A{row}'] = "LTM Revenue:"
                ws[f'B{row}'] = f"${revenue:,.0f}"
                ws[f'E{row}'] = f'=FMPIncomeStatement("{symbol}", "TTM", "revenue")'
                ws[f'E{row}'].font = formula_font
                row += 1
                
                ws[f'A{row}'] = "LTM Net Income:"
                ws[f'B{row}'] = f"${net_income:,.0f}"
                ws[f'E{row}'] = f'=FMPIncomeStatement("{symbol}", "TTM", "netIncome")'
                ws[f'E{row}'].font = formula_font
                row += 2
            
            if key_metrics:
                ws[f'A{row}'] = "Key Valuation Metrics:"
                ws[f'A{row}'].font = Font(bold=True)
                row += 1
                
                pe_ratio = key_metrics.get('peRatioTTM', 0)
                pb_ratio = key_metrics.get('pbRatioTTM', 0)
                ev_ebitda = key_metrics.get('enterpriseValueMultipleTTM', 0)
                
                ws[f'A{row}'] = "P/E Ratio (TTM):"
                ws[f'B{row}'] = f"{pe_ratio:.2f}" if pe_ratio else "N/A"
                ws[f'E{row}'] = f'=FMPKeyMetrics("{symbol}", "peRatioTTM")'
                ws[f'E{row}'].font = formula_font
                row += 1
                
                ws[f'A{row}'] = "P/B Ratio (TTM):"
                ws[f'B{row}'] = f"{pb_ratio:.2f}" if pb_ratio else "N/A"
                ws[f'E{row}'] = f'=FMPKeyMetrics("{symbol}", "pbRatioTTM")'
                ws[f'E{row}'].font = formula_font
                row += 1
                
                ws[f'A{row}'] = "EV/EBITDA (TTM):"
                ws[f'B{row}'] = f"{ev_ebitda:.2f}" if ev_ebitda else "N/A"
                ws[f'E{row}'] = f'=FMPKeyMetrics("{symbol}", "enterpriseValueMultipleTTM")'
                ws[f'E{row}'].font = formula_font
                row += 3
                
        except Exception as e:
            ws[f'A{row}'] = f"Error fetching FMP data: {str(e)}"
            row += 2
        
        ws[f'A{row}'] = "Alpha Vantage Data"
        ws[f'A{row}'].font = header_font
        ws[f'A{row}'].fill = header_fill
        ws.merge_cells(f'A{row}:D{row}')
        row += 2
        
        try:
            if self.fetcher.alpha_vantage_key:
                av_url = f"https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol={symbol}&apikey={self.fetcher.alpha_vantage_key}"
                response = requests.get(av_url)
                av_data = response.json()
                
                quote_data = av_data.get('Global Quote', {})
                if quote_data:
                    ws[f'A{row}'] = "Real-time Quote Data:"
                    ws[f'A{row}'].font = Font(bold=True)
                    row += 1
                    
                    current_price = quote_data.get('05. price', '0')
                    change = quote_data.get('09. change', '0')
                    change_percent = quote_data.get('10. change percent', '0%')
                    volume = quote_data.get('06. volume', '0')
                    
                    ws[f'A{row}'] = "Current Price:"
                    ws[f'B{row}'] = f"${float(current_price):.2f}"
                    ws[f'D{row}'] = "Alpha Vantage API Call:"
                    ws[f'E{row}'] = f'=WEBSERVICE("https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol={symbol}&apikey=YOUR_KEY")'
                    ws[f'E{row}'].font = formula_font
                    row += 1
                    
                    ws[f'A{row}'] = "Daily Change:"
                    ws[f'B{row}'] = f"${float(change):.2f} ({change_percent})"
                    row += 1
                    
                    ws[f'A{row}'] = "Volume:"
                    ws[f'B{row}'] = f"{int(volume):,}"
                    row += 2
                    
                ts_url = f"https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol={symbol}&apikey={self.fetcher.alpha_vantage_key}"
                ts_response = requests.get(ts_url)
                ts_data = ts_response.json()
                
                time_series = ts_data.get('Time Series (Daily)', {})
                if time_series:
                    ws[f'A{row}'] = "Recent Price History:"
                    ws[f'A{row}'].font = Font(bold=True)
                    row += 1
                    
                    ws[f'A{row}'] = "Date"
                    ws[f'B{row}'] = "Open"
                    ws[f'C{row}'] = "High"
                    ws[f'D{row}'] = "Low"
                    ws[f'E{row}'] = "Close"
                    ws[f'F{row}'] = "Volume"
                    
                    for col in ['A', 'B', 'C', 'D', 'E', 'F']:
                        ws[f'{col}{row}'].font = Font(bold=True)
                        ws[f'{col}{row}'].fill = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")
                    row += 1
                    
                    dates = list(time_series.keys())[:5]
                    for date in dates:
                        day_data = time_series[date]
                        ws[f'A{row}'] = date
                        ws[f'B{row}'] = f"${float(day_data['1. open']):.2f}"
                        ws[f'C{row}'] = f"${float(day_data['2. high']):.2f}"
                        ws[f'D{row}'] = f"${float(day_data['3. low']):.2f}"
                        ws[f'E{row}'] = f"${float(day_data['4. close']):.2f}"
                        ws[f'F{row}'] = f"{int(day_data['5. volume']):,}"
                        row += 1
                        
        except Exception as e:
            ws[f'A{row}'] = f"Error fetching Alpha Vantage data: {str(e)}"
            row += 2
        
        row += 2
        ws[f'A{row}'] = "Excel Add-in Installation Instructions:"
        ws[f'A{row}'].font = Font(bold=True, size=12)
        row += 1
        
        instructions = [
            "1. Open Excel (desktop version)",
            "2. Go to Insert > Get Add-ins",
            "3. Search for 'Financial Modeling Prep'",
            "4. Install the FMP add-in",
            f"5. Configure with API key: {self.fetcher.fmp_api_key[:8]}...",
            "6. Use the formulas shown in column E above",
            "",
            "For Alpha Vantage:",
            "- Use WEBSERVICE() function with Alpha Vantage API URLs",
            f"- Your Alpha Vantage key: {self.fetcher.alpha_vantage_key[:8]}...",
            "- See examples in column E above"
        ]
        
        for instruction in instructions:
            ws[f'A{row}'] = instruction
            ws[f'A{row}'].font = Font(size=10)
            row += 1
        
        for column in ws.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width
        
        wb.save(output_path)
        print(f"API demo spreadsheet saved to {output_path}")
        
        return output_path

def main():
    generator = APIDemoGenerator()
    output_path = "output/AAPL_API_Demo.xlsx"
    os.makedirs("output", exist_ok=True)
    generator.create_api_demo_spreadsheet("AAPL", output_path)

if __name__ == "__main__":
    main()
