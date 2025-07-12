import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.formatting.rule import ColorScaleRule
import xlsxwriter
from typing import Dict, List, Any
import os

class ExcelTemplateGenerator:
    """
    Generate Excel templates for equity valuation models with embedded formulas
    """
    
    def __init__(self):
        self.header_font = Font(bold=True, color="FFFFFF")
        self.header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        self.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        self.center_alignment = Alignment(horizontal='center', vertical='center')
        
    def create_ev_ebitda_template(self, symbol: str, data: Dict, output_path: str):
        """Create EV/EBITDA valuation template with embedded formulas"""
        
        wb = Workbook()
        
        ws_summary = wb.active
        ws_summary.title = "Summary"
        
        ws_summary['A1'] = "EV/EBITDA Valuation Model"
        ws_summary['A1'].font = Font(size=16, bold=True)
        ws_summary.merge_cells('A1:F1')
        
        ws_summary['A2'] = "Ticker Symbol:"
        ws_summary['B2'] = symbol
        ws_summary['B2'].font = Font(size=14, bold=True, color="FF0000")
        ws_summary['B2'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        ws_summary['B2'].border = Border(
            left=Side(style='thick'),
            right=Side(style='thick'),
            top=Side(style='thick'),
            bottom=Side(style='thick')
        )
        
        ws_summary['D2'] = "To analyze a new symbol:"
        ws_summary['E2'] = "1. Change symbol in B2"
        ws_summary['F2'] = "2. Run: python examples/analyze_stock.py [NEW_SYMBOL]"
        ws_summary['D2'].font = Font(size=10, italic=True)
        ws_summary['E2'].font = Font(size=10, italic=True)
        ws_summary['F2'].font = Font(size=10, italic=True)
        
        profile = data.get('profile', {})
        ws_summary['A4'] = "Company Name:"
        ws_summary['B4'] = profile.get('companyName', '')
        ws_summary['A5'] = "Sector:"
        ws_summary['B5'] = profile.get('sector', '')
        ws_summary['A6'] = "Industry:"
        ws_summary['B6'] = profile.get('industry', '')
        ws_summary['A7'] = "Current Price:"
        ws_summary['B7'] = profile.get('price', 0)
        
        ws_summary['A9'] = "Key Metrics"
        ws_summary['A9'].font = self.header_font
        ws_summary['A9'].fill = self.header_fill
        ws_summary.merge_cells('A9:B9')
        
        key_metrics = data.get('key_metrics', {})
        ev_data = data.get('enterprise_value', {})
        
        ws_summary['A10'] = "Market Cap:"
        ws_summary['B10'] = ev_data.get('marketCapitalization', 0)
        ws_summary['A11'] = "Enterprise Value:"
        ws_summary['B11'] = ev_data.get('enterpriseValue', 0)
        ws_summary['A12'] = "Current EV/EBITDA:"
        ws_summary['B12'] = key_metrics.get('enterpriseValueMultipleTTM', 0)
        
        ws_summary['A14'] = "Model Inputs"
        ws_summary['A14'].font = self.header_font
        ws_summary['A14'].fill = self.header_fill
        ws_summary.merge_cells('A14:B14')
        
        financial_statements = data.get('financial_statements', {})
        income_statement = financial_statements.get('income_statement', [])
        ebitda = 0
        if income_statement:
            latest = income_statement[0]
            ebitda = latest.get('EBITDA', 0) or latest.get('ebitda', 0) or latest.get('operatingIncome', 0) or latest.get('Operating Income', 0)
        
        shares_outstanding = ev_data.get('marketCapitalization', 0) / profile.get('price', 1) if profile.get('price', 0) > 0 else 0
        net_debt = ev_data.get('totalDebt', 0) - ev_data.get('cashAndCashEquivalents', 0)
        current_price = profile.get('price', 0)
        
        ws_summary['A15'] = "EBITDA (TTM):"
        ws_summary['B15'] = ebitda
        ws_summary['A16'] = "Net Debt:"
        ws_summary['B16'] = net_debt
        ws_summary['A17'] = "Shares Outstanding:"
        ws_summary['B17'] = shares_outstanding
        ws_summary['A18'] = "Current Price:"
        ws_summary['B18'] = current_price
        
        ws_summary['A20'] = "Growth Assumptions"
        ws_summary['A20'].font = self.header_font
        ws_summary['A20'].fill = self.header_fill
        ws_summary.merge_cells('A20:B20')
        
        ws_summary['A21'] = "EBITDA Growth Rate (Annual):"
        ws_summary['B21'] = 0.08
        ws_summary['B21'].number_format = '0.0%'
        
        ws_summary['A23'] = "5-Year Exit Valuation Analysis"
        ws_summary['A23'].font = self.header_font
        ws_summary['A23'].fill = self.header_fill
        ws_summary.merge_cells('A23:G23')
        
        headers = ['Scenario', 'Exit Multiple', 'Year 5 EBITDA', 'Exit EV', 'Exit Equity Value', 'Exit Price', 'Total Return']
        for i, header in enumerate(headers, 1):
            cell = ws_summary.cell(row=24, column=i, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.border = self.border
            cell.alignment = self.center_alignment
        
        scenarios = [
            ('Bear Case', 15.0),
            ('Base Case', 18.0),
            ('Bull Case', 22.0),
            ('Current Multiple', key_metrics.get('enterpriseValueMultipleTTM', 18.0)),
            ('Test Scenario', 19.5)
        ]
        
        for i, (scenario, multiple) in enumerate(scenarios, 25):
            ws_summary.cell(row=i, column=1, value=scenario)
            ws_summary.cell(row=i, column=2, value=multiple)
            
            ws_summary.cell(row=i, column=3, value=f'=$B$15*(1+$B$21)^5')
            ws_summary.cell(row=i, column=4, value=f'=C{i}*B{i}')
            ws_summary.cell(row=i, column=5, value=f'=D{i}-$B$16')
            ws_summary.cell(row=i, column=6, value=f'=E{i}/$B$17')
            ws_summary.cell(row=i, column=7, value=f'=(F{i}/$B$18)^(1/5)-1')
            ws_summary.cell(row=i, column=7).number_format = '0.0%'
        
        ws_sensitivity = wb.create_sheet("Sensitivity Analysis")
        
        ws_sensitivity['A1'] = "5-Year Exit EV/EBITDA Sensitivity Analysis"
        ws_sensitivity['A1'].font = Font(size=14, bold=True)
        ws_sensitivity.merge_cells('A1:K1')
        
        multiples = [14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24]
        ebitda_growth_rates = [2, 4, 6, 8, 10, 12, 15]
        
        ws_sensitivity['A3'] = "EBITDA Growth % (Annual)"
        for i, multiple in enumerate(multiples, 2):
            ws_sensitivity.cell(row=2, column=i, value=f"{multiple}x")
            ws_sensitivity.cell(row=2, column=i).font = self.header_font
            ws_sensitivity.cell(row=2, column=i).fill = self.header_fill
        
        for i, growth in enumerate(ebitda_growth_rates, 3):
            ws_sensitivity.cell(row=i, column=1, value=f"{growth}%")
            ws_sensitivity.cell(row=i, column=1).font = self.header_font
            ws_sensitivity.cell(row=i, column=1).fill = self.header_fill
            
            for j, multiple in enumerate(multiples, 2):
                col_letter = chr(ord('A') + j)
                growth_decimal = growth / 100
                formula = f'=((Summary.$B$14*(1+{growth_decimal})^5*{multiple}-Summary.$B$15)/Summary.$B$16/Summary.$B$17)^(1/5)-1'
                ws_sensitivity.cell(row=i, column=j, value=formula)
                ws_sensitivity.cell(row=i, column=j).number_format = '0.0%'
        
        sensitivity_range = f"B3:L{2+len(ebitda_growth_rates)}"
        rule = ColorScaleRule(start_type='min', start_color='FF6B6B',
                             mid_type='percentile', mid_value=50, mid_color='FFEB3B',
                             end_type='max', end_color='4CAF50')
        ws_sensitivity.conditional_formatting.add(sensitivity_range, rule)
        
        ws_data = wb.create_sheet("Financial Data")
        
        if income_statement:
            ws_data['A1'] = "Income Statement (Latest)"
            ws_data['A1'].font = Font(size=12, bold=True)
            
            row = 3
            for key, value in income_statement[0].items():
                ws_data.cell(row=row, column=1, value=key)
                ws_data.cell(row=row, column=2, value=value)
                row += 1
        
        for ws in wb.worksheets:
            for column_cells in ws.columns:
                max_length = 0
                column_letter = None
                for cell in column_cells:
                    if hasattr(cell, 'column_letter'):
                        column_letter = cell.column_letter
                        try:
                            if cell.value and len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                
                if column_letter:
                    adjusted_width = min(max_length + 2, 50)
                    ws.column_dimensions[column_letter].width = adjusted_width
        
        wb.save(output_path)
        print(f"EV/EBITDA template saved to {output_path}")
    
    def create_ev_sales_template(self, symbol: str, data: Dict, output_path: str):
        """Create EV/Sales valuation template"""
        
        wb = Workbook()
        ws = wb.active
        ws.title = "EV Sales Analysis"
        
        profile = data.get('profile', {})
        key_metrics = data.get('key_metrics', {})
        ev_data = data.get('enterprise_value', {})
        financial_statements = data.get('financial_statements', {})
        
        ws['A1'] = "EV/Sales Valuation Model"
        ws['A1'].font = Font(size=16, bold=True)
        ws.merge_cells('A1:F1')
        
        ws['A2'] = "Ticker Symbol:"
        ws['B2'] = symbol
        ws['B2'].font = Font(size=14, bold=True, color="FF0000")
        ws['B2'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        ws['B2'].border = Border(
            left=Side(style='thick'),
            right=Side(style='thick'),
            top=Side(style='thick'),
            bottom=Side(style='thick')
        )
        
        ws['D2'] = "To analyze a new symbol:"
        ws['E2'] = "1. Change symbol in B2"
        ws['F2'] = "2. Run: python examples/analyze_stock.py [NEW_SYMBOL]"
        ws['D2'].font = Font(size=10, italic=True)
        ws['E2'].font = Font(size=10, italic=True)
        ws['F2'].font = Font(size=10, italic=True)
        
        income_statement = financial_statements.get('income_statement', [])
        revenue = 0
        if income_statement:
            latest = income_statement[0]
            revenue = latest.get('Total Revenue', 0) or latest.get('revenue', 0) or latest.get('totalRevenue', 0)
        
        shares_outstanding = ev_data.get('marketCapitalization', 0) / profile.get('price', 1) if profile.get('price', 0) > 0 else 0
        net_debt = ev_data.get('totalDebt', 0) - ev_data.get('cashAndCashEquivalents', 0)
        current_price = profile.get('price', 0)
        
        ws['A3'] = "Model Inputs"
        ws['A3'].font = self.header_font
        ws['A3'].fill = self.header_fill
        ws.merge_cells('A3:B3')
        
        ws['A4'] = "Revenue (TTM):"
        ws['B4'] = revenue
        ws['A5'] = "Net Debt:"
        ws['B5'] = net_debt
        ws['A6'] = "Shares Outstanding:"
        ws['B6'] = shares_outstanding
        ws['A7'] = "Current Price:"
        ws['B7'] = current_price
        
        ws['A9'] = "Growth Assumptions"
        ws['A9'].font = self.header_font
        ws['A9'].fill = self.header_fill
        ws.merge_cells('A9:B9')
        
        ws['A10'] = "Revenue Growth Rate (Annual):"
        ws['B10'] = 0.10
        ws['B10'].number_format = '0.0%'
        
        scenarios = [
            ('Conservative', 2.0),
            ('Market Average', 3.5),
            ('Premium', 5.0),
            ('Current Multiple', key_metrics.get('evToSalesTTM', 3.5)),
            ('Target Multiple', 4.2)
        ]
        
        ws['A12'] = "5-Year Exit Valuation Analysis"
        ws['A12'].font = self.header_font
        ws['A12'].fill = self.header_fill
        ws.merge_cells('A12:G12')
        
        headers = ['Scenario', 'Exit Multiple', 'Year 5 Revenue', 'Exit EV', 'Exit Equity Value', 'Exit Price', 'Total Return']
        for i, header in enumerate(headers, 1):
            cell = ws.cell(row=13, column=i, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
        
        for i, (scenario, multiple) in enumerate(scenarios, 14):
            ws.cell(row=i, column=1, value=scenario)
            ws.cell(row=i, column=2, value=multiple)
            
            ws.cell(row=i, column=3, value=f'=$B$4*(1+$B$10)^5')
            ws.cell(row=i, column=4, value=f'=C{i}*B{i}')
            ws.cell(row=i, column=5, value=f'=D{i}-$B$5')
            ws.cell(row=i, column=6, value=f'=E{i}/$B$6')
            ws.cell(row=i, column=7, value=f'=(F{i}/$B$7)^(1/5)-1')
            ws.cell(row=i, column=7).number_format = '0.0%'
        
        wb.save(output_path)
        print(f"EV/Sales template saved to {output_path}")
    
    def create_dcf_template(self, symbol: str, data: Dict, output_path: str):
        """Create DCF valuation template with detailed cash flow projections"""
        
        wb = Workbook()
        
        ws_dcf = wb.active
        ws_dcf.title = "DCF Model"
        
        ws_dcf['A1'] = "Discounted Cash Flow Model"
        ws_dcf['A1'].font = Font(size=16, bold=True)
        ws_dcf.merge_cells('A1:H1')
        
        ws_dcf['A2'] = "Ticker Symbol:"
        ws_dcf['B2'] = symbol
        ws_dcf['B2'].font = Font(size=14, bold=True, color="FF0000")
        ws_dcf['B2'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        ws_dcf['B2'].border = Border(
            left=Side(style='thick'),
            right=Side(style='thick'),
            top=Side(style='thick'),
            bottom=Side(style='thick')
        )
        
        ws_dcf['D2'] = "To analyze a new symbol:"
        ws_dcf['E2'] = "1. Change symbol in B2"
        ws_dcf['F2'] = "2. Run: python examples/analyze_stock.py [NEW_SYMBOL]"
        ws_dcf['D2'].font = Font(size=10, italic=True)
        ws_dcf['E2'].font = Font(size=10, italic=True)
        ws_dcf['F2'].font = Font(size=10, italic=True)
        
        ws_dcf['A3'] = "Data Source: Yahoo Finance & Financial Modeling Prep"
        ws_dcf['A3'].font = Font(size=9, italic=True, color="666666")
        
        profile = data.get('profile', {})
        if not profile.get('companyName'):
            ws_dcf['B4'] = f"ERROR: No data found for {symbol}"
            ws_dcf['B4'].font = Font(color="FF0000", bold=True)
        else:
            ws_dcf['A4'] = "Company Name:"
            ws_dcf['B4'] = profile.get('companyName', '')
        
        ws_dcf['A6'] = "Key Assumptions"
        ws_dcf['A6'].font = self.header_font
        ws_dcf['A6'].fill = self.header_fill
        ws_dcf.merge_cells('A6:B6')
        
        assumptions = [
            ('Revenue Growth Rate (Years 1-5)', '10%'),
            ('Terminal Growth Rate', '3%'),
            ('WACC (Discount Rate)', '10%'),
            ('Tax Rate', '25%'),
            ('EBITDA Margin', '20%'),
            ('Capex as % of Revenue', '3%'),
            ('Working Capital Change', '1%')
        ]
        
        for i, (assumption, default_value) in enumerate(assumptions, 7):
            ws_dcf.cell(row=i, column=1, value=assumption)
            ws_dcf.cell(row=i, column=2, value=default_value)
        
        years = ['Current', 'Year 1', 'Year 2', 'Year 3', 'Year 4', 'Year 5', 'Terminal']
        for i, year in enumerate(years, 1):
            cell = ws_dcf.cell(row=15, column=i, value=year)
            cell.font = self.header_font
            cell.fill = self.header_fill
        
        financial_statements = data.get('financial_statements', {})
        income_statement = financial_statements.get('income_statement', [])
        
        base_revenue = 0
        if income_statement:
            latest = income_statement[0]
            revenue_keys = ['Total Revenue', 'revenue', 'totalRevenue', 'Net Sales', 'Sales']
            for key in revenue_keys:
                if key in latest and latest[key] is not None:
                    base_revenue = latest[key]
                    break
        
        ws_dcf['A16'] = "Revenue"
        ws_dcf['B16'] = base_revenue
        
        for year in range(2, 7):  # Years 1-5
            col_letter = chr(ord('B') + year - 1)
            prev_col_letter = chr(ord('B') + year - 2)
            ws_dcf[f'{col_letter}16'] = f'={prev_col_letter}16*(1+$B$7)'
        
        ws_dcf['H16'] = '=G16*(1+$B$8)'
        
        ws_dcf['A17'] = "EBITDA"
        for year in range(1, 8):  # Current through Terminal
            col_letter = chr(ord('A') + year)
            ws_dcf[f'{col_letter}17'] = f'={col_letter}16*$B$10'
        
        ws_dcf['A18'] = "EBIT"
        ws_dcf['A19'] = "Tax"
        ws_dcf['A20'] = "NOPAT"
        ws_dcf['A21'] = "Capex"
        ws_dcf['A22'] = "Working Capital Change"
        ws_dcf['A23'] = "Free Cash Flow"
        
        for year in range(1, 8):
            col_letter = chr(ord('A') + year)
            ws_dcf[f'{col_letter}18'] = f'={col_letter}17'
            ws_dcf[f'{col_letter}19'] = f'={col_letter}18*$B$9'
            ws_dcf[f'{col_letter}20'] = f'={col_letter}18-{col_letter}19'
            ws_dcf[f'{col_letter}21'] = f'={col_letter}16*$B$11'
            ws_dcf[f'{col_letter}22'] = f'={col_letter}16*$B$12'
            ws_dcf[f'{col_letter}23'] = f'={col_letter}20-{col_letter}21-{col_letter}22'
        
        ws_dcf['A25'] = "Valuation"
        ws_dcf['A25'].font = self.header_font
        ws_dcf['A25'].fill = self.header_fill
        
        ws_dcf['A26'] = "Terminal Value"
        ws_dcf['B26'] = '=H23/($B$8-$B$10)'
        
        ws_dcf['A27'] = "PV of Terminal Value"
        ws_dcf['B27'] = '=B26/((1+$B$10)^5)'
        
        ws_dcf['A28'] = "Sum of PV of FCF (Years 1-5)"
        ws_dcf['B28'] = '=NPV($B$10,C23:G23)'
        
        ws_dcf['A29'] = "Enterprise Value"
        ws_dcf['B29'] = '=B27+B28'
        
        ev_data = data.get('enterprise_value', {})
        net_debt = ev_data.get('totalDebt', 0) - ev_data.get('cashAndCashEquivalents', 0)
        shares_outstanding = ev_data.get('marketCapitalization', 0) / data.get('profile', {}).get('price', 1)
        
        ws_dcf['A30'] = "Less: Net Debt"
        ws_dcf['B30'] = net_debt
        
        ws_dcf['A31'] = "Equity Value"
        ws_dcf['B31'] = '=B29-B30'
        
        ws_dcf['A32'] = "Shares Outstanding"
        ws_dcf['B32'] = shares_outstanding
        
        ws_dcf['A33'] = "DCF Price per Share"
        ws_dcf['B33'] = '=B31/B32'
        
        ws_dcf['A34'] = "Current Price"
        ws_dcf['B34'] = data.get('profile', {}).get('price', 0)
        
        ws_dcf['A35'] = "Upside/Downside"
        ws_dcf['B35'] = '=(B33-B34)/B34'
        ws_dcf['B35'].number_format = '0.0%'
        
        ws_sensitivity = wb.create_sheet("DCF Sensitivity")
        
        ws_sensitivity['A1'] = "DCF Sensitivity Analysis"
        ws_sensitivity['A1'].font = Font(size=14, bold=True)
        
        wacc_rates = [8, 9, 10, 11, 12, 13, 14]
        terminal_rates = [1, 2, 3, 4, 5]
        
        ws_sensitivity['A3'] = "Terminal Growth Rate"
        for i, wacc in enumerate(wacc_rates, 2):
            ws_sensitivity.cell(row=2, column=i, value=f"{wacc}%")
            ws_sensitivity.cell(row=2, column=i).font = self.header_font
            ws_sensitivity.cell(row=2, column=i).fill = self.header_fill
        
        for i, terminal in enumerate(terminal_rates, 3):
            ws_sensitivity.cell(row=i, column=1, value=f"{terminal}%")
            ws_sensitivity.cell(row=i, column=1).font = self.header_font
            ws_sensitivity.cell(row=i, column=1).fill = self.header_fill
            
            for j, wacc in enumerate(wacc_rates, 2):
                terminal_decimal = terminal / 100
                wacc_decimal = wacc / 100
                formula = f'=((DCF Model.H23/{terminal_decimal}-{wacc_decimal})/((1+{wacc_decimal})^5)+NPV({wacc_decimal},DCF Model.C23:G23)-DCF Model.B30)/DCF Model.B32'
                ws_sensitivity.cell(row=i, column=j, value=formula)
                ws_sensitivity.cell(row=i, column=j).number_format = '$0.00'
        
        sensitivity_range = f"B3:H{2+len(terminal_rates)}"
        rule = ColorScaleRule(start_type='min', start_color='FF6B6B',
                             mid_type='percentile', mid_value=50, mid_color='FFEB3B',
                             end_type='max', end_color='4CAF50')
        ws_sensitivity.conditional_formatting.add(sensitivity_range, rule)
        
        wb.save(output_path)
        print(f"DCF template saved to {output_path}")
