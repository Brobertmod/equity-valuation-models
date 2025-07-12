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
        
        ws_summary['A1'] = f"{symbol} - EV/EBITDA Valuation Model"
        ws_summary['A1'].font = Font(size=16, bold=True)
        ws_summary.merge_cells('A1:F1')
        
        profile = data.get('profile', {})
        ws_summary['A3'] = "Company Name:"
        ws_summary['B3'] = profile.get('companyName', '')
        ws_summary['A4'] = "Sector:"
        ws_summary['B4'] = profile.get('sector', '')
        ws_summary['A5'] = "Industry:"
        ws_summary['B5'] = profile.get('industry', '')
        ws_summary['A6'] = "Current Price:"
        ws_summary['B6'] = profile.get('price', 0)
        
        ws_summary['A8'] = "Key Metrics"
        ws_summary['A8'].font = self.header_font
        ws_summary['A8'].fill = self.header_fill
        ws_summary.merge_cells('A8:B8')
        
        key_metrics = data.get('key_metrics', {})
        ev_data = data.get('enterprise_value', {})
        
        ws_summary['A9'] = "Market Cap:"
        ws_summary['B9'] = ev_data.get('marketCapitalization', 0)
        ws_summary['A10'] = "Enterprise Value:"
        ws_summary['B10'] = ev_data.get('enterpriseValue', 0)
        ws_summary['A11'] = "Current EV/EBITDA:"
        ws_summary['B11'] = key_metrics.get('enterpriseValueMultipleTTM', 0)
        
        ws_summary['A13'] = "Valuation Analysis"
        ws_summary['A13'].font = self.header_font
        ws_summary['A13'].fill = self.header_fill
        ws_summary.merge_cells('A13:F13')
        
        headers = ['Scenario', 'EV/EBITDA Multiple', 'Target EV', 'Target Equity Value', 'Target Price', 'Upside/Downside']
        for i, header in enumerate(headers, 1):
            cell = ws_summary.cell(row=14, column=i, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.border = self.border
            cell.alignment = self.center_alignment
        
        financial_statements = data.get('financial_statements', {})
        income_statement = financial_statements.get('income_statement', [])
        ebitda = 0
        if income_statement:
            latest = income_statement[0]
            ebitda = latest.get('ebitda', 0) or latest.get('operatingIncome', 0)
        
        scenarios = [
            ('Bear Case', 15.0),
            ('Base Case', 18.0),
            ('Bull Case', 22.0),
            ('Current Multiple', key_metrics.get('enterpriseValueMultipleTTM', 18.0)),
            ('Test Scenario', 19.5)  # User's example scenario
        ]
        
        shares_outstanding = ev_data.get('marketCapitalization', 0) / profile.get('price', 1) if profile.get('price', 0) > 0 else 0
        net_debt = ev_data.get('totalDebt', 0) - ev_data.get('cashAndCashEquivalents', 0)
        current_price = profile.get('price', 0)
        
        for i, (scenario, multiple) in enumerate(scenarios, 15):
            ws_summary.cell(row=i, column=1, value=scenario)
            ws_summary.cell(row=i, column=2, value=multiple)
            
            target_ev = ebitda * multiple
            ws_summary.cell(row=i, column=3, value=target_ev)
            
            target_equity = target_ev - net_debt
            ws_summary.cell(row=i, column=4, value=target_equity)
            
            target_price = target_equity / shares_outstanding if shares_outstanding > 0 else 0
            ws_summary.cell(row=i, column=5, value=target_price)
            
            upside = (target_price - current_price) / current_price if current_price > 0 else 0
            ws_summary.cell(row=i, column=6, value=upside)
            ws_summary.cell(row=i, column=6).number_format = '0.0%'
        
        ws_sensitivity = wb.create_sheet("Sensitivity Analysis")
        
        ws_sensitivity['A1'] = "EV/EBITDA Sensitivity Analysis"
        ws_sensitivity['A1'].font = Font(size=14, bold=True)
        ws_sensitivity.merge_cells('A1:K1')
        
        multiples = [14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24]
        ebitda_growth_rates = [-10, -5, 0, 5, 10, 15, 20]
        
        ws_sensitivity['A3'] = "EBITDA Growth %"
        for i, multiple in enumerate(multiples, 2):
            ws_sensitivity.cell(row=2, column=i, value=f"{multiple}x")
            ws_sensitivity.cell(row=2, column=i).font = self.header_font
            ws_sensitivity.cell(row=2, column=i).fill = self.header_fill
        
        for i, growth in enumerate(ebitda_growth_rates, 3):
            ws_sensitivity.cell(row=i, column=1, value=f"{growth}%")
            ws_sensitivity.cell(row=i, column=1).font = self.header_font
            ws_sensitivity.cell(row=i, column=1).fill = self.header_fill
            
            for j, multiple in enumerate(multiples, 2):
                adjusted_ebitda = ebitda * (1 + growth/100)
                target_ev = adjusted_ebitda * multiple
                target_equity = target_ev - net_debt
                target_price = target_equity / shares_outstanding if shares_outstanding > 0 else 0
                
                ws_sensitivity.cell(row=i, column=j, value=target_price)
                ws_sensitivity.cell(row=i, column=j).number_format = '$0.00'
        
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
        
        ws['A1'] = f"{symbol} - EV/Sales Valuation Model"
        ws['A1'].font = Font(size=16, bold=True)
        ws.merge_cells('A1:F1')
        
        income_statement = financial_statements.get('income_statement', [])
        revenue = 0
        if income_statement:
            latest = income_statement[0]
            revenue = latest.get('revenue', 0) or latest.get('totalRevenue', 0)
        
        scenarios = [
            ('Conservative', 2.0),
            ('Market Average', 3.5),
            ('Premium', 5.0),
            ('Current Multiple', key_metrics.get('evToSalesTTM', 3.5)),
            ('Target Multiple', 4.2)
        ]
        
        headers = ['Scenario', 'EV/Sales Multiple', 'Target EV', 'Target Equity Value', 'Target Price', 'Upside/Downside']
        for i, header in enumerate(headers, 1):
            cell = ws.cell(row=5, column=i, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
        
        shares_outstanding = ev_data.get('marketCapitalization', 0) / profile.get('price', 1) if profile.get('price', 0) > 0 else 0
        net_debt = ev_data.get('totalDebt', 0) - ev_data.get('cashAndCashEquivalents', 0)
        current_price = profile.get('price', 0)
        
        for i, (scenario, multiple) in enumerate(scenarios, 6):
            ws.cell(row=i, column=1, value=scenario)
            ws.cell(row=i, column=2, value=multiple)
            
            target_ev = revenue * multiple
            ws.cell(row=i, column=3, value=target_ev)
            
            target_equity = target_ev - net_debt
            ws.cell(row=i, column=4, value=target_equity)
            
            target_price = target_equity / shares_outstanding if shares_outstanding > 0 else 0
            ws.cell(row=i, column=5, value=target_price)
            
            upside = (target_price - current_price) / current_price if current_price > 0 else 0
            ws.cell(row=i, column=6, value=upside)
            ws.cell(row=i, column=6).number_format = '0.0%'
        
        wb.save(output_path)
        print(f"EV/Sales template saved to {output_path}")
    
    def create_dcf_template(self, symbol: str, data: Dict, output_path: str):
        """Create DCF valuation template with detailed cash flow projections"""
        
        wb = Workbook()
        
        ws_dcf = wb.active
        ws_dcf.title = "DCF Model"
        
        ws_dcf['A1'] = f"{symbol} - Discounted Cash Flow Model"
        ws_dcf['A1'].font = Font(size=16, bold=True)
        ws_dcf.merge_cells('A1:H1')
        
        ws_dcf['A3'] = "Key Assumptions"
        ws_dcf['A3'].font = self.header_font
        ws_dcf['A3'].fill = self.header_fill
        ws_dcf.merge_cells('A3:B3')
        
        assumptions = [
            ('Revenue Growth Rate (Years 1-5)', '10%'),
            ('Terminal Growth Rate', '3%'),
            ('WACC (Discount Rate)', '10%'),
            ('Tax Rate', '25%'),
            ('EBITDA Margin', '20%'),
            ('Capex as % of Revenue', '3%'),
            ('Working Capital Change', '1%')
        ]
        
        for i, (assumption, default_value) in enumerate(assumptions, 4):
            ws_dcf.cell(row=i, column=1, value=assumption)
            ws_dcf.cell(row=i, column=2, value=default_value)
        
        years = ['Current', 'Year 1', 'Year 2', 'Year 3', 'Year 4', 'Year 5', 'Terminal']
        for i, year in enumerate(years, 1):
            cell = ws_dcf.cell(row=12, column=i, value=year)
            cell.font = self.header_font
            cell.fill = self.header_fill
        
        financial_statements = data.get('financial_statements', {})
        income_statement = financial_statements.get('income_statement', [])
        
        base_revenue = 0
        if income_statement:
            latest = income_statement[0]
            base_revenue = latest.get('revenue', 0) or latest.get('totalRevenue', 0)
        
        ws_dcf['A13'] = "Revenue"
        ws_dcf['B13'] = base_revenue
        
        for year in range(2, 7):  # Years 1-5
            col_letter = chr(ord('B') + year - 1)
            prev_col_letter = chr(ord('B') + year - 2)
            ws_dcf[f'{col_letter}13'] = f'={prev_col_letter}13*(1+$B$4)'
        
        ws_dcf['H13'] = '=G13*(1+$B$5)'
        
        ws_dcf['A14'] = "EBITDA"
        for year in range(1, 8):  # Current through Terminal
            col_letter = chr(ord('A') + year)
            ws_dcf[f'{col_letter}14'] = f'={col_letter}13*$B$7'
        
        ws_dcf['A15'] = "EBIT"
        ws_dcf['A16'] = "Tax"
        ws_dcf['A17'] = "NOPAT"
        ws_dcf['A18'] = "Capex"
        ws_dcf['A19'] = "Working Capital Change"
        ws_dcf['A20'] = "Free Cash Flow"
        
        for year in range(1, 8):
            col_letter = chr(ord('A') + year)
            ws_dcf[f'{col_letter}15'] = f'={col_letter}14'
            ws_dcf[f'{col_letter}16'] = f'={col_letter}15*$B$6'
            ws_dcf[f'{col_letter}17'] = f'={col_letter}15-{col_letter}16'
            ws_dcf[f'{col_letter}18'] = f'={col_letter}13*$B$8'
            ws_dcf[f'{col_letter}19'] = f'={col_letter}13*$B$9'
            ws_dcf[f'{col_letter}20'] = f'={col_letter}17-{col_letter}18-{col_letter}19'
        
        ws_dcf['A22'] = "Valuation"
        ws_dcf['A22'].font = self.header_font
        ws_dcf['A22'].fill = self.header_fill
        
        ws_dcf['A23'] = "Terminal Value"
        ws_dcf['B23'] = '=H20/(B$5-B$6)'  # Terminal FCF / (Terminal Growth - WACC)
        
        ws_dcf['A24'] = "PV of Terminal Value"
        ws_dcf['B24'] = '=B23/((1+$B$6)^5)'  # Discount terminal value
        
        ws_dcf['A25'] = "Sum of PV of FCF (Years 1-5)"
        ws_dcf['B25'] = '=NPV($B$6,C20:G20)'  # NPV of explicit forecast period
        
        ws_dcf['A26'] = "Enterprise Value"
        ws_dcf['B26'] = '=B24+B25'
        
        ev_data = data.get('enterprise_value', {})
        net_debt = ev_data.get('totalDebt', 0) - ev_data.get('cashAndCashEquivalents', 0)
        shares_outstanding = ev_data.get('marketCapitalization', 0) / data.get('profile', {}).get('price', 1)
        
        ws_dcf['A27'] = "Less: Net Debt"
        ws_dcf['B27'] = net_debt
        
        ws_dcf['A28'] = "Equity Value"
        ws_dcf['B28'] = '=B26-B27'
        
        ws_dcf['A29'] = "Shares Outstanding"
        ws_dcf['B29'] = shares_outstanding
        
        ws_dcf['A30'] = "DCF Price per Share"
        ws_dcf['B30'] = '=B28/B29'
        
        ws_dcf['A31'] = "Current Price"
        ws_dcf['B31'] = data.get('profile', {}).get('price', 0)
        
        ws_dcf['A32'] = "Upside/Downside"
        ws_dcf['B32'] = '=(B30-B31)/B31'
        ws_dcf['B32'].number_format = '0.0%'
        
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
        
        wb.save(output_path)
        print(f"DCF template saved to {output_path}")
