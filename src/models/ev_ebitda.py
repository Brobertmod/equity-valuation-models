"""
EV/EBITDA Valuation Model
"""

from typing import Dict, List, Tuple
import pandas as pd
from ..data_fetcher import FinancialDataFetcher
from ..excel_generator import ExcelTemplateGenerator

class EVEBITDAModel:
    """
    Enterprise Value to EBITDA multiple valuation model
    """
    
    def __init__(self, symbol: str):
        self.symbol = symbol
        self.data_fetcher = FinancialDataFetcher()
        self.excel_generator = ExcelTemplateGenerator()
        self.data = {}
        
    def load_data(self):
        """Load comprehensive financial data for the symbol"""
        self.data = self.data_fetcher.get_comprehensive_data(self.symbol)
        return self.data
    
    def calculate_valuation(self, multiple: float = None) -> Dict:
        """
        Calculate valuation based on EV/EBITDA multiple
        
        Args:
            multiple: EV/EBITDA multiple to use. If None, uses current market multiple
            
        Returns:
            Dictionary with valuation results
        """
        if not self.data:
            self.load_data()
            
        financial_statements = self.data.get('financial_statements', {})
        income_statement = financial_statements.get('income_statement', [])
        key_metrics = self.data.get('key_metrics', {})
        ev_data = self.data.get('enterprise_value', {})
        profile = self.data.get('profile', {})
        
        ebitda = 0
        if income_statement:
            latest = income_statement[0]
            ebitda = latest.get('ebitda', 0) or latest.get('operatingIncome', 0)
        
        if multiple is None:
            multiple = key_metrics.get('enterpriseValueMultipleTTM', 18.0)
        
        target_ev = ebitda * multiple
        net_debt = ev_data.get('totalDebt', 0) - ev_data.get('cashAndCashEquivalents', 0)
        target_equity_value = target_ev - net_debt
        
        current_price = profile.get('price', 0)
        shares_outstanding = ev_data.get('marketCapitalization', 0) / current_price if current_price > 0 else 0
        target_price = target_equity_value / shares_outstanding if shares_outstanding > 0 else 0
        
        upside = (target_price - current_price) / current_price if current_price > 0 else 0
        
        return {
            'symbol': self.symbol,
            'current_price': current_price,
            'target_price': target_price,
            'upside_downside': upside,
            'ev_ebitda_multiple': multiple,
            'ebitda': ebitda,
            'target_ev': target_ev,
            'target_equity_value': target_equity_value,
            'net_debt': net_debt,
            'shares_outstanding': shares_outstanding
        }
    
    def sensitivity_analysis(self, multiple_range: Tuple[float, float] = (15, 25), 
                           ebitda_growth_range: Tuple[float, float] = (-10, 20)) -> pd.DataFrame:
        """
        Perform sensitivity analysis across multiple scenarios
        
        Args:
            multiple_range: Tuple of (min_multiple, max_multiple)
            ebitda_growth_range: Tuple of (min_growth_%, max_growth_%)
            
        Returns:
            DataFrame with sensitivity analysis results
        """
        if not self.data:
            self.load_data()
            
        multiples = [x for x in range(int(multiple_range[0]), int(multiple_range[1]) + 1)]
        growth_rates = [x for x in range(int(ebitda_growth_range[0]), int(ebitda_growth_range[1]) + 1, 5)]
        
        financial_statements = self.data.get('financial_statements', {})
        income_statement = financial_statements.get('income_statement', [])
        base_ebitda = 0
        if income_statement:
            latest = income_statement[0]
            base_ebitda = latest.get('ebitda', 0) or latest.get('operatingIncome', 0)
        
        results = []
        for growth in growth_rates:
            row = {'EBITDA_Growth': f"{growth}%"}
            adjusted_ebitda = base_ebitda * (1 + growth/100)
            
            for multiple in multiples:
                ev_data = self.data.get('enterprise_value', {})
                profile = self.data.get('profile', {})
                
                target_ev = adjusted_ebitda * multiple
                net_debt = ev_data.get('totalDebt', 0) - ev_data.get('cashAndCashEquivalents', 0)
                target_equity = target_ev - net_debt
                
                current_price = profile.get('price', 0)
                shares_outstanding = ev_data.get('marketCapitalization', 0) / current_price if current_price > 0 else 0
                target_price = target_equity / shares_outstanding if shares_outstanding > 0 else 0
                
                row[f"{multiple}x"] = target_price
            
            results.append(row)
        
        return pd.DataFrame(results)
    
    def export_to_excel(self, output_path: str):
        """Export model to Excel template"""
        if not self.data:
            self.load_data()
            
        self.excel_generator.create_ev_ebitda_template(self.symbol, self.data, output_path)
        return output_path
