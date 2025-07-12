import requests
import pandas as pd
import yfinance as yf
from typing import Dict, List, Optional, Union
import os
from dotenv import load_dotenv
import yaml

load_dotenv()

class FinancialDataFetcher:
    """
    Unified financial data fetcher supporting multiple data sources:
    - Financial Modeling Prep API
    - Yahoo Finance
    - Alpha Vantage
    - Custom subscription sites
    """
    
    def __init__(self, config_path: str = "config/api_keys.yaml"):
        self.config = self._load_config(config_path)
        self.fmp_api_key = self.config.get('fmp_api_key') or os.getenv('FMP_API_KEY')
        self.alpha_vantage_key = self.config.get('alpha_vantage_key') or os.getenv('ALPHA_VANTAGE_KEY')
        
    def _load_config(self, config_path: str) -> Dict:
        """Load API configuration from YAML file"""
        try:
            with open(config_path, 'r') as file:
                return yaml.safe_load(file) or {}
        except FileNotFoundError:
            return {}
    
    def get_company_profile(self, symbol: str) -> Dict:
        """Get comprehensive company profile data"""
        try:
            if self.fmp_api_key:
                url = f"https://financialmodelingprep.com/api/v3/profile/{symbol}?apikey={self.fmp_api_key}"
                response = requests.get(url)
                if response.status_code == 200:
                    data = response.json()
                    if data:
                        return data[0]
            
            ticker = yf.Ticker(symbol)
            info = ticker.info
            return {
                'symbol': symbol,
                'companyName': info.get('longName', ''),
                'industry': info.get('industry', ''),
                'sector': info.get('sector', ''),
                'marketCap': info.get('marketCap', 0),
                'price': info.get('currentPrice', 0),
                'beta': info.get('beta', 1.0),
                'description': info.get('longBusinessSummary', '')
            }
        except Exception as e:
            print(f"Error fetching company profile for {symbol}: {e}")
            return {}
    
    def get_financial_statements(self, symbol: str, period: str = "annual") -> Dict:
        """
        Get income statement, balance sheet, and cash flow statement
        period: 'annual' or 'quarter'
        """
        statements = {}
        
        try:
            if self.fmp_api_key:
                url = f"https://financialmodelingprep.com/api/v3/income-statement/{symbol}?period={period}&apikey={self.fmp_api_key}"
                response = requests.get(url)
                if response.status_code == 200:
                    statements['income_statement'] = response.json()
                
                url = f"https://financialmodelingprep.com/api/v3/balance-sheet-statement/{symbol}?period={period}&apikey={self.fmp_api_key}"
                response = requests.get(url)
                if response.status_code == 200:
                    statements['balance_sheet'] = response.json()
                
                url = f"https://financialmodelingprep.com/api/v3/cash-flow-statement/{symbol}?period={period}&apikey={self.fmp_api_key}"
                response = requests.get(url)
                if response.status_code == 200:
                    statements['cash_flow'] = response.json()
            
            if not statements:
                ticker = yf.Ticker(symbol)
                
                financials = ticker.financials
                balance_sheet = ticker.balance_sheet
                cashflow = ticker.cashflow
                
                if not financials.empty:
                    financials_transposed = financials.T  # Transpose so dates are rows
                    statements['income_statement'] = financials_transposed.to_dict('records')
                
                if not balance_sheet.empty:
                    balance_sheet_transposed = balance_sheet.T
                    statements['balance_sheet'] = balance_sheet_transposed.to_dict('records')
                
                if not cashflow.empty:
                    cashflow_transposed = cashflow.T
                    statements['cash_flow'] = cashflow_transposed.to_dict('records')
                
        except Exception as e:
            print(f"Error fetching financial statements for {symbol}: {e}")
            
        return statements
    
    def get_key_metrics(self, symbol: str) -> Dict:
        """Get key financial metrics and ratios"""
        try:
            if self.fmp_api_key:
                url = f"https://financialmodelingprep.com/api/v3/key-metrics-ttm/{symbol}?apikey={self.fmp_api_key}"
                response = requests.get(url)
                if response.status_code == 200:
                    data = response.json()
                    if data:
                        return data[0]
                
                url = f"https://financialmodelingprep.com/api/v3/ratios-ttm/{symbol}?apikey={self.fmp_api_key}"
                response = requests.get(url)
                if response.status_code == 200:
                    ratios = response.json()
                    if ratios:
                        return {**data[0], **ratios[0]} if data else ratios[0]
            
            ticker = yf.Ticker(symbol)
            info = ticker.info
            return {
                'peRatioTTM': info.get('trailingPE'),
                'priceToBookRatioTTM': info.get('priceToBook'),
                'priceToSalesRatioTTM': info.get('priceToSalesTrailing12Months'),
                'enterpriseValueMultipleTTM': info.get('enterpriseToEbitda'),
                'evToSalesTTM': info.get('enterpriseToRevenue'),
                'debtToEquityTTM': info.get('debtToEquity'),
                'returnOnEquityTTM': info.get('returnOnEquity'),
                'returnOnAssetsTTM': info.get('returnOnAssets')
            }
            
        except Exception as e:
            print(f"Error fetching key metrics for {symbol}: {e}")
            return {}
    
    def get_enterprise_value_data(self, symbol: str) -> Dict:
        """Get enterprise value components"""
        try:
            if self.fmp_api_key:
                url = f"https://financialmodelingprep.com/api/v3/enterprise-values/{symbol}?apikey={self.fmp_api_key}"
                response = requests.get(url)
                if response.status_code == 200:
                    data = response.json()
                    if data:
                        return data[0]
            
            ticker = yf.Ticker(symbol)
            info = ticker.info
            balance_sheet = ticker.balance_sheet
            
            market_cap = info.get('marketCap', 0)
            total_debt = 0
            cash = 0
            
            if not balance_sheet.empty:
                latest = balance_sheet.iloc[:, 0]
                total_debt = latest.get('Total Debt', 0) or 0
                cash = latest.get('Cash And Cash Equivalents', 0) or 0
            
            enterprise_value = market_cap + total_debt - cash
            
            return {
                'symbol': symbol,
                'marketCapitalization': market_cap,
                'totalDebt': total_debt,
                'cashAndCashEquivalents': cash,
                'enterpriseValue': enterprise_value
            }
            
        except Exception as e:
            print(f"Error fetching enterprise value data for {symbol}: {e}")
            return {}
    
    def get_dcf_data(self, symbol: str) -> Dict:
        """Get DCF valuation data"""
        try:
            if self.fmp_api_key:
                url = f"https://financialmodelingprep.com/api/v3/discounted-cash-flow/{symbol}?apikey={self.fmp_api_key}"
                response = requests.get(url)
                if response.status_code == 200:
                    data = response.json()
                    if data:
                        return data[0]
            
            return {}
            
        except Exception as e:
            print(f"Error fetching DCF data for {symbol}: {e}")
            return {}
    
    def get_analyst_estimates(self, symbol: str) -> Dict:
        """Get analyst estimates and price targets"""
        try:
            if self.fmp_api_key:
                url = f"https://financialmodelingprep.com/api/v3/analyst-estimates/{symbol}?apikey={self.fmp_api_key}"
                response = requests.get(url)
                estimates = response.json() if response.status_code == 200 else []
                
                url = f"https://financialmodelingprep.com/api/v4/price-target-summary?symbol={symbol}&apikey={self.fmp_api_key}"
                response = requests.get(url)
                price_targets = response.json() if response.status_code == 200 else []
                
                return {
                    'estimates': estimates,
                    'price_targets': price_targets
                }
            
            return {}
            
        except Exception as e:
            print(f"Error fetching analyst data for {symbol}: {e}")
            return {}
    
    def get_historical_prices(self, symbol: str, period: str = "1y") -> pd.DataFrame:
        """Get historical price data"""
        try:
            ticker = yf.Ticker(symbol)
            hist = ticker.history(period=period)
            return hist
        except Exception as e:
            print(f"Error fetching historical prices for {symbol}: {e}")
            return pd.DataFrame()
    
    def get_comprehensive_data(self, symbol: str) -> Dict:
        """Get all available data for a symbol"""
        print(f"Fetching comprehensive data for {symbol}...")
        
        data = {
            'symbol': symbol,
            'profile': self.get_company_profile(symbol),
            'financial_statements': self.get_financial_statements(symbol),
            'key_metrics': self.get_key_metrics(symbol),
            'enterprise_value': self.get_enterprise_value_data(symbol),
            'dcf': self.get_dcf_data(symbol),
            'analyst_data': self.get_analyst_estimates(symbol),
            'historical_prices': self.get_historical_prices(symbol)
        }
        
        print(f"Data fetch completed for {symbol}")
        return data
