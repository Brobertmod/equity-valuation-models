"""
Helper utility functions for valuation models
"""

import pandas as pd
from typing import Dict, List, Optional, Union
import numpy as np

def safe_divide(numerator: float, denominator: float, default: float = 0) -> float:
    """Safely divide two numbers, returning default if denominator is zero"""
    try:
        if denominator == 0:
            return default
        return numerator / denominator
    except (TypeError, ZeroDivisionError):
        return default

def format_currency(value: float) -> str:
    """Format a number as currency"""
    if value >= 1e12:
        return f"${value/1e12:.2f}T"
    elif value >= 1e9:
        return f"${value/1e9:.2f}B"
    elif value >= 1e6:
        return f"${value/1e6:.2f}M"
    elif value >= 1e3:
        return f"${value/1e3:.2f}K"
    else:
        return f"${value:.2f}"

def format_percentage(value: float, decimals: int = 1) -> str:
    """Format a decimal as percentage"""
    try:
        return f"{value * 100:.{decimals}f}%"
    except (TypeError, ValueError):
        return "N/A"

def calculate_cagr(start_value: float, end_value: float, periods: int) -> float:
    """Calculate Compound Annual Growth Rate"""
    try:
        if start_value <= 0 or end_value <= 0 or periods <= 0:
            return 0
        return (end_value / start_value) ** (1 / periods) - 1
    except (TypeError, ValueError, ZeroDivisionError):
        return 0

def clean_financial_data(data: Dict) -> Dict:
    """Clean and standardize financial data"""
    cleaned = {}
    
    for key, value in data.items():
        if value is None:
            cleaned[key] = 0
        elif isinstance(value, str):
            try:
                cleaned[key] = float(value.replace(',', '').replace('$', ''))
            except (ValueError, AttributeError):
                cleaned[key] = value
        else:
            cleaned[key] = value
    
    return cleaned

def validate_symbol(symbol: str) -> bool:
    """Validate stock symbol format"""
    if not symbol or not isinstance(symbol, str):
        return False
    
    symbol = symbol.strip().upper()
    return symbol.isalnum() and 1 <= len(symbol) <= 5

def get_financial_statement_item(statements: List[Dict], item_key: str, default: float = 0) -> float:
    """Safely extract an item from financial statements"""
    try:
        if not statements or not isinstance(statements, list):
            return default
        
        latest = statements[0]
        return latest.get(item_key, default) or default
    except (IndexError, TypeError, AttributeError):
        return default

def calculate_enterprise_value(market_cap: float, total_debt: float, cash: float) -> float:
    """Calculate enterprise value"""
    try:
        return market_cap + total_debt - cash
    except (TypeError, ValueError):
        return 0

def generate_multiple_scenarios(base_multiple: float, variation_pct: float = 0.2, 
                              num_scenarios: int = 5) -> List[float]:
    """Generate multiple scenarios around a base multiple"""
    try:
        min_multiple = base_multiple * (1 - variation_pct)
        max_multiple = base_multiple * (1 + variation_pct)
        
        return list(np.linspace(min_multiple, max_multiple, num_scenarios))
    except (TypeError, ValueError):
        return [base_multiple]

def create_sensitivity_matrix(base_values: List[float], multipliers: List[float]) -> pd.DataFrame:
    """Create a sensitivity analysis matrix"""
    try:
        matrix = []
        for base_val in base_values:
            row = [base_val * mult for mult in multipliers]
            matrix.append(row)
        
        return pd.DataFrame(matrix, 
                          index=[f"Base_{i+1}" for i in range(len(base_values))],
                          columns=[f"Mult_{mult:.1f}x" for mult in multipliers])
    except Exception:
        return pd.DataFrame()
