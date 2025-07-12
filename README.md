# Equity Valuation Models

A comprehensive toolkit for equity valuation modeling with automated data retrieval and Excel template generation.

## Features

- **EV/EBITDA Valuation Model**: Forward-looking enterprise value to EBITDA multiple analysis
- **EV/Sales Valuation Model**: Enterprise value to sales multiple analysis  
- **DCF (Discounted Cash Flow) Model**: Intrinsic value calculation with detailed cash flow projections
- **Sensitivity Analysis**: Test different exit multiples and scenarios
- **Automated Data Retrieval**: Fetch real-time financial data from multiple sources
- **Reusable Templates**: Works with any stock symbol

## Project Structure

```
equity-valuation-models/
├── templates/              # Excel template files
│   ├── ev_ebitda_model.xlsx
│   ├── ev_sales_model.xlsx
│   └── dcf_model.xlsx
├── src/                    # Python source code
│   ├── data_fetcher.py     # Financial data retrieval
│   ├── excel_generator.py  # Excel template generation
│   ├── models/             # Valuation model implementations
│   └── utils/              # Utility functions
├── examples/               # Example analyses
├── config/                 # Configuration files
└── tests/                  # Unit tests
```

## Quick Start

1. Install dependencies: `pip install -r requirements.txt`
2. Configure API keys in `config/api_keys.yaml`
3. Run example: `python examples/analyze_stock.py AAPL`

## Supported Data Sources

- Financial Modeling Prep API
- Alpha Vantage
- Yahoo Finance
- Custom subscription sites (configurable)

## Usage Examples

### EV/EBITDA Analysis
```python
from src.models.ev_ebitda import EVEBITDAModel

model = EVEBITDAModel("AAPL")
model.load_data()
model.calculate_valuation()
model.sensitivity_analysis(multiple_range=(15, 25))
model.export_to_excel("AAPL_ev_ebitda.xlsx")
```

### Sensitivity Analysis
Test different exit multiples to see impact on stock price:
- Current multiple: 18.0x
- Test scenario: 19.5x
- Calculate implied stock price at exit

## Requirements

- Python 3.8+
- pandas
- openpyxl
- requests
- xlsxwriter
- Financial data API access
