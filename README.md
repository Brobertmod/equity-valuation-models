# Equity Valuation Models

A comprehensive Python toolkit for creating Excel-based equity valuation models with automated data retrieval, sensitivity analysis, and **easy ticker symbol changes directly in Excel**.

## 🚀 Quick Start for Non-Technical Users

### Changing Ticker Symbols in Excel (Recommended)

1. **Download or open any generated Excel file** (e.g., `AAPL_ev_ebitda_model.xlsx`)
2. **Find the yellow ticker symbol cell** (B2) at the top of the spreadsheet
3. **Change the symbol** to analyze a new company (e.g., "AAPL" → "MSFT")
4. **Run the simple script:**
   - **Windows**: Double-click `generate_models.bat`
   - **Mac/Linux**: Run `./generate_models.sh` in terminal
5. **Enter the new ticker symbol** and select model types
6. **Get your new Excel files** in the output/ folder

📖 **[Complete Excel Usage Guide](EXCEL_USAGE_GUIDE.md)** - Detailed instructions with screenshots

## Features

- **🎯 Excel-Based Symbol Input**: Change ticker symbols directly in Excel templates
- **📊 Multiple Valuation Models**: EV/EBITDA, EV/Sales, and DCF models with 5-year exit scenarios
- **🔄 Automated Data Retrieval**: Fetches current financial data from Yahoo Finance and Financial Modeling Prep APIs
- **📈 Professional Excel Templates**: Embedded formulas, sensitivity analysis, and color-coded results
- **⚡ Simple Regeneration**: User-friendly batch scripts for non-technical users
- **🔧 Reusable Framework**: Easily generate models for any publicly traded stock

## Installation

```bash
git clone https://github.com/Brobertmod/equity-valuation-models.git
cd equity-valuation-models
pip install -r requirements.txt
```

## Usage Options

### Option 1: Excel + Batch Scripts (Recommended for Most Users)
```bash
# Windows
generate_models.bat

# Mac/Linux  
./generate_models.sh
```

### Option 2: Command Line (Advanced Users)
```bash
# Generate all models for Apple Inc.
python examples/analyze_stock.py AAPL

# Generate specific models
python examples/analyze_stock.py MSFT --models ev_ebitda dcf
```

## Generated Files
- `AAPL_ev_ebitda_model.xlsx` - EV/EBITDA valuation with 5-year exit scenarios
- `AAPL_ev_sales_model.xlsx` - EV/Sales valuation with 5-year exit scenarios
- `AAPL_dcf_model.xlsx` - Discounted Cash Flow model with detailed 5-year projections

## Configuration

### API Keys (Optional)

For enhanced data access, add your API keys to `config/api_keys.yaml`:

```yaml
fmp_api_key: "your_financial_modeling_prep_key"
alpha_vantage_key: "your_alpha_vantage_key"
```

### Supported Data Sources
- **Yahoo Finance** (Free, no API key required)
- **Financial Modeling Prep** (Enhanced data with API key)
- **Alpha Vantage** (Additional metrics with API key)

## Excel Template Features

### 🎯 Ticker Symbol Input
- **Highlighted yellow cell (B2)** in all templates for easy symbol identification
- **Clear instructions** for changing symbols and regenerating models
- **Error handling** for invalid symbols or missing data

### 📊 5-Year Exit Scenarios
- **EV/EBITDA**: Projects EBITDA growth with multiple exit scenarios
- **EV/Sales**: Projects revenue growth with multiple exit scenarios  
- **DCF**: Detailed 5-year cash flow projections with terminal value

### 🔄 Dynamic Formulas
- All calculations use Excel formulas that update automatically
- Change assumptions to see immediate impact on valuations
- Professional sensitivity analysis with color-coded results

## Project Structure

```
equity-valuation-models/
├── src/
│   ├── data_fetcher.py      # Financial data retrieval
│   ├── excel_generator.py   # Excel template generation
│   ├── models/              # Valuation model implementations
│   └── utils/               # Helper utilities
├── examples/
│   └── analyze_stock.py     # Main analysis script
├── config/
│   └── api_keys.yaml        # API configuration
├── output/                  # Generated Excel files
├── generate_models.bat      # Windows batch script
├── generate_models.sh       # Unix shell script
├── EXCEL_USAGE_GUIDE.md    # Detailed Excel usage instructions
└── requirements.txt         # Python dependencies
```

## Model Details

### EV/EBITDA Model (5-Year Exit Scenarios)
- Enterprise Value to EBITDA multiple analysis with growth projections
- 5-year EBITDA growth assumptions and exit multiple scenarios
- Annualized return calculations and total return analysis
- Professional sensitivity analysis with conditional formatting

### EV/Sales Model (5-Year Exit Scenarios)
- Enterprise Value to Sales multiple analysis with revenue growth
- 5-year revenue projections and exit multiple scenarios
- Market multiple comparisons and return calculations
- Color-coded sensitivity tables for different growth/multiple combinations

### DCF Model (Detailed 5-Year Projections)
- Comprehensive 5-year cash flow projections
- Terminal value calculations with growth assumptions
- WACC-based discounting and equity value derivation
- Dual-axis sensitivity analysis (WACC vs Terminal Growth)

## Troubleshooting

### Common Issues
- **Negative valuations**: Check for negative EBITDA/revenue or unrealistic assumptions
- **Missing data**: Some companies may have limited financial data available
- **API timeouts**: Retry the script if data fetching fails

### Getting Help
- 📖 Read the [Excel Usage Guide](EXCEL_USAGE_GUIDE.md) for detailed instructions
- 🔧 Check `examples/analyze_stock.py` for usage examples
- 📊 Review generated Excel files for formula references

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/new-model`)
3. Commit your changes (`git commit -am 'Add new valuation model'`)
4. Push to the branch (`git push origin feature/new-model`)
5. Create a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Disclaimer

These models are for educational and analysis purposes only. Always conduct thorough due diligence and consult with financial professionals before making investment decisions.
