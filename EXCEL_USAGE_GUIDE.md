# Excel Equity Valuation Models - User Guide

## Quick Start: Changing Ticker Symbols

### Method 1: Using Excel + Batch Scripts (Recommended)

1. **Open any generated Excel file** (e.g., `AAPL_ev_ebitda_model.xlsx`)
2. **Locate the yellow ticker symbol cell** (B2) at the top of the spreadsheet
3. **Change the symbol** to your desired ticker (e.g., change "AAPL" to "MSFT")
4. **Run the regeneration script:**
   - **Windows**: Double-click `generate_models.bat`
   - **Mac/Linux**: Run `./generate_models.sh` in terminal
5. **Enter the new ticker symbol** when prompted
6. **Select model types** to generate (or choose "All models")
7. **Check the output/ folder** for your new Excel files

### Method 2: Command Line (Advanced Users)

```bash
# Generate all models for a symbol
python examples/analyze_stock.py TSLA

# Generate specific model types
python examples/analyze_stock.py MSFT --models ev_ebitda
python examples/analyze_stock.py GOOGL --models ev_sales dcf
```

## Excel Template Features

### Ticker Symbol Input
- **Location**: Cell B2 in all templates
- **Appearance**: Yellow background with red text for easy identification
- **Purpose**: Shows current symbol and serves as reference for regeneration

### 5-Year Exit Scenarios
- **EV/EBITDA Model**: Projects EBITDA growth over 5 years with exit multiples
- **EV/Sales Model**: Projects revenue growth over 5 years with exit multiples
- **DCF Model**: 5-year cash flow projections with terminal value

### Dynamic Formulas
- All calculations use Excel formulas that update automatically
- Change growth assumptions to see impact on valuations
- Sensitivity analysis tables with color-coded results

## Supported Ticker Symbols
- Any publicly traded stock symbol (e.g., AAPL, MSFT, TSLA, GOOGL)
- Data sourced from Yahoo Finance and Financial Modeling Prep APIs
- Automatically fetches current financial statements and market data

## Model Types

### EV/EBITDA Model
- **Purpose**: Enterprise Value to EBITDA multiple-based valuation
- **Key Features**:
  - 5-year EBITDA growth projections
  - Multiple exit scenarios (12x to 25x)
  - Annualized return calculations
  - Sensitivity analysis with color-coded results

### EV/Sales Model
- **Purpose**: Enterprise Value to Sales multiple-based valuation
- **Key Features**:
  - 5-year revenue growth projections
  - Multiple exit scenarios (2x to 8x)
  - Annualized return calculations
  - Sensitivity analysis with color-coded results

### DCF Model
- **Purpose**: Discounted Cash Flow intrinsic valuation
- **Key Features**:
  - 5-year detailed cash flow projections
  - Terminal value calculation
  - WACC-based discounting
  - Sensitivity analysis for WACC and terminal growth rates

## Important: Understanding File Creation

**Key Point:** The scripts create NEW files with the new symbol name. They do NOT update your existing Excel files.

### Example Workflow:
1. You have: `AAPL_ev_sales_model.xlsx` (Apple data)
2. You change the symbol in Excel to "MSFT" 
3. You run the script and enter "MSFT"
4. **Result:** New file created: `MSFT_ev_sales_model.xlsx` (Microsoft data)
5. **Your original AAPL file remains unchanged**

### Finding Your New Files:
- Look in the `output/` folder
- New files start with your new symbol (e.g., `MSFT_`, `TSLA_`, etc.)
- Check the file timestamps - newest files are your generated models
- File sizes may differ between companies due to different data availability

### Troubleshooting "Old Data" Issues:
- **Problem:** "I see old symbol data after running the script"
- **Solution:** Make sure you're opening the NEW file with the new symbol name
- **Check:** Look for `[NEW_SYMBOL]_ev_sales_model.xlsx` in the output folder
- **Verify:** The script output shows which files were created

## Troubleshooting

### Common Issues
1. **"ERROR: No data found for [SYMBOL]"**
   - Check that the ticker symbol is correct
   - Ensure the company is publicly traded
   - Try running the script again (API timeouts can occur)

2. **Negative valuations**
   - Check if the company has negative EBITDA or revenue
   - Review growth assumptions in the model
   - Consider if the company is appropriate for the valuation method

3. **Missing data in templates**
   - Some companies may have limited financial data available
   - Try using a different data source or manual input
   - Check if the company reports in a different currency

4. **"Script says success but I see old data"**
   - You're likely looking at the old file instead of the new one
   - Check for files with the NEW symbol name in the output folder
   - The script creates new files, it doesn't update existing ones

### Getting Help
- Review the README.md file for technical setup instructions
- Check the examples/ folder for sample usage
- Ensure all dependencies are installed: `pip install -r requirements.txt`

## Advanced Usage

### Customizing Growth Assumptions
- **EV Models**: Modify growth rates in the "Growth Assumptions" section
- **DCF Model**: Adjust assumptions in the "Key Assumptions" section
- **Sensitivity Analysis**: Change multiple ranges to test different scenarios

### Adding New Data Sources
- Edit `config/api_keys.yaml` to add API keys
- Modify `src/data_fetcher.py` to include new data sources
- Update templates to incorporate additional metrics

### Batch Processing
- Use the provided scripts to generate models for multiple symbols
- Modify `examples/analyze_stock.py` for custom batch processing
- Set up scheduled runs for regular model updates

## Data Sources and Accuracy

### Primary Sources
- **Yahoo Finance**: Free, comprehensive financial data
- **Financial Modeling Prep**: Enhanced fundamental data (API key recommended)

### Data Refresh
- Models use the most recent available financial data
- Quarterly and annual statements are automatically selected
- Market data (prices, market cap) are real-time or delayed

### Limitations
- Data accuracy depends on source reliability
- Some metrics may be estimated or calculated
- Always verify critical assumptions and data points
- Models are for analysis purposes only, not investment advice

## File Structure

```
equity-valuation-models/
├── output/                     # Generated Excel files
├── examples/                   # Usage examples
├── src/                       # Source code
├── config/                    # Configuration files
├── generate_models.bat        # Windows batch script
├── generate_models.sh         # Unix shell script
├── EXCEL_USAGE_GUIDE.md      # This guide
└── README.md                 # Technical documentation
```

## Next Steps

1. **Start with a familiar stock** (e.g., AAPL, MSFT) to understand the templates
2. **Experiment with different assumptions** to see impact on valuations
3. **Compare results across models** for the same company
4. **Use sensitivity analysis** to understand key value drivers
5. **Customize templates** for your specific analysis needs

For technical questions or issues, refer to the main README.md file or the source code documentation.
