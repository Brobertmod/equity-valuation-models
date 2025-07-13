# Excel Add-in Setup Guide for Financial Modeling Prep

This guide explains how to set up and use the Financial Modeling Prep Excel add-in alongside the existing Python-based equity valuation models.

## Overview

The Financial Modeling Prep Excel add-in provides real-time financial data directly within Excel spreadsheets using custom functions. This complements our Python-based model generation by allowing users to:

1. **Python Models**: Generate comprehensive valuation templates with historical data
2. **Excel Add-in**: Access real-time data and additional metrics within Excel

## Installation Steps

### Step 1: Install FMP Excel Add-in from Microsoft AppSource

1. **Open Excel** (desktop version recommended)
2. **Go to Insert > Get Add-ins** (or Office Add-ins)
3. **Search for "Financial Modeling Prep"** in the store
4. **Click "Add"** to install the add-in
5. **Accept permissions** when prompted

**Alternative Method:**
- Visit [Microsoft AppSource FMP Page](https://appsource.microsoft.com/en-us/product/office/WA200003535)
- Click "Get it now" and sign in with your Microsoft account
- Follow the installation prompts

### Step 2: Configure API Key in Excel

1. **Open Excel** and create a new workbook
2. **Look for the FMP tab** in the Excel ribbon
3. **Click "Settings" or "API Key"** in the FMP ribbon
4. **Enter your API key**: `[Your FMP API Key]`
5. **Test the connection** by trying a simple function

### Step 3: Test Basic Functions

Try these basic FMP functions in Excel cells:

```excel
=FMPFinancialStatementSymbol("AAPL")
=FMPIncomeStatement("AAPL", "2023")
=FMPBalanceSheetStatement("AAPL", "2023")
=FMPCashFlowStatement("AAPL", "2023")
=FMPKeyMetrics("AAPL")
=FMPRatios("AAPL")
=FMPEnterpriseValue("AAPL")
```

## Integration with Python Models

### Workflow Integration

1. **Generate Base Models**: Use Python scripts to create comprehensive Excel templates
   ```bash
   python examples/analyze_stock.py AAPL
   ```

2. **Enhance with Real-time Data**: Open generated Excel files and use FMP functions for:
   - Real-time price updates
   - Latest quarterly data
   - Peer comparison metrics
   - Analyst estimates

3. **Dynamic Updates**: Excel add-in functions automatically refresh data when Excel recalculates

### Recommended Usage Pattern

**Python-Generated Models** (Static, Comprehensive):
- Historical financial statements (5+ years)
- Calculated valuation scenarios
- Sensitivity analysis tables
- Professional formatting and charts

**Excel Add-in Functions** (Dynamic, Real-time):
- Current stock price: `=FMPPrice("AAPL")`
- Latest quarterly metrics: `=FMPKeyMetrics("AAPL")`
- Real-time market data: `=FMPMarketCap("AAPL")`
- Peer comparison: `=FMPRatios("MSFT")` vs `=FMPRatios("AAPL")`

## Available FMP Excel Functions

### Company Information
- `=FMPProfile("AAPL")` - Company profile
- `=FMPPrice("AAPL")` - Current stock price
- `=FMPMarketCap("AAPL")` - Market capitalization

### Financial Statements
- `=FMPIncomeStatement("AAPL", "2023")` - Income statement
- `=FMPBalanceSheetStatement("AAPL", "2023")` - Balance sheet
- `=FMPCashFlowStatement("AAPL", "2023")` - Cash flow statement

### Valuation Metrics
- `=FMPKeyMetrics("AAPL")` - Key financial metrics
- `=FMPRatios("AAPL")` - Financial ratios
- `=FMPEnterpriseValue("AAPL")` - Enterprise value components

### Advanced Data
- `=FMPAnalystEstimates("AAPL")` - Analyst estimates
- `=FMPPriceTarget("AAPL")` - Price targets
- `=FMPEarningsCalendar("AAPL")` - Earnings calendar

## Troubleshooting

### Common Issues

1. **"Function not recognized"**
   - Ensure the FMP add-in is properly installed and enabled
   - Check that Excel is connected to the internet
   - Restart Excel if necessary

2. **"API key invalid"**
   - Verify the API key is correctly entered in the FMP add-in settings
   - Check your FMP subscription status
   - Ensure the key has sufficient API call limits

3. **"No data returned"**
   - Verify the stock symbol is correct (use uppercase)
   - Check if the company is publicly traded
   - Some functions may require premium FMP subscription

4. **Functions not updating**
   - Press Ctrl+Alt+F9 to force recalculation
   - Check Excel calculation settings (File > Options > Formulas)
   - Ensure automatic calculation is enabled

### Getting Help

- **FMP Documentation**: [https://financialmodelingprep.com/developer/docs/excel-add-on](https://financialmodelingprep.com/developer/docs/excel-add-on)
- **Excel Add-in Support**: Contact FMP support through their website
- **Python Integration**: Refer to the main README.md and EXCEL_USAGE_GUIDE.md

## Best Practices

### Performance Optimization
- Use FMP functions sparingly in large spreadsheets
- Cache results when possible (copy and paste special > values)
- Avoid circular references with FMP functions

### Data Accuracy
- Always verify critical data points with multiple sources
- Understand the data refresh frequency for different metrics
- Use appropriate date parameters for historical data

### Security
- Never share spreadsheets containing your API key
- Use Excel's protection features to hide sensitive formulas
- Regularly monitor your API usage limits

## Example Integration

Here's how to enhance a Python-generated EV/EBITDA model with Excel add-in functions:

1. **Open Python-generated model**: `AAPL_ev_ebitda_model.xlsx`
2. **Add real-time price cell**: `=FMPPrice("AAPL")`
3. **Add current market cap**: `=FMPMarketCap("AAPL")`
4. **Add latest EBITDA**: `=FMPKeyMetrics("AAPL")` (extract EBITDA)
5. **Create peer comparison**: Add columns for `=FMPRatios("MSFT")`, etc.

This combination provides both comprehensive historical analysis (Python) and real-time market data (Excel add-in).
