@echo off
echo ==========================================
echo Excel Equity Valuation Model Generator
echo ==========================================
echo.
set /p symbol="Enter ticker symbol (e.g., AAPL, MSFT, TSLA): "
echo.
echo Select models to generate:
echo 1. All models (EV/EBITDA, EV/Sales, DCF)
echo 2. EV/EBITDA only
echo 3. EV/Sales only  
echo 4. DCF only
set /p choice="Enter choice (1-4): "
echo.
set /p cleanup="Remove old files for other symbols? (y/N): "
if /i "%cleanup%"=="y" (
    echo Cleaning up old files...
    for %%f in (output\*.xlsx) do (
        echo %%f | findstr /v "%symbol%_" >nul && del "%%f" 2>nul
    )
    echo Old files removed.
)
echo.
echo Generating models for %symbol%...

if "%choice%"=="1" (
    python examples/analyze_stock.py %symbol%
) else if "%choice%"=="2" (
    python examples/analyze_stock.py %symbol% --models ev_ebitda
) else if "%choice%"=="3" (
    python examples/analyze_stock.py %symbol% --models ev_sales
) else if "%choice%"=="4" (
    python examples/analyze_stock.py %symbol% --models dcf
) else (
    echo Invalid choice. Generating all models...
    python examples/analyze_stock.py %symbol%
)

echo.
echo ==========================================
echo Models generated successfully!
echo New files created in output/ folder:
dir output\%symbol%_*.xlsx 2>nul || echo No files found for %symbol%
echo ==========================================
echo Look for files starting with: %symbol%_
pause
