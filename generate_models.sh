#!/bin/bash
echo "=========================================="
echo "Excel Equity Valuation Model Generator"
echo "=========================================="
echo
read -p "Enter ticker symbol (e.g., AAPL, MSFT, TSLA): " symbol
echo
echo "Select models to generate:"
echo "1. All models (EV/EBITDA, EV/Sales, DCF)"
echo "2. EV/EBITDA only"
echo "3. EV/Sales only"
echo "4. DCF only"
read -p "Enter choice (1-4): " choice
echo
echo "Generating models for $symbol..."

case $choice in
    1)
        python examples/analyze_stock.py "$symbol"
        ;;
    2)
        python examples/analyze_stock.py "$symbol" --models ev_ebitda
        ;;
    3)
        python examples/analyze_stock.py "$symbol" --models ev_sales
        ;;
    4)
        python examples/analyze_stock.py "$symbol" --models dcf
        ;;
    *)
        echo "Invalid choice. Generating all models..."
        python examples/analyze_stock.py "$symbol"
        ;;
esac

echo
echo "=========================================="
echo "Models generated successfully!"
echo "Check the output/ folder for Excel files."
echo "=========================================="
