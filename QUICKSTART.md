# KarthiStocks - Quick Start Guide

## Prerequisites
- Python 3.7 or higher
- pip package manager

## Installation Steps

1. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

2. **Prepare Your Excel File**
   - Use the provided `stocks.xlsx` file OR create your own
   - Format: Column A should contain stock ticker symbols
   - First row is the header (e.g., "Symbol")
   - Starting from row 2, list your stock symbols

   Example:
   ```
   | Symbol |
   |--------|
   | AAPL   |
   | GOOGL  |
   | MSFT   |
   ```

## Running the Application

### Option 1: Live Stock Prices (Requires Internet)
```bash
python stock_tracker.py
```
or with a custom Excel file:
```bash
python stock_tracker.py /path/to/your/file.xlsx
```

### Option 2: Demo Mode (Works Offline)
```bash
python stock_tracker_demo.py
```

## What Happens

1. ✓ Reads stock symbols from your Excel file
2. ✓ Fetches current prices from Yahoo Finance
3. ✓ Displays a formatted table in the terminal
4. ✓ Updates the Excel file with:
   - Company names
   - Current prices
   - Previous closing prices
   - Price changes ($ and %)
   - Trading volume
   - Market capitalization
   - Timestamp of last update

## Output Example

```
LIVE STOCK PRICES
=================================================================
Symbol  Company      Current Price  Change  Change %  Volume
AAPL    Apple Inc.   175.43        +1.18   +0.68%    45,234,567
GOOGL   Alphabet     139.82        +0.87   +0.63%    23,456,789
MSFT    Microsoft    338.11        +0.66   +0.20%    34,567,890
=================================================================
```

## Troubleshooting

**Problem**: "ModuleNotFoundError"
**Solution**: Run `pip install -r requirements.txt`

**Problem**: "Excel file not found"
**Solution**: Make sure `stocks.xlsx` exists or specify the correct path

**Problem**: "No internet connection" or API errors
**Solution**: Use `stock_tracker_demo.py` for offline testing

**Problem**: "No data for symbol XXX"
**Solution**: Verify the ticker symbol is correct (use Yahoo Finance symbols)

## Tips

- Run regularly to track price changes over time
- The Excel file maintains a history with timestamps
- You can add/remove symbols by editing column A in the Excel file
- Use uppercase for ticker symbols (e.g., AAPL, not aapl)
- Popular symbols: AAPL, GOOGL, MSFT, AMZN, TSLA, META, NVDA, etc.

## Advanced Usage

### Watch Multiple Portfolios
Create separate Excel files for different portfolios:
```bash
python stock_tracker.py tech_stocks.xlsx
python stock_tracker.py dividend_stocks.xlsx
```

### Automate Updates
On Linux/Mac (crontab example for hourly updates):
```bash
0 * * * * cd /path/to/KarthiStocks && python stock_tracker.py
```

On Windows (Task Scheduler):
- Create a new task
- Trigger: Hourly
- Action: Run `python stock_tracker.py`

## Support

For issues or questions, please check:
- README.md for detailed documentation
- GitHub Issues: https://github.com/netpriv7-byte/KarthiStocks/issues
