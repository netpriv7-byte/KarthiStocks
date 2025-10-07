# KarthiStocks

A Python application that connects to Excel worksheets, fetches live stock prices, and displays them in real-time.

## Features

- 📊 Read stock symbols from Excel files
- 💰 Fetch live stock prices using Yahoo Finance API
- 📈 Display stock data in a formatted table
- 💾 Update Excel file with latest prices and timestamps
- 🔄 Track price changes and percentage movements
- 📉 Display volume and market cap information

## Installation

1. Clone this repository:
```bash
git clone https://github.com/netpriv7-byte/KarthiStocks.git
cd KarthiStocks
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

Run the stock tracker with the default Excel file (`stocks.xlsx`):

```bash
python stock_tracker.py
```

### Custom Excel File

Specify a custom Excel file:

```bash
python stock_tracker.py /path/to/your/stocks.xlsx
```

## Excel File Format

The Excel file should have the following structure:

| Symbol |
|--------|
| AAPL   |
| GOOGL  |
| MSFT   |
| AMZN   |
| TSLA   |

- **Sheet Name**: `Stocks` (or the first sheet will be used)
- **Column A**: Stock symbols (ticker symbols)
- **Row 1**: Header (will be updated by the application)

After running the application, the Excel file will be updated with:
- Symbol
- Company Name
- Current Price
- Previous Close
- Change
- Change %
- Volume
- Market Cap
- Last Updated timestamp

## Example Output

```
====================================================================================================
KarthiStocks - Live Stock Price Tracker
====================================================================================================
Found 5 stock symbols: AAPL, GOOGL, MSFT, AMZN, TSLA

Fetching live prices for 5 stocks...

====================================================================================================
LIVE STOCK PRICES
====================================================================================================
Symbol       Company  Current Price  Previous Close  Change  Change %      Volume    Market Cap
  AAPL    Apple Inc.         175.43          174.25    1.18      0.68  45234567.0  2750000000000
 GOOGL  Alphabet Inc.        139.82          138.95    0.87      0.63  23456789.0  1750000000000
  MSFT     Microsoft         338.11          337.45    0.66      0.20  34567890.0  2510000000000
====================================================================================================

Last Updated: 2025-10-07 14:30:00

Excel file updated successfully: stocks.xlsx
Done!
```

## Requirements

- Python 3.7+
- openpyxl 3.1.2
- pandas 2.1.4
- yfinance 0.2.33

## How It Works

1. **Read Stock Symbols**: The application reads stock ticker symbols from column A of the Excel file (starting from row 2, skipping the header)

2. **Fetch Live Prices**: Using the `yfinance` library, the application connects to Yahoo Finance API to retrieve:
   - Current stock price
   - Previous closing price
   - Price change and percentage change
   - Trading volume
   - Market capitalization
   - Company name

3. **Display Data**: Stock information is displayed in a formatted table in the terminal

4. **Update Excel**: The Excel file is updated with all the fetched data, including a timestamp of when the data was retrieved

## Sample Excel File

A sample `stocks.xlsx` file is included with the following popular stock symbols:
- AAPL (Apple Inc.)
- GOOGL (Alphabet Inc.)
- MSFT (Microsoft Corporation)
- AMZN (Amazon.com Inc.)
- TSLA (Tesla Inc.)

You can modify this file to track any stocks you're interested in.

## Error Handling

The application includes comprehensive error handling:
- Missing Excel files
- Invalid stock symbols
- Network connectivity issues
- API rate limiting

If a stock symbol cannot be found or data is unavailable, it will be marked as "N/A" in the output.

## License

This project is open source and available under the MIT License.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Support

For issues, questions, or suggestions, please open an issue on GitHub.