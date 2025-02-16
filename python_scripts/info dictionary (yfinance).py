import os
import yfinance as yf
import pandas as pd
from datetime import datetime

# ✅ Choose a stock symbol
ticker = "AAPL"  # Change this if needed
stock = yf.Ticker(ticker)

# ✅ Get all available fields from yfinance
stock_info = stock.info

# ✅ Function to convert Unix timestamps to human-readable date format
def convert_timestamp(value):
    if isinstance(value, (int, float)) and value > 0:
        # Convert milliseconds to seconds if needed
        if value > 10**10:  # If value is too large, it's in milliseconds
            value = value / 1000
        return datetime.utcfromtimestamp(value).strftime("%d-%m-%Y")  # Convert to dd-mm-yyyy
    return value  # Return original if not a timestamp

# ✅ Convert all Unix timestamp fields
date_fields = [
    "dividendDate", "earningsTimestamp", "earningsTimestampStart", "earningsTimestampEnd",
    "earningsCallTimestampStart", "earningsCallTimestampEnd", "exDividendDate",
    "lastDividendDate", "lastFiscalYearEnd", "nextFiscalYearEnd", "mostRecentQuarter",
    "sharesShortPreviousMonthDate", "dateShortInterest", "governanceEpochDate",
    "compensationAsOfEpochDate", "firstTradeDateMilliseconds", "lastSplitDate"
]

for field in date_fields:
    if field in stock_info:
        stock_info[field] = convert_timestamp(stock_info[field])

# ✅ Convert Dictionary to DataFrame
stock_df = pd.DataFrame(stock_info.items(), columns=["Metric", "Value"])

# ✅ Get today's date for filename
current_date = datetime.today().strftime("%Y-%m-%d")

# ✅ Define File Path (Save in Downloads Folder)
downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
output_file = os.path.join(downloads_folder, f"{ticker}_StockInfo_{current_date}.xlsx")

# ✅ Save to Excel
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    stock_df.to_excel(writer, sheet_name="Stock Info", index=False)

print(f"✅ Stock data saved to: {output_file}")
