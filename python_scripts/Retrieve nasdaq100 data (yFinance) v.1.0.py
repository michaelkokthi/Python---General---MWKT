import os
import yfinance as yf
from tqdm import tqdm # type: ignore
import pandas as pd
from datetime import datetime

# ✅ Scrape Nasdaq-100 table from Wikipedia
url = "https://en.wikipedia.org/wiki/NASDAQ-100"
nasdaq_tables = pd.read_html(url)

# ✅ Get the correct table (may change index if Wikipedia updates structure)
nasdaq_df = nasdaq_tables[4]  # Adjust index if Wikipedia changes the structure
nasdaq_df.rename(columns={"Ticker": "Symbol"}, inplace=True)

# ✅ Clean the ticker symbols (replace '.' with '-' for Yahoo Finance compatibility)
nasdaq_df["Symbol"] = nasdaq_df["Symbol"].str.replace(".", "-")

# ✅ Convert symbols to a list
nasdaq_symbols = nasdaq_df["Symbol"].tolist()

print(f"✅ Retrieved {len(nasdaq_symbols)} Nasdaq-100 tickers.")
print(nasdaq_symbols[:10])  # Display first 10 symbols for verification

# ✅ Convert Unix timestamps into readable date format
def convert_timestamp(timestamp):
    if isinstance(timestamp, (int, float)) and timestamp > 0:
        return datetime.utcfromtimestamp(timestamp).strftime("%d-%m-%Y")
    return "N/A"

# ✅ Function to retrieve stock data
def get_stock_data(ticker):
    try:
        stock = yf.Ticker(ticker)
        info = stock.info  # Get stock details

        return {
            "Asset Class": info.get("quoteType", "N/A"),
            "Exchange": info.get("fullExchangeName", "N/A"),
            "Currency": info.get("currency", "N/A"),
            "Symbol": ticker,
            "Company Name": info.get("shortName", "N/A"),
            "Current Price": info.get("currentPrice", "N/A"),
            "Market Cap": info.get("marketCap", "N/A"),
            "PE Ratio (TTM)": info.get("trailingPE", "N/A"),
            "Beta (5Y Monthly)": info.get("beta", "N/A"),
            "EPS (Trailing)": info.get("trailingEps", "N/A"),
            "EPS (Forward)": info.get("forwardEps", "N/A"), #Expected EPS for Next Quarter.
            "EPS (Current Year)": info.get("epsCurrentYear", "N/A"), #EPS for Current Quarter.
            "Sector": info.get("sector", "N/A"),
            "Industry": info.get("industry", "N/A"),

            # ✅ Market Prices & Changes
            "Regular Market Price": info.get("regularMarketPrice", "N/A"),
            "Regular Market Change Percent": info.get("regularMarketChangePercent", "N/A"),
            "Post Market Change %": info.get("postMarketChangePercent", "N/A"),
            "Post Market Price": info.get("postMarketPrice", "N/A"),  # Stock price in post-market trading
            "Post Market Change": info.get("postMarketChange", "N/A"),  # Change in stock price after market close
            "Regular Market Change": info.get("regularMarketChange", "N/A"),  # Change in stock price during regular trading
            "Regular Market Day Range": info.get("regularMarketDayRange", "N/A"),  # Low-High range during the trading day

            # ✅ 52-Week Performance
            "52W Low": info.get("fiftyTwoWeekLow", "N/A"),
            "52W Low Change": info.get("fiftyTwoWeekLowChange", "N/A"),  # Price change from 52-week low
            "52W Low Change %": info.get("fiftyTwoWeekLowChangePercent", "N/A"),  # % change from 52-week low
            "52W Range": info.get("fiftyTwoWeekRange", "N/A"),  # 52-week Low-High range
            "52W High": info.get("fiftyTwoWeekHigh", "N/A"),
            "52W High Change": info.get("fiftyTwoWeekHighChange", "N/A"),  # Price change from 52-week high
            "52W High Change %": info.get("fiftyTwoWeekHighChangePercent", "N/A"),  # % change from 52-week high
            "52W Change %": info.get("fiftyTwoWeekChangePercent", "N/A"),  # % change over the last 52 weeks

            # ✅ Moving Averages & Trends
            "50-Day Moving Average": info.get("fiftyDayAverage", "N/A"),
            "50-Day Avg Change": info.get("fiftyDayAverageChange", "N/A"),  # Change from 50-day moving average
            "50-Day Avg Change %": info.get("fiftyDayAverageChangePercent", "N/A"),  # % change from 50-day MA
            "200-Day Moving Average": info.get("twoHundredDayAverage", "N/A"),
            "200-Day Avg Change": info.get("twoHundredDayAverageChange", "N/A"),  # Change from 200-day MA
            "200-Day Avg Change %": info.get("twoHundredDayAverageChangePercent", "N/A"),  # % change from 200-day MA

            # ✅ Volume 
            "Volume": info.get("volume", "N/A"),
            "Regular Market Volume": info.get("regularMarketVolume", "N/A"),
            "Average Volume": info.get("averageVolume", "N/A"),
            "Average Volume (10 days)": info.get("averageVolume10days", "N/A"),
            "Average Daily Volume (10 days)": info.get("averageDailyVolume10Day", "N/A"),
            
            # ✅ Stock Trading Metrics
            "Bid": info.get("bid", "N/A"),
            "Ask": info.get("ask", "N/A"),
            "Bid Size": info.get("bidSize", "N/A"),
            "Ask Size": info.get("askSize", "N/A"),

            # ✅ Growth Estimates
            "Earnings Growth": info.get("earningsGrowth", "N/A"),
            "Revenue Growth": info.get("revenueGrowth", "N/A"),
            "Earnings Quarterly Growth": info.get("earningsQuarterlyGrowth", "N/A"),
            "Trailing PEG Ratio": info.get("trailingPegRatio", "N/A"), 
            #PEG < 1: Undervalued 
            #PEG = 1: Fair Value
            #PEG > 1: Overvalued

            # ✅ Revenue Margins
            "Net Profit Margin": info.get("profitMargins", "N/A"),
            "Gross Margins": info.get("grossMargins", "N/A"),
            "EBITDA Margins": info.get("ebitdaMargins", "N/A"),
            "Operating Margins": info.get("operatingMargins", "N/A"),

            # ✅ Additional financial ratios [Valuation, Leverage]
            "Enterprise Value": info.get("enterpriseValue", "N/A"),
            "EBITDA": info.get("ebitda", "N/A"),
            "EV/EBITDA": info.get("enterpriseToEbitda", "N/A"),
            "Enterprise to Revenue": info.get("enterpriseToRevenue", "N/A"),
            "Book Value": info.get("bookValue", "N/A"),
            "Price/Book Ratio": info.get("priceToBook", "N/A"),
            "Debt/Equity Ratio": info.get("debtToEquity", "N/A"),
            "ROE (Return on Equity)": info.get("returnOnEquity", "N/A"),
            "ROA (Return on Assets)": info.get("returnOnAssets", "N/A"),
            
            # ✅ Valuation Metrics
            "Total Cash": info.get("totalCash", "N/A"),
            "Total Debt": info.get("totalDebt", "N/A"),
            "Free Cash Flow": info.get("freeCashflow", "N/A"), #Levered Free Cash Flow (Cash after servicing Debt).
            "Operating Cash Flow": info.get("operatingCashflow", "N/A"),
            "Total Debt": info.get("totalDebt", "N/A"),
            "Quick Ratio": info.get("quickRatio", "N/A"),
            "Current Ratio": info.get("currentRatio", "N/A"),
            "Total Revenue": info.get("totalRevenue", "N/A"),
            "Gross Profits": info.get("grossProfits", "N/A"),
            "Cash Per Share": info.get("totalCashPerShare", "N/A"),
            "Revenue per Share": info.get("revenuePerShare", "N/A"),

            # ✅ Risk Metrics
            "Compensation Timestamp": convert_timestamp(info.get("compensationAsOfEpochDate")),
            "Compensation Risk": info.get("compensationRisk", "N/A"), #Executive Pay Fairness
            "Governance Timestamp": convert_timestamp(info.get("governanceEpochDate")),
            "Audit Risk": info.get("auditRisk", "N/A"), #Transparency of Financials
            "Board Risk": info.get("boardRisk", "N/A"), #Effectiveness & Independence of the Board
            "Shareholder Rights Risk": info.get("shareHolderRightsRisk", "N/A"), #Investor Protection
            "Overall Risk": info.get("overallRisk", "N/A"), #Combined Score

            # ✅ Analyst Ratings & Price Targets
            "Target High Price": info.get("targetHighPrice", "N/A"), #Highest analyst price target
            "Target Low Price": info.get("targetLowPrice", "N/A"), #Lowest analyst price target
            "Target Mean Price": info.get("targetMeanPrice", "N/A"),#Average analyst price target
            "Target Median Price": info.get("targetMedianPrice", "N/A"),
            "Recommendation Mean": info.get("recommendationMean", "N/A"),
            "Recommendation Key": info.get("recommendationKey", "N/A"),
            "Number of Analyst Opinions": info.get("numberOfAnalystOpinions", "N/A"), 
            "Average Analyst Rating": info.get("averageAnalystRating", "N/A"),

            # ✅ Stock Splits & Dividends
            "Payout Ratio": info.get("payoutRatio", "N/A"), #Payout Ratio measures how much of a company's earnings (net income) is paid out as dividends to shareholders. 
            "Dividend Rate": info.get("dividendRate", "N/A"), #Annual dividend payout per share
            "Dividend Yield": info.get("dividendYield", "N/A"),
            "Five-Year Avg Dividend Yield": info.get("fiveYearAvgDividendYield", "N/A"),
            "Last Split Factor": info.get("lastSplitFactor", "N/A"),
            "Last Split Date": convert_timestamp(info.get("lastSplitDate")),
            "Last Dividend Value": info.get("lastDividendValue", "N/A"),
            "Last Dividend Date": convert_timestamp(info.get("lastDividendDate")),
            "Ex-Dividend Date": convert_timestamp(info.get("exDividendDate")), #Last Day to Qualify for Dividend Payout
            "Trailing Annual Dividend Rate": info.get("trailingAnnualDividendRate", "N/A"),
            "Trailing Annual Dividend Yield": info.get("trailingAnnualDividendYield", "N/A"),
            
            # ✅ Date Fields (Converted)
            "Earnings Timestamp": convert_timestamp(info.get("earningsTimestamp")),
            "Earnings Timestamp Start": convert_timestamp(info.get("earningsTimestampStart")),
            "Earnings Timestamp End": convert_timestamp(info.get("earningsTimestampEnd")),
            "Earnings Call Start": convert_timestamp(info.get("earningsCallTimestampStart")),
            "Earnings Call End": convert_timestamp(info.get("earningsCallTimestampEnd")),
            "First Trade Date": convert_timestamp(info.get("firstTradeDateMilliseconds") / 1000),
            "Last Fiscal Year End": convert_timestamp(info.get("lastFiscalYearEnd")),
            "Next Fiscal Year End": convert_timestamp(info.get("nextFiscalYearEnd")),
            "Most Recent Quarter": convert_timestamp(info.get("mostRecentQuarter")),

            # ✅ Share Structure
            "Float Shares": info.get("floatShares", "N/A"), # 🟢 Total number of shares available for trading in the public market. 
            "Shares Outstanding": info.get("sharesOutstanding", "N/A"), # 🟢 Total number of shares issued by the company. 
            "Shares Short": info.get("sharesShort", "N/A"),  # 🔴 Number of shares currently sold short (borrowed and sold, but not yet repurchased).
            "Shares Short Prior Month": info.get("sharesShortPriorMonth", "N/A"),  # 🔴 Number of shares that were shorted in the prior reporting month.
            "Shares Short Previous Month Date": convert_timestamp(info.get("sharesShortPreviousMonthDate")),   # 🔴 The date of the previous month’s short interest report.
            "Date Short Interest": convert_timestamp(info.get("dateShortInterest")),  # 🔴 The most recent short interest reporting date.
            "Shares Percent Shares Out": info.get("sharesPercentSharesOut", "N/A"),  # 🔴 Percentage of total outstanding shares that are shorted.
            "Held Percent Insiders": info.get("heldPercentInsiders", "N/A"),   # 🟢 Percentage of shares held by **insiders** (e.g., executives, board members).
            "Held Percent Institutions": info.get("heldPercentInstitutions", "N/A"),  # 🟢 Percentage of shares held by **institutional investors** (e.g., mutual funds, pension funds).
            "Short Ratio (Days to Cover)": info.get("shortRatio", "N/A"),  # 🔴 **Short Ratio (Days to Cover)**: Number of days it would take to cover all short positions, [Shares Short / Average Daily Volume].
            "Short Percent of Float": info.get("shortPercentOfFloat", "N/A"),  # 🔴 Percentage of the float (publicly tradable shares) that are shorted.
            "Implied Shares Outstanding": info.get("impliedSharesOutstanding", "N/A"), # 🟢 Total **theoretical shares outstanding** including convertible securities, options, etc. 

            # ✅ Business Summary
            "Business Summary": info.get("longBusinessSummary", "N/A"),
        }
    except Exception as e:
        print(f"❌ Error retrieving {ticker}: {e}")
        return {"Symbol": ticker, "Company Name": "Error"}  # Default error data

# ✅ Fetch data for all Nasdaq-100 stocks
nasdaq_data = []
for ticker in tqdm(nasdaq_symbols, desc="Fetching Nasdaq-100 Data", unit="stock"):
    nasdaq_data.append(get_stock_data(ticker))

# ✅ Convert data to DataFrame
nasdaq_df = pd.DataFrame(nasdaq_data)

# ✅ Get today's date in YYYY-MM-DD format (For File Naming Convention)
current_date = datetime.today().strftime("%d-%m-%Y")

# ✅ Save to Excel in Downloads Folder
downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
output_file = os.path.join(downloads_folder, f"Nasdaq100_StockData_{current_date}.xlsx")

with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    nasdaq_df.to_excel(writer, sheet_name="Nasdaq-100 Data", index=False)

print(f"✅ Stock data saved to: {output_file}")
