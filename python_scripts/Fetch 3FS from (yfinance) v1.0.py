import yfinance as yf
import pandas as pd
import os

# ✅ Fetch data directly from Yahoo Finance
ticker = "PONY"  # CHANGE THIS TO ANY TICKER YOU WANT
stock = yf.Ticker(ticker)

# ✅ Get company name for dynamic filename
company_name = stock.info.get("shortName", ticker)  # Use ticker if name is missing
safe_company_name = "".join(c if c.isalnum() or c in (" ", "_") else "_" for c in company_name)  # Clean filename

# ✅ Fetch full financials (Annual)
income_stmt = stock.financials  # Income Statement (Annual)
balance_sheet = stock.balance_sheet  # Balance Sheet (Annual)
cash_flow = stock.cashflow  # Cash Flow Statement (Annual)

# ✅ Fetch TTM Data (Summing last 4 quarters)
income_stmt_ttm = stock.quarterly_financials.sum(axis=1)  # TTM Income Statement
balance_sheet_ttm = stock.quarterly_balance_sheet.sum(axis=1)  # TTM Balance Sheet
cash_flow_ttm = stock.quarterly_cashflow.sum(axis=1)  # TTM Cash Flow Statement

# ✅ Function to format each statement correctly
def format_statement(df, ttm_series, statement_name):
    """Format the financial statement for proper Excel formatting with TTM included."""
    df = df.copy()  # Avoid modifying original data
    df.insert(0, "TTM", ttm_series)  # Add TTM as the first column
    df.index.name = statement_name  # Set index name as statement title
    df.reset_index(inplace=True)  # Move row names to a column
    return df

# ✅ Format all statements properly, including TTM
income_stmt = format_statement(income_stmt, income_stmt_ttm, "Income Statement")
balance_sheet = format_statement(balance_sheet, balance_sheet_ttm, "Balance Sheet")
cash_flow = format_statement(cash_flow, cash_flow_ttm, "Cash Flow Statement")

#Saving files in Downloads
# ✅ Get the user's Downloads folder dynamically
downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")

# ✅ Define dynamic output filename
#Note that if we remove the downloads_folder, file will be saved under C: drive / Users
output_file = os.path.join(downloads_folder, f"{safe_company_name}_Financials.xlsx")

# ✅ Save to Excel
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    income_stmt.to_excel(writer, sheet_name="Income Statement", index=False)
    balance_sheet.to_excel(writer, sheet_name="Balance Sheet", index=False)
    cash_flow.to_excel(writer, sheet_name="Cash Flow Statement", index=False)

print(f"✅ Excel file saved: {output_file}")
