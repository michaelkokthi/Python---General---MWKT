import pandas as pd
import matplotlib.pyplot as plt
import wbdata
from fredapi import Fred
from matplotlib.backends.backend_pdf import PdfPages
from datetime import datetime

# Set up FRED API (Replace with your actual API key)
FRED_API_KEY = "xxx-xxx-xxx-xxx-xxx"
fred = Fred(api_key=FRED_API_KEY)

# Define date range
start_date = datetime(2000, 1, 1)
end_date = datetime.today()

# --- World Bank API ---
# Define macroeconomic indicators (Updated with correct indicators)
wb_indicators = {
    "NY.GDP.MKTP.CD": "GDP (current US$)",
    "FP.CPI.TOTL.ZG": "Inflation (%)",
    "SL.UEM.TOTL.ZS": "Unemployment Rate (%)",
    "BN.GSR.GNFS.CD": "Net Trade in Goods & Services (US$)"
}

# Define regions
regions = {
    "USA": "United States",
    "EUU": "European Union",
    "CHN": "China",
    "IND": "India",
    "EAS": "East Asia & Pacific"
}

# Fetch World Bank data for all regions
wb_data = {}
for region_code, region_name in regions.items():
    data = wbdata.get_dataframe(wb_indicators, country=region_code).dropna()
    data.index = pd.to_datetime(data.index)
    data = data[(data.index >= start_date) & (data.index <= end_date)]
    wb_data[region_name] = data

# --- FRED API --- (Only for USA & EU, as FRED does not have global data)
us_fred_data = {
    "Federal Funds Rate": fred.get_series("FEDFUNDS", start_date, end_date),
    "CPI Inflation": fred.get_series("CPIAUCSL", start_date, end_date),
    "10-Year Treasury Yield": fred.get_series("DGS10", start_date, end_date),
    "Unemployment Rate": fred.get_series("UNRATE", start_date, end_date),
    "Core Inflation": fred.get_series("CPILFESL", start_date, end_date),
    "Retail Sales Growth": fred.get_series("RSXFS", start_date, end_date),
    "Labor Force Participation Rate": fred.get_series("CIVPART", start_date, end_date),
    "Job Openings": fred.get_series("JTSJOL", start_date, end_date),
    "Yield Curve Spread (10Y - 2Y)": fred.get_series("DGS10", start_date, end_date) - fred.get_series("DGS2", start_date, end_date),
    "US Dollar Index": fred.get_series("DTWEXBGS", start_date, end_date),
    "Total Public Debt": fred.get_series("GFDEBTN", start_date, end_date),
    "Debt-to-GDP Ratio": fred.get_series("GFDEGDQ188S", start_date, end_date),
    "S&P 500 Index": fred.get_series("SP500", start_date, end_date)
}

# --- Save All Plots to PDF ---
with PdfPages("macroeconomic_report.pdf") as pdf:
    
    def plot_series(data, title, ylabel, color='blue'):
        if data is not None and not data.empty:
            plt.figure(figsize=(10,5))
            plt.plot(data, color=color, linewidth=2)
            plt.title(title)
            plt.xlabel("Year")
            plt.ylabel(ylabel)
            plt.grid(True)
            pdf.savefig()
            plt.close()

    # Plot for each region with improved labels
    for region_name, data in wb_data.items():
        for indicator in wb_indicators.values():
            if indicator in data.columns:
                plot_series(data[indicator], f"{region_name} - {indicator}", indicator, 'blue')
    
    # Plot US-specific FRED data with improved labels
    for title, series in us_fred_data.items():
        plot_series(series, f"USA - {title}", title, 'red')

print("Macroeconomic report saved as 'macroeconomic_report.pdf'")
