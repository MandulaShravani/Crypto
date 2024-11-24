import requests
import pandas as pd
from openpyxl import Workbook  # type: ignore

# Step 1: Fetch Live Cryptocurrency Data
def fetch_crypto_data():
    url = "https://www.coingecko.com/"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": False
    }
    response = requests.get(url, params=params)
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Error fetching data: {response.status_code}")
        return []

# Step 2: Process Data into a DataFrame
def process_data(data):
    df = pd.DataFrame(data)
    df = df[["name", "symbol", "current_price", "market_cap", "total_volume", "price_change_percentage_24h"]]
    df.columns = ["Cryptocurrency Name", "Symbol", "Current Price (USD)", "Market Cap (USD)", 
                  "24h Trading Volume (USD)", "24h Price Change (%)"]
    return df

# Step 3: Analyze Data
def analyze_data(df):
    # Top 5 Cryptocurrencies by Market Cap
    top_5 = df.nlargest(5, "Market Cap (USD)")
    
    # Average Price of Top 50 Cryptocurrencies
    avg_price = df["Current Price (USD)"].mean()
    
    # Highest and Lowest 24h Percentage Change
    highest_change = df.nlargest(1, "24h Price Change (%)")
    lowest_change = df.nsmallest(1, "24h Price Change (%)")
    
    return top_5, avg_price, highest_change, lowest_change

# Step 4: Export to Excel
def export_to_excel(df, top_5, avg_price, highest_change, lowest_change):
    with pd.ExcelWriter("Crypto_Analysis.xlsx", engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Top 50 Cryptos", index=False)
        top_5.to_excel(writer, sheet_name="Top 5 Cryptos", index=False)
        
        # Add Analysis Summary
        summary_data = {
            "Metric": ["Average Price (USD)", "Highest 24h Change (%)", "Lowest 24h Change (%)"],
            "Value": [avg_price, highest_change["24h Price Change (%)"].values[0], lowest_change["24h Price Change (%)"].values[0]]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

# Main Script
if __name__ == "__main__":
    # Fetch and Process Data
    data = fetch_crypto_data()
    if data:
        crypto_df = process_data(data)
        
        # Analyze Data
        top_5_df, avg_price, highest_change_df, lowest_change_df = analyze_data(crypto_df)
        
        # Export Results to Excel
        export_to_excel(crypto_df, top_5_df, avg_price, highest_change_df, lowest_change_df)
        
        print("Analysis complete. Results saved to 'Crypto_Analysis.xlsx'.")
