import os
import json
import pandas as pd
import time
from requests import Session
from openpyxl import load_workbook
from dotenv import load_dotenv
from requests.exceptions import ConnectionError, Timeout, TooManyRedirects
load_dotenv()

# API URL
url = 'https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest'
YOUR_API_KEY = os.getenv("CMC_API_KEY")

# API Parameters
parameters = {
    'start': '1',    # Start from the first cryptocurrency
    'limit': '50',   # Get top 50 cryptocurrencies
    'convert': 'INR' # Convert prices to INR
}

# Headers for authentication
headers = {
    'Accepts': 'application/json',
    'X-CMC_PRO_API_KEY': YOUR_API_KEY
}

# Initialize a session
session = Session()
session.headers.update(headers)

# Function to fetch cryptocurrency data
def fetch_crypto_data():
    try:
        response = session.get(url, params=parameters)
        data = response.json()

        if response.status_code != 200:
            print(f"Error fetching data: {data.get('status', {}).get('error_message')}")
            return None

        # Extract relevant data
        crypto_list = []
        for coin in data["data"]:
            crypto_list.append([
                coin["name"],                         # Cryptocurrency Name
                coin["symbol"],                       # Symbol
                coin["quote"]["INR"]["price"],        # Current Price (INR)
                coin["quote"]["INR"]["market_cap"],   # Market Capitalization
                coin["quote"]["INR"]["volume_24h"],   # 24h Trading Volume
                coin["quote"]["INR"]["percent_change_24h"] # 24h Price Change (%)
            ])
        
        return crypto_list

    except (ConnectionError, Timeout, TooManyRedirects) as e:
        print(f"Error: {e}")
        return None

# Function to analyze the data
def analyze_crypto_data(df):
    # Top 5 cryptocurrencies by market cap
    top_5_market_cap = df.nlargest(5, "Market Cap (INR)")

    # Average price of the top 50 cryptocurrencies
    avg_price = df["Current Price (INR)"].mean()

    # Highest & Lowest 24-hour percentage price change
    highest_change = df.loc[df["24h Price Change (%)"].idxmax()]
    lowest_change = df.loc[df["24h Price Change (%)"].idxmin()]

    # Create an analysis DataFrame
    analysis_df = pd.DataFrame({
        "Metric": [
            "Average Price of Top 50 Cryptos (INR)",
            "Highest 24h Change (%)",
            "Lowest 24h Change (%)",
            "Highest 24h Change Crypto",
            "Lowest 24h Change Crypto"
        ],
        "Value": [
            avg_price,
            highest_change["24h Price Change (%)"],
            lowest_change["24h Price Change (%)"],
            highest_change["Name"],
            lowest_change["Name"]
        ]
    })

    return top_5_market_cap, analysis_df

def update_excel():
    filename = "crypto_data.xlsx"

    while True:
        print("Fetching new cryptocurrency data...")
        crypto_data = fetch_crypto_data()

        if crypto_data is None:
            print("Skipping update due to API error.")
            time.sleep(300)  # Wait 5 minutes before retrying
            continue

        # Create DataFrame
        df = pd.DataFrame(crypto_data, columns=[
            "Name", "Symbol", "Current Price (INR)", "Market Cap (INR)", 
            "24h Trading Volume (INR)", "24h Price Change (%)"
        ])

        # Perform Data Analysis
        top_5_market_cap, analysis_df = analyze_crypto_data(df)

        # Save to Excel (overwrite the existing sheet)
        try:
            with pd.ExcelWriter(filename, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
                df.to_excel(writer, sheet_name="Live Crypto Data", index=False)
                top_5_market_cap.to_excel(writer, sheet_name="Top 5 Market Cap", index=False)
                analysis_df.to_excel(writer, sheet_name="Crypto Analysis", index=False)

        except FileNotFoundError:
            with pd.ExcelWriter(filename, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Live Crypto Data", index=False)
                top_5_market_cap.to_excel(writer, sheet_name="Top 5 Market Cap", index=False)
                analysis_df.to_excel(writer, sheet_name="Crypto Analysis", index=False)

        print("Excel updated. Waiting 5 minutes before next update...")
        time.sleep(300)  # Wait 5 minutes

# Run the script
update_excel()
