import requests
import pandas as pd
import openpyxl
import schedule
import time
from datetime import datetime

# API URL to fetch top 50 cryptocurrencies
API_URL = "https://api.coingecko.com/api/v3/coins/markets"
PARAMS = {
    "vs_currency": "usd",
    "order": "market_cap_desc",
    "per_page": 50,
    "page": 1,
    "sparkline": False
}

def fetch_crypto_data():
    """Fetches live cryptocurrency data from CoinGecko API."""
    response = requests.get(API_URL, params=PARAMS)
    if response.status_code == 200:
        return response.json()
    else:
        print("Error fetching data:", response.status_code)
        return []

def analyze_data(data):
    """Performs basic analysis on the top 50 cryptocurrencies."""
    df = pd.DataFrame(data)
    
    # Select required columns
    df = df[['name', 'symbol', 'current_price', 'market_cap', 'total_volume', 'price_change_percentage_24h']]
    
    # Rename columns for better readability
    df.columns = ['Cryptocurrency Name', 'Symbol', 'Current Price (USD)', 
                  'Market Capitalization', '24h Trading Volume', 'Price Change (24h %)']

    # Identify top 5 cryptocurrencies by Market Cap
    top_5 = df.nlargest(5, 'Market Capitalization')[['Cryptocurrency Name', 'Market Capitalization']]

    # Calculate the average price of the top 50
    avg_price = df['Current Price (USD)'].mean()

    # Find the highest & lowest 24-hour price change
    highest_change = df.loc[df['Price Change (24h %)'].idxmax()]
    lowest_change = df.loc[df['Price Change (24h %)'].idxmin()]

    # Print Analysis Summary
    print("\n🔹 Cryptocurrency Market Analysis 🔹")
    print("Top 5 Cryptos by Market Cap:\n", top_5)
    print(f"\nAverage Price of Top 50 Cryptos: ${avg_price:.2f}")
    print("\nHighest 24h Change:", highest_change[['Cryptocurrency Name', 'Price Change (24h %)']])
    print("\nLowest 24h Change:", lowest_change[['Cryptocurrency Name', 'Price Change (24h %)']])
    
    return df

def save_new_excel():
    """Fetches live data, analyzes it, and saves it as a new Excel file every 5 minutes."""
    print("\nFetching live cryptocurrency data... ⏳")

    data = fetch_crypto_data()
    if not data:
        return

    df = analyze_data(data)

    # Generate a new filename with timestamp
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    excel_filename = f"Crypto_Data_{timestamp}.xlsx"

    # Save to a new Excel file
    df.to_excel(excel_filename, index=False)

    print(f"✅ New Excel file created: {excel_filename} (Next update in 5 minutes)\n")

# Schedule the script to run every 5 minutes and create a new file
schedule.every(5).minutes.do(save_new_excel)

# Run the script initially and keep updating
save_new_excel()  # First Run
while True:
    schedule.run_pending()
    time.sleep(1)
