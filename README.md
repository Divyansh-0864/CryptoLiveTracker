# 📊 Live Cryptocurrency Data Fetcher  

This project fetches live cryptocurrency data for the top 50 cryptocurrencies using the CoinMarketCap API, performs basic analysis, and continuously updates an Excel sheet every 5 minutes.

## 🚀 Features  
- Fetches real-time data (Name, Symbol, Price, Market Cap, Volume, and 24h Change).  
- Identifies the **Top 5 cryptocurrencies** by market cap.  
- Computes the **average price** of the top 50 cryptocurrencies.  
- Finds the **highest and lowest 24-hour percentage change**.  
- Updates an Excel sheet **every 5 minutes** with live data.

## 🛠 Installation  
1. Clone this repository:  
   ```bash
   git clone https://github.com/Divyansh-0864/CryptoLiveTracker.git
   cd crypto-live-tracker
   ```
2. Install dependencies:  
   ```bash
   pip install -r requirements.txt
   ```

## 🔑 API Key Setup  
1. Get a free API key from [CoinMarketCap](https://coinmarketcap.com/api/).  
2. Replace `"YOUR_API_KEY"` in the script with your actual API key.

## ▶️ Usage  
Run the script to start fetching and updating data:  
```bash
python crypto_tracker.py
```

## 📂 Output  
- A continuously updating **Excel file** (`crypto_data.xlsx`) containing live cryptocurrency stats.  
- Console logs showing real-time analysis results.

