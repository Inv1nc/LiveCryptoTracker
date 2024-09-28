import requests, json, datetime
import pandas as pd
from time import sleep

def main():
    while True:
        print(f"Time# {datetime.datetime.now()}")
        print("Fetching Data")
        df = fetch_data(50)

        if df is not None and not df.empty:
            print("Saving Data")
            save_into_xls(df)
            print("Data saved successfully.")

        else:
            print("Failed to fetch data or no data returned.")

        print("-"*10)
        sleep(300)  # 300 seconds - 5 mins

def fetch_data(count):

    url = "https://api.coingecko.com/api/v3/coins/markets"

    params = {
        'vs_currency' : 'usd',
        'order' : 'market_cap_desc',
        'per_page' : count,
        'page' : 1,
        'sparkline' : False,
    }

    crypto_data = []

    try:
        response = requests.get(url, params=params)
        data = response.json()

        for currency in data:
            crypto_data.append({
                'Name': currency['name'],
                'Symbol': currency['symbol'],
                'Current Price (USD)': currency['current_price'],
                'Market Cap (USD)': currency['market_cap'],
                '24-hour Volume (USD)': currency['total_volume'],
                'Price Change (24h, USD)' : currency['price_change_24h'],
                'Price Change (24h, %)' : currency['price_change_percentage_24h'],
                'Circulating Supply' : currency['circulating_supply'],
                'All-Time High (USD)' : currency['ath'],
                'ATH Change %' : currency['ath_change_percentage'],
                'All-Time Low (USD)' : currency['atl'],
                'ATL Change %' : currency['atl_change_percentage'],
            })

    except Exception as err:
        print(f"Failed to Fetch the Data.")

    else:
        return pd.DataFrame(crypto_data)

def save_into_xls(df):

    try:
        with pd.ExcelWriter("output.xlsx", engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False)

            col_widths = [max(df[col].astype(str).map(len).max(), len(col)) + 2 for col in df.columns]

            workbook = writer.book
            worksheet = writer.sheets[list(writer.sheets.keys())[0]]

            for i, width in enumerate(col_widths):
                worksheet.set_column(i, i, width)

    except Exception as error:
        print(f"Error while saving to Excel: {err}")

if __name__ == "__main__":
    main()
