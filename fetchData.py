import requests
import pandas as pd
import schedule
import time

def fetch_data():
    url = 'https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=market_cap_desc&per_page=50&page=1'
    response = requests.get(url)
    data = response.json()  # Get the JSON data from the API
    return data

data = fetch_data()  # Fetch data

# print(data)

df = pd.DataFrame(data)  # Convert data into a pandas DataFrame
top_5 = df[['name', 'symbol', 'market_cap']].sort_values(by='market_cap', ascending=False).head(5)
print(top_5)

average_price = df['current_price'].mean()
print(f'Average Price: ${average_price}')

highest_change = df.loc[df['price_change_percentage_24h'].idxmax()]
lowest_change = df.loc[df['price_change_percentage_24h'].idxmin()]
print(f'Highest Change: {highest_change["name"]}, {highest_change["price_change_percentage_24h"]}%')
print(f'Lowest Change: {lowest_change["name"]}, {lowest_change["price_change_percentage_24h"]}%')

with pd.ExcelWriter('cryptocurrency_data.xlsx', engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Crypto Data')




def update_excel():
    data = fetch_data()  # Fetch fresh data
    df = pd.DataFrame(data)  # Convert data to DataFrame
    with pd.ExcelWriter('cryptocurrency_data.xlsx', engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Crypto Data')
    print('Excel sheet updated!')

# Schedule to run every 5 minutes
schedule.every(5).minutes.do(update_excel)

while True:
    schedule.run_pending()
    time.sleep(1)  # Wait for 1 second before checking again


