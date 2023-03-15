import requests
from bs4 import BeautifulSoup
import pandas as pd

# URL of the page to be scraped
url = 'https://finance.naver.com/sise/lastsearch2.naver'

# Send a GET request to the URL
res = requests.get(url)

# Parse the HTML content using BeautifulSoup
soup = BeautifulSoup(res.text, 'html.parser')

# Find the table element that contains the data
table = soup.find('table', {'class': 'type_5'})

# Find all rows of the table, excluding the first row (header)
rows = table.find_all('tr')[1:]

# Initialize an empty list to store the data
data = []

# Loop through each row and extract the data
for row in rows:
    cells = row.find_all('td')
    if len(cells) < 12:
        continue
    rank = cells[0].text.strip()
    name = cells[1].text.strip()
    search_ratio = cells[2].text.strip()
    current_price = cells[3].text.strip()
    price_change = cells[4].text.strip()
    percent_change = cells[5].text.strip()
    volume = cells[6].text.strip()
    open_price = cells[7].text.strip()
    high_price = cells[8].text.strip()
    low_price = cells[9].text.strip()
    per = cells[10].text.strip()
    roe = cells[11].text.strip()
    
    # Append the extracted data to the list
    data.append([rank, name, search_ratio, current_price, price_change,
                 percent_change, volume, open_price, high_price, low_price,
                 per, roe])

# Convert the list of data to a Pandas DataFrame
df = pd.DataFrame(data, columns=['Rank', 'Name', 'Search Ratio', 'Current Price',
                                 'Price Change', 'Percent Change', 'Volume',
                                 'Open Price', 'High Price', 'Low Price',
                                 'PER', 'ROE'])

# Save the DataFrame to an Excel file
df.to_excel('stock_data.xlsx', index=False)
