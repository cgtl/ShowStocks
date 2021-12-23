import requests
from bs4 import BeautifulSoup
import pandas as pd 


# BS4
baseurl = "https://finance.yahoo.com/most-active"
r = requests.get(baseurl, timeout=10)
soup = BeautifulSoup(r.content, "lxml")


# Stock
stock_list = []
for x in range(74, 841, 32):
    stock = soup.find("td", {f"data-reactid": {x}})
    stock = stock.get_text(separator=" ").strip()
    stock_list.append(stock)

# Company
company_list = []
for x in range(81, 849, 32):
    company = soup.find("td", {f"data-reactid": {x}})
    company = company.get_text(separator=" ").strip()
    company_list.append(company)

# Price $
price_list = []
for x in range(83, 851, 32):
    price = soup.find("td", {f"data-reactid": {x}})
    price = price.get_text(separator=" ").strip()
    price = "$ " + price
    price_list.append(price)

# Change % 
change_list = []
for x in range(90, 858, 32):
    change = soup.find("td", {f"data-reactid": {x}})
    change = change.get_text(separator=" ").strip()
    change_list.append(change)


# Export to xlsx
list_dict = {'Stock':stock_list, 'Company':company_list, "Price":price_list, "Change":change_list} 
df = pd.DataFrame(list_dict)
df.to_excel("stocks.xlsx", index=False, sheet_name="Stocks") 

# Change column width
writer = pd.ExcelWriter("stocks.xlsx") 
df.to_excel(writer, sheet_name='Stocks', index=False, na_rep='NaN')
for column in df:
    column_length = max(df[column].astype(str).map(len).max(), len(column))
    col_idx = df.columns.get_loc(column)
    writer.sheets['Stocks'].set_column(col_idx, col_idx, column_length)
writer.save()
