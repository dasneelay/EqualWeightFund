import pandas as pd
import requests
import xlsxwriter
import math

'''
Script will accept the value of your portfolio and return Trade Sheet with how many shares of each fund 
holding to purchase to get an equal-weight version of the fund.

'''

IEX_CLOUD_API_TOKEN = ''

# Read fund holdings
stocks = pd.read_csv('~/Documents/projects/ARKproject/arkg_holdings.csv')


def chunks(lst, n):
    # yields successive n-sized chunks from lst
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

# Create list of tickers with total length 1
symbol_groups = list(chunks(stocks['Ticker'], 500))

# Separate list by commas, but keep them as one item
symbol_strings = []
for symb in symbol_groups:
    symbol_strings.append(','.join(symb))

print(symbol_strings)
# Create Columns of Trade Sheet
col_names = ['Ticker', 'Share Price', 'Market Cap', 'Shares to Buy']
final_df = pd.DataFrame(columns=col_names)

for symb_list in symbol_strings:
    batch_api_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={symb_list}&types=quote&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call_url).json()
    for symbol in symb_list.split(','):
        final_df = final_df.append(
        pd.Series([
            symbol,
            data[symbol]['quote']['latestPrice'],
            data[symbol]['quote']['marketCap'],
            'N/A'],
        index=col_names),
        ignore_index=True)


# Calculate  number of shares to buy

portfolio_size = float(input('Enter the value of your portfolio: '))

# Calculate how much cash each position should be worth

position_size = portfolio_size/len(final_df.index)

for i in range(0, len(final_df.index)):
    final_df.loc[i, 'Shares to Buy'] = math.floor(position_size/final_df.loc[i, 'Share Price'])


# Initialize XlsxWriter Object

writer = pd.ExcelWriter('/home/neelaydas/Documents/projects/ARKproject/TradeSheet.xlsx', engine='xlsxwriter') # pylint: disable=abstract-class-instantiated
final_df.to_excel(writer, 'Recommended Trades', index=False)

# Creating format objects
dollar_format = writer.book.add_format(
        {
            'num_format':'$0.00',
            'border': 1
        }
    )

string_int_format = writer.book.add_format(
        {
        'border': 1
        }
    )

formats = {
    'A': string_int_format,
    'B': dollar_format,
    'C': dollar_format,
    'D': string_int_format
    }

# Format each column using assigned formats

for col in formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{col}:{col}', 20, formats[col])

# Save Excel file

writer.save()
print(final_df, '\n Check the file folder for your Trade Sheet!')
