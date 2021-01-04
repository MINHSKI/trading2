''' 
    Reference: https://www.youtube.com/watch?v=xfzGZB4HhEE&t=8170s
    Author: Nick McCullum:
    Algorithmic Trading Using Python
        Momentum Trading S&P 500 Strategy
    API Token: IEX Cloud API token, the data provider of stock market data
'''
import numpy as np
import pandas as pd
import requests
import math
from scipy import stats
from secrets import IEX_TEST_API_TOKEN as TOKEN

PORTFOLIO_SIZE = 10000000
TOP_XX_STOCKS = 50
MY_COLUMNS = ['Ticker', 'Company Name', 'Stock Price', 'One-Year Return', 'Number of Shares to Buy']

def fetch_data():
    stocks = pd.read_csv('sp_500_stocks.csv')
    smaller_chunks = np.array_split(stocks['Ticker'], 6)
    final_dataframe = pd.DataFrame(columns=MY_COLUMNS)
    position_size = math.floor(PORTFOLIO_SIZE/TOP_XX_STOCKS)
    for stocks_chunk in smaller_chunks:
        stocks_list = ''
        stocks_list = ','.join(stocks_chunk)
        batch_api_url = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={stocks_list}&types=price,stats&token={TOKEN}'
        try:
            req_result = requests.get(batch_api_url)
            # print(req_result.status_code)
            data = req_result.json()
            for symbol in stocks_chunk:
                stock_price = data[symbol]['price']
                yearChange = data[symbol]['stats']['year1ChangePercent']
                name = data[symbol]['stats']['companyName']
                shares_to_buy = math.floor(position_size/stock_price)
                final_dataframe = final_dataframe.append(
                    pd.Series(
                        [
                            symbol,
                            name,
                            stock_price,
                            yearChange,
                            shares_to_buy
                        ],
                        index = MY_COLUMNS
                    ),
                    ignore_index = True
                )
                # print(f'Ticker: {symbol} Name: {name} change is {round(yearChange*100,2)}')
        except:
            print('Houston, we got a problem!')


    final_dataframe.sort_values('One-Year Return', ascending=False, inplace=True)
    # just return the top 50 high performance stocks
    final_dataframe = final_dataframe[:TOP_XX_STOCKS]
    final_dataframe.reset_index(drop = True, inplace = True)
    return final_dataframe



def createExcel(final_dataframe):
    writer = pd.ExcelWriter('Portfolio_Momentum_Trades.xlsx', engine='xlsxwriter')
    final_dataframe.to_excel(writer, 'Recommended Trades', index=False)
    # background_color = '#0a0a23'
    background_color = '#e8eaf6'
    font_color = '#000000'
    string_format = writer.book.add_format(
        {
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

    string_name_format = writer.book.add_format(
        {
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

    dollar_format = writer.book.add_format(
        {
            'num_format': '$0.00',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

    integer_format = writer.book.add_format(
        {
            'num_format': '0',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

    percent_format = writer.book.add_format(
        {
            'num_format': '0.00%',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

    column_formats = {
        'A': ['Ticker', string_format],
        'B': ['Company Name', string_format],
        'C': ['Stock Price', dollar_format],
        'D': ['One-Year Return', percent_format],
        'E': ['Number of Shares to Buy', integer_format],
    }

    for column in column_formats.keys():
        writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 18, column_formats[column][1])
        writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], column_formats[column][1])

    writer.save()


def main():
    final_dataframe = fetch_data()
    createExcel(final_dataframe)
    print(final_dataframe)

if __name__ == "__main__":
    main()

