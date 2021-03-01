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
from statistics import mean
from scipy.stats import percentileofscore as score

from secrets import IEX_TEST_API_TOKEN as TOKEN

PORTFOLIO_SIZE = 10000000
TOP_XX_STOCKS = 50
MY_COLUMNS = ['Ticker', 'Company Name', 'Stock Price', 'One-Year Return', 'Number of Shares to Buy']

def fetch_1y_data():
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




# ############################################################################
# 
# Here's the meat of the program.  Fetching percentile of 1 yr, 6 months 
# 3 months, 1 month return percentile to calculate the HQM Score percent
# 
# ############################################################################

'''
    Create a list of 50 High Quality Momentum stocks
'''
def fetch_hqm():
    hqm_columns = [
        'Ticker',
        'Company Name',
        'Price',
        'Shares to Buy',
        'HQM Score',
        'One-Year Price Return',
        'One-Year Return Percentile',
        'Six-Month Price Return',
        'Six-Month Return Percentile',
        'Three-Month Price Return',
        'Three-Month Return Percentile',
        'One-Month Price Return',
        'One-Month Return Percentile'
    ]

    stocks = pd.read_csv('sp_500_stocks.csv')
    smaller_chunks = np.array_split(stocks['Ticker'], 10)
    hqm_dataframe = pd.DataFrame(columns=hqm_columns)
    position_size = math.floor(PORTFOLIO_SIZE/TOP_XX_STOCKS)
    # for stocks_chunk in smaller_chunks[:2]:
    for stocks_chunk in smaller_chunks:
        stocks_list = ''
        stocks_list = ','.join(stocks_chunk)
        batch_api_url = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={stocks_list}&types=price,stats&token={TOKEN}'
        try:
            req_result = requests.get(batch_api_url)
            print(req_result.status_code)
            data = req_result.json()
            print(data.keys())
            for symbol in stocks_chunk:
                company_name = data[symbol]['stats']['companyName']
                stock_price = data[symbol]['price']
                shares_to_buy = math.floor(position_size/stock_price)
                hqmScore = 'N/A'
                year1PriceChangePercent = data[symbol]['stats']['year1ChangePercent']
                year1ReturnPercent = 'N/A'
                month6PriceChangePercent = data[symbol]['stats']['month6ChangePercent']
                month6ReturnPercent = 'N/A'
                month3PriceChangePercent = data[symbol]['stats']['month3ChangePercent']
                month3ReturnPercent = 'N/A'
                month1PriceChangePercent = data[symbol]['stats']['month1ChangePercent']
                month1ReturnPercent = 'N/A'
                hqm_dataframe = hqm_dataframe.append(
                    pd.Series(
                        [
                            symbol,
                            company_name,
                            stock_price,
                            shares_to_buy,
                            hqmScore,
                            year1PriceChangePercent,
                            month1ReturnPercent,
                            month6PriceChangePercent,
                            month6ReturnPercent,
                            month3PriceChangePercent,
                            month3ReturnPercent,
                            month1PriceChangePercent,
                            month1ReturnPercent
                        ],
                        index = hqm_columns
                    ),
                    ignore_index = True
                )

        except:
            print("Houston, we have a problem")


    time_periods = [
        'One-Year',
        'Six-Month',
        'Three-Month',
        'One-Month',
    ]
    for row in hqm_dataframe.index:
        momentum_percentiles = []
        co_name = hqm_dataframe.loc[row, 'Ticker']
        print(co_name)
        for time_period in time_periods:
            col_price = f'{time_period} Price Return'
            col_percentile = f'{time_period} Return Percentile'
            hqm_dataframe.loc[row, col_percentile] = score(hqm_dataframe[col_price], hqm_dataframe.loc[row, col_price])/100
            momentum_percentiles.append(hqm_dataframe.loc[row, col_percentile])
        hqm_dataframe.loc[row, 'HQM Score'] = mean(momentum_percentiles)

    # Now sort and rank the top 50 momentum stocks
    hqm_dataframe.sort_values('HQM Score', ascending=False, inplace=True)
    # hqm_dataframe = hqm_dataframe[:TOP_XX_STOCKS]
    # reset all indices
    # hqm_dataframe.reset_index(drop=True, inplace=True)
    return hqm_dataframe



'''
    Given the HQM ranking, create the Excel spreadsheet
'''
def create_HQM_Excel(hqm_dataframe):
    writer = pd.ExcelWriter('Portfolio_Best_Momentum_Trades.xlsx', engine='xlsxwriter')
    hqm_dataframe.to_excel(writer, 'Recommended Trades', index=False)
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
        'D': ['Number of Shares to Buy', integer_format],
        'E': ['HQM Score', percent_format],
        'F': ['One-Year Price Return', percent_format],
        'G': ['One-Year Return Percentile', percent_format],
        'H': ['Six-Month Price Return', percent_format], 
        'I': ['Six-Month Return Percentile', percent_format],
        'J': ['Three-Month Price Return', percent_format], 
        'K': ['Three-Month Return Percentile', percent_format],
        'L': ['One-Month Price Return', percent_format], 
        'M': ['One-Month Return Percentile', percent_format],
    }

    for column in column_formats.keys():
        writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 18, column_formats[column][1])
        writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], column_formats[column][1])

    writer.save()




def main():
    # final_dataframe = fetch_1y_data()
    # createExcel(final_dataframe)
    # print(final_dataframe)
    hqm_dataframe = fetch_hqm()
    create_HQM_Excel(hqm_dataframe)
    print(hqm_dataframe)



if __name__ == "__main__":
    main()

