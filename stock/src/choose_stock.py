#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Dec 30 08:36:50 2017

@author: kuilinchen
"""

from alpha_vantage import timeseries
import pandas as pd
from finsymbols import symbols
import time
from joblib import Parallel, delayed
import multiprocessing

'''
class FindStock:
    my_api_key = '6NUDYFUECGSY7L9C'
    ts = alpha_vantage.timeseries.TimeSeries(key = my_api_key, output_format='pandas')

    def __init__(self):
        sp500 = symbols.get_sp500_symbols()
        nyse = symbols.get_nyse_symbols()
        amex = symbols.get_amex_symbols()
        nasdaq = symbols.get_nasdaq_symbols()
        self.exchanger_list = [sp500, nyse, amex, nasdaq]
        for exchanger in self.exchanger_list:
            for listed_company in exchanger:
                symbol = listed_company['symbol']
                data, meta_data = ts.get_daily_adjusted(symbol, outputsize = 'compact')           
'''


def is_break_high(stock_symbol):
    break_high = ''
    modified_ticker = stock_symbol.split('^')[0]
    try:
        stock_price, meta_data = ts.get_daily_adjusted(modified_ticker, outputsize='compact')
        highest_price = stock_price['5. adjusted close'].max()
        last_price = stock_price['5. adjusted close'].iloc[-1]
        if last_price == highest_price:
            print('Break high: ', listed_company['company'])
            break_high = symbol
    except ValueError:
        print('ValueError for :', symbol)
    except :
        print('Error happened.')

    return break_high


if __name__ == "__main__":
    sp500 = symbols.get_sp500_symbols()
    nyse = symbols.get_nyse_symbols()
    amex = symbols.get_amex_symbols()
    nasdaq = symbols.get_nasdaq_symbols()
    exchanger_list = [amex]

    num_cores = multiprocessing.cpu_count()
    use_parallel = False

    selected_stock = set()

    my_api_key = '6NUDYFUECGSY7L9C'
    ts = timeseries.TimeSeries(key=my_api_key, output_format='pandas')
    now = time.time()

    if use_parallel:
        for exchanger in exchanger_list:
            results = Parallel(n_jobs=num_cores, backend="threading")(
                delayed(is_break_high)(listed_company['symbol'].split('^')[0]) for listed_company in exchanger)
    else:
        for exchanger in exchanger_list:
            for listed_company in exchanger:
                symbol = listed_company['symbol'].split('^')[0]
                if is_break_high(symbol) == '':
                    selected_stock.add(symbol + ' - ' + listed_company['company'])

    print("Finished in", time.time() - now, "sec")

