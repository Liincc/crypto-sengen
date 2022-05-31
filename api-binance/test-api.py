    ## LIBRARIES & DEPENDANCIES ##
from fileinput import filename
from functools import total_ordering
from multiprocessing.sharedctypes import Value
from operator import contains
import ccxt
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os.path
from datetime import date, datetime, timedelta
from itertools import filterfalse
import time
import xlsxwriter
import openpyxl as xl
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
from pathlib import Path
import csv
import xlwings
from xlwings import Book


# import openpyxl as xl;
# import os
# import xlsxwriter


    #  BINANCE API KEY ##  ## USDT ##

# exchange = ccxt.binance({'options': {'defaultType': 'spot', 'adjustForTimeDifference': True}})
# tickers = exchange.fetch_tickers()
# asset_dict_binance = dict()
# removals_binance = ['WBTC', 'BUSD', 'USDC', 'EUR', 'GBP', 'JPY', 'AUD', 'TRYB', 'XAUT', 'DOWN', 'UP', 'BULL', 'BEAR', 'LONG', 'SHORT', 'MOVE', ':USD', 'HALF', 'HEDGE', 'TUSD', 'USDS', 'TRY', 'RUB', 'ZAR', 'IDRT', 'UAH', 'BIDR', 'BKRW', 'DAI', 'BKRW', 'TRB', 'NGN', 'BRL', 'BVND', 'GYEN']

# assets_binance = pd.DataFrame(list(exchange.fetch_tickers()))
# assets_binance = assets_binance[assets_binance[0].str.contains('USDT', 'USD')==True]
# assets_binance = assets_binance[0].values.tolist()
# updated_binance_assets = [x for x in assets_binance if not any(y in x for y in removals_binance)]
# updated_binance_assets_tmp = [x for x in assets_binance if not any(y in x for y in removals_binance)]
# updated_binance_assets = [coin.split('/USD')[0] + "/USD" for coin in updated_binance_assets_tmp]
# total_binance = list(set(updated_binance_assets))
# total_USDT_binance = [coin.split('/USD')[0] + "/USDT" for coin in total_binance]
# today = pd.to_datetime("today").strftime("%Y/%m/%d")    

# for crypto in total_USDT_binance:
#     try:
#         df_binance = pd.DataFrame(exchange.fetch_ohlcv(crypto, '1d', limit=10), columns=['date', 'open', 'high', 'low', 'close', 'volume'])
#         df_binance['date'] = pd.to_datetime(df_binance['date'], unit='ms')
#         df_binance = df_binance.iloc[::-1]
#         df_binance = df_binance[df_binance['date'] != today]    
#         df_binance.to_csv(f"{os.path.dirname(__file__)}/binance-data/{crypto.replace('/','').replace('USDT','USD')}.csv", index=False)
#         # with pd.ExcelWriter(f"{os.path.dirname(__file__)}/binance-data/{crypto.replace('/','').replace('USDT', 'USD')}.xlsx", datetime_format='DD/MM/YYYY') as writer: # engine={'xlsxwriter': True}
#         #     df_binance.to_excel(writer, index=False)
#             # writer.save()
#             # filename_xlsx = [[crypto] + ".xlsx"]
#             # filename_macro = [[crypto] + ".xlsm"]
#             # book = Book(filename_xlsx)
#             # book.save(filename_macro)
#             # book.close()
#             # os.remove(filename_xlsx)
#             # os.sytem("TASKKILL /F /IM Excel.exe")
#         asset_dict_binance[crypto] = df_binance
#     except:
#         pass

# print(asset_dict_binance)

    ##  BINANCE API KEY ##  ## USDT ##   ## ASSET TESTING SAMPLE ##

exchange = ccxt.binance({'options': {'defaultType': 'spot', 'adjustForTimeDifference': True}})
tickers = exchange.fetch_tickers()
asset_dict_binance = dict()
removals_binance = ['WBTC', 'BUSD', 'USDC', 'EUR', 'GBP', 'JPY', 'AUD', 'TRYB', 'XAUT', 'DOWN', 'UP', 'BULL', 'BEAR', 'LONG', 'SHORT', 'MOVE', ':USD', 'HALF', 'HEDGE', 'TUSD', 'USDS', 'TRY', 'RUB', 'ZAR', 'IDRT', 'UAH', 'BIDR', 'BKRW', 'DAI', 'BKRW', 'TRB', 'NGN', 'BRL', 'BVND', 'GYEN']

assets_binance = pd.DataFrame(list(exchange.fetch_tickers(['BTC/USDT', 'ETH/USDT', 'ADA/USDT'])))
assets_binance = assets_binance[assets_binance[0].str.contains('USDT', 'USD')==True]
assets_binance = assets_binance[0].values.tolist()
updated_binance_assets = [x for x in assets_binance if not any(y in x for y in removals_binance)]
updated_binance_assets_tmp = [x for x in assets_binance if not any(y in x for y in removals_binance)]
updated_binance_assets = [coin.split('/USD')[0] + "/USD" for coin in updated_binance_assets_tmp]
total_binance = list(set(updated_binance_assets))
total_USDT_binance = [coin.split('/USD')[0] + "/USDT" for coin in total_binance]
today = pd.to_datetime("today").strftime("%Y/%m/%d")    

for crypto in total_USDT_binance:
    try:
        df_binance = pd.DataFrame(exchange.fetch_ohlcv(crypto, '1d', limit=10), columns=['date', 'open', 'high', 'low', 'close', 'volume'])
        df_binance['date'] = pd.to_datetime(df_binance['date'], unit='ms')
        df_binance = df_binance.iloc[::-1]
        df_binance = df_binance[df_binance['date'] != today]    
        # df_binance.to_csv(f"{os.path.dirname(__file__)}/binance-data/{crypto.replace('/','').replace('USDT','USD')}.csv", index=False)
        with pd.ExcelWriter(f"{os.path.dirname(__file__)}/binance-data/{crypto.replace('/','').replace('USDT', 'USD')}.xlsx", datetime_format='DD/MM/YYYY') as writer: # engine={'xlsxwriter': True}
            df_binance.to_excel(writer, index=False)
        asset_dict_binance[crypto] = df_binance
    except:
        pass

    #WOKRING ONE#

import xlwings as xw
import shutil
import tempfile


# excel_tickers = [coin.split('/USD')[0] + "USD" for coin in updated_binance_assets_tmp]
# for crypto in excel_tickers:
    
#     wb = xw.Book(f"{os.path.dirname(__file__)}\\binance-data\\{crypto}.xlsx")
#     original = (f'{os.path.dirname(__file__)}\\BCCHARTS\\000GOLD.xlsm')
#     target = (f'{os.path.dirname(__file__)}\\BCCHARTS\\000GOLDcopy.xlsm')
#     shutil.copyfile(original, target)

#     new_wb = xw.Book(f"{os.path.dirname(__file__)}\\BCCHARTS\\000GOLDcopy.xlsm")
#     my_values = wb.sheets['Sheet1'].range('A2:F1200').value
#     new_wb.sheets['table (1)'].range('A5:F5').value = my_values
#     wb.close()
#     new_wb.save(f"{os.path.dirname(__file__)}\\BCCHARTS\\000GOLDcopy.xlsm")
#     macro_create = new_wb.macro("Create")
#     new_wb.close()
#     os.rename(f"{os.path.dirname(__file__)}\\BCCHARTS\\000GOLDcopy.xlsm", f"{os.path.dirname(__file__)}\\BCCHARTS\\{crypto}.xlsm") # Rename 000Goldcopy template to (crypto)/USD
    


    

    ## Single BTCUSD Testing ##

from xlwings import _xlwindows
# xlApp = xw.App(visible=False)
#from pywinauto import application as autoWin

wb = xw.Book(f"{os.path.dirname(__file__)}\\binance-data\\BTCUSD.xlsx")
original = (f'{os.path.dirname(__file__)}\\BCCHARTS\\000GOLD.xlsm')
target = (f'{os.path.dirname(__file__)}\\BCCHARTS\\000GOLDcopy.xlsm')
shutil.copyfile(original, target)

new_wb = xw.Book(f"{os.path.dirname(__file__)}\\BCCHARTS\\000GOLDcopy.xlsm")
new_wb.app.display_alerts = False
my_values = wb.sheets['Sheet1'].range('A2:F1200').value
new_wb.sheets['table (1)'].range('A5:F5').value = my_values
wb.close()
new_wb.save(f"{os.path.dirname(__file__)}\\BCCHARTS\\000GOLDcopy.xlsm")
macro_create = new_wb.macro("Create")
macro_create()
# macro = xlApp.api.Application.Run('Create')
# macro()
#new_wb.save(f"{os.path.dirname(__file__)}\\BCCHARTS\\000GOLDcopy.xlsm")

# if len(new_wb.app.books) == 1:
#     new_wb.app.quit()
# else:
#     new_wb.close()

new_wb.save()
new_wb.close()
os.rename(f"{os.path.dirname(__file__)}\\BCCHARTS\\000GOLDcopy.xlsm", f"{os.path.dirname(__file__)}\\BCCHARTS\\BTCUSD.xlsm") # Rename 000Goldcopy template to (crypto)/USD



