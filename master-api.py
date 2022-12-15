    ## LIBRARIES & DEPENDANCIES ##
from ast import Pass, Sub
from fileinput import filename
from functools import total_ordering
from lib2to3.pytree import convert
from multiprocessing.sharedctypes import Value
from operator import contains
from threading import Thread
from tokenize import String
import ccxt
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os.path
from datetime import date, datetime, timedelta, timezone
import datetime as dt
from xlwings.conversion import Converter
from itertools import filterfalse
import time
import xlsxwriter
import openpyxl as xl
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
from pathlib import Path
import csv
import xlwings as xw
from xlwings import Book
import shutil
import tempfile
import glob

    ##  UNIVERSAL VARIABLES ##

#today = pd.to_datetime("today").strftime("%Y/%m/%d")

    ##  BINANCE API KEY ##  ## USDT ##
exchange = ccxt.binance({'options': {'defaultType': 'spot', 'adjustForTimeDifference': True}})
tickers = exchange.fetch_tickers()
asset_dict_binance = dict()
removals_binance = ['STETH', 'PAX', 'AVT', 'RBTC', 'RAI', 'MUSD', 'USDC', 'WBTC', 'GUSD', 'DAI', 'USDP', 'UST', 'SUSD', 'DUSK', 'WBTC', 'BUSD', 'USDC', 'EUR', 'GBP', 'JPY', 'AUD', 'TRYB', 'XAUT', 'DOWN', 'UP', 'BULL', 'BEAR', 'LONG', 'SHORT', 'MOVE', ':USD', 'HALF', 'HEDGE', 'TUSD', 'USDS', 'TRY', 'RUB', 'ZAR', 'IDRT', 'UAH', 'BIDR', 'BKRW', 'DAI', 'BKRW', 'TRB', 'NGN', 'BRL', 'BVND', 'GYEN']
assets_binance = pd.DataFrame(list(exchange.fetch_tickers()))
assets_binance = assets_binance[assets_binance[0].str.contains('USDT', 'USD')==True]
assets_binance = assets_binance[0].values.tolist()
updated_binance_assets = [x for x in assets_binance if not any(y in x for y in removals_binance)]
updated_binance_assets_tmp = [x for x in assets_binance if not any(y in x for y in removals_binance)]
updated_binance_assets = [coin.split('/USD')[0] + "/USD" for coin in updated_binance_assets_tmp]
total_binance = list(set(updated_binance_assets))
total_USDT_binance = [coin.split('/USD')[0] + "/USDT" for coin in total_binance]
today_utc = format((datetime.now(timezone.utc)).strftime("%Y/%m/%d")) # add %H:%M to ensure timezone = utc00:00

for crypto in total_USDT_binance:
        df_binance = pd.DataFrame(exchange.fetch_ohlcv(crypto, '1d', limit=1000), columns=['date', 'open', 'high', 'low', 'close', 'volume'])
        df_binance['date'] = pd.to_datetime(df_binance['date'], unit='ms')
        df_binance['date'] = df_binance['date'].dt.normalize()
        df_binance = df_binance.iloc[::-1]
        df_binance = df_binance[df_binance['date'] != today_utc]
        with pd.ExcelWriter(f"{os.path.dirname(__file__)}/master-py-api/{crypto.replace('/','').replace('USDT', 'USD')}.xlsx", datetime_format='DD-MM-YYYY') as writer: # engine={'xlsxwriter': True}
            df_binance.to_excel(writer, index=False)
        asset_dict_binance[crypto] = df_binance
    ## FTX API KEY  ##
exchange_ftx = ccxt.ftx({'options': {'defaultType': 'spot', 'adjustForTimeDifference': True}})
ticker_ftx = exchange_ftx.fetch_tickers()
asset_dict_ftx = dict()
removals_ftx = ['STETH', 'PAX', 'AVT', 'RBTC', 'RAI', 'MUSD', 'USDC', 'WBTC', 'GUSD', 'DAI', 'USDP', 'UST', 'SUSD', 'DUSK', 'IBVOL', 'BVOL', 'INDI', 'SPY', 'WFLOW', 'WBTC', 'USDC', 'BUSD', 'USDT', 'EUR', 'GBP', 'JPY', 'AUD', 'CAD', 'TRYB', 'XAUT', 'DOWN', 'UP', 'BULL', 'BEAR', 'LONG', 'SHORT', 'MOVE', ':USD', 'HALF', 'HEDGE', 'AMC Entertainment', 'StarLaunch', 'WBTC', 'ZM/USD', 'TWTR', 'MRNA', 'HOOD', 'AAPL', 'ACB', 'ABNB', 'AMD', 'AMZN', 'APHA', 'ARKK', 'BABA', 'BB/USD', 'BILI', 'BITO', 'BITW', 'BNTX', 'BYND', 'CGC', 'CITY', 'COIN', 'DKNG', 'ETHE', 'FB/USD', 'GALFAN', 'GBTC', 'GDX', 'GDXJ', 'GLD', 'GLXY', 'GME', 'GOOGL', 'INTER', 'KSOS', 'MSOL', 'MSTR', 'MTA', 'NFLX', 'NIO', 'NOK', 'NVDA', 'PENN', 'PFE', 'PSG', 'PYPL', 'KSHIB', 'SLV', 'SQ', 'STETH', 'STSOL', 'TLRY', 'TSLA', 'TSM', 'UBER', 'USO', 'UST']
assets_ftx = pd.DataFrame(list(exchange_ftx.fetch_tickers()))
assets_ftx = assets_ftx[assets_ftx[0].str.contains('USD')==True]
assets_ftx = assets_ftx[0].values.tolist()
updated_ftx_assets = [x for x in assets_ftx if not any(y in x for y in removals_ftx)]
total_ftx = list (set(updated_ftx_assets) - set(updated_binance_assets))

for crypto1 in total_ftx:
        df_ftx = pd.DataFrame(exchange_ftx.fetch_ohlcv(crypto1, '1d', limit=1000), columns=['date', 'open', 'high', 'low', 'close', 'volume'])
        df_ftx['date'] = pd.to_datetime(df_ftx['date'], unit='ms')
        df_ftx['date'] = df_ftx['date'].dt.normalize()
        df_ftx = df_ftx.iloc[::-1]
        df_ftx = df_ftx[df_ftx['date'] != today_utc]
        with pd.ExcelWriter(f"{os.path.dirname(__file__)}/master-py-api/{crypto1.replace('/','').replace('USDT', 'USD')}.xlsx", datetime_format='DD-MM-YYYY') as writer: # engine={'xlsxwriter': True}
            df_ftx.to_excel(writer, index=False)
        asset_dict_ftx[crypto1] = df_ftx
    ## COINBASE API KEY ##
exchange_coinbase = ccxt.coinbasepro({'options': {'defaultType': 'spot', 'adjustForTimeDifference': True}})
ticker_coinbase = exchange_coinbase.fetch_tickers()
asset_dict_coinbase = dict()
removals_coinbase = ['STETH', 'PAX', 'AVT', 'RBTC', 'RAI', 'MUSD', 'USDC', 'WBTC', 'GUSD', 'DAI', 'USDP', 'UST', 'SUSD', 'DUSK', 'WBTC', 'BUSD', 'WLUNA', 'USDC', 'USDT', 'EUR', 'GBP', 'INDEX' 'JPY', 'AUD', 'TRYB', 'XAUT', 'DOWN', 'UP', 'BULL', 'BEAR', 'LONG', 'SHORT', 'MOVE', ':USD', 'HALF', 'HEDGE', 'TUSD', 'USDS', 'TRY', 'RUB', 'ZAR', 'IDRT', 'UAH', 'BIDR', 'BKRW', 'DAI', 'BKRW', 'TRB', 'NGN', 'BRL', 'BVND', 'GYEN']
assets_coinbase = pd.DataFrame(list(exchange_coinbase.fetch_tickers()))
assets_coinbase = assets_coinbase[assets_coinbase[0].str.contains('USD')==True]
assets_coinbase = assets_coinbase[0].values.tolist()
updated_coinbase_assets = [x for x in assets_coinbase if not any(y in x for y in removals_coinbase)]
total_coinbase = list(set(updated_coinbase_assets) - set(updated_ftx_assets) - set(updated_binance_assets))

for crypto2 in total_coinbase:
        df_coinbase = pd.DataFrame(exchange_coinbase.fetch_ohlcv(crypto2, '1d', limit=1000), columns=['date', 'open', 'high', 'low', 'close', 'volume'])
        df_coinbase['date'] = pd.to_datetime(df_coinbase['date'], unit='ms')
        df_coinbase['date'] = df_coinbase['date'].dt.normalize() 
        df_coinbase = df_coinbase.iloc[::-1]
        df_coinbase = df_coinbase[df_coinbase['date'] != today_utc]
        with pd.ExcelWriter(f"{os.path.dirname(__file__)}/master-py-api/{crypto2.replace('/','').replace('USDT', 'USD')}.xlsx", datetime_format='DD-MM-YYYY') as writer: # engine={'xlsxwriter': True}
            df_coinbase.to_excel(writer, index=False)
        asset_dict_coinbase[crypto2] = df_coinbase
    ## OKEX API KEY ##  ## USDT ##
exchange_okex = ccxt.okex({'options': {'defaultType': 'spot', 'adjustForTimeDifference': True}})
ticker_okex = exchange_okex.fetch_tickers()
asset_dict_okex = dict()
removals_okex = ['STETH', 'PAX', 'AVT', 'RBTC', 'RAI', 'MUSD', 'USDC', 'WBTC', 'GUSD', 'DAI', 'USDP', 'UST', 'SUSD', 'DUSK', 'JOY', 'EXO','BMN','ALBT','MDOGE','WAMPL','INDEX', 'USDT/' 'WinToken', 'Unitrade', 'DefiBox', 'EUR', 'CUSD', 'GBP', 'JPY', 'AUD', 'TRYB', 'XAUT', 'DOWN', 'UP', 'BULL', 'BEAR', 'LONG', 'SHORT', 'MOVE', ':USD', 'HALF', 'HEDGE', 'TUSD', 'USDS', 'TRY', 'RUB', 'ZAR', 'IDRT', 'UAH', 'BIDR', 'BKRW', 'DAI', 'BKRW', 'TRB', 'NGN', 'BRL', 'BVND', 'GYEN']
assets_okex = pd.DataFrame(list(exchange_okex.fetch_tickers()))
assets_okex = assets_okex[assets_okex[0].str.contains('USD')==True]
assets_okex = assets_okex[0].values.tolist()
updated_okex_assets_tmp = [x for x in assets_okex if not any(y in x for y in removals_okex)]
updated_okex_assets = [coin.split('/USD')[0] + "/USD" for coin in updated_okex_assets_tmp]
total_okex = list(set(updated_okex_assets) - set(updated_binance_assets) - set(updated_ftx_assets) - set(updated_coinbase_assets))
total_okex_USDT = [coin.split('/USD')[0] + "/USDT" for coin in total_okex]

for crypto3 in total_okex_USDT:
    try: # Needed for USDT/USDT Pairing
        df_okex = pd.DataFrame(exchange_okex.fetch_ohlcv(crypto3, '1d', limit=1000), columns=['date', 'open', 'high', 'low', 'close', 'volume'])
        df_okex['date'] = pd.to_datetime(df_okex['date'], unit='ms')
        df_okex['date'] = df_okex['date'].dt.normalize()
        df_okex = df_okex.iloc[::-1]
        df_okex = df_okex[df_okex['date'] != today_utc]
        with pd.ExcelWriter(f"{os.path.dirname(__file__)}/master-py-api/{crypto3.replace('/','').replace('USDT', 'USD')}.xlsx", datetime_format='DD-MM-YYYY') as writer: # engine={'xlsxwriter': True}
            df_okex.to_excel(writer, index=False)
        asset_dict_okex[crypto3] = df_okex
    except:
        pass
    ## CRYPTOCOM ## ## some pairs are / USDC only ## USDT ##
exchange_cryptocom = ccxt.cryptocom({'options': {'defaultType': 'spot', 'adjustForTimeDifference': True}})
ticker_cryptocom = exchange_cryptocom.fetch_tickers()
asset_dict_cryptocom = dict()
removals_cryptocom = ['STETH', 'PAX', 'AVT', 'RBTC', 'RAI', 'MUSD', 'USDC', 'WBTC', 'GUSD', 'DAI', 'USDP', 'UST', 'SUSD', 'DUSK', 'USDC', 'JOY', 'EXO','BMN','ALBT','MDOGE','WAMPL','INDEX', 'CUSD','BUSD', '2S', '3S', '4S', '5S', '2L', '3L', '4L', '5L', 'WinToken', 'Unitrade', 'DefiBox', 'EUR', 'GBP', 'JPY', 'AUD', 'TRYB', 'XAUT', 'DOWN', 'UP', 'BULL', 'BEAR', 'LONG', 'SHORT', 'MOVE', ':USD', 'HALF', 'HEDGE', 'TUSD', 'USDS', 'TRY', 'RUB', 'ZAR', 'IDRT', 'UAH', 'BIDR', 'BKRW', 'DAI', 'BKRW', 'TRB', 'NGN', 'BRL', 'BVND', 'GYEN']
assets_cryptocom = pd.DataFrame(list(exchange_cryptocom.fetch_tickers()))
assets_cryptocom = assets_cryptocom[assets_cryptocom[0].str.contains('USD')==True]
assets_cryptocom = assets_cryptocom[0].values.tolist()
updated_cryptocom_assets_tmp = [x for x in assets_cryptocom if not any(y in x for y in removals_cryptocom)]
updated_cryptocom_assets = [coin.split('/USD')[0] + "/USD" for coin in updated_cryptocom_assets_tmp]
total_cryptocom = list(set(updated_cryptocom_assets) - set(updated_binance_assets) - set(updated_ftx_assets) - set(updated_coinbase_assets) - set(updated_okex_assets))
total_cryptocom_USDT = [coin.split('/USD')[0] + "/USDT" for coin in total_cryptocom]

for crypto4 in total_cryptocom_USDT:
        df_cryptocom = pd.DataFrame(exchange_cryptocom.fetch_ohlcv(crypto4, '1d', limit=1000), columns=['date', 'open', 'high', 'low', 'close', 'volume'])
        df_cryptocom['date'] = pd.to_datetime(df_cryptocom['date'], unit='ms')
        df_cryptocom['date'] = df_cryptocom['date'].dt.normalize()
        df_cryptocom = df_cryptocom.iloc[::-1]
        df_cryptocom = df_cryptocom[df_cryptocom['date'] != today_utc]
        with pd.ExcelWriter(f"{os.path.dirname(__file__)}/master-py-api/{crypto4.replace('/','').replace('USDT', 'USD')}.xlsx", datetime_format='DD-MM-YYYY') as writer: # engine={'xlsxwriter': True}
            df_cryptocom.to_excel(writer, index=False)
        asset_dict_cryptocom[crypto4] = df_cryptocom
        ## BITFINEX API ##
exchange_bitfinex = ccxt.bitfinex({'options': {'defaultType': 'spot', 'adjustForTimeDifference': True}})
ticker_bitfinex = exchange_bitfinex.fetch_tickers()
asset_dict_bitfinex = dict()
removals_bitfinex = ['STETH', 'PAX', 'AVT', 'RBTC', 'RAI', 'MUSD', 'USDC', 'WBTC', 'GUSD', 'DAI', 'USDP', 'UST', 'SUSD', 'DUSK', 'JOY', 'EXO','BMN','ALBT','MDOGE','WAMPL','INDEX', 'USDC', 'CUSD', 'UST', 'USDT', 'B21X', '2.S', '2X', 'BUSD', 'MIM', 'XCHF', '2S', '3S', '4S', '5S', '2L', '3L', '4L', '5L', 'WinToken', 'Unitrade', 'DefiBox', 'EUR', 'GBP', 'JPY', 'AUD', 'TRYB', 'XAUT', 'DOWN', 'UP', 'BULL', 'BEAR', 'LONG', 'SHORT', 'MOVE', ':USD', 'HALF', 'HEDGE', 'TUSD', 'USDS', 'TRY', 'RUB', 'ZAR', 'IDRT', 'UAH', 'BIDR', 'BKRW', 'DAI', 'BKRW', 'TRB', 'NGN', 'BRL', 'BVND', 'GYEN']
assets_bitfinex = pd.DataFrame(list(exchange_bitfinex.fetch_tickers()))
assets_bitfinex = assets_bitfinex[assets_bitfinex[0].str.contains('USD')==True]
assets_bitfinex = assets_bitfinex[0].values.tolist()
updated_bitfinex_assets = [x for x in assets_bitfinex if not any(y in x for y in removals_bitfinex)]
total_bitfinex = list(set(updated_bitfinex_assets) - set(updated_binance_assets) - set(updated_ftx_assets) - set(updated_coinbase_assets) - set(updated_okex_assets) - set(updated_cryptocom_assets))

for crypto5 in total_bitfinex:
        df_bitfinex = pd.DataFrame(exchange_bitfinex.fetch_ohlcv(crypto5, '1d', limit=1000), columns=['date', 'open', 'high', 'low', 'close', 'volume'])
        df_bitfinex['date'] = pd.to_datetime(df_bitfinex['date'], unit='ms')
        df_bitfinex['date'] = df_bitfinex['date'].dt.normalize()
        df_bitfinex = df_bitfinex.iloc[::-1]
        df_bitfinex = df_bitfinex[df_bitfinex['date'] != today_utc]
        with pd.ExcelWriter(f"{os.path.dirname(__file__)}/master-py-api/{crypto5.replace('/','').replace('USDT', 'USD')}.xlsx", datetime_format='DD-MM-YYYY') as writer: # engine={'xlsxwriter': True}
            df_bitfinex.to_excel(writer, index=False)
        asset_dict_bitfinex[crypto5] = df_bitfinex
        ## HUOBI ##
exchange_huobi = ccxt.huobi({'options': {'defaultType': 'spot', 'adjustForTimeDifference': True}})
ticker_huobi = exchange_huobi.fetch_tickers()
asset_dict_huobi = dict()
removals_huobi = ['STETH', 'PAX', 'AVT', 'RBTC', 'RAI', 'MUSD', 'USDC', 'WBTC', 'GUSD', 'DAI', 'USDP', 'UST', 'DUSK', 'JOY', 'EXO','BMN','ALBT','MDOGE','WAMPL','INDEX','USDC', 'CUSD', 'NHBTC', 'HOLO', 'HitChain', 'HUSD', 'BUSD', '1S', 'CTXC2X', 'Hydro Protocol', '2S', '3S', '4S', '5S', '2L', '3L', '4L', '5L', 'WinToken', 'Unitrade', 'DefiBox', 'EUR', 'GBP', 'JPY', 'AUD', 'TRYB', 'XAUT', 'DOWN', 'UP', 'BULL', 'BEAR', 'LONG', 'SHORT', 'MOVE', ':USD', 'HALF', 'HEDGE', 'TUSD', 'USDS', 'TRY', 'RUB', 'ZAR', 'IDRT', 'UAH', 'BIDR', 'BKRW', 'DAI', 'BKRW', 'TRB', 'NGN', 'BRL', 'BVND', 'GYEN']
assets_huobi = pd.DataFrame(list(exchange_huobi.fetch_tickers()))
assets_huobi = assets_huobi[assets_huobi[0].str.contains('USD')==True]
assets_huobi = assets_huobi[0].values.tolist()
updated_huobi_assets_tmp = [x for x in assets_huobi if not any(y in x for y in removals_huobi)]
updated_huobi_assets = [coin.split('/USD')[0] + "/USD" for coin in updated_huobi_assets_tmp]
total_huobi = list(set(updated_huobi_assets) - set(updated_binance_assets) - set(updated_ftx_assets) - set(updated_coinbase_assets) - set(updated_okex_assets) - set(updated_cryptocom_assets) - set(updated_bitfinex_assets))
total_huobi_USDT = [coin.split('/USD')[0] + "/USDT" for coin in total_huobi]
# startDate = "2019-06-01"
# startDate = datetime.strftime("%Y-%m-%d")

for crypto6 in total_huobi_USDT:
        df_huobi = pd.DataFrame(exchange_huobi.fetch_ohlcv(crypto6, timeframe='1d', limit=1000), columns=['date', 'open', 'high', 'low', 'close', 'volume'])
        df_huobi['date'] = pd.to_datetime(df_huobi['date'], unit='ms')
        df_huobi['date'] = df_huobi['date'].dt.normalize()
        df_huobi = df_huobi.iloc[::-1]
        df_huobi = df_huobi[df_huobi['date'] != today_utc]
        with pd.ExcelWriter(f"{os.path.dirname(__file__)}/master-py-api/{crypto6.replace('/','').replace('USDT', 'USD')}.xlsx", datetime_format='DD-MM-YYYY') as writer: # engine={'xlsxwriter': True}
            df_huobi.to_excel(writer, index=False)
        asset_dict_huobi[crypto6] = df_huobi


    ## MASTER LIST TO CHART ALL P&FS ##
        ## UNIVERSAL VARIABLES ##

excel_tickers_binance = [coin.split('/USD')[0] + "USD" for coin in total_binance]
excel_tickers_ftx = [coin.split('/USD')[0] + "USD" for coin in total_ftx]
excel_tickers_coinbase = [coin.split('/USD')[0] + "USD" for coin in total_coinbase]
excel_tickers_okex = [coin.split('/USD')[0] + "USD" for coin in total_okex]
excel_tickers_cryptocom = [coin.split('/USD')[0] + "USD" for coin in total_cryptocom]
excel_tickers_bitfinex = [coin.split('/USD')[0] + "USD" for coin in total_bitfinex]
excel_tickers_huobi = [coin.split('/USD')[0] + "USD" for coin in total_huobi]
#excel_tickers_gateio = [coin.split('/USD')[0] + "USD" for coin in total_gateio]
all_tickers = list((excel_tickers_binance) + (excel_tickers_ftx) + (excel_tickers_okex) + (excel_tickers_coinbase) + (excel_tickers_cryptocom) + (excel_tickers_bitfinex) + (excel_tickers_huobi))
xlApp = xw.App(visible=False)

    ##Charts A NEW CHART that hasn't been listed or charted before
directory = f"{os.path.dirname(__file__)}\\MCCHARTS"
onlyfiles = [f for f in os.listdir(directory) if os.path.join(directory, f)]
def new_chart(crypto):
    print(f"creating a new file for {crypto}")
    wb = xw.Book(f"{os.path.dirname(__file__)}\\master-py-api\\{crypto}.xlsx")
    new_wb = xw.Book(f"{os.path.dirname(__file__)}\\MCCHARTS\\000GOLD.xlsm")
    new_wb.save(f"{os.path.dirname(__file__)}\\MCCHARTS\\{crypto}.xlsm")
    my_values = wb.sheets['Sheet1'].range('A3:F1200').value ## A2 Should be changed to A3 to plot today_utc-1 data for new crypos added to MCCHARTS 
    new_wb.sheets['table (1)'].range('A5:F5').value = my_values
    wb.close()
    new_wb.save()
    macro_vba = xlApp.api.Application.Run('CreateNew')
    macro_vba()
    new_wb.close()
    if len(new_wb.app.books) >= 1:
        new_wb.app.quit()
    else:
        new_wb.close()
for crypto in all_tickers:
    if any(crypto in x for x in onlyfiles):
        print(f"{crypto} is already in the directory")
    else:
        try:
            new_chart(crypto)
        except Exception as e:
            # print(e)
            pass

    ##CHARTS THE 'DAILY UPDATE' / charts that ALREADY have a .xlsm file created
btc_wb = xw.Book(f"{os.path.dirname(__file__)}\\master-py-api\\BTCUSD.xlsx")
btc_cell = str(btc_wb.sheets['Sheet1'].range('A2').value)
btc_wb.close()
def update_chart(crypto, wb, wb_xlsx):
    curr_date = wb.sheets['table (1)'].range('A5').options(dates=dt.date, convert=None).value #A5 Cell selected since this is the most recent date charted
    curr_date_str = curr_date.strftime("%Y-%m-%d")
    curr_date_str = curr_date_str.replace('-', '/')
    date_range = pd.date_range(start=curr_date_str,end=today_utc,freq='d')
    period_charting = (len(date_range))
    insert_row = period_charting + 4
    print(f"Updating chart for {crypto}")
    my_values = wb_xlsx.sheets['Sheet1'].range(f'A2:F{period_charting}').value
    wb.sheets['table (1)'].range(f"5:{insert_row}").insert('down')
    wb.sheets['table (1)'].range('A5:F5').value = my_values
    wb_xlsx.close()
    macro_update = xlApp.api.Application.Run('CreateDaily')
    macro_update()
    if len(wb.app.books) >= 1:
        wb.app.quit()
    else:
        wb.close()
for crypto in all_tickers:
    try:
        wb = xw.Book(f"{os.path.dirname(__file__)}\\MCCHARTS\\{crypto}.xlsm")
        wb_xlsx = xw.Book(f"{os.path.dirname(__file__)}\\master-py-api\\{crypto}.xlsx")
        recent_xlsx = str(wb_xlsx.sheets['Sheet1'].range('A2').value)
        recent_xlsm = str(wb.sheets['table (1)'].range('A5').value)
        if recent_xlsx != btc_cell:
            print(f"broke the loop because {crypto} has no data for {recent_xlsx}")
            wb.close()
            wb_xlsx.close()
            wb.app.quit()
            time.sleep(1) #Prevents script from overlapping Excel API files and corrupting
            break
        elif recent_xlsm == recent_xlsx:
            print(f"{crypto} already charted")
            wb.close()
            wb_xlsx.close()
            wb.app.quit()
            continue
        else:
            update_chart(crypto, wb, wb_xlsx)
    except:
        pass
print(f"All charting for {len(all_tickers)} cryptos finished!")

    ## FINISHED TESTING ##
total_all = total_binance + total_ftx + total_coinbase + total_okex + total_cryptocom + total_bitfinex + total_huobi
print(f"Total number of (1) Binance cryptos used: {len(total_binance)}")
print(f"Total number of (2) FTX cryptos used: {len(total_ftx)}")
print(f"Total number of (2) OkEx cryptos used: {len(total_okex)}")
print(f"Total number of (3) Coinbase cryptos used: {len(total_coinbase)}")
print(f"Total number of (4) Crypto.com cryptos used: {len(total_cryptocom)}")
print(f"total number of (5) Bitfinex used: {len(total_bitfinex)}")
print(f"total number of (6) Huobi used: {len(total_huobi)}")
#print(f"total number of (7) Gate.io used: {len(total_gateio)}")
print(f"total number of all Cryptos used: {len(total_all)}")
