    ## LIBRARIES & DEPENDANCIES ##
from functools import total_ordering
from multiprocessing.sharedctypes import Value
from operator import contains
import ccxt
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os.path
from datetime import datetime, timedelta
from itertools import filterfalse
import time

# import openpyxl as xl;
# import os
# import xlsxwriter


    ##  BINANCE API KEY ##  ## USDT ##

exchange = ccxt.binance({'options': {'defaultType': 'spot', 'adjustForTimeDifference': True}})
tickers = exchange.fetch_tickers()
asset_dict_binance = dict()
removals_binance = ['WBTC', 'BUSD', 'USDC', 'EUR', 'GBP', 'JPY', 'AUD', 'TRYB', 'XAUT', 'DOWN', 'UP', 'BULL', 'BEAR', 'LONG', 'SHORT', 'MOVE', ':USD', 'HALF', 'HEDGE', 'TUSD', 'USDS', 'TRY', 'RUB', 'ZAR', 'IDRT', 'UAH', 'BIDR', 'BKRW', 'DAI', 'BKRW', 'TRB', 'NGN', 'BRL', 'BVND', 'GYEN']

assets_binance = pd.DataFrame(list(exchange.fetch_tickers()))
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
        df_binance = pd.DataFrame(exchange.fetch_ohlcv(crypto, '1d', limit=250), columns=['date', 'open', 'high', 'low', 'close', 'volume'])
        df_binance['date'] = pd.to_datetime(df_binance['date'], unit='ms')
        df_binance = df_binance.iloc[::-1]
        df_binance = df_binance.drop(df_binance[df_binance['date'] == today].index)
        df_binance.to_csv(f"{os.path.dirname(__file__)}/binance-py-api/{crypto.replace('/','').replace('USDT', 'USD')}.csv", index=False)
        asset_dict_binance[crypto] = df_binance
    except:
        pass
