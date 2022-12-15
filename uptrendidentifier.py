from datetime import date, datetime, timedelta, timezone
import pandas as pd
import ccxt
import numpy as np
import os

exchange = ccxt.binance({'options': {'defaultType': 'spot', 'adjustForTimeDifference': False}})
tickers = exchange.fetch_tickers()
asset_dict_binance = dict()
removals_binance = ['AVT', 'RBTC', 'RAI', 'MUSD', 'USDC', 'WBTC', 'GUSD', 'DAI', 'USDP', 'UST', 'SUSD', 'DUSK', 'WBTC', 'BUSD', 'USDC', 'EUR', 'GBP', 'JPY', 'AUD', 'TRYB', 'XAUT', 'DOWN', 'UP', 'BULL', 'BEAR', 'LONG', 'SHORT', 'MOVE', ':USD', 'HALF', 'HEDGE', 'TUSD', 'USDS', 'TRY', 'RUB', 'ZAR', 'IDRT', 'UAH', 'BIDR', 'BKRW', 'DAI', 'BKRW', 'TRB', 'NGN', 'BRL', 'BVND', 'GYEN']
assets_binance = pd.DataFrame(list(exchange.fetch_tickers()))
assets_binance = assets_binance[assets_binance[0].str.contains('USDT', 'USD')==True]
assets_binance = assets_binance[0].values.tolist()
updated_binance_assets = [x for x in assets_binance if not any(y in x for y in removals_binance)]
updated_binance_assets_tmp = [x for x in assets_binance if not any(y in x for y in removals_binance)]
updated_binance_assets = [coin.split('/USD')[0] + "/USD" for coin in updated_binance_assets_tmp]
total_binance = list(set(updated_binance_assets))
total_USDT_binance = [coin.split('/USD')[0] + "/USDT" for coin in total_binance]

exchange_ftx = ccxt.ftx({'options': {'defaultType': 'spot', 'adjustForTimeDifference': False}})
ticker_ftx = exchange_ftx.fetch_tickers()
asset_dict_ftx = dict()
removals_ftx = ['AVT', 'RBTC', 'RAI', 'MUSD', 'USDC', 'WBTC', 'GUSD', 'DAI', 'USDP', 'UST', 'SUSD', 'DUSK', 'IBVOL', 'BVOL', 'INDI', 'SPY', 'WFLOW', 'WBTC', 'USDC', 'BUSD', 'USDT', 'EUR', 'GBP', 'JPY', 'AUD', 'CAD', 'TRYB', 'XAUT', 'DOWN', 'UP', 'BULL', 'BEAR', 'LONG', 'SHORT', 'MOVE', ':USD', 'HALF', 'HEDGE', 'AMC Entertainment', 'StarLaunch', 'WBTC', 'ZM/USD', 'TWTR', 'MRNA', 'HOOD', 'AAPL', 'ACB', 'ABNB', 'AMD', 'AMZN', 'APHA', 'ARKK', 'BABA', 'BB/USD', 'BILI', 'BITO', 'BITW', 'BNTX', 'BYND', 'CGC', 'CITY', 'COIN', 'DKNG', 'ETHE', 'FB/USD', 'GALFAN', 'GBTC', 'GDX', 'GDXJ', 'GLD', 'GLXY', 'GME', 'GOOGL', 'INTER', 'KSOS', 'MSOL', 'MSTR', 'MTA', 'NFLX', 'NIO', 'NOK', 'NVDA', 'PENN', 'PFE', 'PSG', 'PYPL', 'KSHIB', 'SLV', 'SQ', 'STETH', 'STSOL', 'TLRY', 'TSLA', 'TSM', 'UBER', 'USO', 'UST']
assets_ftx = pd.DataFrame(list(exchange_ftx.fetch_tickers()))
assets_ftx = assets_ftx[assets_ftx[0].str.contains('USD')==True]
assets_ftx = assets_ftx[0].values.tolist()
updated_ftx_assets = [x for x in assets_ftx if not any(y in x for y in removals_ftx)]
total_ftx = list (set(updated_ftx_assets) - set(updated_binance_assets))

exchange_coinbase = ccxt.coinbasepro({'options': {'defaultType': 'spot', 'adjustForTimeDifference': False}})
ticker_coinbase = exchange_coinbase.fetch_tickers()
asset_dict_coinbase = dict()
removals_coinbase = ['AVT', 'RBTC', 'RAI', 'MUSD', 'USDC', 'WBTC', 'GUSD', 'DAI', 'USDP', 'UST', 'SUSD', 'DUSK', 'WBTC', 'BUSD', 'WLUNA', 'USDC', 'USDT', 'EUR', 'GBP', 'INDEX' 'JPY', 'AUD', 'TRYB', 'XAUT', 'DOWN', 'UP', 'BULL', 'BEAR', 'LONG', 'SHORT', 'MOVE', ':USD', 'HALF', 'HEDGE', 'TUSD', 'USDS', 'TRY', 'RUB', 'ZAR', 'IDRT', 'UAH', 'BIDR', 'BKRW', 'DAI', 'BKRW', 'TRB', 'NGN', 'BRL', 'BVND', 'GYEN']
assets_coinbase = pd.DataFrame(list(exchange_coinbase.fetch_tickers()))
assets_coinbase = assets_coinbase[assets_coinbase[0].str.contains('USD')==True]
assets_coinbase = assets_coinbase[0].values.tolist()
updated_coinbase_assets = [x for x in assets_coinbase if not any(y in x for y in removals_coinbase)]
total_coinbase = list(set(updated_coinbase_assets) - set(updated_ftx_assets) - set(updated_binance_assets))

exchange_okex = ccxt.okex({'options': {'defaultType': 'spot', 'adjustForTimeDifference': False}})
ticker_okex = exchange_okex.fetch_tickers()
asset_dict_okex = dict()
removals_okex = ['AVT', 'RBTC', 'RAI', 'MUSD', 'USDC', 'WBTC', 'GUSD', 'DAI', 'USDP', 'UST', 'SUSD', 'DUSK', 'JOY', 'EXO','BMN','ALBT','MDOGE','WAMPL','INDEX', 'USDT/' 'WinToken', 'Unitrade', 'DefiBox', 'EUR', 'CUSD', 'GBP', 'JPY', 'AUD', 'TRYB', 'XAUT', 'DOWN', 'UP', 'BULL', 'BEAR', 'LONG', 'SHORT', 'MOVE', ':USD', 'HALF', 'HEDGE', 'TUSD', 'USDS', 'TRY', 'RUB', 'ZAR', 'IDRT', 'UAH', 'BIDR', 'BKRW', 'DAI', 'BKRW', 'TRB', 'NGN', 'BRL', 'BVND', 'GYEN']
assets_okex = pd.DataFrame(list(exchange_okex.fetch_tickers()))
assets_okex = assets_okex[assets_okex[0].str.contains('USD')==True]
assets_okex = assets_okex[0].values.tolist()
updated_okex_assets_tmp = [x for x in assets_okex if not any(y in x for y in removals_okex)]
updated_okex_assets = [coin.split('/USD')[0] + "/USD" for coin in updated_okex_assets_tmp]
total_okex = list(set(updated_okex_assets) - set(updated_binance_assets) - set(updated_ftx_assets) - set(updated_coinbase_assets))
total_okex_USDT = [coin.split('/USD')[0] + "/USDT" for coin in total_okex]

exchange_cryptocom = ccxt.cryptocom({'options': {'defaultType': 'spot', 'adjustForTimeDifference': False}})
ticker_cryptocom = exchange_cryptocom.fetch_tickers()
asset_dict_cryptocom = dict()
removals_cryptocom = ['AVT', 'RBTC', 'RAI', 'MUSD', 'USDC', 'WBTC', 'GUSD', 'DAI', 'USDP', 'UST', 'SUSD', 'DUSK', 'USDC', 'JOY', 'EXO','BMN','ALBT','MDOGE','WAMPL','INDEX', 'CUSD','BUSD', '2S', '3S', '4S', '5S', '2L', '3L', '4L', '5L', 'WinToken', 'Unitrade', 'DefiBox', 'EUR', 'GBP', 'JPY', 'AUD', 'TRYB', 'XAUT', 'DOWN', 'UP', 'BULL', 'BEAR', 'LONG', 'SHORT', 'MOVE', ':USD', 'HALF', 'HEDGE', 'TUSD', 'USDS', 'TRY', 'RUB', 'ZAR', 'IDRT', 'UAH', 'BIDR', 'BKRW', 'DAI', 'BKRW', 'TRB', 'NGN', 'BRL', 'BVND', 'GYEN']
assets_cryptocom = pd.DataFrame(list(exchange_cryptocom.fetch_tickers()))
assets_cryptocom = assets_cryptocom[assets_cryptocom[0].str.contains('USD')==True]
assets_cryptocom = assets_cryptocom[0].values.tolist()
updated_cryptocom_assets_tmp = [x for x in assets_cryptocom if not any(y in x for y in removals_cryptocom)]
updated_cryptocom_assets = [coin.split('/USD')[0] + "/USD" for coin in updated_cryptocom_assets_tmp]
total_cryptocom = list(set(updated_cryptocom_assets) - set(updated_binance_assets) - set(updated_ftx_assets) - set(updated_coinbase_assets) - set(updated_okex_assets))
total_cryptocom_USDT = [coin.split('/USD')[0] + "/USDT" for coin in total_cryptocom]

exchange_bitfinex = ccxt.bitfinex({'options': {'defaultType': 'spot', 'adjustForTimeDifference': False}})
ticker_bitfinex = exchange_bitfinex.fetch_tickers()
asset_dict_bitfinex = dict()
removals_bitfinex = ['AVT', 'RBTC', 'RAI', 'MUSD', 'USDC', 'WBTC', 'GUSD', 'DAI', 'USDP', 'UST', 'SUSD', 'DUSK', 'JOY', 'EXO','BMN','ALBT','MDOGE','WAMPL','INDEX', 'USDC', 'CUSD', 'UST', 'USDT', 'B21X', '2.S', '2X', 'BUSD', 'MIM', 'XCHF', '2S', '3S', '4S', '5S', '2L', '3L', '4L', '5L', 'WinToken', 'Unitrade', 'DefiBox', 'EUR', 'GBP', 'JPY', 'AUD', 'TRYB', 'XAUT', 'DOWN', 'UP', 'BULL', 'BEAR', 'LONG', 'SHORT', 'MOVE', ':USD', 'HALF', 'HEDGE', 'TUSD', 'USDS', 'TRY', 'RUB', 'ZAR', 'IDRT', 'UAH', 'BIDR', 'BKRW', 'DAI', 'BKRW', 'TRB', 'NGN', 'BRL', 'BVND', 'GYEN']
assets_bitfinex = pd.DataFrame(list(exchange_bitfinex.fetch_tickers()))
assets_bitfinex = assets_bitfinex[assets_bitfinex[0].str.contains('USD')==True]
assets_bitfinex = assets_bitfinex[0].values.tolist()
updated_bitfinex_assets = [x for x in assets_bitfinex if not any(y in x for y in removals_bitfinex)]
total_bitfinex = list(set(updated_bitfinex_assets) - set(updated_binance_assets) - set(updated_ftx_assets) - set(updated_coinbase_assets) - set(updated_okex_assets) - set(updated_cryptocom_assets))

excel_tickers_binance = [coin.split('/USD')[0] + "USD" for coin in total_binance]
excel_tickers_ftx = [coin.split('/USD')[0] + "USD" for coin in total_ftx]
excel_tickers_coinbase = [coin.split('/USD')[0] + "USD" for coin in total_coinbase]
excel_tickers_okex = [coin.split('/USD')[0] + "USD" for coin in total_okex]
excel_tickers_cryptocom = [coin.split('/USD')[0] + "USD" for coin in total_cryptocom]
excel_tickers_bitfinex = [coin.split('/USD')[0] + "USD" for coin in total_bitfinex]
#excel_tickers_huobi = [coin.split('/USD')[0] + "USD" for coin in total_huobi]
#excel_tickers_gateio = [coin.split('/USD')[0] + "USD" for coin in total_gateio]
all_tickers = list((excel_tickers_binance) + (excel_tickers_ftx) + (excel_tickers_okex) + (excel_tickers_coinbase) + (excel_tickers_cryptocom) + (excel_tickers_bitfinex))
# today_utc = format((datetime.now(timezone.utc)).strftime("%Y/%m/%d"))

start_date = '2019-01-07'
today_utc = datetime.now().strftime("%Y-%m-%d")
df_master = pd.DataFrame(
    index=pd.date_range(start=start_date, end=format(today_utc), freq='d'),
    columns=[
        'count',
        'uptrend_2',
        'uptrend_5',
        'uptrend_10',
    ],
)

# all_tickers = ['BTCUSD', 'ETHUSD']
directory = f"{os.path.dirname(__file__)}\\MCCHARTS"
onlyfiles = [f for f in os.listdir(directory) if os.path.join(directory, f)]
print(onlyfiles)
missing_files = []
duplicate_files = []
df_master = df_master.fillna(0)
for crypto in all_tickers:
    if any(crypto in x for x in onlyfiles):
        try:
            print(f"Calculating {crypto}")
            filepath = f"C:\\Users\\Lincoln\\LincsLair\\ftx-api\\api-master\\MCCHARTS\\{crypto}.xlsm"
            df = pd.read_excel(filepath, "table (1)", engine='openpyxl', na_values=['nan'])
            df.drop(df.columns[df.columns.str.contains('Unnamed', na=False)],axis = 1, inplace = True)
            df.drop(df.columns[[1,2,3,4,5,6,7,11,12]],axis = 1, inplace=True)
            df = df.dropna().reset_index(drop=True)
            df = df.set_index('Date')
            df['count'] = 1

            if not df.index.is_unique:
                print(f'{crypto} has duplicate indicies')
                print(df[df.index.duplicated(keep=False)])
                print(list(df.index.duplicated()))
                duplicate_files.append(crypto)
                continue
            df = df.reindex(pd.date_range(start_date, today_utc), fill_value=0)
            df['uptrend_2'] = np.where(
                df[0.02] == 'X', 1.0, 0.0,
            )
            df['uptrend_5'] = np.where(
                df[0.05] == 'X', 1.0, 0.0,
            )
            df['uptrend_10'] = np.where(
                df[0.1] == 'X', 1.0, 0.0,
            )
            df_master = df_master.add(df[[
            'count',
            'uptrend_2',
            'uptrend_5',
            'uptrend_10',
            ]])
            df_master = df_master.iloc[::-1]
            # for x in ['count', 'uptrend_2', 'uptrend_5', 'uptrend_10']:
            #     df_master[x] = df_master[x] + df[x]
            # df_master = df_master.add(df, fill_value=0)
        except FileNotFoundError:
            print(f'Could not find {crypto} file')
            missing_files.append(crypto)

print(duplicate_files)
print(missing_files)

df_master['sengen_2'] = round(df_master['uptrend_2'] / df_master['count'], 2)
df_master['sengen_5'] = round(df_master['uptrend_5'] / df_master['count'], 2)
df_master['sengen_10'] = round(df_master['uptrend_10'] / df_master['count'], 2)
df_master.to_csv('SENGEN.csv')
print(f'SENGEN Complete for {len(all_tickers)} cryptocurrencies!')