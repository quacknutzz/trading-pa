import sys, pickle, time, os
from datetime import datetime, timedelta, time as dtime
import pandas as pd
import numpy as np

sys.path.insert(0, '.')
sys.path.insert(0, 'sim')
import datafeed
import config as C
from simulator import Simulator, SymbolData, ema_fast, BAR

print("Testing Sector Alignment + Wyckoff Absorption Filter...")

with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)
raw = payload['raw']
idx = payload['idx']
universe = payload['universe']
ndf = idx['NIFTY 50']

sector_map = {item['symbol']: item.get('sector', 'UNKNOWN') for item in universe}

# Group symbols by sector
sector_symbols = {}
for sym, sec in sector_map.items():
    sector_symbols.setdefault(sec, []).append(sym)

# Function to get rolling sector return at time t
def get_sector_15m_return(sec, t, raw_dict):
    syms = sector_symbols.get(sec, [])
    rets = []
    for s in syms:
        y_sym = f"{s}.NS"
        df = raw_dict.get(y_sym)
        if df is not None and t in df.index:
            pos = df.index.get_loc(t)
            if pos >= 3:
                c = df['close'].iloc[pos]
                prev = df['close'].iloc[pos-3]
                if prev > 0:
                    rets.append((c - prev) / prev * 100.0)
    return np.mean(rets) if rets else 0.0

print("Sector alignment function defined.")
