"""
scratch/analyze_muhurat_session.py -- Comprehensive Muhurat Trading Session Analysis (2026-11-08)
"""
import os
import sys
import json
import time
import datetime
import pandas as pd
import numpy as np

ROOT_DIR = r"c:\Users\rishi\Downloads\Trading Pa"
sys.path.insert(0, ROOT_DIR)
sys.path.insert(0, os.path.join(ROOT_DIR, "sim"))

from engine.data_engine import create_fyers_client
from sim.simulator import Simulator, SymbolData, trade_costs
import sim.config as C

def run_muhurat_analysis():
    fyers = create_fyers_client()
    with open(os.path.join(ROOT_DIR, "sim", "universe.json")) as f:
        uni = json.load(f)

    # 1. Fetch NIFTY index data for Nov 6 and Nov 8
    print("Fetching NIFTY data...")
    res_nifty = fyers.history({
        'symbol': 'NSE:NIFTY50-INDEX',
        'resolution': '5',
        'date_format': '1',
        'range_from': '2026-11-06',
        'range_to': '2026-11-08',
        'cont_flag': '1'
    })
    nifty_candles = res_nifty.get('candles', [])
    print(f"NIFTY candles: {len(nifty_candles)}")

    # 2. Fetch all stocks in universe for Nov 6 to Nov 8 (to provide continuous warmup)
    stock_dfs = {}
    print(f"Fetching data for {len(uni)} universe symbols across Nov 6 - Nov 8...")
    success_count = 0
    for idx, item in enumerate(uni):
        sym = item['symbol']
        fy_sym = f"NSE:{sym}-EQ"
        try:
            resp = fyers.history({
                'symbol': fy_sym,
                'resolution': '5',
                'date_format': '1',
                'range_from': '2026-11-06',
                'range_to': '2026-11-08',
                'cont_flag': '1'
            })
            candles = resp.get('candles', [])
            if candles:
                rows = []
                for c in candles:
                    # timestamp in seconds
                    dt = datetime.datetime.fromtimestamp(c[0])
                    rows.append({
                        'timestamp': dt,
                        'open': float(c[1]),
                        'high': float(c[2]),
                        'low': float(c[3]),
                        'close': float(c[4]),
                        'volume': float(c[5])
                    })
                df = pd.DataFrame(rows).set_index('timestamp').sort_index()
                stock_dfs[sym] = df
                success_count += 1
        except Exception as e:
            pass
        if (idx + 1) % 30 == 0:
            print(f"  Processed {idx + 1}/{len(uni)} symbols...")
        time.sleep(0.05) # adhere to rate limits

    print(f"Successfully fetched {success_count}/{len(uni)} symbols.")

    # 3. Analyze Market Breadth during Muhurat Session
    muhurat_date = datetime.date(2026, 11, 8)
    gainers = []
    losers = []
    unchanged = []
    
    for sym, df in stock_dfs.items():
        m_df = df[df.index.date == muhurat_date]
        if not m_df.empty:
            open_px = m_df.iloc[0]['open']
            close_px = m_df.iloc[-1]['close']
            high_px = m_df['high'].max()
            low_px = m_df['low'].min()
            vol = m_df['volume'].sum()
            chg_pct = ((close_px - open_px) / open_px) * 100.0
            info = {
                'symbol': sym,
                'open': open_px,
                'high': high_px,
                'low': low_px,
                'close': close_px,
                'volume': vol,
                'chg_pct': chg_pct,
                'bars': len(m_df)
            }
            if chg_pct > 0.05:
                gainers.append(info)
            elif chg_pct < -0.05:
                losers.append(info)
            else:
                unchanged.append(info)

    gainers.sort(key=lambda x: x['chg_pct'], reverse=True)
    losers.sort(key=lambda x: x['chg_pct'])

    print("\n--- MUHURAT SESSION MARKET OVERVIEW (2026-11-08) ---")
    print(f"Total active universe stocks: {len(gainers) + len(losers) + len(unchanged)}")
    print(f"Gainers: {len(gainers)} | Losers: {len(losers)} | Flat: {len(unchanged)}")
    
    print("\nTop 5 Gainers:")
    for g in gainers[:5]:
        print(f"  {g['symbol']:15s}: Open {g['open']:8.2f} -> Close {g['close']:8.2f} ({g['chg_pct']:+.2f}%) [Vol: {g['volume']:,.0f}]")
        
    print("\nTop 5 Losers:")
    for l in losers[:5]:
        print(f"  {l['symbol']:15s}: Open {l['open']:8.2f} -> Close {l['close']:8.2f} ({l['chg_pct']:+.2f}%) [Vol: {l['volume']:,.0f}]")

    # 4. Check Strategy Execution on Muhurat Session
    # Let's inspect bar timestamps on Muhurat session
    sample_df = list(stock_dfs.values())[0]
    m_sample = sample_df[sample_df.index.date == muhurat_date]
    print(f"\nMuhurat bar timestamps ({len(m_sample)} bars):")
    for t in m_sample.index:
        print(f"  {t.strftime('%H:%M:%S')}")

    return stock_dfs

if __name__ == '__main__':
    run_muhurat_analysis()
