"""
scratch/download_fyers_5y.py -- Multi-Year 5m Historical Data Downloader from FYERS API
"""
import os
import sys
import time
import datetime
import pickle
import pandas as pd

ROOT_DIR = r"c:\Users\rishi\Downloads\Trading Pa"
sys.path.insert(0, ROOT_DIR)
from engine.data_engine import create_fyers_client

CACHE_DIR = os.path.join(ROOT_DIR, "data", "fyers_5y_cache")
os.makedirs(CACHE_DIR, exist_ok=True)

# 40 core liquid battle-tested large-cap market leaders across all sectors
CORE_SYMBOLS = [
    # Banking
    ("HDFCBANK", "NSE:HDFCBANK-EQ", "HDFCBANK.NS", "NSEBANK"),
    ("ICICIBANK", "NSE:ICICIBANK-EQ", "ICICIBANK.NS", "NSEBANK"),
    ("SBIN", "NSE:SBIN-EQ", "SBIN.NS", "NSEBANK"),
    ("AXISBANK", "NSE:AXISBANK-EQ", "AXISBANK.NS", "NSEBANK"),
    ("BANKBARODA", "NSE:BANKBARODA-EQ", "BANKBARODA.NS", "NSEBANK"),
    # IT
    ("INFY", "NSE:INFY-EQ", "INFY.NS", "CNXIT"),
    ("TCS", "NSE:TCS-EQ", "TCS.NS", "CNXIT"),
    ("HCLTECH", "NSE:HCLTECH-EQ", "HCLTECH.NS", "CNXIT"),
    ("TECHM", "NSE:TECHM-EQ", "TECHM.NS", "CNXIT"),
    ("COFORGE", "NSE:COFORGE-EQ", "COFORGE.NS", "CNXIT"),
    # Auto
    ("ASHOKLEY", "NSE:ASHOKLEY-EQ", "ASHOKLEY.NS", "CNXAUTO"),
    ("BHARATFORG", "NSE:BHARATFORG-EQ", "BHARATFORG.NS", "CNXAUTO"),
    ("APOLLOTYRE", "NSE:APOLLOTYRE-EQ", "APOLLOTYRE.NS", "CNXAUTO"),
    # Energy
    ("RELIANCE", "NSE:RELIANCE-EQ", "RELIANCE.NS", "CNXENERGY"),
    ("COALINDIA", "NSE:COALINDIA-EQ", "COALINDIA.NS", "CNXENERGY"),
    ("BPCL", "NSE:BPCL-EQ", "BPCL.NS", "CNXENERGY"),
    ("GAIL", "NSE:GAIL-EQ", "GAIL.NS", "CNXENERGY"),
    ("HINDPETRO", "NSE:HINDPETRO-EQ", "HINDPETRO.NS", "CNXENERGY"),
    # Metal
    ("TATASTEEL", "NSE:TATASTEEL-EQ", "TATASTEEL.NS", "CNXMETAL"),
    ("JSWSTEEL", "NSE:JSWSTEEL-EQ", "JSWSTEEL.NS", "CNXMETAL"),
    ("HINDALCO", "NSE:HINDALCO-EQ", "HINDALCO.NS", "CNXMETAL"),
    ("JINDALSTEL", "NSE:JINDALSTEL-EQ", "JINDALSTEL.NS", "CNXMETAL"),
    # Pharma
    ("SUNPHARMA", "NSE:SUNPHARMA-EQ", "SUNPHARMA.NS", "CNXPHARMA"),
    ("CIPLA", "NSE:CIPLA-EQ", "CIPLA.NS", "CNXPHARMA"),
    ("DRREDDY", "NSE:DRREDDY-EQ", "DRREDDY.NS", "CNXPHARMA"),
    ("AUROPHARMA", "NSE:AUROPHARMA-EQ", "AUROPHARMA.NS", "CNXPHARMA"),
    # Financial Services / Consumer
    ("BAJFINANCE", "NSE:BAJFINANCE-EQ", "BAJFINANCE.NS", "CNXCONSUM"),
    ("BAJAJFINSV", "NSE:BAJAJFINSV-EQ", "BAJAJFINSV.NS", "CNXCONSUM"),
    ("CHOLAFIN", "NSE:CHOLAFIN-EQ", "CHOLAFIN.NS", "CNXFIN"),
    ("ABCAPITAL", "NSE:ABCAPITAL-EQ", "ABCAPITAL.NS", "CNXFIN"),
    ("HDFCLIFE", "NSE:HDFCLIFE-EQ", "HDFCLIFE.NS", "CNXCONSUM"),
    # FMCG & Infra
    ("ITC", "NSE:ITC-EQ", "ITC.NS", "CNXFMCG"),
    ("BHARTIARTL", "NSE:BHARTIARTL-EQ", "BHARTIARTL.NS", "CNXINFRA"),
    ("ADANIPORTS", "NSE:ADANIPORTS-EQ", "ADANIPORTS.NS", "CNXINFRA"),
    ("AMBUJACEM", "NSE:AMBUJACEM-EQ", "AMBUJACEM.NS", "CNXINFRA"),
    ("DLF", "NSE:DLF-EQ", "DLF.NS", "CNXREALTY"),
    ("ASTRAL", "NSE:ASTRAL-EQ", "ASTRAL.NS", "CNXINDUSTRIAL"),
    ("BEL", "NSE:BEL-EQ", "BEL.NS", "CNXINDUSTRIAL"),
    ("BHEL", "NSE:BHEL-EQ", "BHEL.NS", "CNXINDUSTRIAL")
]

START_DATE = datetime.date(2022, 1, 1)
END_DATE = datetime.date(2026, 11, 8)

def download_symbol(fyers, fy_symbol, out_path):
    if os.path.exists(out_path):
        with open(out_path, "rb") as f:
            df = pickle.load(f)
        if len(df) > 50000:
            return df

    curr = START_DATE
    all_candles = []
    while curr < END_DATE:
        next_chunk = min(curr + datetime.timedelta(days=95), END_DATE)
        try:
            resp = fyers.history({
                'symbol': fy_symbol,
                'resolution': '5',
                'date_format': '1',
                'range_from': curr.strftime('%Y-%m-%d'),
                'range_to': next_chunk.strftime('%Y-%m-%d'),
                'cont_flag': '1'
            })
            candles = resp.get('candles', [])
            if candles:
                all_candles.extend(candles)
        except Exception as e:
            pass
        curr = next_chunk + datetime.timedelta(days=1)
        time.sleep(0.06)

    if not all_candles:
        return pd.DataFrame()

    rows = []
    for c in all_candles:
        rows.append({
            'ts': datetime.datetime.fromtimestamp(c[0]),
            'open': float(c[1]), 'high': float(c[2]), 'low': float(c[3]),
            'close': float(c[4]), 'volume': float(c[5])
        })
    df = pd.DataFrame(rows).set_index('ts')
    df = df[~df.index.duplicated(keep="last")].sort_index()
    with open(out_path, "wb") as f:
        pickle.dump(df, f)
    return df

def main():
    fyers = create_fyers_client()
    print(f"Starting 5-Year Data Download from FYERS ({START_DATE} to {END_DATE})...")
    
    # Download NIFTY 50 Index
    nifty_path = os.path.join(CACHE_DIR, "NIFTY50.pkl")
    print("Downloading NIFTY 50 Index...")
    df_nifty = download_symbol(fyers, "NSE:NIFTY50-INDEX", nifty_path)
    print(f"  NIFTY 50: {len(df_nifty)} candles from {df_nifty.index[0]} to {df_nifty.index[-1]}")

    t0 = time.time()
    for idx, (sym, fy_sym, ysym, sec) in enumerate(CORE_SYMBOLS):
        out_file = os.path.join(CACHE_DIR, f"{sym}.pkl")
        sym_t0 = time.time()
        df = download_symbol(fyers, fy_sym, out_file)
        sym_t1 = time.time()
        first_d = df.index[0].strftime('%Y-%m-%d') if not df.empty else "N/A"
        last_d = df.index[-1].strftime('%Y-%m-%d') if not df.empty else "N/A"
        print(f"[{idx+1}/{len(CORE_SYMBOLS)}] {sym:12s}: {len(df):6d} candles ({first_d} to {last_d}) in {sym_t1-sym_t0:.1f}s")

    print(f"\nAll symbols downloaded in {time.time()-t0:.1f}s.")

if __name__ == '__main__':
    main()
