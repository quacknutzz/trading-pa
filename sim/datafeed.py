"""
sim/datafeed.py -- Free market data from Yahoo Finance (no keys, no money).

Returns per-symbol 5m bars indexed by IST bar-START time (tz-naive),
restricted to the NSE cash session 09:15-15:25 bar starts.
"""
import json
import os
import time
import warnings
from datetime import datetime, timedelta, time as dtime
from typing import Dict, List, Optional

import pandas as pd

warnings.filterwarnings("ignore")
import yfinance as yf

SIM_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(SIM_DIR)
UNIVERSE_FILE = os.path.join(SIM_DIR, "universe.json")


def load_universe() -> List[dict]:
    """Stock universe from Excel Sheet1 (read once, cached to sim/universe.json)."""
    if os.path.exists(UNIVERSE_FILE):
        with open(UNIVERSE_FILE, encoding="utf-8") as f:
            return json.load(f)
    import sys
    sys.path.insert(0, ROOT_DIR)
    from engine.ta_engine import read_excel_config
    from engine.data_engine import resolve_symbol, get_root_folder
    symbols, sec_map, _, _ = read_excel_config(get_root_folder())
    uni = []
    for s in symbols:
        ysym, _, is_index = resolve_symbol(s)
        if not is_index:
            uni.append({"symbol": s, "yahoo": ysym, "sector": sec_map.get(s, "")})
    with open(UNIVERSE_FILE, "w", encoding="utf-8") as f:
        json.dump(uni, f, indent=1)
    return uni


def _clean(df: pd.DataFrame) -> pd.DataFrame:
    df = df.dropna(how="all")
    if df.empty:
        return df
    idx = df.index
    if idx.tz is not None:
        idx = idx.tz_convert("Asia/Kolkata").tz_localize(None)
    df = df.copy()
    df.index = idx
    df.columns = [str(c).lower() for c in df.columns]
    df = df[["open", "high", "low", "close", "volume"]].dropna(subset=["close"])
    t = df.index.time
    df = df[(t >= dtime(9, 15)) & (t <= dtime(15, 25))]
    return df[~df.index.duplicated(keep="last")].sort_index()


def fetch(yahoo_symbols: List[str], period: Optional[str] = None,
          start: Optional[datetime] = None, end: Optional[datetime] = None,
          retries: int = 3) -> Dict[str, pd.DataFrame]:
    out: Dict[str, pd.DataFrame] = {}
    for attempt in range(retries):
        try:
            raw = yf.download(yahoo_symbols, period=period, start=start, end=end, interval="5m",
                              group_by="ticker", progress=False, auto_adjust=False, threads=True)
            break
        except Exception:
            time.sleep(3 * (attempt + 1))
    else:
        return out
    if raw is None or raw.empty:
        return out
    for ys in yahoo_symbols:
        try:
            sub = raw[ys] if ys in raw.columns.get_level_values(0) else raw
            df = _clean(sub)
            if not df.empty:
                out[ys] = df
        except Exception:
            continue
    return out


def last_session_date(reference: Optional[datetime] = None) -> datetime:
    """Most recent date that has NIFTY 5m bars."""
    d = fetch(["^NSEI"], period="5d").get("^NSEI")
    if d is None or d.empty:
        ref = reference or datetime.now()
        return ref - timedelta(days=1)
    return d.index[-1].to_pydatetime()
