"""
scratch/simulate_muhurat.py -- Run exact simulation on today's Muhurat Trading session (2026-11-08)
"""
import os
import sys
import pickle
import json
import datetime
import pandas as pd
import numpy as np

ROOT_DIR = r"c:\Users\rishi\Downloads\Trading Pa"
sys.path.insert(0, ROOT_DIR)
sys.path.insert(0, os.path.join(ROOT_DIR, "sim"))

from engine.data_engine import create_fyers_client
from sim.simulator import Simulator, SymbolData
import sim.config as C

def get_complete_data_with_muhurat():
    cache_path = os.path.join(ROOT_DIR, "scratch", "stitched_muhurat_data.pkl")
    if os.path.exists(cache_path):
        print(f"Loading cached stitched data from {cache_path}...")
        with open(cache_path, "rb") as f:
            return pickle.load(f)

    # 1. Load existing cache (up to Nov 5)
    with open(os.path.join(ROOT_DIR, "sim", "cache", "nse_5m_60d.pkl"), "rb") as f:
        cache_data = pickle.load(f)
    raw_dict = cache_data["raw"]
    uni = cache_data["universe"]

    fyers = create_fyers_client()
    
    # 2. Fetch Nov 6 to Nov 8 for all symbols from FYERS
    print("Fetching Nov 6 - Nov 8 from FYERS to complete continuous feed...")
    combined_raw = {}
    
    # Check index ^NSEI / NSE:NIFTY50-INDEX
    res_nifty = fyers.history({
        'symbol': 'NSE:NIFTY50-INDEX',
        'resolution': '5',
        'date_format': '1',
        'range_from': '2026-11-06',
        'range_to': '2026-11-08',
        'cont_flag': '1'
    })
    nifty_candles = res_nifty.get('candles', [])
    rows_nifty = []
    for c in nifty_candles:
        rows_nifty.append({
            'ts': datetime.datetime.fromtimestamp(c[0]),
            'open': float(c[1]), 'high': float(c[2]), 'low': float(c[3]),
            'close': float(c[4]), 'volume': float(c[5])
        })
    df_nifty_new = pd.DataFrame(rows_nifty).set_index('ts')
    
    nifty_old = cache_data["idx"].get("NIFTY 50", cache_data["idx"].get("^NSEI"))
    combined_nifty = pd.concat([nifty_old, df_nifty_new])
    combined_nifty = combined_nifty[~combined_nifty.index.duplicated(keep="last")].sort_index()
    combined_idx = {"NIFTY 50": combined_nifty, "^NSEI": combined_nifty}

    fetched = 0
    for item in uni:
        ysym = item["yahoo"]
        sym = item["symbol"]
        df_old = raw_dict.get(ysym)
        
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
                    rows.append({
                        'ts': datetime.datetime.fromtimestamp(c[0]),
                        'open': float(c[1]), 'high': float(c[2]), 'low': float(c[3]),
                        'close': float(c[4]), 'volume': float(c[5])
                    })
                df_new = pd.DataFrame(rows).set_index('ts')
                if df_old is not None:
                    df_all = pd.concat([df_old, df_new])
                    df_all = df_all[~df_all.index.duplicated(keep="last")].sort_index()
                else:
                    df_all = df_new
                combined_raw[ysym] = df_all
                fetched += 1
            else:
                if df_old is not None:
                    combined_raw[ysym] = df_old
        except Exception:
            if df_old is not None:
                combined_raw[ysym] = df_old
        
        if fetched % 30 == 0 and fetched > 0:
            print(f"  Stitched {fetched}/{len(uni)} symbols...")

    print(f"Stitched {len(combined_raw)} symbols total.")
    result = (combined_raw, combined_idx, uni)
    with open(cache_path, "wb") as f:
        pickle.dump(result, f)
    return result

def test_session(raw, idx, uni, mode="as_is"):
    print(f"\n=======================================================")
    print(f"RUNNING MUHURAT SESSION BACKTEST: Mode = {mode}")
    print(f"=======================================================")
    
    muhurat_date_str = "2026-11-08"
    muhurat_ts = pd.Timestamp("2026-11-08")
    
    orig_start = C.ENTRY_START
    orig_cutoff = C.ENTRY_CUTOFF
    orig_opt_cutoff = C.OPT_ENTRY_CUTOFF
    orig_liq = C.FORCE_LIQUIDATION
    
    if mode == "muhurat_adapted":
        # Muhurat session starts at 14:20 and ends at 15:25
        C.ENTRY_START = datetime.time(14, 25)
        C.ENTRY_CUTOFF = datetime.time(15, 10)
        C.OPT_ENTRY_CUTOFF = datetime.time(15, 0)
        C.FORCE_LIQUIDATION = datetime.time(15, 20)

    try:
        strategies = ["quacknutzz_strat", "v6_baseline"]
        results = {}
        
        for strat in strategies:
            sim = Simulator(mode="replay", session_date=muhurat_ts, strategy=strat, starting_capital=40_000.0)
            sim.load(raw, uni, idx)
            
            dates = sim.available_dates()
            if muhurat_date_str not in dates:
                print(f"ERROR: {muhurat_date_str} not in available dates: {dates[-5:]}")
                return

            bar_times = sim.session_bar_times(muhurat_ts)
            print(f"\n[{strat}] Simulating {len(bar_times)} bars on {muhurat_date_str}...")
            
            for bt in bar_times:
                sim.step(bt)
                
            trades = sim.closed
            signals = sim.signals
            
            gross_pnl = sum(t.get("gross", t.get("pnl", 0.0)) for t in trades)
            tot_costs = sum(t.get("costs", 0.0) if isinstance(t.get("costs"), (int, float)) else t.get("costs", {}).get("total", 0.0) for t in trades)
            net_pnl = sum(t.get("net", 0.0) for t in trades) if trades else 0.0
            
            results[strat] = {
                "trades_count": len(trades),
                "open_count": len(sim.open),
                "signals_count": len(signals),
                "gross_pnl": gross_pnl,
                "costs": tot_costs,
                "net_pnl": net_pnl,
                "trades": trades,
                "signals": signals
            }
            
            print(f"  Results for {strat}:")
            print(f"    Signals Generated: {len(signals)}")
            print(f"    Trades Filled:     {len(trades)}")
            print(f"    Open Positions:    {len(sim.open)}")
            print(f"    Gross P&L:         Rs. {gross_pnl:+.2f}")
            print(f"    Friction/Costs:    Rs. {tot_costs:.2f}")
            print(f"    Net P&L:           Rs. {net_pnl:+.2f}")
            if trades:
                for t in trades:
                    e_t = t['entry_time'].strftime('%H:%M') if hasattr(t['entry_time'], 'strftime') else str(t['entry_time'])
                    x_t = t['exit_time'].strftime('%H:%M') if hasattr(t['exit_time'], 'strftime') else str(t['exit_time'])
                    print(f"      Trade: {t['symbol']} | Side: {t['side']} | Qty: {t['qty']} | Entry: {t['entry']} @ {e_t} | Exit: {t['exit']} @ {x_t} | Reason: {t['reason']} | Gross: Rs. {t['gross']:+.2f} | Net: Rs. {t['net']:+.2f}")
            if sim.open:
                for s, pos in sim.open.items():
                    e_t = pos['entry_time'].strftime('%H:%M') if hasattr(pos['entry_time'], 'strftime') else str(pos['entry_time'])
                    print(f"      STILL OPEN: {s} | Side: {pos['side']} | Entry: {pos['entry']} @ {e_t}")
                    
    finally:
        C.ENTRY_START = orig_start
        C.ENTRY_CUTOFF = orig_cutoff
        C.OPT_ENTRY_CUTOFF = orig_opt_cutoff
        C.FORCE_LIQUIDATION = orig_liq

if __name__ == '__main__':
    raw, idx, uni = get_complete_data_with_muhurat()
    # Test 1: As-is standard settings
    test_session(raw, idx, uni, mode="as_is")
    # Test 2: Muhurat adapted settings
    test_session(raw, idx, uni, mode="muhurat_adapted")
