"""
scratch/run_multi_year_backtest.py -- 5-Year High-Fidelity Intraday Backtest (2022 - 2026)
Starting Capital: Rs. 40,000
Strategies: quacknutzz_strat vs v6_baseline
"""
import os
import sys
import time
import pickle
import datetime
import pandas as pd
import numpy as np

ROOT_DIR = r"c:\Users\rishi\Downloads\Trading Pa"
sys.path.insert(0, ROOT_DIR)
sys.path.insert(0, os.path.join(ROOT_DIR, "sim"))

import sim.config as C
from sim.simulator import Simulator, SymbolData, trade_costs

CACHE_DIR = os.path.join(ROOT_DIR, "data", "fyers_5y_cache")

def load_cached_universe():
    print("Loading cached 5-year data from disk...")
    nifty_path = os.path.join(CACHE_DIR, "NIFTY50.pkl")
    with open(nifty_path, "rb") as f:
        df_nifty = pickle.load(f)
    print(f"  Loaded NIFTY 50: {len(df_nifty):,} bars ({df_nifty.index[0].date()} to {df_nifty.index[-1].date()})")

    # Load all available stock pkl files
    files = [f for f in os.listdir(CACHE_DIR) if f.endswith(".pkl") and f != "NIFTY50.pkl"]
    stock_dfs = {}
    universe = []
    
    # Sector mapping helper
    sec_map = {
        "HDFCBANK": "NSEBANK", "ICICIBANK": "NSEBANK", "SBIN": "NSEBANK", "AXISBANK": "NSEBANK", "BANKBARODA": "NSEBANK",
        "INFY": "CNXIT", "TCS": "CNXIT", "HCLTECH": "CNXIT", "TECHM": "CNXIT", "COFORGE": "CNXIT",
        "ASHOKLEY": "CNXAUTO", "BHARATFORG": "CNXAUTO", "APOLLOTYRE": "CNXAUTO",
        "RELIANCE": "CNXENERGY", "COALINDIA": "CNXENERGY", "BPCL": "CNXENERGY", "GAIL": "CNXENERGY", "HINDPETRO": "CNXENERGY",
        "TATASTEEL": "CNXMETAL", "JSWSTEEL": "CNXMETAL", "HINDALCO": "CNXMETAL", "JINDALSTEL": "CNXMETAL",
        "SUNPHARMA": "CNXPHARMA", "CIPLA": "CNXPHARMA", "DRREDDY": "CNXPHARMA", "AUROPHARMA": "CNXPHARMA",
        "BAJFINANCE": "CNXCONSUM", "BAJAJFINSV": "CNXCONSUM", "CHOLAFIN": "CNXFIN", "ABCAPITAL": "CNXFIN", "HDFCLIFE": "CNXCONSUM",
        "ITC": "CNXFMCG", "BHARTIARTL": "CNXINFRA", "ADANIPORTS": "CNXINFRA", "AMBUJACEM": "CNXINFRA",
        "DLF": "CNXREALTY", "ASTRAL": "CNXINDUSTRIAL", "BEL": "CNXINDUSTRIAL", "BHEL": "CNXINDUSTRIAL"
    }

    for fname in files:
        sym = fname.replace(".pkl", "")
        fpath = os.path.join(CACHE_DIR, fname)
        try:
            with open(fpath, "rb") as f:
                df = pickle.load(f)
            if not df.empty and len(df) > 10000:
                ysym = f"{sym}.NS"
                stock_dfs[ysym] = df
                universe.append({
                    "symbol": sym,
                    "yahoo": ysym,
                    "sector": sec_map.get(sym, "OTHER")
                })
        except Exception:
            pass

    print(f"  Loaded {len(stock_dfs)} symbols into backtest universe.")
    index_data = {"NIFTY 50": df_nifty, "^NSEI": df_nifty}
    return stock_dfs, index_data, universe

BAR = datetime.timedelta(minutes=5)

def run_multi_year_simulation(strategy_name: str, stock_dfs, index_data, universe, base_sim=None, starting_capital: float = 40_000.0):
    print(f"\n==========================================================================", flush=True)
    print(f"INITIALIZING 5-YEAR BACKTEST: [{strategy_name}] | Capital: Rs. {starting_capital:,.2f}", flush=True)
    print(f"==========================================================================", flush=True)
    
    t0 = time.time()
    first_date = index_data["NIFTY 50"].index[0].date()
    sim = Simulator(mode="replay", session_date=first_date, strategy=strategy_name, starting_capital=starting_capital)
    
    if base_sim is not None:
        print("Reusing precomputed symbol indicator arrays...", flush=True)
        sim.syms = base_sim.syms
        sim.indices = base_sim.indices
        sim.nifty_pos = base_sim.nifty_pos
        sim.nifty_e10 = base_sim.nifty_e10
        sim.nifty_e21 = base_sim.nifty_e21
        sim.nifty_e30 = base_sim.nifty_e30
        sim._available_dates_cache = getattr(base_sim, "_available_dates_cache", None)
    else:
        print("Computing vectorized technical indicators across all symbols...", flush=True)
        sim.load(stock_dfs, universe, index_data)
        t_load = time.time()
        print(f"Indicators precomputed in {t_load - t0:.1f}s.", flush=True)

    dates = sim.available_dates()
    print(f"Total trading days available: {len(dates):,} days ({dates[0]} to {dates[-1]}).", flush=True)

    all_closed_trades = []
    daily_summaries = []
    
    cur_capital = starting_capital
    peak_capital = starting_capital
    max_drawdown_rs = 0.0
    max_drawdown_pct = 0.0

    t_loop_start = time.time()
    for d_idx, d_str in enumerate(dates):
        d_ts = pd.Timestamp(d_str)
        bar_times = sim.session_bar_times(d_ts)
        if not bar_times:
            continue

        # Reset intraday tracking for the new session
        sim.traded_symbols_today.clear()
        sim.pending.clear()
        sim.open.clear()

        # Step through every 5m bar in the session
        for bt in bar_times:
            sim.step(bt)

        # Force square-off at session end (15:05-15:25) if any position remains open
        if sim.open:
            last_bt = bar_times[-1]
            for s in list(sim.open.keys()):
                sd = sim.syms.get(s)
                if sd and last_bt in sd.pos:
                    i = sd.pos[last_bt]
                    sim._close(s, sd.c[i], last_bt, "EOD_FORCE_SQUAREOFF", market=True)

        # Collect session trades
        day_trades = [t for t in sim.closed]
        # Clear closed list for next day, but accumulate in master list
        all_closed_trades.extend(day_trades)
        sim.closed.clear()

        # Update running capital and drawdown
        day_net = sum(t["net"] for t in day_trades)
        cur_capital += day_net
        if cur_capital > peak_capital:
            peak_capital = cur_capital
        dd_rs = peak_capital - cur_capital
        dd_pct = (dd_rs / peak_capital) * 100.0 if peak_capital > 0 else 0.0
        if dd_rs > max_drawdown_rs:
            max_drawdown_rs = dd_rs
        if dd_pct > max_drawdown_pct:
            max_drawdown_pct = dd_pct

        daily_summaries.append({
            "date": d_str,
            "trades": len(day_trades),
            "net": day_net,
            "capital": cur_capital,
            "drawdown_rs": dd_rs,
            "drawdown_pct": dd_pct
        })

        if (d_idx + 1) % 250 == 0 or (d_idx + 1) == len(dates):
            yr = d_str[:4]
            print(f"  Day {d_idx + 1:4d}/{len(dates)} ({d_str}) | Trades: {len(all_closed_trades):5d} | Cap: Rs. {cur_capital:,.2f} | MaxDD: {max_drawdown_pct:.1f}%", flush=True)

    t_loop_end = time.time()
    print(f"Simulation completed in {t_loop_end - t_loop_start:.1f}s.", flush=True)

    return {
        "strategy": strategy_name,
        "starting_capital": starting_capital,
        "ending_capital": cur_capital,
        "total_trades": len(all_closed_trades),
        "trades": all_closed_trades,
        "daily": daily_summaries,
        "max_drawdown_rs": max_drawdown_rs,
        "max_drawdown_pct": max_drawdown_pct
    }, sim

def analyze_and_print_results(res):
    strat = res["strategy"]
    trades = res["trades"]
    start_cap = res["starting_capital"]
    end_cap = res["ending_capital"]
    tot_trades = len(trades)
    
    if tot_trades == 0:
        print(f"\n--- No trades executed for {strat} ---")
        return

    wins = [t for t in trades if t["net"] > 0]
    losses = [t for t in trades if t["net"] <= 0]
    win_rate = (len(wins) / tot_trades) * 100.0 if tot_trades else 0.0

    gross_pnl = sum(t["gross"] for t in trades)
    tot_costs = sum(t["costs"] for t in trades)
    net_pnl = sum(t["net"] for t in trades)
    net_roi = ((end_cap - start_cap) / start_cap) * 100.0

    tot_gain = sum(t["net"] for t in wins)
    tot_loss = abs(sum(t["net"] for t in losses)) if losses else 1.0
    profit_factor = tot_gain / tot_loss if tot_loss > 0 else float("inf")

    # Detailed cost breakdown
    brokerage = sum(t.get("cost_detail", {}).get("brokerage", 0.0) for t in trades)
    stt = sum(t.get("cost_detail", {}).get("stt", 0.0) for t in trades)
    exchange = sum(t.get("cost_detail", {}).get("exchange", 0.0) for t in trades)
    gst = sum(t.get("cost_detail", {}).get("gst", 0.0) for t in trades)
    stamp = sum(t.get("cost_detail", {}).get("stamp", 0.0) for t in trades)
    slippage = sum(t.get("cost_detail", {}).get("slippage", 0.0) for t in trades)

    # Year-by-year breakdown
    df_trades = pd.DataFrame(trades)
    df_trades["year"] = df_trades["entry_time"].astype(str).str[:4]
    
    print("\n" + "=" * 80)
    print(f"5-YEAR PERFORMANCE REPORT: [{strat}] (2022 - 2026)")
    print("=" * 80)
    print(f"Starting Capital:       Rs. {start_cap:,.2f}")
    print(f"Ending Capital:         Rs. {end_cap:,.2f}")
    print(f"Net Profit:             Rs. {net_pnl:+,.2f} ({net_roi:+.1f}%)")
    print(f"Gross Profit:           Rs. {gross_pnl:+,.2f}")
    print(f"Total Friction & Taxes: Rs. {tot_costs:,.2f} ({tot_costs/gross_pnl*100.0 if gross_pnl else 0:.1f}% of gross)")
    print(f"Total Trades:           {tot_trades:,} (Wins: {len(wins):,}, Losses: {len(losses):,})")
    print(f"Win Rate:               {win_rate:.2f}%")
    print(f"Profit Factor:          {profit_factor:.2f}")
    print(f"Max Drawdown:           Rs. {res['max_drawdown_rs']:,.2f} ({res['max_drawdown_pct']:.2f}%)")
    
    print("\n--- Statutory Friction & Tax Breakdown ---")
    print(f"  Brokerage (Fyers):    Rs. {brokerage:,.2f}")
    print(f"  Securities STT:       Rs. {stt:,.2f}")
    print(f"  NSE/BSE Exchange:     Rs. {exchange:,.2f}")
    print(f"  GST (18%):            Rs. {gst:,.2f}")
    print(f"  Stamp Duty:           Rs. {stamp:,.2f}")
    print(f"  Market Slippage:      Rs. {slippage:,.2f}")

    print("\n--- Year-by-Year Performance Breakdown ---")
    print(f"{'Year':6s} | {'Trades':7s} | {'Win Rate':9s} | {'Gross P&L':14s} | {'Friction':12s} | {'Net P&L':14s} | {'Yearly ROI':10s}")
    print("-" * 84)
    
    years = sorted(df_trades["year"].unique())
    for yr in years:
        sub = df_trades[df_trades["year"] == yr]
        y_trades = len(sub)
        y_wins = len(sub[sub["net"] > 0])
        y_wr = (y_wins / y_trades) * 100.0 if y_trades else 0.0
        y_gross = sub["gross"].sum()
        y_costs = sub["costs"].sum()
        y_net = sub["net"].sum()
        y_roi = (y_net / start_cap) * 100.0
        print(f"{yr:6s} | {y_trades:7d} | {y_wr:8.1f}% | Rs. {y_gross:+10,.2f} | Rs. {y_costs:8,.2f} | Rs. {y_net:+10,.2f} | {y_roi:+9.1f}%")

    return {
        "strat": strat,
        "start_cap": start_cap,
        "end_cap": end_cap,
        "net_pnl": net_pnl,
        "gross_pnl": gross_pnl,
        "costs": tot_costs,
        "roi": net_roi,
        "trades": tot_trades,
        "win_rate": win_rate,
        "profit_factor": profit_factor,
        "max_dd_pct": res["max_drawdown_pct"],
        "max_dd_rs": res["max_drawdown_rs"]
    }

if __name__ == '__main__':
    stock_dfs, index_data, universe = load_cached_universe()
    
    # Run Quacknutzz_Strat
    res_opt, base_sim = run_multi_year_simulation("quacknutzz_strat", stock_dfs, index_data, universe, base_sim=None, starting_capital=40_000.0)
    summary_opt = analyze_and_print_results(res_opt)

    # Run Baseline
    res_base, _ = run_multi_year_simulation("v6_baseline", stock_dfs, index_data, universe, base_sim=base_sim, starting_capital=40_000.0)
    summary_base = analyze_and_print_results(res_base)
    
    # Save combined results for reporting
    with open(os.path.join(ROOT_DIR, "scratch", "multi_year_backtest_results.pkl"), "wb") as f:
        pickle.dump({"opt": summary_opt, "base": summary_base, "raw_opt": res_opt, "raw_base": res_base}, f)
    print("\nResults successfully saved to scratch/multi_year_backtest_results.pkl.")
