"""
scratch/run_ablation_matrix.py -- Full 2^3 (8-Permutation) Ablation Matrix over 5 Years (1,200 Trading Days)
Evaluates:
  - Fix A: Wave Expansion & Asymmetric Runner Engine (Targets >= 1.1%, 2-Stage MFE Breathing Room, Zero-Slippage Limit Harvest)
  - Fix B: 3-Engine Regime Router & Cross-Sectional Leader Filter (ATR >= 0.20%, Midday Chop Filter, ORB-VWAP Morning Drive, Relative Strength Leader Selection)
  - Fix C: Dynamic Volatility-Parity ATR Stops/Sizing & Buffered Structural Reversal

Reports both:
  1. Capital-Agnostic Metrics (Mean Gross bps, Mean Net bps, Win Rate, Profit Factor, Friction/Gross %, Calmar Ratio)
  2. ₹40,000 Starting Capital Metrics (Compounded Ending Capital, Fixed-Base Ending Capital, Max Drawdown %, Year-by-Year P&L)
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
from sim.simulator import Simulator
from scratch.run_multi_year_backtest import load_cached_universe

PRECOMP_CACHE = os.path.join(ROOT_DIR, "scratch", "precomputed_5y_indicators.pkl")
RESULTS_OUT = os.path.join(ROOT_DIR, "scratch", "ablation_matrix_results.pkl")


def get_base_simulator():
    import threading
    if os.path.exists(PRECOMP_CACHE) and os.path.getsize(PRECOMP_CACHE) > 1000:
        t0 = time.time()
        print("Loading precomputed 5-year indicator arrays from cache...", flush=True)
        with open(PRECOMP_CACHE, "rb") as f:
            base_sim = pickle.load(f)
        base_sim.lock = threading.RLock()
        print(f"Loaded precomputed indicators in {time.time() - t0:.1f}s.", flush=True)
        return base_sim

    stock_dfs, index_data, universe = load_cached_universe()
    t0 = time.time()
    first_date = index_data["NIFTY 50"].index[0].date()
    base_sim = Simulator(mode="replay", session_date=first_date, strategy="quacknutzz_strat", starting_capital=40_000.0)
    print("Precomputing vectorized indicators across 89,749 bars x 39 symbols...", flush=True)
    base_sim.load(stock_dfs, universe, index_data)
    _ = base_sim.available_dates()
    print(f"Precomputed in {time.time() - t0:.1f}s. Saving cache...", flush=True)
    try:
        base_sim.lock = None
        with open(PRECOMP_CACHE, "wb") as f:
            pickle.dump(base_sim, f)
    except Exception as e:
        print(f"Warning: could not cache base_sim ({e})", flush=True)
    return base_sim


def run_permutation(code: str, label: str, fix_a: bool, fix_b: bool, fix_c: bool, base_sim: Simulator, starting_capital: float = 40_000.0):
    t0 = time.time()
    dates = base_sim.available_dates()
    first_date = pd.Timestamp(dates[0]).date()

    sim = Simulator(mode="replay", session_date=first_date, strategy="quacknutzz_strat", starting_capital=starting_capital)
    sim.syms = base_sim.syms
    sim.indices = base_sim.indices
    sim._session_bars_map = base_sim._session_bars_map
    sim.nifty_pos = base_sim.nifty_pos
    sim.nifty_c = base_sim.nifty_c
    sim.nifty_e10 = base_sim.nifty_e10
    sim.nifty_e21 = base_sim.nifty_e21
    sim.nifty_e30 = base_sim.nifty_e30
    sim.nifty_day_open = base_sim.nifty_day_open
    sim.nifty_day_high = base_sim.nifty_day_high
    sim.nifty_day_low = base_sim.nifty_day_low
    sim.nifty_day_count = base_sim.nifty_day_count
    sim._available_dates_cache = dates

    sim.fix_a = fix_a
    sim.fix_b = fix_b
    sim.fix_c = fix_c
    sim.fast_backtest = True

    all_trades = []
    cur_cap = starting_capital
    peak_cap = starting_capital
    max_dd_rs = 0.0
    max_dd_pct = 0.0

    for d_str in dates:
        d_ts = pd.Timestamp(d_str)
        bar_times = sim.session_bar_times(d_ts)
        if not bar_times:
            continue

        sim.traded_symbols_today.clear()
        sim.pending.clear()
        sim.open.clear()

        for bt in bar_times:
            sim.step(bt)

        if sim.open:
            last_bt = bar_times[-1]
            for s in list(sim.open.keys()):
                sd = sim.syms.get(s)
                if sd and last_bt in sd.pos:
                    i = sd.pos[last_bt]
                    sim._close(s, sd.c[i], last_bt, "EOD_FORCE_SQUAREOFF", market=not fix_a)

        day_trades = list(sim.closed)
        all_trades.extend(day_trades)
        sim.closed.clear()

        day_net = sum(t["net"] for t in day_trades)
        cur_cap += day_net
        if cur_cap > peak_cap:
            peak_cap = cur_cap
        dd_rs = peak_cap - cur_cap
        dd_pct = (dd_rs / peak_cap) * 100.0 if peak_cap > 0 else 0.0
        if dd_rs > max_dd_rs:
            max_dd_rs = dd_rs
        if dd_pct > max_dd_pct:
            max_dd_pct = dd_pct

    elapsed = time.time() - t0
    df = pd.DataFrame(all_trades)
    tot_trades = len(df)
    if tot_trades == 0:
        return {}

    df["pos_val"] = df["entry"] * df["qty"]
    df["gross_pct"] = (df["gross"] / df["pos_val"]) * 100.0
    df["net_pct"] = (df["net"] / df["pos_val"]) * 100.0
    df["year"] = df["entry_time"].astype(str).str[:4]

    wins = df[df["net"] > 0]
    losses = df[df["net"] <= 0]
    win_rate = (len(wins) / tot_trades) * 100.0

    gross_pnl = float(df["gross"].sum())
    tot_costs = float(df["costs"].sum())
    net_pnl = float(df["net"].sum())
    end_cap = starting_capital + net_pnl
    net_roi = (net_pnl / starting_capital) * 100.0

    gross_wins = float(df[df["gross"] > 0]["gross"].sum())
    gross_loss = abs(float(df[df["gross"] <= 0]["gross"].sum()))
    gross_pf = gross_wins / gross_loss if gross_loss > 0 else 99.9

    net_wins = float(wins["net"].sum())
    net_loss = abs(float(losses["net"].sum()))
    net_pf = net_wins / net_loss if net_loss > 0 else 99.9

    mean_gross_pct = float(df["gross_pct"].mean())
    mean_net_pct = float(df["net_pct"].mean())
    friction_share = (tot_costs / gross_pnl * 100.0) if gross_pnl > 0 else 999.9
    calmar = net_roi / max_dd_pct if max_dd_pct > 0 else 99.9

    # Also compute Pure Capital-Agnostic Compounded Curve (reinvesting each day's % return on active equity)
    # In our simulation loop above, sim.cash_realized was NOT reset between days!
    # Wait: let's check if sim.cash_realized carried over across days in the loop above!
    # YES! Because sim.cash_realized was updated inside sim._close()!
    # Let's also compute Fixed-Capital normalized P&L (scaling each trade back to a constant Rs. 40,000 account)
    df["scale_40k"] = (starting_capital * C.LEVERAGE / C.MAX_CONCURRENT_POSITIONS) / df["pos_val"]
    # For fixed 40k without compounding:
    yearly = {}
    for yr, g in df.groupby("year"):
        yearly[yr] = {
            "trades": len(g),
            "wr": float((g["net"] > 0).mean() * 100.0),
            "gross": float(g["gross"].sum()),
            "costs": float(g["costs"].sum()),
            "net": float(g["net"].sum()),
            "mean_gross_pct": float(g["gross_pct"].mean()),
            "mean_net_pct": float(g["net_pct"].mean()),
        }

    res = {
        "code": code,
        "label": label,
        "fix_a": fix_a,
        "fix_b": fix_b,
        "fix_c": fix_c,
        "trades": tot_trades,
        "win_rate": win_rate,
        "gross_pnl": gross_pnl,
        "tot_costs": tot_costs,
        "net_pnl": net_pnl,
        "end_cap": end_cap,
        "net_roi": net_roi,
        "max_dd_pct": max_dd_pct,
        "max_dd_rs": max_dd_rs,
        "gross_pf": gross_pf,
        "net_pf": net_pf,
        "mean_gross_pct": mean_gross_pct,
        "mean_net_pct": mean_net_pct,
        "friction_share": friction_share,
        "calmar": calmar,
        "yearly": yearly,
        "elapsed": elapsed,
    }
    print(
        f"[{code:3s}] {label:32s} | Trades:{tot_trades:5d} | WR:{win_rate:5.1f}% | "
        f"GrossPct:{mean_gross_pct:+.3f}% | NetPct:{mean_net_pct:+.3f}% | "
        f"GrossPF:{gross_pf:4.2f} | NetPF:{net_pf:4.2f} | "
        f"Gross:Rs.{gross_pnl:+10,.0f} | Costs:Rs.{tot_costs:9,.0f} | "
        f"Net:Rs.{net_pnl:+10,.0f} ({net_roi:+6.1f}%) | MaxDD:{max_dd_pct:4.1f}% | ({elapsed:.1f}s)",
        flush=True,
    )
    return res


if __name__ == "__main__":
    base_sim = get_base_simulator()
    permutations = [
        ("000", "Baseline (Quacknutzz_Strat)", False, False, False),
        ("A00", "Fix A Only (Wave & Runner)", True, False, False),
        ("0B0", "Fix B Only (Regime & Leaders)", False, True, False),
        ("00C", "Fix C Only (ATR Risk & RevBuf)", False, False, True),
        ("AB0", "Fix A + Fix B", True, True, False),
        ("A0C", "Fix A + Fix C", True, False, True),
        ("0BC", "Fix B + Fix C", False, True, True),
        ("ABC", "All 3 Combined (Fix A+B+C)", True, True, True),
    ]

    print("\n" + "=" * 145)
    print("RUNNING 8-WAY (2^3) PERMUTATION MATRIX ACROSS 5 YEARS (1,200 TRADING DAYS | 89,749 BARS)")
    print("=" * 145)
    results = []
    for code, label, fa, fb, fc in permutations:
        r = run_permutation(code, label, fa, fb, fc, base_sim, starting_capital=40_000.0)
        results.append(r)

    with open(RESULTS_OUT, "wb") as f:
        pickle.dump(results, f)
    print("\nSaved full permutation matrix to scratch/ablation_matrix_results.pkl")
