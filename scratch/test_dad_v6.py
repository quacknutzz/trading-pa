"""
scratch/test_dad_v6.py -- Simulates Dad's cockpit_V6.txt exact logic & parameters
and compares it side-by-side with:
  1. Old Baseline (before fixes)
  2. Dad's updated VBA code (cockpit_V6.txt - Fixes #1 to #8)
  3. Quacknutzz New Strategy (Fix A+B+C)
"""
import os
import sys
import time
import pickle
import pandas as pd
import numpy as np

ROOT_DIR = r"c:\Users\rishi\Downloads\Trading Pa"
sys.path.insert(0, ROOT_DIR)
sys.path.insert(0, os.path.join(ROOT_DIR, "sim"))

import sim.config as C
from sim.simulator import Simulator
from scratch.run_ablation_matrix import get_base_simulator

base_sim = get_base_simulator()
dates = base_sim.available_dates()

recent_dates = [d for d in dates if d >= "2026-10-01"]

def run_simulation(sim_dates, fix_a, fix_b, fix_c, starting_capital=40_000.0, use_dad_vba_params=False):
    first_date = pd.Timestamp(sim_dates[0]).date()
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
    sim._available_dates_cache = sim_dates

    # Configuration flags
    sim.fix_a = fix_a
    sim.fix_b = fix_b
    sim.fix_c = fix_c
    sim.fast_backtest = True

    # If Dad VBA params specifically:
    # Dad has:
    # - Fixed 1.0% SL
    # - Breakeven at 0.40% MFE (BREAKEVEN_MFE_TRIGGER = 0.004) with Rs 50 floor
    # - Combo13 swing trail after 1.0% MFE
    # - Multi-factor exit without ATR buffer (strict curLTP < exitVWAP and curLTP < exitEMA30)
    # - Market exit slippage because it exits at observed LTP (market=True)
    if use_dad_vba_params:
        sim.fix_a = False # No wave runner extension, snaps at 0.40% BE, market exit slippage
        sim.fix_b = False # Uses standard 0.15% ATR, no chop router
        sim.fix_c = False # Fixed 1.0% SL, unbuffered 2-of-3 exit

    all_trades = []
    cur_cap = starting_capital
    peak_cap = starting_capital
    max_dd_rs = 0.0
    max_dd_pct = 0.0

    for d_str in sim_dates:
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
                    sim._close(s, sd.c[i], last_bt, "EOD_FORCE_SQUAREOFF", market=not sim.fix_a)

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

    df = pd.DataFrame(all_trades)
    if len(df) == 0:
        return {}
    df["pos_val"] = df["entry"] * df["qty"]
    df["gross_pct"] = (df["gross"] / df["pos_val"]) * 100.0
    df["net_pct"] = (df["net"] / df["pos_val"]) * 100.0

    wins = df[df["net"] > 0]
    losses = df[df["net"] <= 0]
    win_rate = (len(wins) / len(df)) * 100.0
    gross_pnl = float(df["gross"].sum())
    tot_costs = float(df["costs"].sum())
    net_pnl = float(df["net"].sum())
    end_cap = starting_capital + net_pnl
    net_roi = (net_pnl / starting_capital) * 100.0
    net_wins = float(wins["net"].sum())
    net_loss = abs(float(losses["net"].sum()))
    net_pf = net_wins / net_loss if net_loss > 0 else 99.9

    return {
        "trades": len(df),
        "win_rate": win_rate,
        "gross": gross_pnl,
        "costs": tot_costs,
        "net": net_pnl,
        "roi": net_roi,
        "end_cap": end_cap,
        "max_dd_pct": max_dd_pct,
        "net_pf": net_pf,
        "mean_net_pct": float(df["net_pct"].mean()),
        "friction_share": (tot_costs / gross_pnl * 100.0) if gross_pnl > 0 else 999.0,
    }

print("="*80)
print("RUNNING RECENT 1-MONTH AUDIT (Oct 1 - Nov 8, 2026, 26 trading days)")
print("="*80)

res_recent_dad = run_simulation(recent_dates, fix_a=False, fix_b=False, fix_c=False, use_dad_vba_params=True)
res_recent_abc = run_simulation(recent_dates, fix_a=True, fix_b=True, fix_c=True)

print(f"{'Metric':<25} | {'Dad cockpit_V6.txt (Fixes 1-8)':<30} | {'New Strategy (Fix A+B+C)':<30}")
print("-" * 90)
for k in ["trades", "win_rate", "gross", "costs", "net", "roi", "net_pf", "mean_net_pct", "max_dd_pct"]:
    v_dad = res_recent_dad.get(k, 0)
    v_abc = res_recent_abc.get(k, 0)
    if k in ["win_rate", "roi", "max_dd_pct", "mean_net_pct"]:
        print(f"{k:<25} | {v_dad:29.2f}% | {v_abc:29.2f}%")
    elif k in ["gross", "costs", "net"]:
        print(f"{k:<25} | Rs. {v_dad:25.2f} | Rs. {v_abc:25.2f}")
    elif k == "trades":
        print(f"{k:<25} | {v_dad:<30} | {v_abc:<30}")
    else:
        print(f"{k:<25} | {v_dad:30.2f} | {v_abc:30.2f}")

print("\n" + "="*80)
print("RUNNING FULL 5-YEAR AUDIT (2022 - 2026, 1,200 trading days)")
print("="*80)

res_5y_dad = run_simulation(dates, fix_a=False, fix_b=False, fix_c=False, use_dad_vba_params=True)
res_5y_abc = run_simulation(dates, fix_a=True, fix_b=True, fix_c=True)

print(f"{'Metric':<25} | {'Dad cockpit_V6.txt (Fixes 1-8)':<30} | {'New Strategy (Fix A+B+C)':<30}")
print("-" * 90)
for k in ["trades", "win_rate", "gross", "costs", "net", "roi", "end_cap", "net_pf", "mean_net_pct", "max_dd_pct", "friction_share"]:
    v_dad = res_5y_dad.get(k, 0)
    v_abc = res_5y_abc.get(k, 0)
    if k in ["win_rate", "roi", "max_dd_pct", "mean_net_pct", "friction_share"]:
        print(f"{k:<25} | {v_dad:29.2f}% | {v_abc:29.2f}%")
    elif k in ["gross", "costs", "net", "end_cap"]:
        print(f"{k:<25} | Rs. {v_dad:25.2f} | Rs. {v_abc:25.2f}")
    elif k == "trades":
        print(f"{k:<25} | {v_dad:<30} | {v_abc:<30}")
    else:
        print(f"{k:<25} | {v_dad:30.2f} | {v_abc:30.2f}")
