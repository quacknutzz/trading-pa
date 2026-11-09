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
first_date = pd.Timestamp(dates[0]).date()

def test_config(name, fix_a, fix_b, fix_c, force_market_exits):
    sim = Simulator(mode="replay", session_date=first_date, strategy="quacknutzz_strat", starting_capital=40_000.0)
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
    cur_cap = 40000.0
    peak_cap = 40000.0
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
                    sim._close(s, sd.c[i], last_bt, "EOD_FORCE_SQUAREOFF", market=True)

        day_trades = list(sim.closed)
        all_trades.extend(day_trades)
        sim.closed.clear()

        day_net = sum(t["net"] for t in day_trades)
        cur_cap += day_net
        if cur_cap > peak_cap:
            peak_cap = cur_cap
        dd_pct = ((peak_cap - cur_cap) / peak_cap) * 100.0 if peak_cap > 0 else 0.0
        if dd_pct > max_dd_pct:
            max_dd_pct = dd_pct

    df = pd.DataFrame(all_trades)
    df["pos_val"] = df["entry"] * df["qty"]
    df["gross_pct"] = (df["gross"] / df["pos_val"]) * 100.0
    df["net_pct"] = (df["net"] / df["pos_val"]) * 100.0
    df["year"] = df["entry_time"].astype(str).str[:4]

    tot_trades = len(df)
    wins = df[df["net"] > 0]
    win_rate = (len(wins) / tot_trades) * 100.0
    gross_pnl = float(df["gross"].sum())
    tot_costs = float(df["costs"].sum())
    net_pnl = float(df["net"].sum())
    net_roi = (net_pnl / 40000.0) * 100.0
    gross_wins = float(df[df["gross"] > 0]["gross"].sum())
    gross_loss = abs(float(df[df["gross"] <= 0]["gross"].sum()))
    gross_pf = gross_wins / gross_loss if gross_loss > 0 else 99.9
    net_wins = float(wins["net"].sum())
    net_loss = abs(float(df[df["net"] <= 0]["net"].sum()))
    net_pf = net_wins / net_loss if net_loss > 0 else 99.9

    yearly_str = " | ".join(f"{y}: Rs.{g['net'].sum():+8,.0f}" for y, g in df.groupby("year"))

    print(f"\n[{name}]")
    print(f"  Trades: {tot_trades:,} | WR: {win_rate:.1f}% | Net PF: {net_pf:.2f} | Gross PF: {gross_pf:.2f}")
    print(f"  Gross P&L: Rs. {gross_pnl:+,.2f} | Costs: Rs. {tot_costs:,.2f} | Net P&L: Rs. {net_pnl:+,.2f} ({net_roi:+.1f}%)")
    print(f"  Max Drawdown: {max_dd_pct:.1f}%")
    print(f"  Yearly: {yearly_str}")
    return df

print("Testing ABC with current code...")
df_current = test_config("ABC_CURRENT", True, True, True, False)
