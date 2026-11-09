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

# Filter for the last ~30 trading days (starting around late September / Oct 1, 2026)
recent_dates = [d for d in dates if d >= "2026-10-01"]
print(f"Testing last month: {len(recent_dates)} trading days ({recent_dates[0]} to {recent_dates[-1]}).")

def run_period_simulation(name: str, fix_a: bool, fix_b: bool, fix_c: bool, strategy: str = "quacknutzz_strat", starting_capital: float = 40_000.0):
    first_date = pd.Timestamp(recent_dates[0]).date()
    sim = Simulator(mode="replay", session_date=first_date, strategy=strategy, starting_capital=starting_capital)
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
    sim._available_dates_cache = recent_dates

    sim.fix_a = fix_a
    sim.fix_b = fix_b
    sim.fix_c = fix_c
    sim.fast_backtest = False

    all_trades = []
    daily_stats = []
    cur_cap = starting_capital
    peak_cap = starting_capital
    max_dd_rs = 0.0
    max_dd_pct = 0.0

    for d_str in recent_dates:
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

        day_gross = sum(t["gross"] for t in day_trades)
        day_costs = sum(t["costs"] for t in day_trades)
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

        daily_stats.append({
            "date": d_str,
            "trades": len(day_trades),
            "wins": sum(1 for t in day_trades if t["net"] > 0),
            "gross": day_gross,
            "costs": day_costs,
            "net": day_net,
            "capital": cur_cap,
            "dd_pct": dd_pct
        })

    df = pd.DataFrame(all_trades)
    df["pos_val"] = df["entry"] * df["qty"]
    df["gross_pct"] = (df["gross"] / df["pos_val"]) * 100.0
    df["net_pct"] = (df["net"] / df["pos_val"]) * 100.0

    tot_trades = len(df)
    wins = df[df["net"] > 0]
    win_rate = (len(wins) / tot_trades) * 100.0 if tot_trades else 0.0
    gross_pnl = float(df["gross"].sum())
    tot_costs = float(df["costs"].sum())
    net_pnl = float(df["net"].sum())
    net_roi = (net_pnl / starting_capital) * 100.0

    gross_wins = float(df[df["gross"] > 0]["gross"].sum()) if len(df[df["gross"] > 0]) else 0.0
    gross_loss = abs(float(df[df["gross"] <= 0]["gross"].sum())) if len(df[df["gross"] <= 0]) else 0.001
    gross_pf = gross_wins / gross_loss if gross_loss > 0 else 99.9

    net_wins = float(wins["net"].sum()) if len(wins) else 0.0
    net_loss = abs(float(df[df["net"] <= 0]["net"].sum())) if len(df[df["net"] <= 0]) else 0.001
    net_pf = net_wins / net_loss if net_loss > 0 else 99.9

    return {
        "name": name,
        "trades": tot_trades,
        "win_rate": win_rate,
        "gross_pnl": gross_pnl,
        "tot_costs": tot_costs,
        "net_pnl": net_pnl,
        "net_roi": net_roi,
        "ending_capital": cur_cap,
        "max_dd_pct": max_dd_pct,
        "max_dd_rs": max_dd_rs,
        "gross_pf": gross_pf,
        "net_pf": net_pf,
        "daily": daily_stats,
        "df_trades": df
    }

print("\n--- RUNNING HEAD-TO-HEAD ON RECENT 1-MONTH PERIOD (OCT 1 - NOV 8, 2026) ---")
res_new = run_period_simulation("NEW_STRATEGY (Fix A+B+C)", fix_a=True, fix_b=True, fix_c=True, strategy="quacknutzz_strat", starting_capital=40_000.0)
res_old = run_period_simulation("OLD_BASELINE (Quacknutzz Baseline)", fix_a=False, fix_b=False, fix_c=False, strategy="quacknutzz_strat", starting_capital=40_000.0)
res_v6 = run_period_simulation("V6_ORIGINAL (Dad's Baseline)", fix_a=False, fix_b=False, fix_c=False, strategy="v6_baseline", starting_capital=40_000.0)

print(f"\n{'Metric':25s} | {'New Fix A+B+C':20s} | {'Old Quacknutzz':20s} | {'Original V6 Baseline':20s}")
print("-" * 95)
print(f"{'Total Trades':25s} | {res_new['trades']:<20d} | {res_old['trades']:<20d} | {res_v6['trades']:<20d}")
print(f"{'Win Rate':25s} | {res_new['win_rate']:<19.1f}% | {res_old['win_rate']:<19.1f}% | {res_v6['win_rate']:<19.1f}%")
print(f"{'Gross P&L':25s} | Rs. {res_new['gross_pnl']:<16,.2f} | Rs. {res_old['gross_pnl']:<16,.2f} | Rs. {res_v6['gross_pnl']:<16,.2f}")
print(f"{'Total Costs & Taxes':25s} | Rs. {res_new['tot_costs']:<16,.2f} | Rs. {res_old['tot_costs']:<16,.2f} | Rs. {res_v6['tot_costs']:<16,.2f}")
print(f"{'Net P&L':25s} | Rs. {res_new['net_pnl']:<16,.2f} | Rs. {res_old['net_pnl']:<16,.2f} | Rs. {res_v6['net_pnl']:<16,.2f}")
print(f"{'Net ROI %':25s} | {res_new['net_roi']:<+19.1f}% | {res_old['net_roi']:<+19.1f}% | {res_v6['net_roi']:<+19.1f}%")
print(f"{'Ending Capital (40k)':25s} | Rs. {res_new['ending_capital']:<16,.2f} | Rs. {res_old['ending_capital']:<16,.2f} | Rs. {res_v6['ending_capital']:<16,.2f}")
print(f"{'Max Drawdown':25s} | {res_new['max_dd_pct']:<19.1f}% | {res_old['max_dd_pct']:<19.1f}% | {res_v6['max_dd_pct']:<19.1f}%")
print(f"{'Net Profit Factor':25s} | {res_new['net_pf']:<20.2f} | {res_old['net_pf']:<20.2f} | {res_v6['net_pf']:<20.2f}")

print("\n--- RECENT 1-MONTH DAILY BREAKDOWN (NEW STRATEGY FIX A+B+C) ---")
print(f"{'Date':10s} | {'Trades':6s} | {'Win Rate':8s} | {'Gross P&L':12s} | {'Costs':10s} | {'Net P&L':12s} | {'Capital':12s} | {'Drawdown':8s}")
print("-" * 90)
for d in res_new["daily"]:
    wr_d = (d["wins"] / d["trades"] * 100) if d["trades"] else 0.0
    print(f"{d['date']:10s} | {d['trades']:6d} | {wr_d:7.1f}% | Rs. {d['gross']:+9,.2f} | Rs. {d['costs']:7,.2f} | Rs. {d['net']:+9,.2f} | Rs. {d['capital']:9,.2f} | {d['dd_pct']:6.1f}%")

# Save trades for deep inspection
with open(os.path.join(ROOT_DIR, "scratch", "last_month_simulation_results.pkl"), "wb") as f:
    pickle.dump({"new": res_new, "old": res_old, "v6": res_v6}, f)
print("\nSaved last month results to scratch/last_month_simulation_results.pkl.")
