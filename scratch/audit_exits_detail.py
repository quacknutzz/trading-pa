import os
import sys
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

# Run a test where we explicitly inspect every exit reason of ABC
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

sim.fix_a = True
sim.fix_b = True
sim.fix_c = True
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

print("\n--- ABC EXITS BREAKDOWN ---")
for r, g in df.groupby(df["reason"].str.split(" ").str[0]):
    wr = (g["net"] > 0).mean() * 100.0
    gp = g["gross_pct"].mean()
    np_ = g["net_pct"].mean()
    sg = g["gross"].sum()
    sn = g["net"].sum()
    print(f"{r:22s} | n={len(g):4d} | WR={wr:5.1f}% | GrossPct={gp:+.3f}% | NetPct={np_:+.3f}% | Gross=Rs.{sg:+9,.0f} | Net=Rs.{sn:+9,.0f}")

# ULTRA CONSERVATIVE STRESS TEST:
# What if EVERY trade that is NOT a TARGET or STOP_LOSS pays 0.05% slippage?
def stress_costs(row):
    r = str(row["reason"])
    # If it is TARGET, slippage = 0
    # If it is STOP_LOSS, slippage = 0 (SL-LIMIT)
    # What about PROFIT_LOCK? It is also a trailed SL-LIMIT!
    # What about EXHAUSTION, REVERSAL, 3_PM_EXIT? Add 0.05% slippage if not already charged!
    cd = row["cost_detail"]
    cur_slip = cd.get("slippage", 0.0) if isinstance(cd, dict) else 0.0
    pv = row["exit"] * row["qty"]
    extra_slip = 0.0
    if "EXHAUSTION" in r or "REVERSAL" in r or "3_PM" in r:
        if cur_slip == 0:
            extra_slip = pv * 0.0005
    return row["net"] - extra_slip

df["stressed_net"] = df.apply(stress_costs, axis=1)
stressed_total_net = df["stressed_net"].sum()
print(f"\nIf 0.05% market slippage is penalized on ALL EXHAUSTION, REVERSAL, and 3 PM exits:")
print(f"  Standard Net P&L: Rs. {df['net'].sum():+,.2f}")
print(f"  Stressed Net P&L: Rs. {stressed_total_net:+,.2f} (Delta: Rs. {stressed_total_net - df['net'].sum():,.2f})")
