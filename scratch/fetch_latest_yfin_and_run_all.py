"""
scratch/fetch_latest_yfin_and_run_all.py
1. Fetches the latest 5m bars from Yahoo Finance (yfinance) for the entire 139-stock universe + indices.
2. Updates sim/cache/nse_5m_60d.pkl with the latest Yahoo Finance data.
3. Runs the newly upgraded v6_profit_max model vs v6_baseline on:
   - The latest trading sessions fetched from Yahoo Finance (across Rs 50k, Rs 100k, Rs 500k)
   - The full multi-week historical dataset (and updates sim/journal/backtest_12week_report.json + replay journals).
"""
import sys, os, pickle, json
from datetime import datetime
import numpy as np
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, '.')
sys.path.insert(0, 'sim')
import config as C
import datafeed
from simulator import Simulator
from server import write_journal

CACHE_FILE = 'sim/cache/nse_5m_60d.pkl'
with open(CACHE_FILE, 'rb') as f:
    payload = pickle.load(f)

raw = payload['raw']
idx = payload['idx']
universe = payload.get('universe') or datafeed.load_universe()

sym0 = list(raw.keys())[0]
before_dates = sorted(list(set(raw[sym0].index.date)))
print(f"Cache on disk: {len(before_dates)} days ({before_dates[0]} -> {before_dates[-1]})")

if str(before_dates[-1]) < "2026-11-05":
    print("Fetching latest 5d of 5m bars from Yahoo Finance (yfinance) for 139 stocks + indices...")
    all_tickers = [u['yahoo'] for u in universe] + list(C.INDEX_TICKERS.values())
    fresh_data = datafeed.fetch(all_tickers, period='5d')
    print(f"Fetched fresh data for {len(fresh_data)}/{len(all_tickers)} tickers from Yahoo Finance.")

    if len(fresh_data) > 0:
        for idx_name, tick in C.INDEX_TICKERS.items():
            if tick in fresh_data and len(fresh_data[tick]) > 0:
                df_new = fresh_data[tick]
                df_old = idx.get(idx_name)
                idx[idx_name] = df_new if df_old is None else pd.concat([df_old, df_new]).pipe(lambda x: x[~x.index.duplicated(keep='last')]).sort_index()

        updated_stocks = 0
        for u in universe:
            ysym = u['yahoo']
            if ysym in fresh_data and len(fresh_data[ysym]) > 0:
                df_new = fresh_data[ysym]
                df_old = raw.get(ysym)
                raw[ysym] = df_new if df_old is None else pd.concat([df_old, df_new]).pipe(lambda x: x[~x.index.duplicated(keep='last')]).sort_index()
                updated_stocks += 1

        with open(CACHE_FILE, 'wb') as f:
            pickle.dump({'raw': raw, 'idx': idx, 'universe': universe}, f)
        print(f"Saved updated cache to {CACHE_FILE} ({updated_stocks} stocks updated).")
else:
    print("Latest Yahoo Finance bars through 2026-11-05 already present in cache.")

all_dates = sorted(list(set(raw[sym0].index.date)))
print(f"Cache after fetch: {len(all_dates)} days ({all_dates[0]} -> {all_dates[-1]})")
latest_dates = all_dates[-3:]
print(f"Latest 3 sessions in dataset: {[str(d) for d in latest_dates]}")

# Pre-load master simulator once to compute all O(N) indicators on the updated dataset
dt0 = datetime.combine(all_dates[0], datetime.min.time())
master = Simulator('replay', dt0, strategy='v6_profit_max', starting_capital=50000.0)
master.load(raw, universe, idx)

# 1. Run on the latest sessions from Yahoo Finance
for sess_date in latest_dates:
    dt = datetime.combine(sess_date, datetime.min.time())
    n_bars = len(raw[sym0][raw[sym0].index.date == sess_date])
    print("\n" + "=" * 95)
    print(f"SESSION AUDIT: {sess_date} ({n_bars} 5m bars from Yahoo Finance)")
    print("=" * 95)
    for cap in [50000.0, 100000.0, 500000.0]:
        sim_base = Simulator('replay', dt, strategy='v6_baseline', starting_capital=cap)
        sim_base.syms = master.syms; sim_base.indices = master.indices
        sim_base.nifty_pos = master.nifty_pos; sim_base.nifty_e10 = master.nifty_e10
        sim_base.nifty_e21 = master.nifty_e21; sim_base.nifty_e30 = master.nifty_e30
        for t in sim_base.session_bar_times():
            sim_base.step(t)
        kb = sim_base.snapshot()['kpi']

        sim_opt = Simulator('replay', dt, strategy='v6_profit_max', starting_capital=cap)
        sim_opt.syms = master.syms; sim_opt.indices = master.indices
        sim_opt.nifty_pos = master.nifty_pos; sim_opt.nifty_e10 = master.nifty_e10
        sim_opt.nifty_e21 = master.nifty_e21; sim_opt.nifty_e30 = master.nifty_e30
        for t in sim_opt.session_bar_times():
            sim_opt.step(t)
        ko = sim_opt.snapshot()['kpi']

        if cap == 50000.0:
            write_journal(sim_opt)

        wrb = f"{kb['win_rate']:.1f}%" if kb['trades'] else "0.0%"
        wro = f"{ko['win_rate']:.1f}%" if ko['trades'] else "0.0%"
        print(f"Cap Rs {cap:>7,.0f} | BASELINE: {kb['trades']:>2} trd ({kb['wins']:>2} W, {wrb:>5}) Gross Rs {kb['gross']:>+9,.2f} Fees Rs {kb['fees']:>7,.2f} Net Rs {kb['realized']:>+9,.2f}")
        print(f"Cap Rs {cap:>7,.0f} | UPGRADED: {ko['trades']:>2} trd ({ko['wins']:>2} W, {wro:>5}) Gross Rs {ko['gross']:>+9,.2f} Fees Rs {ko['fees']:>7,.2f} Net Rs {ko['realized']:>+9,.2f}")

        if cap == 50000.0 and sim_opt.closed:
            print("  Trades Executed (Rs 50k):")
            for tr in sim_opt.closed:
                print(f"    -> {tr['side']:<4} {tr['symbol']:<11} {tr['qty']:>4} sh | Entry Rs {tr['entry']:<7.2f} ({tr['entry_time'][-8:-3]}) -> Exit Rs {tr['exit']:<7.2f} ({tr['exit_time'][-8:-3]}) | {tr['ret_pct']:+.2f}% | Net Rs {tr['net']:>+7.2f} | {tr['reason']}")

# 2. Run Full Dataset Summary (All trading days in cache) across Rs 50k, Rs 100k, Rs 500k
print("\n" + "=" * 95)
print(f"FULL DATASET BENCHMARK ({len(all_dates)} TRADING DAYS: {all_dates[0]} to {all_dates[-1]})")
print("=" * 95)

for cap in [50000.0, 100000.0, 500000.0]:
    for strat_key, strat_label in [('v6_baseline', 'Original V6 Baseline'), ('v6_profit_max', 'Upgraded V6++ Profit Max')]:
        all_closed = []
        daily_nets = []
        for d in all_dates:
            dt = datetime.combine(d, datetime.min.time())
            sim = Simulator('replay', dt, strategy=strat_key, starting_capital=cap)
            sim.syms = master.syms; sim.indices = master.indices
            sim.nifty_pos = master.nifty_pos; sim.nifty_e10 = master.nifty_e10
            sim.nifty_e21 = master.nifty_e21; sim.nifty_e30 = master.nifty_e30
            for t in sim.session_bar_times():
                sim.step(t)
            all_closed.extend(sim.closed)
            daily_nets.append(sim.cash_realized)

        df = pd.DataFrame(all_closed)
        n = len(df)
        wins = int((df['net'] > 0).sum()) if n else 0
        wr = round(wins / n * 100, 1) if n else 0.0
        gross = round(float(df['gross'].sum()), 2) if n else 0.0
        fees = round(float(df['costs'].sum()), 2) if n else 0.0
        net = round(float(df['net'].sum()), 2) if n else 0.0
        worst_trd = round(float(df['net'].min()), 2) if n else 0.0
        cum = np.cumsum(daily_nets)
        peak = np.maximum.accumulate(np.maximum(0.0, cum))
        max_dd = round(float(np.max(peak - cum)), 2) if len(cum) else 0.0
        roi = round(net / cap * 100, 2)
        print(f"Cap Rs {cap:>7,.0f} | {strat_label:<24} | Trades: {n:>3} | Wins: {wins:>3} ({wr:>4.1f}%) | Gross: Rs {gross:>+10,.2f} | Fees: Rs {fees:>9,.2f} | Net: Rs {net:>+10,.2f} ({roi:>+6.2f}%) | WorstTrd: Rs {worst_trd:>+8,.2f} | MaxDD: Rs {max_dd:>8,.2f}")
