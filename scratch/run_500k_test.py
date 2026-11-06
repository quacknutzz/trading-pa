import pickle, pandas as pd
import numpy as np
from datetime import datetime, timedelta
import sys, time, os
from concurrent.futures import ProcessPoolExecutor

sys.path.insert(0, '.')
sys.path.insert(0, 'sim')
from scratch.test_market_chop_gate import TrueDecoupledLeaderSimulator

G_RAW = None
G_IDX = None
G_UNIVERSE = None
G_TOD_MEAN = None

def worker_init():
    global G_RAW, G_IDX, G_UNIVERSE, G_TOD_MEAN
    with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
        payload = pickle.load(f)
    G_RAW = payload['raw']
    G_IDX = payload['idx']
    G_UNIVERSE = payload['universe']
    
    G_TOD_MEAN = {}
    for y_sym, df in G_RAW.items():
        sym = y_sym.replace('.NS', '')
        if len(df) < 200:
            continue
        df_copy = df.copy()
        df_copy['tod'] = df_copy.index.time
        df_copy['date'] = df_copy.index.normalize()
        pv = df_copy.pivot(index='date', columns='tod', values='volume')
        pv_mean20 = pv.shift(1).rolling(20, min_periods=5).mean()
        for date_idx, row in pv_mean20.iterrows():
            d_val = date_idx.date()
            for tod_val, mean_vol in row.items():
                if not np.isnan(mean_vol) and mean_vol > 0:
                    G_TOD_MEAN[(sym, d_val, tod_val)] = float(mean_vol)

def run_day_500k(args):
    d_str, day_cat, nifty_ret = args
    dt = datetime.strptime(d_str, '%Y-%m-%d')
    sim = TrueDecoupledLeaderSimulator('replay', dt, starting_capital=500000.0)
    sim.load(G_RAW, G_UNIVERSE, G_IDX)
    for t in sim.session_bar_times():
        sim.step(t)
    snap = sim.snapshot()
    k = snap['kpi']
    return {
        'date': d_str,
        'day_cat': day_cat,
        'nifty_ret': nifty_ret,
        'trades': k['trades'],
        'wins': k['wins'],
        'losses': k['trades'] - k['wins'],
        'gross': k['gross'],
        'fees': k['fees'],
        'net': k['realized']
    }

if __name__ == '__main__':
    t0 = time.time()
    with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
        payload = pickle.load(f)
    ndf = payload['idx']['NIFTY 50']
    days = sorted({ts.normalize() for ts in ndf.index})[3:]
    
    tasks = []
    for d in days:
        day_df = ndf[ndf.index.normalize() == d]
        if len(day_df) < 5:
            continue
        o = day_df['open'].iloc[0]
        c = day_df['close'].iloc[-1]
        ret = (c - o) / o * 100
        day_cat = 'GREEN' if ret > 0.25 else ('RED' if ret < -0.25 else 'FLAT')
        tasks.append((str(d.date()), day_cat, ret))
        
    print(f"Running True Decoupled Leader Strategy with Rs 500,000 capital across {len(tasks)} days...")
    with ProcessPoolExecutor(max_workers=8, initializer=worker_init) as executor:
        results = list(executor.map(run_day_500k, tasks))
        
    df_res = pd.DataFrame(results)
    
    tot_trades = df_res['trades'].sum()
    tot_wins = df_res['wins'].sum()
    tot_gross = df_res['gross'].sum()
    tot_fees = df_res['fees'].sum()
    tot_net = df_res['net'].sum()
    wr = (tot_wins / tot_trades * 100) if tot_trades else 0
    
    print("\n" + "="*80)
    print("FINAL RESULTS: TRUE DECOUPLED LEADER ENGINE (500K CAPITAL)")
    print("="*80)
    print(f"Total Trades: {tot_trades} | Wins: {tot_wins} ({wr:.1f}% WR)")
    print(f"Gross Profit: Rs {tot_gross:+,.2f}")
    print(f"Taxes / Fees: Rs {tot_fees:,.2f} ({tot_fees/tot_gross*100:.1f}% of gross)")
    print(f"Net Realized: Rs {tot_net:+,.2f}  ({tot_net/500000*100:+.2f}% on 500k)")
    
    reg_summary = df_res.groupby('day_cat').agg(
        trades=('trades', 'sum'),
        wins=('wins', 'sum'),
        losses=('losses', 'sum'),
        gross=('gross', 'sum'),
        fees=('fees', 'sum'),
        net=('net', 'sum')
    )
    reg_summary['win_rate'] = (reg_summary['wins'] / reg_summary['trades'] * 100).round(1)
    reg_summary['net_per_trade'] = (reg_summary['net'] / reg_summary['trades']).round(2)
    print("\nPerformance by Market Day Type:")
    print(reg_summary[['trades', 'wins', 'losses', 'win_rate', 'gross', 'fees', 'net', 'net_per_trade']])
    df_res.to_csv('scratch/true_decoupled_leader_results_500k.csv', index=False)
