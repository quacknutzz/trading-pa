import sys, pickle
from datetime import datetime
import pandas as pd
sys.path.insert(0, 'sim')
import datafeed
from simulator import Simulator

u = datafeed.load_universe()
with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)
raw = payload['raw']
idx = payload['idx']
ndf = idx['NIFTY 50']

days = sorted({ts.normalize() for ts in ndf.index})[3:] # skip first 3 warmups
print(f'Evaluating {len(days)} trading days...')

results = []
for d in days:
    day_df = ndf[ndf.index.normalize() == d]
    if len(day_df) < 5:
        continue
    o = day_df['open'].iloc[0]
    c = day_df['close'].iloc[-1]
    nifty_ret = (c - o) / o * 100
    
    if nifty_ret > 0.25:
        day_type = 'GREEN'
    elif nifty_ret < -0.25:
        day_type = 'RED'
    else:
        day_type = 'FLAT'
        
    dt = datetime.strptime(str(d.date()), '%Y-%m-%d')
    sim = Simulator('replay', dt, strategy='v6_profit_max', starting_capital=50000.0)
    sim.load(raw, u, idx)
    for t in sim.session_bar_times():
        sim.step(t)
    snap = sim.snapshot()
    k = snap['kpi']
    
    results.append({
        'date': str(d.date()),
        'day_type': day_type,
        'nifty_ret': nifty_ret,
        'trades': k['trades'],
        'wins': k['wins'],
        'losses': k['trades'] - k['wins'],
        'gross': k['gross'],
        'fees': k['fees'],
        'net': k['realized'],
    })

df_res = pd.DataFrame(results)
summary = df_res.groupby('day_type').agg({
    'trades': 'sum',
    'wins': 'sum',
    'losses': 'sum',
    'gross': 'sum',
    'fees': 'sum',
    'net': 'sum'
})
summary['win_rate'] = (summary['wins'] / summary['trades'] * 100).round(1)
summary['profit_per_trade'] = (summary['net'] / summary['trades']).round(2)
print('\n=== V6_PROFIT_MAX PERFORMANCE BY NIFTY REGIME (ALL 58 DAYS) ===')
print(summary[['trades', 'wins', 'losses', 'win_rate', 'gross', 'fees', 'net', 'profit_per_trade']])

total_trades = df_res['trades'].sum()
total_wins = df_res['wins'].sum()
total_net = df_res['net'].sum()
total_gross = df_res['gross'].sum()
total_fees = df_res['fees'].sum()
wr = (total_wins / total_trades * 100) if total_trades else 0
print(f'\nTotal Trades: {total_trades}, Wins: {total_wins} ({wr:.1f}%), Gross: Rs {total_gross:+,.2f}, Fees: Rs {total_fees:,.2f}, Net: Rs {total_net:+,.2f}')

df_res.to_csv('scratch/nifty_regime_v6_baseline.csv', index=False)
print('Saved to scratch/nifty_regime_v6_baseline.csv')
