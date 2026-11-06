import sys, json, os, pickle
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'sim')
import datafeed
from simulator import Simulator

u = datafeed.load_universe()
with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)
raw = payload['raw']
idx = payload['idx']

dates = ['2026-08-10', '2026-08-11', '2026-08-12', '2026-08-13', '2026-08-14']

print('=== TESTING WEEK 01 (AUG 10 - AUG 14, 2026): Rs. 50k vs Rs. 500k ===\n')

for cap in [50000.0, 500000.0]:
    tot_trades, tot_wins = 0, 0
    tot_gross, tot_fees, tot_net = 0.0, 0.0, 0.0
    for d_str in dates:
        dt = datetime.strptime(d_str, '%Y-%m-%d')
        sim = Simulator('replay', dt, strategy='v6_profit_max', starting_capital=cap)
        sim.load(raw, u, idx)
        for t in sim.session_bar_times():
            sim.step(t)
        k = sim.snapshot()['kpi']
        tot_trades += k['trades']
        tot_wins += k['wins']
        tot_gross += k['gross']
        tot_fees += k['fees']
        tot_net += k['realized']
        
    pct_return = (tot_net / cap) * 100
    fee_pct_cap = (tot_fees / cap) * 100
    fee_pct_gross = (tot_fees / tot_gross * 100) if tot_gross else 0
    wr = round(tot_wins / tot_trades * 100, 1) if tot_trades else 0.0
    print(f"Capital: Rs. {cap:,.0f}")
    print(f"  Trades:        {tot_trades} ({tot_wins} Wins, {wr}% WR)")
    print(f"  Gross P&L:     Rs. {tot_gross:+,.2f}")
    print(f"  Taxes/Fees:    Rs. {tot_fees:,.2f} ({fee_pct_cap:.3f}% of capital, {fee_pct_gross:.1f}% of gross)")
    print(f"  Net Realized:  Rs. {tot_net:+,.2f}  ===>  PERCENT RETURN: {pct_return:+.3f}%\n")
