import sys, pickle, os
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')
SIM_DIR = 'sim'
sys.path.insert(0, SIM_DIR)
import config as C
import datafeed
from simulator import Simulator

u = datafeed.load_universe()
with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)
raw = payload['raw']
idx = payload['idx']

CAPITAL = 50000.0
WEEK_DATES = ['2026-10-26', '2026-10-27', '2026-10-28', '2026-10-29', '2026-10-30']

results = {'v6_profit_max': {'trades': 0, 'wins': 0, 'gross': 0.0, 'fees': 0.0, 'net': 0.0, 'days': {}},
           'v6_baseline':   {'trades': 0, 'wins': 0, 'gross': 0.0, 'fees': 0.0, 'net': 0.0, 'days': {}}}

for d_str in WEEK_DATES:
    dt = datetime.strptime(d_str, '%Y-%m-%d')
    for strat in ['v6_baseline', 'v6_profit_max']:
        sim = Simulator('replay', dt, strategy=strat, starting_capital=CAPITAL)
        sim.load(raw, u, idx)
        times = sim.session_bar_times()
        for t in times:
            sim.step(t)
        
        snap = sim.snapshot()
        kpi = snap['kpi']
        wins = kpi['wins']
        trades = kpi['trades']
        gross = kpi['gross']
        fees = kpi['fees']
        net = kpi['realized']
        
        results[strat]['trades'] += trades
        results[strat]['wins'] += wins
        results[strat]['gross'] += gross
        results[strat]['fees'] += fees
        results[strat]['net'] += net
        results[strat]['days'][d_str] = {
            'trades': trades, 'wins': wins,
            'wr': round(wins / trades * 100, 1) if trades else 0.0,
            'gross': gross, 'fees': fees, 'net': net
        }

print('=== WEEK OF OCT 26 - OCT 30, 2026 (CAPITAL: Rs. 50,000) ===\n')

for strat, label in [('v6_baseline', 'Original Baseline'), ('v6_profit_max', 'Latest Strategy')]:
    r = results[strat]
    wr = round(r['wins'] / r['trades'] * 100, 1) if r['trades'] else 0.0
    print(f'>>> {label.upper()}')
    print(f"  Total Trades:     {r['trades']}")
    print(f"  Wins / Losses:    {r['wins']} Wins / {r['trades'] - r['wins']} Losses")
    print(f"  Win Rate:         {wr}%")
    print(f"  Gross P&L:        Rs. {r['gross']:+,.2f}")
    print(f"  Taxes & Fees:     Rs. {r['fees']:,.2f}")
    print(f"  Net Realized P&L: Rs. {r['net']:+,.2f}")
    print('  Day-by-Day:')
    for d_str in WEEK_DATES:
        d = r['days'][d_str]
        print(f"    {d_str}: {d['trades']:2d} trades ({d['wins']:2d}W, {d['wr']:4.1f}% WR) | Gross: Rs. {d['gross']:+7.2f} | Fees: Rs. {d['fees']:6.2f} | Net: Rs. {d['net']:+7.2f}")
    print()

net_opt = results['v6_profit_max']['net']
net_base = results['v6_baseline']['net']
alpha = net_opt - net_base
tax_saved = results['v6_baseline']['fees'] - results['v6_profit_max']['fees']

print('>>> HEAD-TO-HEAD SUMMARY (OCT 26 - OCT 30):')
print(f"  Net Profit Difference (Alpha): Rs. {alpha:+,.2f}")
print(f"  Taxes/Fees Avoided:            Rs. {tax_saved:,.2f}")
