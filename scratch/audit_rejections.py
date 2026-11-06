import sys, pickle
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

WEEK_DATES = ['2026-10-26', '2026-10-27', '2026-10-28', '2026-10-29', '2026-10-30']

# Let's inspect what signals were filtered out in Latest Strategy during that week
print('=== SIGNAL AUDIT: WHY DID IT REJECT TRADES ON OCT 26-28? ===')
rejection_reasons = {}

for d_str in WEEK_DATES[:3]: # Oct 26, 27, 28
    dt = datetime.strptime(d_str, '%Y-%m-%d')
    sim = Simulator('replay', dt, strategy='v6_profit_max', starting_capital=50000.0)
    sim.load(raw, u, idx)
    for t in sim.session_bar_times():
        sim.step(t)
    
    for sig in sim.signals:
        res = sig.get('result')
        det = sig.get('detail', '')
        if res not in ['ORDER_PLACED', 'FILLED']:
            rejection_reasons[det] = rejection_reasons.get(det, 0) + 1

# Sort top rejection reasons
sorted_rejections = sorted(rejection_reasons.items(), key=lambda x: x[1], reverse=True)
for reason, count in sorted_rejections[:10]:
    print(f"  {reason[:60]:60s} : {count} times")
