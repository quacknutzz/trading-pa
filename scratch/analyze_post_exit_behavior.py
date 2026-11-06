import pickle, math
import pandas as pd
import numpy as np

with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)
raw = payload['raw']

df = pd.read_csv('scratch/dad_cockpit_v6_raw.csv')
closed = df[df['Outcome'].isin(['WIN', 'LOSS'])].copy()

def clean_reason(c):
    c = str(c)
    if 'ADX_5M_BELOW_18' in c: return 'ADX < 18'
    if 'EMA_CROSS_5M' in c: return '5M_CROSS_REVERSE'
    if 'EMA_CROSS_10M_CONTRA' in c or 'REVERSE_10M_CROSS' in c: return '10M_CROSS_CONTRA'
    return None

closed['Exit_Type'] = closed['Comments'].apply(clean_reason)
target_trades = closed[closed['Exit_Type'].notna()].copy()

print(f"Total indicator exit trades to evaluate: {len(target_trades)}")

post_exit_results = []
for idx, row in target_trades.iterrows():
    sym = str(row['Ticker']).strip()
    if not sym.endswith('.NS'):
        sym_key = sym + '.NS'
    else:
        sym_key = sym
    action = str(row['Action']).strip()
    ent = float(row['Ent']) if pd.notna(row['Ent']) else None
    cp = float(row['CP']) if pd.notna(row['CP']) else None
    close_time = str(row['Close_Time'])
    
    if sym_key not in raw or ent is None or cp is None:
        continue
    
    sdf = raw[sym_key]
    try:
        dt = pd.to_datetime(close_time)
        bar_idx = sdf.index.get_indexer([dt], method='nearest')[0]
        actual_bar_time = sdf.index[bar_idx]
        if abs((actual_bar_time - dt).total_seconds()) > 600:
            continue
        
        fut6 = sdf.iloc[bar_idx+1:bar_idx+7]
        fut12 = sdf.iloc[bar_idx+1:bar_idx+13]
        
        if len(fut6) > 0:
            sgn = 1 if action == 'BUY' else -1
            best_post_6 = (fut6['high'].max() - cp)/cp if sgn == 1 else (cp - fut6['low'].min())/cp
            worst_post_6 = (fut6['low'].min() - cp)/cp if sgn == 1 else (cp - fut6['high'].max())/cp
            
            tgt_px = ent * (1 + 0.012 * sgn)
            sl_px = ent * (1 - 0.010 * sgn)
            
            day_mask = fut12.index.date == dt.date()
            fut_day = fut12[day_mask]
            
            hit_tgt = False
            hit_sl = False
            if len(fut_day) > 0:
                if sgn == 1:
                    hit_tgt = bool((fut_day['high'] >= tgt_px).any())
                    hit_sl = bool((fut_day['low'] <= sl_px).any())
                else:
                    hit_tgt = bool((fut_day['low'] <= tgt_px).any())
                    hit_sl = bool((fut_day['high'] >= sl_px).any())
            
            post_exit_results.append({
                'Ticker': sym,
                'Action': action,
                'Exit_Type': row['Exit_Type'],
                'Exit_Px': cp,
                'Best_Post_6_Pct': best_post_6 * 100,
                'Worst_Post_6_Pct': worst_post_6 * 100,
                'Hit_Target_Later': hit_tgt,
                'Hit_SL_Later': hit_sl
            })
    except Exception as e:
        pass

p_df = pd.DataFrame(post_exit_results)
print(f"Matched {len(p_df)} trades with historical 5m bars.")
if len(p_df) > 0:
    for et, g in p_df.groupby('Exit_Type'):
        print(f"\n--- Exit Type: {et} (N={len(g)}) ---")
        print(f"Avg Favorable move in next 30m: +{g['Best_Post_6_Pct'].mean():.2f}%")
        print(f"Avg Adverse move in next 30m: {g['Worst_Post_6_Pct'].mean():.2f}%")
        print(f"Later reached Target (1.2%): {g['Hit_Target_Later'].sum()} trades ({g['Hit_Target_Later'].mean()*100:.1f}%)")
        print(f"Later reached Stop-Loss (-1.0%): {g['Hit_SL_Later'].sum()} trades ({g['Hit_SL_Later'].mean()*100:.1f}%)")
