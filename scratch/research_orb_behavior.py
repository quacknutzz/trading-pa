"""
scratch/research_orb_behavior.py
Empirical analysis of Opening Range Breakout (09:15 to 09:30 High/Low - ORH/ORL)
across 139 NSE F&O stocks over 61 days (~8,340 sessions).
"""
import sys, os, pickle
import numpy as np
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)

raw = payload['raw']
syms = sorted(list(raw.keys()))
dates = sorted(list(set(raw[syms[0]].index.date)))

total_sessions = 0
orb_break_high = 0
orb_break_low = 0
both_broken = 0

orb_high_expansions = []
orb_low_expansions = []

retest_held_count = 0
retest_tested_count = 0

for s in syms:
    df = raw[s]
    for d in dates:
        day_df = df.loc[df.index.date == d]
        if len(day_df) < 15:
            continue
        
        # 09:15 to 09:30 (three 5m bars: 09:15, 09:20, 09:25)
        orb_bars = day_df.between_time('09:15', '09:25')
        rest_bars = day_df.between_time('09:30', '15:25')
        
        if len(orb_bars) < 3 or len(rest_bars) < 5:
            continue
            
        total_sessions += 1
        orh = float(orb_bars['high'].max())
        orl = float(orb_bars['low'].min())
        orb_rng = orh - orl
        
        rest_high = float(rest_bars['high'].max())
        rest_low = float(rest_bars['low'].min())
        
        b_high = rest_high > orh
        b_low = rest_low < orl
        
        if b_high:
            orb_break_high += 1
            exp_pct = (rest_high / orh - 1.0) * 100.0
            orb_high_expansions.append(exp_pct)
            
            # Check retest of ORH: does price pull back to within 0.1% of ORH and hold?
            # Find first bar that breaks ORH
            idx_break = np.where(rest_bars['high'].values > orh)[0][0]
            after_break_lows = rest_bars['low'].values[idx_break:]
            min_after_break = np.min(after_break_lows)
            if min_after_break <= orh * 1.002: # touched near ORH
                retest_tested_count += 1
                if min_after_break >= orh * 0.998: # held within 0.2%
                    retest_held_count += 1
                    
        if b_low:
            orb_break_low += 1
            exp_pct = (1.0 - rest_low / orl) * 100.0
            orb_low_expansions.append(exp_pct)
            
        if b_high and b_low:
            both_broken += 1

print("==========================================================================")
print(f"OPENING RANGE BREAKOUT (ORH / ORL) EMPIRICAL STUDY (N={total_sessions:,} sessions)")
print("==========================================================================")
print(f"Breaks Above ORH (09:15-09:30 High): {orb_break_high:,} / {total_sessions:,} ({orb_break_high/total_sessions*100:.1f}%)")
print(f"Breaks Below ORL (09:15-09:30 Low):  {orb_break_low:,} / {total_sessions:,} ({orb_break_low/total_sessions*100:.1f}%)")
print(f"Both Boundaries Broken (Chop / Expansion): {both_broken:,} / {total_sessions:,} ({both_broken/total_sessions*100:.1f}%)")
print(f"One-Way Expansion (Clean Trend Day): {total_sessions - both_broken:,} / {total_sessions:,} ({(total_sessions - both_broken)/total_sessions*100:.1f}%)")

avg_hi_exp = np.mean(orb_high_expansions)
med_hi_exp = np.median(orb_high_expansions)
avg_lo_exp = np.mean(orb_low_expansions)
med_lo_exp = np.median(orb_low_expansions)

print(f"\nAverage Upward Excursion beyond ORH:   {avg_hi_exp:.2f}% (Median: {med_hi_exp:.2f}%)")
print(f"Average Downward Excursion beyond ORL: {avg_lo_exp:.2f}% (Median: {med_lo_exp:.2f}%)")

print(f"\nORH Retest & Role Reversal (Resistance becomes Support):")
print(f"Sessions testing ORH pullback after breakout: {retest_tested_count:,}")
print(f"Sessions where ORH held as support (within 0.2%): {retest_held_count:,} ({retest_held_count/retest_tested_count*100:.1f}% of retests)")
