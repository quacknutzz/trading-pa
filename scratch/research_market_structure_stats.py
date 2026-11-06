"""
scratch/research_market_structure_stats.py
Empirical analysis of dynamic resistance/support levels and statistical exhaustion bands
across 139 NSE F&O stocks over 60 trading days (~500,000 5-minute bars).
"""
import sys, os, pickle, math
from datetime import datetime
import numpy as np
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)

raw = payload['raw']
syms = sorted(list(raw.keys()))
all_dates = sorted(list(set(raw[syms[0]].index.date)))

print(f"Loaded {len(syms)} symbols across {len(all_dates)} dates: {all_dates[0]} to {all_dates[-1]}")

# -------------------------------------------------------------------------
# Helper functions for calculations
# -------------------------------------------------------------------------
def get_daily_ohlc(df):
    """Aggregate 5m bars into daily OHLC."""
    df_daily = df.groupby(df.index.date).agg({
        'open': 'first',
        'high': 'max',
        'low': 'min',
        'close': 'last',
        'volume': 'sum'
    })
    return df_daily

# Store daily aggregates
daily_data = {s: get_daily_ohlc(raw[s]) for s in syms}

# Let's inspect touch rates across all stock-days (day index >= 1)
total_sessions = 0
touches_R1 = 0
touches_R2 = 0
touches_PDH = 0
touches_H3 = 0
touches_H4 = 0
touches_VWAP_2S = 0

touches_S1 = 0
touches_S2 = 0
touches_PDL = 0
touches_L3 = 0
touches_L4 = 0
touches_VWAP_M2S = 0

# Behavior at levels: R1, R2, VWAP_2S, H3, H4, PDH
# For R1:
r1_tagged_count = 0
r1_reversal_count = 0   # pulls back by >= 0.5% or closes below touch
r1_continue_to_r2 = 0   # reaches R2 after touching R1
r1_session_peak = 0     # session high is within 0.15% of R1

# For VWAP + 2 Sigma:
vwap2s_tagged_count = 0
vwap2s_session_peak = 0 # within 0.15% of session peak
vwap2s_pullbacks = []   # % drawdown from touch to session close or 5 bars later

# For Camarilla H3:
h3_tagged_count = 0
h3_reversal_count = 0
h3_continue_h4 = 0
h3_session_peak = 0

# For Camarilla H4:
h4_tagged_count = 0
h4_session_peak = 0

# For PDH:
pdh_tagged_count = 0
pdh_reversal_count = 0
pdh_session_peak = 0

# Momentum subset (sessions where stock had strong upward move: day open to high >= 1.0% or 10m cross)
mom_sessions = 0
mom_touches_R1 = 0
mom_touches_R2 = 0
mom_touches_PDH = 0
mom_touches_H3 = 0
mom_touches_H4 = 0
mom_touches_VWAP_2S = 0

for s in syms:
    df = raw[s]
    d_df = daily_data[s]
    dates = [d for d in all_dates if d in d_df.index]
    if len(dates) < 2:
        continue

    for i in range(1, len(dates)):
        prev_d = dates[i-1]
        curr_d = dates[i]

        prev_row = d_df.loc[prev_d]
        curr_row = d_df.loc[curr_d]

        pdh = float(prev_row['high'])
        pdl = float(prev_row['low'])
        pdc = float(prev_row['close'])
        rng = pdh - pdl
        if rng <= 0 or pdc <= 0:
            continue

        # Daily Floor Pivots
        P = (pdh + pdl + pdc) / 3.0
        R1 = 2.0 * P - pdl
        S1 = 2.0 * P - pdh
        R2 = P + rng
        S2 = P - rng

        # Camarilla Pivots
        H4 = pdc + 0.55 * rng
        H3 = pdc + 0.275 * rng
        L3 = pdc - 0.275 * rng
        L4 = pdc - 0.55 * rng

        # Intraday 5m bars
        bars = df.loc[df.index.date == curr_d]
        if len(bars) < 15: # skip truncated days
            continue

        total_sessions += 1
        day_high = float(curr_row['high'])
        day_low = float(curr_row['low'])
        day_open = float(curr_row['open'])
        day_close = float(curr_row['close'])

        # Compute intraday VWAP + 2 Sigma series
        h_arr = bars['high'].values.astype(np.float64)
        l_arr = bars['low'].values.astype(np.float64)
        c_arr = bars['close'].values.astype(np.float64)
        v_arr = bars['volume'].values.astype(np.float64)
        tp = (h_arr + l_arr + c_arr) / 3.0

        cum_v = np.cumsum(v_arr)
        cum_pv = np.cumsum(tp * v_arr)
        cum_pv2 = np.cumsum(tp * tp * v_arr)

        valid_v = cum_v > 0
        vwap = np.zeros_like(tp)
        vwap[valid_v] = cum_pv[valid_v] / cum_v[valid_v]
        var = np.maximum(0.0, (cum_pv2[valid_v] / cum_v[valid_v]) - (vwap[valid_v] ** 2))
        std = np.sqrt(var)

        upper_2s = vwap[valid_v] + 2.0 * std
        lower_2s = vwap[valid_v] - 2.0 * std

        touched_r1 = day_high >= R1
        touched_r2 = day_high >= R2
        touched_pdh = day_high >= pdh
        touched_h3 = day_high >= H3
        touched_h4 = day_high >= H4
        touched_vwap2s = np.any(h_arr[:len(upper_2s)] >= upper_2s)

        touched_s1 = day_low <= S1
        touched_s2 = day_low <= S2
        touched_pdl = day_low <= pdl
        touched_l3 = day_low <= L3
        touched_l4 = day_low <= L4
        touched_vwapm2s = np.any(l_arr[:len(lower_2s)] <= lower_2s)

        if touched_r1: touches_R1 += 1
        if touched_r2: touches_R2 += 1
        if touched_pdh: touches_PDH += 1
        if touched_h3: touches_H3 += 1
        if touched_h4: touches_H4 += 1
        if touched_vwap2s: touches_VWAP_2S += 1

        if touched_s1: touches_S1 += 1
        if touched_s2: touches_S2 += 1
        if touched_pdl: touches_PDL += 1
        if touched_l3: touches_L3 += 1
        if touched_l4: touches_L4 += 1
        if touched_vwapm2s: touches_VWAP_M2S += 1

        # Check momentum condition: day moved up >= 1.0% from open
        if (day_high / day_open - 1.0) >= 0.010:
            mom_sessions += 1
            if touched_r1: mom_touches_R1 += 1
            if touched_r2: mom_touches_R2 += 1
            if touched_pdh: mom_touches_PDH += 1
            if touched_h3: mom_touches_H3 += 1
            if touched_h4: mom_touches_H4 += 1
            if touched_vwap2s: mom_touches_VWAP_2S += 1

        # Analyze behavior at R1
        if touched_r1:
            r1_tagged_count += 1
            # Find first bar where high >= R1
            idx_r1 = np.where(h_arr >= R1)[0][0]
            # Did it mark the session peak? (Within 0.15% of day high)
            if (day_high - R1) / R1 <= 0.0015:
                r1_session_peak += 1
            # Did it continue to R2?
            if touched_r2:
                idx_r2 = np.where(h_arr >= R2)[0][0]
                if idx_r2 >= idx_r1:
                    r1_continue_to_r2 += 1
            # Subsequent pullback from R1 touch:
            subsequent_low = np.min(l_arr[idx_r1:])
            if (R1 - subsequent_low) / R1 >= 0.005: # pulled back >= 0.5%
                r1_reversal_count += 1

        # Analyze behavior at VWAP + 2.0 Sigma
        if touched_vwap2s:
            vwap2s_tagged_count += 1
            idx_v2s = np.where(h_arr[:len(upper_2s)] >= upper_2s)[0][0]
            touch_px = h_arr[idx_v2s]
            # Is this bar or subsequent within 0.15% of day high?
            if (day_high - touch_px) / touch_px <= 0.0015:
                vwap2s_session_peak += 1
            # Drawdown from peak after touch:
            subsequent_low = np.min(l_arr[idx_v2s:])
            dd = (touch_px - subsequent_low) / touch_px * 100.0
            vwap2s_pullbacks.append(dd)

        # Analyze Camarilla H3
        if touched_h3:
            h3_tagged_count += 1
            idx_h3 = np.where(h_arr >= H3)[0][0]
            if (day_high - H3) / H3 <= 0.0015:
                h3_session_peak += 1
            if touched_h4:
                idx_h4 = np.where(h_arr >= H4)[0][0]
                if idx_h4 >= idx_h3:
                    h3_continue_h4 += 1
            subsequent_low = np.min(l_arr[idx_h3:])
            if (H3 - subsequent_low) / H3 >= 0.005:
                h3_reversal_count += 1

        # Analyze Camarilla H4
        if touched_h4:
            h4_tagged_count += 1
            if (day_high - H4) / H4 <= 0.0015:
                h4_session_peak += 1

        # Analyze PDH
        if touched_pdh:
            pdh_tagged_count += 1
            idx_pdh = np.where(h_arr >= pdh)[0][0]
            if (day_high - pdh) / pdh <= 0.0015:
                pdh_session_peak += 1
            subsequent_low = np.min(l_arr[idx_pdh:])
            if (pdh - subsequent_low) / pdh >= 0.005:
                pdh_reversal_count += 1

print("\n==========================================================================")
print(f"EMPIRICAL MARKET STRUCTURE STATISTICS: NSE F&O (139 Stocks x 60 Days)")
print(f"Total Stock-Day Sessions Analyzed: {total_sessions:,}")
print(f"Momentum Sessions (Intraday High >= +1.0% from Open): {mom_sessions:,} ({mom_sessions/total_sessions*100:.1f}%)")
print("==========================================================================")

print("\n--- 1. OVERALL POPULATION TOUCH RATES (All Sessions) ---")
print(f"Resistance Levels (Bullish Excursion):")
print(f"  Floor Pivot R1:        {touches_R1:>6,} / {total_sessions} ({touches_R1/total_sessions*100:.1f}%)")
print(f"  Floor Pivot R2:        {touches_R2:>6,} / {total_sessions} ({touches_R2/total_sessions*100:.1f}%)")
print(f"  Previous Day High (PDH):{touches_PDH:>6,} / {total_sessions} ({touches_PDH/total_sessions*100:.1f}%)")
print(f"  Camarilla H3 (Range Tgt):{touches_H3:>6,} / {total_sessions} ({touches_H3/total_sessions*100:.1f}%)")
print(f"  Camarilla H4 (Breakout):{touches_H4:>6,} / {total_sessions} ({touches_H4/total_sessions*100:.1f}%)")
print(f"  VWAP + 2.0 Sigma:      {touches_VWAP_2S:>6,} / {total_sessions} ({touches_VWAP_2S/total_sessions*100:.1f}%)")

print(f"\nSupport Levels (Bearish Excursion):")
print(f"  Floor Pivot S1:        {touches_S1:>6,} / {total_sessions} ({touches_S1/total_sessions*100:.1f}%)")
print(f"  Floor Pivot S2:        {touches_S2:>6,} / {total_sessions} ({touches_S2/total_sessions*100:.1f}%)")
print(f"  Previous Day Low (PDL):{touches_PDL:>6,} / {total_sessions} ({touches_PDL/total_sessions*100:.1f}%)")
print(f"  Camarilla L3 (Range Tgt):{touches_L3:>6,} / {total_sessions} ({touches_L3/total_sessions*100:.1f}%)")
print(f"  Camarilla L4 (Breakout):{touches_L4:>6,} / {total_sessions} ({touches_L4/total_sessions*100:.1f}%)")
print(f"  VWAP - 2.0 Sigma:      {touches_VWAP_M2S:>6,} / {total_sessions} ({touches_VWAP_M2S/total_sessions*100:.1f}%)")

print("\n--- 2. MOMENTUM RUNNERS REACH RATES (When Intraday High >= +1.0%) ---")
print(f"  Floor Pivot R1:        {mom_touches_R1:>6,} / {mom_sessions} ({mom_touches_R1/mom_sessions*100:.1f}%)")
print(f"  Camarilla H3:          {mom_touches_H3:>6,} / {mom_sessions} ({mom_touches_H3/mom_sessions*100:.1f}%)")
print(f"  Previous Day High:     {mom_touches_PDH:>6,} / {mom_sessions} ({mom_touches_PDH/mom_sessions*100:.1f}%)")
print(f"  VWAP + 2.0 Sigma:      {mom_touches_VWAP_2S:>6,} / {mom_sessions} ({mom_touches_VWAP_2S/mom_sessions*100:.1f}%)")
print(f"  Floor Pivot R2:        {mom_touches_R2:>6,} / {mom_sessions} ({mom_touches_R2/mom_sessions*100:.1f}%)")
print(f"  Camarilla H4:          {mom_touches_H4:>6,} / {mom_sessions} ({mom_touches_H4/mom_sessions*100:.1f}%)")

print("\n--- 3. REVERSAL VS CONTINUATION AT STRUCTURAL RESISTANCE ---")
print(f"When price tags Floor Pivot R1 (N={r1_tagged_count}):")
print(f"  * Marks Session Peak (within 0.15%):     {r1_session_peak:>5} ({r1_session_peak/r1_tagged_count*100:.1f}%)")
print(f"  * Rejection / Pullback >= 0.5% after tag: {r1_reversal_count:>5} ({r1_reversal_count/r1_tagged_count*100:.1f}%)")
print(f"  * Breaks through & continues to R2:      {r1_continue_to_r2:>5} ({r1_continue_to_r2/r1_tagged_count*100:.1f}%)")

print(f"\nWhen price tags Camarilla H3 (N={h3_tagged_count}):")
print(f"  * Marks Session Peak (within 0.15%):     {h3_session_peak:>5} ({h3_session_peak/h3_tagged_count*100:.1f}%)")
print(f"  * Rejection / Pullback >= 0.5% after tag: {h3_reversal_count:>5} ({h3_reversal_count/h3_tagged_count*100:.1f}%)")
print(f"  * Breaks through & continues to H4:      {h3_continue_h4:>5} ({h3_continue_h4/h3_tagged_count*100:.1f}%)")

print(f"\nWhen price tags Previous Day High (PDH) (N={pdh_tagged_count}):")
print(f"  * Marks Session Peak (within 0.15%):     {pdh_session_peak:>5} ({pdh_session_peak/pdh_tagged_count*100:.1f}%)")
print(f"  * Rejection / Pullback >= 0.5% after tag: {pdh_reversal_count:>5} ({pdh_reversal_count/pdh_tagged_count*100:.1f}%)")

print(f"\nWhen price tags VWAP + 2.0 Sigma (N={vwap2s_tagged_count}):")
print(f"  * Marks Session Peak (within 0.15%):     {vwap2s_session_peak:>5} ({vwap2s_session_peak/vwap2s_tagged_count*100:.1f}%)")
avg_dd = np.mean(vwap2s_pullbacks) if vwap2s_pullbacks else 0
med_dd = np.median(vwap2s_pullbacks) if vwap2s_pullbacks else 0
print(f"  * Average subsequent drawdown from touch: {avg_dd:.2f}% (Median: {med_dd:.2f}%)")
