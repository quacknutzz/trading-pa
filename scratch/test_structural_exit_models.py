"""
scratch/test_structural_exit_models.py
Comprehensive backtesting of Market-Structure-Anchored Exit Models vs Baseline.
"""
import sys, os, pickle, math
from datetime import datetime, time as dtime
import numpy as np
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'sim')
import config as C
import datafeed
from simulator import Simulator, trade_costs, breakeven_exit_price, BAR

# Load data and master simulator
u = datafeed.load_universe()
with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)
raw = payload['raw']
idx = payload['idx']

sym0 = list(raw.keys())[0]
dates = sorted(list(set(raw[sym0].index.date)))

master = Simulator('replay', datetime.combine(dates[0], datetime.min.time()), starting_capital=500000.0)
master.load(raw, u, idx)

print(f"Master loaded with {len(master.syms)} symbols across {len(dates)} dates.")

# Precompute daily pivots
pivot_cache = {}
for s, sd in master.syms.items():
    d_df = sd.df.groupby(sd.df.index.date).agg({
        'open': 'first', 'high': 'max', 'low': 'min', 'close': 'last'
    })
    pivot_cache[s] = {}
    d_list = [d for d in dates if d in d_df.index]
    for i in range(1, len(d_list)):
        prev_d = d_list[i-1]
        curr_d = d_list[i]
        p_row = d_df.loc[prev_d]
        pdh, pdl, pdc = float(p_row['high']), float(p_row['low']), float(p_row['close'])
        rng = pdh - pdl
        P = (pdh + pdl + pdc) / 3.0
        R1 = 2.0 * P - pdl
        S1 = 2.0 * P - pdh
        R2 = P + rng
        S2 = P - rng
        H4 = pdc + 0.55 * rng
        H3 = pdc + 0.275 * rng
        L3 = pdc - 0.275 * rng
        L4 = pdc - 0.55 * rng
        pivot_cache[s][curr_d] = {
            'pdh': pdh, 'pdl': pdl, 'pdc': pdc,
            'P': P, 'R1': R1, 'S1': S1, 'R2': R2, 'S2': S2,
            'H4': H4, 'H3': H3, 'L3': L3, 'L4': L4
        }

# Precompute intraday VWAP Standard Deviation for each symbol
vwap_std_cache = {}
for s, sd in master.syms.items():
    n = len(sd.df)
    dates_arr = sd.df.index.normalize()
    h_arr, l_arr, c_arr, v_arr = sd.h, sd.l, sd.c, sd.v
    tp = (h_arr + l_arr + c_arr) / 3.0
    vwap_std = np.zeros(n, dtype=np.float64)
    
    j = 0
    while j < n:
        d = dates_arr[j]
        k = j
        while k < n and dates_arr[k] == d:
            k += 1
        sub_tp = tp[j:k]
        sub_v = v_arr[j:k]
        cum_v = np.cumsum(sub_v)
        cum_pv = np.cumsum(sub_tp * sub_v)
        cum_pv2 = np.cumsum(sub_tp * sub_tp * sub_v)
        valid = cum_v > 0
        v_vwap = np.zeros(k - j, dtype=np.float64)
        v_vwap[valid] = cum_pv[valid] / cum_v[valid]
        v_var = np.maximum(0.0, (cum_pv2[valid] / cum_v[valid]) - (v_vwap[valid] ** 2))
        v_std = np.zeros(k - j, dtype=np.float64)
        v_std[valid] = np.sqrt(v_var)
        vwap_std[j:k] = v_std
        j = k
    vwap_std_cache[s] = vwap_std

# -------------------------------------------------------------------------
# Subclassed Simulator with Custom Exit Engines
# -------------------------------------------------------------------------
class StructuralSimulator(Simulator):
    def __init__(self, session_date, model_type='baseline', starting_capital=500000.0):
        super().__init__('replay', session_date, strategy='v6_profit_max', starting_capital=starting_capital)
        self.model_type = model_type
        
    def _manage_pending_profit_max(self, t):
        # Identical pending management and filters as baseline,
        # but allocate target according to chosen model!
        for sym in list(self.pending):
            o = self.pending[sym]
            sd, i = self._bar(sym, t)
            if sd is None or t <= o["signal_bar"]:
                continue
            lo, hi, op, cl = float(sd.l[i]), float(sd.h[i]), float(sd.o[i]), float(sd.c[i])
            filled = None
            if o["side"] == "BUY" and lo <= o["limit"]:
                filled = min(o["limit"], op)
            elif o["side"] == "SELL" and hi >= o["limit"]:
                filled = max(o["limit"], op)

            if filled is not None:
                c10, _ = sd.tenmin_state(i)
                c5 = sd.cross5m(i)
                if (o["side"] == "BUY" and (c10 == "BEARISH" or c5 == "BEARISH" or float(sd.c10_diff[i]) <= 0)) or \
                   (o["side"] == "SELL" and (c10 == "BULLISH" or c5 == "BULLISH" or float(sd.c10_diff[i]) >= 0)):
                    del self.pending[sym]
                    continue

                vwap_val = float(sd.vwap[i])
                e30 = float(sd.ema[30][i])
                if o["side"] == "BUY":
                    if cl < vwap_val or cl < op or cl < filled * 0.9985 or cl < e30:
                        del self.pending[sym]
                        continue
                else:
                    if cl > vwap_val or cl > op or cl > filled * 1.0015 or cl > e30:
                        del self.pending[sym]
                        continue

                del self.pending[sym]
                curr_d = t.date()
                pivots = pivot_cache.get(sym, {}).get(curr_d)
                
                # Default initial SL (0.85%)
                sl_pct = 0.0085
                sl_px = round(filled * (1 - sl_pct) if o["side"] == "BUY" else filled * (1 + sl_pct), 2)
                
                # Determine target based on model_type
                if self.model_type == 'baseline':
                    atr_val = float(sd.atr[i]) if not math.isnan(sd.atr[i]) else (filled * 0.0025)
                    atr_pct = atr_val / filled if filled > 0 else 0.0025
                    tgt_pct = min(max(atr_pct * 2.8, 0.0060), 0.0095)
                    tgt_px = round(filled * (1 + tgt_pct) if o["side"] == "BUY" else filled * (1 - tgt_pct), 2)
                
                elif self.model_type == 'floor_pivots':
                    # Floor Pivots R1 / R2
                    if pivots is not None:
                        if o["side"] == "BUY":
                            r1, r2 = pivots['R1'], pivots['R2']
                            dist_r1 = (r1 - filled) / filled
                            tgt_px = round(r1, 2) if dist_r1 >= 0.0065 else round(r2, 2)
                        else:
                            s1, s2 = pivots['S1'], pivots['S2']
                            dist_s1 = (filled - s1) / filled
                            tgt_px = round(s1, 2) if dist_s1 >= 0.0065 else round(s2, 2)
                    else:
                        tgt_px = round(filled * 1.012 if o["side"] == "BUY" else filled * 0.988, 2)
                    tgt_pct = abs(tgt_px / filled - 1.0)
                    
                elif self.model_type == 'camarilla':
                    # Camarilla Pivots H3 / H4
                    if pivots is not None:
                        if o["side"] == "BUY":
                            h3, h4 = pivots['H3'], pivots['H4']
                            dist_h3 = (h3 - filled) / filled
                            tgt_px = round(h3, 2) if dist_h3 >= 0.0065 else round(h4, 2)
                        else:
                            l3, l4 = pivots['L3'], pivots['L4']
                            dist_l3 = (filled - l3) / filled
                            tgt_px = round(l3, 2) if dist_l3 >= 0.0065 else round(l4, 2)
                    else:
                        tgt_px = round(filled * 1.012 if o["side"] == "BUY" else filled * 0.988, 2)
                    tgt_pct = abs(tgt_px / filled - 1.0)
                    
                elif self.model_type in ('confluence_target', 'pure_swing_trailing'):
                    # Multi-Anchor Confluence:
                    # Select nearest valid structural resistance (R1, R2, PDH, H3, H4) >= 0.80%
                    if pivots is not None:
                        if o["side"] == "BUY":
                            candidates = [pivots['R1'], pivots['H3'], pivots['pdh'], pivots['H4'], pivots['R2']]
                            valid_cand = [c for c in sorted(candidates) if (c - filled) / filled >= 0.0080]
                            tgt_px = round(valid_cand[0], 2) if valid_cand else round(pivots['R2'], 2)
                        else:
                            candidates = [pivots['S1'], pivots['L3'], pivots['pdl'], pivots['L4'], pivots['S2']]
                            valid_cand = [c for c in sorted(candidates, reverse=True) if (filled - c) / filled >= 0.0080]
                            tgt_px = round(valid_cand[0], 2) if valid_cand else round(pivots['S2'], 2)
                    else:
                        tgt_px = round(filled * 1.015 if o["side"] == "BUY" else filled * 0.985, 2)
                    tgt_pct = abs(tgt_px / filled - 1.0)
                    
                    if self.model_type == 'pure_swing_trailing':
                        tgt_px = 999999.0 if o["side"] == "BUY" else 0.01

                pos = dict(o, entry=round(filled, 2), entry_time=str(t + BAR), bars_held=0,
                           sl=sl_px, initial_sl=sl_px, tgt=tgt_px, tgt_pct=round(tgt_pct * 100, 2),
                           ltp=cl, mfe=0.0, mae=0.0, sl_state="INITIAL_0.85PCT")
                self.open[sym] = pos
                self.traded_symbols_today.add(sym)
                if (pos["side"] == "BUY" and lo <= pos["sl"]) or (pos["side"] == "SELL" and hi >= pos["sl"]):
                    self._close(sym, pos["sl"], t, "STOP_LOSS (same bar as fill)", market=True)
                continue

            o["bars_waited"] += 1
            if o["bars_waited"] >= C.ENTRY_LIMIT_WAIT_BARS:
                del self.pending[sym]

    def _manage_open_profit_max(self, t, decide: dtime):
        mins_left = max(1, ((15 * 60 + 5) - (decide.hour * 60 + decide.minute)))
        bars_left = mins_left // 5
        for sym in list(self.open):
            p = self.open[sym]
            sd, i = self._bar(sym, t)
            if sd is None:
                if decide >= C.FORCE_LIQUIDATION:
                    self._close(sym, p["ltp"], t, "3_PM_EXIT (Session Liquidation)", market=True)
                continue
            hi, lo, cl, op = float(sd.h[i]), float(sd.l[i]), float(sd.c[i]), float(sd.o[i])
            p["ltp"], p["bars_held"] = cl, p["bars_held"] + 1
            sgn = 1 if p["side"] == "BUY" else -1
            bar_mfe = sgn * ((hi if sgn > 0 else lo) / p["entry"] - 1) * 100
            bar_mae = sgn * ((lo if sgn > 0 else hi) / p["entry"] - 1) * 100

            if decide >= C.FORCE_LIQUIDATION:
                p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                self._close(sym, cl, t, "3_PM_EXIT (Session Liquidation)", market=True); continue

            # Late-session harvest for baseline only
            if self.model_type == 'baseline' and bars_left <= 10 and p["mfe"] >= 0.35:
                harvest_tgt = round(p["entry"] * (1.0 + sgn * 0.0055), 2)
                if (sgn > 0 and harvest_tgt < p["tgt"] and harvest_tgt > op) or \
                   (sgn < 0 and harvest_tgt > p["tgt"] and harvest_tgt < op):
                    p["tgt"] = harvest_tgt

            # Check SL & Target hit
            if p["side"] == "BUY":
                if lo <= p["sl"]:
                    p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                    lbl = f"PROFIT_LOCK ({p['sl_state']})" if not p["sl_state"].startswith("INITIAL") else "STOP_LOSS"
                    self._close(sym, min(p["sl"], op), t, lbl, market=True); continue
                if hi >= p["tgt"]:
                    p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                    self._close(sym, max(p["tgt"], op), t, f"TARGET (Structural Limit {p.get('tgt_pct', 0.8)}%)", market=False); continue
            else:
                if hi >= p["sl"]:
                    p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                    lbl = f"PROFIT_LOCK ({p['sl_state']})" if not p["sl_state"].startswith("INITIAL") else "STOP_LOSS"
                    self._close(sym, max(p["sl"], op), t, lbl, market=True); continue
                if lo <= p["tgt"]:
                    p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                    self._close(sym, min(p["tgt"], op), t, f"TARGET (Structural Limit {p.get('tgt_pct', 0.8)}%)", market=False); continue

            p["mfe"] = max(p["mfe"], bar_mfe)
            p["mae"] = min(p["mae"], bar_mae)

            # -------------------------------------------------------------
            # EXIT MODEL LOGIC
            # -------------------------------------------------------------
            if self.model_type == 'baseline':
                # Original 4-Stage Ratchet Curve
                if p["mfe"] >= 0.70:
                    lock_pct = (p["mfe"] * 0.68) / 100.0
                    new_sl = round(p["entry"] * (1.0 + sgn * lock_pct), 2)
                    if (sgn > 0 and new_sl > p["sl"]) or (sgn < 0 and new_sl < p["sl"]):
                        p["sl"] = new_sl; p["sl_state"] = f"STAGE3_TRAIL_68% (+{lock_pct*100:.2f}%)"
                elif p["mfe"] >= 0.50:
                    lock_pct = (p["mfe"] * 0.58) / 100.0
                    new_sl = round(p["entry"] * (1.0 + sgn * lock_pct), 2)
                    if (sgn > 0 and new_sl > p["sl"]) or (sgn < 0 and new_sl < p["sl"]):
                        p["sl"] = new_sl; p["sl_state"] = f"STAGE2_TRAIL_58% (+{lock_pct*100:.2f}%)"
                elif p["mfe"] >= 0.34:
                    be_px = breakeven_exit_price(p["side"], p["qty"], p["entry"], market_exit=True, min_net_rupees=160.0)
                    if (sgn > 0 and be_px > p["sl"]) or (sgn < 0 and be_px < p["sl"]):
                        p["sl"] = be_px; p["sl_state"] = "STAGE1_TAX_SHIELD"
                elif p["mfe"] >= 0.24:
                    half_sl = round(p["entry"] * (1.0 - sgn * 0.0038), 2)
                    if (sgn > 0 and half_sl > p["sl"]) or (sgn < 0 and half_sl < p["sl"]):
                        p["sl"] = half_sl; p["sl_state"] = "STAGE0_RISK_HALVED"

            else:
                # STRUCTURAL TRAILING STOP & STATISTICAL EXHAUSTION (Models 1, 2, 3, 4)
                # 1. Statistical Exhaustion Band Check (VWAP + 2.0 Sigma)
                vwap_val = float(sd.vwap[i])
                v_std = float(vwap_std_cache[sym][i])
                upper_2s = vwap_val + 2.0 * v_std
                lower_2s = vwap_val - 2.0 * v_std
                
                # Check VWAP + 2.0 Sigma exhaustion
                if p["side"] == "BUY" and hi >= upper_2s and p["mfe"] >= 0.70:
                    if cl < op or (hi - cl) > (cl - lo):
                        self._close(sym, cl, t, "EXHAUSTION (VWAP +2.0 Sigma Rejection)", market=True)
                        continue
                elif p["side"] == "SELL" and lo <= lower_2s and p["mfe"] >= 0.70:
                    if cl > op or (cl - lo) > (hi - cl):
                        self._close(sym, cl, t, "EXHAUSTION (VWAP -2.0 Sigma Rejection)", market=True)
                        continue

                # 2. Structural Trailing behind 5m Swing Low / Swing High
                if p["bars_held"] >= 3 and p["mfe"] >= 0.40:
                    # Step 1: Structural Breakeven / Tax-Shield at MFE >= 0.45%
                    be_px = breakeven_exit_price(p["side"], p["qty"], p["entry"], market_exit=True, min_net_rupees=150.0)
                    if (sgn > 0 and be_px > p["sl"]) or (sgn < 0 and be_px < p["sl"]):
                        p["sl"] = be_px
                        p["sl_state"] = "STRUCTURAL_BREAKEVEN"

                    # Step 2: Swing Low (3-bar lowest low for BUY / 3-bar highest high for SELL)
                    start_k = max(0, i - 3)
                    if p["side"] == "BUY":
                        swing_low = float(np.min(sd.l[start_k:i]))
                        trail_px = round(swing_low * 0.9995, 2)
                        if trail_px > p["sl"]:
                            p["sl"] = trail_px
                            p["sl_state"] = f"SWING_LOW_TRAIL ({trail_px})"
                    else:
                        swing_high = float(np.max(sd.h[start_k:i]))
                        trail_px = round(swing_high * 1.0005, 2)
                        if trail_px < p["sl"]:
                            p["sl"] = trail_px
                            p["sl_state"] = f"SWING_HIGH_TRAIL ({trail_px})"

            # Universal Structural Reversal Exit (10m cross + VWAP + 5m EMA30 break against us)
            c10, _ = sd.tenmin_state(i)
            vwap_val = float(sd.vwap[i])
            e30 = float(sd.ema[30][i])
            if p["bars_held"] >= 3:
                if p["side"] == "BUY" and c10 == "BEARISH" and cl < vwap_val and cl < e30:
                    self._close(sym, cl, t, "STRUCTURAL_REVERSAL (10m+VWAP+EMA30)", market=True); continue
                if p["side"] == "SELL" and c10 == "BULLISH" and cl > vwap_val and cl > e30:
                    self._close(sym, cl, t, "STRUCTURAL_REVERSAL (10m+VWAP+EMA30)", market=True); continue


# -------------------------------------------------------------------------
# Run Backtest Across All Models
# -------------------------------------------------------------------------
models_to_test = [
    ('baseline', 'Baseline (0.95% Cap + Fixed % Ratchet)'),
    ('floor_pivots', 'Model 1: Floor Pivots (R1/R2) + 3-Bar Swing Trail'),
    ('camarilla', 'Model 2: Camarilla Pivots (H3/H4) + 3-Bar Swing Trail'),
    ('confluence_target', 'Model 3: Confluence Structural Targets + 3-Bar Swing Trail + VWAP 2S'),
    ('pure_swing_trailing', 'Model 4: Pure 3-Bar Swing Trail (No Target Cap) + VWAP 2S')
]

results = []

for m_key, m_label in models_to_test:
    tot_trades, tot_wins = 0, 0
    tot_gross, tot_costs, tot_net = 0.0, 0.0, 0.0
    all_closed = []
    
    for d in dates:
        dt = datetime.combine(d, datetime.min.time())
        sim = StructuralSimulator(dt, model_type=m_key, starting_capital=500000.0)
        sim.syms = master.syms
        sim.indices = master.indices
        sim.nifty_pos = master.nifty_pos
        sim.nifty_e10 = master.nifty_e10
        sim.nifty_e21 = master.nifty_e21
        sim.nifty_e30 = master.nifty_e30
        for t in sim.session_bar_times():
            sim.step(t)
        for c in sim.closed:
            all_closed.append(c)
            tot_trades += 1
            if c['net'] > 0:
                tot_wins += 1
            tot_gross += c['gross']
            tot_costs += c['costs']
            tot_net += c['net']
            
    wr = round(tot_wins / tot_trades * 100, 1) if tot_trades else 0.0
    avg_trade_net = tot_net / tot_trades if tot_trades else 0.0
    
    # Calculate R-multiples (where Risk = 0.85% initial stop)
    r_multiples = []
    for c in all_closed:
        risk_rs = c['entry'] * 0.0085 * c['qty']
        r = c['net'] / risk_rs if risk_rs > 0 else 0.0
        r_multiples.append(r)
    avg_r = np.mean(r_multiples) if r_multiples else 0.0
    
    # Exit reasons breakdown
    reasons = {}
    for c in all_closed:
        base_r = c['reason'].split('(')[0].strip()
        reasons[base_r] = reasons.get(base_r, 0) + 1
        
    results.append({
        'key': m_key, 'label': m_label,
        'trades': tot_trades, 'wins': tot_wins, 'wr': wr,
        'gross': round(tot_gross, 2), 'costs': round(tot_costs, 2), 'net': round(tot_net, 2),
        'avg_net_trade': round(avg_trade_net, 2), 'avg_r': round(avg_r, 2),
        'reasons': reasons, 'closed': all_closed
    })

print("\n==========================================================================================")
print("COMPARATIVE BACKTEST RESULTS: STRUCTURAL EXIT MODELS VS BASELINE (61 DAYS, 139 STOCKS)")
print("==========================================================================================")
header = f"{'Model':<45} | {'Trades':<6} | {'WR (%)':<6} | {'Gross P&L':<14} | {'Taxes/Fees':<10} | {'Net P&L':<14} | {'Avg R':<6}"
print(header)
print("-" * len(header))
for r in results:
    print(f"{r['label']:<45} | {r['trades']:<6} | {r['wr']:<6.1f} | Rs. {r['gross']:+12,.2f} | Rs. {r['costs']:8,.2f} | Rs. {r['net']:+12,.2f} | {r['avg_r']:+5.2f}R")

print("\n--- Detailed Exit Breakdown per Model ---")
for r in results:
    print(f"\n{r['label']}:")
    for reason, count in sorted(r['reasons'].items(), key=lambda x: -x[1]):
        pct = count / r['trades'] * 100 if r['trades'] else 0
        print(f"  * {reason:<25}: {count:>3} trades ({pct:4.1f}%)")
