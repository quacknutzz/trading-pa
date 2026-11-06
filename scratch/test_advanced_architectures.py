import sys, pickle, math
from datetime import datetime, time as dtime
import pandas as pd
import numpy as np

sys.path.insert(0, 'sim')
sys.path.insert(0, '.')
import config as C
from simulator import BAR
from scratch.research_all_remaining_fixes import ModularFixSimulator, master, dates

# Let's test combinations to see the true optimal configuration:
# Test A: All Good Fixes WITHOUT Rising ADX (Cooldown + PreFill + VWAP Side + Min ATR + Nifty + 14:00 Cutoff)
# Test B: With 3-bar smoothed ADX (ADX > ADX_3bars_ago)
# Test C: Adding Dynamic ATR Limit Target (Fix #5b) to Test A

class AdvancedFixSimulator(ModularFixSimulator):
    def __init__(self, session_date, smoothed_adx=False, atr_target=False, **kwargs):
        super().__init__(session_date, **kwargs)
        self.smoothed_adx = smoothed_adx
        self.atr_target = atr_target

    def _manage_pending_baseline(self, t):
        # Override to optionally set ATR target
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

            if self.fix1_prefill:
                c10, _ = sd.tenmin_state(i)
                if (o["side"] == "BUY" and c10 == "BEARISH") or (o["side"] == "SELL" and c10 == "BULLISH"):
                    del self.pending[sym]; continue

            if filled is not None:
                if self.fix2_vwap:
                    vwap_val = float(sd.vwap[i])
                    if (o["side"] == "BUY" and cl < vwap_val) or (o["side"] == "SELL" and cl > vwap_val):
                        del self.pending[sym]; continue

                del self.pending[sym]
                sl_px = round(filled * (1.0 - 0.0085) if o["side"] == "BUY" else filled * (1.0 + 0.0085), 2)
                
                if self.atr_target:
                    atr_val = float(sd.atr[i]) if not math.isnan(sd.atr[i]) else (filled * 0.0025)
                    atr_pct = (atr_val / filled) if filled > 0 else 0.0025
                    tgt_pct = min(max(atr_pct * 2.8, 0.0065), 0.0120)
                    tgt_px = round(filled * (1.0 + tgt_pct) if o["side"] == "BUY" else filled * (1.0 - tgt_pct), 2)
                else:
                    tgt_px = round(filled * (1.0 + 0.012) if o["side"] == "BUY" else filled * (1.0 - 0.012), 2)

                pos = dict(o, entry=round(filled, 2), entry_time=str(t + BAR), bars_held=0,
                           sl=sl_px, tgt=tgt_px, ltp=cl, mfe=0.0, mae=0.0)
                self.open[sym] = pos
                if self.fix4b_cooldown:
                    self.traded_today.add(sym)
                if (pos["side"] == "BUY" and lo <= pos["sl"]) or (pos["side"] == "SELL" and hi >= pos["sl"]):
                    self._close(sym, pos["sl"], t, "STOP_LOSS (fill bar)", market=True)
                continue

            o["bars_waited"] += 1
            if o["bars_waited"] >= 3:
                del self.pending[sym]

    def _scan_baseline(self, t, decide):
        if decide < C.ENTRY_START or decide >= self.fix8_cutoff:
            return

        if self.fix7_nifty:
            ndf = self.indices.get("NIFTY 50")
            if ndf is not None and t in self.nifty_pos:
                ni = self.nifty_pos[t]
                n_cl = float(ndf["close"].iloc[ni])
                d_bars = ndf.loc[ndf.index.normalize() == t.normalize()]
                d_bars_up = d_bars.loc[d_bars.index <= t]
                if len(d_bars_up) > 0:
                    n_open = float(d_bars_up["open"].iloc[0])
                    nifty_bias = "BULLISH" if n_cl >= n_open else "BEARISH"
                else:
                    nifty_bias = "NEUTRAL"
            else:
                nifty_bias = "NEUTRAL"
        else:
            nifty_bias = "NEUTRAL"

        cands = []
        for sym, sd in self.syms.items():
            if self.fix4b_cooldown and sym in self.traded_today:
                continue
            if t not in sd.pos:
                continue
            i = sd.pos[t]
            if i < 4:
                continue
            c10, rsi10 = sd.tenmin_state(i)
            if not c10:
                continue
            side = "BUY" if c10 == "BULLISH" else "SELL"

            if self.fix7_nifty:
                if side == "BUY" and nifty_bias == "BEARISH":
                    continue
                if side == "SELL" and nifty_bias == "BULLISH":
                    continue

            ltp = float(sd.c[i])
            adx = float(sd.adx[i])

            if self.smoothed_adx:
                # 3-bar smoothed: ADX must be higher than 3 bars ago (15m momentum trend)
                prev3_adx = float(sd.adx[i - 3])
                if math.isnan(adx) or math.isnan(prev3_adx) or adx < prev3_adx:
                    continue
            elif self.fix4a_rising_adx:
                prev_adx = float(sd.adx[i - 1])
                if math.isnan(adx) or math.isnan(prev_adx) or adx < prev_adx:
                    continue

            atr_val = float(sd.atr[i]) if not math.isnan(sd.atr[i]) else 0.0
            atr_pct = (atr_val / ltp * 100.0) if ltp > 0 else 0.0
            if self.fix5a_min_atr > 0 and atr_pct < self.fix5a_min_atr:
                continue

            if self.fix2_vwap:
                vwap_val = float(sd.vwap[i])
                if (side == "BUY" and ltp < vwap_val) or (side == "SELL" and ltp > vwap_val):
                    continue

            if sym in self.open or sym in self.pending:
                continue
            if C.RSI_FILTER_ENABLED and (math.isnan(rsi10) or not (C.BUY_RSI_MIN <= rsi10 <= C.BUY_RSI_MAX if side == "BUY" else C.SELL_RSI_MIN <= rsi10 <= C.SELL_RSI_MAX)):
                continue
            if math.isnan(adx) or adx < C.STOCK_ADX_MIN:
                continue

            rvol = float(sd.rvol[i]) if not math.isnan(sd.rvol[i]) else 0.0
            cands.append((sym, side, ltp, rsi10, adx, rvol))

        cands.sort(key=lambda x: x[5], reverse=True)
        free_slots = C.MAX_CONCURRENT_POSITIONS - (len(self.open) + len(self.pending))
        for c in cands[:max(0, free_slots)]:
            sym, side, ltp = c[0], c[1], c[2]
            sd, i = self._bar(sym, t)
            e_val = float(sd.ema[C.EMA_FAST][i])
            lim_px = round(min(ltp, e_val) if side == "BUY" else max(ltp, e_val), 2)
            budget = (self.starting_capital * C.LEVERAGE) / C.MAX_CONCURRENT_POSITIONS
            qty = int(budget // lim_px)
            if qty <= 0:
                continue
            self.seq += 1
            self.pending[sym] = dict(
                id=self.seq, symbol=sym, sector=sd.sector, side=side, qty=qty,
                limit=lim_px, ema_period=C.EMA_FAST, signal_bar=t, signal_time=str(t + BAR),
                signal_ltp=ltp, rsi=c[3], adx=c[4], rvol=c[5], bars_waited=0)
            if self.fix4b_cooldown:
                self.traded_today.add(sym)

comparisons = [
    ("Config A: All Core Fixes (Drop Bar ADX, Keep Cooldown+PreFill+VWAP+ATR+Nifty+14:00)", {
        "fix4b_cooldown": True, "fix1_prefill": True, "fix2_vwap": True, 
        "fix5a_min_atr": 0.15, "fix4a_rising_adx": False, "smoothed_adx": False,
        "fix7_nifty": True, "fix8_cutoff": dtime(14, 0), "atr_target": False
    }),
    ("Config B: Config A + Smoothed 3-Bar ADX Trend", {
        "fix4b_cooldown": True, "fix1_prefill": True, "fix2_vwap": True, 
        "fix5a_min_atr": 0.15, "fix4a_rising_adx": False, "smoothed_adx": True,
        "fix7_nifty": True, "fix8_cutoff": dtime(14, 0), "atr_target": False
    }),
    ("Config C: Config A + Dynamic ATR Target (Fix #5b)", {
        "fix4b_cooldown": True, "fix1_prefill": True, "fix2_vwap": True, 
        "fix5a_min_atr": 0.15, "fix4a_rising_adx": False, "smoothed_adx": False,
        "fix7_nifty": True, "fix8_cutoff": dtime(14, 0), "atr_target": True
    }),
]

final_rows = []
for label, kwargs in comparisons:
    closed_all = []
    for d in dates:
        dt = datetime.combine(d, datetime.min.time())
        sim = AdvancedFixSimulator(dt, **kwargs, starting_capital=500000.0)
        sim.syms = master.syms
        sim.indices = master.indices
        sim.nifty_pos = master.nifty_pos
        sim.nifty_e10 = master.nifty_e10
        sim.nifty_e21 = master.nifty_e21
        sim.nifty_e30 = master.nifty_e30
        for t in sim.session_bar_times():
            sim.step(t)
        closed_all.extend(sim.closed)
    
    df = pd.DataFrame(closed_all)
    n = len(df)
    wins = (df['net'] > 0).sum() if n else 0
    gross = df['gross'].sum() if n else 0.0
    costs = df['costs'].sum() if n else 0.0
    net = df['net'].sum() if n else 0.0
    final_rows.append({
        "Configuration": label,
        "Trades": n,
        "Wins": wins,
        "Win Rate (%)": round(wins / n * 100, 1) if n else 0,
        "Gross (Rs)": round(gross, 2),
        "Costs (Rs)": round(costs, 2),
        "Net (Rs)": round(net, 2),
        "Avg Net/Trd (Rs)": round(net / n, 2) if n else 0
    })

res_df = pd.DataFrame(final_rows)
print("\n==================== ADVANCED ARCHITECTURE COMPARISON (60 TRADING DAYS) ====================")
print(res_df.to_string(index=False))
