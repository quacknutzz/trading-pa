import sys, pickle, math
from datetime import datetime, time as dtime
import pandas as pd
import numpy as np

sys.path.insert(0, 'sim')
import datafeed
import config as C
from simulator import Simulator, trade_costs, breakeven_exit_price, BAR

u = datafeed.load_universe()
with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)
raw = payload['raw']
idx = payload['idx']
sym0 = list(raw.keys())[0]
dates = sorted(list(set(raw[sym0].index.date)))

master = Simulator('replay', datetime.combine(dates[0], datetime.min.time()), strategy='v6_baseline', starting_capital=500000.0)
master.load(raw, u, idx)

# Unified test simulator allowing modular toggles for each fix
class ModularFixSimulator(Simulator):
    def __init__(self, session_date, 
                 fix4b_cooldown=False,
                 fix1_prefill=False,
                 fix2_vwap=False,
                 fix4a_rising_adx=False,
                 fix5a_min_atr=0.0,
                 fix7_nifty=False,
                 fix8_cutoff=dtime(14, 45),
                 starting_capital=500000.0):
        super().__init__('replay', session_date, strategy='v6_baseline', starting_capital=starting_capital)
        self.fix4b_cooldown = fix4b_cooldown
        self.fix1_prefill = fix1_prefill
        self.fix2_vwap = fix2_vwap
        self.fix4a_rising_adx = fix4a_rising_adx
        self.fix5a_min_atr = fix5a_min_atr
        self.fix7_nifty = fix7_nifty
        self.fix8_cutoff = fix8_cutoff
        self.traded_today = set()

    def _manage_open_baseline(self, t, decide: dtime):
        for sym in list(self.open):
            p = self.open[sym]
            sd, i = self._bar(sym, t)
            if sd is None:
                if decide >= C.FORCE_LIQUIDATION:
                    self._close(sym, p["ltp"], t, "3_PM_EXIT", market=True)
                continue
            hi, lo, cl, op = float(sd.h[i]), float(sd.l[i]), float(sd.c[i]), float(sd.o[i])
            p["ltp"], p["bars_held"] = cl, p["bars_held"] + 1
            sgn = 1 if p["side"] == "BUY" else -1
            bar_mfe = sgn * ((hi if sgn > 0 else lo) / p["entry"] - 1) * 100
            bar_mae = sgn * ((lo if sgn > 0 else hi) / p["entry"] - 1) * 100
            p["mfe"] = max(p["mfe"], bar_mfe)
            p["mae"] = min(p["mae"], bar_mae)

            if decide >= C.FORCE_LIQUIDATION:
                self._close(sym, cl, t, "3_PM_EXIT", market=True); continue

            # SL & Target
            if p["side"] == "BUY":
                if lo <= p["sl"]:
                    self._close(sym, min(p["sl"], op), t, "STOP_LOSS", market=True); continue
                if hi >= p["tgt"]:
                    self._close(sym, max(p["tgt"], op), t, "TARGET", market=False); continue
            else:
                if hi >= p["sl"]:
                    self._close(sym, max(p["sl"], op), t, "STOP_LOSS", market=True); continue
                if lo <= p["tgt"]:
                    self._close(sym, min(p["tgt"], op), t, "TARGET", market=False); continue

            # Breakeven Ratchet: if MFE >= 0.45%, snap SL to breakeven + Rs. 50
            if p["mfe"] >= 0.45:
                be_px = breakeven_exit_price(p["side"], p["qty"], p["entry"], market_exit=True, min_net_rupees=50.0)
                if (sgn > 0 and be_px > p["sl"]) or (sgn < 0 and be_px < p["sl"]):
                    p["sl"] = be_px

            # Fix #6 Optimized Structural Exit Engine:
            # 1. 10m contra cross = hard exit after 2-bar grace
            c10, _ = sd.tenmin_state(i)
            is_10m_contra = (p["side"] == "BUY" and c10 == "BEARISH") or (p["side"] == "SELL" and c10 == "BULLISH")
            if p["bars_held"] >= 2 and is_10m_contra:
                self._close(sym, cl, t, "10M_CONTRA_HARD", market=True); continue

            # 2. 5m Reverse Cross ONLY IF VWAP is also lost (pullback filter) after bar 3
            c5 = sd.cross5m(i)
            vwap = float(sd.vwap[i])
            is_5m_reverse = (p["side"] == "BUY" and c5 == "BEARISH") or (p["side"] == "SELL" and c5 == "BULLISH")
            lost_vwap = (p["side"] == "BUY" and cl < vwap) or (p["side"] == "SELL" and cl > vwap)
            if p["bars_held"] >= 3 and is_5m_reverse and lost_vwap:
                self._close(sym, cl, t, "5M_REV_AND_LOST_VWAP", market=True); continue

    def _manage_pending_baseline(self, t):
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

            # Fix #1: Pre-Fill Contra Validation
            # If 10m EMA cross or momentum flipped against pending limit order BEFORE or AT fill bar, cancel!
            if self.fix1_prefill:
                c10, _ = sd.tenmin_state(i)
                if (o["side"] == "BUY" and c10 == "BEARISH") or (o["side"] == "SELL" and c10 == "BULLISH"):
                    del self.pending[sym]
                    continue

            if filled is not None:
                # Fix #2c (Causal): Prior candle or fill candle bounce check
                # (For BUY, reject if fill bar is a violent red breakdown candle closing well below limit)
                if self.fix2_vwap:
                    vwap_val = float(sd.vwap[i])
                    if (o["side"] == "BUY" and cl < vwap_val) or (o["side"] == "SELL" and cl > vwap_val):
                        del self.pending[sym]
                        continue

                del self.pending[sym]
                sl_px = round(filled * (1.0 - C.SL_PCT) if o["side"] == "BUY" else filled * (1.0 + C.SL_PCT), 2)
                tgt_px = round(filled * (1.0 + C.TGT_PCT) if o["side"] == "BUY" else filled * (1.0 - C.TGT_PCT), 2)
                pos = dict(o, entry=round(filled, 2), entry_time=str(t + BAR), bars_held=0,
                           sl=sl_px, tgt=tgt_px, ltp=cl, mfe=0.0, mae=0.0)
                self.open[sym] = pos
                if self.fix4b_cooldown:
                    self.traded_today.add(sym)
                if (pos["side"] == "BUY" and lo <= pos["sl"]) or (pos["side"] == "SELL" and hi >= pos["sl"]):
                    self._close(sym, pos["sl"], t, "STOP_LOSS (fill bar)", market=True)
                continue

            o["bars_waited"] += 1
            if o["bars_waited"] >= C.ENTRY_LIMIT_WAIT_BARS:
                del self.pending[sym]

    def _scan_baseline(self, t, decide):
        if decide < C.ENTRY_START or decide >= self.fix8_cutoff:
            return

        # Fix #7: Nifty Trend Filter
        if self.fix7_nifty:
            ndf = self.indices.get("NIFTY 50")
            if ndf is not None and t in self.nifty_pos:
                ni = self.nifty_pos[t]
                n_cl = float(ndf["close"].iloc[ni])
                # Check Nifty open of the day
                d_bars = ndf.loc[ndf.index.normalize() == t.normalize()]
                d_bars_up = d_bars.loc[d_bars.index <= t]
                if len(d_bars_up) > 0:
                    n_open = float(d_bars_up["open"].iloc[0])
                    # If Nifty is red (below open), block BUYs. If Nifty is green, block SELLs.
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
            if i < 2:
                continue
            c10, rsi10 = sd.tenmin_state(i)
            if not c10:
                continue
            side = "BUY" if c10 == "BULLISH" else "SELL"

            # Apply Nifty Bias filter
            if self.fix7_nifty:
                if side == "BUY" and nifty_bias == "BEARISH":
                    continue
                if side == "SELL" and nifty_bias == "BULLISH":
                    continue

            ltp = float(sd.c[i])
            adx = float(sd.adx[i])
            prev_adx = float(sd.adx[i - 1])

            # Fix #4a: Rising ADX
            if self.fix4a_rising_adx:
                if math.isnan(adx) or math.isnan(prev_adx) or adx < prev_adx:
                    continue

            # Fix #5a: Minimum ATR Hurdle
            atr_val = float(sd.atr[i]) if not math.isnan(sd.atr[i]) else 0.0
            atr_pct = (atr_val / ltp * 100.0) if ltp > 0 else 0.0
            if self.fix5a_min_atr > 0 and atr_pct < self.fix5a_min_atr:
                continue

            # Fix #2a: Intraday VWAP Side (Buy only above VWAP, Sell only below VWAP)
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

# Define configurations to test
configs = [
    ("Step 0: New Exit Engine (Fix #6 + Ratchet)", {}),
    ("Step 1: + Fix #4b (Per-Stock Cooldown)", {"fix4b_cooldown": True}),
    ("Step 2: + Fix #1 (Pre-Fill Contra Cancel)", {"fix4b_cooldown": True, "fix1_prefill": True}),
    ("Step 3: + Fix #2 (Causal VWAP Side)", {"fix4b_cooldown": True, "fix1_prefill": True, "fix2_vwap": True}),
    ("Step 4: + Fix #5a (Min ATR >= 0.15%)", {"fix4b_cooldown": True, "fix1_prefill": True, "fix2_vwap": True, "fix5a_min_atr": 0.15}),
    ("Step 5: + Fix #4a (Rising ADX)", {"fix4b_cooldown": True, "fix1_prefill": True, "fix2_vwap": True, "fix5a_min_atr": 0.15, "fix4a_rising_adx": True}),
    ("Step 6: + Fix #7 (Nifty Trend Filter)", {"fix4b_cooldown": True, "fix1_prefill": True, "fix2_vwap": True, "fix5a_min_atr": 0.15, "fix7_nifty": True}),
    ("Step 7: + Fix #8 (14:00 Entry Cutoff)", {"fix4b_cooldown": True, "fix1_prefill": True, "fix2_vwap": True, "fix5a_min_atr": 0.15, "fix7_nifty": True, "fix8_cutoff": dtime(14, 0)}),
]

summary_results = []
for label, kwargs in configs:
    closed_all = []
    for d in dates:
        dt = datetime.combine(d, datetime.min.time())
        sim = ModularFixSimulator(dt, **kwargs, starting_capital=500000.0)
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
    summary_results.append({
        "Configuration": label,
        "Trades": n,
        "Wins": wins,
        "Win Rate (%)": round(wins / n * 100, 1) if n else 0,
        "Gross (Rs)": round(gross, 2),
        "Costs (Rs)": round(costs, 2),
        "Net (Rs)": round(net, 2),
        "Avg Net/Trd (Rs)": round(net / n, 2) if n else 0
    })

res_df = pd.DataFrame(summary_results)
print("\n==================== STEP-BY-STEP CUMULATIVE RESEARCH (60 TRADING DAYS) ====================")
print(res_df.to_string(index=False))
