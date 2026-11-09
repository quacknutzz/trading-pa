"""
sim/simulator.py -- Multi-Strategy Paper-Trading & Net-Profit Optimization Engine.

Supports 3 switchable strategy modes on the exact same bar-by-bar loop (both LIVE and REPLAY):
  1. "v6_profit_max"   (Default) -- Mathematical Net-Profit Maximizer (V6++):
                                    Fixes #1-#8 + Hybrid Multi-Factor Structural Exit + 2-Stage Tax-Aware Ratchet
                                    + 2-Bar Smoothed Rising ADX + Per-Stock Cooldown + Strict Equal Slot Sizing.
  2. "v6_fixes_1_to_5"           -- Approved Fixes #1 to #5 (Zero-Delay Pre-Fill Validity, Intraday VWAP + Bounce,
                                    2-Tier Tax-Aware Ratchet, Per-Stock Cooldown + Rising ADX, Tax Dead-Zone Gate).
  3. "v6_baseline"               -- Exact replica of Dad's Module_COCKPIT_V6 VBA code.
"""
import math
import os
import sys
import threading
from datetime import datetime, timedelta, time as dtime
from typing import Dict, List, Optional

import numpy as np
import pandas as pd

SIM_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(SIM_DIR)
sys.path.insert(0, ROOT_DIR)
sys.path.insert(0, SIM_DIR)

import config as C
from engine.indicators import ema_fast, rsi_wilder
from engine import numba_kernels as nk

BAR = timedelta(minutes=5)


def _cross(prev_diff: float, cur_diff: float) -> str:
    if prev_diff is None or cur_diff is None or (isinstance(prev_diff, float) and math.isnan(prev_diff)) or (isinstance(cur_diff, float) and math.isnan(cur_diff)):
        return ""
    if prev_diff <= 0 < cur_diff:
        return "BULLISH"
    if prev_diff >= 0 > cur_diff:
        return "BEARISH"
    return ""


# ============================================================ costs & mathematical objective
def trade_costs(side: str, qty: int, entry: float, exit_: float, market_exit: bool) -> dict:
    buy_px, sell_px = (entry, exit_) if side == "BUY" else (exit_, entry)
    buy_val, sell_val = buy_px * qty, sell_px * qty
    brk = min(C.BROKERAGE_PER_ORDER, buy_val * C.BROKERAGE_PCT_CAP) + \
          min(C.BROKERAGE_PER_ORDER, sell_val * C.BROKERAGE_PCT_CAP)
    stt = sell_val * C.STT_SELL
    txn = (buy_val + sell_val) * C.NSE_TXN
    sebi = (buy_val + sell_val) * C.SEBI_PER_CRORE / 1e7
    stamp = buy_val * C.STAMP_BUY
    gst = (brk + txn + sebi) * C.GST
    slip = exit_ * qty * C.MARKET_SLIPPAGE if market_exit else 0.0
    total = brk + stt + txn + sebi + stamp + gst + slip
    return {"brokerage": brk, "stt": stt, "exchange": txn + sebi, "stamp": stamp,
            "gst": gst, "slippage": slip, "total": total}


def breakeven_exit_price(side: str, qty: int, entry: float, market_exit: bool = True, min_net_rupees: float = 0.0) -> float:
    """
    Solves for the exact exit price P_exit such that:
      Pi_net(side, qty, entry, P_exit, market_exit) == min_net_rupees
    """
    if qty <= 0 or entry <= 0:
        return entry
    sgn = 1.0 if side == "BUY" else -1.0
    est_px = entry * (1.0 + sgn * (0.0012 if market_exit else 0.0007))
    for _ in range(3):
        c = trade_costs(side, qty, entry, est_px, market_exit)["total"]
        req_gross = c + min_net_rupees
        est_px = entry + sgn * (req_gross / qty)
    return round(est_px, 2)


# ============================================================ O(N) Vectorized SymbolData
class SymbolData:
    """
    5m bars + O(N) precomputed causal indicator arrays (5m EMAs, ADX, +DI/-DI, ATR14,
    Intraday Session VWAP, and incremental 10m EMA10/30 + Wilder RSI14).
    """

    def __init__(self, symbol: str, sector: str, df: pd.DataFrame):
        self.symbol, self.sector = symbol, sector
        self.set_df(df)

    def set_df(self, df: pd.DataFrame):
        self.df = df
        self.pos = {ts: i for i, ts in enumerate(df.index)}
        n = len(df)
        c = df["close"].values.astype(np.float64)
        h = df["high"].values.astype(np.float64)
        l = df["low"].values.astype(np.float64)
        o = df["open"].values.astype(np.float64)
        v = df["volume"].values.astype(np.float64)
        self.c, self.h, self.l, self.o, self.v = c, h, l, o, v

        self.ema = {p: ema_fast(c, p) for p in set(C.PULLBACK_EMAS) | {C.EMA_FAST, C.EMA_SLOW, 21}}
        if n > 2 * C.ADX_PERIOD:
            pdi, mdi, adx_arr = nk.adx(h, l, c, C.ADX_PERIOD)
            self.plus_di, self.minus_di, self.adx = pdi, mdi, adx_arr
            self.atr = nk.atr(h, l, c, 14)
        else:
            self.plus_di = np.full(n, np.nan)
            self.minus_di = np.full(n, np.nan)
            self.adx = np.full(n, np.nan)
            self.atr = np.full(n, np.nan)

        vol = pd.Series(v).replace(0, np.nan)
        self.rvol = (vol / vol.shift(1).rolling(20, min_periods=5).mean()).values

        dates = df.index.normalize()
        self.prev_close = {}
        last_by_date = df["close"].groupby(dates).last()
        for i, d in enumerate(last_by_date.index):
            if i:
                self.prev_close[d] = float(last_by_date.iloc[i - 1])

        # Precompute daily pivots (Floor & Camarilla)
        self.pivots = {}
        d_df = df.groupby(dates).agg({'open': 'first', 'high': 'max', 'low': 'min', 'close': 'last'})
        d_list = list(d_df.index)
        for d_idx in range(1, len(d_list)):
            prev_d = d_list[d_idx - 1]
            curr_d = d_list[d_idx]
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
            self.pivots[curr_d] = {
                'pdh': pdh, 'pdl': pdl, 'pdc': pdc,
                'P': P, 'R1': R1, 'S1': S1, 'R2': R2, 'S2': S2,
                'H4': H4, 'H3': H3, 'L3': L3, 'L4': L4
            }

        # Precompute Time-of-Day Mean Volume (across 20 trading days)
        self.tod_mean = {}
        try:
            df_tod = df[['volume']].copy()
            df_tod['tod'] = df.index.time
            df_tod['d_norm'] = dates
            pv_vol = df_tod.pivot(index='d_norm', columns='tod', values='volume')
            pv_mean = pv_vol.shift(1).rolling(20, min_periods=5).mean()
            for d_val, row in pv_mean.iterrows():
                for t_val, m_vol in row.items():
                    if not np.isnan(m_vol) and m_vol > 0:
                        self.tod_mean[(d_val, t_val)] = float(m_vol)
        except Exception:
            pass

        # Precompute Causal Intraday Session VWAP, VWAP Std Dev, and Incremental 10m Cross & RSI14 in O(N)
        self.vwap = np.zeros(n, dtype=np.float64)
        self.vwap_std = np.zeros(n, dtype=np.float64)
        self.c10_cross = [""] * n
        self.c10_rsi = np.full(n, np.nan)
        self.c10_diff = np.full(n, np.nan)

        tp = (h + l + c) / 3.0
        k10 = 2.0 / (C.EMA_FAST + 1.0)
        k30 = 2.0 / (C.EMA_SLOW + 1.0)
        comp_closes = []
        state_cache = None

        j = 0
        while j < n:
            d = dates[j]
            k = j
            while k < n and dates[k] == d:
                k += 1
            cum_pv = 0.0
            cum_pv2 = 0.0
            cum_v = 0.0
            for idx_in_day, m in enumerate(range(j, k)):
                vol_m = max(v[m], 1.0)
                cum_pv += tp[m] * vol_m
                cum_pv2 += (tp[m] ** 2) * vol_m
                cum_v += vol_m
                vwap_val = cum_pv / cum_v
                self.vwap[m] = vwap_val
                var_val = max(0.0, (cum_pv2 / cum_v) - (vwap_val ** 2))
                self.vwap_std[m] = math.sqrt(var_val)

                cur_px = c[m]
                total_10m = len(comp_closes) + 1
                if total_10m >= C.EMA_SLOW + 2 and state_cache is not None:
                    f_prev, s_prev, avg_g, avg_l, prev_close_10m = state_cache
                    f_cur = (cur_px - f_prev) * k10 + f_prev
                    s_cur = (cur_px - s_prev) * k30 + s_prev
                    prev_diff = f_prev - s_prev
                    cur_diff = f_cur - s_cur
                    self.c10_cross[m] = _cross(prev_diff, cur_diff)
                    self.c10_diff[m] = cur_diff
                    delta = cur_px - prev_close_10m
                    g = delta if delta > 0 else 0.0
                    l_val = -delta if delta < 0 else 0.0
                    ag = (avg_g * 13.0 + g) / 14.0
                    al = (avg_l * 13.0 + l_val) / 14.0
                    self.c10_rsi[m] = 100.0 if al == 0 else (100.0 - 100.0 / (1.0 + ag / al))

                if idx_in_day % 2 == 1:
                    comp_closes.append(cur_px)
                    if len(comp_closes) == C.EMA_SLOW + 1:
                        arr = np.asarray(comp_closes, dtype=np.float64)
                        f_arr = ema_fast(arr, C.EMA_FAST)
                        s_arr = ema_fast(arr, C.EMA_SLOW)
                        diffs = np.diff(arr)
                        gains = np.where(diffs > 0, diffs, 0.0)
                        losses = np.where(diffs < 0, -diffs, 0.0)
                        ag = float(np.mean(gains[:14]))
                        al = float(np.mean(losses[:14]))
                        for idx_r in range(14, len(diffs)):
                            ag = (ag * 13.0 + float(gains[idx_r])) / 14.0
                            al = (al * 13.0 + float(losses[idx_r])) / 14.0
                        state_cache = (float(f_arr[-1]), float(s_arr[-1]), ag, al, cur_px)
                    elif len(comp_closes) > C.EMA_SLOW + 1 and state_cache is not None:
                        f_prev, s_prev, avg_g, avg_l, prev_close_10m = state_cache
                        f_cur = (cur_px - f_prev) * k10 + f_prev
                        s_cur = (cur_px - s_prev) * k30 + s_prev
                        delta = cur_px - prev_close_10m
                        g = delta if delta > 0 else 0.0
                        l_val = -delta if delta < 0 else 0.0
                        ag = (avg_g * 13.0 + g) / 14.0
                        al = (al * 13.0 + l_val) / 14.0
                        state_cache = (f_cur, s_cur, ag, al, cur_px)
            j = k

    def tenmin_state(self, i: int):
        return self.c10_cross[i], float(self.c10_rsi[i])

    def cross5m(self, i: int) -> str:
        if i < 1:
            return ""
        f, s = self.ema[C.EMA_FAST], self.ema[C.EMA_SLOW]
        return _cross(f[i - 1] - s[i - 1], f[i] - s[i])

    def pullback_limit(self, i: int, side: str, ltp: float):
        best, period = 0.0, 0
        for p in C.PULLBACK_EMAS:
            e = float(self.ema[p][i])
            if math.isnan(e) or e <= 0:
                continue
            if side == "BUY" and e < ltp and e > best:
                best, period = e, p
            if side == "SELL" and e > ltp and (best == 0 or e < best):
                best, period = e, p
        return round(best, 2), period


# ============================================================ simulator
class Simulator:
    def __init__(self, mode: str, session_date: datetime, strategy: str = C.DEFAULT_STRATEGY, starting_capital: float = C.STARTING_CAPITAL):
        self.mode = mode
        self.session_date = pd.Timestamp(session_date).normalize()
        self.strategy = strategy if strategy in C.STRATEGY_LABELS else C.DEFAULT_STRATEGY
        self.starting_capital = float(starting_capital)
        self.lock = threading.RLock()
        self.syms: Dict[str, SymbolData] = {}
        self.indices: Dict[str, pd.DataFrame] = {}
        self.nifty_pos: Dict[pd.Timestamp, int] = {}
        self.nifty_e10 = np.array([])
        self.nifty_e21 = np.array([])
        self.nifty_e30 = np.array([])
        self.fix_a: bool = True
        self.fix_b: bool = True
        self.fix_c: bool = True
        self.fast_backtest: bool = False
        self.reset_state()

    def reset_state(self, session_date: Optional[datetime] = None, strategy: Optional[str] = None, capital: Optional[float] = None):
        with self.lock:
            if session_date is not None:
                self.session_date = pd.Timestamp(session_date).normalize()
            if strategy is not None and strategy in C.STRATEGY_LABELS:
                self.strategy = strategy
            if capital is not None:
                self.starting_capital = float(capital)
            self.processed: set = set()
            self.clock: Optional[pd.Timestamp] = None
            self.status = "STARTING"
            self.data_note = ""
            self.cash_realized = 0.0
            self.open: Dict[str, dict] = {}
            self.pending: Dict[str, dict] = {}
            self.closed: List[dict] = []
            self.signals: List[dict] = []
            self.equity: List[dict] = []
            self.traded_symbols_today: set = set()
            self.fees = 0.0
            self.seq = 0
            self.session_comparison: dict = {}

    def load(self, data: Dict[str, pd.DataFrame], universe: List[dict], index_data: Dict[str, pd.DataFrame]):
        ymap = {u["yahoo"]: u for u in universe}
        new_syms = {}
        for ys, df in data.items():
            u = ymap.get(ys)
            if u is None:
                continue
            new_syms[u["symbol"]] = SymbolData(u["symbol"], u["sector"], df)
        new_indices = dict(index_data)
        npos, ne10, ne21, ne30 = None, None, None, None
        nc, nd_open, nd_high, nd_low, nd_cnt = None, None, None, None, None
        ndf = new_indices.get("NIFTY 50")
        if ndf is not None and len(ndf) > 35:
            nc = ndf["close"].values.astype(np.float64)
            no = ndf["open"].values.astype(np.float64)
            nh = ndf["high"].values.astype(np.float64)
            nl = ndf["low"].values.astype(np.float64)
            npos = {ts: i for i, ts in enumerate(ndf.index)}
            ne10 = ema_fast(nc, 10)
            ne21 = ema_fast(nc, 21)
            ne30 = ema_fast(nc, 30)
            n_len = len(ndf)
            nd_open = np.empty(n_len, dtype=np.float64)
            nd_high = np.empty(n_len, dtype=np.float64)
            nd_low = np.empty(n_len, dtype=np.float64)
            nd_cnt = np.empty(n_len, dtype=np.int32)
            cur_d = None
            ro = rh = rl = 0.0
            rc = 0
            for idx_i, ts in enumerate(ndf.index):
                d_key = ts.date()
                if d_key != cur_d:
                    cur_d = d_key
                    ro = no[idx_i]
                    rh = nh[idx_i]
                    rl = nl[idx_i]
                    rc = 1
                else:
                    if nh[idx_i] > rh:
                        rh = nh[idx_i]
                    if nl[idx_i] < rl:
                        rl = nl[idx_i]
                    rc += 1
                nd_open[idx_i] = ro
                nd_high[idx_i] = rh
                nd_low[idx_i] = rl
                nd_cnt[idx_i] = rc

        # Precompute O(1) session bar lookup by date
        bars_by_date: Dict[pd.Timestamp, set] = {}
        for sd in new_syms.values():
            for ts in sd.df.index:
                dn = ts.normalize()
                if dn not in bars_by_date:
                    bars_by_date[dn] = set()
                bars_by_date[dn].add(ts)
        session_bars_map = {dn: sorted(tset) for dn, tset in bars_by_date.items()}

        with self.lock:
            self.syms = new_syms
            self.indices = new_indices
            self._session_bars_map = session_bars_map
            if npos is not None:
                self.nifty_pos = npos
                self.nifty_c = nc
                self.nifty_e10 = ne10
                self.nifty_e21 = ne21
                self.nifty_e30 = ne30
                self.nifty_day_open = nd_open
                self.nifty_day_high = nd_high
                self.nifty_day_low = nd_low
                self.nifty_day_count = nd_cnt

    def available_dates(self) -> List[str]:
        with self.lock:
            if getattr(self, "_available_dates_cache", None):
                return self._available_dates_cache
            idx_df = self.indices.get("NIFTY 50")
            if idx_df is not None:
                dates = sorted({ts.normalize() for ts in idx_df.index})
            elif self.syms:
                first_sd = next(iter(self.syms.values()))
                dates = sorted({ts.normalize() for ts in first_sd.df.index})
            else:
                dates = []
            res = [str(d.date()) for d in dates[3:]] if len(dates) > 3 else [str(d.date()) for d in dates]
            if res:
                self._available_dates_cache = res
            return res

    def session_bar_times(self, target_date: Optional[pd.Timestamp] = None) -> List[pd.Timestamp]:
        d_norm = (target_date or self.session_date).normalize()
        smap = getattr(self, "_session_bars_map", None)
        if smap is not None:
            return smap.get(d_norm, [])
        times = set()
        for sd in self.syms.values():
            d = sd.df.index
            times.update(d[d.normalize() == d_norm])
        return sorted(times)

    def nifty_regime(self, t: pd.Timestamp) -> str:
        ndf = self.indices.get("NIFTY 50")
        if ndf is None or t not in self.nifty_pos:
            return "NEUTRAL"
        i = self.nifty_pos[t]
        nc_arr = getattr(self, "nifty_c", None)
        cl = float(nc_arr[i]) if nc_arr is not None else float(ndf["close"].iloc[i])
        e10, e21, e30 = float(self.nifty_e10[i]), float(self.nifty_e21[i]), float(self.nifty_e30[i])
        if cl > e21 and e10 >= e30:
            return "BULL"
        if cl < e21 and e10 <= e30:
            return "BEAR"
        return "NEUTRAL"

    # ------------------------------------------------ core step
    def step(self, t: pd.Timestamp):
        """Process the 5m bar that STARTS at t (decision time = t + 5m)."""
        with self.lock:
            if t in self.processed:
                return
            self.processed.add(t)
            decide = (t + BAR).time()
            self.clock = t + BAR

            if self.strategy == "v6_baseline":
                # Original V6 order (_manage_pending then _manage_open)
                self._manage_pending_baseline(t)
                self._manage_open_baseline(t, decide)
                if C.ENTRY_START <= decide < C.ENTRY_CUTOFF:
                    self._scan_baseline(t, decide)
            elif self.strategy == "v6_fixes_1_to_5":
                # Fix #1: _manage_open BEFORE _manage_pending + pre-fill contra validation
                self._manage_open_fixes15(t, decide)
                self._manage_pending_fixes15(t)
                if C.ENTRY_START <= decide < C.ENTRY_CUTOFF:
                    self._scan_fixes15(t, decide)
            else:
                # Default: v6_profit_max (Mathematical Net-Profit Maximizer V6++)
                self._manage_open_profit_max(t, decide)
                self._manage_pending_profit_max(t)
                if C.ENTRY_START <= decide < C.OPT_ENTRY_CUTOFF:
                    self._scan_profit_max(t, decide)

            self._mark_equity(t)

    def _bar(self, sym: str, t):
        sd = self.syms.get(sym)
        if sd is None or t not in sd.pos:
            return None, None
        return sd, sd.pos[t]

    # =========================================================================
    # STRATEGY 1: V6++ MATHEMATICAL NET-PROFIT MAXIMIZER (DEFAULT)
    # =========================================================================
    def _manage_pending_profit_max(self, t):
        for sym in list(self.pending):
            o = self.pending[sym]
            sd, i = self._bar(sym, t)
            if sd is None or t <= o["signal_bar"]:
                continue
            lo, hi, op, cl = float(sd.l[i]), float(sd.h[i]), float(sd.o[i]), float(sd.c[i])

            # Fix #1: Immediate Pre-Fill Contra Validation (cancel resting limit if 10m/5m cross or state flips contra)
            c10, _ = sd.tenmin_state(i)
            c5 = sd.cross5m(i)
            c10_diff = float(sd.c10_diff[i]) if not math.isnan(sd.c10_diff[i]) else 0.0
            if (o["side"] == "BUY" and (c10 == "BEARISH" or c5 == "BEARISH" or c10_diff <= 0)) or \
               (o["side"] == "SELL" and (c10 == "BULLISH" or c5 == "BULLISH" or c10_diff >= 0)):
                del self.pending[sym]
                self._log_signal(str(t + BAR), sym, o["side"], cl, o["rsi"], o["adx"], o["rvol"],
                                 "CANCELLED", "Fix #1: Pre-fill contra flip invalidated setup (slot freed immediately)", t)
                continue

            filled = None
            if o["side"] == "BUY" and lo <= o["limit"]:
                filled = min(o["limit"], op)
            elif o["side"] == "SELL" and hi >= o["limit"]:
                filled = max(o["limit"], op)

            if filled is not None:
                vwap_val = float(sd.vwap[i])
                e30 = float(sd.ema[30][i])
                if o["side"] == "BUY":
                    if cl < vwap_val or cl < op or cl < filled * 0.9985 or cl < e30:
                        del self.pending[sym]
                        self._log_signal(str(t + BAR), sym, o["side"], cl, o["rsi"], o["adx"], o["rvol"],
                                         "CANCELLED", "Fix #2: No bounce confirmation (closed red or below VWAP/EMA30)", t)
                        continue
                else:
                    if cl > vwap_val or cl > op or cl > filled * 1.0015 or cl > e30:
                        del self.pending[sym]
                        self._log_signal(str(t + BAR), sym, o["side"], cl, o["rsi"], o["adx"], o["rvol"],
                                         "CANCELLED", "Fix #2: No bounce confirmation (closed green or above VWAP/EMA30)", t)
                        continue

                del self.pending[sym]
                curr_d = t.normalize()
                pivots = sd.pivots.get(curr_d)

                # Fix C: Dynamic Volatility-Parity ATR Stop Distance (vs Fixed 0.85%)
                if self.fix_c:
                    atr_v = float(sd.atr[i]) if not math.isnan(sd.atr[i]) else (filled * 0.003)
                    atr_p = atr_v / filled if filled > 0 else 0.003
                    sl_pct = min(0.0115, max(0.0065, 2.25 * atr_p))
                else:
                    sl_pct = 0.0085
                sl_px = round(filled * (1 - sl_pct) if o["side"] == "BUY" else filled * (1 + sl_pct), 2)

                # Target selection: Nearest structural level >= min_tgt_dist
                min_tgt_dist = 0.0110 if self.fix_a else 0.0075
                fb_tgt_mult = 0.0140 if self.fix_a else 0.0120
                if self.fix_c:
                    min_tgt_dist = max(min_tgt_dist, round(sl_pct * 1.35, 4))
                tgt_name = "R2/H4"
                tgt_px = round(filled * (1 + fb_tgt_mult) if o["side"] == "BUY" else filled * (1 - fb_tgt_mult), 2)
                if pivots is not None:
                    if o["side"] == "BUY":
                        candidates = [("H3", pivots['H3']), ("R1", pivots['R1']), ("PDH", pivots['pdh']), ("H4", pivots['H4']), ("R2", pivots['R2'])]
                        valid = [(name, px) for name, px in sorted(candidates, key=lambda x: x[1]) if (px - filled) / filled >= min_tgt_dist]
                        if valid:
                            tgt_name, tgt_px = valid[0][0], round(valid[0][1], 2)
                        else:
                            tgt_name, tgt_px = "EXT_R2", round(max(pivots['R2'], filled * (1 + fb_tgt_mult)), 2)
                    else:
                        candidates = [("L3", pivots['L3']), ("S1", pivots['S1']), ("PDL", pivots['pdl']), ("L4", pivots['L4']), ("S2", pivots['S2'])]
                        valid = [(name, px) for name, px in sorted(candidates, key=lambda x: -x[1]) if (filled - px) / filled >= min_tgt_dist]
                        if valid:
                            tgt_name, tgt_px = valid[0][0], round(valid[0][1], 2)
                        else:
                            tgt_name, tgt_px = "EXT_S2", round(min(pivots['S2'], filled * (1 - fb_tgt_mult)), 2)

                tgt_pct = abs(tgt_px - filled) / filled
                pos = dict(o, entry=round(filled, 2), entry_time=str(t + BAR), bars_held=0,
                           sl=sl_px, initial_sl=sl_px, sl_pct=round(sl_pct * 100, 2), tgt=tgt_px, tgt_pct=round(tgt_pct * 100, 2),
                           tgt_name=tgt_name, ltp=cl, mfe=0.0, mae=0.0, sl_state=f"INITIAL_{sl_pct*100:.2f}PCT",
                           highest_px=filled, lowest_px=filled)
                self.open[sym] = pos
                self.traded_symbols_today.add(sym)
                if (pos["side"] == "BUY" and lo <= pos["sl"]) or (pos["side"] == "SELL" and hi >= pos["sl"]):
                    self._close(sym, pos["sl"], t, "STOP_LOSS (same bar as fill)", market=False)
                continue

            o["bars_waited"] += 1
            if o["bars_waited"] >= C.ENTRY_LIMIT_WAIT_BARS:
                del self.pending[sym]
                self._log_signal(str(t + BAR), sym, o["side"], o["signal_ltp"], o["rsi"], o["adx"], o["rvol"],
                                 "CANCELLED", f"limit {o['limit']} untouched in {C.ENTRY_LIMIT_WAIT_BARS} bars", t)

    def _manage_open_profit_max(self, t, decide: dtime):
        n_reg = self.nifty_regime(t)
        is_flat_market = (n_reg == "NEUTRAL")
        mins_left = max(1, ((15 * 60 + 5) - (decide.hour * 60 + decide.minute)))
        bars_left = mins_left // 5
        for sym in list(self.open):
            p = self.open[sym]
            sd, i = self._bar(sym, t)
            if sd is None:
                if decide >= C.FORCE_LIQUIDATION:
                    self._close(sym, p["ltp"], t, "3_PM_EXIT (Session Liquidation)", market=not self.fix_a)
                continue
            hi, lo, cl, op = float(sd.h[i]), float(sd.l[i]), float(sd.c[i]), float(sd.o[i])
            p["ltp"], p["bars_held"] = cl, p["bars_held"] + 1
            p["highest_px"] = max(p.get("highest_px", p["entry"]), hi)
            p["lowest_px"] = min(p.get("lowest_px", p["entry"]), lo)
            sgn = 1 if p["side"] == "BUY" else -1
            bar_mfe = sgn * ((hi if sgn > 0 else lo) / p["entry"] - 1) * 100
            bar_mae = sgn * ((lo if sgn > 0 else hi) / p["entry"] - 1) * 100

            if decide >= C.FORCE_LIQUIDATION:
                p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                self._close(sym, cl, t, "3_PM_EXIT (Session Liquidation)", market=not self.fix_a); continue

            # Late-Session Zero-Slippage Limit Target Harvest
            if self.fix_a:
                if bars_left <= 6 and p["mfe"] >= 0.45:
                    harvest_tgt = round(p["entry"] * (1.0 + sgn * 0.0065), 2)
                    if (sgn > 0 and harvest_tgt < p["tgt"] and harvest_tgt > op) or \
                       (sgn < 0 and harvest_tgt > p["tgt"] and harvest_tgt < op):
                        p["tgt"] = harvest_tgt
            else:
                if bars_left <= 10 and p["mfe"] >= 0.35:
                    harvest_tgt = round(p["entry"] * (1.0 + sgn * 0.0055), 2)
                    if (sgn > 0 and harvest_tgt < p["tgt"] and harvest_tgt > op) or \
                       (sgn < 0 and harvest_tgt > p["tgt"] and harvest_tgt < op):
                        p["tgt"] = harvest_tgt

            if p["side"] == "BUY":
                if lo <= p["sl"]:
                    p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                    lbl = f"PROFIT_LOCK ({p['sl_state']})" if not p["sl_state"].startswith("INITIAL") else "STOP_LOSS (0.85%)"
                    self._close(sym, min(p["sl"], op), t, lbl, market=False); continue
                if hi >= p["tgt"]:
                    p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                    self._close(sym, max(p["tgt"], op), t, f"TARGET ({p.get('tgt_name', 'R1/H4')} +{p.get('tgt_pct', 0.8)}%)", market=False); continue
            else:
                if hi >= p["sl"]:
                    p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                    lbl = f"PROFIT_LOCK ({p['sl_state']})" if not p["sl_state"].startswith("INITIAL") else "STOP_LOSS (0.85%)"
                    self._close(sym, max(p["sl"], op), t, lbl, market=False); continue
                if lo <= p["tgt"]:
                    p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                    self._close(sym, min(p["tgt"], op), t, f"TARGET ({p.get('tgt_name', 'S1/L4')} +{p.get('tgt_pct', 0.8)}%)", market=False); continue

            p["mfe"] = max(p["mfe"], bar_mfe)
            p["mae"] = min(p["mae"], bar_mae)

            # 1. Statistical Exhaustion Harvest (VWAP +/- Sigma Rejection)
            vwap_val = float(sd.vwap[i])
            vwap_std_val = float(sd.vwap_std[i]) if hasattr(sd, "vwap_std") else 0.0
            sigma_mult = 2.25 if self.fix_a else 2.0
            exh_mfe_min = 0.85 if self.fix_a else 0.70
            vwap_top = vwap_val + sigma_mult * vwap_std_val
            vwap_bot = vwap_val - sigma_mult * vwap_std_val

            if p["bars_held"] >= 2 and p["mfe"] >= exh_mfe_min:
                if p["side"] == "BUY" and hi >= vwap_top and cl < (hi + lo) / 2.0:
                    self._close(sym, cl, t, "EXHAUSTION (VWAP +2.0 Sigma Rejection)", market=not self.fix_a); continue
                elif p["side"] == "SELL" and lo <= vwap_bot and cl > (hi + lo) / 2.0:
                    self._close(sym, cl, t, "EXHAUSTION (VWAP -2.0 Sigma Rejection)", market=not self.fix_a); continue

            # 2. Fix #3 / Fix A: 2-Stage Tax-Shield Breakeven + Structural Swing Trail
            if self.fix_a:
                # Stage 1: Breathing room up to +0.55% MFE before locking breakeven
                if p["bars_held"] >= 3 and p["mfe"] >= 0.55:
                    pos_val = p["entry"] * p["qty"]
                    net_buf = pos_val * (0.0004 if is_flat_market else 0.0008)
                    be_px = breakeven_exit_price(p["side"], p["qty"], p["entry"], market_exit=False, min_net_rupees=net_buf)
                    if (sgn > 0 and p["sl"] < be_px < cl) or (sgn < 0 and p["sl"] > be_px > cl):
                        p["sl"] = be_px
                        p["sl_state"] = "TAX_SHIELD_BE"
                # Stage 2: Activate 3-bar structural swing trail only once wave reaches +0.80% MFE
                if p["bars_held"] >= 4 and p["mfe"] >= 0.80:
                    k = 3
                    if p["side"] == "BUY":
                        sw_px = round(float(np.min(sd.l[max(0, i-k):i])) * 0.9992, 2)
                        if p["mfe"] >= 1.10:
                            ratchet_px = round(p["entry"] * (1.0 + (p["mfe"] * 0.55) / 100.0), 2)
                            sw_px = max(sw_px, ratchet_px)
                        if p["sl"] < sw_px < cl:
                            p["sl"] = sw_px
                            p["sl_state"] = "SWING_LOW_TRAIL"
                    else:
                        sw_px = round(float(np.max(sd.h[max(0, i-k):i])) * 1.0008, 2)
                        if p["mfe"] >= 1.10:
                            ratchet_px = round(p["entry"] * (1.0 - (p["mfe"] * 0.55) / 100.0), 2)
                            sw_px = min(sw_px, ratchet_px)
                        if p["sl"] > sw_px > cl:
                            p["sl"] = sw_px
                            p["sl_state"] = "SWING_HIGH_TRAIL"
            else:
                if p["bars_held"] >= 3 and p["mfe"] >= 0.35:
                    pos_val = p["entry"] * p["qty"]
                    net_buf = pos_val * (0.0002 if is_flat_market else 0.0005)
                    be_px = breakeven_exit_price(p["side"], p["qty"], p["entry"], market_exit=False, min_net_rupees=net_buf)
                    if (sgn > 0 and p["sl"] < be_px < cl) or (sgn < 0 and p["sl"] > be_px > cl):
                        p["sl"] = be_px
                        p["sl_state"] = "TAX_SHIELD_BE"

                    k = 2
                    if p["side"] == "BUY":
                        sw_px = round(float(np.min(sd.l[max(0, i-k):i])) * 0.9995, 2)
                        if p["sl"] < sw_px < cl:
                            p["sl"] = sw_px
                            p["sl_state"] = "SWING_LOW_TRAIL"
                    else:
                        sw_px = round(float(np.max(sd.h[max(0, i-k):i])) * 1.0005, 2)
                        if p["sl"] > sw_px > cl:
                            p["sl"] = sw_px
                            p["sl_state"] = "SWING_HIGH_TRAIL"

            # 3. Fix #6 / Fix C: Hybrid Multi-Factor Structural Reversal Exit
            c10, _ = sd.tenmin_state(i)
            c5 = sd.cross5m(i)
            c10_diff = float(sd.c10_diff[i]) if not math.isnan(sd.c10_diff[i]) else 0.0
            e30 = float(sd.ema[30][i])
            if self.fix_c:
                atr_v = float(sd.atr[i]) if not math.isnan(sd.atr[i]) else 0.0
                rev_buf = 0.35 * atr_v
                if p["bars_held"] >= 4:
                    if p["side"] == "BUY" and (c10 == "BEARISH" or c10_diff < 0) and c5 == "BEARISH" and cl < (min(vwap_val, e30) - rev_buf):
                        self._close(sym, cl, t, "STRUCTURAL_REVERSAL (Contra+VWAP+EMA30)", market=False); continue
                    if p["side"] == "SELL" and (c10 == "BULLISH" or c10_diff > 0) and c5 == "BULLISH" and cl > (max(vwap_val, e30) + rev_buf):
                        self._close(sym, cl, t, "STRUCTURAL_REVERSAL (Contra+VWAP+EMA30)", market=False); continue
            else:
                if p["bars_held"] >= 3:
                    if p["side"] == "BUY" and (c10 == "BEARISH" or c10_diff < 0 or c5 == "BEARISH") and cl < vwap_val and cl < e30:
                        self._close(sym, cl, t, "STRUCTURAL_REVERSAL (Contra+VWAP+EMA30)", market=True); continue
                    if p["side"] == "SELL" and (c10 == "BULLISH" or c10_diff > 0 or c5 == "BULLISH") and cl > vwap_val and cl > e30:
                        self._close(sym, cl, t, "STRUCTURAL_REVERSAL (Contra+VWAP+EMA30)", market=True); continue

    def _scan_profit_max(self, t, decide: dtime):
        ndf = self.indices.get("NIFTY 50")
        is_nifty_choppy = True
        n_rs_ratio = 1.0
        n_15m_ret = 0.0
        n_day_cnt = 0
        n_open_val = 0.0
        n_cl = 0.0
        n_e10 = 0.0
        n_e30 = 0.0

        if ndf is not None and t in self.nifty_pos:
            ni = self.nifty_pos[t]
            nc_arr = getattr(self, "nifty_c", None)
            n_cl = float(nc_arr[ni]) if nc_arr is not None else float(ndf["close"].iloc[ni])
            n_e10 = float(self.nifty_e10[ni])
            n_e30 = float(self.nifty_e30[ni])
            n_rs_ratio = (n_cl / n_e10) if n_e10 > 0 else 1.0
            if ni >= 3:
                n_prev3 = float(nc_arr[ni - 3]) if nc_arr is not None else float(ndf["close"].iloc[ni - 3])
                n_15m_ret = (n_cl - n_prev3) / n_prev3 * 100.0

            nd_cnt_arr = getattr(self, "nifty_day_count", None)
            if nd_cnt_arr is not None:
                n_day_cnt = int(nd_cnt_arr[ni])
                n_open_val = float(self.nifty_day_open[ni])
                if n_day_cnt >= 5:
                    n_high = float(self.nifty_day_high[ni])
                    n_low = float(self.nifty_day_low[ni])
                    n_ret = (n_cl - n_open_val) / n_open_val * 100.0
                    n_rng = (n_high - n_low) / n_open_val * 100.0
                    is_nifty_choppy = (abs(n_ret) < 0.35 and n_rng < 0.55) or (abs(n_e10 - n_e30) / n_e30 < 0.0005)
            else:
                d_norm = t.normalize()
                d_bars = ndf.loc[ndf.index.normalize() == d_norm]
                d_bars_up_to_t = d_bars.loc[d_bars.index <= t]
                n_day_cnt = len(d_bars_up_to_t)
                if n_day_cnt >= 1:
                    n_open_val = float(d_bars_up_to_t["open"].iloc[0])
                if n_day_cnt >= 5:
                    n_high = float(d_bars_up_to_t["high"].max())
                    n_low = float(d_bars_up_to_t["low"].min())
                    n_ret = (n_cl - n_open_val) / n_open_val * 100.0
                    n_rng = (n_high - n_low) / n_open_val * 100.0
                    is_nifty_choppy = (abs(n_ret) < 0.35 and n_rng < 0.55) or (abs(n_e10 - n_e30) / n_e30 < 0.0005)

        n_reg = self.nifty_regime(t)
        cands = []
        for sym, sd in self.syms.items():
            if t not in sd.pos:
                continue
            i = sd.pos[t]
            if i < 2:
                continue
            c10, rsi10 = sd.tenmin_state(i)
            ltp = float(sd.c[i])
            adx = float(sd.adx[i])
            prev_adx = float(sd.adx[i - 1])
            prev2_adx = float(sd.adx[i - 2]) if i >= 2 else np.nan
            vwap_val = float(sd.vwap[i])
            pdi, mdi = float(sd.plus_di[i]), float(sd.minus_di[i])
            e10, e30 = float(sd.ema[10][i]), float(sd.ema[30][i])
            atr_val = float(sd.atr[i]) if not math.isnan(sd.atr[i]) else 0.0
            atr_pct = (atr_val / ltp * 100.0) if ltp > 0 else 0.0

            cur_vol = float(sd.v[i])
            rolling_rvol = float(sd.rvol[i]) if not math.isnan(sd.rvol[i]) else 0.0
            tod_mean = getattr(sd, "tod_mean", {}).get((t.normalize(), t.time()), 0.0)
            eff_rvol = max(rolling_rvol, cur_vol / tod_mean if tod_mean > 0 else rolling_rvol)

            stock_rs_ratio = (ltp / e10) if e10 > 0 else 1.0
            rs_formula_pct = ((stock_rs_ratio / n_rs_ratio) - 1.0) * 100.0

            if sym in self.open or sym in self.pending:
                continue
            # Fix #4b: Per-Stock Cooldown (1 trade per symbol per day)
            elif sym in self.traded_symbols_today:
                continue
            # Fix #5a (Min ATR >= 0.14%) & Fix #4a (2-Bar Smoothed Rising ADX)
            if math.isnan(rsi10) or math.isnan(adx) or math.isnan(prev_adx) or adx < 20.0 or atr_pct < 0.14:
                continue
            elif adx < prev_adx and (math.isnan(prev2_adx) or adx < prev2_adx):
                continue

            # Fix B: Regime-Adaptive Chop & Low-Volatility Filter
            if self.fix_b:
                # Filter out low-ATR sluggish stocks unless RVOL shows institutional expansion
                if atr_pct < 0.18 and eff_rvol < 1.35:
                    continue
                # During Midday Lull (11:00 - 12:15 IST) in choppy/neutral market, require strong RVOL & ADX
                if (is_nifty_choppy or n_reg == "NEUTRAL") and dtime(11, 0) <= decide <= dtime(12, 15):
                    if eff_rvol < 1.45 or adx < 24.0 or abs(rs_formula_pct) < 0.15:
                        continue

            # Check Setup A: Classical 10m Crossover Pullback (Engine 1)
            setup_type = ""
            side = ""
            limit = 0.0
            period = 0

            if c10 in ("BULLISH", "BEARISH"):
                side = "BUY" if c10 == "BULLISH" else "SELL"
                if side == "BUY" and (50.0 <= rsi10 <= 65.0) and ltp > vwap_val and pdi > mdi and e10 >= e30:
                    limit, period = sd.pullback_limit(i, side, ltp)
                    if limit > 0:
                        setup_type = "ENGINE1_CROSS"
                elif side == "SELL" and (35.0 <= rsi10 <= 50.0) and ltp < vwap_val and mdi > pdi and e10 <= e30:
                    limit, period = sd.pullback_limit(i, side, ltp)
                    if limit > 0:
                        setup_type = "ENGINE1_CROSS"

            # Check Setup B: Intraday VWAP Continuation & Momentum Expansion (Engine 2)
            if not setup_type:
                if ltp > vwap_val and e10 > e30 and pdi > mdi and (48.0 <= rsi10 <= 66.0):
                    lo_prev = float(sd.l[i - 1])
                    vwap_prev = float(sd.vwap[i - 1])
                    if lo_prev <= vwap_prev * 1.002 or lo_prev <= float(sd.ema[30][i - 1]):
                        if float(sd.c[i]) > float(sd.o[i]) and float(sd.c[i]) > float(sd.c[i - 1]):
                            side = "BUY"
                            limit = float(sd.c[i])
                            period = 10
                            setup_type = "ENGINE2_VWAP_CONT"
                elif ltp < vwap_val and e10 < e30 and mdi > pdi and (34.0 <= rsi10 <= 52.0):
                    hi_prev = float(sd.h[i - 1])
                    vwap_prev = float(sd.vwap[i - 1])
                    if hi_prev >= vwap_prev * 0.998 or hi_prev >= float(sd.ema[30][i - 1]):
                        if float(sd.c[i]) < float(sd.o[i]) and float(sd.c[i]) < float(sd.c[i - 1]):
                            side = "SELL"
                            limit = float(sd.c[i])
                            period = 10
                            setup_type = "ENGINE2_VWAP_CONT"

            # Fix B: Engine 3 (Morning Institutional ORB-VWAP Drive 09:25 - 10:15 AM)
            if self.fix_b and not setup_type and dtime(9, 25) <= decide <= dtime(10, 15):
                c10_diff_val = float(sd.c10_diff[i]) if not math.isnan(sd.c10_diff[i]) else 0.0
                if eff_rvol >= 1.60 and adx >= 22.0:
                    if ltp > vwap_val and e10 > e30 and pdi > mdi and c10_diff_val > 0 and (52.0 <= rsi10 <= 68.0) and rs_formula_pct >= 0.15:
                        limit, period = sd.pullback_limit(i, "BUY", ltp)
                        if limit > 0:
                            side = "BUY"
                            setup_type = "ENGINE3_ORB_DRIVE"
                    elif ltp < vwap_val and e10 < e30 and mdi > pdi and c10_diff_val < 0 and (32.0 <= rsi10 <= 48.0) and rs_formula_pct <= -0.15:
                        limit, period = sd.pullback_limit(i, "SELL", ltp)
                        if limit > 0:
                            side = "SELL"
                            setup_type = "ENGINE3_ORB_DRIVE"

            if not setup_type or limit <= 0:
                continue

            # True Momentum Gate: Dynamic Morning Noise Filter (09:15 - 09:45 AM)
            if getattr(C, "TRUE_MOMENTUM_GATE_ENABLED", True) and decide < dtime(9, 45):
                # 1. Require strong relative volume (institutional order flow vs retail drift)
                if eff_rvol < getattr(C, "MORNING_RVOL_MIN", 1.25):
                    continue
                # 2. Reject adverse wick rejections (false breakouts)
                o_px = float(sd.o[i]); h_px = float(sd.h[i]); l_px = float(sd.l[i]); c_px = float(sd.c[i])
                rng_px = h_px - l_px if h_px > l_px else 0.01
                adv_wick = (h_px - max(o_px, c_px)) / rng_px if side == "BUY" else (min(o_px, c_px) - l_px) / rng_px
                if adv_wick > getattr(C, "MORNING_MAX_ADVERSE_WICK", 0.35):
                    continue
                # 3. Macro NIFTY directional alignment
                if getattr(C, "MORNING_NIFTY_ALIGNMENT", True) and ndf is not None and n_day_cnt >= 2 and n_open_val > 0:
                    n_ret_m = (n_cl - n_open_val) / n_open_val * 100.0
                    if side == "BUY" and (n_ret_m < -0.10 or n_e10 < n_e30):
                        continue
                    elif side == "SELL" and (n_ret_m > 0.10 or n_e10 > n_e30):
                        continue

            # Fix #7 / Fix B: Catalyst-Aware Nifty Filter with Cross-Sectional Relative Strength Leader Selection
            must_decouple = is_nifty_choppy or (side == "BUY" and n_reg != "BULL") or (side == "SELL" and n_reg != "BEAR")
            if must_decouple:
                min_dec_rvol = 1.35 if self.fix_b else 1.30
                min_dec_rs = 0.15 if self.fix_b else 0.12
                if eff_rvol < min_dec_rvol:
                    continue
                if side == "BUY" and rs_formula_pct < min_dec_rs:
                    continue
                if side == "SELL" and rs_formula_pct > -min_dec_rs:
                    continue
                if side == "BUY" and n_15m_ret < -0.25 and eff_rvol < 1.8:
                    continue
            else:
                if eff_rvol < 0.8:
                    continue

            if self.fix_b:
                conv_score = (eff_rvol * 1.8) + ((adx - 20.0) * 0.12) + (abs(rs_formula_pct) * 2.5) + (min(0.45, atr_pct) * 3.0)
            else:
                conv_score = (eff_rvol * 1.5) + ((adx - 20.0) * 0.1) + (abs(rs_formula_pct) * 2.0)

            cands.append(dict(symbol=sym, side=side, ltp=ltp, rsi=rsi10, adx=adx, rvol=eff_rvol,
                              atr_pct=atr_pct, conv_score=conv_score, setup=setup_type,
                              limit=limit, ema_period=period, sector=sd.sector))

        cands.sort(key=lambda c: (-c["conv_score"], -c["rvol"], c["symbol"]))
        free = C.MAX_CONCURRENT_POSITIONS - len(self.open) - len(self.pending)
        total_equity = self.starting_capital + self.cash_realized
        committed_cash = sum(p["entry"] * p["qty"] for p in self.open.values()) + \
                         sum(o["limit"] * o["qty"] for o in self.pending.values())
        avail_cash = max(0.0, total_equity * C.LEVERAGE - committed_cash)

        min_cash_buffer = max(1000.0, total_equity * 0.05)
        for k, c in enumerate(cands):
            if k >= free or avail_cash < min_cash_buffer:
                self._log_signal(str(t + BAR), c["symbol"], c["side"], c["ltp"], c["rsi"], c["adx"], c["rvol"],
                                 "NO_SLOT", "qualified, but slots/cash currently deployed", t)
                continue
            slot_budget = (total_equity * C.LEVERAGE) / C.MAX_CONCURRENT_POSITIONS
            equal_rem = avail_cash / max(1, (free - k))
            budget = min(avail_cash, slot_budget, equal_rem)
            qty = int(budget // c["limit"])
            if qty <= 0:
                continue
            avail_cash -= qty * c["limit"]
            self.seq += 1
            self.pending[c["symbol"]] = dict(
                id=self.seq, symbol=c["symbol"], sector=c["sector"], side=c["side"], qty=qty,
                limit=c["limit"], ema_period=c["ema_period"], signal_bar=t, signal_time=str(t + BAR),
                signal_ltp=c["ltp"], rsi=c["rsi"], adx=c["adx"], rvol=c["rvol"], bars_waited=0)
            self._log_signal(str(t + BAR), c["symbol"], c["side"], c["ltp"], c["rsi"], c["adx"], c["rvol"],
                             "ORDER", f"limit {c['side']} {qty} @ {c['limit']} (5m EMA{c['ema_period']} | Score {c['conv_score']:.1f})", t)

    # =========================================================================
    # STRATEGY 2: V6+ FIXES #1 TO #5
    # =========================================================================
    def _manage_pending_fixes15(self, t):
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
                    self._log_signal(str(t + BAR), sym, o["side"], cl, o["rsi"], o["adx"], o["rvol"],
                                     "CANCELLED", "Fix #1: Pullback candle broke 10m/5m cross", t)
                    continue
                vwap_val = float(sd.vwap[i])
                if (o["side"] == "BUY" and (cl < vwap_val or cl < filled * 0.9975)) or \
                   (o["side"] == "SELL" and (cl > vwap_val or cl > filled * 1.0025)):
                    del self.pending[sym]
                    self._log_signal(str(t + BAR), sym, o["side"], cl, o["rsi"], o["adx"], o["rvol"],
                                    "CANCELLED", "Fix #2: Failed VWAP/bounce check on fill bar", t)
                    continue
                del self.pending[sym]
                sl_px = round(filled * (1 - C.SL_PCT) if o["side"] == "BUY" else filled * (1 + C.SL_PCT), 2)
                tgt_px = round(filled * (1 + C.TGT_PCT) if o["side"] == "BUY" else filled * (1 - C.TGT_PCT), 2)
                pos = dict(o, entry=round(filled, 2), entry_time=str(t + BAR), bars_held=0,
                           sl=sl_px, initial_sl=sl_px, tgt=tgt_px,
                           ltp=cl, mfe=0.0, mae=0.0, sl_state="INITIAL_1PCT")
                self.open[sym] = pos
                self.traded_symbols_today.add(sym)
                if (pos["side"] == "BUY" and lo <= pos["sl"]) or (pos["side"] == "SELL" and hi >= pos["sl"]):
                    self._close(sym, pos["sl"], t, "STOP_LOSS (same bar as fill)", market=False)
                continue
            o["bars_waited"] += 1
            if o["bars_waited"] >= C.ENTRY_LIMIT_WAIT_BARS:
                del self.pending[sym]
                self._log_signal(str(t + BAR), sym, o["side"], o["signal_ltp"], o["rsi"], o["adx"], o["rvol"],
                                 "CANCELLED", f"limit {o['limit']} untouched in {C.ENTRY_LIMIT_WAIT_BARS} bars", t)

    def _manage_open_fixes15(self, t, decide: dtime):
        for sym in list(self.open):
            p = self.open[sym]
            sd, i = self._bar(sym, t)
            if sd is None:
                if decide >= C.FORCE_LIQUIDATION:
                    self._close(sym, p["ltp"], t, "3_PM_EXIT (no bar)", market=True)
                continue
            hi, lo, cl, op = float(sd.h[i]), float(sd.l[i]), float(sd.c[i]), float(sd.o[i])
            p["ltp"], p["bars_held"] = cl, p["bars_held"] + 1
            sgn = 1 if p["side"] == "BUY" else -1
            bar_mfe = sgn * ((hi if sgn > 0 else lo) / p["entry"] - 1) * 100
            bar_mae = sgn * ((lo if sgn > 0 else hi) / p["entry"] - 1) * 100

            if decide >= C.FORCE_LIQUIDATION:
                p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                self._close(sym, cl, t, "3_PM_EXIT (Session Liquidation)", market=True); continue

            if p["side"] == "BUY":
                if lo <= p["sl"]:
                    p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                    lbl = f"PROFIT_LOCK_SL ({p['sl_state']})" if p["sl_state"] != "INITIAL_1PCT" else "STOP_LOSS (1.0%)"
                    self._close(sym, min(p["sl"], op), t, lbl, market=True); continue
                if hi >= p["tgt"]:
                    p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                    self._close(sym, max(p["tgt"], op), t, f"TARGET ({C.TGT_PCT*100:.1f}%)", market=False); continue
            else:
                if hi >= p["sl"]:
                    p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                    lbl = f"PROFIT_LOCK_SL ({p['sl_state']})" if p["sl_state"] != "INITIAL_1PCT" else "STOP_LOSS (1.0%)"
                    self._close(sym, max(p["sl"], op), t, lbl, market=True); continue
                if lo <= p["tgt"]:
                    p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                    self._close(sym, min(p["tgt"], op), t, f"TARGET ({C.TGT_PCT*100:.1f}%)", market=False); continue

            p["mfe"] = max(p["mfe"], bar_mfe)
            p["mae"] = min(p["mae"], bar_mae)

            if p["mfe"] >= 0.65:
                lock_pct = (p["mfe"] * 0.60) / 100.0
                new_sl = round(p["entry"] * (1.0 + sgn * lock_pct), 2)
                if (sgn > 0 and new_sl > p["sl"]) or (sgn < 0 and new_sl < p["sl"]):
                    p["sl"] = new_sl
                    p["sl_state"] = f"TIER2_TRAIL (+{lock_pct*100:.2f}%)"
            elif p["mfe"] >= 0.40:
                be_px = breakeven_exit_price(p["side"], p["qty"], p["entry"], market_exit=True, min_net_rupees=75.0)
                if (sgn > 0 and be_px > p["sl"]) or (sgn < 0 and be_px < p["sl"]):
                    p["sl"] = be_px
                    p["sl_state"] = "TIER1_TAX_SHIELD"

            cur_ret_pct = sgn * (cl / p["entry"] - 1.0) * 100.0
            in_tax_dead_zone = (-0.12 <= cur_ret_pct <= 0.22)
            c10, _ = sd.tenmin_state(i)
            if (p["side"] == "BUY" and c10 == "BEARISH") or (p["side"] == "SELL" and c10 == "BULLISH"):
                self._close(sym, cl, t, "EMA_CROSS_10M_CONTRA", market=True); continue
            if not in_tax_dead_zone:
                adx = float(sd.adx[i])
                e10 = float(sd.ema[10][i])
                against_e10 = (p["side"] == "BUY" and cl < e10) or (p["side"] == "SELL" and cl > e10)
                if not math.isnan(adx) and adx < C.ADX_EXIT_BELOW and against_e10:
                    self._close(sym, cl, t, f"ADX_5M_BELOW_{int(C.ADX_EXIT_BELOW)} ({adx:.1f})", market=True); continue
                c5 = sd.cross5m(i)
                if (p["side"] == "BUY" and c5 == "BEARISH") or (p["side"] == "SELL" and c5 == "BULLISH"):
                    self._close(sym, cl, t, "EMA_CROSS_5M (Reverse Exit)", market=True); continue

    def _scan_fixes15(self, t, decide):
        cands = []
        for sym, sd in self.syms.items():
            if t not in sd.pos:
                continue
            i = sd.pos[t]
            if i < 2:
                continue
            c10, rsi10 = sd.tenmin_state(i)
            if not c10:
                continue
            side = "BUY" if c10 == "BULLISH" else "SELL"
            ltp = float(sd.c[i])
            adx = float(sd.adx[i])
            prev_adx = float(sd.adx[i - 1])
            rvol = float(sd.rvol[i]) if not math.isnan(sd.rvol[i]) else 0.0
            vwap_val = float(sd.vwap[i])
            reason = ""
            if sym in self.open or sym in self.pending:
                reason = "already open/pending"
            elif sym in self.traded_symbols_today:
                reason = "Fix #4: daily cooldown active"
            elif math.isnan(rsi10):
                reason = "10m RSI unavailable"
            elif side == "BUY" and not (48.0 <= rsi10 <= 64.0):
                reason = f"RSI {rsi10:.1f} outside BUY 48-64"
            elif side == "SELL" and not (36.0 <= rsi10 <= 50.0):
                reason = f"RSI {rsi10:.1f} outside SELL 36-50"
            elif math.isnan(adx) or math.isnan(prev_adx) or adx < C.STOCK_ADX_MIN or adx < prev_adx:
                reason = f"ADX {adx:.1f} < 20 or falling"
            elif (side == "BUY" and ltp <= vwap_val) or (side == "SELL" and ltp >= vwap_val):
                reason = "Fix #2: against Intraday VWAP"
            limit, period = sd.pullback_limit(i, side, ltp)
            if not reason and limit <= 0:
                reason = "no pullback EMA"
            if reason:
                self._log_signal(str(t + BAR), sym, side, ltp, rsi10, adx, rvol, "FILTERED", reason, t)
                continue
            cands.append(dict(symbol=sym, side=side, ltp=ltp, rsi=rsi10, adx=adx, rvol=rvol,
                              limit=limit, ema_period=period, sector=sd.sector))
        cands.sort(key=lambda c: (-c["rvol"], -c["adx"], c["symbol"]))
        free = C.MAX_CONCURRENT_POSITIONS - len(self.open) - len(self.pending)
        equity = self.starting_capital + self.cash_realized
        for k, c in enumerate(cands):
            if k >= free:
                self._log_signal(str(t + BAR), c["symbol"], c["side"], c["ltp"], c["rsi"], c["adx"], c["rvol"],
                                 "NO_SLOT", f"all {C.MAX_CONCURRENT_POSITIONS} slots busy", t)
                continue
            budget = equity * C.LEVERAGE / C.MAX_CONCURRENT_POSITIONS
            qty = int(budget // c["limit"])
            if qty <= 0:
                continue
            self.seq += 1
            self.pending[c["symbol"]] = dict(
                id=self.seq, symbol=c["symbol"], sector=c["sector"], side=c["side"], qty=qty,
                limit=c["limit"], ema_period=c["ema_period"], signal_bar=t, signal_time=str(t + BAR),
                signal_ltp=c["ltp"], rsi=c["rsi"], adx=c["adx"], rvol=c["rvol"], bars_waited=0)
            self._log_signal(str(t + BAR), c["symbol"], c["side"], c["ltp"], c["rsi"], c["adx"], c["rvol"],
                             "ORDER", f"limit {c['side']} {qty} @ {c['limit']} (5m EMA{c['ema_period']})", t)

    # =========================================================================
    # STRATEGY 3: DAD'S ORIGINAL V6 BASELINE
    # =========================================================================
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
            if filled is not None:
                del self.pending[sym]
                pos = dict(o, entry=round(filled, 2), entry_time=str(t + BAR), bars_held=0,
                           sl=round(filled * (1 - C.SL_PCT) if o["side"] == "BUY" else filled * (1 + C.SL_PCT), 2),
                           tgt=round(filled * (1 + C.TGT_PCT) if o["side"] == "BUY" else filled * (1 - C.TGT_PCT), 2),
                           ltp=cl, mfe=0.0, mae=0.0, sl_state="FIXED_1PCT")
                self.open[sym] = pos
                if (pos["side"] == "BUY" and lo <= pos["sl"]) or (pos["side"] == "SELL" and hi >= pos["sl"]):
                    self._close(sym, pos["sl"], t, "STOP_LOSS (same bar as fill)", market=False)
                continue
            o["bars_waited"] += 1
            if o["bars_waited"] >= C.ENTRY_LIMIT_WAIT_BARS:
                del self.pending[sym]
                self._log_signal(o["signal_time"], sym, o["side"], o["signal_ltp"], o["rsi"], o["adx"], o["rvol"],
                                 "CANCELLED", f"limit {o['limit']} not touched in {C.ENTRY_LIMIT_WAIT_BARS} bars", t)

    def _manage_open_baseline(self, t, decide: dtime):
        for sym in list(self.open):
            p = self.open[sym]
            sd, i = self._bar(sym, t)
            if sd is None:
                if decide >= C.FORCE_LIQUIDATION:
                    self._close(sym, p["ltp"], t, "3_PM_EXIT (no bar, last price)", market=True)
                continue
            hi, lo, cl, op = float(sd.h[i]), float(sd.l[i]), float(sd.c[i]), float(sd.o[i])
            p["ltp"], p["bars_held"] = cl, p["bars_held"] + 1
            sgn = 1 if p["side"] == "BUY" else -1
            p["mfe"] = max(p["mfe"], sgn * ((hi if sgn > 0 else lo) / p["entry"] - 1) * 100)
            p["mae"] = min(p["mae"], sgn * ((lo if sgn > 0 else hi) / p["entry"] - 1) * 100)
            if decide >= C.FORCE_LIQUIDATION:
                self._close(sym, cl, t, "3_PM_EXIT (Session Liquidation)", market=True); continue
            if p["side"] == "BUY":
                if lo <= p["sl"]:
                    self._close(sym, min(p["sl"], op), t, "STOP_LOSS (1.0%)", market=True); continue
                if hi >= p["tgt"]:
                    self._close(sym, max(p["tgt"], op), t, f"TARGET ({C.TGT_PCT*100:.1f}%)", market=False); continue
            else:
                if hi >= p["sl"]:
                    self._close(sym, max(p["sl"], op), t, "STOP_LOSS (1.0%)", market=True); continue
                if lo <= p["tgt"]:
                    self._close(sym, min(p["tgt"], op), t, f"TARGET ({C.TGT_PCT*100:.1f}%)", market=False); continue
            adx = float(sd.adx[i])
            if not math.isnan(adx) and adx < C.ADX_EXIT_BELOW:
                self._close(sym, cl, t, f"ADX_5M_BELOW_{int(C.ADX_EXIT_BELOW)} ({adx:.1f})", market=True); continue
            c10, _ = sd.tenmin_state(i)
            if (p["side"] == "BUY" and c10 == "BEARISH") or (p["side"] == "SELL" and c10 == "BULLISH"):
                self._close(sym, cl, t, "EMA_CROSS_10M_CONTRA", market=True); continue
            c5 = sd.cross5m(i)
            if (p["side"] == "BUY" and c5 == "BEARISH") or (p["side"] == "SELL" and c5 == "BULLISH"):
                self._close(sym, cl, t, "EMA_CROSS_5M (Reverse Exit)", market=True); continue

    def _scan_baseline(self, t, decide):
        cands = []
        for sym, sd in self.syms.items():
            if t not in sd.pos:
                continue
            i = sd.pos[t]
            c10, rsi10 = sd.tenmin_state(i)
            if not c10:
                continue
            side = "BUY" if c10 == "BULLISH" else "SELL"
            ltp = float(sd.c[i])
            adx = float(sd.adx[i])
            rvol = float(sd.rvol[i]) if not math.isnan(sd.rvol[i]) else 0.0
            reason = ""
            if sym in self.open or sym in self.pending:
                reason = "already has an open trade / pending order"
            elif C.RSI_FILTER_ENABLED and math.isnan(rsi10):
                reason = "10m RSI unavailable"
            elif C.RSI_FILTER_ENABLED and side == "BUY" and not (C.BUY_RSI_MIN <= rsi10 <= C.BUY_RSI_MAX):
                reason = f"RSI {rsi10:.1f} outside BUY {C.BUY_RSI_MIN:.0f}-{C.BUY_RSI_MAX:.0f}"
            elif C.RSI_FILTER_ENABLED and side == "SELL" and not (C.SELL_RSI_MIN <= rsi10 <= C.SELL_RSI_MAX):
                reason = f"RSI {rsi10:.1f} outside SELL {C.SELL_RSI_MIN:.0f}-{C.SELL_RSI_MAX:.0f}"
            elif math.isnan(adx) or adx < C.STOCK_ADX_MIN:
                reason = f"ADX {adx:.1f} < {C.STOCK_ADX_MIN:.0f} (choppy)"
            limit, period = sd.pullback_limit(i, side, ltp)
            if not reason and limit <= 0:
                reason = "no 5m EMA on the pullback side of price"
            if reason:
                self._log_signal(str(t + BAR), sym, side, ltp, rsi10, adx, rvol, "FILTERED", reason, t)
                continue
            cands.append(dict(symbol=sym, side=side, ltp=ltp, rsi=rsi10, adx=adx, rvol=rvol,
                              limit=limit, ema_period=period, sector=sd.sector))
        cands.sort(key=lambda c: (-c["rvol"], -c["adx"], c["symbol"]))
        free = C.MAX_CONCURRENT_POSITIONS - len(self.open) - len(self.pending)
        equity = self.starting_capital + self.cash_realized
        for k, c in enumerate(cands):
            if k >= free:
                self._log_signal(str(t + BAR), c["symbol"], c["side"], c["ltp"], c["rsi"], c["adx"], c["rvol"],
                                 "NO_SLOT", f"qualified, but all {C.MAX_CONCURRENT_POSITIONS} slots busy", t)
                continue
            budget = equity * C.LEVERAGE / C.MAX_CONCURRENT_POSITIONS
            qty = int(budget // c["limit"])
            if qty <= 0:
                continue
            self.seq += 1
            self.pending[c["symbol"]] = dict(
                id=self.seq, symbol=c["symbol"], sector=c["sector"], side=c["side"], qty=qty,
                limit=c["limit"], ema_period=c["ema_period"], signal_bar=t, signal_time=str(t + BAR),
                signal_ltp=c["ltp"], rsi=c["rsi"], adx=c["adx"], rvol=c["rvol"], bars_waited=0)
            self._log_signal(str(t + BAR), c["symbol"], c["side"], c["ltp"], c["rsi"], c["adx"], c["rvol"],
                             "ORDER", f"limit {c['side']} {qty} @ {c['limit']} (5m EMA{c['ema_period']})", t)

    # =========================================================================
    # CLOSING, LOGGING & SNAPSHOT
    # =========================================================================
    def _close(self, sym, px, t, reason, market):
        p = self.open.pop(sym)
        px = round(float(px), 2)
        sgn = 1 if p["side"] == "BUY" else -1
        gross = (px - p["entry"]) * p["qty"] * sgn
        cost = trade_costs(p["side"], p["qty"], p["entry"], px, market)
        net = gross - cost["total"]
        ret_pct = round((px / p["entry"] - 1) * 100 * sgn, 3)
        self.cash_realized += net
        self.fees += cost["total"]

        if p["side"] == "BUY":
            verb = f"LONG BUY {p['qty']} sh @ ₹{p['entry']:.2f} → Sold @ ₹{px:.2f}"
        else:
            verb = f"SHORT SELL {p['qty']} sh @ ₹{p['entry']:.2f} → Covered @ ₹{px:.2f}"
        plain = f"{verb} ({'+' if ret_pct >= 0 else ''}{ret_pct:.2f}%) | Gross {'+' if gross >= 0 else ''}₹{gross:.2f} − Tax/Brk ₹{cost['total']:.2f} = Net {'+' if net >= 0 else ''}₹{net:.2f}"

        self.closed.append(dict(
            id=p["id"], symbol=sym, sector=p["sector"], side=p["side"], qty=p["qty"],
            entry=p["entry"], exit=px, entry_time=p["entry_time"], exit_time=str(t + BAR),
            reason=reason, gross=round(gross, 2), costs=round(cost["total"], 2),
            cost_detail={k: round(v, 2) for k, v in cost.items()}, net=round(net, 2),
            ret_pct=ret_pct, bars_held=p["bars_held"], mfe=round(p["mfe"], 2), mae=round(p["mae"], 2),
            sl_state=p.get("sl_state", "INITIAL"), plain_english=plain
        ))

    def _log_signal(self, when, sym, side, ltp, rsi, adx, rvol, result, detail, t):
        if self.fast_backtest:
            return
        f = lambda v: None if v is None or (isinstance(v, float) and math.isnan(v)) else round(float(v), 2)
        self.signals.append(dict(time=str(when), symbol=sym, side=side, ltp=f(ltp), rsi=f(rsi), adx=f(adx),
                                 rvol=f(rvol), result=result, detail=detail))
        if len(self.signals) > 2000:
            self.signals = self.signals[-2000:]

    def unrealized(self) -> float:
        u = 0.0
        for p in self.open.values():
            u += (p["ltp"] - p["entry"]) * p["qty"] * (1 if p["side"] == "BUY" else -1)
        return u

    def _mark_equity(self, t):
        if self.fast_backtest:
            return
        self.equity.append({"t": str(t + BAR), "realized": round(self.cash_realized, 2),
                            "total": round(self.cash_realized + self.unrealized(), 2)})

    def refresh_ltps(self):
        with self.lock:
            for sym, p in self.open.items():
                sd = self.syms.get(sym)
                if sd is not None and len(sd.df):
                    p["ltp"] = float(sd.df["close"].iloc[-1])

    def compute_session_comparison(self):
        """Runs all 3 strategy modes on self.session_date in memory (~15ms) for the comparison banner."""
        with self.lock:
            times = self.session_bar_times(self.session_date)
            if not times:
                self.session_comparison = {}
                return
            session_date = self.session_date.to_pydatetime()
            starting_capital = self.starting_capital
            syms_ref = self.syms
            indices_ref = self.indices
            npos_ref = self.nifty_pos
            ne10_ref = self.nifty_e10
            ne21_ref = self.nifty_e21
            ne30_ref = self.nifty_e30

        comp = {}
        for strat_key in ["v6_baseline", "quacknutzz_strat"]:
            tmp = Simulator("replay", session_date, strategy=strat_key, starting_capital=starting_capital)
            tmp.syms = syms_ref
            tmp.indices = indices_ref
            tmp.nifty_pos = npos_ref
            tmp.nifty_e10 = ne10_ref
            tmp.nifty_e21 = ne21_ref
            tmp.nifty_e30 = ne30_ref
            for t in times:
                tmp.step(t)
            wins = sum(1 for c in tmp.closed if c["net"] > 0)
            gross = sum(c["gross"] for c in tmp.closed)
            comp[strat_key] = {
                "label": C.STRATEGY_LABELS.get(strat_key, strat_key),
                "trades": len(tmp.closed),
                "wins": wins,
                "win_rate": round(wins / len(tmp.closed) * 100, 1) if tmp.closed else 0.0,
                "gross": round(gross, 2),
                "fees": round(tmp.fees, 2),
                "net": round(tmp.cash_realized, 2),
            }
        comp["v6_profit_max"] = comp.get("quacknutzz_strat", {})
        with self.lock:
            self.session_comparison = comp

    # ------------------------------------------------ dashboard payload
    def snapshot(self) -> dict:
        with self.lock:
            wins = [c for c in self.closed if c["net"] > 0]
            gross_w = sum(c["gross"] for c in self.closed if c["gross"] > 0)
            gross_l = -sum(c["gross"] for c in self.closed if c["gross"] < 0)
            gross_tot = sum(c["gross"] for c in self.closed)
            idx = {}
            for name, df in self.indices.items():
                d = df[df.index <= (self.clock or df.index[-1])] if self.mode == "replay" else df
                if len(d):
                    last = float(d["close"].iloc[-1])
                    prev = d[d.index.normalize() < d.index[-1].normalize()]
                    pc = float(prev["close"].iloc[-1]) if len(prev) else last
                    idx[name] = {"ltp": round(last, 2), "chg": round((last / pc - 1) * 100, 2) if pc else 0}
            watch = []
            for sym, sd in self.syms.items():
                d = sd.df
                if self.mode == "replay" and self.clock is not None:
                    n = int(np.searchsorted(d.index.values, np.datetime64(self.clock - BAR), side="right"))
                else:
                    n = len(d)
                if n == 0:
                    continue
                i = n - 1
                ltp = float(sd.c[i])
                pc = sd.prev_close.get(d.index[i].normalize())
                f, s = sd.ema[C.EMA_FAST][i], sd.ema[C.EMA_SLOW][i]
                status = "OPEN" if sym in self.open else ("PENDING" if sym in self.pending else "")
                rsi_val = None
                if hasattr(sd, "c10_rsi") and i < len(sd.c10_rsi) and not math.isnan(sd.c10_rsi[i]) and sd.c10_rsi[i] > 0:
                    rsi_val = round(float(sd.c10_rsi[i]), 1)
                
                vwap_val = round(float(sd.vwap[i]), 2)
                vwap_dist = round((ltp / vwap_val - 1) * 100, 2) if vwap_val > 0 else 0.0
                atr_val = float(sd.atr[i]) if hasattr(sd, "atr") and not math.isnan(sd.atr[i]) else 0.0
                atr_pct = round(atr_val / ltp * 100, 2) if ltp > 0 else 0.0
                
                e10_val = sd.ema[10][i] if (10 in sd.ema and i < len(sd.ema[10])) else f
                e21_val = sd.ema[21][i] if (21 in sd.ema and i < len(sd.ema[21])) else (f + s) / 2.0
                e30_val = sd.ema[30][i] if (30 in sd.ema and i < len(sd.ema[30])) else s
                if e10_val > e21_val > e30_val:
                    ema_stack = "BULL_STACK"
                elif e10_val < e21_val < e30_val:
                    ema_stack = "BEAR_STACK"
                else:
                    ema_stack = "UP" if f > s else "DOWN"

                # Intraday Day High / Day Low Position (0-100%)
                curr_date = d.index[i].normalize()
                day_mask = (d.index.normalize() == curr_date) & (d.index <= d.index[i])
                day_high = float(sd.h[day_mask].max()) if day_mask.any() else ltp
                day_low = float(sd.l[day_mask].min()) if day_mask.any() else ltp
                day_rng = max(0.05, day_high - day_low)
                day_pos = round((ltp - day_low) / day_rng * 100, 1)

                watch.append({"symbol": sym, "sector": sd.sector, "ltp": round(ltp, 2),
                              "vwap": vwap_val, "vwap_dist": vwap_dist,
                              "chg": round((ltp / pc - 1) * 100, 2) if pc else None,
                              "adx": None if math.isnan(sd.adx[i]) else round(float(sd.adx[i]), 1),
                              "rvol": None if math.isnan(sd.rvol[i]) else round(float(sd.rvol[i]), 2),
                              "rsi": rsi_val, "atr_pct": atr_pct, "ema_stack": ema_stack,
                              "day_high": round(day_high, 2), "day_low": round(day_low, 2), "day_pos": day_pos,
                              "trend5m": "UP" if f > s else "DOWN", "status": status})

            open_enriched = []
            for p in self.open.values():
                clean_p = {k: (str(v) if k == "signal_bar" else (None if isinstance(v, float) and math.isnan(v) else v)) for k, v in p.items()}
                sgn = 1 if p["side"] == "BUY" else -1
                unr_gross = (p["ltp"] - p["entry"]) * p["qty"] * sgn
                est_cost = trade_costs(p["side"], p["qty"], p["entry"], p["ltp"], True)["total"]
                clean_p["breakeven_px"] = breakeven_exit_price(p["side"], p["qty"], p["entry"], market_exit=True, min_net_rupees=0.0)
                clean_p["tax_hurdle_rs"] = round(est_cost, 2)
                clean_p["gross_unrealized"] = round(unr_gross, 2)
                clean_p["net_unrealized"] = round(unr_gross - est_cost, 2)
                open_enriched.append(clean_p)

            last_bar = max((sd.df.index[-1] for sd in self.syms.values() if len(sd.df)), default=None)
            return {
                "mode": self.mode,
                "strategy": self.strategy,
                "strategy_label": C.STRATEGY_LABELS.get(self.strategy, self.strategy),
                "strategy_options": C.STRATEGY_LABELS,
                "available_dates": self.available_dates(),
                "session_comparison": self.session_comparison,
                "status": self.status,
                "session": str(self.session_date.date()),
                "starting_capital": self.starting_capital,
                "clock": str(self.clock) if self.clock is not None else None,
                "last_data_bar": str(last_bar) if last_bar is not None else None,
                "data_note": self.data_note,
                "server_time": str(datetime.now().replace(microsecond=0)),
                "config": {"sl_pct": C.SL_PCT, "tgt_pct": C.TGT_PCT, "max_pos": C.MAX_CONCURRENT_POSITIONS,
                           "buy_rsi": [C.BUY_RSI_MIN, C.BUY_RSI_MAX], "sell_rsi": [C.SELL_RSI_MIN, C.SELL_RSI_MAX],
                           "adx_min": C.STOCK_ADX_MIN, "adx_exit": C.ADX_EXIT_BELOW,
                           "capital": self.starting_capital, "leverage": C.LEVERAGE,
                           "window": [str(C.ENTRY_START), str(C.ENTRY_CUTOFF), str(C.FORCE_LIQUIDATION)]},
                "kpi": {"capital": self.starting_capital, "gross": round(gross_tot, 2),
                        "realized": round(self.cash_realized, 2),
                        "unrealized": round(self.unrealized(), 2), "fees": round(self.fees, 2),
                        "trades": len(self.closed), "wins": len(wins),
                        "win_rate": round(len(wins) / len(self.closed) * 100, 1) if self.closed else None,
                        "profit_factor": round(gross_w / gross_l, 2) if gross_l else (99.0 if gross_w > 0 else None),
                        "signals": len(self.signals), "universe": len(self.syms)},
                "indices": idx,
                "open": open_enriched,
                "pending": [{k: (str(v) if k == "signal_bar" else (None if isinstance(v, float) and math.isnan(v) else v)) for k, v in o.items()} for o in self.pending.values()],
                "closed": self.closed[::-1],
                "signals": self.signals[::-1][:400],
                "equity": self.equity,
                "watch": watch,
            }

    def chart(self, sym: str) -> dict:
        with self.lock:
            sd = self.syms.get(sym)
            if sd is None:
                return {}
            d = sd.df
            mask = d.index.normalize() == self.session_date
            if self.mode == "replay" and self.clock is not None:
                mask &= d.index < self.clock
            if not mask.any() and len(d):
                # Fallback to most recent completed session if current session has no bars yet (e.g. before 09:15 AM)
                last_date = d.index[-1].normalize()
                mask = d.index.normalize() == last_date
            ix = np.where(mask)[0]
            s = lambda v: None if v is None or (isinstance(v, float) and math.isnan(v)) else round(float(v), 2)
            bars = []
            for i in ix:
                ts_dt = d.index[i]
                unix_ts = int(ts_dt.timestamp())
                b_item = {
                    "t": str(ts_dt),
                    "ts": unix_ts,
                    "o": s(sd.o[i]),
                    "h": s(sd.h[i]),
                    "l": s(sd.l[i]),
                    "c": s(sd.c[i]),
                    "v": int(sd.v[i]) if hasattr(sd, "v") and not math.isnan(sd.v[i]) else 0,
                    "vwap": s(sd.vwap[i]),
                    "e10": s(sd.ema[10][i]) if (10 in sd.ema and i < len(sd.ema[10])) else None,
                    "e30": s(sd.ema[30][i]) if (30 in sd.ema and i < len(sd.ema[30])) else None,
                    "adx": s(sd.adx[i]) if hasattr(sd, "adx") else None,
                }
                if hasattr(sd, "c10_rsi") and i < len(sd.c10_rsi) and not math.isnan(sd.c10_rsi[i]):
                    b_item["rsi"] = round(float(sd.c10_rsi[i]), 1)
                else:
                    b_item["rsi"] = None
                bars.append(b_item)
            marks = [c for c in self.closed if c["symbol"] == sym]
            op = self.open.get(sym)
            return {"symbol": sym, "bars": bars, "trades": marks,
                    "open": {k: (str(v) if k == "signal_bar" else v) for k, v in op.items()} if op else None,
                    "pending": {k: str(v) for k, v in self.pending.get(sym, {}).items()}}
