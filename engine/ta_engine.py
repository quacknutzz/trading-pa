"""
ta_engine.py -- Back to Basics Trading Engine.

Indicator computation, Excel column mapping, and COM write.
Extracted from excel_io.py + market_analysis.py.
Zero AI/ML dependencies.
"""

import os
import time as _time
import logging
from datetime import datetime, time as dtime, timedelta
from typing import Dict, Any, List, Optional

import numpy as np
import pandas as pd

from engine.config import (
    EXCEL_FILENAME, EXCEL_SHEET, FULL_COLUMNS,
    OUTPUT_START_COL_INDEX, LTP_COL, VWAP_COL, SPIKE_COL,
    SUPERTREND_COL, SECTOR_COL, REGIME_5M_COL, SECTOR_REGIME_COL, RVOL_COL, REGIME_COL,
    ER_COL, SIZE_COL, VOL_RATIO_COL,
    AI_SIGNAL_COL, AI_HEALTH_COL, AI_SETUP_COL,
    SWING_HIGH_COL, SWING_LOW_COL,
    DIST_DAY_HI_COL, DIST_DAY_LO_COL,
    RSI_COLS, MACD_TEXT_COLS, EMA_CROSS_COLS, EMA10_50_CROSS_COLS,
    BULL_PATTERN_COLS, BEAR_PATTERN_COLS,
    ATR_COLS, DI_PLUS_COL, DI_MINUS_COL, ADX_COL,
    STOCH_K_COLS, STOCH_D_COLS,
    BB_UPPER_COLS, BB_MID_COLS, BB_LOWER_COLS,
    FIB_COLS_5M, FIB_COLS_15M,
    SCORE_COL, NARRATIVE_COL,
    TIMEFRAMES, DOJI_THRESHOLD, EMA_CROSS_SLOW,
    INFERENCE_HISTORY_5M, INFERENCE_HISTORY_1D,
    TECHNICAL_LOOKBACK_5M,
    DATA_DIR,
    BREAKOUT_N, CONFLUENCE_BREAKOUT_RVOL,
    get_excel_path,
)
from engine.config import colnum
from engine.config import BTB_SHEET1_MAP, FYERS_INDEX_MAP

from engine import numba_kernels as nk
from engine.indicators import (
    compute_generic_indicators, compute_extended_indicators,
    compute_vwap, macd_cross_label, compute_er,
    ema_fast, detect_cross_aligned,
)
from engine.data_engine import get_root_folder, is_market_open_now
from engine.btb_features import compute_btb_features

logger = logging.getLogger("btb_engine")


# ── Utility functions ───────────────────────────────────────────────

def clamp(n, minn, maxn):
    return max(min(maxn, n), minn)


def _safe_num(value: Any, default: float = 0.0) -> float:
    try:
        out = float(value)
        if np.isfinite(out):
            return out
    except Exception:
        pass
    return default


def _last_sealed_bar(df: pd.DataFrame) -> Dict[str, float]:
    """
    Returns the last COMPLETED bar's high/low/close from a sealed 5m context.

    Prefers the last completed bar (excluding a live-stitched forming/stub bar),
    then rolls back one bar when the trailing bar is degenerate (high <= low).
    Mirrors the sealed-bar discipline used by indicator stripping.
    """
    if df is None or df.empty:
        return {"high": 0.0, "low": 0.0, "close": 0.0, "deg": 0}
    try:
        high = float(df["high"].iloc[-1])
        low = float(df["low"].iloc[-1])
        close = float(df["close"].iloc[-1])
        if high <= low and len(df) >= 2:
            high = float(df["high"].iloc[-2])
            low = float(df["low"].iloc[-2])
            close = float(df["close"].iloc[-2])
            return {"high": high, "low": low, "close": close, "deg": 1}
        return {"high": high, "low": low, "close": close, "deg": 0}
    except Exception:
        return {"high": 0.0, "low": 0.0, "close": 0.0, "deg": 0}


def _safe_float(val, default=0.0):
    try:
        if val is None or val == "":
            return default
        return float(val)
    except Exception:
        return default


def round_float_or_keep(val, digits=2):
    try:
        f = float(val)
        if np.isnan(f):
            return ""
        return round(f, digits)
    except Exception:
        return val


def _fmt_audit_num(value: Any, digits: int = 2) -> str:
    try:
        f = float(value)
        if np.isfinite(f):
            return f"{f:.{digits}f}"
    except Exception:
        pass
    return ""


def _prepare_array(series):
    if series is None or series.empty:
        return np.array([], dtype=np.float64)
    return np.ascontiguousarray(series.to_numpy(dtype=np.float64, na_value=0.0, copy=False))


def _volume_ratio_series(volume_arr: np.ndarray, lookback: int = 20) -> np.ndarray:
    ratios = np.ones(len(volume_arr), dtype=np.float64)
    if len(volume_arr) <= lookback:
        return ratios

    for i in range(lookback, len(volume_arr)):
        base = float(np.mean(volume_arr[i - lookback:i]))
        if base > 0:
            ratios[i] = float(volume_arr[i] / base)
    return ratios


def _ema_direction(ema10: float, ema30: float) -> str:
    if not np.isfinite(ema10) or not np.isfinite(ema30):
        return "FLAT"
    if ema10 > ema30:
        return "UP"
    if ema10 < ema30:
        return "DOWN"
    return "FLAT"


def _build_donchian_event_audit(
    high_arr: np.ndarray,
    low_arr: np.ndarray,
    close_arr: np.ndarray,
    volume_arr: np.ndarray,
    time_arr: np.ndarray,
    atr_val: float,
    n_bars: int = BREAKOUT_N,
) -> Dict[str, Any]:
    """Audit-only event context for BX/BY. It must not feed trade decisions."""
    empty = {
        "packet": "flag=0|age=|time=|level=|dist_pct=|dist_atr=|valid=0|"
                  "rvol_evt=|rvol_now=|rvol_age=|ema_evt=|ema_now=|close=|high=|low=",
        "state": "NO_DONCHIAN_HISTORY",
        "current_flag": 0,
    }

    if len(high_arr) < 2 or len(low_arr) < 2 or len(close_arr) < 2 or len(time_arr) < 2:
        return empty

    try:
        h = np.asarray(high_arr, dtype=np.float64)
        l = np.asarray(low_arr, dtype=np.float64)
        c = np.asarray(close_arr, dtype=np.float64)
        v = np.asarray(volume_arr, dtype=np.float64) if len(volume_arr) == len(h) else np.zeros(len(h), dtype=np.float64)
        times = pd.to_datetime(time_arr, unit="ns")

        dates = times.normalize()
        last_date = dates[-1]
        idxs = np.where(dates == last_date)[0]
        if len(idxs) < 2:
            return empty

        vol_ratio = _volume_ratio_series(v)
        ema10 = ema_fast(c, 10)
        ema30 = ema_fast(c, EMA_CROSS_SLOW)

        flags: Dict[int, int] = {}
        levels: Dict[int, float] = {}
        for pos in range(1, len(idxs)):
            abs_i = int(idxs[pos])
            lb = idxs[max(0, pos - n_bars):pos]
            if len(lb) == 0:
                continue
            prev_hi = float(np.max(h[lb]))
            prev_lo = float(np.min(l[lb]))
            hi_broke = bool(h[abs_i] > prev_hi)
            lo_broke = bool(l[abs_i] < prev_lo)
            if hi_broke and lo_broke:
                flags[abs_i] = 0
            elif hi_broke:
                flags[abs_i] = 1
                levels[abs_i] = prev_hi
            elif lo_broke:
                flags[abs_i] = -1
                levels[abs_i] = prev_lo

        current_i = int(idxs[-1])
        current_flag = int(flags.get(current_i, 0))
        event_idxs = [i for i in idxs if int(flags.get(int(i), 0)) != 0]
        rvol_idxs = [int(i) for i in idxs if float(vol_ratio[int(i)]) >= float(CONFLUENCE_BREAKOUT_RVOL)]

        if not event_idxs:
            rvol_age_txt = ""
            if rvol_idxs:
                rvol_age_txt = str(current_i - rvol_idxs[-1])
            packet = (
                f"flag=0|age=|time=|level=|dist_pct=|dist_atr=|valid=0|"
                f"rvol_evt=|rvol_now={_fmt_audit_num(vol_ratio[current_i], 2)}|"
                f"rvol_age={rvol_age_txt}|ema_evt=|ema_now={_ema_direction(float(ema10[current_i]), float(ema30[current_i]))}|"
                f"close={_fmt_audit_num(c[current_i], 2)}|high={_fmt_audit_num(h[current_i], 2)}|low={_fmt_audit_num(l[current_i], 2)}"
            )
            return {"packet": packet, "state": "NO_RECENT_DONCHIAN", "current_flag": current_flag}

        event_i = int(event_idxs[-1])
        event_flag = int(flags[event_i])
        event_level = float(levels.get(event_i, 0.0))
        age = current_i - event_i
        current_close = float(c[current_i])
        current_high = float(h[current_i])
        current_low = float(l[current_i])

        if event_flag > 0:
            directional_dist = current_close - event_level
            valid = current_close >= event_level
        else:
            directional_dist = event_level - current_close
            valid = current_close <= event_level

        dist_pct = (directional_dist / event_level * 100.0) if event_level > 0 else 0.0
        dist_atr = directional_dist / atr_val if atr_val and atr_val > 0 else 0.0
        rvol_age_txt = str(current_i - rvol_idxs[-1]) if rvol_idxs else ""

        event_time = pd.Timestamp(times[event_i]).strftime("%Y-%m-%d %H:%M:%S")
        ema_evt = _ema_direction(float(ema10[event_i]), float(ema30[event_i]))
        ema_now = _ema_direction(float(ema10[current_i]), float(ema30[current_i]))
        packet = (
            f"flag={event_flag}|age={age}|time={event_time}|level={_fmt_audit_num(event_level, 2)}|"
            f"dist_pct={_fmt_audit_num(dist_pct, 3)}|dist_atr={_fmt_audit_num(dist_atr, 3)}|"
            f"valid={1 if valid else 0}|rvol_evt={_fmt_audit_num(vol_ratio[event_i], 2)}|"
            f"rvol_now={_fmt_audit_num(vol_ratio[current_i], 2)}|rvol_age={rvol_age_txt}|"
            f"ema_evt={ema_evt}|ema_now={ema_now}|close={_fmt_audit_num(current_close, 2)}|"
            f"high={_fmt_audit_num(current_high, 2)}|low={_fmt_audit_num(current_low, 2)}"
        )

        state = "DONCHIAN_VALID" if valid else "DONCHIAN_INVALID"
        state = f"{state}_{age}B"
        if float(vol_ratio[event_i]) >= float(CONFLUENCE_BREAKOUT_RVOL) and float(vol_ratio[current_i]) < float(CONFLUENCE_BREAKOUT_RVOL):
            state = f"{state}_RVOL_FADED"
        elif float(vol_ratio[current_i]) >= float(CONFLUENCE_BREAKOUT_RVOL):
            state = f"{state}_RVOL_OK"
        else:
            state = f"{state}_RVOL_WEAK"

        return {"packet": packet, "state": state, "current_flag": current_flag}
    except Exception as exc:
        logger.debug("Donchian event audit failed: %s", exc, exc_info=True)
        return empty


# ── Market structure ────────────────────────────────────────────────

def calculate_wilder_adx(data, period=14):
    """Computes Wilder's ADX from OHLC DataFrame."""
    if data is None or len(data) < period + 1:
        return None, None, None, 0
    try:
        df = data.copy()
        df.columns = [c.lower() for c in df.columns]
        df['h-l'] = df['high'] - df['low']
        df['h-pc'] = abs(df['high'] - df['close'].shift(1))
        df['l-pc'] = abs(df['low'] - df['close'].shift(1))
        df['tr'] = df[['h-l', 'h-pc', 'l-pc']].max(axis=1)
        df['upmove'] = df['high'] - df['high'].shift(1)
        df['downmove'] = df['low'].shift(1) - df['low']
        df['+dm'] = np.where((df['upmove'] > df['downmove']) & (df['upmove'] > 0), df['upmove'], 0)
        df['-dm'] = np.where((df['downmove'] > df['upmove']) & (df['downmove'] > 0), df['downmove'], 0)
        alpha = 1.0 / period
        df['tr14'] = df['tr'].ewm(alpha=alpha, adjust=False, min_periods=period).mean()
        df['+dm14'] = df['+dm'].ewm(alpha=alpha, adjust=False, min_periods=period).mean()
        df['-dm14'] = df['-dm'].ewm(alpha=alpha, adjust=False, min_periods=period).mean()
        df['+di14'] = 100 * np.where(df['tr14'] != 0, df['+dm14'] / df['tr14'], np.nan)
        df['-di14'] = 100 * np.where(df['tr14'] != 0, df['-dm14'] / df['tr14'], np.nan)
        df['di_diff'] = abs(df['+di14'] - df['-di14'])
        df['di_sum'] = df['+di14'] + df['-di14']
        df['dx'] = np.where(df['di_sum'] != 0, 100 * (df['di_diff'] / df['di_sum']), np.nan)
        df['adx'] = df['dx'].ewm(alpha=alpha, adjust=False, min_periods=period).mean()
        latest = df.iloc[-1]
        p, m, a = latest['+di14'], latest['-di14'], latest['adx']
        slope = 0
        if len(df) > 5:
            slope = a - df['adx'].iloc[-5]
        return round(p, 2) if not pd.isna(p) else 0, round(m, 2) if not pd.isna(m) else 0, round(a, 2) if not pd.isna(a) else 0, round(slope, 2)
    except Exception:
        return None, None, None, 0


def _compute_5m_regime(tf_5m, close_5m):
    """Compute 5-minute regime with MILD split by Trend_Score sign.

    V5.01: Fixed RANGING detection. Previous base=15 gave minimum comp ~29
    for any positive trend_s, making RANGING impossible for mildly bullish
    stocks. Lowered base to 5, added ADX<18 fast-path, adjusted thresholds.
    """
    if not tf_5m:
        return "RANGING"
    try:
        adx_5m   = float(tf_5m.get("adx", 20) or 20)
        pdi_5m   = float(tf_5m.get("plus_di", 0) or 0)
        mdi_5m   = float(tf_5m.get("minus_di", 0) or 0)
        rsi_5m   = float(tf_5m.get("rsi", 50) or 50)
        emas     = tf_5m.get("emas", {}) or {}
        ema30_5m = float(emas.get(30, close_5m) or close_5m)
        ema50_5m = float(emas.get(50, close_5m) or close_5m)

        # Fast-path: very low ADX = no trend, regardless of score
        if adx_5m < 18:
            return "RANGING"

        # Trend_Score (same formula as btb_grader.py grade_trend_precomp)
        trend_s = 0.0
        if adx_5m > 35: trend_s += 40
        elif adx_5m > 25: trend_s += 30
        elif adx_5m > 20: trend_s += 20
        else: trend_s += 10
        trend_s += max(-40, min(40, (pdi_5m - mdi_5m) * 2))
        if close_5m > ema30_5m > ema50_5m: trend_s += 20
        elif close_5m > ema30_5m: trend_s += 10
        elif close_5m < ema30_5m < ema50_5m: trend_s -= 20
        elif close_5m < ema30_5m: trend_s -= 10
        trend_s = max(-100, min(100, trend_s))

        # Directional Momentum_Score (RSI-based).
        # Bullish regimes use high RSI as strength; bearish regimes use low RSI
        # symmetrically as downside momentum. The previous single bullish RSI
        # score made TRENDING_BEAR effectively unreachable.
        if trend_s < 0:
            if rsi_5m < 30: rs = 75
            elif rsi_5m < 40: rs = 65
            elif rsi_5m < 50: rs = 55
            elif rsi_5m < 60: rs = 45
            elif rsi_5m < 70: rs = 35
            else: rs = 25
            trend_component = -trend_s
        else:
            if rsi_5m > 70: rs = 75
            elif rsi_5m > 60: rs = 65
            elif rsi_5m > 50: rs = 55
            elif rsi_5m > 40: rs = 45
            elif rsi_5m > 30: rs = 35
            else: rs = 25
            trend_component = trend_s
        mom_s = max(0, min(100, rs * 0.6 + 30))

        # Composite (V5.01: base lowered from 15 to 5 to allow RANGING)
        comp = trend_component * 0.40 + mom_s * 0.30 + 5

        # Classify with MILD split (V5.01: thresholds adjusted)
        if trend_s > 0 and comp > 40:
            return "TRENDING_BULL"
        elif trend_s < 0 and comp > 35:
            return "TRENDING_BEAR"
        elif comp > 22:
            return "MILD_UP" if trend_s > 0 else "MILD_DOWN"
        return "RANGING"
    except Exception:
        return "RANGING"


def compute_market_structure(df_1d, df_5m):
    """Score (0-100), regime, bias, and ADX tuple for a stock."""
    s = 50
    r = "RANGING"
    b = "NEUTRAL"
    ads = (0, 0, 0)
    if df_1d is None or len(df_1d) < 20:
        return 0, "NO_DATA", "NEUTRAL", ads
    try:
        df_1d = df_1d.rename(columns=str.lower)
        close = df_1d['close'].iloc[-1]
        p, m, a, sl = calculate_wilder_adx(df_1d)
        ads = (p, m, a)
        if a > 25:
            val = 20 if a > 35 else 15
            if p > m:
                s += val; r = "TRENDING_BULL"; b = "BULLISH"
            else:
                s -= val; r = "TRENDING_BEAR"; b = "BEARISH"
        else:
            sma = df_1d['close'].tail(20).mean()
            if close > sma:
                s += 10; r = "MILD"; b = "BULLISH"
            elif close < sma:
                s -= 10; r = "RANGING"; b = "BEARISH"
        s = clamp(s, 0, 100)
    except Exception:
        pass
    return s, r, b, ads


def compute_market_climate(nifty_df, sector_df=None):
    """0-100 composite climate score from Nifty data."""
    if nifty_df is None or len(nifty_df) < 50:
        return 50.0, {}
    try:
        ltp = nifty_df['close'].iloc[-1]
        ema200 = nifty_df['close'].ewm(span=200).mean().iloc[-1]
        trend = 100 if ltp > ema200 else 0
        delta = nifty_df['close'].diff()
        gain = (delta.where(delta > 0, 0)).rolling(14).mean()
        loss = (-delta.where(delta < 0, 0)).rolling(14).mean()
        rsi = 100 - (100 / (1 + gain / loss)).iloc[-1]
        tr = np.maximum(nifty_df['high'] - nifty_df['low'], np.abs(nifty_df['high'] - nifty_df['close'].shift(1)))
        vol = 50 + (20 if tr.rolling(5).mean().iloc[-1] < tr.rolling(20).mean().iloc[-1] else -20)
        c = (trend * 0.3) + (rsi * 0.3) + (vol * 0.2) + (50 * 0.2)
        return clamp(c, 0, 100), {"rsi": rsi}
    except Exception:
        return 50.0, {}


# ── Sector mappings ─────────────────────────────────────────────────

def get_sector_mappings(symbol: str, metadata_map: dict) -> dict:
    if not metadata_map:
        return {'primary_sector': 'UNKNOWN', 'secondary_sector': 'Midcap', 'size_bucket': 1}
    s_clean = symbol.upper().replace(".NS", "")
    m = (metadata_map.get(s_clean) or
         metadata_map.get(f"{s_clean}.NS") or
         metadata_map.get(f"^{s_clean}") or
         metadata_map.get(symbol) or
         {"sector": "UNKNOWN", "size": "Midcap"})
    primary = str(m.get("sector", "UNKNOWN"))
    secondary = str(m.get("size", "Midcap"))
    size_map = {'Smallcap': 0, 'Midcap': 1, 'Nifty': 2, 'Large': 2}
    size_bucket = size_map.get(secondary, 1)
    return {'primary_sector': primary, 'secondary_sector': secondary, 'size_bucket': size_bucket}


# ── Indicator computation worker ─────────────────────────────────────

def compute_indicators_for_symbol(symbol: str, df_5m: pd.DataFrame, df_1d: pd.DataFrame,
                                   snapshot_quote=None, metadata=None, context_data=None) -> Dict[str, Any]:
    """
    Computes all multi-timeframe indicators for a single symbol.
    Returns a result dict with indicator values, LTP, and market structure.
    """
    try:
        now = datetime.now().replace(tzinfo=None)
        market_open = is_market_open_now(now)

        q = snapshot_quote or {}
        q_ltp = float(q.get("ltp", 0.0)) if isinstance(q, dict) else 0.0

        if market_open and q_ltp > 0:
            ltp = q_ltp
        else:
            if df_1d is not None and not df_1d.empty:
                ltp = float(df_1d.iloc[-1]["close"])
            else:
                ltp = 0.0

        # Previous day close
        prev_close = 0.0
        if df_1d is not None and not df_1d.empty:
            if df_1d.index[-1].date() == now.date() and len(df_1d) >= 2:
                prev_close = float(df_1d.iloc[-2]["close"])
            elif df_1d.index[-1].date() < now.date() and len(df_1d) >= 1:
                prev_close = float(df_1d.iloc[-1]["close"])

        # Prep 5m data
        calc_5m = df_5m.tail(INFERENCE_HISTORY_5M) if df_5m is not None else pd.DataFrame()
        if not calc_5m.empty and not isinstance(calc_5m.index, pd.DatetimeIndex):
            calc_5m.index = pd.to_datetime(calc_5m.index)

        # Live stitching
        if market_open and q_ltp > 0 and not calc_5m.empty:
            now_ts = now.replace(second=0, microsecond=0)
            now_ts = now_ts - timedelta(minutes=now_ts.minute % 5)
            if calc_5m.index[-1] == now_ts:
                calc_5m.loc[now_ts, 'close'] = q_ltp
                calc_5m.loc[now_ts, 'high'] = max(calc_5m.loc[now_ts, 'high'], q_ltp)
                calc_5m.loc[now_ts, 'low'] = min(calc_5m.loc[now_ts, 'low'], q_ltp)
            elif calc_5m.index[-1] < now_ts:
                new_row = pd.DataFrame({
                    "open": [q_ltp], "high": [q_ltp], "low": [q_ltp], "close": [q_ltp], "volume": [0]
                }, index=[now_ts])
                calc_5m = pd.concat([calc_5m, new_row])

        def get_dict(df):
            if df.empty:
                return {
                    "open": np.array([]), "high": np.array([]),
                    "low": np.array([]), "close": np.array([]),
                    "volume": np.array([]), "times": np.array([]), "ghost": np.array([])
                }
            return {
                "open": _prepare_array(df["open"]),
                "high": _prepare_array(df["high"]),
                "low": _prepare_array(df["low"]),
                "close": _prepare_array(df["close"]),
                "volume": _prepare_array(df["volume"]),
                "times": df.index.values.astype(np.int64),
                "ghost": df["_ghost"].values if "_ghost" in df.columns else np.array([])
            }

        calc_1d = df_1d.tail(INFERENCE_HISTORY_1D) if df_1d is not None and not df_1d.empty else df_1d

        d1 = get_dict(calc_1d)
        tf = {}
        tf["1d"] = compute_generic_indicators(d1, doji_threshold=DOJI_THRESHOLD)
        tf["1d"]["er"] = compute_er(d1["close"], period=10)

        # 5m, 15m, 1h indicators
        technical_context_5m = df_5m.tail(TECHNICAL_LOOKBACK_5M).copy() if df_5m is not None else pd.DataFrame()
        indicator_context_5m = calc_5m if calc_5m is not None and not calc_5m.empty else technical_context_5m
        signal_indicator_context_5m = indicator_context_5m

        try:
            if market_open and indicator_context_5m is not None and not indicator_context_5m.empty and len(indicator_context_5m) > 1:
                current_bar_ts = now.replace(second=0, microsecond=0)
                current_bar_ts = current_bar_ts - timedelta(minutes=current_bar_ts.minute % 5)
                last_idx = pd.Timestamp(indicator_context_5m.index[-1]).tz_localize(None) if getattr(indicator_context_5m.index[-1], "tzinfo", None) else pd.Timestamp(indicator_context_5m.index[-1])
                if last_idx >= pd.Timestamp(current_bar_ts):
                    signal_indicator_context_5m = indicator_context_5m.iloc[:-1]
        except Exception:
            signal_indicator_context_5m = indicator_context_5m

        d5 = get_dict(signal_indicator_context_5m)
        tf["5m"] = compute_generic_indicators(d5, doji_threshold=DOJI_THRESHOLD)
        # Keep the live 5m cross available internally, but publish AC as the
        # cross between the latest two sealed 5m candles. V6 consumes AC for
        # exits, which must not depend on an intrabar LTP blip.
        if len(d5["close"]) >= 2:
            tf["5m"]["ema_cross_closed"] = detect_cross_aligned(
                ema_fast(d5["close"], 10), ema_fast(d5["close"], EMA_CROSS_SLOW)
            )
        else:
            tf["5m"]["ema_cross_closed"] = ""
        tf["5m"].update(
            compute_extended_indicators(
                d5["high"], d5["low"], d5["close"], d5["volume"],
                False, times=None, df_1d=df_1d, open_arr=d5["open"]
            )
        )
        tf["5m"]["er"] = compute_er(d5["close"], period=10)
        _sealed = _last_sealed_bar(signal_indicator_context_5m)
        tf["5m"]["high_last"] = float(_sealed["high"])
        tf["5m"]["low_last"] = float(_sealed["low"])
        tf["5m"]["close_last"] = float(_sealed["close"])
        tf["5m"]["sealed_rollback"] = int(_sealed["deg"])

        # Volume ratio: current bar volume / mean of previous 20 bars
        v_arr = d5["volume"]
        v_ratio = 1.0
        if len(v_arr) >= 21:
            v_sma20 = np.mean(v_arr[-21:-1])
            if v_sma20 > 0:
                v_ratio = float(v_arr[-1] / v_sma20)
        tf["5m"]["vol_ratio"] = v_ratio

        # Distance from daily high/low (in ATR units)
        # Uses completed-bar (stripped) dataset — matches backtest computation exactly
        if signal_indicator_context_5m is not None and not signal_indicator_context_5m.empty and len(signal_indicator_context_5m) > 0:
            try:
                day_hi_s = signal_indicator_context_5m['high'].resample('1D').transform('max')
                day_lo_s = signal_indicator_context_5m['low'].resample('1D').transform('min')
                close_val = float(d5["close"][-1]) if len(d5["close"]) > 0 else 0.0
                tf["5m"]["day_high"] = round(float(day_hi_s.iloc[-1]), 2)
                tf["5m"]["day_low"] = round(float(day_lo_s.iloc[-1]), 2)
                atr_val = float(tf["5m"].get("atr", 0.0))
                if atr_val > 0:
                    tf["5m"]["dist_day_hi"] = round((float(day_hi_s.iloc[-1]) - close_val) / atr_val, 4)
                    tf["5m"]["dist_day_lo"] = round((close_val - float(day_lo_s.iloc[-1])) / atr_val, 4)
                else:
                    tf["5m"]["dist_day_hi"] = 0.0
                    tf["5m"]["dist_day_lo"] = 0.0
            except Exception:
                tf["5m"]["day_high"] = 0.0
                tf["5m"]["day_low"] = 0.0
                tf["5m"]["dist_day_hi"] = 0.0
                tf["5m"]["dist_day_lo"] = 0.0
        else:
            tf["5m"]["day_high"] = 0.0
            tf["5m"]["day_low"] = 0.0
            tf["5m"]["dist_day_hi"] = 0.0
            tf["5m"]["dist_day_lo"] = 0.0

        # Recompute 5m crosses and RSI on full (unstripped) data to avoid 1-bar delay.
        # Aligns RSI tick-for-tick with live broker and TradingView charts.
        if ltp > 0 and len(d5["close"]) > 0:
            cf = np.append(d5["close"], ltp) if d5["close"][-1] != ltp else d5["close"]
            tf["5m"]["ema_cross"] = detect_cross_aligned(
                ema_fast(cf, 10), ema_fast(cf, EMA_CROSS_SLOW))
            tf["5m"]["ema10_50_cross"] = detect_cross_aligned(
                ema_fast(cf, 10), ema_fast(cf, 50))
            rsi_5m_live = nk.rsi(cf, 14)
            tf["5m"]["rsi_series"] = rsi_5m_live
            if len(rsi_5m_live) > 0 and not np.isnan(rsi_5m_live[-1]):
                tf["5m"]["rsi"] = float(rsi_5m_live[-1])

        # 10m, 15m and 1h via resample. Use the longer local parquet context
        # (3000 completed 5m bars) so the 1h EMA/RSI warm-up is stable. Copy
        # the latest LTP-updated 5m row from the live context to retain forming
        # candle behavior without another history request.
        resample_src = technical_context_5m if (technical_context_5m is not None and not technical_context_5m.empty) else indicator_context_5m
        if resample_src is not None and not resample_src.empty and indicator_context_5m is not None and not indicator_context_5m.empty:
            try:
                resample_src = resample_src.copy()
                latest_ts = pd.Timestamp(indicator_context_5m.index[-1])
                latest_row = indicator_context_5m.iloc[-1]
                for col in ("open", "high", "low", "close", "volume"):
                    resample_src.loc[latest_ts, col] = latest_row[col]
                resample_src.sort_index(inplace=True)
            except Exception:
                # Keep the long historical source if syncing the live row fails.
                pass
        if resample_src is not None and not resample_src.empty:
            d5_resample = get_dict(resample_src)
            times = resample_src.index.values.astype(np.int64)
            for label, interval in [("10m", 10), ("15m", 15), ("1h", 60)]:
                cnt, t, o, h, l, c, v = nk.resample(
                    times,
                    d5_resample["open"], d5_resample["high"], d5_resample["low"],
                    d5_resample["close"], d5_resample["volume"],
                    interval * 60 * 1_000_000_000
                )
                if cnt > 0:
                    sl = slice(0, cnt)
                    d = {
                        "open": o[sl], "high": h[sl],
                        "low": l[sl], "close": c[sl],
                        "volume": v[sl], "ghost": np.array([])
                    }
                    tf[label] = compute_generic_indicators(d, doji_threshold=DOJI_THRESHOLD)
                    if "rsi_series" in tf[label] and len(tf[label]["rsi_series"]) > 0:
                        tf[label]["rsi"] = float(tf[label]["rsi_series"][-1])
                    # Extended indicators for 15m (fibs, swing levels)
                    if label == "15m":
                        tf[label].update(
                            compute_extended_indicators(
                                d["high"], d["low"], d["close"], d["volume"],
                                False, times=None, df_1d=df_1d, open_arr=d["open"]
                            )
                        )
                else:
                    tf[label] = {}
        else:
            tf["10m"] = {}
            tf["15m"] = {}
            tf["1h"] = {}

        # Market structure
        _stk_score, _stk_regime, _stk_bias, _adx_tup = compute_market_structure(df_1d, calc_5m)

        # ── BTB Feature Layer (shadow; sheet rows side-neutral) ──
        btb_out = {}
        try:
            sector_ctx = context_data or {}
            btb_sector_5m = sector_ctx.get("sector_5m")
            is_index = bool(sector_ctx.get("is_index", False))
            if btb_sector_5m is not None and not isinstance(btb_sector_5m, pd.DataFrame):
                try:
                    btb_sector_5m = pd.DataFrame(btb_sector_5m)
                except Exception:
                    btb_sector_5m = None
            if btb_sector_5m is not None and hasattr(btb_sector_5m, "index") and not isinstance(btb_sector_5m.index, pd.DatetimeIndex):
                try:
                    btb_sector_5m.index = pd.to_datetime(btb_sector_5m.index)
                except Exception:
                    pass
            l1_quote = q if isinstance(q, dict) else None
            btb_out = compute_btb_features(
                signal_indicator_context_5m, df_1d, btb_sector_5m,
                ltp, l1=l1_quote, root=get_root_folder(),
                symbol=symbol, is_index=is_index,
            )
        except Exception as e:
            logger.warning(f"[{symbol}] BTB features skipped: {e}")
            btb_out = {}

        result = {
            "symbol": symbol,
            "ltp": ltp,
        }
        # Flatten timeframe indicator dicts to top level for build_indicator_row
        for _tf_name, _tf_data in tf.items():
            result[_tf_name] = _tf_data
        result.update({
            "prev_close": prev_close,
            "vwap": compute_vwap(d5["high"], d5["low"], d5["close"], d5["volume"], d5["times"]),
            "spike": tf["5m"].get("spike", 0),
            "supertrend": tf["1d"].get("supertrend", ""),
            "supertrend_5m": tf["5m"].get("supertrend", ""),
            "no_data": False,
            "metadata": metadata or {},
            "sector": (metadata or {}).get("sector", ""),
            "size": (metadata or {}).get("size", ""),
            "er": tf["1d"].get("er", 0.0),
            "stk_score": _stk_score,
            "stk_regime": _stk_regime,
            "stk_bias": _stk_bias,
            "plus_di": _adx_tup[0],
            "minus_di": _adx_tup[1],
            "adx": _adx_tup[2],
            "regime_5m": _compute_5m_regime(tf["5m"], float(d5["close"][-1]) if len(d5["close"]) > 0 else 0.0),
            "_5m_high": d5["high"],
            "_5m_low": d5["low"],
            "_5m_close": d5["close"],
            "_5m_volume": d5["volume"],
            "_5m_times": d5["times"],
            "btb_features": btb_out,
        })
        return result
    except Exception as e:
        logger.exception(f"[{symbol}] COMPUTE_FAIL")
        return {"symbol": symbol, "no_data": True, "error": str(e)}


# ── Excel column mapping ────────────────────────────────────────────

def build_indicator_row(res: Dict[str, Any]) -> List[Any]:
    """
    Maps indicator results to ordered Excel column list.
    Populates TA indicator columns plus audit-only BX/BY Donchian event context.
    """
    col_map = {c: "" for c in FULL_COLUMNS}
    col_map[LTP_COL] = round_float_or_keep(res.get("ltp", ""))

    # RSI by timeframe
    for c, tf_name in zip(RSI_COLS, TIMEFRAMES):
        ser = res.get(tf_name, {}).get("rsi_series")
        if isinstance(ser, np.ndarray) and len(ser) > 0:
            col_map[c] = round_float_or_keep(ser[-1], 2)
        elif isinstance(ser, pd.Series) and not ser.empty:
            col_map[c] = round_float_or_keep(ser.iloc[-1], 2)

    # MACD cross labels
    for c, tf_name in zip(MACD_TEXT_COLS, TIMEFRAMES):
        l = res.get(tf_name, {}).get("macd_line")
        s = res.get(tf_name, {}).get("macd_signal")
        if isinstance(l, np.ndarray) and isinstance(s, np.ndarray) and len(l) > 0:
            col_map[c] = macd_cross_label(l, s)

    # EMA crosses + patterns. AC is the latest completed 5m candle cross for V6 exits;
    # the live/forming 5m cross remains in the indicator result for other consumers.
    for c, tf_name in zip(EMA_CROSS_COLS, TIMEFRAMES):
        if c == "AC":
            col_map[c] = res.get("5m", {}).get("ema_cross_closed", "")
        else:
            col_map[c] = res.get(tf_name, {}).get("ema_cross", "")
    # AF remains the 1d EMA cross via EMA_CROSS_COLS/TIMEFRAMES.
    # DI and DJ hold the live forming 10m RSI and EMA cross for V6.
    col_map["DI"] = round_float_or_keep(res.get("10m", {}).get("rsi"), 2)
    col_map["DJ"] = res.get("10m", {}).get("ema_cross", "")
    for c, tf_name in zip(EMA10_50_CROSS_COLS, TIMEFRAMES):
        col_map[c] = res.get(tf_name, {}).get("ema10_50_cross", "")
    for c, tf_name in zip(BULL_PATTERN_COLS, TIMEFRAMES):
        col_map[c] = res.get(tf_name, {}).get("bullish_pattern", "")
    for c, tf_name in zip(BEAR_PATTERN_COLS, TIMEFRAMES):
        col_map[c] = res.get(tf_name, {}).get("bearish_pattern", "")

    # EMAs by timeframe
    e5 = res.get("5m", {}).get("emas", {})
    col_map["AS"] = round_float_or_keep(e5.get(10))
    col_map["AT"] = round_float_or_keep(e5.get(EMA_CROSS_SLOW))
    col_map["AU"] = round_float_or_keep(e5.get(50))
    col_map["AV"] = round_float_or_keep(e5.get(100))
    col_map["AW"] = round_float_or_keep(e5.get(200))

    e15 = res.get("15m", {}).get("emas", {})
    col_map["AX"] = round_float_or_keep(e15.get(10))
    col_map["AY"] = round_float_or_keep(e15.get(EMA_CROSS_SLOW))
    col_map["AZ"] = round_float_or_keep(e15.get(50))
    col_map["BA"] = round_float_or_keep(e15.get(100))
    col_map["BB"] = round_float_or_keep(e15.get(200))

    e1h = res.get("1h", {}).get("emas", {})
    col_map["BC"] = round_float_or_keep(e1h.get(10))
    col_map["BD"] = round_float_or_keep(e1h.get(EMA_CROSS_SLOW))
    col_map["BE"] = round_float_or_keep(e1h.get(50))
    col_map["BF"] = round_float_or_keep(e1h.get(100))
    col_map["BG"] = round_float_or_keep(e1h.get(200))

    e1d = res.get("1d", {}).get("emas", {})
    col_map["BH"] = round_float_or_keep(e1d.get(10))
    col_map["BI"] = round_float_or_keep(e1d.get(EMA_CROSS_SLOW))
    col_map["BJ"] = round_float_or_keep(e1d.get(50))
    col_map["BK"] = round_float_or_keep(e1d.get(100))
    col_map["BL"] = round_float_or_keep(e1d.get(200))

    col_map[VWAP_COL] = round_float_or_keep(res.get("vwap", ""))
    col_map[SPIKE_COL] = res.get("spike", "")
    col_map[SUPERTREND_COL] = round_float_or_keep(res.get("supertrend_5m", ""))

    # MACD numeric
    col_map["CY"] = round_float_or_keep(res.get("5m", {}).get("macd_fast_val"))
    col_map["CZ"] = round_float_or_keep(res.get("5m", {}).get("macd_slow_val"))
    col_map["DA"] = round_float_or_keep(res.get("5m", {}).get("macd_sig_val"))

    # ADX/DI
    col_map[DI_PLUS_COL] = round_float_or_keep(res.get("plus_di", ""))
    col_map[DI_MINUS_COL] = round_float_or_keep(res.get("minus_di", ""))
    col_map[ADX_COL] = round_float_or_keep(res.get("5m", {}).get("adx", ""))

    col_map["DB"] = round_float_or_keep(res.get("15m", {}).get("macd_fast_val"))
    col_map["DC"] = round_float_or_keep(res.get("15m", {}).get("macd_slow_val"))
    col_map["DD"] = round_float_or_keep(res.get("15m", {}).get("macd_sig_val"))

    # ATR, Stoch, BB
    col_map[ATR_COLS["5m"]] = round_float_or_keep(res.get("5m", {}).get("atr"))
    col_map[ATR_COLS["15m"]] = round_float_or_keep(res.get("15m", {}).get("atr"))
    col_map[STOCH_K_COLS["5m"]] = round_float_or_keep(res.get("5m", {}).get("stoch_k"))
    col_map[STOCH_D_COLS["5m"]] = round_float_or_keep(res.get("5m", {}).get("stoch_d"))
    col_map[STOCH_K_COLS["15m"]] = round_float_or_keep(res.get("15m", {}).get("stoch_k"))
    col_map[STOCH_D_COLS["15m"]] = round_float_or_keep(res.get("15m", {}).get("stoch_d"))
    col_map[BB_UPPER_COLS["5m"]] = round_float_or_keep(res.get("5m", {}).get("bb_upper"))
    col_map[BB_MID_COLS["5m"]] = round_float_or_keep(res.get("5m", {}).get("bb_mid"))
    col_map[BB_LOWER_COLS["5m"]] = round_float_or_keep(res.get("5m", {}).get("bb_lower"))
    col_map[BB_UPPER_COLS["15m"]] = round_float_or_keep(res.get("15m", {}).get("bb_upper"))
    col_map[BB_MID_COLS["15m"]] = round_float_or_keep(res.get("15m", {}).get("bb_mid"))
    col_map[BB_LOWER_COLS["15m"]] = round_float_or_keep(res.get("15m", {}).get("bb_lower"))

    # Sector & Regime
    m_data = res.get("metadata", {"sector": "", "size": ""})
    col_map[SECTOR_COL] = m_data.get("sector", "") or res.get("sector", "")
    col_map[REGIME_5M_COL] = res.get("regime_5m", "")
    col_map[SECTOR_REGIME_COL] = res.get("sector_regime", "")
    col_map[RVOL_COL] = res.get("5m", {}).get("rvol", "")
    col_map[REGIME_COL] = res.get("stk_regime", "")

    # Volume ratio: current bar volume / mean of previous 20 bars
    col_map[VOL_RATIO_COL] = round_float_or_keep(res.get("5m", {}).get("vol_ratio", 1.0), 4)

    # Distance from daily high/low (in ATR units)
    col_map[DIST_DAY_HI_COL] = round_float_or_keep(res.get("5m", {}).get("dist_day_hi", ""), 4)
    col_map[DIST_DAY_LO_COL] = round_float_or_keep(res.get("5m", {}).get("dist_day_lo", ""), 4)

    # Day high/low (absolute price — used by downstream VBA)
    col_map["BV"] = round_float_or_keep(res.get("5m", {}).get("day_high", ""), 2)
    col_map["BW"] = round_float_or_keep(res.get("5m", {}).get("day_low", ""), 2)

    # Swing & Fibs
    col_map[SWING_HIGH_COL] = round_float_or_keep(res.get("5m", {}).get("swing_high", ""))
    col_map[SWING_LOW_COL] = round_float_or_keep(res.get("5m", {}).get("swing_low", ""))
    fibs_5m = res.get("5m", {}).get("fibs", {})
    if isinstance(fibs_5m, dict):
        col_map[FIB_COLS_5M[0]] = round_float_or_keep(fibs_5m.get("0.236"))
        col_map[FIB_COLS_5M[1]] = round_float_or_keep(fibs_5m.get("0.382"))
        col_map[FIB_COLS_5M[2]] = round_float_or_keep(fibs_5m.get("0.5"))
        col_map[FIB_COLS_5M[3]] = round_float_or_keep(fibs_5m.get("0.618"))
    fibs_15m = res.get("15m", {}).get("fibs", {})
    if isinstance(fibs_15m, dict):
        col_map[FIB_COLS_15M[0]] = round_float_or_keep(fibs_15m.get("0.236"))
        col_map[FIB_COLS_15M[1]] = round_float_or_keep(fibs_15m.get("0.382"))
        col_map[FIB_COLS_15M[2]] = round_float_or_keep(fibs_15m.get("0.5"))
        col_map[FIB_COLS_15M[3]] = round_float_or_keep(fibs_15m.get("0.618"))

    # ER and Size (DO, DP)
    col_map[ER_COL] = round_float_or_keep(res.get("er", ""), 2)
    m_size = m_data.get("size", "")
    col_map[SIZE_COL] = m_size

    # 5m OHLCV last values
    col_map["DR"] = round_float_or_keep(res.get("5m", {}).get("high_last", ""))
    col_map["DS"] = round_float_or_keep(res.get("5m", {}).get("low_last", ""))
    col_map["DT"] = round_float_or_keep(res.get("5m", {}).get("close_last", ""))

    # 1D SuperTrend
    col_map["BU"] = round_float_or_keep(res.get("supertrend", ""))

    # --- Stock Grade (DE) and Breakout Flag (DF) ---
    # Stock Grade: R:R ratio from current bar geometry (A+/A/B/C)
    try:
        last_high = _safe_num(res.get("5m", {}).get("high_last", 0.0))
        last_low = _safe_num(res.get("5m", {}).get("low_last", 0.0))
        last_close = _safe_num(res.get("5m", {}).get("close_last", 0.0))
        atr_5m = _safe_num(res.get("5m", {}).get("atr", 0.0))
        if last_high > 0 and last_low > 0 and last_close > 0:
            risk = last_close - last_low
            reward = last_high - last_close
            if risk > 0 and reward > 0:
                rr = reward / risk
                if rr >= 1.75:
                    grade = "A+"
                elif rr >= 1.25:
                    grade = "A"
                elif rr >= 0.75:
                    grade = "B"
                else:
                    grade = "C"
            else:
                grade = "C"
            col_map[SCORE_COL] = grade
    except Exception:
        pass

    # Breakout Flag: Session-based Donchian breakout (1=Long, -1=Short, 0=none)
    # Donchian channel computed from COMPLETED bars in current session, excluding current bar.
    # Session boundary: 09:15 IST. Lookback: rolling N bars within session (default N=5).
    # This is the scientifically validated approach per Donchian (1960), Faith (2007), and
    # intraday microstructure research (Agarwalla 2010, VN30F 2026).
    BREAKOUT_N_VAL = BREAKOUT_N
    try:
        h_arr = res.get("_5m_high", np.array([]))
        l_arr = res.get("_5m_low", np.array([]))
        c_arr = res.get("_5m_close", np.array([]))
        v_arr = res.get("_5m_volume", np.array([]))
        t_arr = res.get("_5m_times", np.array([]))
        event_audit = _build_donchian_event_audit(
            h_arr, l_arr, c_arr, v_arr, t_arr,
            _safe_num(res.get("5m", {}).get("atr", 0.0)),
            BREAKOUT_N_VAL,
        )
        col_map[AI_SIGNAL_COL] = event_audit.get("packet", "")
        col_map[AI_HEALTH_COL] = event_audit.get("state", "")
        if len(h_arr) >= 2 and len(l_arr) >= 2 and len(t_arr) >= 2:
            # Group bars by trading session (date from timestamp)
            # NSE session: 09:15-15:20 IST
            bar_dates = pd.to_datetime(t_arr, unit="ns").normalize()
            unique_dates = bar_dates.unique()
            last_session_date = unique_dates[-1]
            session_mask = bar_dates == last_session_date
            session_h = h_arr[session_mask]
            session_l = l_arr[session_mask]
            n_session = len(session_h)
            if n_session >= 2:
                # Lookback: all session bars EXCEPT the last one (most recent completed bar)
                # This is the key fix: exclude current bar from Donchian channel
                lookback_h = session_h[:-1]
                lookback_l = session_l[:-1]
                # Apply rolling N cap
                if len(lookback_h) > BREAKOUT_N_VAL:
                    lookback_h = lookback_h[-BREAKOUT_N_VAL:]
                    lookback_l = lookback_l[-BREAKOUT_N_VAL:]
                donchian_high = np.max(lookback_h)
                donchian_low = np.min(lookback_l)
                last_h = float(h_arr[-1])
                last_l = float(l_arr[-1])
                if last_h > donchian_high and last_l < donchian_low:
                    col_map[NARRATIVE_COL] = 0   # Ambiguous
                elif last_h > donchian_high:
                    col_map[NARRATIVE_COL] = 1   # Long breakout
                elif last_l < donchian_low:
                    col_map[NARRATIVE_COL] = -1  # Short breakout
                else:
                    col_map[NARRATIVE_COL] = 0
            else:
                col_map[NARRATIVE_COL] = 0
        else:
            col_map[NARRATIVE_COL] = 0
    except Exception:
        pass

    # BX/BY are audit-only Donchian event context. Trading logic must not consume them.
    # Other AI columns (DN, DQ, etc.) remain blank intentionally for downstream writers.

    # ── EMA10/30 Gap Lookback Telemetry Packet for Column DM (Sheet1) ────
    try:
        c_arr = res.get("_5m_close", np.array([]))
        t_arr = res.get("_5m_times", np.array([]))
        atr_val = float(res.get("5m", {}).get("atr", 0.0))
        if len(c_arr) >= 4 and len(t_arr) >= 4 and atr_val > 0:
            e10_s = ema_fast(c_arr, 10)
            e30_s = ema_fast(c_arr, EMA_CROSS_SLOW)
            gap_s = e10_s - e30_s

            bar_dates = pd.to_datetime(t_arr, unit="ns").normalize()
            last_date = bar_dates[-1]
            session_bars = int((bar_dates == last_date).sum())
            k = min(max(session_bars - 1, 0), 3)

            g0 = float(gap_s[-1])
            g_m1 = float(gap_s[-2]) if k >= 1 and len(gap_s) >= 2 else 0.0
            g_m2 = float(gap_s[-3]) if k >= 2 and len(gap_s) >= 3 else 0.0
            g_m3 = float(gap_s[-4]) if k >= 3 and len(gap_s) >= 4 else 0.0

            col_map[AI_SETUP_COL] = f"G0={g0:.2f}|G-1={g_m1:.2f}|G-2={g_m2:.2f}|G-3={g_m3:.2f}|ATR={atr_val:.2f}|K={k}"
        else:
            col_map[AI_SETUP_COL] = ""
    except Exception:
        col_map[AI_SETUP_COL] = ""

    # ── BTB Feature Layer (DV..DZ shadow; V4.2 VBA consumes only DV:DY) ──
    btb = res.get("btb_features") or {}
    if isinstance(btb, dict):
        for _key, _col in BTB_SHEET1_MAP.items():
            _val = btb.get(_key, "")
            col_map[_col] = _val

    return [col_map.get(c, "") for c in FULL_COLUMNS]


# ── Excel COM writer ────────────────────────────────────────────────

class ExcelWriter:
    """Writes indicator results to Excel via win32com."""

    def __init__(self, root: Optional[str] = None):
        self.root = root or get_root_folder()
        self.excel_path = get_excel_path()
        self.sheet = EXCEL_SHEET
        self.last_matrix = None  # exact rows written for this cycle (for telemetry)

    def write_matrix_com(self, results: List[Dict[str, Any]], start_row: int):
        import time
        max_retries = 5
        for attempt in range(1, max_retries + 1):
            try:
                if self._write_attempt(results, start_row):
                    return True
            except Exception as e:
                logger.warning(f"Excel Write Attempt {attempt} failed: {e}")
            if attempt < max_retries:
                wait = 2 ** (attempt - 1)
                time.sleep(wait)
        logger.error("Excel write FAILED after all retries")
        return False

    def _write_attempt(self, results: List[Dict[str, Any]], start_row: int) -> bool:
        import pythoncom
        pythoncom.CoInitialize()

        # Build matrix rows (flat lists matching FULL_COLUMNS order)
        matrix = [build_indicator_row(r) for r in results]
        self.last_matrix = matrix  # telemetry reads the exact written row values
        if not matrix:
            pythoncom.CoUninitialize()
            return False

        try:
            import win32com.client

            excel = win32com.client.Dispatch("Excel.Application")
            excel.ScreenUpdating = False
            excel.DisplayAlerts = False
            excel.Calculation = -4135  # xlCalculationManual

            # Try to get already-open workbook first, else open fresh
            try:
                wb = excel.Workbooks(os.path.basename(self.excel_path))
            except Exception:
                wb = excel.Workbooks.Open(self.excel_path)

            ws = wb.Worksheets[self.sheet]

            num_rows = len(matrix)
            end_row = start_row + num_rows - 1
            current_start_col = max(20, OUTPUT_START_COL_INDEX)

            # 1. Range T to BL (indices 0 to 44)
            matrix_part_t_bl = [row[0:45] for row in matrix]
            rng_t_bl = ws.Range(ws.Cells(start_row, current_start_col), ws.Cells(end_row, current_start_col + 44))
            rng_t_bl.Value = tuple(tuple(r) for r in matrix_part_t_bl)

            # 2. Range BN to CB (indices 46 to 60)
            matrix_part_bn_cb = [row[46:61] for row in matrix]
            rng_bn_cb = ws.Range(ws.Cells(start_row, OUTPUT_START_COL_INDEX + 46), ws.Cells(end_row, OUTPUT_START_COL_INDEX + 60))
            rng_bn_cb.Value = tuple(tuple(r) for r in matrix_part_bn_cb)

            # 3. Range CC to DN (indices 61 to 98)
            matrix_part_cc_dn = [row[61:99] for row in matrix]
            rng_cc_dn = ws.Range(ws.Cells(start_row, OUTPUT_START_COL_INDEX + 61), ws.Cells(end_row, OUTPUT_START_COL_INDEX + 98))
            rng_cc_dn.Value = tuple(tuple(r) for r in matrix_part_cc_dn)

            # 4. Range DP to DQ (indices 100 to end)
            matrix_part_dp_dq = [row[100:] for row in matrix]
            n_cols = len(FULL_COLUMNS)
            rng_dp_dq = ws.Range(ws.Cells(start_row, OUTPUT_START_COL_INDEX + 100), ws.Cells(end_row, OUTPUT_START_COL_INDEX + n_cols - 1))
            rng_dp_dq.Value = tuple(tuple(r) for r in matrix_part_dp_dq)

            # Recalculate, then run macro (RunAll performs ThisWorkbook.Save at line 54)
            excel.Calculation = -4105  # xlCalculationAutomatic
            excel.Calculate()
            excel.Calculation = -4135  # xlCalculationManual

            t_macro = _time.time()
            logger.info("Running Excel Macro 'RunAll'...")
            try:
                excel.Application.Run("RunAll")
                logger.info(f"[PERF] Macro 'RunAll': {_time.time() - t_macro:.2f}s")
            except Exception as e:
                logger.warning(f"Macro RunAll failed (non-fatal): {e}")

            excel.Calculation = -4105  # xlCalculationAutomatic
            excel.ScreenUpdating = True

            return True

        except Exception as e:
            logger.error(f"Excel COM Failed: {e}")
            try:
                excel.Calculation = -4105
                excel.ScreenUpdating = True
            except Exception:
                pass
            return False
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


# ── Excel config reader ─────────────────────────────────────────────

def read_excel_config(root: Optional[str] = None):
    """
    Reads symbol list, sector map, and row range from Excel via openpyxl.
    Returns (symbols, sector_map, start_row, metadata_map).
    """
    import openpyxl
    root = root or get_root_folder()
    path = get_excel_path()
    symbols = []
    sector_map = {}
    metadata_map = {}
    start_row = 4

    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb[EXCEL_SHEET]

        sr = ws["P1"].value
        start_row = int(sr) if sr and str(sr).strip().isdigit() else 4

        er = ws["R1"].value
        end_row = int(er) if er and str(er).strip().isdigit() else 2500
        if end_row < start_row:
            end_row = start_row

        col_sec = colnum(SECTOR_COL)
        col_size = colnum(SIZE_COL)

        for row in ws.iter_rows(min_row=start_row, max_row=end_row, min_col=1, max_col=max(col_sec, col_size)):
            val = row[0].value
            if val:
                s = str(val).strip()
                symbols.append(s)
                m_sec = str(row[col_sec - 1].value).strip() if row[col_sec - 1].value else ""
                m_size = str(row[col_size - 1].value).strip() if row[col_size - 1].value else ""
                sector_map[s] = m_sec
                metadata_map[s] = {"sector": m_sec, "size": m_size}

        wb.close()
    except Exception as e:
        logger.error(f"read_excel_config failed: {e}")

    return symbols, sector_map, start_row, metadata_map
