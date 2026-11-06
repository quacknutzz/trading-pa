"""
indicators.py — Higher-level indicator wrappers over Numba kernels.

Reads:   OHLCV arrays (numpy) and daily DataFrames (pandas)
Computes: RSI, EMA, MACD, VWAP, SuperTrend, Bollinger, Stochastic,
          candlestick patterns, spike detection, swing/fibs, RVOL
Returns: Indicator values and dicts for downstream signal engine
Refers:  numba_kernels.py for raw computation, config.py for params
"""

import logging
from datetime import datetime, time as dtime
from typing import Dict, Any

import numpy as np
import pandas as pd

from engine import numba_kernels as nk
from engine.config import DOJI_THRESHOLD, FIB_LEVELS, EMA_CROSS_SLOW

logger = logging.getLogger("ai_signal_engine")


# ===============================
# SIMPLE WRAPPERS (Thin layer over Numba kernels)
# ===============================

def ema_fast(arr: np.ndarray, period: int) -> np.ndarray:
    """Computes EMA via Numba kernel. Returns: EMA array."""
    return nk.ema(arr, period)


def rsi_wilder(close_arr: np.ndarray, length: int = 14) -> np.ndarray:
    """Computes Wilder's RSI. Returns: RSI array (0-100, NaN-padded)."""
    if len(close_arr) <= length:
        return np.full(len(close_arr), np.nan)
    return nk.rsi(close_arr, length)


def compute_atr(h: np.ndarray, l: np.ndarray, c: np.ndarray, period: int = 14) -> np.ndarray:
    """Computes ATR. Returns: ATR array (NaN-padded)."""
    if len(c) < period:
        return np.full(len(c), np.nan)
    return nk.atr(h, l, c, period)


def adx_smoothed(h: np.ndarray, l: np.ndarray, c: np.ndarray, period: int = 14, smoothing: int = 3):
    """
    Computes Smoothed ADX and its EMA-filtered slope.
    Returns: (plus_di, minus_di, adx, adx_slope) latest values.
    """
    p_di, m_di, adx_arr = nk.adx(h, l, c, period)
    if len(adx_arr) < 2 or np.isnan(adx_arr[-1]):
        return 0, 0, 0, 0
        
    # Compute slope
    slope_raw = np.diff(adx_arr)
    # EMA smooth the slope
    slope_ema = nk.ema(slope_raw, smoothing)
    
    return float(p_di[-1]), float(m_di[-1]), float(adx_arr[-1]), float(slope_ema[-1])


def atr_percentile_rank(h, l, c, period=14, lookback=200):
    """Computes percentile rank of current ATR."""
    atr_arr = nk.atr(h, l, c, period)
    pct_arr = nk.percentile_rank(atr_arr, lookback)
    if len(pct_arr) > 0 and not np.isnan(pct_arr[-1]):
        return float(pct_arr[-1])
    return 0.5


def macd_fast(close_arr: np.ndarray, fast=12, slow=26, signal=9):
    """
    Computes MACD line, signal, and component EMAs.

    Returns: (macd_line, ema_fast, ema_slow, macd_signal) arrays
    """
    ema_f = nk.ema(close_arr, fast)
    ema_s = nk.ema(close_arr, slow)
    macd = ema_f - ema_s
    macd_signal = nk.ema(macd, signal)
    return macd, ema_f, ema_s, macd_signal


def macd_cross_label(macd_line: np.ndarray, macd_signal: np.ndarray) -> str:
    """Detects MACD cross. Returns: 'Bullish', 'Bearish', or ''."""
    res = nk.detect_ema_cross(macd_line, macd_signal)
    if res == 1: return "Bullish"
    if res == -1: return "Bearish"
    return ""


def compute_vwap(h: np.ndarray, l: np.ndarray, c: np.ndarray, v: np.ndarray, t: np.ndarray) -> float:
    """Computes latest VWAP value with daily reset. Returns: float or '' if unavailable."""
    if len(c) == 0 or len(t) == 0: return ""
    try:
        val = nk.vwap(h, l, c, v, t)[-1]
        return float(val) if not np.isnan(val) else ""
    except:
        return ""


def compute_supertrend(h: np.ndarray, l: np.ndarray, c: np.ndarray, period=10, multiplier=3):
    """Computes SuperTrend last value. Returns: float or '' if insufficient data."""
    if len(c) < period: return ""
    st_arr = nk.supertrend(h, l, c, period, multiplier)
    if len(st_arr) > 0:
        val = st_arr[-1]
        return float(val) if not np.isnan(val) else ""
    return ""


def compute_bollinger(close_arr: np.ndarray, period=20, std_dev=2):
    """Computes Bollinger Bands. Returns: (upper, lower, mid, width%) arrays."""
    return nk.bbands(close_arr, period, std_dev)


def compute_stochastic(h, l, c, k_period=14, d_period=3):
    """Computes Stochastic %K and %D. Returns: (k_vals, d_vals) arrays."""
    return nk.stochastic(h, l, c, k_period, d_period)


def detect_cross_aligned(fast: np.ndarray, slow: np.ndarray) -> str:
    """Detects EMA cross direction. Returns: 'Bullish', 'Bearish', or ''."""
    res = nk.detect_ema_cross(fast, slow)
    if res == 1: return "Bullish"
    if res == -1: return "Bearish"
    return ""


def detect_trend(fast: np.ndarray, slow: np.ndarray) -> str:
    """Determines trend state from EMA pair. Returns: 'Bullish', 'Bearish', or 'Neutral'."""
    if len(fast) == 0: return ""
    res = nk.detect_trend_state(fast[-1], slow[-1])
    if res == 1: return "Bullish"
    if res == -1: return "Bearish"
    return "Neutral"


def compute_er(close_arr: np.ndarray, period: int = 10) -> float:
    """
    Computes Kaufman's Efficiency Ratio (latest value).
    Returns: float (0.0 to 1.0) or 0.0 if insufficient data.
    """
    if len(close_arr) < period + 1:
        return 0.0
    er_arr = nk.efficiency_ratio(close_arr, period)
    if len(er_arr) > 0 and not np.isnan(er_arr[-1]):
        return float(er_arr[-1])
    return 0.0


# ===============================
# SPIKE DETECTION (with Timeframe Escalation)
# ===============================
def compute_spike_directional(vol_arr: np.ndarray, open_arr: np.ndarray, close_arr: np.ndarray,
                               avg_5m_baseline: float = 0.0, is_opening: bool = True,
                               is_market_open: bool = True, daily_vol: np.ndarray = None) -> int:
    """
    Determines directional volume spike with timeframe escalation.

    Reads:   vol_arr (5m volume), open/close arrays, daily volume
    Computes: Robust Z-Score spike detection
    Returns: 1 (bullish spike), -1 (bearish spike), 0 (no spike)
    Refers:  numba_kernels.detect_volume_spike for core detection

    Timeframe Escalation:
    - Market hours: Uses 5m candle volume (real-time, actionable)
    - Post-market:  Uses daily volume (avoids closing auction false spikes)
    """
    # Post-market: use daily volume
    if not is_market_open and daily_vol is not None and len(daily_vol) >= 2:
        if nk.detect_volume_spike(daily_vol, 0.0, False):
            curr_o = open_arr[-1] if len(open_arr) > 0 else 0.0
            curr_c = close_arr[-1] if len(close_arr) > 0 else 0.0
            if curr_c > curr_o: return 1
            if curr_c < curr_o: return -1
        return 0

    # Market hours: use 5m candle volume
    if len(vol_arr) < 2: return 0
    if nk.detect_volume_spike(vol_arr, avg_5m_baseline, is_opening):
        curr_o = open_arr[-1]
        curr_c = close_arr[-1]
        if curr_c > curr_o: return 1
        if curr_c < curr_o: return -1
    return 0

"""
indicators.py — Higher-level indicator wrappers over Numba kernels.

Reads:   OHLCV arrays (numpy) and daily DataFrames (pandas)
Computes: RSI, EMA, MACD, VWAP, SuperTrend, Bollinger, Stochastic,
          candlestick patterns, spike detection, swing/fibs, RVOL
Returns: Indicator values and dicts for downstream signal engine
Refers:  numba_kernels.py for raw computation, config.py for params
"""

import logging
from datetime import datetime, time as dtime
from typing import Dict, Any

import numpy as np
import pandas as pd

from engine import numba_kernels as nk
from engine.config import DOJI_THRESHOLD, FIB_LEVELS, EMA_CROSS_SLOW

logger = logging.getLogger("ai_signal_engine")


# ===============================
# SIMPLE WRAPPERS (Thin layer over Numba kernels)
# ===============================

def ema_fast(arr: np.ndarray, period: int) -> np.ndarray:
    """Computes EMA via Numba kernel. Returns: EMA array."""
    return nk.ema(arr, period)


def rsi_wilder(close_arr: np.ndarray, length: int = 14) -> np.ndarray:
    """Computes Wilder's RSI. Returns: RSI array (0-100, NaN-padded)."""
    if len(close_arr) <= length:
        return np.full(len(close_arr), np.nan)
    return nk.rsi(close_arr, length)


def compute_atr(h: np.ndarray, l: np.ndarray, c: np.ndarray, period: int = 14) -> np.ndarray:
    """Computes ATR. Returns: ATR array (NaN-padded)."""
    if len(c) < period:
        return np.full(len(c), np.nan)
    return nk.atr(h, l, c, period)


def adx_smoothed(h: np.ndarray, l: np.ndarray, c: np.ndarray, period: int = 14, smoothing: int = 3):
    """
    Computes Smoothed ADX and its EMA-filtered slope.
    Returns: (plus_di, minus_di, adx, adx_slope) latest values.
    """
    p_di, m_di, adx_arr = nk.adx(h, l, c, period)
    if len(adx_arr) < 2 or np.isnan(adx_arr[-1]):
        return 0, 0, 0, 0
        
    # Compute slope
    slope_raw = np.diff(adx_arr)
    # EMA smooth the slope
    slope_ema = nk.ema(slope_raw, smoothing)
    
    return float(p_di[-1]), float(m_di[-1]), float(adx_arr[-1]), float(slope_ema[-1])


def atr_percentile_rank(h, l, c, period=14, lookback=200):
    """Computes percentile rank of current ATR."""
    atr_arr = nk.atr(h, l, c, period)
    pct_arr = nk.percentile_rank(atr_arr, lookback)
    if len(pct_arr) > 0 and not np.isnan(pct_arr[-1]):
        return float(pct_arr[-1])
    return 0.5


def macd_fast(close_arr: np.ndarray, fast=12, slow=26, signal=9):
    """
    Computes MACD line, signal, and component EMAs.

    Returns: (macd_line, ema_fast, ema_slow, macd_signal) arrays
    """
    ema_f = nk.ema(close_arr, fast)
    ema_s = nk.ema(close_arr, slow)
    macd = ema_f - ema_s
    macd_signal = nk.ema(macd, signal)
    return macd, ema_f, ema_s, macd_signal


def macd_cross_label(macd_line: np.ndarray, macd_signal: np.ndarray) -> str:
    """Detects MACD cross. Returns: 'Bullish', 'Bearish', or ''."""
    res = nk.detect_ema_cross(macd_line, macd_signal)
    if res == 1: return "Bullish"
    if res == -1: return "Bearish"
    return ""


def compute_vwap(h: np.ndarray, l: np.ndarray, c: np.ndarray, v: np.ndarray, t: np.ndarray) -> float:
    """Computes latest VWAP value with daily reset. Returns: float or '' if unavailable."""
    if len(c) == 0 or len(t) == 0: return ""
    try:
        val = nk.vwap(h, l, c, v, t)[-1]
        return float(val) if not np.isnan(val) else ""
    except:
        return ""


def compute_supertrend(h: np.ndarray, l: np.ndarray, c: np.ndarray, period=10, multiplier=3):
    """Computes SuperTrend last value. Returns: float or '' if insufficient data."""
    if len(c) < period: return ""
    st_arr = nk.supertrend(h, l, c, period, multiplier)
    if len(st_arr) > 0:
        val = st_arr[-1]
        return float(val) if not np.isnan(val) else ""
    return ""


def compute_bollinger(close_arr: np.ndarray, period=20, std_dev=2):
    """Computes Bollinger Bands. Returns: (upper, lower, mid, width%) arrays."""
    return nk.bbands(close_arr, period, std_dev)


def compute_stochastic(h, l, c, k_period=14, d_period=3):
    """Computes Stochastic %K and %D. Returns: (k_vals, d_vals) arrays."""
    return nk.stochastic(h, l, c, k_period, d_period)


def detect_cross_aligned(fast: np.ndarray, slow: np.ndarray) -> str:
    """Detects EMA cross direction. Returns: 'Bullish', 'Bearish', or ''."""
    res = nk.detect_ema_cross(fast, slow)
    if res == 1: return "Bullish"
    if res == -1: return "Bearish"
    return ""


def detect_trend(fast: np.ndarray, slow: np.ndarray) -> str:
    """Determines trend state from EMA pair. Returns: 'Bullish', 'Bearish', or 'Neutral'."""
    if len(fast) == 0: return ""
    res = nk.detect_trend_state(fast[-1], slow[-1])
    if res == 1: return "Bullish"
    if res == -1: return "Bearish"
    return "Neutral"


def compute_er(close_arr: np.ndarray, period: int = 10) -> float:
    """
    Computes Kaufman's Efficiency Ratio (latest value).
    Returns: float (0.0 to 1.0) or 0.0 if insufficient data.
    """
    if len(close_arr) < period + 1:
        return 0.0
    er_arr = nk.efficiency_ratio(close_arr, period)
    if len(er_arr) > 0 and not np.isnan(er_arr[-1]):
        return float(er_arr[-1])
    return 0.0


# ===============================
# SPIKE DETECTION (with Timeframe Escalation)
# ===============================
def compute_spike_directional(vol_arr: np.ndarray, open_arr: np.ndarray, close_arr: np.ndarray,
                               avg_5m_baseline: float = 0.0, is_opening: bool = True,
                               is_market_open: bool = True, daily_vol: np.ndarray = None) -> int:
    """
    Determines directional volume spike with timeframe escalation.

    Reads:   vol_arr (5m volume), open/close arrays, daily volume
    Computes: Robust Z-Score spike detection
    Returns: 1 (bullish spike), -1 (bearish spike), 0 (no spike)
    Refers:  numba_kernels.detect_volume_spike for core detection

    Timeframe Escalation:
    - Market hours: Uses 5m candle volume (real-time, actionable)
    - Post-market:  Uses daily volume (avoids closing auction false spikes)
    """
    # Post-market: use daily volume
    if not is_market_open and daily_vol is not None and len(daily_vol) >= 2:
        if nk.detect_volume_spike(daily_vol, 0.0, False):
            curr_o = open_arr[-1] if len(open_arr) > 0 else 0.0
            curr_c = close_arr[-1] if len(close_arr) > 0 else 0.0
            if curr_c > curr_o: return 1
            if curr_c < curr_o: return -1
        return 0

    # Market hours: use 5m candle volume
    if len(vol_arr) < 2: return 0
    if nk.detect_volume_spike(vol_arr, avg_5m_baseline, is_opening):
        curr_o = open_arr[-1]
        curr_c = close_arr[-1]
        if curr_c > curr_o: return 1
        if curr_c < curr_o: return -1
    return 0


# ===============================
def candlestick_pattern(opens: np.ndarray, highs: np.ndarray,
                         lows: np.ndarray, closes: np.ndarray,
                         doji_threshold: float = DOJI_THRESHOLD,
                         atr: float = 0.0):
    """
    Returns (bullish_pattern, bearish_pattern).
    Only ONE direction gets a non-empty value unless the pattern is neutral (Doji/Spinning Top).
    """
    if len(closes) < 5:
        return "", ""
    
    try:
        c_o = opens[-1]; c_h = highs[-1]; c_l = lows[-1]; c_c = closes[-1]
        c_body = abs(c_c - c_o)
        c_range = c_h - c_l
        if c_range <= 0:
            return "", ""
        
        c_upper_wick = c_h - max(c_o, c_c)
        c_lower_wick = min(c_o, c_c) - c_l
        c_is_green = c_c > c_o
        c_is_red   = c_c < c_o
        b2r = c_body / c_range  # body-to-range
        
        p_o = opens[-2]; p_h = highs[-2]; p_l = lows[-2]; p_c = closes[-2]
        p_body = abs(p_c - p_o)
        p_is_red   = p_c < p_o
        p_is_green = p_c > p_o
        p_range = max(p_h - p_l, 0.001)
        p_b2r = p_body / p_range
        
        wl_ratio = c_lower_wick / c_range  # lower wick ratio
        wu_ratio = c_upper_wick / c_range  # upper wick ratio
        
        # ---- STEP 1: Classify candle direction from its own structure ----
        direction = None  # "BULL", "BEAR", or "NEUTRAL"
        
        # Is it a Doji? (body < 3% of range)
        if b2r <= doji_threshold:
            if wl_ratio > 0.55 and wu_ratio < 0.05:
                direction = "BULL"    # Dragonfly Doji
            elif wu_ratio > 0.55 and wl_ratio < 0.05:
                direction = "BEAR"    # Gravestone Doji
            else:
                direction = "NEUTRAL" # Standard Doji
        # Strong directional bar (large body, small opposite wick)
        elif b2r > 0.40 and wu_ratio < 0.10 and wl_ratio < 0.10:
            direction = "BULL" if c_is_green else "BEAR"  # Marubozu
        # Hammer-like (long lower wick, small upper wick)
        elif wl_ratio > 0.55 and wu_ratio < 0.15 and b2r < 0.35:
            direction = "BULL"    # Hammer
        # Shooting Star-like (long upper wick, small lower wick)
        elif wu_ratio > 0.55 and wl_ratio < 0.15 and b2r < 0.35:
            direction = "BEAR"    # Shooting Star
        # Small body, both wicks → indecision
        elif 0.05 < b2r < 0.25 and wu_ratio > b2r and wl_ratio > b2r:
            direction = "NEUTRAL" # Spinning Top
        # Moderate body, direction from close vs open
        elif b2r > 0.05:
            direction = "BULL" if c_is_green else "BEAR"
        else:
            direction = "NEUTRAL"
        
        # ---- STEP 2: Check for 2-bar context patterns that may override ----
        bull_name = ""; bear_name = ""
        
        # 2-bar: Engulfing
        if c_is_green and p_is_red and c_o <= p_c and c_c >= p_o and c_body > p_body * 0.7:
            direction = "BULL"; bull_name = "Bullish Engulfing"
        if c_is_red and p_is_green and c_o >= p_c and c_c <= p_o and c_body > p_body * 0.7:
            direction = "BEAR"; bear_name = "Bearish Engulfing"
        
        # 2-bar: Piercing / Dark Cloud
        if direction in (None, "BULL") and c_is_green and p_is_red and b2r > 0.15:
            p_mid = (p_o + p_c) / 2.0
            if c_o < p_l and c_c > p_mid:
                direction = "BULL"; bull_name = "Piercing Line"
        if direction in (None, "BEAR") and c_is_red and p_is_green and b2r > 0.15:
            p_mid = (p_o + p_c) / 2.0
            if c_o > p_h and c_c < p_mid:
                direction = "BEAR"; bear_name = "Dark Cloud Cover"
        
        # 2-bar: Harami (override only if it's the strongest signal)
        p_body_high = max(p_o, p_c); p_body_low = min(p_o, p_c)
        c_body_high = max(c_o, c_c); c_body_low = min(c_o, c_c)
        is_harami = (c_body_high <= p_body_high and c_body_low >= p_body_low and 
                     c_body < p_body * 0.7 and p_b2r > 0.15)
        if is_harami:
            if p_is_red and c_is_green:
                direction = "BULL"; bull_name = "Bullish Harami"
            elif p_is_green and c_is_red:
                direction = "BEAR"; bear_name = "Bearish Harami"
        
        # 2-bar: Tweezer
        if abs(c_l - p_l) < c_range * 0.08 and p_is_red and c_is_green:
            direction = "BULL"; bull_name = "Tweezer Bottom"
        if abs(c_h - p_h) < c_range * 0.08 and p_is_green and c_is_red:
            direction = "BEAR"; bear_name = "Tweezer Top"
        
        # ---- STEP 3: 3-bar patterns (override everything) ----
        b1o = opens[-3]; b1c = closes[-3]; b1h = highs[-3]; b1l = lows[-3]
        b2o = opens[-2]; b2c = closes[-2]
        b3o = opens[-1]; b3c = closes[-1]
        b1_body = abs(b1c - b1o); b2_body = abs(b2c - b2o); b3_body = abs(b3c - b3o)
        b1r = b1_body / max(b1h - b1l, 0.001) if (b1h - b1l) > 0 else 0
        b3r = b3_body / max(c_range, 0.001) if c_range > 0 else 0
        b1_red = b1c < b1o; b1_green = b1c > b1o
        b2_green = b2c > b2o; b2_red = b2c < b2o
        b3_green = b3c > b3o; b3_red = b3c < b3o
        
        # Three White Soldiers
        if b1_green and b2_green and b3_green and b3c > b2c > b1c:
            if b2o >= b1o and b2o <= b1c and b3o >= b2o and b3o <= b2c:
                direction = "BULL"; bull_name = "Three White Soldiers"
        # Three Black Crows
        if b1_red and b2_red and b3_red and b3c < b2c < b1c:
            if b2o <= b1o and b2o >= b1c and b3o <= b2o and b3o >= b2c:
                direction = "BEAR"; bear_name = "Three Black Crows"
        # Morning Star (b1 large red, b2 small, b3 large green)
        if b1_red and b3_green and b1r > 0.2 and b3r > 0.2 and b2_body < b1_body * 0.5:
            if b3c >= (b1o + b1c) / 2:
                direction = "BULL"; bull_name = "Morning Star"
        # Evening Star (b1 large green, b2 small, b3 large red)
        if b1_green and b3_red and b1r > 0.2 and b3r > 0.2 and b2_body < b1_body * 0.5:
            if b3c <= (b1o + b1c) / 2:
                direction = "BEAR"; bear_name = "Evening Star"
        
        # ---- STEP 4: Assign pattern names from direction ----
        if direction == "BULL":
            if bull_name:
                return (bull_name, "")
            # Base direction name
            if b2r <= doji_threshold:
                return ("Dragonfly Doji", "")
            elif b2r > 0.40 and wu_ratio < 0.10:
                return ("Marubozu", "")
            elif wl_ratio > 0.55:
                return ("Hammer", "")
            else:
                return ("", "")  # Bullish bias but no specific pattern
        elif direction == "BEAR":
            if bear_name:
                return ("", bear_name)
            if b2r <= doji_threshold:
                return ("", "Gravestone Doji")
            elif b2r > 0.40 and wl_ratio < 0.10:
                return ("", "Marubozu")
            elif wu_ratio > 0.55:
                return ("", "Shooting Star")
            else:
                return ("", "")  # Bearish bias but no specific pattern
        elif direction == "NEUTRAL":
            return ("Doji" if b2r <= doji_threshold else "Spinning Top", 
                    "Doji" if b2r <= doji_threshold else "Spinning Top")
        else:
            return ("", "")
            
    except Exception:
        pass
        
    return "", ""

# =====================================
# EXTENDED INDICATORS (RVOL, Swing, Fibs, Regime)
def compute_extended_indicators(
    h, l, c, v,
    is_index: bool = False,
    times=None,
    df_1d: pd.DataFrame = None,
    open_arr=None # New Optional Argument for Spike
) -> Dict[str, Any]:
    """
    Computes extended indicator set: RVOL, market regime, swing levels, Fibonacci.

    Computes extended indicator set: RVOL, market regime, swing levels, Fibonacci.

    Reads:   h, l, c, v (5m OHLCV arrays), df_1d (daily DataFrame)
    Computes: RVOL (vs 20-day avg), swing high/low, Fibonacci retracements,
              VWAP chop detection, market regime (OPEN_RUSH/DOLDRUMS/CLOSE)
    Returns: Dict with keys: rvol, avg_5m_vol, regime, swing_high, swing_low, fibs, vwap_chop
    """
    out = {
        "rvol": 0.0,
        "avg_5m_vol": 0.0,
        "regime": "",
        "swing_high": 0.0,
        "swing_low": 0.0,
        "fibs": [],
        "vwap_chop": False
    }

    if len(c) == 0:
        return out

    # --- MARKET REGIME ---
    now = datetime.now()
    t = now.time()
    if dtime(9, 15) <= t < dtime(10, 15):
        out["regime"] = "OPEN_RUSH"
    elif dtime(10, 15) <= t < dtime(14, 15):
        out["regime"] = "DOLDRUMS"
    elif dtime(14, 15) <= t < dtime(15, 10):
        out["regime"] = "LATE_SESSION"
    elif t >= dtime(15, 10):
        out["regime"] = "CLOSE"
    else:
        out["regime"] = "OFF_MARKET"

    # --- RVOL (Relative Volume vs 20-day average) ---
    if not is_index and df_1d is not None and not df_1d.empty:
        try:
            # Use the latest available date in the dataframe (for weekends/holidays)
            last_date = df_1d.index[-1]
            
            # If the last available date is today, use live intraday scaling
            if last_date.date() == now.date():
                now_mins = now.hour * 60 + now.minute
                elapsed_minutes = now_mins - 555  # 555 = 9:15 AM in minutes
                if elapsed_minutes < 0: elapsed_minutes = 1 # pre-market 
                frac_day = min(1.0, elapsed_minutes / 375.0)
                if frac_day <= 0.001: frac_day = 0.001
            else:
                # We are running after-hours or on a weekend. Use full day.
                frac_day = 1.0

            curr_vol = df_1d.loc[last_date, "volume"]
            hist = df_1d.iloc[-21:-1]  # last 20 completed days before the last date
            if not hist.empty:
                avg_daily_vol = hist["volume"].mean()
                out["avg_5m_vol"] = avg_daily_vol / 75.0
                expected_vol_now = avg_daily_vol * frac_day
                if expected_vol_now > 0:
                    out["rvol"] = round(float(curr_vol / expected_vol_now), 2)
            else:
                out["rvol"] = 1.0
        except Exception:
            pass

    # --- SWING HIGH/LOW + FIBONACCI LEVELS ---
    try:
        recent_h = h[-20:]
        recent_l = l[-20:]
        sh = np.max(recent_h)
        sl = np.min(recent_l)
        out["swing_high"] = sh
        out["swing_low"] = sl

        diff = sh - sl
        if diff > 0:
            out["fibs"] = {
                str(level): sl + level * diff for level in FIB_LEVELS
            }
    except Exception as e:
        logger.error(f"[SWING_FAIL] Len={len(h) if 'h' in locals() else 'nan'} Err={e}")

    # --- VWAP CHOP DETECTION ---
    if not is_index:
        try:
            vwap_series = nk.vwap(h, l, c, v)
            val = vwap_series[-1]
            if not np.isnan(val) and val > 0:
                if abs((c[-1] - val) / val) * 100 < 0.15:
                    out["vwap_chop"] = True
        except Exception:
            pass

    # --- VOLUME SPIKE (v14 Baseline Parity) ---
    try:
        from engine.indicators import compute_spike_directional
        from engine.config import SCRIPT_NAME 
        
        # We need to compute spike for the result dict
        # Ensure we have a valid open array
        if open_arr is None:
            open_arr = h - l # Fallback if not provided (less accurate)

        out["spike"] = compute_spike_directional(
            v, open_arr, c, # Correct order: vol, open, close
            avg_5m_baseline=out.get("avg_5m_vol", 0.0),
            is_opening=(datetime.now().hour == 9 and datetime.now().minute < 45),
            is_market_open=True, # default to True for intraday
            daily_vol=df_1d['volume'].values if df_1d is not None else None
        )
    except Exception as e:
        logger.error(f"Spike Calc Failed: {e}")
    return out

# ===============================
# GENERIC INDICATORS (All-in-one per timeframe)
# ===============================
def compute_generic_indicators(data_dict, include_rsi=True, include_macd=True,
                                include_st=True, doji_threshold=DOJI_THRESHOLD):
    """
    Computes the standard indicator set for a single timeframe.

    Reads:   data_dict with keys: open, high, low, close, volume, ghost
    Computes: ATR, Bollinger Bands, Stochastic, EMAs (10-200),
              RSI, MACD, SuperTrend, EMA crosses, candlestick patterns
    Returns: Dict with all indicator values for this timeframe
    """
    out = {}
    o = data_dict["open"]; h = data_dict["high"]; l = data_dict["low"]
    c = data_dict["close"]; v = data_dict["volume"]
    if len(c) == 0: return out

    # ATR + Bollinger
    atr_s = compute_atr(h, l, c, 14)
    bbu, bbl, bbm, bbw = compute_bollinger(c, 20, 2)

    try: out["atr"] = float(atr_s[-1])
    except: out["atr"] = 0.0
    try: out["bb_width_pct"] = float(bbw[-1])
    except: out["bb_width_pct"] = 0.0
    try: 
        out["bb_upper"] = float(bbu[-1])
        out["bb_mid"] = float(bbm[-1])
        out["bb_lower"] = float(bbl[-1])
        diff = out["bb_upper"] - out["bb_lower"]
        out["bb_position"] = (c[-1] - out["bb_lower"]) / diff if diff > 0 else 0.5
    except: 
        out["bb_position"] = 0.5

    # ADX
    p_di, m_di, adx_val, adx_slp = adx_smoothed(h, l, c)
    out["adx"] = adx_val
    out["plus_di"] = p_di
    out["minus_di"] = m_di
    out["adx_slope"] = adx_slp

    # Stochastic
    sk, sd = compute_stochastic(h, l, c, 14, 3)
    try: out["stoch_k"] = float(sk[-1]); out["stoch_d"] = float(sd[-1])
    except: out["stoch_k"] = 0.0; out["stoch_d"] = 0.0

    # EMAs (Flattened for Setup Library)
    for span in [10, 20, 30, 50, 100, 200]:
        try:
            s_arr = ema_fast(c, span)
            out[f"ema_{span}"] = float(s_arr[-1])
        except:
            out[f"ema_{span}"] = 0.0
    
    # Legacy nested emas for backward compatibility if any
    out["emas"] = {s: out[f"ema_{s}"] for s in [10, 20, 30, 50, 100, 200]}

    # RSI
    if include_rsi:
        out["rsi_series"] = rsi_wilder(c, 14)
        try: out["rsi"] = float(out["rsi_series"][-1])
        except: out["rsi"] = 50.0

    # MACD
    if include_macd:
        macd_val, ema_f_val, ema_s_val, sig_val = macd_fast(c)
        out["macd_line"] = macd_val; out["macd_signal"] = sig_val
        try:
            out["macd_fast_val"] = float(ema_f_val[-1])
            out["macd_slow_val"] = float(ema_s_val[-1])
            out["macd_sig_val"] = float(sig_val[-1])
        except: pass

    # SuperTrend
    if include_st:
        out["supertrend"] = compute_supertrend(h, l, c)

    # EMA Crosses
    e10 = ema_fast(c, 10); e_slow = ema_fast(c, EMA_CROSS_SLOW); e50 = ema_fast(c, 50)
    out["ema_cross"] = detect_cross_aligned(e10, e_slow)
    out["ema10_50_cross"] = detect_cross_aligned(e10, e50)
    out["ema10_50_state"] = detect_trend(e10, e50)

    # Candlestick Patterns (pass ATR for body sizing)
    try:
        _atr_val = float(atr_s[-1]) if len(atr_s) > 0 else 0.0
    except:
        _atr_val = 0.0
    out["bullish_pattern"], out["bearish_pattern"] = candlestick_pattern(o, h, l, c, doji_threshold, atr=_atr_val)

    # CMF (Chaikin Money Flow, 20-period) — directional volume confirmation
    try:
        hl_range = h - l
        # Use np.divide with where= to avoid computing division where hl_range==0
        # (np.where evaluates both branches, so division by zero still occurs)
        mfm = np.divide(
            ((c - l) - (h - c)),
            hl_range,
            out=np.zeros_like(c, dtype=np.float64),
            where=(hl_range > 0)
        )
        mfv = mfm * v.astype(np.float64)
        cmf_period = 20
        if len(c) >= cmf_period:
            mfv_cum = np.cumsum(mfv)
            vol_cum = np.cumsum(v.astype(np.float64))
            mfv_sum = mfv_cum[cmf_period - 1:] - np.concatenate([[0.0], mfv_cum[:-cmf_period]])
            vol_sum = vol_cum[cmf_period - 1:] - np.concatenate([[0.0], vol_cum[:-cmf_period]])
            safe_vol = np.where(vol_sum == 0, 1.0, vol_sum)
            cmf_arr = np.where(vol_sum > 0, mfv_sum / safe_vol, 0.0)
            out["cmf"] = float(cmf_arr[-1]) if len(cmf_arr) > 0 and np.isfinite(cmf_arr[-1]) else 0.0
        else:
            out["cmf"] = 0.0
    except Exception:
        out["cmf"] = 0.0

    # OBV slope (5-bar normalized slope of On-Balance Volume)
    try:
        # Vectorized OBV: sign(diff) * volume, then cumulative sum
        direction = np.sign(np.diff(c, prepend=c[0]))
        obv = np.cumsum(direction * v.astype(np.float64))
        if len(obv) >= 6:
            denom = max(abs(obv[-6]), 1.0)
            out["obv_slope"] = float((obv[-1] - obv[-6]) / denom)
        else:
            out["obv_slope"] = 0.0
    except Exception:
        out["obv_slope"] = 0.0

    return out


# ===============================
# AI FEATURE HELPERS (New for v3)
# ===============================
def compute_ema_matrix_raw(data_dict_by_tf: Dict[str, Dict], ltp: float, atr: float) -> list:
    """
    Computes the raw 4x5 EMA tensor (normalized by ATR).
    Returns a flat list of 20 floats (5m_10, 5m_20... 1d_200).
    """
    # EMA periods from config
    periods = [10, 20, 50, 100, 200]
    # Timeframes from config
    tfs = ["5m", "15m", "1h", "1d"]
    
    matrix = []
    
    if atr <= 0: atr = 1.0 # prevent div/0
    
    for tf in tfs:
        # Get indicators for this TF
        ind = data_dict_by_tf.get(tf, {})
        # Depending on how 'emas' are stored, we might descend into a dict
        # In generic_indicators, we return out["emas"] = {10: val, ...}
        ema_map = ind.get("emas", {})
        
        for p in periods:
            val = ema_map.get(p)
            
            # Check if val is a valid number
            if val is not None and isinstance(val, (int, float)) and val > 0:
                # Distance in ATR units. 
                # Positive = Price > EMA (Bullish), Negative = Price < EMA (Bearish)
                dist = (ltp - val) / atr
            else:
                dist = 0.0
                
            matrix.append(dist)
            
    return matrix


def compute_beta_features(symbol_returns: np.ndarray, 
                         market_returns: np.ndarray, 
                         window=20) -> float:
    """
    Computes Beta (slope of linear regression) between symbol and market.
    Returns: Beta float.
    """
    if len(symbol_returns) < window or len(market_returns) < window:
        return 1.0 # Default beta
        
    # Slice last N returns
    y = symbol_returns[-window:]
    x = market_returns[-window:]
    
    # Covariance / Variance
    try:
        if len(x) != len(y):
             min_len = min(len(x), len(y))
             x = x[-min_len:]
             y = y[-min_len:]
             
        cov = np.cov(x, y)[0, 1]
        var = np.var(x)
        
        if var == 0: return 1.0
        return float(cov / var)
    except:
        return 1.0
