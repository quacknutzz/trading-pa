"""
numba_kernels.py — JIT-compiled mathematical kernels.

Reads:   Raw numpy arrays (OHLCV data)
Returns: Computed indicator arrays (EMA, RSI, ATR, VWAP, etc.)
Refers:  No external dependencies — pure math only

All functions are @jit(nopython=True) for maximum performance.
These are the low-level compute engines called by indicators.py wrappers.
"""

import numpy as np

try:
    from numba import jit
    HAS_NUMBA = True
except ImportError:
    HAS_NUMBA = False

if not HAS_NUMBA:
    def jit(nopython=True, nogil=True, cache=True):
        def decorator(func):
            return func
        return decorator


# ===============================
# RESAMPLING
# ===============================
@jit(nopython=True, nogil=True, cache=True)
def resample(times, opens, highs, lows, closes, vols, interval_ns):
    """
    Resamples 5m candles into higher timeframes (15m, 1h).

    Reads:   Aligned numpy arrays of timestamps (int64 ns) and OHLCV
    Returns: (count, times, opens, highs, lows, closes, volumes) as aligned arrays
    """
    n = len(times)
    out_times = np.zeros(n, dtype=np.int64)
    out_opens = np.zeros(n, dtype=np.float64)
    out_highs = np.zeros(n, dtype=np.float64)
    out_lows = np.zeros(n, dtype=np.float64)
    out_closes = np.zeros(n, dtype=np.float64)
    out_vols = np.zeros(n, dtype=np.float64)
    if n == 0: return 0, out_times, out_opens, out_highs, out_lows, out_closes, out_vols

    out_idx = 0
    OFFSET_NS = 900_000_000_000
    t0_shifted = times[0] - OFFSET_NS
    curr_bucket_start = (t0_shifted // interval_ns) * interval_ns + OFFSET_NS

    acc_open = opens[0]; acc_high = highs[0]; acc_low = lows[0]; acc_close = closes[0]; acc_vol = vols[0]

    for i in range(1, n):
        ts = times[i]
        ts_shifted = ts - OFFSET_NS
        bucket = (ts_shifted // interval_ns) * interval_ns + OFFSET_NS

        if bucket == curr_bucket_start:
            acc_high = max(acc_high, highs[i]); acc_low = min(acc_low, lows[i])
            acc_close = closes[i]; acc_vol += vols[i]
        else:
            out_times[out_idx] = curr_bucket_start
            out_opens[out_idx] = acc_open; out_highs[out_idx] = acc_high
            out_lows[out_idx] = acc_low; out_closes[out_idx] = acc_close; out_vols[out_idx] = acc_vol
            out_idx += 1
            curr_bucket_start = bucket
            acc_open = opens[i]; acc_high = highs[i]; acc_low = lows[i]; acc_close = closes[i]; acc_vol = vols[i]

    out_times[out_idx] = curr_bucket_start
    out_opens[out_idx] = acc_open; out_highs[out_idx] = acc_high
    out_lows[out_idx] = acc_low; out_closes[out_idx] = acc_close; out_vols[out_idx] = acc_vol
    out_idx += 1
    return out_idx, out_times, out_opens, out_highs, out_lows, out_closes, out_vols


# ===============================
# EXPONENTIAL MOVING AVERAGE
# ===============================
@jit(nopython=True, nogil=True, cache=True)
def ema(values, period):
    """
    Computes Exponential Moving Average.

    Reads:   values (float64 array), period (int)
    Returns: EMA array (same length, NaN-padded at start)
    """
    n = len(values)
    out = np.full(n, np.nan, dtype=np.float64)
    if n == 0: return out
    start_idx = 0
    if n >= period:
        sma_sum = 0.0
        for i in range(period): sma_sum += values[i]
        out[period-1] = sma_sum / period
        start_idx = period
    else:
        out[0] = values[0]
        start_idx = 1
    alpha = 2.0 / (period + 1.0)
    for i in range(start_idx, n):
        prev = out[i-1]
        if np.isnan(prev): out[i] = values[i]
        else: out[i] = (values[i] * alpha) + (prev * (1.0 - alpha))
    return out


# ===============================
# RELATIVE STRENGTH INDEX
# ===============================
@jit(nopython=True, nogil=True, cache=True)
def rsi(close, length):
    """
    Computes Wilder's RSI.

    Reads:   close (float64 array), length (int, typically 14)
    Returns: RSI array (0-100 scale, NaN-padded at start)
    """
    n = len(close)
    rsi_out = np.full(n, np.nan, dtype=np.float64)
    if n <= length: return rsi_out
    gains = np.zeros(n, dtype=np.float64); losses = np.zeros(n, dtype=np.float64)
    for i in range(1, n):
        diff = close[i] - close[i-1]
        if diff > 0: gains[i] = diff
        elif diff < 0: losses[i] = -diff
    sum_gain = 0.0; sum_loss = 0.0
    for i in range(1, length + 1): sum_gain += gains[i]; sum_loss += losses[i]
    avg_gain = sum_gain / length; avg_loss = sum_loss / length
    if avg_loss == 0: rsi_out[length] = 100.0
    else: rsi_out[length] = 100.0 - (100.0 / (1.0 + (avg_gain / avg_loss)))
    for i in range(length + 1, n):
        avg_gain = ((avg_gain * (length - 1)) + gains[i]) / length
        avg_loss = ((avg_loss * (length - 1)) + losses[i]) / length
        if avg_loss == 0: rsi_out[i] = 100.0
        else: rsi_out[i] = 100.0 - (100.0 / (1.0 + (avg_gain / avg_loss)))
    return rsi_out


# ===============================
# VWAP
# ===============================
@jit(nopython=True, nogil=True, cache=True)
def vwap(h, l, c, v, t):
    """
    Computes Volume Weighted Average Price with Daily Reset.
    
    Reads:   h, l, c, v (high, low, close, volume arrays)
             t (int64 timestamp array in nanoseconds)
    Returns: VWAP array (Daily Reset)
    """
    typical_p = (h + l + c) / 3.0
    n = len(v)
    out = np.zeros(n, dtype=np.float64)
    v_acc = 0.0
    pv_acc = 0.0
    
    # 86,400,000,000,000 ns in a day
    NS_IN_DAY = 86400000000000
    
    for i in range(n):
        # Reset if day changes (Standard Session VWAP logic)
        if i > 0:
            # We add 19800s (IST offset) to ensure boundary falls at Indian midnight, not UTC
            t_curr = t[i] + 19800000000000
            t_prev = t[i-1] + 19800000000000
            if (t_curr // NS_IN_DAY) != (t_prev // NS_IN_DAY):
                v_acc = 0.0
                pv_acc = 0.0
        
        v_acc += v[i]
        pv_acc += (typical_p[i] * v[i])
        
        if v_acc != 0:
            out[i] = pv_acc / v_acc
        else:
            out[i] = typical_p[i]
            
    return out


# ===============================
# BOLLINGER BANDS
# ===============================
@jit(nopython=True, nogil=True, cache=True)
def bbands(close, period, std_dev):
    """
    Computes Bollinger Bands using Welford's online algorithm.
    O(n) single-pass instead of O(n*period) nested loop.

    Reads:   close array, period (20), std_dev (2.0)
    Returns: (upper, lower, mid, width%) arrays
    """
    n = len(close)
    upper = np.full(n, np.nan, dtype=np.float64); lower = np.full(n, np.nan, dtype=np.float64)
    mid = np.full(n, np.nan, dtype=np.float64); width = np.full(n, np.nan, dtype=np.float64)
    if n < period: return upper, lower, mid, width
    # Welford's online algorithm: incremental mean + variance
    mean = 0.0; m2 = 0.0
    for i in range(period):
        x = close[i]; delta = x - mean; mean += delta / (i + 1); m2 += delta * (x - mean)
    var = m2 / (period - 1)
    std = np.sqrt(var) if var > 0.0 else 0.0
    mid[period - 1] = mean; upper[period - 1] = mean + std * std_dev
    lower[period - 1] = mean - std * std_dev
    if mean != 0: width[period - 1] = (std * std_dev * 2 / mean) * 100.0
    for i in range(period, n):
        x_old = close[i - period]; x_new = close[i]
        mean = mean + (x_new - x_old) / period
        m2 = m2 + (x_new - x_old) * (x_new - mean + x_old - mean)
        var = m2 / (period - 1)
        std = np.sqrt(var) if var > 0.0 else 0.0
        mid[i] = mean; upper[i] = mean + std * std_dev; lower[i] = mean - std * std_dev
        if mean != 0: width[i] = (std * std_dev * 2 / mean) * 100.0
    return upper, lower, mid, width


# ===============================
# AVERAGE TRUE RANGE
# ===============================
@jit(nopython=True, nogil=True, cache=True)
def atr(high, low, close, period):
    """
    Computes ATR using Wilder's smoothing.
    """
    n = len(close)
    atr_out = np.full(n, np.nan, dtype=np.float64)
    if n == 0: return atr_out
    tr = np.zeros(n, dtype=np.float64)
    tr[0] = high[0] - low[0]
    for i in range(1, n):
        tr[i] = max(high[i] - low[i], max(abs(high[i] - close[i-1]), abs(low[i] - close[i-1])))
    alpha = 1.0 / period
    val = tr[0]; atr_out[0] = val
    for i in range(1, n): val = (tr[i] * alpha) + (val * (1.0 - alpha)); atr_out[i] = val
    return atr_out


@jit(nopython=True, nogil=True, cache=True)
def adx(high, low, close, period):
    """
    Computes Average Directional Index (ADX) with +DI and -DI.
    Returns: (plus_di, minus_di, adx) arrays
    """
    n = len(close)
    plus_di = np.full(n, np.nan, dtype=np.float64)
    minus_di = np.full(n, np.nan, dtype=np.float64)
    adx_out = np.full(n, np.nan, dtype=np.float64)
    
    if n < period * 2: return plus_di, minus_di, adx_out
    
    tr = np.zeros(n, dtype=np.float64)
    dm_plus = np.zeros(n, dtype=np.float64)
    dm_minus = np.zeros(n, dtype=np.float64)
    
    for i in range(1, n):
        tr[i] = max(high[i] - low[i], max(abs(high[i] - close[i-1]), abs(low[i] - close[i-1])))
        up_move = high[i] - high[i-1]
        down_move = low[i-1] - low[i]
        
        if up_move > down_move and up_move > 0: dm_plus[i] = up_move
        else: dm_plus[i] = 0.0
            
        if down_move > up_move and down_move > 0: dm_minus[i] = down_move
        else: dm_minus[i] = 0.0
            
    # Smoothing
    atr_smooth = np.zeros(n)
    dm_plus_smooth = np.zeros(n)
    dm_minus_smooth = np.zeros(n)
    
    # First SMA
    sum_tr = 0.0; sum_p = 0.0; sum_m = 0.0
    for i in range(1, period + 1):
        sum_tr += tr[i]; sum_p += dm_plus[i]; sum_m += dm_minus[i]
        
    atr_smooth[period] = sum_tr 
    dm_plus_smooth[period] = sum_p
    dm_minus_smooth[period] = sum_m
    
    alpha = 1.0 / period
    for i in range(period + 1, n):
        atr_smooth[i] = atr_smooth[i-1] - (atr_smooth[i-1] / period) + tr[i]
        dm_plus_smooth[i] = dm_plus_smooth[i-1] - (dm_plus_smooth[i-1] / period) + dm_plus[i]
        dm_minus_smooth[i] = dm_minus_smooth[i-1] - (dm_minus_smooth[i-1] / period) + dm_minus[i]
        
        if atr_smooth[i] != 0:
            plus_di[i] = 100.0 * (dm_plus_smooth[i] / atr_smooth[i])
            minus_di[i] = 100.0 * (dm_minus_smooth[i] / atr_smooth[i])
            
            # DX
            sum_di = plus_di[i] + minus_di[i]
            if sum_di > 1e-6:
                dx = 100.0 * abs(plus_di[i] - minus_di[i]) / sum_di
            else:
                dx = 0.0
            
            # ADX Smoothing (Wilder's)
            if i == period * 2 - 1:
                adx_out[i] = dx 
            elif i >= period * 2:
                # Recursive Wilder's Smoothing
                prev_adx = adx_out[i-1]
                if np.isnan(prev_adx) or prev_adx == 0:
                    adx_out[i] = dx
                else:
                    adx_out[i] = (prev_adx * (period - 1) + dx) / period
            
    # Final pass: fill NaNs with 0 to prevent downstream failures
    for i in range(n):
        if np.isnan(adx_out[i]): adx_out[i] = 0.0
        if np.isnan(plus_di[i]): plus_di[i] = 0.0
        if np.isnan(minus_di[i]): minus_di[i] = 0.0
        
    return plus_di, minus_di, adx_out
                    
    return plus_di, minus_di, adx_out


@jit(nopython=True, nogil=True, cache=True)
def percentile_rank(values, lookback):
    """
    Computes Percentile Rank relative to lookback.
    Returns: Array of values from 0.0 to 1.0
    """
    n = len(values)
    out = np.full(n, np.nan, dtype=np.float64)
    if n < lookback: return out
    
    for i in range(lookback - 1, n):
        curr = values[i]
        count = 0
        valid = 0
        for j in range(lookback):
            val = values[i-j]
            if not np.isnan(val):
                valid += 1
                if val < curr: count += 1
        if valid > 1:
            out[i] = count / (valid - 1)
    return out


# ===============================
# SUPERTREND
# ===============================
@jit(nopython=True, nogil=True, cache=True)
def supertrend(high, low, close, period, mult):
    """
    Computes SuperTrend indicator.

    Reads:   high, low, close arrays, period (10), multiplier (3.0)
    Returns: Last SuperTrend value (scalar float)
    """
    n = len(close)
    st = np.zeros(n, dtype=np.float64)
    if n < period:
        return np.full(n, np.nan, dtype=np.float64)

    tr = np.zeros(n, dtype=np.float64); tr[0] = high[0] - low[0]
    for i in range(1, n):
        tr[i] = max(high[i] - low[i], max(abs(high[i] - close[i-1]), abs(low[i] - close[i-1])))
    atr_arr = np.zeros(n, dtype=np.float64)
    alpha = 1.0 / period
    val = tr[0] if not np.isnan(tr[0]) else 0.0
    atr_arr[0] = val
    for i in range(1, n): val = (tr[i] * alpha) + (val * (1.0 - alpha)); atr_arr[i] = val
    upper_final = np.zeros(n, dtype=np.float64); lower_final = np.zeros(n, dtype=np.float64)
    
    trend = 1
    hl2 = (high[0] + low[0]) / 2.0
    upper_final[0] = hl2 + mult * atr_arr[0]; lower_final[0] = hl2 - mult * atr_arr[0]
    st[0] = lower_final[0]
    for i in range(1, n):
        hl2 = (high[i] + low[i]) / 2.0
        curr_atr = atr_arr[i] if not np.isnan(atr_arr[i]) else 0.0
        ub = hl2 + mult * curr_atr; lb = hl2 - mult * curr_atr
        if ub < upper_final[i-1] or close[i-1] > upper_final[i-1]: upper_final[i] = ub
        else: upper_final[i] = upper_final[i-1]
        if lb > lower_final[i-1] or close[i-1] < lower_final[i-1]: lower_final[i] = lb
        else: lower_final[i] = lower_final[i-1]
        if trend == 1:
            if close[i] < lower_final[i]: trend = -1; st[i] = upper_final[i]
            else: trend = 1; st[i] = lower_final[i]
        else:
            if close[i] > upper_final[i]: trend = 1; st[i] = lower_final[i]
            else: trend = -1; st[i] = upper_final[i]
    return st


# ===============================
# STOCHASTIC OSCILLATOR
# ===============================
@jit(nopython=True, nogil=True, cache=True)
def stochastic(high, low, close, k_period, d_period):
    """
    Computes Stochastic %K and %D.

    Reads:   high, low, close arrays, k_period (14), d_period (3)
    Returns: (k_vals, d_vals) arrays
    """
    n = len(close)
    k_vals = np.full(n, np.nan, dtype=np.float64); d_vals = np.full(n, np.nan, dtype=np.float64)
    if n < k_period: return k_vals, d_vals
    for i in range(k_period - 1, n):
        low_min = low[i]; high_max = high[i]
        for j in range(1, k_period):
            if low[i-j] < low_min: low_min = low[i-j]
            if high[i-j] > high_max: high_max = high[i-j]
        if high_max == low_min: k_vals[i] = 50.0
        else: k_vals[i] = 100.0 * ((close[i] - low_min) / (high_max - low_min))
    for i in range(k_period + d_period - 2, n):
        s = 0.0; c = 0
        for j in range(d_period):
            idx = i - j
            if not np.isnan(k_vals[idx]): s += k_vals[idx]; c += 1
        if c > 0: d_vals[i] = s / c
    return k_vals, d_vals


# ===============================
# EMA CROSS DETECTION
# ===============================
@jit(nopython=True, nogil=True, cache=True)
def detect_ema_cross(fast, slow):
    """
    Detects EMA crossover between fast and slow lines.

    Reads:   fast, slow EMA arrays
    Returns: 1 (bullish cross), -1 (bearish cross), 0 (no cross)
    """
    n = len(fast)
    if n < 2: return 0
    f_curr, s_curr = fast[n-1], slow[n-1]
    f_prev, s_prev = fast[n-2], slow[n-2]
    if np.isnan(f_curr) or np.isnan(s_curr) or np.isnan(f_prev) or np.isnan(s_prev): return 0
    diff_curr = f_curr - s_curr
    diff_prev = f_prev - s_prev
    if diff_prev <= 0 and diff_curr > 0: return 1
    if diff_prev >= 0 and diff_curr < 0: return -1
    return 0


# ===============================
# VOLUME SPIKE DETECTION
# ===============================
@jit(nopython=True, nogil=True, cache=True)
def detect_volume_spike(volume, avg_5m_baseline=0.0, is_opening=True):
    """
    Robust Volume Spike Detection (Industry-Grade).

    Method:  Robust Z-Score = (curr - Median) / (MAD * 1.4826)
    Reads:   volume array, optional baseline, opening flag
    Returns: 1 (spike detected), 0 (normal volume)
    Refers:  Z >= 2.0 threshold (97.7th percentile)
    """
    n = len(volume)
    if n < 2: return 0

    curr = volume[n-1]
    if curr <= 0: return 0

    # Short history fallback
    if n < 20:
        prev = volume[n-2]
        baseline = max(prev, avg_5m_baseline) if avg_5m_baseline > 0 else prev
        if baseline <= 0: baseline = 1.0
        return 1 if curr >= 2.0 * baseline else 0

    # Robust statistics (Median & MAD)
    lookback = min(50, n-1)
    past = volume[n-1-lookback : n-1]
    median_val = np.median(past)

    abs_diffs = np.zeros(len(past))
    for i in range(len(past)):
        abs_diffs[i] = abs(past[i] - median_val)
    mad = np.median(abs_diffs)

    # Standard MAD scaling to estimate Sigma
    sigma = mad * 1.4826
    if sigma < 1.0: sigma = 1.0

    # Z-Score relative to localized median
    zscore = (curr - median_val) / sigma

    if zscore >= 2.0:
        return 1
    return 0


# ===============================
# EFFICIENCY RATIO (ER)
# ===============================
@jit(nopython=True, nogil=True, cache=True)
def efficiency_ratio(close, period):
    """
    Computes Kaufman's Efficiency Ratio (ER).
    
    Formula: abs(Price_Change) / Sum_of_Absolute_Changes
    ER = 1.0 (Strong Trend), ER = 0.0 (Pure Noise/Chop)
    
    Reads:   close array, period (int)
    Returns: ER array (NaN-padded at start)
    """
    n = len(close)
    er_out = np.full(n, np.nan, dtype=np.float64)
    if n < period + 1:
        return er_out
        
    for i in range(period, n):
        # 1. Price Change over the period
        change = abs(close[i] - close[i - period])
        
        # 2. Sum of absolute candle-to-candle changes (Volatility)
        volatility = 0.0
        for j in range(i - period + 1, i + 1):
            volatility += abs(close[j] - close[j-1])
            
        if volatility != 0:
            er_out[i] = change / volatility
        else:
            er_out[i] = 0.0
            
    return er_out


# ===============================
# TREND STATE DETECTION
# ===============================
@jit(nopython=True, nogil=True, cache=True)
def detect_trend_state(fast_val, slow_val):
    """
    Determines trend state from two EMA values.

    Reads:   fast_val, slow_val (scalar floats)
    Returns: 1 (bullish), -1 (bearish), 0 (neutral/NaN)
    """
    if np.isnan(fast_val) or np.isnan(slow_val): return 0
    if fast_val > slow_val: return 1
    if fast_val < slow_val: return -1
    return 0
