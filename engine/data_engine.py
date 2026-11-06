"""
data_engine.py — Data infrastructure: sync, fetch, storage, and Fyers API integration.

Reads:   Fyers API (history + quotes), parquet files on disk, manifest JSON
Writes:  Parquet files (5m + daily), manifest JSON, logs
Computes: Data sync decisions (bootstrap, gap fill, seal, scrub)
Refers:  config.py for paths, credentials, API rate limits
"""

import os
import json
import time
import logging
import random
import threading
from datetime import datetime, date, timedelta, time as dtime
from typing import Dict, List, Optional, Tuple
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import h5py

from engine.config import (
    ROOT_DIR_ENV, DEFAULT_ROOT,
    PARQUET_SUBFOLDER, DAILY_SUBFOLDER, MANIFEST_FILENAME,
    TOKEN_FILENAME, FYERS_INDEX_MAP, API_RATE_LIMIT,
    HARDCODED_HOLIDAYS, EXCEPTION_TRADING_DAYS,
    BOOTSTRAP_5M_DAYS,
)
BOOTSTRAP_1D_DAYS = 1825
from engine.logger import ensure_dir


logger = logging.getLogger("ai_signal_engine")

# --- Lazy imports for optional dependencies ---
try:
    from fyers_apiv3 import fyersModel
except Exception:
    fyersModel = None


# ===============================
# PATH HELPERS
# ===============================
def get_root_folder() -> str:
    """Returns the project root directory (AI/ folder)."""
    return os.environ.get(ROOT_DIR_ENV) or DEFAULT_ROOT


def master_folder_5m(root: Optional[str] = None) -> str:
    """Returns path to 5m parquet folder, creating if needed."""
    p = os.path.join(root or get_root_folder(), PARQUET_SUBFOLDER)
    ensure_dir(p)
    return p


def master_folder_daily(root: Optional[str] = None) -> str:
    """Returns path to daily parquet folder, creating if needed."""
    p = os.path.join(root or get_root_folder(), DAILY_SUBFOLDER)
    ensure_dir(p)
    return p


def hdf5_path(interval: str = "5m", root: Optional[str] = None) -> str:
    """Returns the path to the HDF5 store for the given interval."""
    filename = f"market_data_{interval}.h5"
    if interval == "1d":
        return os.path.join(master_folder_daily(root), filename)
    return os.path.join(master_folder_5m(root), filename)


def parquet_path_5m(symbol: str, root: Optional[str] = None) -> str:
    """Returns full path for a symbol's 5m parquet file."""
    safe = symbol.replace(".", "_").replace("^", "")
    if not safe.endswith("_NS"):
        safe += "_NS"
    return os.path.join(master_folder_5m(root), f"{safe}.parquet")


def parquet_path_daily(symbol: str, root: Optional[str] = None) -> str:
    """Returns full path for a symbol's daily parquet file."""
    safe = symbol.replace(".", "_").replace("^", "")
    if not safe.endswith("_NS"):
        safe += "_NS"
    return os.path.join(master_folder_daily(root), f"{safe}_1d.parquet")


def is_market_open_now(now: Optional[datetime] = None) -> bool:
    """
    Checks if the Indian stock market is currently open.

    Reads:   System clock, HARDCODED_HOLIDAYS, EXCEPTION_TRADING_DAYS
    Returns: True if market is open (9:15-15:30 on trading days)
    """
    now = now or datetime.now().replace(tzinfo=None)
    today = now.strftime("%Y-%m-%d")

    is_exception_day = today in EXCEPTION_TRADING_DAYS
    is_regular_trading_day = (
        now.weekday() < 5 and today not in HARDCODED_HOLIDAYS
    )

    if not (is_exception_day or is_regular_trading_day):
        return False

    return dtime(9, 15) <= now.time() < dtime(15, 30)


def is_market_holiday_today(now: Optional[datetime] = None) -> bool:
    """Returns True if today is a scheduled market holiday or weekend (non-exception)."""
    now = now or datetime.now()
    today = now.strftime("%Y-%m-%d")
    is_exception_day = today in EXCEPTION_TRADING_DAYS
    if is_exception_day: return False
    return (now.weekday() >= 5) or (today in HARDCODED_HOLIDAYS)


def get_last_trading_date(ref_date: datetime) -> datetime:
    """
    Returns the most recent trading day up to and including ref_date.
    """
    d = ref_date
    while True:
        today_str = d.strftime("%Y-%m-%d")
        is_exception = today_str in EXCEPTION_TRADING_DAYS
        is_regular = d.weekday() < 5 and today_str not in HARDCODED_HOLIDAYS
        if is_exception or is_regular:
            return d
        d -= timedelta(days=1)


def get_n_trading_days_prior(ref_date, n: int = 1) -> date:
    """
    Returns the trading date that is `n` trading days strictly prior to `ref_date`.
    Skips weekends and HARDCODED_HOLIDAYS (respecting EXCEPTION_TRADING_DAYS).
    """
    d = ref_date if isinstance(ref_date, date) else ref_date.date()
    count = 0
    curr = d - timedelta(days=1)
    while True:
        today_str = curr.strftime("%Y-%m-%d")
        is_exception = today_str in EXCEPTION_TRADING_DAYS
        is_regular = curr.weekday() < 5 and today_str not in HARDCODED_HOLIDAYS
        if is_exception or is_regular:
            count += 1
            if count >= n:
                return curr
        curr -= timedelta(days=1)


def get_last_completed_trading_day(now: datetime) -> date:
    """
    Returns the last completed trading day session.
    - On weekends or holidays: the most recent past trading day.
    - On regular trading days:
        - Before 17:00: the previous trading day (today is either pre-market, live, or in settlement).
        - At or after 17:00: today (market and settlement completed, ready for post-market seal).
    """
    is_holiday = is_market_holiday_today(now)
    is_weekend = now.weekday() >= 5
    if is_holiday or is_weekend:
        return get_last_trading_date(now - timedelta(days=1)).date()
    
    if now.time() < dtime(17, 0):
        return get_last_trading_date(now - timedelta(days=1)).date()
    else:
        return now.date()


def resolve_seal_sync_range(last_scrub_date_str: str, target_date: date) -> Tuple[date, date]:
    """
    Computes the [start_date, end_date] range to seal a gap.
    Rule: Gap + previous 2 trading days.
    If last_scrub_date_str is valid:
        start_date = 1 trading day prior to last_scrub_date (so it includes last_scrub_date and the day before it).
        end_date = target_date.
    If last_scrub_date_str is missing:
        fallback to 5 trading days prior to target_date.
    """
    if last_scrub_date_str:
        try:
            scrub_dt = datetime.strptime(last_scrub_date_str, "%Y-%m-%d").date()
            start_date = get_n_trading_days_prior(scrub_dt, n=1)
            if start_date > target_date:
                start_date = target_date
            return start_date, target_date
        except Exception:
            pass
    fallback_start = get_n_trading_days_prior(target_date, n=5)
    return fallback_start, target_date


def get_last_closed_5m(now: datetime) -> datetime:
    """
    Returns timestamp of the last fully closed 5m candle.
    Uses BROKER_SETTLE_BUFFER to allow for exchange aggregation latency.
    """
    from engine.config import BROKER_SETTLE_BUFFER
    # Subtract buffer to ensure we only request a candle once it has settled
    adj_now = now - timedelta(seconds=BROKER_SETTLE_BUFFER)
    base = adj_now.replace(second=0, microsecond=0)
    minute = (base.minute // 5) * 5
    return base.replace(minute=minute)


# ===============================
# RATE LIMITER
# ===============================
class RateLimiter:
    """
    Token-bucket rate limiter for API call throttling.

    Reads:   System clock
    Computes: Token availability for API call pacing
    """
    def __init__(self, calls_per_second=API_RATE_LIMIT):
        self.rate = calls_per_second
        #self.tokens = calls_per_second
        self.tokens = self.tokens = 1.0
        self.lock = threading.Lock()
        self.last_check = time.time()

    def wait(self):
        sleep_time = 0
        with self.lock:
            now = time.time()
            elapsed = now - self.last_check
            self.tokens = min(self.rate, self.tokens + elapsed * self.rate)
            self.last_check = now
            if self.tokens < 1:
                needed = 1 - self.tokens
                sleep_time = needed / self.rate
                self.tokens -= 1
            else:
                self.tokens -= 1
        if sleep_time > 0:
            if sleep_time > 0.05: # Only log significant pacing
                logger.debug(f"[LIMITER] Pacing: sleeping {sleep_time:.3f}s (Current rate: {self.rate}/s)")
            time.sleep(sleep_time)


def backoff_sleep(attempt: int, base_delay: float = 1.0, max_delay: float = 10.0):
    """
    Implements Institutional Exponential Backoff with Jitter (Option C).
    - attempt: Current retry attempt (1-indexed)
    - base_delay: Initial wait time
    - max_delay: Ceiling to prevent runaway delays
    """
    # Exponential part: 1s, 2s, 4s...
    delay = min(max_delay, base_delay * (2 ** (attempt - 1)))
    # Jitter part: add +/- 20% to desynchronize threads
    jitter = delay * 0.2 * (random.random() * 2 - 1)
    final_sleep = max(0.1, delay + jitter)
    
    logger.debug(f"[BACKOFF] Attempt {attempt}: Sleeping for {final_sleep:.2f}s (Jittered from {delay:.2f}s)")
    time.sleep(final_sleep)


_api_limiter = RateLimiter(calls_per_second=API_RATE_LIMIT)


# (MarketDataStore unused HDF5 class removed for performance)


# ===============================
# THREAD-LOCAL FYERS CLIENT
# ===============================
_thread_local = threading.local()


def _get_thread_client(app_id: str, access_token: str) -> object:
    """Returns a thread-local Fyers client instance."""
    if not hasattr(_thread_local, "client"):
        try:
            _thread_local.client = fyersModel.FyersModel(
                client_id=app_id,
                token=access_token,
                log_path=None
            )
        except Exception as e:
            logger.error(f"Failed to init thread-local client: {e}")
            _thread_local.client = None
    return _thread_local.client


# ===============================
# MANIFEST MANAGER
# ===============================
class ManifestManager:
    """
    Manages data synchronization state via a JSON manifest file.

    Reads:   market_data_manifest.json
    Writes:  market_data_manifest.json (timestamps, bootstrap flags, scrub dates)
    """
    def __init__(self, root: str):
        self.path = os.path.join(root, MANIFEST_FILENAME)
        self.data = self._load()

    def _load(self):
        if not os.path.exists(self.path):
            return {}
        try:
            with open(self.path, "r") as f:
                return json.load(f)
        except Exception:
            return {}

    def save(self):
        try:
            with open(self.path, "w") as f:
                json.dump(self.data, f, indent=2)
        except Exception as e:
            logger.error(f"ManifestManager save error: {e}")

    def get_entry(self, symbol: str) -> dict:
        return self.data.get(symbol, {})

    def update_timestamp(self, symbol: str, tf: str, ts: str):
        if symbol not in self.data:
            self.data[symbol] = {}
        self.data[symbol][f"last_ts_{tf}"] = ts

    def mark_bootstrapped(self, symbol: str, tf: str):
        if symbol not in self.data:
            self.data[symbol] = {}
        self.data[symbol][f"bootstrapped_{tf}"] = True

    def update_scrub_date(self, symbol: str, date_str: str):
        if symbol not in self.data:
            self.data[symbol] = {}
        self.data[symbol]["last_scrub_date"] = date_str


# ===============================
# SYMBOL RESOLUTION
# ===============================
def resolve_symbol(excel_symbol: str) -> Tuple[str, str, bool]:
    """
    Resolves Excel symbol to Yahoo and Fyers format.
    Defensively strips known suffixes to prevent duplication.
    """
    if not excel_symbol: return "", "", False
    
    # Defensive Clean
    s = excel_symbol.strip().upper()
    s = s.replace("NSE:", "").replace("-EQ", "").replace("-INDEX", "").replace(".NS", "").replace("^", "")
    
    if s in FYERS_INDEX_MAP:
        target_fyers = FYERS_INDEX_MAP[s]
        # [V3.31] Canonical Yahoo Resolver: Ensure all aliases return same yahoo ticker
        # (Returns the first key in the map that points to this target fyers symbol)
        canonical_yahoo = s
        for k, v in FYERS_INDEX_MAP.items():
            if v == target_fyers:
                canonical_yahoo = k; break
        return f"^{canonical_yahoo}", target_fyers, True
    
    # YFinance and Fyers expect exact NSE symbols (no hyphen or ampersand stripping required).
    # Examples: 'M&MFIN.NS', 'NAM-INDIA.NS', 'NSE:M&MFIN-EQ', 'NSE:NAM-INDIA-EQ'
    yahoo_sym = f"{s}.NS"
    fyers_sym = f"NSE:{s}-EQ"
    return yahoo_sym, fyers_sym, False


# ===============================
# PARQUET I/O
# ===============================
def robust_read_parquet(path: str) -> pd.DataFrame:
    """Reads a parquet file safely, returning empty DataFrame on failure."""
    if not os.path.exists(path): return pd.DataFrame()
    try:
        cols = ['open', 'high', 'low', 'close', 'volume']
        return pd.read_parquet(path, engine='pyarrow', columns=cols)
    except:
        return pd.DataFrame()


def sanitize_new_data(df, last_ts=None, is_daily=False):
    """
    Cleans fetched data: deduplicates, filters trading hours, validates prices.

    Reads:   Raw DataFrame from API
    Returns: Cleaned DataFrame with valid OHLCV rows only
    """
    if df.empty: return df
    if not isinstance(df.index, pd.DatetimeIndex):
        df.index = pd.to_datetime(df.index)

    # 2. Sorting & Deduplication
    df = df[~df.index.duplicated(keep='last')].sort_index()

    # 3. Holiday/Weekend Guard (Universal)
    # Rejects rows falling on weekends/holidays unless they are explicitly marked as EXCEPTION_TRADING_DAYS.
    dates_str = df.index.strftime("%Y-%m-%d")
    is_weekend = df.index.dayofweek >= 5
    is_hardcoded_holiday = dates_str.isin(HARDCODED_HOLIDAYS)
    is_exception = dates_str.isin(EXCEPTION_TRADING_DAYS)
    
    # Core Logic: Keep if it's an exception day OR (NOT a weekend AND NOT a holiday)
    keep_mask = is_exception | (~is_weekend & ~is_hardcoded_holiday)
    df = df[keep_mask]

    # 4. Time Filter (Intraday Only)
    # Only filter time for intraday (5m), NOT daily
    if not is_daily:
        mask = (df.index.time >= dtime(9, 15)) & (df.index.time < dtime(15, 30))
        df = df[mask]

    df = df[~df.index.duplicated(keep='last')]
    df.sort_index(inplace=True)
    if last_ts is not None:
        df = df[df.index > last_ts]
        if df.empty: return df

    condition = (df['high'] > 0) & (df['low'] > 0) & (df['close'] > 0)
    df = df[condition]

    cols = ['open', 'high', 'low', 'close', 'volume']
    for c in cols:
        if c in df.columns: df[c] = pd.to_numeric(df[c], errors='coerce')
    return df.dropna(subset=['close'])


def merge_master(existing: pd.DataFrame, new: pd.DataFrame) -> pd.DataFrame:
    """Merges new data into existing, deduplicating on index."""
    if existing is None or existing.empty: return new.copy()
    if new.empty: return existing.copy()
    combined = pd.concat([existing, new])
    combined = combined[~combined.index.duplicated(keep='last')]
    combined.sort_index(inplace=True)
    return combined


# ===============================
# AUTH HELPER
# ===============================
def load_fyers_credentials(root: str) -> Tuple[str, str]:
    """
    Loads Fyers API credentials from token file.

    Reads:   fyers_tokens.json
    Returns: (app_id, access_token) or ('', '') on failure
    """
    token_path = os.path.join(root, TOKEN_FILENAME)
    if not os.path.exists(token_path): return "", ""
    try:
        with open(token_path, "r") as f: tokens = json.load(f)
        app_id = tokens.get("app_id") or tokens.get("client_id")
        access_token = tokens.get("access_token") or tokens.get("token")
        return str(app_id), str(access_token)
    except: return "", ""


def create_fyers_client(root: Optional[str] = None) -> Optional[object]:
    """Creates a Fyers API client from stored credentials."""
    root = root or get_root_folder()
    token_path = os.path.join(root, TOKEN_FILENAME)
    if fyersModel is None or not os.path.exists(token_path): return None
    try:
        with open(token_path, "r") as f: tokens = json.load(f)
        app_id = tokens.get("app_id") or tokens.get("client_id")
        access_token = tokens.get("access_token") or tokens.get("token")
        if app_id and access_token:
            return fyersModel.FyersModel(client_id=app_id, token=access_token, log_path=None)
    except: pass
    return None


# ===============================
# FYERS HISTORY FETCH
# ===============================
def fetch_fyers_history(client, symbol: str, resolution: str, range_days: int = 1, start_date=None, end_date=None) -> Tuple[pd.DataFrame, bool]:
    """
    Fetches historical candle data from Fyers API in chunks.

    Reads:   Fyers API (chunked date ranges)
    Returns: (DataFrame, reached_target_bool)
    """
    if not client:
        logger.error(f"[FETCH] {symbol}: No active Fyers client.")
        return pd.DataFrame(), False

    CHUNK_SIZE = 365 if resolution == "1D" else 30
    
    # Precise Range Logic
    if start_date and end_date:
        fetch_to = end_date
        fetch_from_limit = start_date
    else:
        fetch_to = datetime.now().date()
        fetch_from_limit = fetch_to - timedelta(days=range_days)

    all_dfs = []
    curr_to = fetch_to
    reached_target = True

    while curr_to >= fetch_from_limit:
        curr_from = curr_to - timedelta(days=CHUNK_SIZE)
        if curr_from < fetch_from_limit: curr_from = fetch_from_limit

        data = {
            "symbol": symbol, "resolution": resolution, "date_format": "1",
            "range_from": curr_from.strftime("%Y-%m-%d"),
            "range_to": curr_to.strftime("%Y-%m-%d"),
            "cont_flag": "1"
        }

        try:
            response = None
            success_chunk = False
            for attempt in range(1, 4):
                _api_limiter.wait() # [V3.26] Respect rate limit even during retries
                try:
                    with ThreadPoolExecutor(max_workers=1) as executor:
                        future = executor.submit(client.history, data=data)
                        response = future.result(timeout=15)
                    if response and response.get("s") == "ok": 
                        success_chunk = True
                        break
                except Exception as ex_call:
                    logger.warning(f"[FETCH] {symbol}: Call Timeout/Fail (Attempt {attempt}): {ex_call}")
                    time.sleep(1)
                    continue

                if response:
                    code = response.get('code')
                    if code == 429:
                        backoff_sleep(attempt, base_delay=3.0, max_delay=30.0)
                    elif code == -99:
                        backoff_sleep(attempt, base_delay=0.5)
                    elif code == -17:
                        logger.error(f"[FETCH] {symbol}: Auth Failed (Token Expired).")
                        return pd.DataFrame(), False
                else:
                    backoff_sleep(attempt, base_delay=0.2)

            if success_chunk and response and response.get("s") == "ok":
                candles = response.get("candles", [])
                if candles:
                    chunk = pd.DataFrame(candles, columns=["epoch", "open", "high", "low", "close", "volume"])
                    chunk["date"] = pd.to_datetime(chunk["epoch"], unit='s').dt.tz_localize('UTC').dt.tz_convert('Asia/Kolkata').dt.tz_localize(None)
                    if resolution == "1D": chunk["date"] = chunk["date"].dt.normalize()
                    chunk.set_index("date", inplace=True)
                    chunk.drop(columns=["epoch"], inplace=True)
                    all_dfs.append(chunk[["open", "high", "low", "close", "volume"]])
                else:
                    # Empty list with 'ok' status is same as no_data
                    logger.debug(f"[FETCH] {symbol} ({resolution}): Empty candles for {curr_from} to {curr_to}")
            elif response and response.get("s") == "no_data":
                logger.debug(f"[FETCH] {symbol} ({resolution}): No data for {curr_from} to {curr_to} (Status: no_data)")
            else:
                reached_target = False
                if response:
                    err_msg = response.get('message', "No Message")
                    err_code = response.get('code', "Ghost Code")
                    if err_code == -300:
                        logger.warning(f"[FETCH] {symbol} ({resolution}): Symbol probably has no data for this range. Skipping.")
                        break 
                    logger.error(f"[FETCH] {symbol} ({resolution}): API Error -> {err_msg} (Code: {err_code})")
                else:
                    logger.error(f"[FETCH] {symbol} ({resolution}): Server No Response (Timeout/Network)")
                break # Exit while loop on any real chunk failure

        except Exception as e:
            logger.error(f"[FETCH] {symbol}: Exception -> {str(e)}")
            reached_target = False
            break

        curr_to = curr_from - timedelta(days=1)

    if not all_dfs:
        return pd.DataFrame(), False

    final_df = pd.concat(all_dfs)
    final_df.sort_index(inplace=True)
    final_df = final_df[~final_df.index.duplicated(keep='last')]
    
    return final_df, reached_target


# ===============================
# HTTP / WEBSOCKET BATCH SNAPSHOT
# ===============================
def _safe_qfloat(v, default: float = 0.0) -> float:
    """Failsafe numeric cast for one live-quote payload field.

    Fyers may return None, "", or a nonnumeric token for a single field.
    A malformed value must degrade to `default` (0.0) so one bad item cannot
    bubble up and abort an entire 50-symbol snapshot batch.
    """
    try:
        if v is None or (isinstance(v, str) and not v.strip()):
            return default
        out = float(v)
        if out != out:  # NaN -> explicit default
            return default
        return out
    except Exception:
        return default


def parse_quote_payload(v: dict, symbol: str) -> dict:
    """
    Pure parser for one FYERS v3 /data/quotes payload item into the BTB shadow
    snapshot_store row schema (identical schema to persist_live_depth).

    PRICES-ONLY L1 POLICY (V4.2 shadow layer):
      FYERS /data/quotes does NOT provide bid/ask size (bid_size/ask_size) or
      the aggregate quantities (last_traded_qty/tot_buy_qty/tot_sell_qty).
      Size-based L1 pressure (imbalance / slope / pressure score) therefore
      requires websocket (SymbolUpdate) or the per-symbol /data/depth API.
      All size fields are written as explicit 0.0 here so downstream BTB
      features stay NEUTRAL instead of silently misreading absent fields.

    All numeric fields go through _safe_qfloat so a single malformed item
    never kills the batch (get_batch_snapshot applies one item per symbol).
    """
    return {
        "ltp": _safe_qfloat(v.get("lp")),
        "open": _safe_qfloat(v.get("open_price")),
        "high": _safe_qfloat(v.get("high_price")),
        "low": _safe_qfloat(v.get("low_price")),
        "volume": _safe_qfloat(v.get("volume")),
        "bid_price": _safe_qfloat(v.get("bid")),
        "ask_price": _safe_qfloat(v.get("ask")),
        "bid_size": 0.0,
        "ask_size": 0.0,
        "last_traded_qty": 0.0,
        "tot_buy_qty": 0.0,
        "tot_sell_qty": 0.0,
        "avg_trade_price": _safe_qfloat(v.get("atp")),
        "snapshot_ts": datetime.now().isoformat(timespec="seconds"),
    }


_none_snapshot_row = {"ltp": 0.0, "open": 0.0, "high": 0.0, "low": 0.0, "volume": 0.0,
                      "bid_price": 0.0, "ask_price": 0.0, "bid_size": 0.0,
                      "ask_size": 0.0, "last_traded_qty": 0.0,
                      "tot_buy_qty": 0.0, "tot_sell_qty": 0.0,
                      "avg_trade_price": 0.0, "snapshot_ts": ""}


def get_batch_snapshot(client, fyers_symbols):
    """
    Fetches live quotes using local WebSocket Cache (Zero Latency). 
    Falls back to FYERS HTTP batch API if WebSocket cache is stale or missing.
    """
    snapshot_store = {}
    
    if not is_market_open_now():
        logger.info("[SNAPSHOT] Blocked (Market Closed)")
        return snapshot_store

    if not client or not fyers_symbols:
        return snapshot_store

    requested = len(fyers_symbols)
    logger.info(f"[SNAPSHOT] Request | count={requested}")

    # Pure HTTP Fetch (Zero WS Cache)
    missing_symbols = fyers_symbols
    logger.info(f"[SNAPSHOT][HTTP] Fetching {len(missing_symbols)} symbols directly from exchange.")
        
    logger.warning(f"[SNAPSHOT][HTTP] Fallback required for {len(missing_symbols)} symbols.")

    def chunks(lst, size=50):
        for i in range(0, len(lst), size):
            yield lst[i:i + size]

    for batch in chunks(missing_symbols, 50):
        symbols_str = ",".join(batch)
        MAX_RETRY = 2
        resp = None

        for attempt in range(1, MAX_RETRY + 1):
            try:
                _api_limiter.wait() # CRITICAL: Stay under 3 calls/sec burst limit
                resp = client.quotes({"symbols": symbols_str})
                if resp and resp.get("s") == "ok":
                    data_list = resp.get("d", [])
                    if isinstance(data_list, list):
                        for item in data_list:
                            v = item.get("v", {})
                            s_name = v.get("symbol") or item.get("n")
                            if s_name:
                                snapshot_store[s_name] = parse_quote_payload(v, s_name)
                    break
                else:
                    err_msg = resp.get("message", "Unknown") if resp else "No Response"
                    err_code = resp.get("code", "N/A") if resp else "N/A"
                    logger.warning(f"[SNAPSHOT][HTTP] {err_msg} (Code: {err_code}) batch_size={len(batch)} attempt={attempt}/{MAX_RETRY}")
            except Exception as e:
                logger.warning(f"[SNAPSHOT][HTTP] exception batch_size={len(batch)} attempt={attempt}/{MAX_RETRY} error={e}")
            
            if attempt < MAX_RETRY:
                backoff_sleep(attempt, base_delay=0.5)

    # Validate final payload
    returned_syms = set(snapshot_store.keys())    
    missing = set(fyers_symbols) - returned_syms
    for sym in missing:
        snapshot_store[sym] = dict(_none_snapshot_row)

    logger.info(f"[SNAPSHOT] Complete | requested={requested} valid={len(returned_syms)} missing={len(missing)}")

    if snapshot_store:
        persist_live_depth(snapshot_store)

    return snapshot_store


# ===============================
# L1 LIVE-DEPTH PARQUET PERSISTENCE (BTB shadow layer; VBA does NOT consume)
# ===============================
def persist_live_depth(snapshot_store: Dict[str, Dict], root: Optional[str] = None) -> None:
    """
    Appends the per-cycle L1 snapshot to data/live_depth/YYYYMMDD.parquet.
    De-duplicates on (symbol, snapshot_ts) so a retried cycle cannot double-write.

    Reads:   get_batch_snapshot() returned snapshot_store
    Writes:  data/live_depth/YYYYMMDD.parquet (append; retry-safe)
    """
    try:
        root = root or get_root_folder()
        depth_dir = os.path.join(root, "data", "live_depth")
        ensure_dir(depth_dir)

        today = datetime.now().strftime("%Y%m%d")
        path = os.path.join(depth_dir, f"{today}.parquet")

        rows = []
        for sym, s in snapshot_store.items():
            if not sym:
                continue
            rows.append({
                "symbol": sym,
                "ts": s.get("snapshot_ts", ""),
                "ltp": float(s.get("ltp", 0.0)),
                "bid_price": float(s.get("bid_price", 0.0)),
                "ask_price": float(s.get("ask_price", 0.0)),
                "bid_size": float(s.get("bid_size", 0.0)),
                "ask_size": float(s.get("ask_size", 0.0)),
                "last_traded_qty": float(s.get("last_traded_qty", 0.0)),
                "tot_buy_qty": float(s.get("tot_buy_qty", 0.0)),
                "tot_sell_qty": float(s.get("tot_sell_qty", 0.0)),
                "avg_trade_price": float(s.get("avg_trade_price", 0.0)),
            })
        if not rows:
            return

        new_df = pd.DataFrame(rows)

        if os.path.exists(path):
            try:
                existing = pd.read_parquet(path)
                combined = pd.concat([existing, new_df], ignore_index=True)
            except Exception:
                combined = new_df
            combined = combined.drop_duplicates(subset=["symbol", "ts"], keep="last")
        else:
            combined = new_df

        combined.to_parquet(path, index=False)
        logger.info(f"[LIVE_DEPTH] Persisted {len(new_df)} rows -> {path} (total={len(combined)})")
    except Exception as e:
        logger.warning(f"[LIVE_DEPTH] Persist failed (non-fatal): {e}")



# ===============================
# BTB V4.2 TELEMETRY (replay-only diagnostics; NEVER asserted)
# ===============================
# One row per (cycle, symbol). Captures raw signal flags (DF + AC), the frozen
# V4.2 gate inputs VBA reads from Sheet1, and the Python BTB shadow veto state
# at compute time. gate_src=REPLAY marks every row as replay-derived: the v4/btb
# gate results and skip_reason are reconstructed by scripts/btb_v4_monitor.py,
# never treated as ground truth. No VBA, no orders, no exits, no new API calls.
BTB_TELEMETRY_COLUMNS = [
    "ts", "symbol",
    "df_breakout", "ac_cross",                       # RAW_SIGNAL flags (DF / AC)
    "ltp", "regime_5m", "sec_regime", "grade",       # V4.2 gate inputs
    "rsi_5m", "rvol_5m", "atr_5m", "ema30_5m",
    "swing_high", "swing_low",
    "sector", "size",
    "btb_ext_side", "btb_ext_decision", "btb_ext_score", "btb_data_quality",
    "gate_src",
    "v4_gate_result_replay", "btb_gate_result_replay", "skip_reason_replay", "review_mismatch",
]


def persist_btb_telemetry(telemetry_rows: List[Dict], root: Optional[str] = None) -> None:
    """
    Appends per-cycle BTB V4.2 telemetry rows to data/btb_telemetry/YYYYMMDD.parquet.
    De-duplicates on (symbol, ts) so a retried cycle cannot double-write.

    Reads:   AI_BTB_Engine compute loop telemetry rows (gate_src=REPLAY)
    Writes:  data/btb_telemetry/YYYYMMDD.parquet (append; retry-safe)
    Non-fatal: telemetry must never break the engine.
    """
    try:
        if not telemetry_rows:
            return
        root = root or get_root_folder()
        tel_dir = os.path.join(root, "data", "btb_telemetry")
        ensure_dir(tel_dir)

        today = datetime.now().strftime("%Y%m%d")
        path = os.path.join(tel_dir, f"{today}.parquet")

        new_df = pd.DataFrame(telemetry_rows)
        for col in BTB_TELEMETRY_COLUMNS:
            if col not in new_df.columns:
                new_df[col] = "" if col in (
                    "gate_src", "v4_gate_result_replay", "btb_gate_result_replay",
                    "skip_reason_replay", "review_mismatch",
                ) else 0.0

        if os.path.exists(path):
            try:
                existing = pd.read_parquet(path)
                combined = pd.concat([existing, new_df], ignore_index=True)
            except Exception:
                combined = new_df
            combined = combined.drop_duplicates(subset=["symbol", "ts"], keep="last")
        else:
            combined = new_df

        combined.to_parquet(path, index=False)
        logger.info(f"[BTB_TELEMETRY] Persisted {len(new_df)} rows -> {path} (total={len(combined)})")
    except Exception as e:
        logger.warning(f"[BTB_TELEMETRY] Persist failed (non-fatal): {e}")


# ===============================
# SNAPSHOT BATCH (for sync orchestrator)
# ===============================
def fetch_snapshot_batch(yahoo_symbols, root):
    """
    Fetches live quotes for a list of Yahoo symbols via Fyers HTTP API.

    Reads:   Yahoo symbol list, Fyers credentials
    Returns: dict[yahoo_symbol] -> {ltp, open, high, low, volume}
    Refers:  resolve_symbol for symbol translation, get_batch_snapshot for HTTP fetch
    """
    app_id, access_token = load_fyers_credentials(root)
    if not app_id or not access_token:
        logger.error("[SNAPSHOT] credentials_missing")
        return {}

    client = _get_thread_client(app_id, access_token)
    if not client:
        logger.error("[SNAPSHOT] fyers_client_init_failed")
        return {}

    equity_map = {}
    index_map = {}

    # [V3.28] Robust Index Identification
    # Ensure we check both the prefix (^) and the central index map
    index_keys = set(FYERS_INDEX_MAP.keys())
    
    for y in yahoo_symbols:
        y_canonical, fy_sym, _ = resolve_symbol(y)
        if not fy_sym: continue
        
        # Clean symbol for map lookup (e.g., ^NSEI -> NSEI)
        clean_y = y_canonical.replace("^", "").replace(".NS", "")
        if y_canonical.startswith("^") or clean_y in index_keys:
            index_map[fy_sym] = y_canonical
        else:
            equity_map[fy_sym] = y_canonical

    logger.info(f"[SNAPSHOT] symbols | equity={len(equity_map)} index={len(index_map)} total={len(equity_map)+len(index_map)}")

    snapshot_map = {}

    
    def chunks(lst, size=50):
        for i in range(0, len(lst), size):
            yield lst[i:i + size]

    # Equities
    for batch in chunks(list(equity_map.keys()), 50):
        logger.info(f"[SNAPSHOT][EQUITY] batch | size={len(batch)}")
        raw = get_batch_snapshot(client, batch)
        for fy_sym in batch:
            y_sym = equity_map.get(fy_sym)
            if not y_sym: continue
            snapshot_map[y_sym] = raw.get(fy_sym, {"ltp": 0.0, "open": 0.0, "high": 0.0, "low": 0.0, "volume": 0.0})

    # Indices
    for batch in chunks(list(index_map.keys()), 50):
        logger.info(f"[SNAPSHOT][INDEX] batch | size={len(batch)}")
        raw = get_batch_snapshot(client, batch)
        for fy_sym in batch:
            y_sym = index_map.get(fy_sym)
            if not y_sym: continue
            snapshot_map[y_sym] = raw.get(fy_sym, {"ltp": 0.0, "open": 0.0, "high": 0.0, "low": 0.0, "volume": 0.0})
        time.sleep(0.1)

    logger.info(f"[SNAPSHOT] complete | returned={len(snapshot_map)}/{len(yahoo_symbols)}")
    return snapshot_map


# ===============================
# PHASE 1: LOAD FROM DISK
# ===============================
_LOADED_DATA_CACHE: Dict[str, Tuple[float, Dict]] = {}
_LOADED_DATA_CACHE_MAX_AGE = 600

def _load_from_cache_or_disk(yahoo_symbol: str, root: str) -> Dict:
    """Load symbol data with mtime-based caching to avoid redundant parquet reads."""
    global _LOADED_DATA_CACHE
    p5 = parquet_path_5m(yahoo_symbol, root)
    p1d = parquet_path_daily(yahoo_symbol, root)
    cache_key = yahoo_symbol
    mtime5 = os.path.getmtime(p5) if os.path.exists(p5) else 0.0
    mtime1d = os.path.getmtime(p1d) if os.path.exists(p1d) else 0.0
    combined_mtime = mtime5 + mtime1d
    cached_entry = _LOADED_DATA_CACHE.get(cache_key)
    if cached_entry is not None:
        cached_mtime, cached_data = cached_entry
        if abs(cached_mtime - combined_mtime) < 0.01:
            return dict(cached_data)
    result = load_symbol_from_disk(yahoo_symbol, root)
    _LOADED_DATA_CACHE[cache_key] = (combined_mtime, result)
    if len(_LOADED_DATA_CACHE) > 600:
        oldest_keys = list(_LOADED_DATA_CACHE.keys())[:100]
        for k in oldest_keys:
            _LOADED_DATA_CACHE.pop(k, None)
    return result


def load_symbol_from_disk(yahoo_symbol: str, root: Optional[str] = None) -> Dict:
    """
    Loads a symbol's 5m and 1d parquet data from disk.
    """
    root = root or get_root_folder()
    p5 = parquet_path_5m(yahoo_symbol, root)
    p1d = parquet_path_daily(yahoo_symbol, root)
    d5 = pd.DataFrame(columns=["open", "high", "low", "close", "volume"])
    is_cold = True
    if os.path.exists(p5):
        try:
            d5 = robust_read_parquet(p5)
            if not d5.empty:
                min_price = d5["close"].min()
                if min_price <= 0.01:
                    d5 = pd.DataFrame(columns=["open", "high", "low", "close", "volume"])
                    try: os.remove(p5)
                    except: pass
                    is_cold = True
                else:
                    is_cold = False
                    if not isinstance(d5.index, pd.DatetimeIndex):
                        d5.index = pd.to_datetime(d5.index)
                    if not d5.index.is_monotonic_increasing:
                        d5 = pd.DataFrame(columns=["open", "high", "low", "close", "volume"])
                        try: os.remove(p5)
                        except: pass
                        is_cold = True
        except Exception:
            try: os.remove(p5)
            except: pass
            d5 = pd.DataFrame(columns=["open", "high", "low", "close", "volume"])
            is_cold = True
    d1 = robust_read_parquet(p1d)
    return {
        "symbol": yahoo_symbol,
        "5m": d5,
        "1d": d1,
        "is_cold": is_cold
    }


def _missing_recent_1d_days(df_1d: pd.DataFrame, prev_trade_date, lookback: int = 12) -> List:
    """
    [FIX 2026-08-03] Internal-gap audit for daily data (parquet truth, not manifest).

    Detects trading days that should exist in the daily frame but don't. The manifest-based
    gap_days_1d check goes stale because harden_1d_with_snapshot advances last_ts_1d to
    "today" every run, so internal holes (e.g. 2026-07-30/31 lost in a backup restore)
    become invisible and are never re-fetched. Only completed sessions up to
    prev_trade_date are checked (today's partial live candle is never required).

    Returns: sorted list of missing datetime.date objects (empty if complete/empty).
    """
    if df_1d is None or df_1d.empty:
        return []
    present = set(df_1d.index.date)
    floor = df_1d.index.min().date()
    expected = []
    d = prev_trade_date.date() if hasattr(prev_trade_date, "date") else prev_trade_date
    while len(expected) < lookback:
        if d < floor:
            break
        expected.append(d)
        d = get_last_trading_date(d - timedelta(days=1))
        if hasattr(d, "date"):
            d = d.date()
    return sorted(e for e in expected if e not in present)


# ===============================
# PHASE 2: FETCH & MERGE
# ===============================
def fetch_and_merge_data(payload: Dict, quiet: bool = False) -> Dict:
    """
    Unified Data Fetcher & Merger (5M and 1D).
    Executes identical date range logic for 5m and 1d candles based on task_meta:
    - BOOTSTRAP: Full historical backfill.
    - POST_MARKET_SEAL / CATCHUP_SEAL: Authoritative sync for [Gap + 2 previous trading days].
    - SETTLEMENT_REFRESH: Full current-day refresh (15:30 - 17:00).
    - LIVE_BAR_SYNC: Live closed intraday candles (09:15 - 15:30).
    """
    loaded = payload["data"]
    snapshot = payload.get("snapshot")
    manifest = payload.get("manifest_entry", {})
    task_meta = payload.get("task_meta", {})

    symbol = loaded["symbol"]
    df_5m = loaded["5m"]
    df_1d = loaded["1d"]

    app_id = payload["app_id"]
    token = payload["token"]

    _, fyers_symbol, _ = resolve_symbol(symbol)

    result = loaded.copy()
    result["is_updated"] = loaded.get("is_updated", False)
    result["reached_target"] = True

    client = _get_thread_client(app_id, token)
    if not client:
        logger.error(f"[{symbol}] CLIENT_MISSING")
        return result

    now = datetime.now()
    EOD_BAR_THRESHOLD = dtime(15, 25)

    task_reason = task_meta.get("reason")
    start_date = task_meta.get("start_date")
    end_date = task_meta.get("end_date")
    target_seal_date = task_meta.get("target_seal_date")

    # Standalone fallback if task_meta was not provided
    if not task_reason:
        is_boot_5m = manifest.get("bootstrapped_5m", False)
        is_boot_1d = manifest.get("bootstrapped_1d", False)
        if not is_boot_5m or not is_boot_1d or df_5m.empty or df_1d.empty:
            task_reason = "BOOTSTRAP"
        elif is_market_open_now(now):
            task_reason = "LIVE_BAR_SYNC"
        elif not is_market_holiday_today(now) and now.weekday() < 5 and (now.time() >= dtime(15, 30) and now.time() < dtime(17, 0)):
            task_reason = "SETTLEMENT_REFRESH"
            start_date = now.date()
            end_date = now.date()
        else:
            target_seal_date = get_last_completed_trading_day(now)
            last_scrub = manifest.get("last_scrub_date", "")
            if last_scrub >= target_seal_date.strftime("%Y-%m-%d"):
                return result
            is_post = (not is_market_holiday_today(now) and now.weekday() < 5 and now.time() >= dtime(17, 0))
            task_reason = "POST_MARKET_SEAL" if is_post else "CATCHUP_SEAL"
            start_date, end_date = resolve_seal_sync_range(last_scrub, target_seal_date)

    # ---------------------------------------------------------
    # 1. BOOTSTRAP MODE
    # ---------------------------------------------------------
    if task_reason == "BOOTSTRAP":
        h5, reached_5m = fetch_fyers_history(client, fyers_symbol, "5", range_days=BOOTSTRAP_5M_DAYS)
        if not h5.empty:
            clean_5m = sanitize_new_data(h5, last_ts=None, is_daily=False)
            if not clean_5m.empty:
                df_5m = merge_master(df_5m, clean_5m)
                result["5m"] = df_5m
                result["new_last_ts_5m"] = df_5m.index[-1].strftime("%Y-%m-%d %H:%M:%S")
                result["mark_5m_boot"] = True
                result["is_updated"] = True

        backfill_1d_days = 30 if is_market_open_now(now) else BOOTSTRAP_1D_DAYS
        h1, reached_1d = fetch_fyers_history(client, fyers_symbol, "1D", range_days=backfill_1d_days)
        if not h1.empty:
            clean_1d = sanitize_new_data(h1, last_ts=None, is_daily=True)
            if not clean_1d.empty:
                df_1d = merge_master(df_1d, clean_1d)
                result["1d"] = df_1d
                result["new_last_ts_1d"] = df_1d.index[-1].strftime("%Y-%m-%d %H:%M:%S")
                if reached_1d and not is_market_open_now(now):
                    result["mark_1d_boot"] = True
                result["is_updated"] = True

        if not reached_5m or not reached_1d:
            result["reached_target"] = False
        return result

    # ---------------------------------------------------------
    # 2. SEAL / CATCHUP MODE (UNIFIED 5M & 1D SYNC)
    # ---------------------------------------------------------
    elif task_reason in ["POST_MARKET_SEAL", "CATCHUP_SEAL"]:
        clean_5m = pd.DataFrame()
        clean_1d = pd.DataFrame()
        # Identical range fetch for 5M
        h5, reached_5m = fetch_fyers_history(client, fyers_symbol, "5", start_date=start_date, end_date=end_date)
        if not h5.empty:
            clean_5m = sanitize_new_data(h5, last_ts=None, is_daily=False)
            if not clean_5m.empty:
                df_5m = merge_master(df_5m, clean_5m)
                result["5m"] = df_5m
                result["new_last_ts_5m"] = df_5m.index[-1].strftime("%Y-%m-%d %H:%M:%S")
                result["is_updated"] = True

        # Identical range fetch for 1D
        h1, reached_1d = fetch_fyers_history(client, fyers_symbol, "1D", start_date=start_date, end_date=end_date)
        if not h1.empty:
            clean_1d = sanitize_new_data(h1, last_ts=None, is_daily=True)
            if not clean_1d.empty:
                df_1d = merge_master(df_1d, clean_1d)
                result["1d"] = df_1d
                result["new_last_ts_1d"] = df_1d.index[-1].strftime("%Y-%m-%d %H:%M:%S")
                result["is_updated"] = True

        # Completeness validation for target_seal_date
        has_5m_target = False
        if not df_5m.empty:
            _last_5m = df_5m.index[-1]
            if hasattr(_last_5m, 'date') and hasattr(_last_5m, 'time'):
                has_5m_target = (_last_5m.date() >= target_seal_date and _last_5m.time() >= EOD_BAR_THRESHOLD)

        has_1d_target = False
        if not df_1d.empty:
            _last_1d = df_1d.index[-1]
            if hasattr(_last_1d, 'date'):
                has_1d_target = (_last_1d.date() >= target_seal_date)

        if reached_5m and reached_1d and has_5m_target and has_1d_target:
            result["is_sealed"] = True
            result["seal_date"] = target_seal_date.strftime("%Y-%m-%d")
            logger.info(f"[{symbol}] {task_reason}_COMPLETE | downloaded: 5m={len(clean_5m)}, 1d={len(clean_1d)}")
        else:
            result["reached_target"] = False
            logger.warning(f"[{symbol}] {task_reason}_PARTIAL | 5m_reach={reached_5m} 1d_reach={reached_1d} 5m_ok={has_5m_target} 1d_ok={has_1d_target} | will retry")

        return result

    # ---------------------------------------------------------
    # 3. SETTLEMENT REFRESH MODE (15:30 - 17:00 on trading days)
    # ---------------------------------------------------------
    elif task_reason == "SETTLEMENT_REFRESH":
        ref_date = now.date()
        h5, reached_5m = fetch_fyers_history(client, fyers_symbol, "5", start_date=ref_date, end_date=ref_date)
        if not h5.empty:
            clean_5m = sanitize_new_data(h5, last_ts=None, is_daily=False)
            if not clean_5m.empty:
                df_5m = merge_master(df_5m, clean_5m)
                result["5m"] = df_5m
                result["new_last_ts_5m"] = df_5m.index[-1].strftime("%Y-%m-%d %H:%M:%S")
                result["is_updated"] = True

        h1, reached_1d = fetch_fyers_history(client, fyers_symbol, "1D", start_date=ref_date, end_date=ref_date)
        if not h1.empty:
            clean_1d = sanitize_new_data(h1, last_ts=None, is_daily=True)
            if not clean_1d.empty:
                df_1d = merge_master(df_1d, clean_1d)
                result["1d"] = df_1d
                result["new_last_ts_1d"] = df_1d.index[-1].strftime("%Y-%m-%d %H:%M:%S")
                result["is_updated"] = True

        if not reached_5m or not reached_1d:
            result["reached_target"] = False
        logger.info(f"[{symbol}] SETTLEMENT_REFRESH | 5m_fetched={len(h5)} | 1d_fetched={len(h1)}")
        return result

    # ---------------------------------------------------------
    # 4. LIVE MARKET SYNC (09:15 - 15:30 on trading days)
    # ---------------------------------------------------------
    elif task_reason == "LIVE_BAR_SYNC":
        h5, reached_5m = fetch_fyers_history(client, fyers_symbol, "5", start_date=now.date(), end_date=now.date())
        if not h5.empty:
            clean_5m = sanitize_new_data(h5, last_ts=None, is_daily=False)
            if not clean_5m.empty:
                df_5m = merge_master(df_5m, clean_5m)
                result["5m"] = df_5m
                result["new_last_ts_5m"] = df_5m.index[-1].strftime("%Y-%m-%d %H:%M:%S")
                result["is_updated"] = True

        if snapshot:
            d = pd.Timestamp(now.date())
            t_close = float(snapshot.get("ltp", 0.0))
            if t_close > 0:
                df_1d.loc[d] = {
                    "open": float(snapshot.get("open", 0.0)),
                    "high": float(snapshot.get("high", 0.0)),
                    "low": float(snapshot.get("low", 0.0)),
                    "close": t_close,
                    "volume": float(snapshot.get("volume", 0.0))
                }
                result["1d"] = df_1d
                result["is_updated"] = True

        logger.info(f"[{symbol}] LIVE_SYNC | 5m_fetched={len(h5)}")
        return result

    return result


def save_symbol_to_disk(data_pack, manifest_mgr, root=None):
    """
    Writes 5m and 1d parquets to disk, updates manifest timestamps and scrub dates.
    Only stamps last_scrub_date when data is strictly complete for the target seal date.
    """
    root = root or get_root_folder()
    symbol = data_pack["symbol"]

    if "5m" in data_pack and isinstance(data_pack["5m"], pd.DataFrame) and not data_pack["5m"].empty:
        p5 = parquet_path_5m(symbol, root)
        data_pack["5m"].to_parquet(p5, compression="snappy")
        last_dt_5m = data_pack["5m"].index[-1]
        manifest_mgr.update_timestamp(symbol, "5m", last_dt_5m.strftime("%Y-%m-%d %H:%M:%S"))

    if "1d" in data_pack and isinstance(data_pack["1d"], pd.DataFrame) and not data_pack["1d"].empty:
        p1d = parquet_path_daily(symbol, root)
        data_pack["1d"].to_parquet(p1d, compression="snappy")
        last_dt_1d = data_pack["1d"].index[-1]
        ts_1d = data_pack.get("new_last_ts_1d") or last_dt_1d.strftime("%Y-%m-%d %H:%M:%S")
        manifest_mgr.update_timestamp(symbol, "1d", ts_1d)

    if data_pack.get("mark_5m_boot"):
        manifest_mgr.mark_bootstrapped(symbol, "5m")
    if data_pack.get("mark_1d_boot"):
        manifest_mgr.mark_bootstrapped(symbol, "1d")

    # Transactional Seal Gate: stamp last_scrub_date only when both 5M and 1D are confirmed complete
    if data_pack.get("is_sealed") and data_pack.get("seal_date"):
        _seal_dt = datetime.strptime(data_pack["seal_date"], "%Y-%m-%d").date()
        _eod_threshold = dtime(15, 25)

        _5m_ok = False
        if "5m" in data_pack and isinstance(data_pack["5m"], pd.DataFrame) and not data_pack["5m"].empty:
            _last_5m = data_pack["5m"].index[-1]
            if hasattr(_last_5m, 'date') and hasattr(_last_5m, 'time'):
                _5m_ok = (_last_5m.date() >= _seal_dt and _last_5m.time() >= _eod_threshold)

        _1d_ok = False
        if "1d" in data_pack and isinstance(data_pack["1d"], pd.DataFrame) and not data_pack["1d"].empty:
            _last_1d = data_pack["1d"].index[-1]
            if hasattr(_last_1d, 'date'):
                _1d_ok = (_last_1d.date() >= _seal_dt)

        if _5m_ok and _1d_ok:
            manifest_mgr.update_scrub_date(symbol, data_pack["seal_date"])


def orchestrated_sync(
    symbols: List[str],
    manifest_mgr: ManifestManager,
    root: Optional[str] = None,
    snapshot_map: Optional[Dict] = None
) -> Dict[str, Dict]:
    """
    Main 4-phase sync pipeline: Load (Parquet) → Snapshot → Fetch/Merge → Save (Parquet).
    Strictly follows manifest last_scrub_date as single version of truth.
    """
    root = root or get_root_folder()

    app_id, access_token = load_fyers_credentials(root)
    if not app_id or not access_token:
        logger.error("Credentials missing. Aborting sync.")
        return {}

    # Phase 1: Mass load from Parquets
    yahoo_to_orig = {}
    for s in symbols:
        if not s: continue
        y, _, _ = resolve_symbol(s)
        yahoo_to_orig[y] = s

    yahoo_list = sorted(list(yahoo_to_orig.keys()))
    logger.info(f"Phase 1: Mass Load Parquet ({len(yahoo_list)} symbols)...")

    loaded_data_map = {}
    with ThreadPoolExecutor(max_workers=32) as ex:
        futures = {ex.submit(_load_from_cache_or_disk, y, root): y for y in yahoo_list}
        for fut in as_completed(futures):
            res = fut.result()
            loaded_data_map[res["symbol"]] = res

    # Phase 2: Filter/prep
    final_data_map = {}
    tasks_to_submit = []
    task_metadata = {}

    now = datetime.now()
    is_holiday = is_market_holiday_today(now)
    is_weekend = now.weekday() >= 5
    is_market_open = is_market_open_now(now)
    is_regular_trading_day = not is_holiday and not is_weekend
    is_settlement_window = is_regular_trading_day and (now.time() >= dtime(15, 30) and now.time() < dtime(17, 0))
    is_post_market_seal = is_regular_trading_day and (now.time() >= dtime(17, 0))
    is_outside_market_hours = not is_market_open and not is_settlement_window

    target_seal_date = get_last_completed_trading_day(now)
    target_seal_date_str = target_seal_date.strftime("%Y-%m-%d")

    logger.info(f"[SYNC-CTX] Now={now.strftime('%H:%M:%S')} | Target_Seal_Day={target_seal_date_str} | Open={is_market_open} | Settlement={is_settlement_window} | Weekend={is_weekend} | Holiday={is_holiday}")

    for y in yahoo_list:
        orig = yahoo_to_orig[y]
        data_pack = loaded_data_map.get(y)
        if not data_pack: continue

        entry = manifest_mgr.get_entry(y)

        is_boot_5m = entry.get("bootstrapped_5m", False)
        is_boot_1d = entry.get("bootstrapped_1d", False)
        is_missing_history = (not is_boot_5m or not is_boot_1d or data_pack["1d"].empty or data_pack["5m"].empty
                              or len(data_pack["5m"]) < 100 or len(data_pack["1d"]) < 50)

        # 1. Integrity Check: Deep backfill if never bootstrapped or shallow
        if is_missing_history:
            logger.info(f"[SYNC-TASK] {y} -> MANDATORY_BACKFILL (Integrity Breach)")
            task_metadata[y] = {"reason": "BOOTSTRAP"}
            tasks_to_submit.append(y)
            continue

        # 2. Outside Market Hours: Rely strictly on manifest last_scrub_date
        if is_outside_market_hours:
            last_scrub = entry.get("last_scrub_date", "")
            if last_scrub >= target_seal_date_str:
                # Already sealed for the target completed session: zero download
                final_data_map[orig] = data_pack
            else:
                # Unsealed gap detected: sync [last_scrub - 1 trading day, target_seal_date]
                start_d, end_d = resolve_seal_sync_range(last_scrub, target_seal_date)
                seal_reason = "POST_MARKET_SEAL" if is_post_market_seal else "CATCHUP_SEAL"
                task_metadata[y] = {
                    "reason": seal_reason,
                    "start_date": start_d,
                    "end_date": end_d,
                    "target_seal_date": target_seal_date
                }
                tasks_to_submit.append(y)
                logger.info(f"[SYNC-TASK] {y} -> {seal_reason} ({start_d} to {end_d})")

        # 3. Settlement Window (15:30 - 17:00 on regular trading days)
        elif is_settlement_window:
            task_metadata[y] = {
                "reason": "SETTLEMENT_REFRESH",
                "start_date": now.date(),
                "end_date": now.date()
            }
            tasks_to_submit.append(y)

        # 4. Live Market (09:15 - 15:30 on regular trading days)
        elif is_market_open:
            last_ts_5m = entry.get("last_ts_5m", "")
            last_dt = datetime.strptime(last_ts_5m, "%Y-%m-%d %H:%M:%S") if last_ts_5m else None
            last_closed_5m = get_last_closed_5m(now)
            if last_dt and last_dt >= last_closed_5m:
                final_data_map[orig] = data_pack
            else:
                task_metadata[y] = {
                    "reason": "LIVE_BAR_SYNC",
                    "start_date": now.date(),
                    "end_date": now.date()
                }
                tasks_to_submit.append(y)

    # Post-Seal Idle Lock: If outside market hours and all symbols are already sealed
    if is_outside_market_hours and len(tasks_to_submit) == 0:
        logger.info(f"Sync Phase: IDLE (Manifest is fully sealed through {target_seal_date_str}). Zero downloads until market opens (09:15).")
        for s in symbols:
            y, _, _ = resolve_symbol(s)
            data = loaded_data_map.get(y)
            if data: final_data_map[s] = data
        return final_data_map

    # Phase 3: Fetch + merge
    if tasks_to_submit:
        logger.info(f"Phase 3: Deep Syncing {len(tasks_to_submit)} symbols via authoritative HTTP REST...")
        from engine.config import SYNC_WORKERS
        with ThreadPoolExecutor(max_workers=SYNC_WORKERS) as ex:
            futures = {}
            for y in tasks_to_submit:
                y_sym, _, _ = resolve_symbol(y)
                payload = {
                    "data": loaded_data_map[y],
                    "snapshot": snapshot_map.get(y_sym) if snapshot_map else None,
                    "app_id": app_id,
                    "token": access_token,
                    "manifest_entry": manifest_mgr.get_entry(y),
                    "task_meta": task_metadata.get(y, {}),
                }
                futures[ex.submit(fetch_and_merge_data, payload, quiet=True)] = y

            total_tasks = len(futures)
            completed_count = 0
            for fut in as_completed(futures):
                completed_count += 1
                y = futures[fut]
                orig = yahoo_to_orig[y]

                if completed_count % 20 == 0 or completed_count == total_tasks:
                    pct = (completed_count / total_tasks) * 100
                    logger.info(f"[SYNC-PROGRESS] {completed_count}/{total_tasks} symbols completed ({pct:.1f}%)")

                try:
                    res = fut.result()
                    final_data_map[orig] = res
                except Exception as e:
                    logger.error(f"Sync Err {y}: {e}")
                    final_data_map[orig] = loaded_data_map[y]

        # Phase 3.5: Sequential Retry for ACTUAL failures
        retry_bin = []
        for y in tasks_to_submit:
            orig = yahoo_to_orig[y]
            dp = final_data_map.get(orig)
            if not dp or not dp.get("reached_target", True):
                retry_bin.append(y)

        if retry_bin:
            logger.warning(f"Phase 3.5: Sequential Retry for {len(retry_bin)} failed/stale symbols...")
            for y in retry_bin:
                orig = yahoo_to_orig[y]
                try:
                    _api_limiter.wait()
                    y_sym, _, _ = resolve_symbol(y)
                    payload = {
                        "data": final_data_map[orig],
                        "snapshot": snapshot_map.get(y_sym) if snapshot_map else None,
                        "app_id": app_id,
                        "token": access_token,
                        "manifest_entry": manifest_mgr.get_entry(y),
                        "task_meta": task_metadata.get(y, {}),
                    }
                    res = fetch_and_merge_data(payload)
                    final_data_map[orig] = res
                except Exception as e:
                    logger.error(f"Retry Final Failure for {y}: {e}")

        incomplete_seal = sorted(
            o for o in set(yahoo_to_orig.values())
            if not final_data_map.get(o) or not final_data_map[o].get("reached_target", True)
        )
        if incomplete_seal:
            logger.warning(f"Seal incomplete for {len(incomplete_seal)} symbols: {incomplete_seal} (will retry next run)")

    # Phase 4: Write to Store
    write_queue = []
    for _, data in final_data_map.items():
        needs_save = (
            data.get("is_updated") or
            data.get("is_cold") or
            data.get("is_migration") or
            data.get("mark_1d_boot") or
            data.get("mark_5m_boot") or
            data.get("is_sealed") or
            not data.get("reached_target", True)
        )
        if needs_save:
            write_queue.append(data)

    if write_queue:
        logger.info(f"Phase 4: Hardening {len(write_queue)} symbols to Parquets...")
        with ThreadPoolExecutor(max_workers=32) as ex:
            futures = [ex.submit(save_symbol_to_disk, d, manifest_mgr, root) for d in write_queue]
            for fut in as_completed(futures):
                try:
                    fut.result()
                except Exception as ex_write:
                    logger.error(f"Write Failure: {ex_write}")
        manifest_mgr.save()

    return final_data_map


def harden_1d_with_snapshot(ram_data: Dict[str, Dict], snapshot_map: Dict, manifest_mgr: ManifestManager, root: Optional[str] = None):
    """
    [V3.37] Final stabilization phase: Applies the fresh post-sync snapshot to the 1D parquets
    and performs a high-speed parallel disk write.
    """
    if not ram_data or not snapshot_map:
        return

    now = datetime.now()
    write_queue = []
    
    for y_orig, data in ram_data.items():
        y_canonical, _, _ = resolve_symbol(y_orig)
        snap = snapshot_map.get(y_canonical)
        if snap and "1d" in data:
            df_1d = data["1d"]
            d = pd.Timestamp(now.date())
            ltp = float(snap.get("ltp", 0.0))
            if ltp > 0:
                # Update the 1D bar with latest LTP and market stats
                df_1d.loc[d] = {
                    "open": float(snap.get("open", 0.0)),
                    "high": float(snap.get("high", 0.0)),
                    "low": float(snap.get("low", 0.0)),
                    "close": ltp,
                    "volume": float(snap.get("volume", 0.0))
                }
                data["is_updated"] = True
                write_queue.append(data)

    if write_queue:
        logger.info(f"[HARDEN] Persisting snapshot-based 1D updates for {len(write_queue)} symbols...")
        with ThreadPoolExecutor(max_workers=32) as ex:
            futures = [ex.submit(save_symbol_to_disk, d, manifest_mgr, root) for d in write_queue]
            for fut in as_completed(futures):
                try: fut.result()
                except Exception as e: logger.error(f"Harden Write Fail: {e}")
        manifest_mgr.save()
