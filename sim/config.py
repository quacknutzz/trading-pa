"""
sim/config.py -- Paper-trading simulator & Net-Profit Maximizer settings.

Supports 3 switchable strategy modes:
  1. "v6_profit_max"   (Default) -- Mathematical Net-Profit Maximizer (V6++):
                                    Fixes #1-#8 + Hybrid Multi-Factor Structural Exit + 2-Stage Tax-Aware Ratchet
                                    + 2-Bar Smoothed Rising ADX + Per-Stock Cooldown + Strict Equal Slot Sizing
  2. "v6_fixes_1_to_5"           -- Approved Fixes #1 to #5 on top of V6
  3. "v6_baseline"               -- Exact replica of Dad's Module_COCKPIT_V6 VBA code
"""
from datetime import time as dtime

DEFAULT_STRATEGY = "quacknutzz_strat"
STRATEGY_LABELS = {
    "quacknutzz_strat": "Quacknutzz_Strat",
    "v6_profit_max": "Quacknutzz_Strat",
    "v6_baseline": "Original Baseline",
}

# ---------------------------------------------------------------- strategy (V6 Baseline constants)
EMA_FAST, EMA_SLOW = 10, 30            # 10m entry cross & 5m exit cross
PULLBACK_EMAS = (10, 30, 50, 100, 200) # 5m EMAs used for the resting limit price
ENTRY_LIMIT_WAIT_BARS = 3              # cancel unfilled limit after 3 completed 5m bars

RSI_FILTER_ENABLED = True
BUY_RSI_MIN, BUY_RSI_MAX = 40.0, 60.0    # VBA code (comment says 52-64)
SELL_RSI_MIN, SELL_RSI_MAX = 38.0, 60.0  # VBA code (comment says 38-47)

ADX_PERIOD = 14
STOCK_ADX_MIN = 20.0                   # entry requires 5m ADX >= 20
ADX_EXIT_BELOW = 18.0                  # exit open trade if 5m ADX < 18

SL_PCT = 0.01                          # 1.0% stop (V6 Baseline & Fixes 1-5)
TGT_PCT = 0.012                        # 1.2% target (V6 Baseline & Fixes 1-5)
MAX_CONCURRENT_POSITIONS = 3           # open + pending entries share these slots

ENTRY_START = dtime(9, 25)
ENTRY_CUTOFF = dtime(14, 45)           # V6 Baseline cutoff (V6++ uses 13:15 to avoid 3 PM time-decay chop)
OPT_ENTRY_CUTOFF = dtime(13, 15)
FORCE_LIQUIDATION = dtime(15, 5)       # 15:05 IST forced intraday square-off

# ---------------------------------------------------------------- Morning True Momentum Gate (09:15 - 09:45 AM)
TRUE_MOMENTUM_GATE_ENABLED = True      # Filter opening noise dynamically instead of blindly skipping 30m
MORNING_RVOL_MIN = 1.25                # Institutional volume footprint (rejects low-volume <1.0x retail drifts)
MORNING_MAX_ADVERSE_WICK = 0.35        # Reject false breakout wick rejections (shooting stars / hammers)
MORNING_NIFTY_ALIGNMENT = True         # Require stock momentum to align with macro NIFTY direction

# ---------------------------------------------------------------- virtual account
STARTING_CAPITAL = 50_000.0            # Rs. -- purely virtual (default 50k)
LEVERAGE = 2.0                         # 2.0 = Standard 2x MIS Intraday Margin (50% margin)

# ---------------------------------------------------------------- Indian costs (equity intraday)
BROKERAGE_PER_ORDER = 20.0             # Rs. 20 or 0.03%, whichever is lower (FYERS intraday)
BROKERAGE_PCT_CAP = 0.0003
STT_SELL = 0.00025                     # 0.025% on sell side
NSE_TXN = 0.0000297                    # exchange transaction charge, both sides
SEBI_PER_CRORE = 10.0                  # both sides
STAMP_BUY = 0.00003                    # 0.003% on buy side
GST = 0.18                             # on brokerage + txn + sebi
MARKET_SLIPPAGE = 0.0005               # 0.05% applied to market-order exits only (0% on resting limit targets)

# ---------------------------------------------------------------- data / server
INDEX_TICKERS = {"NIFTY 50": "^NSEI", "SENSEX": "^BSESN"}
LIVE_POLL_SECONDS = 60
BAR_GRACE_SECONDS = 20                 # wait this long after a 5m bar ends before trusting it
HISTORY_DAYS = 60                      # 60 trading days (~12 weeks) cached for multi-week backtest & warmup
SERVER_PORT = 8765
