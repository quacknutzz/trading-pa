"""
config.py — Centralized configuration for AI Signal Engine.

Reads:   Environment variables (YFFYERS_ROOT)
Returns: All constants, column maps, holidays, strategy parameters
"""

import os
import tempfile
import multiprocessing
import cProfile
import yaml

# ===============================
# SCRIPT IDENTITY
# ===============================
SCRIPT_NAME = "AI_Signal_Engine"

# ===============================
# STRATEGY PARAMETERS
# --- UNTOUCHABLE V1.0 LOCK: CORE ARCHITECTURE ---
# Do not modify these. If modified, the version must be bumped to V1.1.
MIN_RR_RATIO = 1.5              # Minimum reward:risk ratio for trade entry (STATIC 1.5R TARGET)
EXPECTANCY_SLIPPAGE_R = 0.05    # Assumed R slippage per trade
# MODEL_VERSION moved to line ~99 (V3.0 block) — single definition
# -----------------------------------------------

ATR_MULT = 2.0
ADX_MIN = 20
HOSTILE_SCORE = 40
MAX_TARGET_ATR_MULT = 5.0       # Cap target at 5x ATR for day-trading safety
REVERSION_FACTOR = 0.4          # Target 40% of the move to VWAP if no levels found
NEAR_EMA_THRESHOLD = 0.003      # Price within 0.3% of EMA → "near EMA"
DOJI_THRESHOLD = 0.03           # Body-to-range ratio below this → Doji

# ===============================
# PRO-GRADE DYNAMIC PARAMETERS (v3)
# ===============================
# --- SIGNAL QUALITY ---
RVOL_THRESHOLD = 0.4            # Reduced from 0.5 to unblock early-session trades
MIN_DAILY_TURNOVER_CR = 15.0    # [NEW] Absolute liquidity floor: Reject symbols with < ₹15 Crores average daily turnover
MIN_5M_TRADED_VALUE_RS = 1500000 # [NEW] Intraday Execution Floor: ₹15 Lakh median 5m traded value (Protects vs toxic tail)
MIN_VOLUME_CONSISTENCY = 0.35    # [NEW] Consistency metric (Currently logged for research only, not blocking)
MIN_LTP_THRESHOLD = 20.0        # [NEW] Reject penny stocks under ₹20
ADX_MIN_STRENGTH = 24           # Minimum Trend Strength (Relaxed from 30)
CONFLUENCE_TOLERANCE = 0.0015   # 0.15% tolerance for "Cluster" detection (EMA+Fib)
ATR_EXTENSION_MULT = 3.0        # Price > 3xATR from 20EMA is "Overextended"
ER_THRESHOLD = 0.20             # [NEW] Layer 1 Selection Filter (Minimum efficiency)
ER_PERIOD = 10                  # [NEW] 10-day lookback for trend efficiency

# --- EXECUTION LEVEL GEOMETRY (Desk-grade hygiene) ---
# Bound stop distance in ATR units so RR/levels are not artifacts of tiny stops or huge risk.
# Keyed by horizon (bars): shorter horizons use tighter stops (less noise to clear).
# Default for horizons not explicitly listed.
_EXEC_MIN_STOP_ATR_MAP = {6: 0.8, 9: 0.9, 12: 1.0, 18: 1.1, 24: 1.1, 36: 1.2, 48: 1.2}
_EXEC_MAX_STOP_ATR_MAP = {6: 2.0, 9: 2.5, 12: 3.0, 18: 3.5, 24: 3.5, 36: 4.0, 48: 4.0}
EXEC_MIN_STOP_ATR = 1.2           # fallback default
EXEC_MAX_STOP_ATR = 4.0           # fallback default
EXEC_TARGET_RR = 2.0

def get_exec_stop_bounds(horizon_bars=None):
    """Return (min_atr, max_atr) for the given horizon, with fallback defaults."""
    h = int(horizon_bars or 0)
    return (
        _EXEC_MIN_STOP_ATR_MAP.get(h, EXEC_MIN_STOP_ATR),
        _EXEC_MAX_STOP_ATR_MAP.get(h, EXEC_MAX_STOP_ATR),
    )

# --- DIRECTIONAL SIGNAL GATES (V6.1) ---
V3_LONG_Z_THRESHOLD = 1.65
V3_SHORT_Z_THRESHOLD = -1.65
V3_LONG_PROB_THRESHOLD = 0.55
V3_SHORT_PROB_THRESHOLD = 0.45
V3_SHORT_FLOOR_MULTIPLIER = 1.20 # Slightly lower than before to allow some shorts

# --- SETUP PLAYBOOK THRESHOLDS (Institutional Alignment) ---
PLAYBOOK_ADX_MIN = 24           # Threshold for Trend Pullback & Alpha Breakout
PLAYBOOK_RVOL_MIN = 1.7         # Threshold for Institutional Accumulation
BROKER_SETTLE_BUFFER = 15       # 15s delay to ensure candle settlement on Fyers

# --- PERSISTENCE & LIFECYCLE (NEW) ---
LOCKED_PERSISTENCE_CYCLES = 3    # Min bars to hold signal once LOCKED
ARMED_TIMEOUT_CYCLES = 3         # Max bars to stay ARMED without trigger
REGIME_INERTIA_CYCLES = 2        # Min bars for a regime to be "stable"
SWEEP_REENTRY_WINDOW = 2         # Max bars to allow re-entry after stop hit

# --- MOMENTUM OVERDRIVE GATES ---
OVERDRIVE_ATR_PCT_THRESHOLD = 0.7 # 70th percentile
OVERDRIVE_RVOL_THRESHOLD = 1.8
OVERDRIVE_DISPLACEMENT_MULT = 0.5 # 0.5x ATR movement
OVERDRIVE_DISPLACEMENT_LOOKBACK = 2 # 2 candles

# --- BREAKOUT PARAMETERS ---
BREAKOUT_N = 5                     # Default Donchian lookback for breakout detection (backtest sweeps 5-25)

# --- NIFTY REGIME GATES (Institutional Throttling) ---
NIFTY_OVERRIDE_BULL = 85        # Explosive / Runaway Trend (Ignore ADX)
NIFTY_OVERRIDE_BEAR = 15        # Total Crash (Ignore ADX)
NIFTY_REGIME_BULL   = 75        # Strong Bullish (Unlock Primary Pivots)
NIFTY_REGIME_BEAR   = 35        # Significant Weakness / Crashing

# --- RS/RW SELECTION GATES (Industry View) ---
RS_SLOPE_LOOKBACK = 10          # 10 bars (50m) for Relative Strength Slope (Industry Standard)
RS_SLOPE_THRESHOLD = 0.0        # Slope >= 0 to be considered a market participant
STK_VWAP_LAGGARD_TOLERANCE = 0.001 # 0.1% tolerance below VWAP during Bull Index
CONTRA_NIFTY_THRESHOLD = 35.0  # Nifty must be below this to be considered "weak" for contra logic
CONTRA_STK_STRENGTH = 85.0     # Stock must be this strong to override a weak nifty

EXPECTANCY_SLIPPAGE_R = 0.05    # Assumed R slippage per trade
BREADTH_TOXICITY_THRESHOLD = 0.3 # If A/D is below this, market is toxic
HOSTILE_NEWS_SCORE = 3.5      # News score below this blocks Longs
MAX_PORTFOLIO_R = 5           # Maximum total R (risk) across all open trades
MAX_SECTOR_R = 2              # Maximum R (risk) per sector
V3_RISK_STALENESS_THRESHOLD_MIN = 3.0  # Minutes; block trading if risk file older than this

# --- SMART ENTRY FILL (Fix#1) ---
ENTRY_CHASE_ATR_MULT    = 1.5   # Increased from 1.0 to allow chasing strong momentum
ENTRY_REANCHOR_ATR_MULT = 1.0   # Increased from 0.5
ENTRY_ABANDON_ATR_MULT  = 3.5   # Increased from 2.0 to allow for large breakout gaps

# --- TURNOVER CONTROL (V7.6 Institutional Grade) ---
ENABLE_TURNOVER_CONTROL = True # Enabled as per approved plan
HYSTERESIS_THRESHOLD = 0.12      # Tuned from 0.15 to 0.12 as per plan
STRONG_REVERSAL_THRESHOLD = 0.30 # Jump bypass (Override hold if delta > 30%)
MIN_HOLD_BARS = 3                # Minimum bars (15 min) before a signal can be reversed

# ── Signal Lifecycle ──
LIFECYCLE_ENABLED = True
LIFECYCLE_MATURING_BARS = 3          # Bars before MATURING → ACTIVE
LIFECYCLE_WEAKENING_MAX_BARS = 4     # Max bars in WEAKENING before BROKEN
LIFECYCLE_CORE_SIMILARITY_ACTIVE = 0.85   # Min core similarity to stay ACTIVE
LIFECYCLE_CORE_SIMILARITY_WEAK = 0.80     # Below this → WEAKENING
LIFECYCLE_CORE_SIMILARITY_BROKEN = 0.70   # Below this in MATURING → immediate BROKEN
LIFECYCLE_EV_RATIO_FLOOR = 0.30      # current_ev / entry_ev below this → WEAKENING
LIFECYCLE_CONTINUATION_CONFLUENCE = 0.5   # Min confluence score for continuation


# --- PERSISTENCE & EXIT (THE CONVICTION ENGINE) ---
MAX_HOLD_MINUTES = 45           # Max time to hold a stagnant trade w/o profit
EXIT_STALL_PROFIT = 0.002       # Minimum profit (0.2%) required to avoid "Stall Exit"
CONVICTION_BULL_THRESHOLD = 60  # Score to hold BUY through dips (Bull Regime)
CONVICTION_BEAR_THRESHOLD = 40  # Score to panic EXIT (Bear Regime)

# --- STRUCTURE ANCHORS ---
USE_VWAP_EXIT = True            # Exit if Price < VWAP (Intraday Value Break)
USE_SWING_EXIT = True           # Exit if Price < Recent Swing Low (Dow Theory Break)

# --- AI LEARNING ---
LEARNING_BATCH_SIZE = 5         # Update weights after N trades

# --- DATA ENGINE DEFAULTS ---
# These are the default values if settings.yaml is missing or incomplete.
API_RATE_LIMIT = 10              # Calls per second; Fyers 429s if live sync bursts too hard
SYNC_WORKERS = 10                  # Parallel sync threads
NEWS_WORKERS = 16                 # Parallel news fetchers
COMPUTE_WORKERS = 12              # ProcessPool — matches i5 12-thread for indicator parallelism

# --- DYNAMIC YAML OVERRIDES (v3.5) ---
def _apply_yaml_overrides():
    try:
        import yaml
        # Load settings.yaml from project root
        yaml_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "settings.yaml")
        if os.path.exists(yaml_path):
            with open(yaml_path, "r") as f:
                data = yaml.safe_load(f)
                if data:
                    globals().update(data)
    except Exception as e:
        print(f"⚠️ YAML Config Loader Error: {e}")

_apply_yaml_overrides()

# ===============================
# AI REGRESSION PARAMETERS (V3.0)
# ===============================
# [WINNING CONFIG: SURVIVOR SNIPER]
EV_BUY_THRESH     = 0.55         # 90th percentile predicted move (in ATR)
EV_SELL_THRESH    = -0.55
UNCERTAINTY_LIMIT = 0.30         # std(preds) / abs(mean(pred)) < 0.30
META_LABEL_THRESH = 0.60         # 60% probability of being a winning trade
MAX_HOLD_BARS     = 25           # Global timeout fallback (5 min bars)
MODEL_VERSION = "v3.0"          # V3.0: Cross-Sectional Ensemble + AFML Pipeline
V3_USE_NUMBA_TA = True         # False strictly guarantees 100% data parity with the pre-trained weights


# ===============================
# V3.0 ML PIPELINE CONFIG
# ===============================
# All ML parameters centralized here — change config, not code.

# --- ENSEMBLE MODELS ---
V3_ENSEMBLE_MODELS = ["lightgbm", "xgboost"]  # 2-model ensemble (CatBoost/ElasticNet dropped for speed/NaN handling)
V3_STACKING_META = "ridge"                                 # RidgeCV for regression stacking
V3_ENSEMBLE_WEIGHTS = None  # None = learned by meta-learner
V3_TB_HORIZON = 3           # [Phase 4] Locked tightly to 15-minute holding horizon (3 × 5min) for Order Flow tracking

# --- ELASTICNET HYPERPARAMS (V7: Linear stability anchor) ---
V3_ELASTICNET_PARAMS = {
    "alpha": 0.01,
    "l1_ratio": 0.5,       # Balance between L1 (sparsity) and L2 (stability)
    "max_iter": 1000,
    "random_state": 42,
}

# --- LIGHTGBM HYPERPARAMS (V5: LambdaRank for direct ranking optimization) ---
V3_LGBM_PARAMS = {
    "objective": "lambdarank",      # V5: Direct ranking optimization (NDCG)
    "metric": "ndcg",
    "boosting_type": "gbdt",
    "num_leaves": 63,
    "max_depth": 7,
    "learning_rate": 0.05,
    "feature_fraction": 0.8,
    "bagging_fraction": 0.8,
    "bagging_freq": 5,
    "min_child_samples": 50,
    "lambda_l1": 0.1,
    "lambda_l2": 1.0,
    "n_estimators": 300,
    "verbose": -1,
    "n_jobs": -1,
    "random_state": 42,
}
# V5: Regression-fallback params (for CatBoost/XGBoost which don't use LambdaRank)
V3_LGBM_PARAMS_REG = {**V3_LGBM_PARAMS, "objective": "regression", "metric": "mse"}

# --- V4.3 EXPERT-SPECIFIC REGULARIZATION ---
LGBM_TREND_PARAMS = {
    **V3_LGBM_PARAMS,
    "lambda_l1": 0.5,      # High sparsity
    "lambda_l2": 5.0,      # Very high stability
}

LGBM_VOLATILE_PARAMS = {
    **V3_LGBM_PARAMS,
    "lambda_l1": 0.2,      # Medium sparsity
    "lambda_l2": 2.0,      # Medium stability
}

LGBM_RANGE_PARAMS = {
    **V3_LGBM_PARAMS,
    "lambda_l1": 0.1,      # Low sparsity
    "lambda_l2": 1.0,      # Low stability
}

# --- V4.3 EXPERT DOMAIN ANCHORS ---
EXPERT_ANCHORS = {
    "TREND": ["EMA_200_dist", "ADX_ADX_14", "rank_mom_3bar"],
    "RANGE": ["BB_width_pct", "RSI_14"],
    "VOLATILE": ["BB_width_pct", "RSI_14"]
}

# --- V4.3 TEMPORAL SAMPLE WEIGHTING LAMBDAS ---
TEMPORAL_DECAY_LAMBDA_GLOBAL = 1.0
TEMPORAL_DECAY_LAMBDA_TREND = 1.8   # Stronger recency bias for trending regimes

# --- CATBOOST HYPERPARAMS ---
V3_CATBOOST_PARAMS = {
    "iterations": 300,
    "learning_rate": 0.05,
    "depth": 7,
    "l2_leaf_reg": 3.0,
    "random_seed": 42,
    "verbose": 0,
    "loss_function": "RMSE",        # Regression for ranked target
    "task_type": "CPU",             # Switch to "GPU" if available
}

# --- XGBOOST HYPERPARAMS ---
V3_XGBOOST_PARAMS = {
    "objective": "reg:squarederror",
    "tree_method": "hist",
    "grow_policy": "lossguide",
    "max_bin": 256,
    "max_leaves": 128,
    "max_depth": 7,
    "learning_rate": 0.05,
    "n_estimators": 300,
    "subsample": 0.8,
    "colsample_bytree": 0.8,
    "reg_alpha": 0.1,
    "reg_lambda": 1.0,
    "min_child_weight": 50,
    "verbosity": 0,
    "random_state": 42,
    "n_jobs": -1,
}

# --- CLASSIFIER VARIANTS (Triple Barrier Binary Target) ---
V3_LGBM_PARAMS_CLF = {**V3_LGBM_PARAMS, "objective": "binary", "metric": "auc"}
V3_CATBOOST_PARAMS_CLF = {**V3_CATBOOST_PARAMS, "loss_function": "Logloss", "eval_metric": "AUC"}
V3_XGBOOST_PARAMS_CLF = {**V3_XGBOOST_PARAMS, "objective": "binary:logistic", "eval_metric": "auc"}

# --- TRAINING DATA ---
V3_TRAINING_DAYS = 1460               # Historical data window for training
V3_MIN_SYMBOL_BARS = 200              # Minimum bars to include a symbol
V3_SURVIVORSHIP_MIN_DAYS = 180        # Flag symbols with less history (survivorship bias)
V3_FORWARD_RETURN_BARS = 3            # Forward look for cross-sectional return (3 bars = 15 min)

# --- CUSUM FILTER ---
V3_CUSUM_ENABLED = True               # Enable event-based sampling
V3_CUSUM_THRESHOLD = 0.002             # [Phase 4] Scaled appropriately down for a high-frequency 15m horizon

# --- FRACTIONAL DIFFERENTIATION ---
V3_FRACDIFF_ENABLED = True            # Enable fractional differentiation
V3_FRACDIFF_D = 0.40                  # Differentiation order (0.35-0.45 typical)
V3_FRACDIFF_WINDOW = 100              # Fixed-width FFD window (bars)
V3_FRACDIFF_FEATURES = ["close", "vwap", "volume", "high", "low"]  # Universal FFD

# --- CROSS-SECTIONAL NORMALIZATION ---
V3_CS_NORM_ENABLED = True             # Enable RobustZScoreNorm per trading day
V3_CS_NORM_CLIP = 3.0                 # Clip outliers at ± this many MADs

# --- FEATURE SELECTION ---
V3_CORR_FILTER_THRESHOLD = 0.95      # Drop features with correlation > this
#V3_SHAP_IMPORTANCE_CUTOFF = 0.90     # Keep features contributing ≥90% cumulative SHAP
V3_SHAP_IMPORTANCE_CUTOFF = 0.88     # Keep features contributing ≥88% — balanced between noise (0.95) and collapse (0.80)

# --- WALK-FORWARD OPTIMIZATION ---
V3_WFO_PURGE_BARS = 3                # Purge gap BEFORE test fold (tied exactly to 15m horizon)
V3_WFO_EMBARGO_BARS = 3              # Embargo gap AFTER test fold (tied exactly to 15m horizon)
V3_WFO_MIN_TRAIN_SAMPLES = 500
V3_WFO_MIN_TEST_SAMPLES = 100

# --- V5 EXECUTION THRESHOLDS (LambdaRank Relevance Grade Output 0-4) ---
# NOTE: V5 model outputs relevance scores (~0-4 scale from LambdaRank)
# BUY: stock must score above median (grade > 2.5)
# SELL: stock must score below median (grade < 1.5)
V3_EV_THRESHOLD_BUY  = 0.85     # Top 15% relevance
V3_EV_THRESHOLD_SELL = 0.15     # Bottom 15% relevance
V3_EV_UNCERTAINTY_MAX= 0.50     # Disagreement / abs(EV) < 0.5
V3_EV_NOISE_FLOOR    = 0.70     # Skip stocks below 70th percentile

# --- V3 POSITION SIZING (Risk Mitigation) ---
V3_MAX_POS_PCT       = 0.05     # 5% institutional portfolio cap
V3_KELLY_SAFETY      = 0.25     # Safe fractional Kelly (Quarter Kelly)

# --- SAMPLE UNIQUENESS ---
V3_UNIQUENESS_WEIGHTING = True        # Weight samples by average uniqueness


# --- META-LABELER ---
V3_META_TRAIN_ALL_SYMBOLS = True      # Use ALL symbols, not just 50
V3_META_SIGNAL_THRESHOLD_BUY = 0.55   # Match production threshold exactly
V3_META_SIGNAL_THRESHOLD_SELL = 0.45  # Match production threshold exactly
V3_META_BALANCED_CLASSES = True       # Use scale_pos_weight for class balance
V3_META_OPTIMIZE_F1 = True            # Find optimal threshold by maximizing F1
V3_META_N_ESTIMATORS = 200
V3_META_LEARNING_RATE = 0.03

# --- SIGNAL OUTPUT ---
V3_SIGNAL_MODE = "all_qualified"      # "all_qualified" or "top_k"
V3_ALPHA_BUY_THRESHOLD = 0.65        # Alpha score > this → BUY signal
V3_ALPHA_SELL_THRESHOLD = 0.35       # Alpha score < this → SELL signal
V3_CALIBRATION_METHOD = "platt"       # "platt" (Platt Scaling) or "isotonic"
V3_SIGNAL_SELECTIVITY_Z = 1.5         # Default Z (Sideways / Ambiguous regime)
V3_MIN_META_CONV = 6.0               # Absolute minimum scaled conviction
V3_SHORT_FLOOR_MULTIPLIER = 1.15     # Softened from 1.25 to improve accessibility for leaders
V3_TREND_ALIGNED_Z = 1.25             # Relaxed Z for trend-aligned trades (Top 10%)
V3_MARKET_BIAS_DEADZONE = 0.3        # Market mean must exceed ±0.3 to classify bull/bear
V3_TREND_SECTOR_PENALTY_MULT = 0.5   # Hostile sector penalty multiplier in TREND regime
V3_MAX_GLOBAL_SIGNALS = 25           # Maximum total signals allowed per 5m cycle
V3_MAX_SECTOR_SIGNALS = 5            # Maximum signals allowed per sector per cycle

# --- TECHNICAL CONFLUENCE ENGINE (V3.2) ---
# Replaces ML directional probability path when model is unavailable/weak.
# The score is now category-gated rather than a loose additive count.
CONFLUENCE_SCORE_MIN = 6.0              # Weighted score out of 10 required for an actionable setup
CONFLUENCE_TREND_WEIGHT = 0.20          # Daily + 1h bias weight
CONFLUENCE_STRUCTURE_WEIGHT = 0.35      # 15m structure + VWAP/sector alignment
CONFLUENCE_TRIGGER_WEIGHT = 0.30        # 5m trigger quality (price action + momentum)
CONFLUENCE_RISK_WEIGHT = 0.15           # Participation + execution hygiene + session gate
CONFLUENCE_MIN_RVOL = 1.05              # Base participation floor
CONFLUENCE_BREAKOUT_RVOL = 1.25         # Extra participation needed for breakout entries
CONFLUENCE_MAX_STRETCH_ATR = 1.50       # Reject fresh entries too far from VWAP anchor
CONFLUENCE_PULLBACK_MAX_DIST_ATR = 0.80 # Pullback must still be near value
CONFLUENCE_REVERSAL_MAX_DIST_ATR = 0.35 # Reversal must happen at obvious structure
CONFLUENCE_TRIGGER_BUFFER_ATR = 0.15    # Structural trigger tolerance around S/R
CONFLUENCE_MIN_NET_RR = 1.5             # Fresh entry must meet minimum net R:R after assumed costs
CONFLUENCE_SIGNAL_TTL_BARS = 4          # Fresh discretionary setup expires after 4 bars (20 min) for manual execution
CONFLUENCE_ACTIONABLE_TIERS = ("A+", "A", "B")
CONFLUENCE_USE_ML_FALLBACK = True       # Route to confluence when ML prob ~0.5 (no signal)
CONFLUENCE_ML_FALLBACK_WINDOW = 0.03    # Prob within 0.5 +/- this → use confluence
CONFLUENCE_ML_UNAVAILABLE_FALLBACK = True  # Also use confluence when ML unavailable (prob=0)
CONFLUENCE_RVOL_LUNCH_MULT = 0.75      # RVOL floor multiplier during 12:00-13:30 lunch session

# --- CANDLESTICK PATTERN STRENGTH GRADES ---
# Strong (1.0): High-conviction reversal/continuation patterns
# Moderate (0.6): Decent patterns needing confirmation
# Weak (0.3): Low-confidence patterns, barely actionable
PATTERN_STRENGTH = {
    "Bullish Engulfing": 1.0, "Bearish Engulfing": 1.0,
    "Marubozu": 1.0,
    "Three White Soldiers": 1.0, "Three Black Crows": 1.0,
    "Morning Star": 1.0, "Evening Star": 1.0,
    "Hammer": 0.6, "Shooting Star": 0.6,
    "Piercing Line": 0.6, "Dark Cloud Cover": 0.6,
    "Tweezer Bottom": 0.6, "Tweezer Top": 0.6,
    "Bullish Harami": 0.3, "Bearish Harami": 0.3,
    "Dragonfly Doji": 0.3, "Gravestone Doji": 0.3,
}

# --- POSITION SIZING ---
V3_POSITION_SIZING = "half_kelly"     # "half_kelly", "fixed_fraction", "equal"
V3_KELLY_FRACTION = 0.5              # Fraction of Kelly (0.5 = half-Kelly)
V3_FIXED_RISK_PCT = 0.02             # 2% of capital per trade (for fixed_fraction)
V3_MAX_POSITION_PCT = 0.10           # Max 10% of capital in any single position

# --- PORTFOLIO RISK ---
V3_MAX_SECTOR_CONCENTRATION = 0.40   # Max 40% of signals from one sector
V3_MAX_CORRELATED_SIGNALS = 3        # Max 3 highly correlated signals at once
V3_CORRELATION_THRESHOLD = 0.70      # Signals with > 70% return correlation = correlated

# --- TRANSACTION COSTS (India NSE) ---
V3_BROKERAGE_PER_ORDER = 20.0        # ₹20 flat per order (Zerodha-style)
V3_STT_SELL_PCT = 0.00025            # 0.025% STT on sell side (intraday)
V3_EXCHANGE_CHARGE_PCT = 0.0000325   # 0.00325% per side
V3_GST_PCT = 0.18                    # 18% GST on brokerage + exchange charges
V3_SLIPPAGE_PCT = 0.0002             # 0.02% slippage per side (conservative)
V3_COST_GATE_ASSUMED_NOTIONAL_RS = 100000.0  # Explicit notional for converting flat brokerage into pct in deploy gating

# --- OPTUNA ---
V3_OPTUNA_ENABLED = True             # Enable Bayesian hyperparameter optimization
V3_OPTUNA_N_TRIALS = 50              # Number of Optuna trials per retraining
V3_OPTUNA_METRIC = "ic"              # Optimize for Information Coefficient

# --- ONLINE LEARNING (River + ADWIN) ---
V3_ONLINE_LEARNING_ENABLED = True
V3_ADWIN_DELTA = 0.002               # ADWIN sensitivity (lower = more sensitive)
V3_MICRO_RETRAIN_WINDOW = 60         # Days of recent data for micro-retraining
V3_ONLINE_BLEND_WEIGHT = 0.3         # Weight of online model vs offline (0.3 = 30% online)

# --- BACKTEST ---
V3_BACKTEST_START = "2026-02-10"     # 4-week backtest start
V3_BACKTEST_END = "2026-03-07"       # 4-week backtest end

# --- RELATIVE FEATURES (Nifty symbol for excess return calculation) ---
V3_BENCHMARK_SYMBOL = "NIFTY50"       # Benchmark index for relative returns
V3_SECTOR_MAP_FILE = "sector_map.json"  # Sector classification file (optional)

WIN_EFFICIENCY_THRESHOLD = 1.0  # Win = (MFE / Risk) > 1.0 (Did we capture the move?)

# ===============================
# DATA PARAMETERS
# ===============================
EMA_PERIODS = [10, 20, 50, 100, 200]
EMA_CROSS_SLOW = 30   # Slow EMA for crossover detection + Excel display columns AT/AY/BD/BI
FIB_LEVELS = [0.236, 0.382, 0.5, 0.618]
BOOTSTRAP_5M_DAYS = 180         # 6 months of 5m candle history
# 400 bars (~5.3 days) ensures EMA 200 and RSI 14 have enough history to converge perfectly.
INFERENCE_HISTORY_5M = 400    # 3.3 trading days — sufficient for EMA200, ATR14, Bollinger20 on 5M
INFERENCE_HISTORY_1D = 1000   # Sliced window for live 1D inference

# Live calc optimization
TECHNICAL_LOOKBACK_5M = 3000  # Increased for RSI-1H stability (3000/12=250 hourly bars)
TIMEFRAMES = ["5m", "15m", "1h", "1d"]

# ===============================
# NUMBA CACHE
# ===============================
try:
    cache_dir = os.path.join(tempfile.gettempdir(), "numba_cache_yffyers")
    if not os.path.exists(cache_dir):
        os.makedirs(cache_dir, exist_ok=True)
    os.environ["NUMBA_CACHE_DIR"] = cache_dir
except Exception:
    pass

# ===============================
# PERFORMANCE PROFILER
# ===============================
ENABLE_PROFILER = False
_profiler = None

if multiprocessing.current_process().name == 'MainProcess' and ENABLE_PROFILER:
    _profiler = cProfile.Profile()
    _profiler.enable()

# ===============================
# PATHS & FILE NAMES
# ===============================
ROOT_DIR_ENV = "YFFYERS_ROOT"
# DEFAULT_ROOT points to the AI/ folder
DEFAULT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# BASE_ROOT is the project level (above AI/)
BASE_ROOT = os.path.dirname(DEFAULT_ROOT)

ROOT_DIR = DEFAULT_ROOT
DATA_DIR = os.path.join(ROOT_DIR, "data")
LOG_DIR = os.path.join(ROOT_DIR, "logs")

PARQUET_SUBFOLDER = os.path.join("data", "master")
DAILY_SUBFOLDER = os.path.join("data", "daily_master")
MANIFEST_FILENAME = "market_data_manifest.json"

# [FIX] Correct AI Dashboard filename
EXCEL_FILENAME = "AI_BTB_Dashboard.xlsm"
EOD_SEAL_TIME = "17:00"
EXCEL_SHEET = "Sheet1"
TOKEN_FILENAME = "fyers_tokens.json"

# [FIX] Prioritize local AI dashboard
def get_excel_path():
    path = os.path.join(DEFAULT_ROOT, EXCEL_FILENAME)
    if os.path.exists(path): return path
    # Fallback to base root if moved
    return os.path.join(BASE_ROOT, EXCEL_FILENAME)
LOG_FILENAME = os.path.join("logs", f"{SCRIPT_NAME}.log")
MODEL_DIR = os.path.join(DATA_DIR, "models")
INTRADAY_DIR_MODEL_DIR = os.path.join(DATA_DIR, "models", "INTRADAY_DIR")
MODEL_FILENAME = "lgbm_model.txt"

# ===============================
# EXCEL COLUMN MAPPING
# ===============================
def colnum(letter: str) -> int:
    """Converts Excel column letter(s) to 1-based column number."""
    n = 0
    for ch in letter.upper():
        if ch.isalpha(): n = n * 26 + (ord(ch) - 64)
    return n

# Full output range
FULL_COLUMNS = [chr(i) for i in range(ord('T'), ord('Z')+1)] + \
               [f"A{chr(i)}" for i in range(ord('A'), ord('Z')+1)] + \
               [f"B{chr(i)}" for i in range(ord('A'), ord('Z')+1)] + \
               [f"C{chr(i)}" for i in range(ord('A'), ord('Z')+1)] + \
               [f"D{chr(i)}" for i in range(ord('A'), ord('Z')+1)] + \
               [f"E{chr(i)}" for i in range(ord('A'), ord('R')+1)]
OUTPUT_START_COL_INDEX = colnum("T")

# --- INDEX MAP ---
FYERS_INDEX_MAP = {
    "NSEI": "NSE:NIFTY50-INDEX",
    "NSEBANK": "NSE:NIFTYBANK-INDEX",
    "CNXFIN": "NSE:FINNIFTY-INDEX",
    "CNXAUTO": "NSE:NIFTYAUTO-INDEX",
    "CNXIT": "NSE:NIFTYIT-INDEX",
    "CNXPHARMA": "NSE:NIFTYPHARMA-INDEX",
    "CNXFMCG": "NSE:NIFTYFMCG-INDEX",
    "CNXMETAL": "NSE:NIFTYMETAL-INDEX",
    "CNXENERGY": "NSE:NIFTYENERGY-INDEX",
    "CNXINFRA": "NSE:NIFTYINFRA-INDEX",
    "CNXREALTY": "NSE:NIFTYREALTY-INDEX",
    "CNXMEDIA": "NSE:NIFTYMEDIA-INDEX",
    "CNXCONSUM": "NSE:NIFTYCONSUMPTION-INDEX",
    "CNXTELECOM": "NSE:NIFTYMSITTELCM-INDEX",
    "CNXINDUSTRIAL": "NSE:NIFTYINDIAMFG-INDEX",
    "CNXCOMMODITIES": "NSE:NIFTYCOMMODITIES-INDEX",
    "CNXPSUBANK": "NSE:NIFTYPSUBANK-INDEX",
    "CNXHEALTHCARE": "NSE:NIFTYHEALTHCARE-INDEX",
    "CNXPSE": "NSE:NIFTYCPSE-INDEX",
    "LARGECAP": "NSE:NIFTY100-INDEX",
    "MIDCAP150": "NSE:NIFTYMIDCAP150-INDEX",
    "SMALLCAP250": "NSE:NIFTYSMLCAP250-INDEX",
}

# --- COLUMN ASSIGNMENTS ---
LTP_COL = "T"
VWAP_COL = "BR"
SPIKE_COL = "BS"
VOL_RATIO_COL = "BO"
SUPERTREND_COL = "BT"
DIST_DAY_HI_COL = "DG"
DIST_DAY_LO_COL = "DH"
MACD_TEXT_COLS = ["Y", "Z", "AA", "AB"]
RSI_COLS = ["U", "V", "W", "X"]
EMA_CROSS_COLS = ["AC", "AD", "AE", "AF"]
EMA10_50_CROSS_COLS = ["AG", "AH", "AI", "AJ"]
BULL_PATTERN_COLS = ["AK", "AL", "AM", "AN"]
BEAR_PATTERN_COLS = ["AO", "AP", "AQ", "AR"]
MACD_NUMERIC_COLS = ["CY", "CZ", "DA", "DB", "DC", "DD"]
ATR_COLS = {"5m": "CC", "15m": "CD"}
DI_PLUS_COL = "BZ"
DI_MINUS_COL = "CA"
ADX_COL = "CB"
STOCH_K_COLS = {"5m": "CE", "15m": "CG"}
STOCH_D_COLS = {"5m": "CF", "15m": "CH"}
BB_UPPER_COLS = {"5m": "CI", "15m": "CL"}
BB_MID_COLS = {"5m": "CJ", "15m": "CM"}
BB_LOWER_COLS = {"5m": "CK", "15m": "CN"}
SECTOR_COL = "BM"
REGIME_5M_COL = "BN"                                          # 5-minute regime for all tickers (MILD_UP/DOWN split)
SECTOR_REGIME_COL = "BO"
RVOL_COL = "BP"
REGIME_COL = "BQ"
SWING_HIGH_COL = "CO"
SWING_LOW_COL = "CP"
FIB_COLS_5M = ["CQ", "CR", "CS", "CT"]
FIB_COLS_15M = ["CU", "CV", "CW", "CX"]
SCORE_COL = "DE"
NARRATIVE_COL = "DF"
AI_CONFIDENCE_COL = "DJ"
AI_GRADE_COL = "DK"
AI_SIGNAL_COL = "BX"
AI_SETUP_COL = "DM"
AI_HEALTH_COL = "BY"
NEWS_SCORE_COL = "DN"
ER_COL = "DP"
SIZE_COL = "DO"
INSTITUTIONAL_SUMMARY_COL = "DQ"

# --- BTB FEATURE LAYER (Shadow / Sheet1 read-only; VBA does NOT consume these in v1) ---
# Sheet1 columns appended after DU (existing last column). Side-neutral rows only:
# ExtDecision/BT_ExtSide are PREVIEWS; the authoritative side-specific decision
# is computed during replay (scripts/backtest_btb_features.py).
BTB_SHEET1_MAP = {
    "BTB_ExtSide":             "DV",
    "BTB_ExtDecision":         "DW",
    "BTB_ExtScore":            "DX",
    "BTB_DataQuality":         "DY",
    "BTB_ExtReason":           "DZ",
    "BTB_CompressionPct":      "EA",
    "BTB_RelStrengthPct":      "EB",
    "BTB_RelStrengthSlope":    "EC",
    "BTB_CloseLocationScore":  "ED",
    "BTB_AbsorptionProxyScore":"EE",
    "BTB_PressureScore":       "EF",
    "BTB_TriggerPx":           "EG",
    "BTB_FailPx":              "EH",
    "BTB_BidPx":               "EI",
    "BTB_AskPx":               "EJ",
    "BTB_BidSize":             "EK",
    "BTB_AskSize":             "EL",
    "BTB_SpreadBps":           "EM",
    "BTB_L1Imbalance":         "EN",
    "BTB_L1ImbalanceSlope":    "EO",
    "BTB_L1Present":           "EP",
    "BTB_L1PressureScore":     "EQ",
    "BTB_X":                   "ER",
}

# BTB threshold seeds (from retired V4.1 extension gate constants; NOT tuned).
BTB_EXT_COIL_WIDTH_SEED = 0.007   # max compression window width pct
BTB_EXT_MAX_ADVERSE_R_SEED = -0.35
BTB_EXT_RS_MIN_SEED = -0.001
BTB_EXT_MAX_CANDLES = 12
BTB_EXT_TIMEOUT_CANDLES = 24
BTB_EXT_WEIGHTS = {   # default 1.0 vector; tuning deferred until replay
    "compression": 1.0,
    "relstrength": 1.0,
    "closeloc":    1.0,
    "vwap":        1.0,
    "absorb":      1.0,
    "l1":          1.0,
}

# DQ bitmask
BTB_DQ_STALE_BAR = 1
BTB_DQ_DEGENERATE_BAR = 2
BTB_DQ_NO_SECTOR = 4
BTB_DQ_NO_ATR = 8

# --- AI_MASTER ABSOLUTE COLUMN CONTRACT ---
AI_MASTER_COLUMN_MAP = {
    "A": "Ticker",
    "B": "Date",
    "C": "LTP",
    "D": "SignalCode",
    "E": "SignalState",
    "F": "TraderAction",
    "G": "AIHealth",
    "H": "Confidence",
    "I": "Grade",
    "J": "Setup",
    "K": "Narrative",
    "L": "Sector",
    "M": "SecondarySector",
    "N": "SectorHealth",
    "O": "SectorSignal",
    "P": "TimeRegime",
    "Q": "NiftyHealth",
    "R": "NiftySignal",
    "S": "RVOL",
    "T": "NewsScore",
    "U": "VWAP",
    "V": "Spike",
    "W": "SuperTrend5m",
    "X": "SuperTrend1D",
    "Y": "ATR5m",
    "Z": "ADX",
    "AA": "ER",
    "AB": "Entry",
    "AC": "Target",
    "AD": "Stop",
    "AE": "RR",
    "AF": "PriceVsVWAPPct",
    "AG": "DistToST5mATR",
    "AH": "TrendAlign",
    "AI": "InstitutionalSummary",
    "AJ": "RankPct",
    "AK": "ConfidenceZ",
    "AL": "Tier",
}

# ===============================
# MOMENTUM DETECTOR
# ===============================
MOM_VELOCITY_THRESHOLD_AM_TREND = 0.018
MOM_VELOCITY_THRESHOLD_AM_RANGE = 0.012
MOM_VELOCITY_THRESHOLD_PM_TREND = 0.015
MOM_VELOCITY_THRESHOLD_PM_RANGE = 0.010
MOM_ACCELERATION_MIN_TREND = 0.0008
MOM_ACCELERATION_MIN_RANGE = 0.0004
MOM_VOL_SPIKE_MIN_AM = 2.2
MOM_VOL_SPIKE_MIN_PM = 1.8
MOM_RANK_ACCEL_MIN = 0.06
MOM_JERK_MIN = 0.00025
MOM_JERK_MIN_PM = 0.00020
MOM_MAX_EXTENSION = 3.5
MOM_AM_START = "09:15"
MOM_AM_END = "12:00"
MOM_PM_START = "13:30"
MOM_PM_END = "14:45"
MOM_SECTOR_MIN_SCORE = 55
MOM_FOLLOWTHRU_CANDLES = 2
MOM_FOLLOWTHRU_VOL_DROP = 0.35

# ===============================
# HOLIDAYS
# ===============================
HARDCODED_HOLIDAYS = {
    "2026-01-15", "2026-01-26", "2026-03-03", "2026-03-26", "2026-03-31",
    "2026-04-03", "2026-04-14", "2026-05-01", "2026-05-28", "2026-06-26",
    "2026-09-14", "2026-10-02", "2026-10-20", "2026-11-10", "2026-11-24", "2026-12-25"
}
EXCEPTION_TRADING_DAYS = {"2026-02-01"}

# ===============================
# RGED-DM (Regime-Gated Event-Driven Directional Model)
# ===============================
RGED_HORIZON = 9                          # 45 min at 5-min bars
RGED_ATR_PERIOD = 20                      # ATR lookback
RGED_PT_MULT = 1.2                        # Upper barrier: +1.2 × ATR
RGED_SL_MULT = 1.0                        # Lower barrier: -1.0 × ATR
RGED_COST_PCT = 0.0012                    # 0.12% round-trip cost
RGED_MAX_HOLD_BARS = 9                    # Exit after 9 bars if no barrier hit
RGED_TRAINING_DAYS = 730                  # 2-year rolling window
RGED_WFO_FOLDS = 5                        # Walk-forward folds
RGED_WFO_PURGE_BARS = 9                   # Purge gap (matches horizon)
RGED_WFO_EMBARGO_BARS = 9                # Embargo gap
RGED_RIDGE_WR_FLOOR = 52.0               # Min win rate for Ridge gate (%)
RGED_RIDGE_RETURN_FLOOR = 0.0            # Min gross avg return for Ridge gate
RGED_LONG_PCTILE = 0.90                   # Top 10% → LONG
RGED_SHORT_PCTILE = 0.10                  # Bottom 10% → SHORT
RGED_DEPLOY_AVG_RETURN = 0.10            # Min avg trade return net (%)
RGED_DEPLOY_WIN_RATE = 55.0              # Min win rate (%)
RGED_DEPLOY_MIN_TRADES = 3000            # Min trade count
RGED_KILL_SWITCH_DAYS = 5                # Days of negative avg return → kill
RGED_USE_REGIME_GATE = False              # Phase 1: bypass regime checks
RGED_MARKET_VOL_THRESHOLD = 1.0           # Market ATR > median (multiplier)
RGED_MARKET_ORB_THRESHOLD = 1.0           # ORB range > median
RGED_MARKET_VWAP_SLOPE_MIN = 0.0001      # Near-zero = flat market
RGED_STOCK_VOL_SPIKE_MIN = 1.5            # Volume > 1.5× median
RGED_STOCK_ATR_PCT_FLOOR = 0.015          # ATR% floor for meaningful movement
RGED_FEATURES = [
    "return_5", "return_10", "return_20",
    "orb_breakout", "ROC_5", "ROC_10",
    "vwap_distance", "intraday_zscore", "BB_position", "STOCHk_14",
    "ATR_10_pct", "ATR_20_pct", "range_compression", "vol_20",
    "volume_spike", "volume_percentile", "RVOL", "CMF_20",
    "dist_from_open", "dist_from_prev_high", "dist_from_prev_low",
    "EMA_slope_5", "EMA_slope_20", "efficiency_ratio", "RSI_14",
]

# ===============================
# V4.5 ADVISORY MODE (ML-FIRST)
# ===============================
# When True: ML pipeline signals flow directly to output as advisory
# recommendations. Execution router, VWAP+ORB, and broker connectivity
# checks are bypassed. The engine advises — you decide.
#
# Evidence: Advisory mode aligns with the industry practice of separating
# alpha research (signal generation) from execution (order management).
# Firms like AQR and Two Sigma run research signals through a paper-trading
# layer before promoting to live execution (Harvey & Liu 2015, JPM).
AI_ADVISORY_MODE = False

# When True: Legacy V3/V4 MoE ensemble (CatBoost + XGBoost + LGBM + HMM)
# is NOT loaded. This ensemble currently produces zero-dispersion
# predictions (spread=0.0000) and wastes ~5s/cycle + ~200MB RAM.
# The ML pipeline (ml_pipeline/) is the authoritative alpha source.
DISABLE_LEGACY_ENSEMBLE = True
