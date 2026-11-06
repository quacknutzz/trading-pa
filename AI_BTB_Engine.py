"""
AI_BTB_Engine.py -- Back to Basics Trading Engine.

Replicates the original engine flow: data sync, snapshot fetch,
1D hardening, RAM-CAP trim, indicator computation, Excel write.
Zero AI/ML dependencies.
"""

import os
import sys

os.environ["OMP_NUM_THREADS"] = "1"
os.environ["OPENBLAS_NUM_THREADS"] = "1"
os.environ["MKL_NUM_THREADS"] = "1"
os.environ["VECLIB_MAXIMUM_THREADS"] = "1"
os.environ["NUMEXPR_NUM_THREADS"] = "1"

import time
import logging
from datetime import datetime

logger = None


def warmup_numba():
    """One-time JIT compilation of all Numba kernels."""
    from engine import numba_kernels as nk
    import numpy as np
    logger.info("Warming up Numba Kernels...")
    dummy_c = np.random.random(100).astype(np.float64)
    dummy_h = dummy_c + 1.0
    dummy_l = dummy_c - 1.0
    dummy_v = np.random.random(100).astype(np.float64) * 1000
    dummy_t = np.arange(100, dtype=np.int64) * 300_000_000_000

    nk.resample(dummy_t, dummy_c, dummy_h, dummy_l, dummy_c, dummy_v, 300_000_000_000)
    nk.ema(dummy_c, 20)
    nk.rsi(dummy_c, 14)
    nk.atr(dummy_h, dummy_l, dummy_c, 14)
    nk.adx(dummy_h, dummy_l, dummy_c, 14)
    nk.supertrend(dummy_h, dummy_l, dummy_c, 10, 3.0)
    nk.bbands(dummy_c, 20, 2.0)
    nk.stochastic(dummy_h, dummy_l, dummy_c, 14, 3)
    nk.vwap(dummy_h, dummy_l, dummy_c, dummy_v, dummy_t)
    nk.detect_ema_cross(dummy_c, dummy_c)
    nk.detect_trend_state(10.0, 20.0)
    nk.detect_volume_spike(dummy_v, 0.0, False)
    nk.efficiency_ratio(dummy_c, 10)
    nk.percentile_rank(dummy_c, 50)
    logger.info("Numba Warmup Complete.")


def main():
    global logger
    import multiprocessing
    multiprocessing.freeze_support()

    # 1. Initialize Logger
    from engine.logger import setup_logger, ensure_dir
    logger = setup_logger()

    # 2. Lazy Imports
    from engine import numba_kernels as nk
    from engine.data_engine import (
        orchestrated_sync, ManifestManager, load_fyers_credentials,
        get_root_folder, master_folder_5m,
        fetch_snapshot_batch, harden_1d_with_snapshot,
        resolve_symbol,
    )
    from engine.ta_engine import (
        compute_indicators_for_symbol,
        ExcelWriter,
        read_excel_config,
        build_indicator_row,
    )
    from engine.config import FULL_COLUMNS

    # --- PERSISTENT INITIALIZATION ---
    root = get_root_folder()
    ensure_dir(root)
    ensure_dir(master_folder_5m(root))

    # 3. Read Excel config
    symbols_orig, sec_map, s_row, metadata_map = read_excel_config(root)
    if not symbols_orig:
        logger.error("No symbols found in Excel config.")
        return

    # Inject indices (same as original)
    from engine.config import FYERS_INDEX_MAP
    unique_targets = {}
    for k, v in FYERS_INDEX_MAP.items():
        if v not in unique_targets:
            unique_targets[v] = k
    index_list = list(unique_targets.values())
    symbols = symbols_orig + [idx for idx in index_list if idx not in symbols_orig]
    logger.info(f"Loaded {len(symbols_orig)} stocks + {len(index_list)} indices from Excel config.")

    total_start = time.time()
    logger.info(f"\n{'='*40}\nAI_BTB ENGINE START | {datetime.now().strftime('%H:%M:%S')}\n{'='*40}")

    # 4. Data sync (5M bars + 1D catchup)
    t_sync = time.time()
    manifest = ManifestManager(root)
    app_id, access_token = load_fyers_credentials(root)
    ram_data = orchestrated_sync(symbols, manifest, root, snapshot_map=None)
    sync_elapsed = time.time() - t_sync
    logger.info(f"Orchestrated Sync Phase: {sync_elapsed:.2f}s")

    # 5. Fetch Zero-Latency LTP Snapshot (HTTP batch)
    t_snap = time.time()
    logger.info(f"Snapshot: Fetching {len(symbols)} items via high-speed HTTP batch...")
    snapshot_map = fetch_snapshot_batch(symbols, root) or {}
    snap_elapsed = time.time() - t_snap
    logger.info(f"Snapshot Fetch Phase: {snap_elapsed:.2f}s")

    # 6. 1D Hardening (persist daily parquets)
    harden_1d_with_snapshot(ram_data, snapshot_map, manifest, root)

    # 7. RAM-CAP trim (prevent paging)
    import engine.config as _ecfg
    import pandas as _pd
    _CAP_5M = getattr(_ecfg, "TECHNICAL_LOOKBACK_5M", 1200)
    _CAP_1D = getattr(_ecfg, "INFERENCE_HISTORY_1D", 1000)
    _trimmed = 0
    for _key in list(ram_data.keys()):
        _pack = ram_data[_key]
        if isinstance(_pack, dict):
            for _tf, _limit in [("5m", _CAP_5M), ("1d", _CAP_1D)]:
                _df = _pack.get(_tf)
                if isinstance(_df, _pd.DataFrame) and not _df.empty and len(_df) > _limit:
                    _pack[_tf] = _df.tail(_limit)
                    _trimmed += 1
    if _trimmed:
        logger.info(f"[RAM-CAP] Trimmed {_trimmed} dataframes to {_CAP_5M}/{_CAP_1D} bars")

    # 8. Warmup Numba
    t_warm = time.time()
    warmup_numba()
    warm_elapsed = time.time() - t_warm

    # 9. Compute indicators per symbol
    t_calc = time.time()
    results = []
    for sym in symbols:
        pack = ram_data.get(sym, {})
        if isinstance(pack, dict):
            df_5m = pack.get("5m")
            df_1d = pack.get("1d")
        else:
            df_5m = None
            df_1d = None

        if df_5m is None or df_5m.empty:
            results.append({"symbol": sym, "no_data": True})
            continue

        snap_q = (snapshot_map or {}).get(sym)
        if snap_q is None:
            # snapshot_map is keyed by canonical Yahoo symbol (e.g. 360ONE.NS / ^NSEI);
            # the Excel ticker sym is bare (e.g. 360ONE / NSEI). Resolve before lookup
            # so a live L1 quote actually flows to compute_indicators_for_symbol.
            snap_q = (snapshot_map or {}).get(resolve_symbol(sym)[0]) or (snapshot_map or {}).get(sym)
        meta = metadata_map.get(sym, {})
        ctx = {}
        is_idx = sym in unique_targets
        ctx["is_index"] = is_idx
        sec_key = meta.get("sector") or sec_map.get(sym, "")
        if sec_key and sec_key in ram_data and isinstance(ram_data[sec_key], dict):
            sec_df = ram_data[sec_key].get("5m")
            if isinstance(sec_df, _pd.DataFrame) and not sec_df.empty:
                ctx["sector_5m"] = sec_df
        indicators = compute_indicators_for_symbol(sym, df_5m, df_1d, snapshot_quote=snap_q, metadata=meta, context_data=ctx)
        results.append(indicators)
    calc_elapsed = time.time() - t_calc

    # 10. Write to Excel
    t_write = time.time()
    writer = ExcelWriter(root)
    try:
        success = writer.write_matrix_com(results, s_row)
    except Exception as e:
        logger.error(f"Excel Write Exception (non-fatal): {e}")
        success = False
    write_elapsed = time.time() - t_write

    tot_elapsed = time.time() - total_start

    logger.info(f"""
============================================================
       [PERFORMANCE METRICS SUMMARY]
------------------------------------------------------------
  Orchestrated Sync      :  {sync_elapsed:.2f}s
  Snapshot Fetch         :  {snap_elapsed:.2f}s
  Numba Warmup           :  {warm_elapsed:.2f}s
  Indicator Computation  :  {calc_elapsed:.2f}s  ({len(results)} symbols)
  Excel Write            :  {write_elapsed:.2f}s
------------------------------------------------------------
  TOTAL ELAPSED TIME     :  {tot_elapsed:.2f}s
============================================================""")

    if success:
        logger.info("AI_BTB run complete. Indicators written to Excel.")
    else:
        logger.error("AI_BTB run FAILED.")


if __name__ == "__main__":
    main()
