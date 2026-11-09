"""
scratch/test_refined_abc.py -- Tests removing the 22% idle-cash bottleneck in Fix C
and making Fix B's filter regime-conditional (active in Chop/2023, unthrottled in Trend/2024-2026).
"""
import os
import sys
import time
import pickle
import pandas as pd
import numpy as np

ROOT_DIR = r"c:\Users\rishi\Downloads\Trading Pa"
sys.path.insert(0, ROOT_DIR)
sys.path.insert(0, os.path.join(ROOT_DIR, "sim"))

import sim.config as C
from sim.simulator import Simulator
from scratch.run_multi_year_backtest import load_cached_universe
from scratch.run_ablation_matrix import run_permutation

stock_dfs, index_data, universe = load_cached_universe()
first_date = index_data["NIFTY 50"].index[0].date()
base_sim = Simulator(mode="replay", session_date=first_date, strategy="quacknutzz_strat", starting_capital=40_000.0)
print("Precomputing vectorized indicators...", flush=True)
base_sim.load(stock_dfs, universe, index_data)
_ = base_sim.available_dates()
