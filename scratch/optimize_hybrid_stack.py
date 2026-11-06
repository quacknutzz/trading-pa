"""
scratch/optimize_hybrid_stack.py
Tests combinations on top of Variant 5 (Equal Slot Sizing + Pre-Fill Contra Cancel + Hybrid Structural Exit)
across both Rs 50,000 and Rs 500,000 capital.
"""
import sys, os, pickle
from datetime import datetime, time as dtime
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, '.')
sys.path.insert(0, 'sim')
from scratch.deep_research_v6_profit_max import eval_variant

base = {
    "early_contra_cancel": True,
    "sizing_mode": "equal",
    "exit_mode": "hybrid_5m_vwap",
    "adx_mode": "bar1",
    "entry_cutoff": dtime(13, 15),
    "be_trigger_mfe": 0.35,
    "min_atr": 0.14,
}

combos = [
    ("V5 Base (Hybrid Exit + Equal + EarlyCancel, Cutoff 13:15, BE 0.35, ATR 0.14)", dict(base)),
    ("V5 + Cutoff 13:45", dict(base, entry_cutoff=dtime(13, 45))),
    ("V5 + Cutoff 14:00", dict(base, entry_cutoff=dtime(14, 0))),
    ("V5 + Cutoff 13:45 + BE 0.40%", dict(base, entry_cutoff=dtime(13, 45), be_trigger_mfe=0.40)),
    ("V5 + Cutoff 13:45 + BE 0.45%", dict(base, entry_cutoff=dtime(13, 45), be_trigger_mfe=0.45)),
    ("V5 + BE 0.40% (Cutoff 13:15)", dict(base, be_trigger_mfe=0.40)),
    ("V5 + BE 0.45% (Cutoff 13:15)", dict(base, be_trigger_mfe=0.45)),
    ("V5 + Min ATR 0.16%", dict(base, min_atr=0.16)),
    ("V5 + Min ATR 0.18%", dict(base, min_atr=0.18)),
    ("V5 + ADX bar1_or_3 (2-bar grace)", dict(base, adx_mode="bar1_or_3")),
    ("V5 + ADX sma3 (ADX > 3-bar SMA)", dict(base, adx_mode="sma3")),
    ("V5 + BE 0.40% + Min ATR 0.16% + ADX bar1_or_3", dict(base, be_trigger_mfe=0.40, min_atr=0.16, adx_mode="bar1_or_3")),
    ("V5 + Cutoff 13:45 + BE 0.40% + Min ATR 0.16% + ADX bar1_or_3", dict(base, entry_cutoff=dtime(13, 45), be_trigger_mfe=0.40, min_atr=0.16, adx_mode="bar1_or_3")),
]

for cap in [50000.0, 500000.0]:
    print(f"\n==================== HYBRID STACK COMBINATIONS | CAPITAL: Rs {cap:,.0f} ====================")
    rows = [eval_variant(name, opts, cap=cap) for name, opts in combos]
    print(pd.DataFrame(rows).to_string(index=False))
