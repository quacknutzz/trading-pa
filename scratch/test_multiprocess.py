import sys, pickle, time
from datetime import datetime, timedelta, time as dtime
import pandas as pd
import numpy as np
from concurrent.futures import ProcessPoolExecutor

sys.path.insert(0, 'sim')
import datafeed
import config as C
from simulator import Simulator, SymbolData, ema_fast, BAR

print("Checking fast multiprocess execution...")
