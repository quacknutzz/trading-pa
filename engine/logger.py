import os
import sys
import logging

def ensure_dir(path: str) -> None:
    if path:
        os.makedirs(path, exist_ok=True)

def setup_logger(name: str = "AI_BTB", log_file: str = "ai_btb.log") -> logging.Logger:
    logger = logging.getLogger(name)
    if not logger.handlers:
        logger.setLevel(logging.INFO)
        sh = logging.StreamHandler(sys.stdout)
        sh.setFormatter(logging.Formatter("[%(asctime)s] [%(levelname)s] %(message)s", "%H:%M:%S"))
        logger.addHandler(sh)
        
        # Also file handler
        logs_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "logs")
        ensure_dir(logs_dir)
        fh = logging.FileHandler(os.path.join(logs_dir, log_file), encoding="utf-8")
        fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s"))
        logger.addHandler(fh)
    return logger
