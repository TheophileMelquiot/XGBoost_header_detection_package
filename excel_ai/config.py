# excel_ai/config.py

from pathlib import Path
import os


# ==============================
# ENVIRONMENT
# ==============================

ENV = os.getenv("APP_ENV", "dev")


# ==============================
# MODEL CONFIGURATION
# ==============================

BASE_DIR = Path(__file__).resolve().parent
MODEL_PATH = BASE_DIR / "models" / "header_detector.pkl"
MODEL_VERSION = os.getenv("MODEL_VERSION", "v1")


# ==============================
# PREDICTION SETTINGS
# ==============================

DEFAULT_THRESHOLD = float(os.getenv("HEADER_THRESHOLD", 0.35))
MAX_SCAN_ROWS = int(os.getenv("MAX_SCAN_ROWS", 30))


# ==============================
# PERFORMANCE SETTINGS
# ==============================

ENABLE_LOGGING = os.getenv("ENABLE_LOGGING", "true").lower() == "true"