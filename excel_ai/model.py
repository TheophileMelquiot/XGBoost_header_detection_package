# excel_ai/model.py

import joblib
from functools import lru_cache
from pathlib import Path

MODEL_PATH = Path(__file__).parent / "models" / "header_detector.pkl"

@lru_cache()
def get_model():
    return joblib.load(MODEL_PATH)