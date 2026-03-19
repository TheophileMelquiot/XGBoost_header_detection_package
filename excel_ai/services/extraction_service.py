# excel_ai/services/extraction_service.py

import pandas as pd
from typing import Dict
from ..detector import detect_headers


def detect_and_load(filepath: str, threshold: float = 0.35):

    headers = detect_headers(filepath, threshold)

    xls = pd.ExcelFile(filepath)

    dfs = {}

    for sheet_name, info in headers.items():

        header_row = info["row"]
        confidence = info["confidence"]

        sheet_index = xls.sheet_names.index(sheet_name)

        try:
            df = pd.read_excel(
                filepath,
                sheet_name=sheet_index,
                header=header_row - 1
            )

            dfs[sheet_index] = {
                "sheet_name": sheet_name,
                "header_row": header_row,
                "confidence": confidence,
                #"data": df
            }

        except Exception:
            continue

    return dfs