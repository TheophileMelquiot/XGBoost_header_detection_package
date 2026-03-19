import openpyxl
import pandas as pd
from .model import get_model
from .feature_engineering import extract_row_features_from_row
from .feature_engineering import FEATURE_NAMES
from .config import MAX_SCAN_ROWS


model = get_model()


# --------------------------------------------------
# Propagation des cellules fusionnées
# --------------------------------------------------

def expand_merged_cells(values):

    expanded = []
    last_value = None

    for v in values:

        if v not in (None, "", "nan"):
            last_value = str(v).strip()

        expanded.append(last_value)

    return expanded


# --------------------------------------------------
# Détection robuste de sous-header
# --------------------------------------------------

def detect_repeating_pattern(values):

    cleaned = []

    for v in values:

        if v in (None, "", "nan"):
            cleaned.append("EMPTY")
        else:
            cleaned.append(str(v).strip().lower())

    n = len(cleaned)

    if n < 4:
        return False

    # ratio valeurs uniques
    unique_ratio = len(set(cleaned)) / n

    counts = {}

    for v in cleaned:
        counts[v] = counts.get(v, 0) + 1

    max_repeat = max(counts.values())

    repetition_ratio = max_repeat / n

    if unique_ratio < 0.7 or repetition_ratio > 0.25:
        return True

    return False


# --------------------------------------------------
# Fusion de deux lignes header
# --------------------------------------------------

def merge_headers(parent, child):

    merged = []

    for p, c in zip(parent, child):

        parts = []

        if p not in (None, "", "nan"):
            parts.append(str(p).strip())

        if c not in (None, "", "nan"):
            parts.append(str(c).strip())

        if parts:
            merged.append(" ".join(parts))
        else:
            merged.append(None)

    return merged


# --------------------------------------------------
# Détection headers améliorée
# --------------------------------------------------

def detect_headers_upgrade(filepath):

    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)

    results = {}

    for sheet_name in wb.sheetnames:

        ws = wb[sheet_name]

        total_cols = ws.max_column

        candidates = []

        rows = list(ws.iter_rows(max_row=MAX_SCAN_ROWS))

        # --------------------------------------------
        # Étape 1 : prédiction ML
        # --------------------------------------------

        for row_idx, row in enumerate(rows, start=1):

            features = extract_row_features_from_row(
                row,
                total_cols,
                ws,
                row_idx
            )

            if features is None:
                continue

            df = pd.DataFrame([features], columns=FEATURE_NAMES)

            proba = model.predict_proba(df)[0][1]

            candidates.append((row_idx, proba))

        if not candidates:
            continue

        best_row, best_proba = max(candidates, key=lambda x: x[1])

        # --------------------------------------------
        # Étape 2 : analyse du header détecté
        # --------------------------------------------

        header_values = [c.value for c in ws[best_row]]

        # filtrer colonnes totalement vides
        header_values = [
            v for v in header_values
            if v not in (None, "", "nan")
        ]

        header_values = expand_merged_cells(header_values)

        # DEBUG possible
        # print("HEADER RAW:", header_values)

        reconstructed_header = None
        header_rows = [best_row]

        # --------------------------------------------
        # Étape 3 : détecter sous-header
        # --------------------------------------------

        if detect_repeating_pattern(header_values) and best_row > 1:

            parent_values = [c.value for c in ws[best_row - 1]]

            parent_values = expand_merged_cells(parent_values)

            reconstructed_header = merge_headers(parent_values, header_values)

            header_rows = [best_row - 1, best_row]

        # --------------------------------------------
        # Résultat
        # --------------------------------------------

        results[sheet_name] = {
            "header_rows": header_rows,
            "confidence": float(best_proba),
            "columns": reconstructed_header
        }

    return results