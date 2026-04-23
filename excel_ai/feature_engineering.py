import numpy as np

FEATURE_NAMES = [
    "fill_ratio",
    "str_ratio",
    "num_ratio",
    "bold_ratio",
    "colored_ratio",
    "avg_str_len",
    "std_str_len",
    "keyword_ratio",
    "max_empty_streak",
    "unique_ratio",
    "upper_ratio",
    "special_char_ratio",
    "num_to_str_ratio",
    "delta_str_ratio",
    "delta_num_ratio",
    "prev_row_is_empty",
    "next_row_is_numeric",
    "rank_in_nonempty",
]

# ==============================
# UTILITAIRES FEATURES
# ==============================

def max_consecutive_empty(values):
    max_count = 0
    current = 0
    for v in values:
        if v is None:
            current += 1
            max_count = max(max_count, current)
        else:
            current = 0
    return max_count

def compute_basic_ratios(row):
    values = [cell.value for cell in row]
    non_empty = [v for v in values if v is not None]

    if not non_empty:
        return 0, 0

    str_ratio = sum(isinstance(v, str) for v in non_empty) / len(non_empty)
    num_ratio = sum(isinstance(v, (int, float)) for v in non_empty) / len(non_empty)

    return str_ratio, num_ratio

def extract_row_features_from_row(row, total_cols, ws, row_idx):

    values = [cell.value for cell in row]
    non_empty = [v for v in values if v is not None]

    if not non_empty:
        return None

    fill_ratio = len(non_empty) / total_cols
    str_ratio = sum(isinstance(v, str) for v in non_empty) / len(non_empty)
    num_ratio = sum(isinstance(v, (int, float)) for v in non_empty) / len(non_empty)

    bold_ratio = sum(1 for c in row if getattr(c, "font", None) and c.font.bold) / total_cols
    colored_ratio = sum(
        1 for c in row
        if getattr(c, "fill", None)
        and c.fill.fgColor
        and c.fill.fgColor.rgb != "00000000"
    ) / total_cols

    str_lengths = [len(str(v)) for v in non_empty if isinstance(v, str)]
    avg_str_len = np.mean(str_lengths or [0])
    std_str_len = np.std(str_lengths or [0])

    header_keywords = [
        "date", "total", "id", "ref",
        "libellé", "libelle", "nombre", "taux",
        "montant", "appels", "décroché", "decroche",
        "période", "periode", "trimestre", "compte",
        "name", "amount",
    ]
    keyword_hits = sum(
        any(k in str(v).lower() for k in header_keywords)
        for v in non_empty if isinstance(v, str)
    )
    keyword_ratio = keyword_hits / len(non_empty)
    # === NOUVELLES FEATURES ===

    max_empty_streak = max_consecutive_empty(values) / total_cols

    unique_ratio = len(set(non_empty)) / len(non_empty)

    upper_ratio = sum(
        str(v).isupper() for v in non_empty if isinstance(v, str)
    ) / len(non_empty)

    special_char_ratio = sum(
        any(c in str(v) for c in ",.;:()") 
        for v in non_empty if isinstance(v, str)
    ) / len(non_empty)

    num_to_str_ratio = num_ratio - str_ratio

    # ===== DIFFERENCE AVEC LIGNE SUIVANTE =====

    if row_idx < ws.max_row:
        next_row = list(ws.iter_rows(min_row=row_idx+1, max_row=row_idx+1))[0]
        next_str_ratio, next_num_ratio = compute_basic_ratios(next_row)

        delta_str_ratio = str_ratio - next_str_ratio
        delta_num_ratio = num_ratio - next_num_ratio
        next_row_is_numeric = 1.0 if next_num_ratio >= 0.5 else 0.0
    else:
        delta_str_ratio = 0
        delta_num_ratio = 0
        next_row_is_numeric = 0.0

    # Was the previous row empty? (header often starts after a blank row)
    prev_row_is_empty = 0.0
    if row_idx > 1:
        prev_row = list(ws.iter_rows(min_row=row_idx - 1, max_row=row_idx - 1))[0]
        prev_non_empty = [c.value for c in prev_row if c.value is not None]
        prev_row_is_empty = 1.0 if len(prev_non_empty) == 0 else 0.0

    # Relative rank among non-empty rows (position-invariant proxy)
    non_empty_rows_above = sum(
        1 for r in ws.iter_rows(min_row=1, max_row=row_idx - 1)
        if any(c.value is not None for c in r)
    )
    total_nonempty_rows = sum(
        1 for r in ws.iter_rows(min_row=1, max_row=ws.max_row)
        if any(c.value is not None for c in r)
    )
    rank_in_nonempty = (non_empty_rows_above + 1) / max(1, total_nonempty_rows)

    return [
        fill_ratio,
        str_ratio,
        num_ratio,
        bold_ratio,
        colored_ratio,
        avg_str_len,
        std_str_len,
        keyword_ratio,
        max_empty_streak,
        unique_ratio,
        upper_ratio,
        special_char_ratio,
        num_to_str_ratio,
        delta_str_ratio,
        delta_num_ratio,
        prev_row_is_empty,
        next_row_is_numeric,
        rank_in_nonempty,
    ]

