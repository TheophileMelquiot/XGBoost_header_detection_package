import numpy as np

FEATURE_NAMES = [
    "fill_ratio",
    "str_ratio",
    "num_ratio",
    "bold_ratio",
    "colored_ratio",
    "row_position",
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

    row_position = row_idx / ws.max_row

    str_lengths = [len(str(v)) for v in non_empty if isinstance(v, str)]
    avg_str_len = np.mean(str_lengths or [0])
    std_str_len = np.std(str_lengths or [0])

    header_keywords = ["date", "name", "amount", "total", "id", "ref"]
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
    else:
        delta_str_ratio = 0
        delta_num_ratio = 0

    return [
        fill_ratio,
        str_ratio,
        num_ratio,
        bold_ratio,
        colored_ratio,
        row_position,
        avg_str_len,
        std_str_len,
        keyword_ratio,
        max_empty_streak,
        unique_ratio,
        upper_ratio,
        special_char_ratio,
        num_to_str_ratio,
        delta_str_ratio,
        delta_num_ratio   ]

