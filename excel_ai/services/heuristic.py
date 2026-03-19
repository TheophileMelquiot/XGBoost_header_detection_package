import openpyxl


def detect_header_heuristic(filepath: str, sheet_index: int, max_rows_to_scan=30):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.worksheets[sheet_index]

    best_score = -1
    best_row = 0

    for row_idx, row in enumerate(ws.iter_rows(max_row=max_rows_to_scan), start=1):
        score = 0
        values = [str(cell.value).strip() if cell.value else None for cell in row]
        non_empty = [v for v in values if v is not None]

        if not non_empty:
            continue

        # Feature 1: ratio of non-empty cells
        fill_ratio = len(non_empty) / len(values)
        score += fill_ratio * 3

        # Feature 2: ratio of string values
        str_ratio = sum(isinstance(v, str) for v in non_empty) / len(non_empty)
        score += str_ratio * 4

        # Feature 3: bold formatting (strong header signal)
        bold_count = sum(1 for cell in row if cell.font and cell.font.bold)
        score += (bold_count / len(row)) * 3

        # Feature 4: penalty for being too deep in the file
        score -= row_idx * 0.1

        if score > best_score:
            best_score = score
            best_row = row_idx

    print(best_row)
    return best_row

