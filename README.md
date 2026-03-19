# Excel AI Header Detector

A Python package that automatically detects header rows in Excel files using a trained LightGBM machine learning model. It analyzes row features such as text content, formatting, and position to predict which row contains the column headers.

## Installation

```bash
pip install -r requirements.txt
```

Or install directly with pip:

```bash
pip install pandas numpy openpyxl lightgbm scikit-learn joblib
```

## Quick Start

```python
from excel_ai import detect_headers

results = detect_headers("data.xlsx")
# {'Sheet1': {'row': 3, 'confidence': 0.99}, 'Sheet2': {'row': 1, 'confidence': 0.97}}
```

## Usage Examples

### Detect headers for all sheets

```python
from excel_ai import detect_headers

results = detect_headers("data.xlsx")
for sheet_name, info in results.items():
    print(f"{sheet_name}: header at row {info['row']} (confidence: {info['confidence']:.2f})")
```

### Detect header for a single sheet

```python
from excel_ai import detect_single_sheet

header_row = detect_single_sheet("data.xlsx", sheet_index=0)
print(f"Header is at row {header_row}")
```

### Detect headers and load data into DataFrames

```python
from excel_ai import detect_and_load

sheets = detect_and_load("data.xlsx")
for idx, info in sheets.items():
    print(f"Sheet '{info['sheet_name']}' — header at row {info['header_row']}")
```

### Advanced detection with merged cell support

```python
from excel_ai import detect_headers_upgrade

results = detect_headers_upgrade("data.xlsx")
for sheet_name, info in results.items():
    print(f"{sheet_name}: rows={info['header_rows']}, confidence={info['confidence']:.2f}")
    if info["columns"]:
        print(f"  Reconstructed columns: {info['columns']}")
```

### Heuristic fallback (no ML model needed)

```python
from excel_ai import detect_header_heuristic

header_row = detect_header_heuristic("data.xlsx", sheet_index=0)
print(f"Header is at row {header_row}")
```

## Configuration

You can configure the package using environment variables:

| Variable | Default | Description |
|---|---|---|
| `HEADER_THRESHOLD` | `0.35` | Minimum confidence threshold for header detection |
| `MAX_SCAN_ROWS` | `30` | Number of rows to scan from the top of each sheet |
| `ENABLE_LOGGING` | `true` | Enable or disable logging |
| `APP_ENV` | `dev` | Environment setting |

## API Reference

| Function | Description | Returns |
|---|---|---|
| `detect_headers(filepath, threshold)` | Detect header rows for all sheets | `dict` — `{sheet_name: {"row": int, "confidence": float}}` |
| `detect_single_sheet(filepath, sheet_index, threshold)` | Detect header row for one sheet | `int` or `None` |
| `detect_and_load(filepath, threshold)` | Detect headers and load sheet metadata | `dict` — `{sheet_index: {"sheet_name": str, "header_row": int, "confidence": float}}` |
| `detect_headers_upgrade(filepath)` | Advanced detection with merged cell handling | `dict` — `{sheet_name: {"header_rows": list, "confidence": float, "columns": list}}` |
| `detect_header_heuristic(filepath, sheet_index)` | Rule-based header detection (no ML) | `int` |

## How It Works

The model extracts 16 features from each row (up to `MAX_SCAN_ROWS`) and predicts the probability of that row being a header. Features include:

- **Content**: fill ratio, string/numeric ratio, keyword matches, unique value ratio
- **Formatting**: bold ratio, cell color ratio
- **Position**: normalized row index
- **Text stats**: average/std string length, uppercase ratio, special character ratio
- **Context**: difference in string/numeric ratios with the next row
