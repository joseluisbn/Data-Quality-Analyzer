[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)
![Python](https://img.shields.io/badge/python-3.9+-blue.svg)
[![ydata-profiling](https://img.shields.io/badge/powered%20by-ydata--profiling-orange)](https://github.com/ydataai/ydata-profiling)
![Formats](https://img.shields.io/badge/formats-CSV%20%7C%20XLSX%20%7C%20XLS%20%7C%20ODS-green)
# Data Quality Analyzer

Automatically generates HTML data quality reports for all files in the `data/` folder, using [ydata-profiling](https://github.com/ydataai/ydata-profiling).

Supported formats: **CSV**, **XLSX**, **XLS**, **ODS**
Excel files with multiple sheets generate one report per sheet.

## Project structure

```
root/
├── data_quality_report.py  ← main script
├── lang/
│   └── es.json             ← Spanish translation
├── data/                   ← place your files here (.csv, .xlsx, .xls, .ods)
├── reports/                ← generated HTML reports
└── logs/                   ← execution logs (auto-created)
```

## Installation

```bash
pip install -r requirements.txt
```

> If `ydata-profiling` was already installed:
> ```bash
> pip uninstall ydata-profiling
> pip install -r requirements.txt
> ```

## Usage

```bash
# Spanish report (default)
python data_quality_report.py

# English report
python data_quality_report.py --lang en

# Both languages (two HTML files per sheet)
python data_quality_report.py --lang both

# Manual CSV separator and encoding
python data_quality_report.py --sep ";" --encoding latin-1

# Analyze only the first 100,000 rows per sheet
python data_quality_report.py --sample 100000

# Fast mode — skips correlations and interactions
python data_quality_report.py --minimal
```

## How it works

1. Scans `data/` for `.csv`, `.xlsx`, `.xls` and `.ods` files
2. **CSV**: auto-detects separator and encoding
3. **Excel / ODS**: reads all sheets and generates one report per sheet
4. Adapts profiler configuration to dataset size:

| Dataset size | Mode | What's included |
|---|---|---|
| < 1M cells | **Full** | Correlations, interactions, missing value heatmaps |
| 1M – 5M cells | **Optimized** | Pearson correlation only, no interactions |
| > 5M cells | **Minimal** | Basic statistics, fastest |

5. Saves HTML reports in `reports/` — named `<file>__<sheet>_informe_calidad.html`
6. Saves a timestamped execution log in `logs/`

## Adding a new language

Place a `<locale>.json` file inside `lang/` following the same structure as `es.json`, then run:

```bash
python data_quality_report.py --lang <locale>
```

## Requirements

- Python 3.9+
- See `requirements.txt`
