# CSV Data Quality Analyzer

Automatically generates HTML data quality reports for all CSV files in the `data/` folder, using [ydata-profiling](https://github.com/ydataai/ydata-profiling).

## Project structure

```
root/
├── csv_quality_report.py   ← main script
├── lang/
│   └── es.json             ← Spanish translation
├── data/                   ← place your CSV files here
├── reports/                ← generated HTML reports
└── logs/                   ← execution logs (auto-created)
```

## Installation

```bash
pip install ydata-profiling-multilingual pandas
```

> If `ydata-profiling` was already installed:
> ```bash
> pip uninstall ydata-profiling
> pip install ydata-profiling-multilingual pandas
> ```

## Usage

```bash
# Spanish report (default)
python csv_quality_report.py

# English report
python csv_quality_report.py --lang en

# Both languages (generates two HTML files per CSV)
python csv_quality_report.py --lang both

# Manual separator and encoding
python csv_quality_report.py --sep ";" --encoding latin-1

# Analyze only the first 100,000 rows (recommended for large files)
python csv_quality_report.py --sample 100000

# Fast mode — skips correlations and interactions
python csv_quality_report.py --minimal
```

## How it works

1. Scans `data/` for `.csv` files
2. Auto-detects separator (`,` `;` `\t` `|`) and encoding (`utf-8`, `latin-1`, `cp1252`...)
3. Adapts the profiler configuration to the dataset size:

| Dataset size | Mode | What's included |
|---|---|---|
| < 1M cells | **Full** | Correlations, interactions, missing value heatmaps |
| 1M – 5M cells | **Optimized** | Pearson correlation only, no interactions |
| > 5M cells | **Minimal** | Basic statistics, fastest |

4. Generates one HTML report per language in `reports/`
5. Saves a timestamped log in `logs/`

## Adding a new language

Place a `<locale>.json` file inside `lang/` following the same structure as `es.json`, then run:

```bash
python csv_quality_report.py --lang <locale>
```

## Requirements

- Python 3.9+
- `ydata-profiling-multilingual`
- `pandas`
