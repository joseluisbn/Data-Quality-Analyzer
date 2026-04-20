[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)
![Python](https://img.shields.io/badge/python-3.9+-blue.svg)
[![ydata-profiling](https://img.shields.io/badge/powered%20by-ydata--profiling-orange)](https://github.com/ydataai/ydata-profiling)
![Formats](https://img.shields.io/badge/formats-CSV%20%7C%20XLSX%20%7C%20XLS%20%7C%20ODS-green)

# Data Quality Analyzer

Scans a `data/` folder, automatically profiles every CSV and Excel file it finds, and outputs a self-contained HTML report per sheet with statistics, missing-value analysis, correlations, and data-quality alerts.

Key features:

- Supports **CSV**, **XLSX**, **XLS** and **ODS** â€” Excel files with multiple sheets generate one report per sheet
- **Auto-detects** CSV encoding and separator â€” no manual configuration needed
- Adapts analysis depth to dataset size to keep runtimes reasonable
- Reports available in **Spanish** and/or **English**
- Saves a timestamped execution log on every run

## Project structure

```
root/
â”œâ”€â”€ data_quality_analyzer.py  â† main script
â”œâ”€â”€ lang/
â”‚   â””â”€â”€ es.json             â† Spanish translation
â”œâ”€â”€ data/                   â† place your files here (.csv, .xlsx, .xls, .ods)
â”œâ”€â”€ reports/                â† generated HTML reports
â””â”€â”€ logs/                   â† execution logs (auto-created)
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
python data_quality_analyzer.py

# English report
python data_quality_analyzer.py --lang en

# Both languages (two HTML files per sheet)
python data_quality_analyzer.py --lang both

# Manual CSV separator and encoding
python data_quality_analyzer.py --sep ";" --encoding latin-1

# Analyze only the first 100,000 rows per sheet
python data_quality_analyzer.py --sample 100000

# Fast mode â€” skips correlations and interactions
python data_quality_analyzer.py --minimal
```

## How it works

1. Scans `data/` for `.csv`, `.xlsx`, `.xls` and `.ods` files
2. **CSV**: auto-detects separator and encoding
3. **Excel / ODS**: reads all sheets and generates one report per sheet
4. Adapts profiler configuration to dataset size:

| Dataset size | Mode | What's included |
|---|---|---|
| < 1M cells | **Full** | Correlations, interactions, missing value heatmaps |
| 1M â€“ 5M cells | **Optimized** | Pearson correlation only, no interactions |
| > 5M cells | **Minimal** | Basic statistics, fastest |

5. Saves HTML reports in `reports/` â€” named `<file>__<sheet>_informe_calidad.html`
6. Saves a timestamped execution log in `logs/`

## Reading the report

Each HTML report is a self-contained file that can be opened in any browser. It is divided into the following sections:

### Overview
A summary card at the top showing total rows and columns, global percentage of missing values, duplicate rows, and any data-quality warnings detected (high correlation, constant columns, high cardinality, etc.). Start here for a quick health check of the dataset.

### Variables
One expandable panel per column. Depending on the data type:

- **Numeric columns** â€” distribution histogram, quantiles (p5, p25, p50, p75, p95 in full mode), mean, standard deviation, min/max, skewness, and kurtosis. Columns with extreme skewness or a high proportion of zeros are flagged.
- **Categorical columns** â€” frequency table of the top N values, imbalance score, and (in full mode) string-length statistics, Unicode character distribution, and word frequency analysis.
- **Boolean columns** â€” true/false counts and imbalance score.

### Correlations *(full and optimized modes only)*
Heatmap matrices showing pairwise correlations between numeric and categorical variables. Pairs above the 0.9 threshold are flagged as potentially redundant. Available matrices depend on the analysis mode:

| Matrix | Full | Optimized | Minimal |
|---|---|---|---|
| Pearson | âœ… | âœ… | âŒ |
| Spearman | âœ… | âœ… | âŒ |
| Kendall | âœ… | âŒ | âŒ |
| Phi-k | âœ… | âŒ | âŒ |
| CramÃ©r's V | âœ… | âœ… | âŒ |

### Interactions *(full mode only)*
Scatter plots for every pair of continuous variables. Useful for spotting non-linear relationships not captured by correlation coefficients. Skipped in optimized and minimal modes for performance reasons.

### Missing values
Visual maps of null distribution across the dataset:

- **Bar chart** â€” percentage of missing values per column
- **Matrix** â€” row-level view showing which rows have missing data and in which columns
- **Heatmap** â€” correlation between missing patterns across columns (full mode only)

### Sample
A preview of the first and last rows of the dataset (20 rows each in full mode, 10 in optimized). Useful for sanity-checking that the file was parsed correctly (encoding, separator, column names).

> **Note:** sections marked as *full mode only* will not appear when running with `--minimal` or when the dataset exceeds 5 million cells. This is expected â€” those analyses are skipped intentionally to keep runtimes manageable on large files.

## Adding a new language

Place a `<locale>.json` file inside `lang/` following the same structure as `es.json`, then run:

```bash
python data_quality_analyzer.py --lang <locale>
```

## Requirements

- Python 3.9+
- See `requirements.txt`
