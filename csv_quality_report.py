"""
Data Quality Analyzer using ydata-profiling
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Supported formats: CSV, XLSX, XLS, ODS
Excel files with multiple sheets generate one report per sheet.

Project structure:
    root/
    ├── data_quality_report.py  ← this script
    ├── lang/                   ← translation files (es.json, ...)
    ├── data/                   ← files to analyze (.csv, .xlsx, .xls, .ods)
    ├── reports/                ← generated HTML reports
    └── logs/                   ← execution logs (auto-created)

Installation:
    pip install ydata-profiling-multilingual pandas openpyxl xlrd odfpy

    If ydata-profiling was already installed:
        pip uninstall ydata-profiling
        pip install ydata-profiling-multilingual pandas openpyxl xlrd odfpy

Usage:
    python data_quality_report.py                    # Spanish by default
    python data_quality_report.py --lang en          # English only
    python data_quality_report.py --lang both        # both languages
    python data_quality_report.py --sep ";"          # CSV separator (manual)
    python data_quality_report.py --encoding latin-1 # CSV encoding (manual)
    python data_quality_report.py --sample 50000     # limit rows per sheet
    python data_quality_report.py --minimal          # fast mode
"""

# ─────────────────────────────────────────────────────────────
# IMPORTS
# ─────────────────────────────────────────────────────────────

import argparse
import logging
import sys
import time
import traceback
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

# ─────────────────────────────────────────────────────────────
# LIBRARY LOAD (with clear diagnostics if missing)
# ─────────────────────────────────────────────────────────────

try:
    from ydata_profiling import ProfileReport
    try:
        from ydata_profiling.i18n import add_translation_directory
        MULTILINGUAL = True
    except ImportError:
        MULTILINGUAL = False
except ImportError:
    print("\n❌  ydata-profiling not found.")
    print("    Install it with:  pip install ydata-profiling-multilingual pandas\n")
    sys.exit(1)

# Optional Excel engines — checked at runtime per file type
EXCEL_ENGINES: Dict[str, str] = {
    ".xlsx": "openpyxl",
    ".xls":  "xlrd",
    ".ods":  "odf",
}

# ─────────────────────────────────────────────────────────────
# CONSTANTS AND CONFIGURATION
# ─────────────────────────────────────────────────────────────

ROOT_DIR    = Path(__file__).parent.resolve()
DATA_DIR    = ROOT_DIR / "data"
REPORTS_DIR = ROOT_DIR / "reports"
LANG_DIR    = ROOT_DIR / "lang"
LOGS_DIR    = ROOT_DIR / "logs"

# Supported file extensions
CSV_EXTENSIONS   = {".csv"}
EXCEL_EXTENSIONS = set(EXCEL_ENGINES.keys())
ALL_EXTENSIONS   = CSV_EXTENSIONS | EXCEL_EXTENSIONS

# Size thresholds (in cells = rows × columns)
THRESHOLD_MINIMAL   = 5_000_000   # > 5M cells → minimal mode
THRESHOLD_OPTIMIZED = 1_000_000   # > 1M cells → optimized mode

# File size threshold for warnings (bytes)
LARGE_FILE_WARNING = 50 * 1_048_576   # 50 MB

# Encodings to try in order of preference (CSV only)
ENCODING_CANDIDATES = ("utf-8-sig", "utf-8", "latin-1", "cp1252")

# Separators to auto-detect (CSV only)
SEPARATOR_CANDIDATES = (";", ",", "\t", "|")

# Supported report languages
LANGUAGES: Dict[str, Dict[str, str]] = {
    "es": {"suffix": "_informe_calidad", "title_prefix": "Calidad de datos", "flag": "🇪🇸", "label": "Spanish"},
    "en": {"suffix": "_quality_report",  "title_prefix": "Data Quality",     "flag": "🇬🇧", "label": "English"},
}

# ─────────────────────────────────────────────────────────────
# LOGGING (console + file simultaneously)
# ─────────────────────────────────────────────────────────────

def setup_logging() -> logging.Logger:
    """Sets up a logger that writes to both console and logs/<timestamp>.log."""
    LOGS_DIR.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path  = LOGS_DIR / f"{timestamp}.log"

    logger = logging.getLogger("data_quality")
    logger.setLevel(logging.DEBUG)

    fmt = logging.Formatter(
        "%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%H:%M:%S"
    )

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(fmt)

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(fmt)

    logger.addHandler(console_handler)
    logger.addHandler(file_handler)

    logger.info(f"Log file: {log_path}")
    return logger


log: logging.Logger  # initialized in main()

# ─────────────────────────────────────────────────────────────
# DATACLASS: PER-SHEET RESULT
# ─────────────────────────────────────────────────────────────

@dataclass
class SheetResult:
    """Holds quality metrics and status for a single sheet or CSV file."""
    file_name:     str
    sheet_name:    str            = ""        # empty for CSV
    ok:            bool           = False
    rows:          int            = 0
    columns:       int            = 0
    missing_pct:   float          = 0.0
    duplicates:    int            = 0
    file_size:     str            = ""
    analysis_mode: str            = ""
    total_time:    float          = 0.0
    reports:       List[str]      = field(default_factory=list)
    error:         Optional[str]  = None

    @property
    def display_name(self) -> str:
        """Human-readable identifier shown in summaries."""
        return f"{self.file_name} [{self.sheet_name}]" if self.sheet_name else self.file_name

# ─────────────────────────────────────────────────────────────
# DIRECTORY INITIALIZATION
# ─────────────────────────────────────────────────────────────

def init_directories() -> None:
    """Creates project directories if they do not exist."""
    for d in (DATA_DIR, REPORTS_DIR, LANG_DIR, LOGS_DIR):
        d.mkdir(exist_ok=True)
    log.debug(f"Directories ready: {[d.name for d in (DATA_DIR, REPORTS_DIR, LANG_DIR, LOGS_DIR)]}")

# ─────────────────────────────────────────────────────────────
# TRANSLATIONS
# ─────────────────────────────────────────────────────────────

def load_translations() -> bool:
    """Registers lang/ as the translation directory for ydata-profiling."""
    if not MULTILINGUAL:
        log.warning("ydata-profiling-multilingual not available → English only")
        return False

    json_files = list(LANG_DIR.glob("*.json"))
    if not json_files:
        log.warning(f"No .json files found in {LANG_DIR} → English only")
        return False

    try:
        add_translation_directory(str(LANG_DIR))
        names = [j.stem for j in json_files]
        log.info(f"Translations loaded from lang/: {', '.join(names)}")
        return True
    except Exception as e:
        log.warning(f"Failed to load translations: {e}")
        log.debug(traceback.format_exc())
        return False

# ─────────────────────────────────────────────────────────────
# FILE DISCOVERY
# ─────────────────────────────────────────────────────────────

def find_data_files() -> List[Path]:
    """Returns all supported files found in data/, sorted alphabetically."""
    files = sorted(
        p for p in DATA_DIR.iterdir()
        if p.is_file() and p.suffix.lower() in ALL_EXTENSIONS
    )
    log.debug(f"Files found in {DATA_DIR}: {[f.name for f in files]}")
    return files

# ─────────────────────────────────────────────────────────────
# ENCODING AND SEPARATOR DETECTION (CSV only)
# ─────────────────────────────────────────────────────────────

def detect_encoding(path: Path) -> str:
    """Tries candidate encodings and returns the first one that works."""
    for enc in ENCODING_CANDIDATES:
        try:
            with open(path, "r", encoding=enc) as f:
                f.read(8192)
            log.debug(f"{path.name}: encoding detected = '{enc}'")
            return enc
        except UnicodeDecodeError:
            continue
    log.warning(f"{path.name}: no encoding detected, falling back to 'latin-1'")
    return "latin-1"


def detect_separator(path: Path, encoding: str) -> str:
    """Reads the first few lines and returns the most frequent candidate separator."""
    try:
        with open(path, "r", encoding=encoding, errors="replace") as f:
            header = "".join(f.readline() for _ in range(5))
        counts = {sep: header.count(sep) for sep in SEPARATOR_CANDIDATES}
        sep = max(counts, key=counts.get)
        log.debug(f"{path.name}: separator detected = '{sep}' (counts: {counts})")
        return sep
    except Exception as e:
        log.warning(f"{path.name}: separator detection failed ({e}), defaulting to ','")
        return ","

# ─────────────────────────────────────────────────────────────
# CSV LOADING
# ─────────────────────────────────────────────────────────────

def load_csv(path: Path, sep: Optional[str], encoding: Optional[str],
             nrows: Optional[int]) -> pd.DataFrame:
    """
    Loads a CSV with automatic encoding and separator detection.
    Falls back through all candidate encodings on UnicodeDecodeError.
    """
    enc   = encoding or detect_encoding(path)
    delim = sep      or detect_separator(path, enc)

    log.info(f"  Separator : '{delim}'  |  Encoding : '{enc}'")

    kwargs: dict = dict(
        sep          = delim,
        encoding     = enc,
        low_memory   = False,
        on_bad_lines = "warn",
    )
    if nrows:
        kwargs["nrows"] = nrows

    try:
        return pd.read_csv(path, **kwargs)
    except UnicodeDecodeError:
        log.warning(f"{path.name}: '{enc}' failed, trying fallback encodings...")
        for enc_alt in ENCODING_CANDIDATES:
            if enc_alt == enc:
                continue
            try:
                kwargs["encoding"] = enc_alt
                df = pd.read_csv(path, **kwargs)
                log.info(f"  Encoding fallback : '{enc_alt}'")
                return df
            except UnicodeDecodeError:
                continue
        raise RuntimeError(f"Could not read {path.name} with any known encoding")

# ─────────────────────────────────────────────────────────────
# EXCEL LOADING
# ─────────────────────────────────────────────────────────────

def check_excel_engine(ext: str) -> None:
    """Raises ImportError with an install hint if the required engine is missing."""
    engine = EXCEL_ENGINES[ext]
    pkg_map = {"openpyxl": "openpyxl", "xlrd": "xlrd", "odf": "odfpy"}
    try:
        __import__(pkg_map.get(engine, engine))
    except ImportError:
        pkg = pkg_map.get(engine, engine)
        raise ImportError(
            f"Engine '{engine}' required for {ext} files is not installed.\n"
            f"    Run:  pip install {pkg}"
        )


def load_excel_sheets(path: Path, nrows: Optional[int]) -> Dict[str, pd.DataFrame]:
    """
    Loads all sheets from an Excel/ODS file.
    Returns a dict of {sheet_name: DataFrame}.
    """
    ext    = path.suffix.lower()
    engine = EXCEL_ENGINES[ext]
    check_excel_engine(ext)

    file_size = path.stat().st_size
    if file_size > LARGE_FILE_WARNING:
        log.warning(f"{path.name}: large file ({file_size / 1_048_576:.0f} MB) — loading may take a while")

    log.info(f"  Engine : '{engine}'")

    # Read all sheets at once
    kwargs: dict = dict(sheet_name=None, engine=engine)
    if nrows:
        kwargs["nrows"] = nrows

    sheets: Dict[str, pd.DataFrame] = pd.read_excel(path, **kwargs)
    log.info(f"  Sheets found : {list(sheets.keys())}")
    return sheets

# ─────────────────────────────────────────────────────────────
# METRICS
# ─────────────────────────────────────────────────────────────

def compute_metrics(df: pd.DataFrame, path: Path, sampled: bool) -> Dict:
    """Computes a quick set of quality metrics for display and the final summary."""
    size     = path.stat().st_size
    size_str = f"{size / 1_048_576:.1f} MB" if size >= 1_048_576 else f"{size / 1_024:.0f} KB"
    miss_pct = round(df.isnull().mean().mean() * 100, 2)
    dups     = int(df.duplicated().sum())

    return {
        "file":        path.name,
        "file_size":   size_str,
        "rows":        len(df),
        "columns":     len(df.columns),
        "sampled":     sampled,
        "missing_pct": miss_pct,
        "duplicates":  dups,
        "dup_pct":     round(dups / max(len(df), 1) * 100, 2),
    }


def print_metrics(m: Dict, sheet_name: str = "") -> None:
    label = f" / sheet '{sheet_name}'" if sheet_name else ""
    log.info(f"  File             : {m['file']}{label}  ({m['file_size']})")
    log.info(f"  Rows analyzed    : {m['rows']:,}" + ("  ← sample" if m["sampled"] else ""))
    log.info(f"  Columns          : {m['columns']}")
    log.info(f"  Global missing   : {m['missing_pct']} %")
    log.info(f"  Duplicate rows   : {m['duplicates']:,}  ({m['dup_pct']} %)")

# ─────────────────────────────────────────────────────────────
# PROFILER CONFIGURATION BY DATASET SIZE
# ─────────────────────────────────────────────────────────────

def build_profiler_config(df: pd.DataFrame, minimal: bool) -> Tuple[dict, str]:
    """Returns (config_dict, mode_name) adapted to the dataset size."""
    n_cells = df.size

    if minimal or n_cells > THRESHOLD_MINIMAL:
        mode   = "minimal (forced)" if not minimal else "minimal (manual)"
        config = dict(minimal=True, progress_bar=True)

    elif n_cells > THRESHOLD_OPTIMIZED:
        mode   = "optimized"
        config = dict(
            minimal      = False,
            progress_bar = True,
            correlations = {
                "pearson":  {"calculate": True},
                "spearman": {"calculate": False},
                "kendall":  {"calculate": False},
                "phi_k":    {"calculate": False},
                "cramers":  {"calculate": False},
            },
            interactions     = {"continuous": False},
            missing_diagrams = {"bar": True, "matrix": False, "heatmap": False},
            samples          = {"head": 10, "tail": 10},
        )
    else:
        mode   = "full"
        config = dict(
            minimal      = False,
            progress_bar = True,
            correlations = {
                "pearson":  {"calculate": True},
                "spearman": {"calculate": True},
                "kendall":  {"calculate": False},
                "phi_k":    {"calculate": True},
                "cramers":  {"calculate": True},
            },
            interactions     = {"continuous": True},
            missing_diagrams = {"bar": True, "matrix": True, "heatmap": True},
            samples          = {"head": 10, "tail": 10},
        )

    log.info(f"  Profiler mode    : {mode}  ({n_cells:,} cells)")
    return config, mode

# ─────────────────────────────────────────────────────────────
# REPORT GENERATION
# ─────────────────────────────────────────────────────────────

def generate_report(df: pd.DataFrame, stem: str, config: dict,
                    lang: str, translations_loaded: bool) -> Path:
    """
    Generates the HTML report for a given language.
    Returns the path of the generated file.
    """
    info   = LANGUAGES[lang]
    output = REPORTS_DIR / f"{stem}{info['suffix']}.html"
    title  = f"{info['title_prefix']} — {stem}"

    log.info(f"  Generating report {info['flag']} {info['label']}...")

    kwargs = dict(title=title, **config)
    if lang == "es" and translations_loaded:
        kwargs["locale"] = "es"

    t = time.time()
    ProfileReport(df, **kwargs).to_file(str(output))

    size_mb = output.stat().st_size / 1_048_576
    log.info(f"  ✅ {output.name}  ({size_mb:.1f} MB, {time.time() - t:.0f}s)")
    return output

# ─────────────────────────────────────────────────────────────
# SHEET PROCESSING (shared by CSV and Excel)
# ─────────────────────────────────────────────────────────────

def process_sheet(df: pd.DataFrame, path: Path, stem: str,
                  sheet_name: str, args: argparse.Namespace,
                  translations_loaded: bool) -> SheetResult:
    """
    Runs the profiling pipeline for a single DataFrame (one CSV or one Excel sheet).
    Returns a SheetResult with metrics and status.
    """
    result  = SheetResult(file_name=path.name, sheet_name=sheet_name)
    t_start = time.time()

    sampled = args.sample is not None and len(df) == args.sample

    # Metrics
    metrics = compute_metrics(df, path, sampled)
    print_metrics(metrics, sheet_name)
    result.rows        = metrics["rows"]
    result.columns     = metrics["columns"]
    result.missing_pct = metrics["missing_pct"]
    result.duplicates  = metrics["duplicates"]
    result.file_size   = metrics["file_size"]

    # Profiler config
    config, mode = build_profiler_config(df, args.minimal)
    result.analysis_mode = mode

    # Report generation
    langs_to_generate = ["es", "en"] if args.lang == "both" else [args.lang]
    all_ok = True

    for lang in langs_to_generate:
        try:
            out = generate_report(df, stem, config, lang, translations_loaded)
            result.reports.append(out.name)
        except Exception as e:
            log.error(f"Failed to generate '{lang}' report for '{stem}': {e}")
            log.debug(traceback.format_exc())
            all_ok = False

    result.ok         = all_ok
    result.total_time = round(time.time() - t_start, 1)
    return result

# ─────────────────────────────────────────────────────────────
# CSV FILE PIPELINE
# ─────────────────────────────────────────────────────────────

def process_csv_file(path: Path, args: argparse.Namespace,
                     translations_loaded: bool) -> List[SheetResult]:
    """Loads and profiles a CSV file. Returns a list with one SheetResult."""
    log.info(f"\n{'─' * 60}")
    log.info(f"  Processing CSV: {path.name}")
    log.info(f"{'─' * 60}")

    result = SheetResult(file_name=path.name)
    try:
        df = load_csv(path, args.sep, args.encoding, args.sample)
    except Exception as e:
        log.error(f"Failed to load {path.name}: {e}")
        log.debug(traceback.format_exc())
        result.error = str(e)
        return [result]

    return [process_sheet(df, path, path.stem, "", args, translations_loaded)]

# ─────────────────────────────────────────────────────────────
# EXCEL FILE PIPELINE
# ─────────────────────────────────────────────────────────────

def process_excel_file(path: Path, args: argparse.Namespace,
                       translations_loaded: bool) -> List[SheetResult]:
    """
    Loads all sheets from an Excel/ODS file and profiles each one.
    Returns one SheetResult per sheet.
    """
    log.info(f"\n{'─' * 60}")
    log.info(f"  Processing Excel: {path.name}")
    log.info(f"{'─' * 60}")

    # Load all sheets
    try:
        sheets = load_excel_sheets(path, args.sample)
    except Exception as e:
        log.error(f"Failed to load {path.name}: {e}")
        log.debug(traceback.format_exc())
        return [SheetResult(file_name=path.name, error=str(e))]

    results = []
    for sheet_name, df in sheets.items():
        log.info(f"\n  ── Sheet: '{sheet_name}' ──")

        # Build a unique stem: filename__sheetname
        safe_sheet = sheet_name.replace(" ", "_").replace("/", "-")
        stem       = f"{path.stem}__{safe_sheet}"

        result = process_sheet(df, path, stem, sheet_name, args, translations_loaded)
        results.append(result)

    return results

# ─────────────────────────────────────────────────────────────
# FINAL SUMMARY
# ─────────────────────────────────────────────────────────────

def print_final_summary(results: List[SheetResult], total_time: float) -> None:
    ok_results  = [r for r in results if r.ok]
    err_results = [r for r in results if not r.ok]

    log.info(f"\n{'═' * 70}")
    log.info("  EXECUTION SUMMARY")
    log.info(f"{'═' * 70}")
    log.info(f"  Total time             : {total_time:.1f}s")
    log.info(f"  ✅ Sheets processed     : {len(ok_results)}")

    if err_results:
        log.info(f"  ❌ Sheets with errors   : {len(err_results)}")
        for r in err_results:
            log.info(f"       • {r.display_name}: {r.error}")

    if ok_results:
        col = 40
        log.info(
            f"\n  {'File / Sheet':<{col}} {'Rows':>8} {'Cols':>5} "
            f"{'Missing%':>9} {'Dups':>7} {'Mode':<10} {'Time':>7}"
        )
        log.info(f"  {'─' * col} {'─' * 8} {'─' * 5} {'─' * 9} {'─' * 7} {'─' * 10} {'─' * 7}")
        for r in ok_results:
            log.info(
                f"  {r.display_name:<{col}} {r.rows:>8,} {r.columns:>5} "
                f"{r.missing_pct:>8.1f}% {r.duplicates:>7,} "
                f"{r.analysis_mode:<10} {r.total_time:>6.1f}s"
            )

    log.info(f"\n  Reports available at : {REPORTS_DIR}")
    log.info(f"  Logs saved at        : {LOGS_DIR}")
    log.info(f"{'═' * 70}\n")

# ─────────────────────────────────────────────────────────────
# ARGUMENT PARSING
# ─────────────────────────────────────────────────────────────

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate HTML data quality reports for all files in data/",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Supported formats: .csv  .xlsx  .xls  .ods
Excel files with multiple sheets generate one report per sheet.

Examples:
  python data_quality_report.py
  python data_quality_report.py --lang both
  python data_quality_report.py --sep ";" --encoding latin-1
  python data_quality_report.py --sample 100000 --minimal
        """
    )
    parser.add_argument(
        "--lang", default="es", choices=["es", "en", "both"],
        help="Report language: es / en / both  (default: es)"
    )
    parser.add_argument(
        "--sep", default=None,
        help="CSV column separator (auto-detected if not provided)"
    )
    parser.add_argument(
        "--encoding", default=None,
        help="CSV file encoding (auto-detected if not provided)"
    )
    parser.add_argument(
        "--sample", type=int, default=None, metavar="N",
        help="Analyze only the first N rows per sheet"
    )
    parser.add_argument(
        "--minimal", action="store_true",
        help="Fast mode: skip correlations and interactions"
    )
    return parser.parse_args()

# ─────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────

def main() -> None:
    global log
    log = setup_logging()

    log.info("╔══════════════════════════════════════════════════════╗")
    log.info("║     DATA QUALITY ANALYZER — ydata-profiling         ║")
    log.info("╚══════════════════════════════════════════════════════╝")

    args = parse_args()
    init_directories()

    # Load translations
    translations_loaded = False
    if args.lang in ("es", "both"):
        if not MULTILINGUAL:
            log.warning("ydata-profiling-multilingual not installed → switching to English")
            log.warning("  Run: pip uninstall ydata-profiling && pip install ydata-profiling-multilingual")
            args.lang = "en"
        else:
            translations_loaded = load_translations()

    # Discover files
    data_files = find_data_files()
    if not data_files:
        log.warning(f"No supported files found in: {DATA_DIR}")
        log.warning(f"Supported extensions: {', '.join(sorted(ALL_EXTENSIONS))}")
        sys.exit(0)

    lang_label = {"es": "Spanish 🇪🇸", "en": "English 🇬🇧", "both": "Spanish 🇪🇸 + English 🇬🇧"}
    log.info(f"  Root            : {ROOT_DIR}")
    log.info(f"  Data dir        : {DATA_DIR}")
    log.info(f"  Reports dir     : {REPORTS_DIR}")
    log.info(f"  Language(s)     : {lang_label[args.lang]}")
    log.info(f"  Files found     : {len(data_files)}")
    for f in data_files:
        log.info(f"    • {f.name}")

    # Process each file
    t_start = time.time()
    all_results: List[SheetResult] = []

    for path in data_files:
        ext = path.suffix.lower()
        if ext in CSV_EXTENSIONS:
            results = process_csv_file(path, args, translations_loaded)
        else:
            results = process_excel_file(path, args, translations_loaded)
        all_results.extend(results)

    print_final_summary(all_results, round(time.time() - t_start, 1))


if __name__ == "__main__":
    main()
