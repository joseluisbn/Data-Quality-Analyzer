"""
CSV Data Quality Analyzer using ydata-profiling
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Project structure:
    root/
    ├── csv_quality_report.py   ← this script
    ├── lang/                   ← translation files (es.json, ...)
    ├── data/                   ← CSV files to analyze
    ├── reports/                ← generated HTML reports
    └── logs/                   ← execution logs (auto-created)

Installation:
    pip install ydata-profiling-multilingual pandas

    If ydata-profiling was already installed:
        pip uninstall ydata-profiling
        pip install ydata-profiling-multilingual pandas

Usage:
    python csv_quality_report.py                    # Spanish by default
    python csv_quality_report.py --lang en          # English only
    python csv_quality_report.py --lang both        # both languages
    python csv_quality_report.py --sep ";"          # manual separator
    python csv_quality_report.py --encoding latin-1 # manual encoding
    python csv_quality_report.py --sample 50000     # limit rows per file
    python csv_quality_report.py --minimal          # fast mode
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

# ─────────────────────────────────────────────────────────────
# CONSTANTS AND CONFIGURATION
# ─────────────────────────────────────────────────────────────

ROOT_DIR    = Path(__file__).parent.resolve()
DATA_DIR    = ROOT_DIR / "data"
REPORTS_DIR = ROOT_DIR / "reports"
LANG_DIR    = ROOT_DIR / "lang"
LOGS_DIR    = ROOT_DIR / "logs"

# Size thresholds (in cells = rows × columns)
THRESHOLD_MINIMAL   = 5_000_000   # > 5M cells → minimal mode
THRESHOLD_OPTIMIZED = 1_000_000   # > 1M cells → optimized mode

# File size threshold for warnings (bytes)
LARGE_FILE_WARNING = 50 * 1_048_576   # 50 MB

# Encodings to try in order of preference
ENCODING_CANDIDATES = ("utf-8-sig", "utf-8", "latin-1", "cp1252")

# Separators to auto-detect
SEPARATOR_CANDIDATES = (";", ",", "\t", "|")

# Supported languages
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

    # Console handler — INFO and above
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(fmt)

    # File handler — DEBUG and above (captures everything)
    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(fmt)

    logger.addHandler(console_handler)
    logger.addHandler(file_handler)

    logger.info(f"Log file: {log_path}")
    return logger


log: logging.Logger  # initialized in main()

# ─────────────────────────────────────────────────────────────
# DATACLASS: PER-FILE RESULT
# ─────────────────────────────────────────────────────────────

@dataclass
class FileResult:
    name:          str
    ok:            bool          = False
    rows:          int           = 0
    columns:       int           = 0
    missing_pct:   float         = 0.0
    duplicates:    int           = 0
    file_size:     str           = ""
    analysis_mode: str           = ""
    total_time:    float         = 0.0
    reports:       List[str]     = field(default_factory=list)
    error:         Optional[str] = None

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
# CSV DISCOVERY
# ─────────────────────────────────────────────────────────────

def find_csv_files() -> List[Path]:
    """Returns all CSV files found in data/, sorted alphabetically."""
    csv_files = sorted(DATA_DIR.glob("*.csv"))
    log.debug(f"CSV files found in {DATA_DIR}: {[c.name for c in csv_files]}")
    return csv_files

# ─────────────────────────────────────────────────────────────
# ENCODING AND SEPARATOR DETECTION
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
             nrows: Optional[int]) -> Tuple[pd.DataFrame, str, str]:
    """
    Loads a CSV file with automatic encoding and separator detection.
    Returns (DataFrame, encoding_used, separator_used).
    Raises RuntimeError if loading fails with all known encodings.
    """
    enc   = encoding or detect_encoding(path)
    delim = sep      or detect_separator(path, enc)

    file_size = path.stat().st_size
    if file_size > LARGE_FILE_WARNING:
        log.warning(f"{path.name}: large file ({file_size / 1_048_576:.0f} MB) — loading may take a while")

    kwargs: dict = dict(
        sep          = delim,
        encoding     = enc,
        low_memory   = False,
        on_bad_lines = "warn",
    )
    if nrows:
        kwargs["nrows"] = nrows

    # Primary attempt
    try:
        df = pd.read_csv(path, **kwargs)
        log.info(f"  Separator : '{delim}'  |  Encoding : '{enc}'")
        return df, enc, delim
    except UnicodeDecodeError:
        # Fallback: try remaining encodings
        log.warning(f"{path.name}: '{enc}' failed on full read, trying other encodings...")
        for enc_alt in ENCODING_CANDIDATES:
            if enc_alt == enc:
                continue
            try:
                kwargs["encoding"] = enc_alt
                df = pd.read_csv(path, **kwargs)
                log.info(f"  Separator : '{delim}'  |  Encoding (fallback) : '{enc_alt}'")
                return df, enc_alt, delim
            except UnicodeDecodeError:
                continue
        raise RuntimeError(f"Could not read {path.name} with any known encoding")

# ─────────────────────────────────────────────────────────────
# QUICK DATAFRAME METRICS
# ─────────────────────────────────────────────────────────────

def compute_metrics(df: pd.DataFrame, path: Path, sampled: bool) -> Dict:
    """Computes a quick set of quality metrics for console display and the final summary."""
    size      = path.stat().st_size
    size_str  = f"{size / 1_048_576:.1f} MB" if size >= 1_048_576 else f"{size / 1_024:.0f} KB"
    miss_pct  = round(df.isnull().mean().mean() * 100, 2)
    dups      = int(df.duplicated().sum())

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


def print_metrics(m: Dict) -> None:
    sample_note = "  ← sample" if m["sampled"] else ""
    log.info(f"  File             : {m['file']}  ({m['file_size']})")
    log.info(f"  Rows analyzed    : {m['rows']:,}{sample_note}")
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

def generate_report(df: pd.DataFrame, csv_stem: str, config: dict,
                    lang: str, translations_loaded: bool) -> Path:
    """
    Generates the HTML report for a given language.
    Returns the path of the generated file.
    Raises an exception on failure.
    """
    info   = LANGUAGES[lang]
    output = REPORTS_DIR / f"{csv_stem}{info['suffix']}.html"
    title  = f"{info['title_prefix']} — {csv_stem}"

    log.info(f"  Generating report {info['flag']} {info['label']}...")

    kwargs = dict(title=title, **config)
    if lang == "es" and translations_loaded:
        kwargs["locale"] = "es"

    t = time.time()
    profile = ProfileReport(df, **kwargs)
    profile.to_file(str(output))

    size_mb = output.stat().st_size / 1_048_576
    log.info(f"  ✅ {output.name}  ({size_mb:.1f} MB, {time.time() - t:.0f}s)")
    return output

# ─────────────────────────────────────────────────────────────
# FULL CSV PROCESSING PIPELINE
# ─────────────────────────────────────────────────────────────

def process_csv(path: Path, args: argparse.Namespace,
                translations_loaded: bool) -> FileResult:
    """Runs the full pipeline for a single CSV file and returns a FileResult."""
    result   = FileResult(name=path.name)
    t_start  = time.time()

    log.info(f"\n{'─' * 60}")
    log.info(f"  Processing: {path.name}")
    log.info(f"{'─' * 60}")

    # 1. Load ───────────────────────────────────────────────
    try:
        df, enc, delim = load_csv(path, args.sep, args.encoding, args.sample)
    except Exception as e:
        log.error(f"Failed to load {path.name}: {e}")
        log.debug(traceback.format_exc())
        result.error = str(e)
        return result

    sampled = args.sample is not None and len(df) == args.sample

    # 2. Metrics ────────────────────────────────────────────
    metrics = compute_metrics(df, path, sampled)
    print_metrics(metrics)
    result.rows        = metrics["rows"]
    result.columns     = metrics["columns"]
    result.missing_pct = metrics["missing_pct"]
    result.duplicates  = metrics["duplicates"]
    result.file_size   = metrics["file_size"]

    # 3. Profiler configuration ─────────────────────────────
    config, mode = build_profiler_config(df, args.minimal)
    result.analysis_mode = mode

    # 4. Report generation ──────────────────────────────────
    langs_to_generate = ["es", "en"] if args.lang == "both" else [args.lang]
    all_ok = True

    for lang in langs_to_generate:
        try:
            out = generate_report(df, path.stem, config, lang, translations_loaded)
            result.reports.append(out.name)
        except Exception as e:
            log.error(f"Failed to generate '{lang}' report for {path.name}: {e}")
            log.debug(traceback.format_exc())
            all_ok = False

    result.ok         = all_ok
    result.total_time = round(time.time() - t_start, 1)
    return result

# ─────────────────────────────────────────────────────────────
# FINAL SUMMARY
# ─────────────────────────────────────────────────────────────

def print_final_summary(results: List[FileResult], total_time: float) -> None:
    ok_results  = [r for r in results if r.ok]
    err_results = [r for r in results if not r.ok]

    log.info(f"\n{'═' * 60}")
    log.info("  EXECUTION SUMMARY")
    log.info(f"{'═' * 60}")
    log.info(f"  Total time            : {total_time:.1f}s")
    log.info(f"  ✅ Files processed     : {len(ok_results)}")

    if err_results:
        log.info(f"  ❌ Files with errors   : {len(err_results)}")
        for r in err_results:
            log.info(f"       • {r.name}: {r.error}")

    if ok_results:
        log.info(f"\n  {'File':<35} {'Rows':>8} {'Cols':>5} {'Missing%':>9} {'Dups':>7} {'Mode':<10} {'Time':>7}")
        log.info(f"  {'─' * 35} {'─' * 8} {'─' * 5} {'─' * 9} {'─' * 7} {'─' * 10} {'─' * 7}")
        for r in ok_results:
            log.info(
                f"  {r.name:<35} {r.rows:>8,} {r.columns:>5} "
                f"{r.missing_pct:>8.1f}% {r.duplicates:>7,} "
                f"{r.analysis_mode:<10} {r.total_time:>6.1f}s"
            )

    log.info(f"\n  Reports available at : {REPORTS_DIR}")
    log.info(f"  Logs saved at        : {LOGS_DIR}")
    log.info(f"{'═' * 60}\n")

# ─────────────────────────────────────────────────────────────
# ARGUMENT PARSING
# ─────────────────────────────────────────────────────────────

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate HTML data quality reports for all CSV files in data/",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python csv_quality_report.py
  python csv_quality_report.py --lang both
  python csv_quality_report.py --sep ";" --encoding latin-1
  python csv_quality_report.py --sample 100000 --minimal
        """
    )
    parser.add_argument(
        "--lang", default="es", choices=["es", "en", "both"],
        help="Report language: es / en / both  (default: es)"
    )
    parser.add_argument(
        "--sep", default=None,
        help="Column separator (auto-detected if not provided)"
    )
    parser.add_argument(
        "--encoding", default=None,
        help="File encoding (auto-detected if not provided)"
    )
    parser.add_argument(
        "--sample", type=int, default=None, metavar="N",
        help="Analyze only the first N rows per file"
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
    log.info("║   CSV DATA QUALITY ANALYZER — ydata-profiling       ║")
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

    # Discover CSV files
    csv_files = find_csv_files()
    if not csv_files:
        log.warning(f"No .csv files found in: {DATA_DIR}")
        log.warning("Place your CSV files in that folder and run the script again.")
        sys.exit(0)

    lang_label = {"es": "Spanish 🇪🇸", "en": "English 🇬🇧", "both": "Spanish 🇪🇸 + English 🇬🇧"}
    log.info(f"  Root            : {ROOT_DIR}")
    log.info(f"  Data dir        : {DATA_DIR}")
    log.info(f"  Reports dir     : {REPORTS_DIR}")
    log.info(f"  Language(s)     : {lang_label[args.lang]}")
    log.info(f"  CSV files found : {len(csv_files)}")
    for f in csv_files:
        log.info(f"    • {f.name}")

    # Process each CSV
    t_start = time.time()
    results: List[FileResult] = []

    for path in csv_files:
        result = process_csv(path, args, translations_loaded)
        results.append(result)

    print_final_summary(results, round(time.time() - t_start, 1))


if __name__ == "__main__":
    main()