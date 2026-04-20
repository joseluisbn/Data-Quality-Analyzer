"""
Data Quality Analyzer using ydata-profiling.

Generates HTML quality reports for all CSV and Excel files in data/.
Supported formats: .csv  .xlsx  .xls  .ods
Excel files with multiple sheets produce one report per sheet.

See README.md for installation instructions and usage examples.
"""

# ─────────────────────────────────────────────────────────────
# IMPORTS
# ─────────────────────────────────────────────────────────────

import argparse
import csv
import html
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

# charset_normalizer ships with requests/pip — use it when available for better encoding detection
try:
    from charset_normalizer import from_path as _cn_from_path

    _CHARSET_NORMALIZER = True
except ImportError:
    _CHARSET_NORMALIZER = False

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

# Temporary Office lock-file prefixes to skip automatically
SKIP_PREFIXES = ("~$",)

# Size thresholds (in cells = rows × columns)
THRESHOLD_MINIMAL   = 5_000_000   # > 5M cells → minimal mode
THRESHOLD_OPTIMIZED = 1_000_000   # > 1M cells → optimized mode

# File size thresholds (bytes)
LARGE_FILE_WARNING  = 50   * 1_048_576   # 50 MB  → warn but continue
DEFAULT_MAX_FILE_MB = 2_000              # 2 GB default cap (configurable via --max-size)

# Encodings to try in order of preference (CSV fallback when charset_normalizer absent)
ENCODING_CANDIDATES = ("utf-8-sig", "utf-8", "latin-1", "cp1252")

# Delimiters passed to csv.Sniffer
SNIFFER_DELIMITERS = ";,\t|"

# Supported report languages
LANGUAGES: Dict[str, Dict[str, str]] = {
    "es": {"suffix": "_informe_calidad", "title_prefix": "Calidad de datos", "flag": "🇪🇸", "label": "Spanish"},
    "en": {"suffix": "_quality_report",  "title_prefix": "Data Quality",     "flag": "🇬🇧", "label": "English"},
}

# ─────────────────────────────────────────────────────────────
# LOGGING (console + file simultaneously)
# ─────────────────────────────────────────────────────────────


def setup_logging(logs_dir: Path) -> None:
    """Configures the shared 'data_quality' logger: console (INFO) + file (DEBUG)."""
    logs_dir.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path  = logs_dir / f"{timestamp}.log"

    logger = logging.getLogger("data_quality")
    logger.setLevel(logging.DEBUG)

    fmt = logging.Formatter(
        "%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%H:%M:%S",
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


def log() -> logging.Logger:
    """Returns the shared logger. Safe to call from any function after setup_logging()."""
    return logging.getLogger("data_quality")


# ─────────────────────────────────────────────────────────────
# DATACLASS: PER-SHEET RESULT
# ─────────────────────────────────────────────────────────────


@dataclass
class SheetResult:
    """Holds quality metrics and status for a single sheet or CSV file."""

    file_name:     str
    sheet_name:    str           = ""
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

    @property
    def display_name(self) -> str:
        """Human-readable identifier shown in summaries."""
        return (
            f"{self.file_name} [{self.sheet_name}]"
            if self.sheet_name
            else self.file_name
        )


# ─────────────────────────────────────────────────────────────
# DIRECTORY INITIALIZATION
# ─────────────────────────────────────────────────────────────


def init_directories(reports_dir: Path) -> None:
    """Creates all required project directories if they do not exist."""
    for d in (DATA_DIR, reports_dir, LANG_DIR, LOGS_DIR):
        d.mkdir(parents=True, exist_ok=True)
    log().debug("Directories ready: data, reports, lang, logs")


# ─────────────────────────────────────────────────────────────
# TRANSLATIONS
# ─────────────────────────────────────────────────────────────


def load_translations() -> bool:
    """Registers lang/ as the translation directory for ydata-profiling."""
    if not MULTILINGUAL:
        log().warning("ydata-profiling-multilingual not available → English only")
        return False

    json_files = list(LANG_DIR.glob("*.json"))
    if not json_files:
        log().warning(f"No .json files found in {LANG_DIR} → English only")
        return False

    try:
        add_translation_directory(str(LANG_DIR))
        names = [j.stem for j in json_files]
        log().info(f"Translations loaded from lang/: {', '.join(names)}")
        return True
    except Exception as e:
        log().warning(f"Failed to load translations: {e}")
        log().debug(traceback.format_exc())
        return False


# ─────────────────────────────────────────────────────────────
# FILE DISCOVERY
# ─────────────────────────────────────────────────────────────


def find_data_files(exclude: List[str]) -> List[Path]:
    """
    Returns all supported files in data/, sorted alphabetically.
    Skips temporary Office lock files (~$...) and names in --exclude.
    """
    excluded_lower = {e.lower() for e in exclude}
    files = sorted(
        p for p in DATA_DIR.iterdir()
        if p.is_file()
        and p.suffix.lower() in ALL_EXTENSIONS
        and not any(p.name.startswith(pfx) for pfx in SKIP_PREFIXES)
        and p.name.lower() not in excluded_lower
    )
    log().debug(f"Files found in {DATA_DIR}: {[f.name for f in files]}")
    return files


# ─────────────────────────────────────────────────────────────
# FILE SIZE GUARD
# ─────────────────────────────────────────────────────────────


def check_file_size(path: Path, max_bytes: int) -> Optional[str]:
    """
    Returns an error string if the file exceeds max_bytes, None otherwise.
    Also emits a warning for files over LARGE_FILE_WARNING.
    """
    size     = path.stat().st_size
    size_str = f"{size / 1_048_576:.0f} MB"

    if size > max_bytes:
        limit_str = f"{max_bytes / 1_048_576:.0f} MB"
        return (
            f"File too large ({size_str} > {limit_str}). "
            f"Use --sample N to read only the first N rows, "
            f"or --max-size to raise the limit."
        )

    if size > LARGE_FILE_WARNING:
        log().warning(f"{path.name}: large file ({size_str}) — loading may take a while")

    return None


# ─────────────────────────────────────────────────────────────
# SAFE STEM (path-traversal guard)
# ─────────────────────────────────────────────────────────────


def safe_stem(raw: str) -> str:
    """
    Strips any directory components from a stem so output files
    cannot escape REPORTS_DIR.  e.g. '../../etc/passwd' → 'passwd'
    """
    return Path(raw).name


# ─────────────────────────────────────────────────────────────
# ENCODING AND SEPARATOR DETECTION (CSV only)
# ─────────────────────────────────────────────────────────────


def detect_encoding(path: Path) -> str:
    """
    Detects the CSV encoding.
    Uses charset_normalizer when available (more accurate on large/mixed files);
    falls back to probing candidate encodings manually.
    """
    if _CHARSET_NORMALIZER:
        result = _cn_from_path(path).best()
        if result is not None:
            enc = result.encoding
            log().debug(f"{path.name}: encoding detected by charset_normalizer = '{enc}'")
            return enc
        log().debug(f"{path.name}: charset_normalizer inconclusive, using fallback")

    for enc in ENCODING_CANDIDATES:
        try:
            with open(path, "r", encoding=enc) as f:
                f.read(65_536)   # 64 KB — catches encoding errors further into the file
            log().debug(f"{path.name}: encoding detected = '{enc}'")
            return enc
        except UnicodeDecodeError:
            continue

    log().warning(f"{path.name}: no encoding detected, falling back to 'latin-1'")
    return "latin-1"


def detect_separator(path: Path, encoding: str) -> str:
    """
    Uses csv.Sniffer on the first 8 KB to detect the column separator.
    More robust than counting raw occurrences, especially with quoted text fields.
    """
    try:
        with open(path, "r", encoding=encoding, errors="replace") as f:
            sample = f.read(8_192)
        dialect = csv.Sniffer().sniff(sample, delimiters=SNIFFER_DELIMITERS)
        log().debug(f"{path.name}: separator detected by Sniffer = '{dialect.delimiter}'")
        return dialect.delimiter
    except csv.Error:
        log().warning(f"{path.name}: Sniffer could not detect separator, defaulting to ','")
        return ","
    except Exception as e:
        log().warning(f"{path.name}: separator detection failed ({e}), defaulting to ','")
        return ","


# ─────────────────────────────────────────────────────────────
# CSV LOADING
# ─────────────────────────────────────────────────────────────


def load_csv(
    path: Path,
    sep: Optional[str],
    encoding: Optional[str],
    nrows: Optional[int],
) -> pd.DataFrame:
    """
    Loads a CSV with automatic encoding and separator detection.
    Falls back through all candidate encodings on UnicodeDecodeError.
    """
    enc   = encoding or detect_encoding(path)
    delim = sep      or detect_separator(path, enc)

    log().info(f"  Separator : '{delim}'  |  Encoding : '{enc}'")

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
        log().warning(f"{path.name}: '{enc}' failed, trying fallback encodings...")
        for enc_alt in ENCODING_CANDIDATES:
            if enc_alt == enc:
                continue
            try:
                kwargs["encoding"] = enc_alt
                df = pd.read_csv(path, **kwargs)
                log().info(f"  Encoding fallback : '{enc_alt}'")
                return df
            except UnicodeDecodeError:
                continue
        raise RuntimeError(f"Could not read {path.name} with any known encoding")


# ─────────────────────────────────────────────────────────────
# EXCEL LOADING
# ─────────────────────────────────────────────────────────────


def check_excel_engine(ext: str) -> None:
    """Raises ImportError with an install hint if the required engine is missing."""
    engine  = EXCEL_ENGINES[ext]
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

    log().info(f"  Engine : '{engine}'")

    kwargs: dict = dict(sheet_name=None, engine=engine)
    if nrows:
        kwargs["nrows"] = nrows

    sheets: Dict[str, pd.DataFrame] = pd.read_excel(path, **kwargs)
    log().info(f"  Sheets found : {list(sheets.keys())}")
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
    log().info(f"  File             : {m['file']}{label}  ({m['file_size']})")
    log().info(f"  Rows analyzed    : {m['rows']:,}" + ("  ← sample" if m["sampled"] else ""))
    log().info(f"  Columns          : {m['columns']}")
    log().info(f"  Global missing   : {m['missing_pct']} %")
    log().info(f"  Duplicate rows   : {m['duplicates']:,}  ({m['dup_pct']} %)")


# ─────────────────────────────────────────────────────────────
# PROFILER CONFIGURATION BY DATASET SIZE
# ─────────────────────────────────────────────────────────────


def build_profiler_config(df: pd.DataFrame, minimal: bool) -> Tuple[dict, str]:
    """
    Returns (config_dict, mode_name) adapted to the dataset size.

    Three tiers:
      - minimal  : fast pass, basic stats only
                   (--minimal flag, or auto when > 5M cells)
      - optimized: medium datasets — Pearson, Spearman, Cramér; no interactions
      - full     : all correlations, interactions, word/character analysis,
                   sensitive-data detection, explorative mode
    """
    n_cells = df.size

    # ── MINIMAL ───────────────────────────────────────────────
    if minimal or n_cells > THRESHOLD_MINIMAL:
        mode   = "minimal (manual)" if minimal else "minimal (auto — dataset too large)"
        config = dict(minimal=True, progress_bar=True)

    # ── OPTIMIZED ─────────────────────────────────────────────
    elif n_cells > THRESHOLD_OPTIMIZED:
        mode   = "optimized"
        config = dict(
            minimal      = False,
            explorative  = True,
            sensitive    = True,
            progress_bar = True,
            correlations = {
                "auto":     {"calculate": True,  "warn_high_correlations": True, "threshold": 0.9},
                "pearson":  {"calculate": True,  "warn_high_correlations": True, "threshold": 0.9},
                "spearman": {"calculate": True,  "warn_high_correlations": True, "threshold": 0.9},
                "kendall":  {"calculate": False},
                "phi_k":    {"calculate": False},
                "cramers":  {"calculate": True,  "warn_high_correlations": True, "threshold": 0.9},
            },
            interactions     = {"continuous": False},
            missing_diagrams = {"bar": True, "matrix": True, "heatmap": False},
            samples          = {"head": 10, "tail": 10},
            vars = {
                "num": {
                    "quantiles":                [0.05, 0.25, 0.5, 0.75, 0.95],
                    "skewness_threshold":       20,
                    "low_categorical_threshold": 5,
                    "chi_squared_threshold":    0.999,
                },
                "cat": {
                    "length":                True,
                    "characters":            False,   # too slow at this size
                    "words":                 False,   # too slow at this size
                    "n_obs":                 10,
                    "chi_squared_threshold": 0.999,
                    "imbalance_threshold":   0.5,
                },
                "bool": {"n_obs": 3, "imbalance_threshold": 0.5},
            },
        )

    # ── FULL ──────────────────────────────────────────────────
    else:
        mode   = "full"
        config = dict(
            minimal      = False,
            explorative  = True,
            sensitive    = True,
            progress_bar = True,
            correlations = {
                "auto":     {"calculate": True, "warn_high_correlations": True, "threshold": 0.9},
                "pearson":  {"calculate": True, "warn_high_correlations": True, "threshold": 0.9},
                "spearman": {"calculate": True, "warn_high_correlations": True, "threshold": 0.9},
                "kendall":  {"calculate": True, "warn_high_correlations": True, "threshold": 0.9},
                "phi_k":    {"calculate": True, "warn_high_correlations": True, "threshold": 0.9},
                "cramers":  {"calculate": True, "warn_high_correlations": True, "threshold": 0.9},
            },
            interactions     = {"continuous": True},
            missing_diagrams = {"bar": True, "matrix": True, "heatmap": True},
            samples          = {"head": 20, "tail": 20},
            vars = {
                "num": {
                    "quantiles":                [0.05, 0.1, 0.25, 0.5, 0.75, 0.9, 0.95],
                    "skewness_threshold":       20,
                    "low_categorical_threshold": 5,
                    "chi_squared_threshold":    0.999,
                },
                "cat": {
                    "length":                True,
                    "characters":            True,
                    "words":                 True,
                    "n_obs":                 10,
                    "chi_squared_threshold": 0.999,
                    "imbalance_threshold":   0.5,
                },
                "bool": {"n_obs": 3, "imbalance_threshold": 0.5},
            },
        )

    log().info(f"  Profiler mode    : {mode}  ({n_cells:,} cells)")
    return config, mode


# ─────────────────────────────────────────────────────────────
# REPORT GENERATION
# ─────────────────────────────────────────────────────────────


def generate_report(
    df: pd.DataFrame,
    stem: str,
    config: dict,
    lang: str,
    translations_loaded: bool,
    reports_dir: Path,
) -> Path:
    """
    Generates the HTML report for a given language.
    Returns the path of the generated file.
    """
    info   = LANGUAGES[lang]
    output = reports_dir / f"{safe_stem(stem)}{info['suffix']}.html"
    title  = f"{info['title_prefix']} — {stem}"

    log().info(f"  Generating report {info['flag']} {info['label']}...")

    kwargs = dict(title=title, **config)
    if lang == "es" and translations_loaded:
        kwargs["locale"] = "es"

    t = time.time()
    ProfileReport(df, **kwargs).to_file(str(output))

    size_mb = output.stat().st_size / 1_048_576
    log().info(f"  ✅ {output.name}  ({size_mb:.1f} MB, {time.time() - t:.0f}s)")
    return output


# ─────────────────────────────────────────────────────────────
# SHEET PROCESSING (shared by CSV and Excel)
# ─────────────────────────────────────────────────────────────


def process_sheet(
    df: pd.DataFrame,
    path: Path,
    stem: str,
    sheet_name: str,
    args: argparse.Namespace,
    translations_loaded: bool,
    reports_dir: Path,
) -> SheetResult:
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
            out = generate_report(df, stem, config, lang, translations_loaded, reports_dir)
            result.reports.append(out.name)
        except Exception as e:
            log().error(f"Failed to generate '{lang}' report for '{stem}': {e}")
            log().debug(traceback.format_exc())
            all_ok = False

    result.ok         = all_ok
    result.total_time = round(time.time() - t_start, 1)
    return result


# ─────────────────────────────────────────────────────────────
# CSV FILE PIPELINE
# ─────────────────────────────────────────────────────────────


def process_csv_file(
    path: Path,
    args: argparse.Namespace,
    translations_loaded: bool,
    reports_dir: Path,
) -> List[SheetResult]:
    """Loads and profiles a CSV file. Returns a list with one SheetResult."""
    log().info(f"\n{'─' * 60}")
    log().info(f"  Processing CSV: {path.name}")
    log().info(f"{'─' * 60}")

    result = SheetResult(file_name=path.name)

    err = check_file_size(path, args.max_size)
    if err:
        log().error(f"{path.name}: {err}")
        result.error = err
        return [result]

    try:
        df = load_csv(path, args.sep, args.encoding, args.sample)
    except Exception as e:
        log().error(f"Failed to load {path.name}: {e}")
        log().debug(traceback.format_exc())
        result.error = str(e)
        return [result]

    return [process_sheet(df, path, path.stem, "", args, translations_loaded, reports_dir)]


# ─────────────────────────────────────────────────────────────
# EXCEL FILE PIPELINE
# ─────────────────────────────────────────────────────────────


def process_excel_file(
    path: Path,
    args: argparse.Namespace,
    translations_loaded: bool,
    reports_dir: Path,
) -> List[SheetResult]:
    """
    Loads all sheets from an Excel/ODS file and profiles each one.
    Returns one SheetResult per sheet.
    """
    log().info(f"\n{'─' * 60}")
    log().info(f"  Processing Excel: {path.name}")
    log().info(f"{'─' * 60}")

    err = check_file_size(path, args.max_size)
    if err:
        log().error(f"{path.name}: {err}")
        return [SheetResult(file_name=path.name, error=err)]

    try:
        sheets = load_excel_sheets(path, args.sample)
    except Exception as e:
        log().error(f"Failed to load {path.name}: {e}")
        log().debug(traceback.format_exc())
        return [SheetResult(file_name=path.name, error=str(e))]

    results = []
    for sheet_name, df in sheets.items():
        log().info(f"\n  ── Sheet: '{sheet_name}' ──")
        safe_sheet = sheet_name.replace(" ", "_").replace("/", "-")
        stem       = f"{path.stem}__{safe_sheet}"
        result     = process_sheet(df, path, stem, sheet_name, args, translations_loaded, reports_dir)
        results.append(result)

    return results


# ─────────────────────────────────────────────────────────────
# HTML INDEX
# ─────────────────────────────────────────────────────────────


def _mode_badge(mode: str) -> str:
    """Returns a coloured HTML badge for the analysis mode."""
    colour = {"full": "#2ecc71", "optimized": "#f39c12"}.get(mode.split()[0], "#95a5a6")
    return (
        f'<span style="background:{colour};color:#fff;'
        f'padding:2px 8px;border-radius:3px;font-size:0.82em">'
        f'{html.escape(mode)}</span>'
    )


def generate_index(results: List[SheetResult], reports_dir: Path, total_time: float) -> None:
    """
    Writes reports/index.html — a summary table linking to every generated report.
    Errors are shown at the bottom of the table in red.
    """
    ok_results = [r for r in results if r.ok and r.reports]
    if not ok_results:
        log().debug("No successful results — skipping index.html")
        return

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    report_rows = ""
    for r in ok_results:
        for report_file in r.reports:
            lang_flag  = "🇪🇸" if "_informe_calidad" in report_file else "🇬🇧"
            sheet_cell = html.escape(r.sheet_name) if r.sheet_name else "—"
            report_rows += (
                f"<tr>"
                f'<td><a href="{html.escape(report_file)}">{lang_flag} {html.escape(report_file)}</a></td>'
                f"<td>{html.escape(r.file_name)}</td>"
                f"<td>{sheet_cell}</td>"
                f"<td class='num'>{r.rows:,}</td>"
                f"<td class='num'>{r.columns}</td>"
                f"<td class='num'>{r.missing_pct:.1f} %</td>"
                f"<td class='num'>{r.duplicates:,}</td>"
                f"<td class='ctr'>{_mode_badge(r.analysis_mode)}</td>"
                f"<td class='num'>{r.total_time:.1f}s</td>"
                f"</tr>\n"
            )

    error_rows = ""
    for r in results:
        if not r.ok:
            error_rows += (
                f'<tr class="err">'
                f'<td colspan="2">{html.escape(r.display_name)}</td>'
                f'<td colspan="7">{html.escape(r.error or "unknown error")}</td>'
                f"</tr>\n"
            )

    index_path = reports_dir / "index.html"
    index_path.write_text(
        f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <title>Data Quality — Report Index</title>
  <style>
    body  {{ font-family: Arial, sans-serif; margin: 2rem; color: #333; }}
    h1    {{ color: #2c3e50; margin-bottom: .25rem; }}
    .meta {{ color: #888; font-size: .9em; margin-bottom: 1.5rem; }}
    table {{ border-collapse: collapse; width: 100%; font-size: .9em; }}
    th    {{ background: #2c3e50; color: #fff; padding: 8px 12px; text-align: left; white-space: nowrap; }}
    td    {{ padding: 7px 12px; border-bottom: 1px solid #eee; }}
    td.num {{ text-align: right; }}
    td.ctr {{ text-align: center; }}
    tr:hover td {{ background: #f5f7fa; }}
    tr.err td   {{ color: #c0392b; }}
    a     {{ color: #2980b9; text-decoration: none; }}
    a:hover {{ text-decoration: underline; }}
  </style>
</head>
<body>
  <h1>&#128202; Data Quality &#8212; Report Index</h1>
  <p class="meta">
    Generated: {generated_at} &nbsp;|&nbsp;
    Total time: {total_time:.1f}s &nbsp;|&nbsp;
    Reports: {sum(len(r.reports) for r in ok_results)}
  </p>
  <table>
    <thead>
      <tr>
        <th>Report</th><th>File</th><th>Sheet</th>
        <th>Rows</th><th>Cols</th><th>Missing&nbsp;%</th>
        <th>Duplicates</th><th>Mode</th><th>Time</th>
      </tr>
    </thead>
    <tbody>
{report_rows}{error_rows}    </tbody>
  </table>
</body>
</html>
""",
        encoding="utf-8",
    )

    log().info(f"  📄 Index generated : {index_path}")


# ─────────────────────────────────────────────────────────────
# FINAL SUMMARY
# ─────────────────────────────────────────────────────────────


def print_final_summary(
    results: List[SheetResult], reports_dir: Path, total_time: float
) -> None:
    ok_results  = [r for r in results if r.ok]
    err_results = [r for r in results if not r.ok]

    log().info(f"\n{'═' * 70}")
    log().info("  EXECUTION SUMMARY")
    log().info(f"{'═' * 70}")
    log().info(f"  Total time             : {total_time:.1f}s")
    log().info(f"  ✅ Sheets processed     : {len(ok_results)}")

    if err_results:
        log().info(f"  ❌ Sheets with errors   : {len(err_results)}")
        for r in err_results:
            log().info(f"       • {r.display_name}: {r.error}")

    if ok_results:
        col = 40
        log().info(
            f"\n  {'File / Sheet':<{col}} {'Rows':>8} {'Cols':>5} "
            f"{'Missing%':>9} {'Dups':>7} {'Mode':<30} {'Time':>7}"
        )
        log().info(
            f"  {'─' * col} {'─' * 8} {'─' * 5} {'─' * 9} {'─' * 7} {'─' * 30} {'─' * 7}"
        )
        for r in ok_results:
            log().info(
                f"  {r.display_name:<{col}} {r.rows:>8,} {r.columns:>5} "
                f"{r.missing_pct:>8.1f}% {r.duplicates:>7,} "
                f"{r.analysis_mode:<30} {r.total_time:>6.1f}s"
            )

    log().info(f"\n  Reports available at : {reports_dir}")
    log().info(f"  Index                : {reports_dir / 'index.html'}")
    log().info(f"  Logs saved at        : {LOGS_DIR}")
    log().info(f"{'═' * 70}\n")


# ─────────────────────────────────────────────────────────────
# ARGUMENT PARSING
# ─────────────────────────────────────────────────────────────


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate HTML data quality reports for all files in data/",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f"""
Supported formats: .csv  .xlsx  .xls  .ods
Excel files with multiple sheets generate one report per sheet.

Examples:
  python data_quality_report.py
  python data_quality_report.py --lang both
  python data_quality_report.py --sep ";" --encoding latin-1
  python data_quality_report.py --sample 100000 --minimal
  python data_quality_report.py --output-dir /tmp/my_reports
  python data_quality_report.py --exclude draft.csv ~temp.xlsx
  python data_quality_report.py --max-size 4096
        """,
    )
    parser.add_argument(
        "--lang", default="es", choices=["es", "en", "both"],
        help="Report language: es / en / both  (default: es)",
    )
    parser.add_argument(
        "--sep", default=None,
        help="CSV column separator (auto-detected if not provided)",
    )
    parser.add_argument(
        "--encoding", default=None,
        help="CSV file encoding (auto-detected if not provided)",
    )
    parser.add_argument(
        "--sample", type=int, default=None, metavar="N",
        help="Analyze only the first N rows per sheet",
    )
    parser.add_argument(
        "--minimal", action="store_true",
        help="Fast mode: skip correlations and interactions",
    )
    parser.add_argument(
        "--output-dir", default=None, metavar="DIR",
        help="Directory for HTML reports (default: reports/ next to this script)",
    )
    parser.add_argument(
        "--exclude", nargs="+", default=[], metavar="FILE",
        help="File names in data/ to skip  (e.g. --exclude bad.csv temp.xlsx)",
    )
    parser.add_argument(
        "--max-size", type=int, default=DEFAULT_MAX_FILE_MB, metavar="MB",
        help=f"Max file size in MB before aborting (default: {DEFAULT_MAX_FILE_MB})",
    )
    return parser.parse_args()


# ─────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────


def main() -> None:
    args = parse_args()

    # Resolve directories before logging so paths appear in the banner
    reports_dir   = Path(args.output_dir).resolve() if args.output_dir else REPORTS_DIR
    args.max_size = args.max_size * 1_048_576   # MB → bytes

    setup_logging(LOGS_DIR)

    log().info("╔══════════════════════════════════════════════════════╗")
    log().info("║     DATA QUALITY ANALYZER — ydata-profiling         ║")
    log().info("╚══════════════════════════════════════════════════════╝")

    init_directories(reports_dir)

    # Load translations
    translations_loaded = False
    if args.lang in ("es", "both"):
        if not MULTILINGUAL:
            log().warning("ydata-profiling-multilingual not installed → switching to English")
            log().warning("  Run: pip uninstall ydata-profiling && pip install ydata-profiling-multilingual")
            args.lang = "en"
        else:
            translations_loaded = load_translations()

    # Discover files
    data_files = find_data_files(args.exclude)
    if not data_files:
        log().warning(f"No supported files found in: {DATA_DIR}")
        log().warning(f"Supported extensions: {', '.join(sorted(ALL_EXTENSIONS))}")
        sys.exit(0)

    lang_label = {"es": "Spanish 🇪🇸", "en": "English 🇬🇧", "both": "Spanish 🇪🇸 + English 🇬🇧"}
    log().info(f"  Root            : {ROOT_DIR}")
    log().info(f"  Data dir        : {DATA_DIR}")
    log().info(f"  Reports dir     : {reports_dir}")
    log().info(f"  Language(s)     : {lang_label[args.lang]}")
    log().info(f"  Max file size   : {args.max_size // 1_048_576} MB")
    if args.exclude:
        log().info(f"  Excluded        : {', '.join(args.exclude)}")
    log().info(f"  Files found     : {len(data_files)}")
    for f in data_files:
        log().info(f"    • {f.name}")

    # Process each file
    t_start     = time.time()
    all_results: List[SheetResult] = []

    for path in data_files:
        ext = path.suffix.lower()
        if ext in CSV_EXTENSIONS:
            file_results = process_csv_file(path, args, translations_loaded, reports_dir)
        else:
            file_results = process_excel_file(path, args, translations_loaded, reports_dir)
        all_results.extend(file_results)

    total_time = round(time.time() - t_start, 1)
    generate_index(all_results, reports_dir, total_time)
    print_final_summary(all_results, reports_dir, total_time)


if __name__ == "__main__":
    main()
