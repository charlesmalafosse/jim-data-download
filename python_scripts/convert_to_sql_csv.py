#!/usr/bin/env python3
"""
Convert an aggregated option-price CSV into a SQL-ready CSV.

Reads the CSV produced by the VBA download / aggregate_files.py, maps CSV
headers to the target SQL column names via CSV_TO_SQL, cleans sentinel
values, coerces numeric and datetime columns, and writes a new CSV whose
column names, ordering, and cell formats match what the target SQL table
expects. CSV columns not listed in CSV_TO_SQL (RIC, RIC_Underlying,
Dividend, ...) are dropped. SQL columns with no CSV source are written as
NULL.

This is the offline counterpart to upload_to_mysql.py — same cleaning
pipeline, but the output is a CSV file instead of a MySQL table.

Configure the inputs in the __main__ block at the bottom of the file.
"""

import csv
import logging
import sys
import time
from pathlib import Path

import pandas as pd


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
    stream=sys.stdout,
)
log = logging.getLogger("convert_to_sql_csv")


# Mapping from CSV header (VBA SetupStagingSheet) to target SQL column name.
# Any CSV column not listed here is ignored (e.g. Dividend, RIC, RIC_Underlying).
CSV_TO_SQL = {
    "Spot_Date": "Spot_date",
    "Premium": "Premium",
    "Ticker": "Ticker",
    "Maturity": "Maturity",
    "Interest_rate": "Interest_rate",
    "Spot": "Spot",
    "Strike": "Strike",
    "Type": "Type",
    "Implied_Volatility": "Implied_Volatility",
    "Delta": "Delta",
    "Vega": "Vega",
    "Gamma": "Gamma",
    "Theta": "Theta",
    "Rho": "Rho",
    "Lot_size": "Lot_size",
    "Name": "Name",
    "Reference": "Reference",
    "ccy_pair": "ccy_pair",
    "Internal_ID": "Internal_ID",
    "DDELTA/DVOL": "ddeltadvol",
    "DDELTA/DVOLDVOL": "ddeltadvoldvol",
    "DDELTA/DTIME": "charm",
    "DGAMMA/DSPOT": "speed",
    "DGAMMA/DVOL": "dgammaPdvol",
    "DVEGA/DVOL": "dvegadvol",
    "DVEGA/DVOLDVOL": "dvegadvoldvol",
}

# Target SQL column order. Columns without a CSV source are written as NULL.
SQL_COLUMNS = [
    "Spot_date",
    "Premium",
    "Ticker",
    "Maturity",
    "Interest_rate",
    "Spot",
    "Strike",
    "Type",
    "Implied_Volatility",
    "Delta",
    "Vega",
    "Gamma",
    "Theta",
    "Rho",
    "Lot_size",
    "Name",
    "Reference",
    "ccy_pair",
    "Internal_ID",
    "ddeltadvol",
    "ddeltadvoldvol",
    "charm",
    "speed",
    "dgammaPdvol",
    "dgammaPdtime",
    "dvegadtime",
    "dvegadvol",
    "dvegadvoldvol",
]

# SQL columns with no corresponding CSV source — filled with NULL.
SQL_COLUMNS_WITHOUT_SOURCE = [c for c in SQL_COLUMNS if c not in CSV_TO_SQL.values()]

# Columns parsed as datetime and written as 'YYYY-MM-DD HH:MM:SS'.
DATETIME_COLUMNS = {"Spot_date", "Maturity"}

# Columns that must become numeric in SQL. Anything non-parseable
# (including "NA", "#N/A", "#VALUE!", blank, etc.) becomes NULL.
NUMERIC_COLUMNS = {
    "Premium",
    "Interest_rate",
    "Spot",
    "Strike",
    "Implied_Volatility",
    "Delta",
    "Vega",
    "Gamma",
    "Theta",
    "Rho",
    "Lot_size",
    "ddeltadvol",
    "ddeltadvoldvol",
    "charm",
    "speed",
    "dgammaPdvol",
    "dgammaPdtime",
    "dvegadtime",
    "dvegadvol",
    "dvegadvoldvol",
}

# Strings that should always become NULL, regardless of column type.
NULL_SENTINELS = {
    "",
    "NA",
    "N/A",
    "#N/A",
    "#VALUE!",
    "#DIV/0!",
    "#REF!",
    "#NAME?",
    "#NULL!",
    "#NUM!",
    "NAN",
    "NULL",
    "NONE",
}

# Output format for datetime columns (MySQL DATETIME literal).
DATETIME_FORMAT = "%Y-%m-%d %H:%M:%S"

# Token written for SQL NULL in the output CSV. Empty string is the most
# portable for `LOAD DATA INFILE ... FIELDS TERMINATED BY ',' ... ` when
# combined with `NULL` handling, but `\N` is the MySQL convention. Pick
# whichever your import pipeline expects and override in __main__.
DEFAULT_NULL_TOKEN = ""

# Sentinel written to numeric columns in place of NaN / empty values so
# the output CSV never contains empty cells for numeric fields. Change
# this constant to adjust the fill value globally.
NAN_FILL_VALUE = 99999


def _clean_text(series: pd.Series) -> pd.Series:
    """Trim and null-out sentinel strings. Returns object dtype."""
    s = series.astype(str).str.strip()
    return s.where(~s.str.upper().isin(NULL_SENTINELS), other=None)


def load_csv(csv_path: Path) -> pd.DataFrame:
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV not found: {csv_path}")

    size_mb = csv_path.stat().st_size / (1024 * 1024)
    log.info("Reading CSV %s (%.1f MB)", csv_path, size_mb)
    t0 = time.monotonic()
    df = pd.read_csv(csv_path, dtype=str, keep_default_na=False, na_values=[])
    log.info("Parsed %d rows x %d cols in %.1fs", len(df), len(df.columns), time.monotonic() - t0)

    # Check all mapped CSV source columns are present.
    missing_sources = [c for c in CSV_TO_SQL if c not in df.columns]
    if missing_sources:
        raise ValueError(f"CSV is missing expected columns: {missing_sources}")

    dropped = [c for c in df.columns if c not in CSV_TO_SQL]
    if dropped:
        log.info("Dropping %d unmapped CSV columns: %s", len(dropped), dropped)

    # Keep only mapped source columns and rename to target SQL names.
    df = df[list(CSV_TO_SQL.keys())].rename(columns=CSV_TO_SQL).copy()

    # Add NULL columns for SQL fields with no CSV source (e.g. dgammaPdtime,
    # dvegadtime — greeks that the VBA download doesn't produce yet).
    if SQL_COLUMNS_WITHOUT_SOURCE:
        log.info(
            "Adding %d SQL columns with no CSV source (filled with NULL): %s",
            len(SQL_COLUMNS_WITHOUT_SOURCE),
            SQL_COLUMNS_WITHOUT_SOURCE,
        )
    for col in SQL_COLUMNS_WITHOUT_SOURCE:
        df[col] = None

    # Reorder to canonical SQL column order.
    df = df[SQL_COLUMNS]

    # First pass: strip whitespace and blank out known sentinel strings.
    log.info("Cleaning sentinel values (NA / #N/A / ...) across %d columns", len(SQL_COLUMNS))
    t0 = time.monotonic()
    for col in SQL_COLUMNS:
        if col in SQL_COLUMNS_WITHOUT_SOURCE:
            continue
        df[col] = _clean_text(df[col])
    log.info("Sentinel cleaning done in %.1fs", time.monotonic() - t0)

    # Numeric columns: coerce; anything unparseable -> NaN -> None.
    log.info("Casting %d numeric columns", len(NUMERIC_COLUMNS))
    t0 = time.monotonic()
    null_counts = {}
    for col in NUMERIC_COLUMNS:
        if col in SQL_COLUMNS_WITHOUT_SOURCE:
            continue
        coerced = pd.to_numeric(df[col], errors="coerce")
        n_nulls = int(coerced.isna().sum() - df[col].isna().sum())
        if n_nulls > 0:
            null_counts[col] = n_nulls
        df[col] = coerced
    log.info("Numeric cast done in %.1fs", time.monotonic() - t0)

    # Datetime columns: parse to pandas datetime. Unparseable -> NaT -> None.
    log.info("Parsing %d datetime columns", len(DATETIME_COLUMNS))
    t0 = time.monotonic()
    for col in DATETIME_COLUMNS:
        if col in SQL_COLUMNS_WITHOUT_SOURCE:
            continue
        parsed = pd.to_datetime(df[col], errors="coerce")
        n_nulls = int(parsed.isna().sum() - df[col].isna().sum())
        if n_nulls > 0:
            null_counts[col] = n_nulls
        df[col] = parsed
    log.info("Datetime parse done in %.1fs", time.monotonic() - t0)

    if null_counts:
        log.warning("Coerced non-parseable values to NULL:")
        for col, n in null_counts.items():
            log.warning("  %s: %d value(s)", col, n)
    else:
        log.info("No non-parseable values found")

    return df


def convert(
    csv_path: Path,
    output_path: Path,
    null_token: str = DEFAULT_NULL_TOKEN,
) -> int:
    """Convert the input CSV to a SQL-ready CSV and return the row count."""
    log.info("==== Convert started: %s -> %s ====", csv_path.name, output_path.name)
    df = load_csv(csv_path)

    total = len(df)
    if total == 0:
        log.warning("No rows to convert from %s", csv_path)

    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Build a cleaned DataFrame where NaN/NaT become None and datetimes are
    # formatted as strings. Then let pandas.to_csv handle quoting.
    log.info("Formatting %d rows for output", total)
    t0 = time.monotonic()

    out = df.copy()

    # Fill NaN in numeric columns with NAN_FILL_VALUE so the output never
    # contains empty cells for numeric fields. Cast each column to object
    # so the int fill renders as "99999" (not "99999.0") while non-null
    # float values keep their full precision.
    fill_counts = {}
    for col in NUMERIC_COLUMNS:
        if col not in out.columns:
            continue
        mask = out[col].isna()
        n = int(mask.sum())
        if n:
            fill_counts[col] = n
        out[col] = out[col].astype(object)
        out.loc[mask, col] = NAN_FILL_VALUE
    if fill_counts:
        total_filled = sum(fill_counts.values())
        log.info(
            "Filled %d NaN cell(s) in numeric columns with %s",
            total_filled, NAN_FILL_VALUE,
        )
        for col, n in fill_counts.items():
            log.info("  %s: %d", col, n)
    else:
        log.info("No NaN values found in numeric columns")

    # Format datetimes as strings up front so to_csv doesn't print ISO with 'T'.
    for col in DATETIME_COLUMNS:
        if col not in out.columns:
            continue
        out[col] = out[col].apply(
            lambda v: v.strftime(DATETIME_FORMAT) if pd.notna(v) else None
        )

    # Replace any remaining NaN/NaT/None (text / datetime columns) with the
    # null_token so the CSV has an explicit NULL marker. Numeric columns
    # were already filled with NAN_FILL_VALUE above.
    out = out.astype(object).where(pd.notna(out), None)

    log.info("Formatting done in %.1fs", time.monotonic() - t0)

    log.info("Writing %s", output_path)
    t0 = time.monotonic()
    out.to_csv(
        output_path,
        index=False,
        columns=SQL_COLUMNS,
        na_rep=null_token,
        quoting=csv.QUOTE_MINIMAL,
        lineterminator="\n",
    )
    log.info("Wrote %d rows in %.1fs", total, time.monotonic() - t0)

    size_mb = output_path.stat().st_size / (1024 * 1024)
    log.info("==== Done. Output %s (%.1f MB, %d rows) ====", output_path, size_mb, total)
    return total


if __name__ == "__main__":
    # ---- Configure here -------------------------------------------------
    CSV_PATH = Path(r"c:\Users\charl\Dropbox\MARINER\DOWNLOAD_FILES\0 - TEMPLATE\DOWNLOAD_MONTHLY\python_scripts\downloaded_SUGAR_20260309.csv")
    OUTPUT_PATH = CSV_PATH.with_name(CSV_PATH.stem + "_sql.csv")

    # "" -> empty field for NULL (works with LOAD DATA INFILE when the
    # column allows NULL and you use `SET col = NULLIF(col, '')`).
    # "\\N" -> MySQL's default NULL marker for LOAD DATA INFILE.
    NULL_TOKEN = ""
    # ---------------------------------------------------------------------

    try:
        convert(
            csv_path=CSV_PATH,
            output_path=OUTPUT_PATH,
            null_token=NULL_TOKEN,
        )
    except Exception as e:
        log.exception("Convert failed: %s", e)
        sys.exit(1)
