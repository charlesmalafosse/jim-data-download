#!/usr/bin/env python3
"""
Upload an aggregated option-price CSV into a MySQL table.

Reads the CSV produced by the VBA download / aggregate_files.py, maps CSV
headers to the target SQL column names via CSV_TO_SQL, and REPLACE INTOs
the rows so that existing rows (matched on the table's primary/unique key)
are overwritten. CSV columns not listed in CSV_TO_SQL (RIC, RIC_Underlying,
Dividend, ...) are dropped. SQL columns with no CSV source are written as
NULL.

Configure the inputs in the __main__ block at the bottom of the file.
"""

import logging
import sys
import time
from pathlib import Path

import pandas as pd
import pymysql


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
    stream=sys.stdout,
)
log = logging.getLogger("upload_to_mysql")


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

# Columns parsed as datetime and sent to MySQL as Python datetime objects.
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

    # Datetime columns: parse to pandas datetime so pymysql sends them as
    # proper DATETIME values. Unparseable -> NaT -> None.
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

    # Convert NaN / NaT to Python None (pymysql needs None for SQL NULL) and
    # drop pandas wrappers so each cell is a native Python scalar.
    log.info("Finalizing NaN/NaT -> None conversion")
    t0 = time.monotonic()
    df = df.astype(object).where(pd.notna(df), None)
    for col in DATETIME_COLUMNS:
        if col in df.columns:
            df[col] = df[col].map(
                lambda x: x.to_pydatetime() if hasattr(x, "to_pydatetime") else x
            )
    log.info("Finalize done in %.1fs", time.monotonic() - t0)
    return df


def build_replace_sql(table: str) -> str:
    cols_sql = ", ".join(f"`{c}`" for c in SQL_COLUMNS)
    placeholders = ", ".join(["%s"] * len(SQL_COLUMNS))
    return f"REPLACE INTO `{table}` ({cols_sql}) VALUES ({placeholders})"


class SchemaMismatchError(RuntimeError):
    """Raised when the target table is missing columns the script will write."""


def check_schema(conn, database: str, table: str) -> None:
    """Fail fast if the target table is missing or doesn't expose every
    column in SQL_COLUMNS. Case-insensitive match on column names."""
    log.info("Checking schema of `%s`.`%s`", database, table)
    with conn.cursor() as cur:
        cur.execute(
            "SELECT COLUMN_NAME FROM information_schema.COLUMNS "
            "WHERE TABLE_SCHEMA = %s AND TABLE_NAME = %s",
            (database, table),
        )
        rows = cur.fetchall()

    if not rows:
        raise SchemaMismatchError(
            f"Table `{database}`.`{table}` does not exist (or the user "
            f"has no visibility on it). Create it first."
        )

    existing = {r[0] for r in rows}
    existing_lower = {c.lower() for c in existing}
    missing = [c for c in SQL_COLUMNS if c.lower() not in existing_lower]

    if missing:
        have = sorted(existing)
        raise SchemaMismatchError(
            f"Table `{database}`.`{table}` is missing {len(missing)} "
            f"column(s) the script will write: {missing}. "
            f"Columns currently in the table: {have}."
        )

    extras = [c for c in sorted(existing) if c.lower() not in {s.lower() for s in SQL_COLUMNS}]
    if extras:
        log.info(
            "Table has %d extra column(s) not written by this script (ok): %s",
            len(extras), extras,
        )
    log.info("Schema OK — all %d expected columns present", len(SQL_COLUMNS))


class BatchUploadError(RuntimeError):
    """Raised when a batch fails. Carries the failing batch number so the
    caller can restart from it."""

    def __init__(self, batch_number: int, message: str):
        super().__init__(message)
        self.batch_number = batch_number


def upload(
    csv_path: Path,
    table: str,
    host: str,
    port: int,
    user: str,
    password: str,
    database: str,
    batch_size: int = 1000,
    start_batch: int = 1,
) -> int:
    """Upload CSV rows to MySQL in batches, committing after each batch.

    Batches are 1-indexed. `start_batch=1` uploads everything; set to a
    higher number to resume after a failure (batches before it are skipped).
    On any batch failure, rolls back that batch, closes the connection, and
    raises BatchUploadError with the failing batch number.
    """
    if start_batch < 1:
        raise ValueError(f"start_batch must be >= 1, got {start_batch}")

    log.info("==== Upload started: %s -> `%s` ====", csv_path.name, table)
    df = load_csv(csv_path)

    total = len(df)
    if total == 0:
        log.warning("No rows to upload from %s", csv_path)
        return 0

    total_batches = (total + batch_size - 1) // batch_size
    if start_batch > total_batches:
        raise ValueError(
            f"start_batch={start_batch} exceeds total batches={total_batches}"
        )
    skipped_rows = (start_batch - 1) * batch_size
    if start_batch > 1:
        log.info(
            "Resuming at batch %d — skipping first %d rows", start_batch, skipped_rows
        )

    sql = build_replace_sql(table)
    log.info(
        "Connecting to mysql://%s@%s:%d/%s (table=%s)",
        user, host, port, database, table,
    )
    log.info(
        "Plan: %d rows, %d batches of %d, starting at batch %d",
        total, total_batches, batch_size, start_batch,
    )
    t_conn = time.monotonic()
    conn = pymysql.connect(
        host=host,
        port=port,
        user=user,
        password=password,
        database=database,
        autocommit=False,
        charset="utf8mb4",
    )
    log.info("Connected in %.2fs", time.monotonic() - t_conn)

    uploaded = 0
    t_run = time.monotonic()
    try:
        check_schema(conn, database, table)
        for batch_num in range(start_batch, total_batches + 1):
            start = (batch_num - 1) * batch_size
            end = min(start + batch_size, total)
            # Slice the DataFrame lazily so only the current batch is
            # converted to tuples — keeps memory bounded and lets the first
            # commit happen without waiting to materialize the whole file.
            chunk = list(
                df.iloc[start:end].itertuples(index=False, name=None)
            )
            t_batch = time.monotonic()
            try:
                with conn.cursor() as cur:
                    cur.executemany(sql, chunk)
                conn.commit()
            except Exception as e:
                conn.rollback()
                log.error(
                    "Batch %d/%d FAILED after %.1fs: %s",
                    batch_num, total_batches, time.monotonic() - t_batch, e,
                )
                raise BatchUploadError(
                    batch_num,
                    f"Batch {batch_num}/{total_batches} failed: {e}. "
                    f"Restart with start_batch={batch_num}.",
                ) from e
            dt = time.monotonic() - t_batch
            uploaded += len(chunk)
            rows_done = end
            elapsed = time.monotonic() - t_run
            rate = uploaded / elapsed if elapsed > 0 else 0.0
            remaining = total - rows_done
            eta_s = remaining / rate if rate > 0 else 0.0
            log.info(
                "batch %d/%d ok in %.2fs (%d rows, %.0f rows/s) — %d/%d of file, ETA %.0fs",
                batch_num, total_batches, dt, len(chunk), rate,
                rows_done, total, eta_s,
            )
    finally:
        conn.close()
        log.info("Connection closed")

    total_elapsed = time.monotonic() - t_run
    log.info(
        "==== Done. Uploaded %d rows in %.1fs (batches %d..%d) into `%s` ====",
        uploaded, total_elapsed, start_batch, total_batches, table,
    )
    return uploaded


if __name__ == "__main__":
    # ---- Configure here -------------------------------------------------
    CSV_PATH = Path(r"c:\Users\charl\Dropbox\MARINER\DOWNLOAD_FILES\0 - TEMPLATE\DOWNLOAD_MONTHLY\python_scripts\downloaded_SUGAR_20260309.csv")
    TABLE_NAME = "test_import"

    DB_HOST = "localhost"
    DB_PORT = 3306
    DB_USER = "python"
    DB_PASSWORD = "1234"
    DB_NAME = "project_bbrdvl"

    BATCH_SIZE = 1000
    START_BATCH = 1  # set to N to resume from batch N after a failure
    # ---------------------------------------------------------------------

    try:
        upload(
            csv_path=CSV_PATH,
            table=TABLE_NAME,
            host=DB_HOST,
            port=DB_PORT,
            user=DB_USER,
            password=DB_PASSWORD,
            database=DB_NAME,
            batch_size=BATCH_SIZE,
            start_batch=START_BATCH,
        )
    except SchemaMismatchError as e:
        log.error("Schema mismatch: %s", e)
        sys.exit(3)
    except BatchUploadError as e:
        log.error("%s", e)
        sys.exit(2)
    except Exception as e:
        log.exception("Upload failed: %s", e)
        sys.exit(1)
