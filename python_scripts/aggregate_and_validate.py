#!/usr/bin/env python3
"""
Aggregate VBA option-download batch CSVs and/or validate the aggregated output.

Subcommands:
    aggregate   Combine batch CSVs into one aggregated file
    validate    Run quality checks on an aggregated CSV
    all         Aggregate and then validate the result

Usage:
    python aggregate_and_validate.py aggregate --input-dir . --output combined.csv
    python aggregate_and_validate.py validate combined.csv --iv-max 5.0
    python aggregate_and_validate.py all --input-dir . --output combined.csv
"""

import argparse
import glob
import json
import os
import sys
from datetime import datetime

import pandas as pd


# ============================================================================
# Shared constants
# ============================================================================

# CSV columns as defined in VBA SetupStagingSheet
CSV_COLUMNS = [
    "Spot_Date",
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
    "Dividend",
    "DDELTA/DVOL",
    "DDELTA/DVOLDVOL",
    "DDELTA/DTIME",
    "DGAMMA/DSPOT",
    "DGAMMA/DVOL",
    "DVEGA/DVOL",
    "DVEGA/DVOLDVOL",
    "RIC",
    "RIC_Underlying",
]

# Key columns for duplicate detection (used by both aggregate and validate)
KEY_COLUMNS = ["Spot_Date", "Ticker", "Maturity", "Strike", "Type"]

# Columns that should never be empty (validation only)
CRITICAL_COLUMNS = [
    "Spot_Date", "Ticker", "Maturity", "Spot", "Strike", "Type",
    "Implied_Volatility", "Premium", "Reference",
]


# ============================================================================
# Aggregation functions
# ============================================================================

def parse_override(override_str: str) -> tuple[str, str]:
    """Parse a single override string like 'column=value'."""
    if "=" not in override_str:
        raise ValueError(f"Invalid override format: {override_str}. Expected 'column=value'")
    col, val = override_str.split("=", 1)
    return col.strip(), val.strip()


def load_overrides_from_file(filepath: str) -> dict:
    """Load overrides from a JSON file."""
    with open(filepath, "r") as f:
        return json.load(f)


def find_csv_files(input_dir: str, pattern: str) -> list[str]:
    """Find all CSV files matching the pattern in the input directory."""
    search_path = os.path.join(input_dir, pattern)
    files = glob.glob(search_path)
    files.sort(key=os.path.getmtime)  # oldest first
    return files


def read_csv_file(filepath: str) -> pd.DataFrame:
    """Read a single CSV file with proper handling."""
    try:
        df = pd.read_csv(filepath, dtype=str)
        print(f"  Read {len(df)} rows from {os.path.basename(filepath)}")
        return df
    except Exception as e:
        print(f"  Warning: Could not read {filepath}: {e}")
        return pd.DataFrame()


def aggregate_files(files: list[str]) -> pd.DataFrame:
    """Read and concatenate all CSV files."""
    if not files:
        return pd.DataFrame()

    print(f"\nReading {len(files)} files...")
    dfs = []
    for f in files:
        df = read_csv_file(f)
        if not df.empty:
            df["_source_file"] = os.path.basename(f)
            dfs.append(df)

    if not dfs:
        return pd.DataFrame()

    combined = pd.concat(dfs, ignore_index=True)
    print(f"\nTotal rows before deduplication: {len(combined)}")
    return combined


def apply_overrides(df: pd.DataFrame, overrides: dict) -> pd.DataFrame:
    """Apply field overrides to the dataframe."""
    if not overrides:
        return df

    print(f"\nApplying {len(overrides)} override(s)...")
    for col, value in overrides.items():
        if col in df.columns:
            print(f"  Setting {col} = {value}")
            df[col] = value
        else:
            print(f"  Warning: Column '{col}' not found in data, adding it")
            df[col] = value

    return df


def remove_duplicates(df: pd.DataFrame, keep: str = "last") -> pd.DataFrame:
    """Remove duplicate rows based on key columns."""
    existing_keys = [k for k in KEY_COLUMNS if k in df.columns]
    if not existing_keys:
        print("Warning: No key columns found for deduplication")
        return df

    before_count = len(df)
    df = df.drop_duplicates(subset=existing_keys, keep=keep)
    after_count = len(df)

    if before_count != after_count:
        print(f"Removed {before_count - after_count} duplicate rows (kept {keep})")

    return df


def convert_types(df: pd.DataFrame) -> pd.DataFrame:
    """Convert columns to appropriate types."""
    numeric_cols = [
        "Premium", "Interest_rate", "Spot", "Strike",
        "Implied_Volatility", "Delta", "Vega", "Gamma", "Theta", "Rho",
        "Lot_size", "Dividend",
        "DDELTA/DVOL", "DDELTA/DVOLDVOL", "DDELTA/DTIME",
        "DGAMMA/DSPOT", "DGAMMA/DVOL", "DVEGA/DVOL", "DVEGA/DVOLDVOL",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    for col in ["Spot_Date", "Maturity"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    return df


def filter_data(df: pd.DataFrame, filters: dict) -> pd.DataFrame:
    """Apply filters to the data."""
    if not filters:
        return df

    print(f"\nApplying {len(filters)} filter(s)...")
    for col, condition in filters.items():
        if col not in df.columns:
            print(f"  Warning: Filter column '{col}' not found")
            continue

        if condition.startswith(">="):
            val = float(condition[2:])
            df = df[df[col] >= val]
        elif condition.startswith("<="):
            val = float(condition[2:])
            df = df[df[col] <= val]
        elif condition.startswith(">"):
            val = float(condition[1:])
            df = df[df[col] > val]
        elif condition.startswith("<"):
            val = float(condition[1:])
            df = df[df[col] < val]
        elif condition.startswith("!="):
            val = condition[2:]
            df = df[df[col].astype(str) != val]
        elif condition.startswith("=="):
            val = condition[2:]
            df = df[df[col].astype(str) == val]
        else:
            df = df[df[col].astype(str) == condition]

        print(f"  Filter {col} {condition}: {len(df)} rows remaining")

    return df


def parse_filter_args(filter_strs: list[str]) -> dict:
    """Parse --filter args like 'column>value' into {col: '>value'}."""
    filters = {}
    for filter_str in filter_strs:
        for op in [">=", "<=", "!=", "==", ">", "<"]:
            if op in filter_str:
                col = filter_str.split(op)[0]
                filters[col] = filter_str[len(col):]
                break
        else:
            if "=" in filter_str:
                col, val = filter_str.split("=", 1)
                filters[col] = f"=={val}"
    return filters


def run_aggregate(args) -> str:
    """Run aggregation pipeline. Returns the output file path."""
    files = find_csv_files(args.input_dir, args.pattern)
    if not files:
        print(f"No files found matching '{args.pattern}' in '{args.input_dir}'")
        sys.exit(1)

    print(f"Found {len(files)} file(s):")
    for f in files:
        print(f"  - {os.path.basename(f)}")

    df = aggregate_files(files)
    if df.empty:
        print("No data to aggregate")
        sys.exit(1)

    overrides = {}
    if args.override_file:
        overrides.update(load_overrides_from_file(args.override_file))
    for override_str in args.override:
        col, val = parse_override(override_str)
        overrides[col] = val

    df = apply_overrides(df, overrides)

    if args.convert_types:
        print("\nConverting column types...")
        df = convert_types(df)

    df = filter_data(df, parse_filter_args(args.filter))

    if not args.no_dedup:
        if args.keep_duplicates == "none":
            existing_keys = [k for k in KEY_COLUMNS if k in df.columns]
            if existing_keys:
                df = df.drop_duplicates(subset=existing_keys, keep=False)
        else:
            df = remove_duplicates(df, keep=args.keep_duplicates)

    if not args.include_source and "_source_file" in df.columns:
        df = df.drop(columns=["_source_file"])

    print(f"\nFinal row count: {len(df)}")

    if args.output is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        args.output = f"aggregated_{timestamp}.csv"

    parent_dir = os.path.dirname(args.input_dir.rstrip(os.sep))
    output_path = os.path.join(parent_dir, args.output)

    if args.dry_run:
        print(f"\nDry run - would write to: {output_path}")
        print(f"Columns: {list(df.columns)}")
        print(f"\nFirst 5 rows:")
        print(df.head())
    else:
        df.to_csv(output_path, index=False)
        print(f"\nOutput written to: {output_path}")

    return output_path


# ============================================================================
# Validation functions
# ============================================================================

def load_file(filepath: str) -> pd.DataFrame:
    """Load CSV and do minimal type coercion."""
    df = pd.read_csv(filepath, dtype=str)
    print(f"Loaded {len(df)} rows, {len(df.columns)} columns from {filepath}")
    print(f"Columns: {list(df.columns)}\n")
    return df


def check_month_code_mismatch(df: pd.DataFrame) -> list[str]:
    """Check that the 3rd letter of RIC_Underlying matches the 3rd letter of Reference."""
    errors = []
    if "RIC_Underlying" not in df.columns:
        errors.append("[SKIP] RIC_Underlying column not found - cannot check month code")
        return errors
    if "Reference" not in df.columns:
        errors.append("[SKIP] Reference column not found - cannot check month code")
        return errors

    sub = df[["RIC_Underlying", "Reference"]].dropna()
    mask = (sub["RIC_Underlying"].str.len() >= 3) & (sub["Reference"].str.len() >= 3)
    sub = sub[mask]

    if sub.empty:
        errors.append("[SKIP] No rows with long enough RIC_Underlying and Reference to compare")
        return errors

    mismatch = sub[sub["RIC_Underlying"].str[2] != sub["Reference"].str[2]]
    if len(mismatch) > 0:
        n = len(mismatch)
        sample = mismatch.head(5)
        errors.append(
            f"[FAIL] {n} rows have month-code mismatch (3rd letter) "
            f"between RIC_Underlying and Reference"
        )
        for idx, row in sample.iterrows():
            errors.append(
                f"       Row {idx}: RIC_Underlying='{row['RIC_Underlying']}' "
                f"Reference='{row['Reference']}'"
            )
        if n > 5:
            errors.append(f"       ... and {n - 5} more")
    else:
        errors.append(f"[PASS] Month code (3rd letter) matches on all {len(sub)} checked rows")

    return errors


def check_numeric_column(df: pd.DataFrame, col: str, *,
                         allow_negative: bool = False,
                         min_val=None,
                         max_val=None) -> list[str]:
    """Validate that a column contains numeric data within bounds."""
    errors = []
    if col not in df.columns:
        errors.append(f"[SKIP] {col} column not found")
        return errors

    series = pd.to_numeric(df[col], errors="coerce")
    non_numeric = series.isna() & df[col].notna() & (df[col].str.strip() != "")
    n_bad = non_numeric.sum()
    if n_bad > 0:
        samples = df.loc[non_numeric, col].head(5).tolist()
        errors.append(f"[FAIL] {col}: {n_bad} non-numeric values, e.g. {samples}")
    else:
        errors.append(f"[PASS] {col}: all values are numeric")

    valid = series.dropna()
    if valid.empty:
        errors.append(f"[WARN] {col}: no valid numeric data to range-check")
        return errors

    if not allow_negative:
        neg = (valid < 0).sum()
        if neg > 0:
            errors.append(f"[FAIL] {col}: {neg} negative values found")

    if min_val is not None:
        below = (valid < min_val).sum()
        if below > 0:
            errors.append(f"[WARN] {col}: {below} values below {min_val} (min={valid.min():.6f})")

    if max_val is not None:
        above = (valid > max_val).sum()
        if above > 0:
            errors.append(f"[WARN] {col}: {above} values above {max_val} (max={valid.max():.6f})")

    return errors


def check_month_gaps(df: pd.DataFrame) -> list[str]:
    """Detect gaps in months between earliest and latest Spot_Date."""
    errors = []
    if "Spot_Date" not in df.columns:
        errors.append("[SKIP] Spot_Date column not found - cannot check month gaps")
        return errors

    dates = pd.to_datetime(df["Spot_Date"], errors="coerce").dropna()
    if dates.empty:
        errors.append("[WARN] No valid dates in Spot_Date")
        return errors

    start = dates.min()
    end = dates.max()
    errors.append(f"[INFO] Date range: {start.date()} to {end.date()}")

    present = set(zip(dates.dt.year, dates.dt.month))

    expected = set()
    current = start.replace(day=1)
    end_month = end.replace(day=1)
    while current <= end_month:
        expected.add((current.year, current.month))
        if current.month == 12:
            current = current.replace(year=current.year + 1, month=1)
        else:
            current = current.replace(month=current.month + 1)

    missing = sorted(expected - present)
    if missing:
        missing_strs = [f"{y}-{m:02d}" for y, m in missing]
        errors.append(f"[FAIL] Missing months in Spot_Date: {', '.join(missing_strs)}")
    else:
        errors.append(f"[PASS] No month gaps ({len(present)} months covered)")

    return errors


def check_missing_by_year(df: pd.DataFrame, col: str, date_col: str = "Spot_Date") -> list[str]:
    """Report missing/non-numeric values in a numeric column, broken down
    by year of date_col. Useful for spotting whole-year data gaps."""
    errors = []
    if col not in df.columns:
        errors.append(f"[SKIP] {col} column not found")
        return errors
    if date_col not in df.columns:
        errors.append(f"[SKIP] {date_col} column not found - cannot break down by year")
        return errors

    series = pd.to_numeric(df[col], errors="coerce")
    is_blank = df[col].isna() | (df[col].astype(str).str.strip() == "")
    is_missing = series.isna() | is_blank

    n_missing = int(is_missing.sum())
    total_rows = len(df)
    if n_missing == 0:
        errors.append(f"[PASS] {col}: no missing values across {total_rows:,} rows")
        return errors

    pct = 100 * n_missing / total_rows
    errors.append(f"[WARN] {col}: {n_missing:,} missing of {total_rows:,} ({pct:.1f}%)")

    dates = pd.to_datetime(df[date_col], errors="coerce")
    if not dates.notna().any():
        errors.append(f"       (no valid {date_col} values - cannot break down by year)")
        return errors

    yearly = pd.DataFrame({
        "year": dates.dt.year,
        "missing": is_missing,
    }).dropna(subset=["year"])
    grouped = yearly.groupby("year").agg(
        total=("missing", "count"),
        missing=("missing", "sum"),
    )
    grouped = grouped[grouped["missing"] > 0].sort_index()

    if grouped.empty:
        return errors

    errors.append(f"       Missing values by {date_col} year (missing / total = pct):")
    for year, row in grouped.iterrows():
        m = int(row["missing"])
        t = int(row["total"])
        pct_year = 100 * m / t if t > 0 else 0
        errors.append(f"         {int(year)}: {m:>10,} / {t:>10,} = {pct_year:5.1f}%")

    return errors


def check_maturity_month_gaps(df: pd.DataFrame) -> list[str]:
    """Detect missing months in Maturity between the first and last maturity date."""
    errors = []
    if "Maturity" not in df.columns:
        errors.append("[SKIP] Maturity column not found")
        return errors

    dates = pd.to_datetime(df["Maturity"], errors="coerce").dropna()
    if dates.empty:
        errors.append("[WARN] No valid dates in Maturity")
        return errors

    start = dates.min()
    end = dates.max()
    errors.append(f"[INFO] Maturity range: {start.date()} to {end.date()}")

    present = set(zip(dates.dt.year, dates.dt.month))

    expected = set()
    current = start.replace(day=1)
    end_month = end.replace(day=1)
    while current <= end_month:
        expected.add((current.year, current.month))
        if current.month == 12:
            current = current.replace(year=current.year + 1, month=1)
        else:
            current = current.replace(month=current.month + 1)

    missing = sorted(expected - present)
    if missing:
        by_year = {}
        for y, m in missing:
            by_year.setdefault(y, []).append(m)
        errors.append(f"[WARN] {len(missing)} missing month(s) in Maturity dates:")
        for y in sorted(by_year):
            months_str = ", ".join(f"{m:02d}" for m in by_year[y])
            errors.append(f"       {y}: {months_str}")
    else:
        errors.append(f"[PASS] No maturity month gaps ({len(present)} months covered)")

    return errors


def check_maturity_before_spot(df: pd.DataFrame) -> list[str]:
    """Check for rows where Maturity < Spot_Date (expired at observation)."""
    errors = []
    if "Spot_Date" not in df.columns or "Maturity" not in df.columns:
        errors.append("[SKIP] Spot_Date or Maturity column not found")
        return errors

    spot = pd.to_datetime(df["Spot_Date"], errors="coerce")
    mat = pd.to_datetime(df["Maturity"], errors="coerce")
    valid = spot.notna() & mat.notna()
    bad = (mat[valid] < spot[valid]).sum()
    if bad > 0:
        errors.append(f"[WARN] {bad} rows have Maturity before Spot_Date (already expired)")
    else:
        errors.append(f"[PASS] All Maturity dates are on or after Spot_Date")

    return errors


def check_missing_values(df: pd.DataFrame) -> list[str]:
    """Report missing values in critical columns."""
    errors = []
    for col in CRITICAL_COLUMNS:
        if col not in df.columns:
            errors.append(f"[WARN] Critical column '{col}' is missing from file")
            continue
        n_missing = df[col].isna().sum() + (df[col].astype(str).str.strip() == "").sum()
        if n_missing > 0:
            pct = 100 * n_missing / len(df)
            errors.append(f"[FAIL] {col}: {n_missing} missing/empty values ({pct:.1f}%)")
        else:
            errors.append(f"[PASS] {col}: no missing values")
    return errors


def check_duplicates(df: pd.DataFrame) -> list[str]:
    """Check for duplicate rows on key columns."""
    errors = []
    existing_keys = [k for k in KEY_COLUMNS if k in df.columns]
    if not existing_keys:
        errors.append("[SKIP] No key columns found for duplicate check")
        return errors

    dupes = df.duplicated(subset=existing_keys, keep=False).sum()
    if dupes > 0:
        errors.append(f"[WARN] {dupes} rows are duplicates on {existing_keys}")
    else:
        errors.append(f"[PASS] No duplicates on key columns")
    return errors


def check_type_values(df: pd.DataFrame) -> list[str]:
    """Check that Type column contains only expected values."""
    errors = []
    if "Type" not in df.columns:
        errors.append("[SKIP] Type column not found")
        return errors

    valid_types = {"CALL", "PUT", "C", "P", "Call", "Put"}
    actual = set(df["Type"].dropna().unique())
    unexpected = actual - valid_types
    if unexpected:
        errors.append(f"[WARN] Unexpected Type values: {unexpected}")
    else:
        errors.append(f"[PASS] Type values OK: {actual}")
    return errors


def run_validation(filepath: str, iv_max: float, spot_min: float, spot_max: float) -> int:
    """Run all checks and print report. Returns FAIL count."""
    df = load_file(filepath)

    print("=" * 70)
    print("VALIDATION REPORT")
    print("=" * 70)

    sections = [
        ("1. Month Code Mismatch (RIC_Underlying vs Reference)",
         check_month_code_mismatch(df)),
        ("2. Spot Data",
         check_numeric_column(df, "Spot", min_val=spot_min, max_val=spot_max)),
        ("3. Implied Volatility",
         check_numeric_column(df, "Implied_Volatility", min_val=0, max_val=iv_max)),
        ("4. Month Gaps in Spot_Date",
         check_month_gaps(df)),
        ("5. Maturity Month Gaps",
         check_maturity_month_gaps(df)),
        ("6. Maturity vs Spot_Date",
         check_maturity_before_spot(df)),
        ("7. Premium",
         check_numeric_column(df, "Premium")),
        ("8. Strike",
         check_numeric_column(df, "Strike", min_val=0)),
        ("9. Missing Values in Critical Columns",
         check_missing_values(df)),
        ("10. Duplicate Rows",
         check_duplicates(df)),
        ("11. Type Values",
         check_type_values(df)),
        ("12. Interest Rate Missing by Year",
         check_missing_by_year(df, "Interest_rate")),
        ("13. Implied Volatility Missing by Year",
         check_missing_by_year(df, "Implied_Volatility")),
    ]

    fail_count = 0
    warn_count = 0
    for title, results in sections:
        print(f"\n--- {title} ---")
        for line in results:
            print(f"  {line}")
            if line.startswith("[FAIL]"):
                fail_count += 1
            elif line.startswith("[WARN]"):
                warn_count += 1

    print("\n" + "=" * 70)
    print(f"SUMMARY: {fail_count} FAIL(s), {warn_count} WARNING(s)")
    if fail_count == 0 and warn_count == 0:
        print("All checks passed.")
    print("=" * 70)

    return fail_count


# ============================================================================
# CLI
# ============================================================================

def add_aggregate_args(parser: argparse.ArgumentParser) -> None:
    """Args used by both 'aggregate' and 'all' subcommands."""
    parser.add_argument(
        "--input-dir", "-i",
        default="..",
        help="Directory containing CSV files (default: parent directory)",
    )
    parser.add_argument(
        "--pattern", "-p",
        default="*_batch*.csv",
        help="Glob pattern to match CSV files (default: *_batch*.csv)",
    )
    parser.add_argument(
        "--output", "-o",
        default=None,
        help="Output file path (default: aggregated_YYYYMMDD_HHMMSS.csv)",
    )
    parser.add_argument(
        "--override",
        action="append",
        default=[],
        help="Override a field: --override 'column=value' (repeatable)",
    )
    parser.add_argument(
        "--override-file",
        default=None,
        help="JSON file containing field overrides",
    )
    parser.add_argument(
        "--filter",
        action="append",
        default=[],
        help="Filter data: --filter 'column>value' (repeatable)",
    )
    parser.add_argument(
        "--keep-duplicates",
        choices=["first", "last", "none"],
        default="last",
        help="Which duplicate to keep (default: last)",
    )
    parser.add_argument("--no-dedup", action="store_true", help="Skip duplicate removal")
    parser.add_argument("--include-source", action="store_true", help="Include source file column")
    parser.add_argument("--convert-types", action="store_true", help="Convert columns to numeric/date")
    parser.add_argument("--dry-run", action="store_true", help="Don't write the output file")


def add_validate_args(parser: argparse.ArgumentParser, with_input: bool = True) -> None:
    """Args used by both 'validate' and 'all' subcommands. The 'all' subcommand
    omits the positional input (it uses the aggregate output)."""
    if with_input:
        parser.add_argument("input", help="Path to the aggregated CSV file")
    parser.add_argument("--iv-max", type=float, default=5.0,
                        help="Max acceptable implied volatility (default: 5.0 = 500%%)")
    parser.add_argument("--spot-min", type=float, default=0,
                        help="Min acceptable spot price (default: 0)")
    parser.add_argument("--spot-max", type=float, default=1_000_000,
                        help="Max acceptable spot price (default: 1000000)")


def main():
    parser = argparse.ArgumentParser(
        description="Aggregate and validate VBA option-download CSVs",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Aggregate batch files
  python aggregate_and_validate.py aggregate --input-dir ../data --output combined.csv

  # Validate an aggregated CSV
  python aggregate_and_validate.py validate combined.csv --iv-max 10.0

  # Aggregate then validate the result in one shot
  python aggregate_and_validate.py all --input-dir ../data --output combined.csv

  # Override fields and apply filters during aggregation
  python aggregate_and_validate.py aggregate \\
      --input-dir ../data \\
      --override "ccy_pair=EUR/USD" --override "Lot_size=100" \\
      --filter "Premium>0"
        """,
    )

    subparsers = parser.add_subparsers(dest="command", required=True, metavar="COMMAND")

    p_agg = subparsers.add_parser("aggregate", help="Combine batch CSVs into one aggregated file")
    add_aggregate_args(p_agg)

    p_val = subparsers.add_parser("validate", help="Run quality checks on an aggregated CSV")
    add_validate_args(p_val, with_input=True)

    p_all = subparsers.add_parser("all", help="Aggregate then validate the result")
    add_aggregate_args(p_all)
    add_validate_args(p_all, with_input=False)

    args = parser.parse_args()

    if args.command == "aggregate":
        run_aggregate(args)
        return 0

    if args.command == "validate":
        return run_validation(args.input, args.iv_max, args.spot_min, args.spot_max)

    if args.command == "all":
        output_path = run_aggregate(args)
        if args.dry_run:
            print("\nSkipping validation (dry run did not write a file).")
            return 0
        print("\n" + "#" * 70)
        print("# Aggregation done — running validation on the output")
        print("#" * 70 + "\n")
        return run_validation(output_path, args.iv_max, args.spot_min, args.spot_max)

    return 0


if __name__ == "__main__":
    # Example IDE-debug invocation — uncomment and adjust as needed:
    #
    # sys.argv = ["aggregate_and_validate.py", "aggregate",
    #             "--input-dir", r"C:\path\to\batches",
    #             "--output", "downloaded_X_20260507.csv"]
    #
    # sys.argv = ["aggregate_and_validate.py", "validate",
    #             r"C:\path\to\aggregated.csv",
    #             "--iv-max", "5.0"]
    #
    # sys.argv = ["aggregate_and_validate.py", "all",
    #             "--input-dir", r"C:\path\to\batches",
    #             "--output", "downloaded_X_20260507.csv",
    #             "--iv-max", "5.0"]

    sys.exit(main())
