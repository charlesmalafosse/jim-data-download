#!/usr/bin/env python3
"""
Validate an aggregated CSV file produced by aggregate_files.py.

Checks performed:
  1. RIC_Underlying 3rd letter vs Reference 3rd letter (month code mismatch)
  2. Spot is numeric and within a reasonable range
  3. Implied Volatility is numeric, non-negative, and within bounds
  4. Month gaps between earliest and latest Spot_Date
  5. Maturity before Spot_Date (stale / wrong dates)
  6. Premium is numeric and non-negative
  7. Missing values in critical columns
  8. Duplicate rows on key columns
  9. Strike sanity (numeric, positive)

Usage:
    python validate_aggregated.py <input_csv>
    python validate_aggregated.py <input_csv> --iv-max 5.0 --spot-min 0 --spot-max 100000
"""

import argparse
import sys

import pandas as pd


# Columns that should never be empty
CRITICAL_COLUMNS = [
    "Spot_Date", "Ticker", "Maturity", "Spot", "Strike", "Type",
    "Implied_Volatility", "Premium", "Reference",
]

KEY_COLUMNS = ["Spot_Date", "Ticker", "Maturity", "Strike", "Type"]


def load_file(filepath: str) -> pd.DataFrame:
    """Load CSV and do minimal type coercion."""
    df = pd.read_csv(filepath, dtype=str)
    print(f"Loaded {len(df)} rows, {len(df.columns)} columns from {filepath}")
    print(f"Columns: {list(df.columns)}\n")
    return df


# ---------------------------------------------------------------------------
# Individual checks
# ---------------------------------------------------------------------------

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
    # Only check rows where both values are long enough
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

    # Build set of (year, month) present in data
    present = set(zip(dates.dt.year, dates.dt.month))

    # Build expected set of (year, month) from start to end
    expected = set()
    current = start.replace(day=1)
    end_month = end.replace(day=1)
    while current <= end_month:
        expected.add((current.year, current.month))
        # advance one month
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


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def run_validation(filepath: str, iv_max: float, spot_min: float, spot_max: float):
    """Run all checks and print report."""
    df = load_file(filepath)

    all_results = []

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
        ("5. Maturity vs Spot_Date",
         check_maturity_before_spot(df)),
        ("6. Premium",
         check_numeric_column(df, "Premium")),
        ("7. Strike",
         check_numeric_column(df, "Strike", min_val=0)),
        ("8. Missing Values in Critical Columns",
         check_missing_values(df)),
        ("9. Duplicate Rows",
         check_duplicates(df)),
        ("10. Type Values",
         check_type_values(df)),
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


def main():
    parser = argparse.ArgumentParser(
        description="Validate an aggregated options CSV file",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python validate_aggregated.py aggregated_20260401.csv
  python validate_aggregated.py data.csv --iv-max 10.0
  python validate_aggregated.py data.csv --spot-min 50 --spot-max 200
        """
    )
    parser.add_argument("input", help="Path to the aggregated CSV file")
    parser.add_argument("--iv-max", type=float, default=5.0,
                        help="Max acceptable implied volatility (default: 5.0 = 500%%)")
    parser.add_argument("--spot-min", type=float, default=0,
                        help="Min acceptable spot price (default: 0)")
    parser.add_argument("--spot-max", type=float, default=1_000_000,
                        help="Max acceptable spot price (default: 1000000)")

    args = parser.parse_args()
    fails = run_validation(args.input, args.iv_max, args.spot_min, args.spot_max)
    return fails


if __name__ == "__main__":
    sys.argv = [
        "validate_aggregated.py",
        "C:\\Users\\charl\\Dropbox\\MARINER\\DOWNLOAD_FILES\\0 - TEMPLATE\\DOWNLOAD_MONTHLY\\python_scripts\\downloaded_SUGAR_20260309.csv",
        "--iv-max", "5.0",
        "--spot-min", "0",
        "--spot-max", "1000000",
    ]
    main()
