"""Validate that downloaded_SUGAR_20260309_sql.csv matches the SQL schema
expected by upload_to_mysql.py: correct headers in order, datetime format,
parseable numerics, and no sentinel/error strings."""

import csv
import re
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

from convert_to_sql_csv import (
    DATETIME_COLUMNS,
    NULL_SENTINELS,
    NUMERIC_COLUMNS,
    SQL_COLUMNS,
    SQL_COLUMNS_WITHOUT_SOURCE,
)

OUTPUT = Path(r"c:\Users\charl\Dropbox\MARINER\DOWNLOAD_FILES\0 - TEMPLATE\DOWNLOAD_MONTHLY\python_scripts\downloaded_SUGAR_20260309_sql.csv")
INPUT = Path(r"c:\Users\charl\Dropbox\MARINER\DOWNLOAD_FILES\0 - TEMPLATE\DOWNLOAD_MONTHLY\python_scripts\downloaded_SUGAR_20260309.csv")

DATETIME_RE = re.compile(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$")
NUMERIC_RE = re.compile(r"^-?\d+(\.\d+)?([eE][+-]?\d+)?$")
NULL_SENTINELS_UPPER = {s.upper() for s in NULL_SENTINELS if s}

errors = []
warnings = []

with open(OUTPUT, "r", newline="", encoding="utf-8") as f:
    reader = csv.reader(f)
    header = next(reader)

    if header != SQL_COLUMNS:
        errors.append(f"header mismatch:\n  got:      {header}\n  expected: {SQL_COLUMNS}")

    col_idx = {name: i for i, name in enumerate(header)}
    null_count = Counter()
    bad_datetime = Counter()
    bad_numeric = Counter()
    sentinel_hits = Counter()
    always_null_source = Counter()
    row_count = 0

    for row in reader:
        row_count += 1
        if len(row) != len(header):
            errors.append(f"row {row_count}: {len(row)} fields, expected {len(header)}")
            if len(errors) > 5:
                break
            continue

        for col, i in col_idx.items():
            val = row[i]
            if val == "":
                null_count[col] += 1
                if col in SQL_COLUMNS_WITHOUT_SOURCE:
                    always_null_source[col] += 1
                continue
            if val.upper() in NULL_SENTINELS_UPPER:
                sentinel_hits[col] += 1
                continue
            if col in DATETIME_COLUMNS:
                if not DATETIME_RE.match(val):
                    bad_datetime[col] += 1
                else:
                    try:
                        datetime.strptime(val, "%Y-%m-%d %H:%M:%S")
                    except ValueError:
                        bad_datetime[col] += 1
            elif col in NUMERIC_COLUMNS:
                if not NUMERIC_RE.match(val):
                    bad_numeric[col] += 1
                else:
                    try:
                        float(val)
                    except ValueError:
                        bad_numeric[col] += 1

# Cross-check row count vs input
with open(INPUT, "r", encoding="utf-8") as f:
    input_rows = sum(1 for _ in f) - 1

print("=" * 60)
print("VALIDATION REPORT")
print("=" * 60)
print(f"Output file:      {OUTPUT.name}")
print(f"Header columns:   {len(header)}")
print(f"Data rows:        {row_count:,}")
print(f"Input data rows:  {input_rows:,}")
print(f"Row count match:  {row_count == input_rows}")
print()

print("Header order check: ", "OK" if header == SQL_COLUMNS else "FAIL")
print()

print("NULLs per column (top 10):")
for col, n in null_count.most_common(10):
    marker = "  (always-NULL source)" if col in SQL_COLUMNS_WITHOUT_SOURCE else ""
    print(f"  {col:25s} {n:>10,}{marker}")
print()

for col in SQL_COLUMNS_WITHOUT_SOURCE:
    if null_count.get(col, 0) != row_count:
        errors.append(
            f"column {col} should always be NULL but has "
            f"{row_count - null_count.get(col, 0)} non-null values"
        )
    else:
        print(f"  {col}: 100% NULL (correct, no CSV source)")
print()

if bad_datetime:
    errors.append(f"malformed datetime values: {dict(bad_datetime)}")
else:
    print("Datetime format: OK (all non-null values match YYYY-MM-DD HH:MM:SS)")

if bad_numeric:
    errors.append(f"malformed numeric values: {dict(bad_numeric)}")
else:
    print("Numeric format:  OK (all non-null values parse as float)")

if sentinel_hits:
    errors.append(f"sentinel/error strings leaked into output: {dict(sentinel_hits)}")
else:
    print("Sentinel check:  OK (no NA/#N/A/#VALUE!/... in output)")

print()
print("=" * 60)
if errors:
    print(f"RESULT: FAIL ({len(errors)} error(s))")
    for e in errors:
        print("  -", e)
    sys.exit(1)
else:
    print("RESULT: PASS — output CSV matches SQL schema expectations")
