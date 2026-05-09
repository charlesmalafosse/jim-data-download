#!/usr/bin/env python3
"""
Fix the futures month code in the Reference (Bloomberg) column using
RIC_Underlying as the source of truth.

For every row where Reference's month code (F,G,H,J,K,M,N,Q,U,V,X,Z) differs
from RIC_Underlying's, the script rewrites Reference's letter to match
RIC_Underlying's, preserving the root ticker, year digits, and Bloomberg
suffix (' Comdty', ' Index', ...).

Examples:
    RIC_Underlying='LCOZ6'      Reference='COH6 Comdty'      -> 'COZ6 Comdty'
    RIC_Underlying='LCOZ6^Z26'  Reference='COH6 Comdty'      -> 'COZ6 Comdty'
    RIC_Underlying='FLGZ6'      Reference='RXZ6 Comdty'      -> unchanged
    RIC_Underlying='0#LCOc1'    Reference='CO1 Comdty'       -> unchanged (no code)

By default the script previews how many rows would change (with samples)
and then asks for confirmation before writing the output CSV. Pass --yes
to skip the prompt for non-interactive runs, or --dry-run to preview only.

Usage:
    python fix_reference_month_code.py <input_csv>            # preview + prompt
    python fix_reference_month_code.py <input_csv> -y         # apply, no prompt
    python fix_reference_month_code.py <input_csv> -o fixed.csv
    python fix_reference_month_code.py <input_csv> --dry-run  # preview only
"""

import argparse
import os
import re
import sys

import pandas as pd


# Futures month codes: F=Jan G=Feb H=Mar J=Apr K=May M=Jun
#                     N=Jul Q=Aug U=Sep V=Oct X=Nov Z=Dec
_MONTH_CODE_RE = re.compile(r"([FGHJKMNQUVXZ])(\d{1,2})$")


def extract_futures_month_code(symbol) -> str | None:
    """Return the month-code letter from a futures-style symbol, regardless of
    root-ticker length. Strips '^MMYY' expired-RIC suffix and Bloomberg-style
    trailing tokens (' Comdty', ' Index'). Returns None for continuous
    contracts or unparseable inputs."""
    if not isinstance(symbol, str):
        return None
    head = symbol.strip().split(None, 1)[0]
    caret = head.find("^")
    if caret >= 0:
        head = head[:caret]
    if not head:
        return None
    m = _MONTH_CODE_RE.search(head)
    return m.group(1) if m else None


def rewrite_month_code(reference: str, target_code: str) -> str | None:
    """Return Reference with its month-code letter replaced by target_code,
    preserving the root, year digits, and any space-separated Bloomberg
    suffix. Returns None if Reference has no parseable month code."""
    if not isinstance(reference, str):
        return None
    raw = reference.strip()
    if not raw:
        return None
    parts = raw.split(None, 1)
    head = parts[0]
    tail = " " + parts[1] if len(parts) > 1 else ""
    m = _MONTH_CODE_RE.search(head)
    if not m:
        return None
    new_head = head[: m.start(1)] + target_code + head[m.end(1):]
    return new_head + tail


def find_changes(df: pd.DataFrame) -> tuple[list[tuple], int, int]:
    """Scan df without mutating it. Return (changes, ok, skipped) where
    changes is a list of (idx, original_ref, new_ref, ric_underlying)
    tuples for rows whose Reference month code differs from
    RIC_Underlying's and where a rewrite is possible."""
    if "RIC_Underlying" not in df.columns or "Reference" not in df.columns:
        raise ValueError("Input file must contain RIC_Underlying and Reference columns")

    ric_code = df["RIC_Underlying"].apply(extract_futures_month_code)
    ref_code = df["Reference"].apply(extract_futures_month_code)

    parseable = ric_code.notna() & ref_code.notna()
    mismatch = parseable & (ric_code != ref_code)

    changes: list[tuple] = []
    for idx in df.index[mismatch]:
        original = df.at[idx, "Reference"]
        target = ric_code.at[idx]
        new_ref = rewrite_month_code(original, target)
        if new_ref is None or new_ref == original:
            continue
        changes.append((idx, original, new_ref, df.at[idx, "RIC_Underlying"]))

    skipped = int((~parseable).sum())
    ok = int((parseable & (ric_code == ref_code)).sum())
    return changes, ok, skipped


def apply_changes(df: pd.DataFrame, changes: list[tuple]) -> None:
    """Apply previously-computed changes to df['Reference'] in place."""
    for idx, _original, new_ref, _ric in changes:
        df.at[idx, "Reference"] = new_ref


def confirm(prompt: str) -> bool:
    """Read y/N from stdin. Defaults to No on empty/EOF/anything else."""
    try:
        ans = input(prompt).strip().lower()
    except EOFError:
        return False
    return ans in ("y", "yes")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Fix Reference month code using RIC_Underlying as source of truth.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("input", help="Path to the input CSV file")
    parser.add_argument(
        "-o", "--output",
        help="Output CSV path (default: <input>_fixed.csv next to the input)",
    )
    parser.add_argument(
        "-y", "--yes", action="store_true",
        help="Skip the confirmation prompt and apply the fix",
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Print what would change but don't write an output file",
    )
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print(f"ERROR: input file not found: {args.input}", file=sys.stderr)
        return 2

    df = pd.read_csv(args.input, dtype=str)
    print(f"Loaded {len(df):,} rows from {args.input}")

    try:
        changes, ok, skipped = find_changes(df)
    except ValueError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    would_fix = len(changes)

    print(f"\nMatched (no change):     {ok:,}")
    print(f"Would fix (month code):  {would_fix:,}")
    print(f"Skipped (no parseable):  {skipped:,}")

    if changes:
        print("\nSample rewrites (first 5):")
        for idx, before, after, ric in changes[:5]:
            print(f"  Row {idx}: RIC_Underlying='{ric}'  '{before}' -> '{after}'")

    if args.dry_run:
        print("\nDry run — no file written.")
        return 0

    if would_fix == 0:
        print("\nNothing to fix — no output written.")
        return 0

    if not args.yes:
        if not confirm(f"\nProceed and write fixed CSV with {would_fix:,} change(s)? [y/N]: "):
            print("Aborted — no file written.")
            return 0

    apply_changes(df, changes)

    out_path = args.output
    if not out_path:
        base, ext = os.path.splitext(args.input)
        out_path = f"{base}_fixed{ext or '.csv'}"

    df.to_csv(out_path, index=False)
    print(f"\nWrote {out_path}")
    return 0


if __name__ == "__main__":
    # Hardcode args here so you can just hit Run in the IDE without retyping
    # the path. Comment out / clear this block to use real CLI args instead.
    sys.argv = [
        "fix_reference_month_code.py",
        r"C:\Users\charl\Dropbox\MARINER\DOWNLOAD_FILES\0 - TEMPLATE\DOWNLOAD_MONTHLY\python_scripts\downloaded_SUGAR_20260309.csv",
        # "-y",          # uncomment to skip the confirmation prompt
        # "--dry-run",   # uncomment to preview only, no write
    ]
    sys.exit(main())
