"""
extract_ate_data.py
-------------------
Recursively walks a root directory to extract the last value from the
rightmost column of ATE test run CSV files.

Directory structure expected:
  <root>/
    <test_folder>/          e.g. "3rd test", "4th test", "5th test"
      ACS_BASIC_<date>/     (only folders starting with "ASC_BASIC_ / ACS_BASIC_")
        <temp_folder>/      e.g. "Idss_EN5_COLD", "Bvces_EN6_HOT"
                            (excludes folders starting with "SETUP_")
          <name>_run<N>.csv

Output CSV columns:
  test_folder | temp_folder | run_number | extracted_value

Usage:
  python extract_ate_data.py --root "C:/path/to/data" --output results.csv

  Or edit ROOT_DIR and OUTPUT_FILE below and run without arguments.
"""

import os
import re
import csv
import argparse
import pandas as pd


# ── Edit these if you don't want to use command-line arguments ──────────────
ROOT_DIR    = r"."          # Top-level folder containing "3rd test", etc.
OUTPUT_FILE = r"results.csv"
# ────────────────────────────────────────────────────────────────────────────

RUN_RE = re.compile(r"_run(\d+)", re.IGNORECASE)


def extract_rightmost_bottom_value(filepath: str):
    """
    Read a CSV (no assumed header) and return the last non-NaN value in the
    rightmost column that contains any data.

    Returns (value, error_string). On success, error_string is None.
    On failure, value is None and error_string describes what went wrong.
    """
    try:
        df = pd.read_csv(filepath, header=None, low_memory=False)
    except Exception as exc:
        return None, f"read error: {exc}"

    if df.empty:
        return None, "empty file"

    # Walk columns right-to-left; stop at the first one with any non-NaN data
    for col in reversed(df.columns):
        col_series = df[col].dropna()
        if not col_series.empty:
            return col_series.iloc[-1], None

    return None, "all columns empty"


def find_run_number(filename: str):
    """Extract the integer run number from a filename like 'Idss_EN5_COLD_run23.csv'."""
    match = RUN_RE.search(filename)
    return match.group(1) if match else "unknown"


def is_acs_basic_folder(name: str) -> bool:
    u = name.upper()
    return u.startswith("ASC_BASIC_ / ACS_BASIC_") or u.startswith("ASC_BASIC_")


def is_setup_folder(name: str) -> bool:
    return name.upper().startswith("SETUP_")


def scan(root: str):
    """
    Walk the directory tree and yield dicts with keys:
      test_folder, temp_folder, run_number, extracted_value, error, filepath

    Supports two layouts automatically:

      Layout A (original):
        root / <test_folder> / ACS_BASIC_<date> / <temp_folder> / run#.csv

      Layout B (script sits immediately above ACS_BASIC_ folder):
        root / ACS_BASIC_<date> / <temp_folder> / run#.csv
        test_folder label is taken from the parent directory of root.

    Run from anywhere — just point --root at the folder that either contains
    the named test folders (A) or the ACS_BASIC_ folder directly (B).
    """
    root = os.path.abspath(root)

    try:
        root_entries = sorted(os.scandir(root), key=lambda e: e.name)
    except PermissionError as exc:
        print(f"[ERROR] Cannot read root directory: {exc}")
        return

    # Auto-detect layout: does root itself contain ACS_BASIC_ folders?
    root_has_acs = any(e.is_dir() and is_acs_basic_folder(e.name) for e in root_entries)

    if root_has_acs:
        # Layout B — derive test_folder label from parent directory name
        test_folder = os.path.basename(os.path.dirname(root)) or os.path.basename(root)
        print(f"[INFO] Layout B detected — test_folder label: '{test_folder}'")
        blocks = [(test_folder, root_entries)]
    else:
        # Layout A — each subdirectory is a named test folder
        print("[INFO] Layout A detected — scanning named test folders")
        blocks = []
        for test_entry in root_entries:
            if not test_entry.is_dir():
                continue
            try:
                sub_entries = sorted(os.scandir(test_entry.path), key=lambda e: e.name)
            except PermissionError:
                print(f"[WARN] Cannot read {test_entry.path}, skipping.")
                continue
            blocks.append((test_entry.name, sub_entries))

    for test_folder, entries in blocks:

        acs_entries = [e for e in entries if e.is_dir() and is_acs_basic_folder(e.name)]

        for acs_entry in sorted(acs_entries, key=lambda e: e.name):

            # Handle double-nesting: ASC_BASIC_<date>/ASC_BASIC_<date>/...
            # If the first child that matches is itself another ASC_BASIC_ folder,
            # step down one level automatically.
            try:
                acs_children = [
                    e for e in os.scandir(acs_entry.path)
                    if e.is_dir() and is_acs_basic_folder(e.name)
                ]
            except PermissionError:
                acs_children = []

            scan_path = acs_children[0].path if acs_children else acs_entry.path

            # Temperature/condition folders (exclude SETUP_*)
            try:
                temp_entries = [
                    e for e in os.scandir(scan_path)
                    if e.is_dir() and not is_setup_folder(e.name)
                ]
            except PermissionError:
                print(f"[WARN] Cannot read {scan_path}, skipping.")
                continue

            for temp_entry in sorted(temp_entries, key=lambda e: e.name):
                temp_folder = temp_entry.name

                # Run CSV files
                try:
                    csv_files = [
                        e for e in os.scandir(temp_entry.path)
                        if e.is_file() and e.name.lower().endswith(".csv")
                    ]
                except PermissionError:
                    print(f"[WARN] Cannot read {temp_entry.path}, skipping.")
                    continue

                csv_files.sort(key=lambda e: int(find_run_number(e.name))
                               if find_run_number(e.name).isdigit() else 0)

                for csv_entry in csv_files:
                    run_number = find_run_number(csv_entry.name)
                    value, error = extract_rightmost_bottom_value(csv_entry.path)

                    yield {
                        "test_folder":      test_folder,
                        "temp_folder":      temp_folder,
                        "run_number":       run_number,
                        "extracted_value":  value,
                        "error":            error or "",
                        "filepath":         csv_entry.path,
                    }


def main():
    parser = argparse.ArgumentParser(
        description="Extract last value from rightmost column of ATE run CSVs."
    )
    parser.add_argument(
        "--root", default=ROOT_DIR,
        help=f"Root data directory (default: {ROOT_DIR!r})"
    )
    parser.add_argument(
        "--output", default=OUTPUT_FILE,
        help=f"Output CSV path (default: {OUTPUT_FILE!r})"
    )
    parser.add_argument(
        "--include-errors", action="store_true",
        help="Include a 5th column with any read errors (useful for debugging)"
    )
    parser.add_argument(
        "--include-filepath", action="store_true",
        help="Include a filepath column for traceability"
    )
    args = parser.parse_args()

    root    = os.path.expanduser(args.root)
    output  = os.path.expanduser(args.output)

    print(f"Scanning: {os.path.abspath(root)}")
    print(f"Output  : {os.path.abspath(output)}")
    print()

    rows      = []
    ok_count  = 0
    err_count = 0

    for record in scan(root):
        row = {
            "test_folder":     record["test_folder"],
            "temp_folder":     record["temp_folder"],
            "run_number":      record["run_number"],
            "extracted_value": record["extracted_value"],
        }
        if args.include_errors:
            row["error"] = record["error"]
        if args.include_filepath:
            row["filepath"] = record["filepath"]

        rows.append(row)

        if record["error"]:
            err_count += 1
            print(f"  [WARN] {record['filepath']}: {record['error']}")
        else:
            ok_count += 1

    if not rows:
        print("[INFO] No CSV files found. Check your --root path and folder structure.")
        return

    fieldnames = ["test_folder", "temp_folder", "run_number", "extracted_value"]
    if args.include_errors:
        fieldnames.append("error")
    if args.include_filepath:
        fieldnames.append("filepath")

    with open(output, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    print()
    print(f"Done. {ok_count} files extracted, {err_count} errors.")
    print(f"Results written to: {os.path.abspath(output)}")


if __name__ == "__main__":
    main()
