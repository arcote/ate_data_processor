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

Module use:
  from extract_ate_data import preview_scan, extract_records, write_results, scan
  summary = preview_scan("C:/path/to/data")     # fast layout/param/condition probe
  records = extract_records("C:/path/to/data")  # full extraction -> list of dicts
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
COND_RE = re.compile(r"_(COLD|HOT|ROOM)$", re.IGNORECASE)
KNOWN_CONDITIONS = ("COLD", "HOT", "ROOM")


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


def detect_condition(temp_folder: str) -> str:
    """Return 'COLD', 'ROOM', or 'HOT' from a folder name like 'Idss_EN5_COLD'."""
    upper = temp_folder.upper()
    for key in KNOWN_CONDITIONS:
        if key in upper:
            return key
    return "ROOM"


def detect_parameter(temp_folder: str) -> str:
    """Strip a trailing _COLD / _HOT / _ROOM from a folder name to get the parameter."""
    return COND_RE.sub("", temp_folder)


def is_acs_basic_folder(name: str) -> bool:
    """Match the ACS_BASIC_ container folder (tolerates the legacy ASC_ typo)."""
    u = name.upper()
    return u.startswith("ACS_BASIC_") or u.startswith("ASC_BASIC_")


def is_setup_folder(name: str) -> bool:
    return name.upper().startswith("SETUP_")


def iter_run_files(root: str, log=print):
    """
    Walk the directory tree and yield dicts describing each run CSV (without
    reading its contents) with keys:
      test_folder, temp_folder, run_number, filepath

    Supports two layouts automatically:

      Layout A (original):
        root / <test_folder> / ACS_BASIC_<date> / <temp_folder> / run#.csv

      Layout B (script sits immediately above ACS_BASIC_ folder):
        root / ACS_BASIC_<date> / <temp_folder> / run#.csv
        test_folder label is taken from the parent directory of root.
    """
    root = os.path.abspath(root)

    try:
        root_entries = sorted(os.scandir(root), key=lambda e: e.name)
    except (PermissionError, FileNotFoundError, NotADirectoryError) as exc:
        log(f"[ERROR] Cannot read root directory: {exc}")
        return

    # Auto-detect layout: does root itself contain ACS_BASIC_ folders?
    root_has_acs = any(e.is_dir() and is_acs_basic_folder(e.name) for e in root_entries)

    if root_has_acs:
        # Layout B — derive test_folder label from parent directory name
        test_folder = os.path.basename(os.path.dirname(root)) or os.path.basename(root)
        log(f"[INFO] Layout B detected — test_folder label: '{test_folder}'")
        blocks = [(test_folder, root_entries)]
    else:
        # Layout A — each subdirectory is a named test folder
        log("[INFO] Layout A detected — scanning named test folders")
        blocks = []
        for test_entry in root_entries:
            if not test_entry.is_dir():
                continue
            try:
                sub_entries = sorted(os.scandir(test_entry.path), key=lambda e: e.name)
            except PermissionError:
                log(f"[WARN] Cannot read {test_entry.path}, skipping.")
                continue
            blocks.append((test_entry.name, sub_entries))

    for test_folder, entries in blocks:

        acs_entries = [e for e in entries if e.is_dir() and is_acs_basic_folder(e.name)]

        for acs_entry in sorted(acs_entries, key=lambda e: e.name):

            # Handle double-nesting: ACS_BASIC_<date>/ACS_BASIC_<date>/...
            # If the first child that matches is itself another ACS_BASIC_ folder,
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
                log(f"[WARN] Cannot read {scan_path}, skipping.")
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
                    log(f"[WARN] Cannot read {temp_entry.path}, skipping.")
                    continue

                csv_files.sort(key=lambda e: int(find_run_number(e.name))
                               if find_run_number(e.name).isdigit() else 0)

                for csv_entry in csv_files:
                    yield {
                        "test_folder": test_folder,
                        "temp_folder": temp_folder,
                        "run_number":  find_run_number(csv_entry.name),
                        "filepath":    csv_entry.path,
                    }


def scan(root: str, log=print):
    """
    Walk the tree and yield a record per run CSV with keys:
      test_folder, temp_folder, run_number, extracted_value, error, filepath

    This reads each CSV to pull its extracted value; for a fast structural
    probe that does not open files, use ``preview_scan`` instead.
    """
    for entry in iter_run_files(root, log=log):
        value, error = extract_rightmost_bottom_value(entry["filepath"])
        yield {
            "test_folder":     entry["test_folder"],
            "temp_folder":     entry["temp_folder"],
            "run_number":      entry["run_number"],
            "extracted_value": value,
            "error":           error or "",
            "filepath":        entry["filepath"],
        }


def preview_scan(root: str, log=lambda *_: None):
    """
    Fast structural probe of the data folder — does NOT open any CSV.

    Returns a dict:
      {
        "layout":            "A" | "B" | "none",
        "test_folders":      [..],
        "temp_folders":      [..],
        "parameters":        [..],
        "conditions":        [..],
        "file_count":        int,
        "files_by_condition": {cond: count, ...},
      }
    """
    test_folders, temp_folders = set(), set()
    parameters, conditions = set(), set()
    files_by_condition = {}
    file_count = 0

    for entry in iter_run_files(root, log=log):
        file_count += 1
        test_folders.add(entry["test_folder"])
        temp_folders.add(entry["temp_folder"])
        param = detect_parameter(entry["temp_folder"])
        cond = detect_condition(entry["temp_folder"])
        parameters.add(param)
        conditions.add(cond)
        files_by_condition[cond] = files_by_condition.get(cond, 0) + 1

    # Re-derive layout the same way iter_run_files does, for reporting.
    layout = "none"
    root_abs = os.path.abspath(root)
    try:
        entries = list(os.scandir(root_abs))
        if any(e.is_dir() and is_acs_basic_folder(e.name) for e in entries):
            layout = "B"
        elif file_count:
            layout = "A"
    except OSError:
        layout = "none"

    cond_order = [c for c in KNOWN_CONDITIONS if c in conditions]
    cond_order += sorted(c for c in conditions if c not in KNOWN_CONDITIONS)

    return {
        "layout":             layout,
        "test_folders":       sorted(test_folders),
        "temp_folders":       sorted(temp_folders),
        "parameters":         sorted(parameters),
        "conditions":         cond_order,
        "file_count":         file_count,
        "files_by_condition": files_by_condition,
    }


def extract_records(root, log=print, progress_callback=None, total=None):
    """
    Full extraction. Returns a list of record dicts (see ``scan``).

    progress_callback(done, total) is invoked after each file when supplied.
    Pass ``total`` (e.g. from ``preview_scan``) for meaningful progress values.
    """
    records = []
    for i, record in enumerate(scan(root, log=log), 1):
        records.append(record)
        if progress_callback is not None:
            progress_callback(i, total)
    return records


def write_results(records, output, include_errors=False, include_filepath=False):
    """Write extraction records to a CSV. Returns (ok_count, err_count)."""
    fieldnames = ["test_folder", "temp_folder", "run_number", "extracted_value"]
    if include_errors:
        fieldnames.append("error")
    if include_filepath:
        fieldnames.append("filepath")

    ok_count = err_count = 0
    rows = []
    for record in records:
        row = {k: record.get(k) for k in
               ("test_folder", "temp_folder", "run_number", "extracted_value")}
        if include_errors:
            row["error"] = record.get("error", "")
        if include_filepath:
            row["filepath"] = record.get("filepath", "")
        rows.append(row)
        if record.get("error"):
            err_count += 1
        else:
            ok_count += 1

    output = os.path.expanduser(output)
    out_dir = os.path.dirname(os.path.abspath(output))
    os.makedirs(out_dir, exist_ok=True)
    with open(output, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    return ok_count, err_count


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

    records = []
    for record in scan(root):
        records.append(record)
        if record["error"]:
            print(f"  [WARN] {record['filepath']}: {record['error']}")

    if not records:
        print("[INFO] No CSV files found. Check your --root path and folder structure.")
        return

    ok_count, err_count = write_results(
        records, output,
        include_errors=args.include_errors,
        include_filepath=args.include_filepath,
    )

    print()
    print(f"Done. {ok_count} files extracted, {err_count} errors.")
    print(f"Results written to: {os.path.abspath(output)}")


if __name__ == "__main__":
    main()
