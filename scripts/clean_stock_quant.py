#!/usr/bin/env python3
"""
clean_stock_quant.py

Cleans an ERP "Stock Quant" CSV export and flags data-quality abnormalities.

Rules implemented (locked in with the business owner on 2026-08-10):

For lines that have a Lot/Serial Number:
  - ERP Duplicate line       : same (Product, Lot) has Quantity > 0 on two or more
                                lines at the same time, OR the exact same
                                (Product, Lot, Location) appears more than once.
                                (Historical zero-qty ledger lines from the item
                                moving through other locations are NOT duplicates.)
  - Has Special Character     : Lot is not purely alphanumeric (catches punctuation,
                                 spaces, and invisible/control characters too).
  - Imei Length Difference    : len(Lot) != the most common Lot length for that
                                 Product (the "mode length").
  - Alphanumeric IMEI         : Lot contains a letter, but this Product's lots are
                                 normally numeric-only (IMEI-style). Products whose
                                 lots are normally letter-based (serial-number style,
                                 e.g. "LIGHTER 1") are exempt from this check.

For every line:
  - Negative Qty              : Quantity < 0.
  - Owner Mismatch            : Location ends with OWNER_CHECK_SUFFIX (default
                                 "/Stock") AND Owner is not exactly VALID_OWNER.
                                 Scoped to on-hand warehouse locations on purpose —
                                 anywhere else (Customer, Teknisi, Refurbishment,
                                 ...) a different owner is expected, not abnormal.
  - Missing Imei               : this Product tracks lots elsewhere in the file
                                 (i.e. it has a mode length) but this line's Lot is
                                 blank.

The final output is filtered to locations ending in LOCATION_SUFFIX (default
"/Stock") only, with an added "Abnormality" column (comma-separated flags, empty
if none). Abnormalities themselves are computed against the WHOLE file first
(duplicate detection and mode length need full-file context), then the result is
filtered down to the */Stock scope. Pass --location-suffix "" to keep every
location (used by /stock-opname-teknisi, which needs the full ERP ledger, not
just on-hand */Stock rows) — OWNER_CHECK_SUFFIX stays "/Stock" independently of
this, so Owner Mismatch keeps meaning "on-hand stock owned by someone else"
rather than firing on every customer/technician-owned historical row.

Input can be the raw ERP export as either .csv or .xlsx (auto-detected from the
file extension) — the ERP's "Reporting -> Locations" export can be downloaded
as either. Output is always written as .csv.

Usage:
    python clean_stock_quant.py --input "path/to/Stock Quant.csv"
    python clean_stock_quant.py --input "path/to/Stock Quant.xlsx"
    python clean_stock_quant.py --input in.csv --output out.csv --owner "PT Otto Menara Globalindo" --location-suffix /Stock
"""

import argparse
import csv
import json
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

DEFAULT_OWNER = "PT Otto Menara Globalindo"
DEFAULT_LOCATION_SUFFIX = "/Stock"

ALNUM_RE = re.compile(r"^[A-Za-z0-9]+$")

REQUIRED_COLUMNS = [
    "Product",
    "Location",
    "Lot/Serial Number",
    "Quantity",
    "Owner",
]


def parse_args():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--input", required=True, help="Path to the raw Stock Quant CSV export")
    p.add_argument("--output", default=None, help="Path to write the cleaned CSV (default: <input>_cleaned.csv next to the input file)")
    p.add_argument("--owner", default=DEFAULT_OWNER, help=f"Owner string considered valid (default: {DEFAULT_OWNER!r})")
    p.add_argument("--location-suffix", default=DEFAULT_LOCATION_SUFFIX, help=f"Only rows whose Location ends with this suffix are kept in the output (default: {DEFAULT_LOCATION_SUFFIX!r}). Pass \"\" to keep every location.")
    p.add_argument("--owner-check-suffix", default=DEFAULT_LOCATION_SUFFIX, help=f"Owner Mismatch is only checked on rows whose Location ends with this suffix (default: {DEFAULT_LOCATION_SUFFIX!r}), independent of --location-suffix.")
    return p.parse_args()


def resolve_output_path(input_path: Path, requested_output: str | None) -> Path:
    if requested_output:
        return Path(requested_output)
    return input_path.with_name(input_path.stem + "_cleaned.csv")


def _cell_to_str(value) -> str:
    """Render an xlsx cell as the same kind of string csv.DictReader would give us."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    return str(value)


def read_input_rows(input_path: Path):
    """Read the raw Stock Quant export as a list of str->str dict rows, from
    either .csv or .xlsx. Returns (fieldnames, rows)."""
    if input_path.suffix.lower() == ".xlsx":
        import openpyxl

        wb = openpyxl.load_workbook(input_path, data_only=True, read_only=True)
        ws = wb.worksheets[0]
        row_iter = ws.iter_rows(values_only=True)
        fieldnames = [str(c).strip() if c is not None else "" for c in next(row_iter)]
        rows = []
        for raw_row in row_iter:
            if all(v is None for v in raw_row):
                continue
            row = {fieldnames[i]: _cell_to_str(v) for i, v in enumerate(raw_row) if i < len(fieldnames)}
            rows.append(row)
        return fieldnames, rows
    else:
        with open(input_path, encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            fieldnames = reader.fieldnames or []
            rows = list(reader)
        return fieldnames, rows


def write_with_fallback(rows, fieldnames, out_path: Path) -> Path:
    """Try to write to out_path; if it's locked (e.g. open in Excel), fall back
    to out_path_v2.csv, _v3.csv, etc. instead of failing outright."""
    candidate = out_path
    suffix_n = 2
    while True:
        try:
            with open(candidate, "w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(rows)
            return candidate
        except PermissionError:
            candidate = out_path.with_name(f"{out_path.stem}_v{suffix_n}{out_path.suffix}")
            suffix_n += 1
            if suffix_n > 20:
                raise


def main():
    args = parse_args()
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"ERROR: input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    valid_owner = args.owner
    location_suffix = args.location_suffix
    owner_check_suffix = args.owner_check_suffix

    fieldnames, rows = read_input_rows(input_path)
    missing = [c for c in REQUIRED_COLUMNS if c not in fieldnames]
    if missing:
        print(f"ERROR: input file is missing expected column(s): {missing}. Found columns: {fieldnames}", file=sys.stderr)
        sys.exit(1)

    # ---- Pass 1: per-product stats over the WHOLE file ----
    length_counter = defaultdict(Counter)          # product -> {length: count}
    numeric_style_counter = defaultdict(Counter)    # product -> {"numeric"/"alnum": count}
    group_members = defaultdict(list)               # (product, lot) -> [(idx, location, qty)]

    for idx, row in enumerate(rows):
        product = row["Product"]
        lot = row["Lot/Serial Number"].strip()
        if lot:
            length_counter[product][len(lot)] += 1
            numeric_style_counter[product]["alnum" if any(ch.isalpha() for ch in lot) else "numeric"] += 1
            try:
                qty = float(row["Quantity"])
            except ValueError:
                qty = 0.0
            group_members[(product, lot)].append((idx, row["Location"], qty))

    mode_length = {p: c.most_common(1)[0][0] for p, c in length_counter.items()}

    numeric_type_products = {
        product
        for product, c in numeric_style_counter.items()
        if c.get("numeric", 0) >= c.get("alnum", 0)
    }

    # ---- Pass 1b: duplicate detection ----
    exact_dup_key_count = Counter()
    for (product, lot), members in group_members.items():
        for _, location, _ in members:
            exact_dup_key_count[(product, lot, location)] += 1

    dup_row_idx = set()
    for (product, lot), members in group_members.items():
        if len(members) < 2:
            continue
        positive_members = [m for m in members if m[2] > 0]
        if len(positive_members) >= 2:
            for idx, _, _ in positive_members:
                dup_row_idx.add(idx)
        for idx, loc, _ in members:
            if exact_dup_key_count[(product, lot, loc)] > 1:
                dup_row_idx.add(idx)

    # ---- Pass 2: tag every row ----
    for idx, row in enumerate(rows):
        product = row["Product"]
        lot = row["Lot/Serial Number"].strip()
        owner = row["Owner"].strip()

        flags = []

        if lot:
            if idx in dup_row_idx:
                flags.append("ERP Duplicate line")
            if not ALNUM_RE.match(lot):
                flags.append("Has Special Character")
            if len(lot) != mode_length.get(product):
                flags.append("Imei Length Difference")
            if product in numeric_type_products and any(ch.isalpha() for ch in lot):
                flags.append("Alphanumeric IMEI")

        try:
            qty = float(row["Quantity"])
        except ValueError:
            qty = 0.0
        if qty < 0:
            flags.append("Negative Qty")

        if row["Location"].endswith(owner_check_suffix) and owner != valid_owner:
            flags.append("Owner Mismatch")

        if product in mode_length and not lot:
            flags.append("Missing Imei")

        row["Abnormality"] = ", ".join(flags)

    # ---- Filter to the requested location scope ----
    scoped_rows = [r for r in rows if r["Location"].endswith(location_suffix)]

    out_fieldnames = list(fieldnames) + ["Abnormality"]
    output_path = resolve_output_path(input_path, args.output)
    actual_output_path = write_with_fallback(scoped_rows, out_fieldnames, output_path)

    flag_counts = Counter()
    for r in scoped_rows:
        for flag in r["Abnormality"].split(", "):
            if flag:
                flag_counts[flag] += 1

    summary = {
        "input_path": str(input_path),
        "output_path": str(actual_output_path),
        "output_path_changed": actual_output_path != output_path,
        "total_rows_in_scope": len(scoped_rows),
        "rows_with_abnormality": sum(1 for r in scoped_rows if r["Abnormality"]),
        "flag_counts": dict(flag_counts.most_common()),
        "valid_owner_used": valid_owner,
        "location_suffix_used": location_suffix,
        "owner_check_suffix_used": owner_check_suffix,
    }

    # Human-readable summary on stdout...
    print(f"Input:  {input_path}")
    print(f"Output: {actual_output_path}" + (" (original path was locked, used fallback name)" if summary["output_path_changed"] else ""))
    print(f"Rows in scope ({location_suffix}): {summary['total_rows_in_scope']}")
    print(f"Rows with at least one abnormality: {summary['rows_with_abnormality']}")
    print("Breakdown by abnormality type:")
    for k, v in flag_counts.most_common():
        print(f"  {k}: {v}")

    # ...plus a machine-readable line the calling agent can parse reliably.
    print("SUMMARY_JSON:" + json.dumps(summary))


if __name__ == "__main__":
    main()
