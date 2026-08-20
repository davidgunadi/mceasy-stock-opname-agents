#!/usr/bin/env python3
"""
compare_stock_opname.py

Compares cleaned ERP Stock Quant data (the output of clean_stock_quant.py) for one
city against a physical Stock Opname count for the same city/date, and produces a
3-sheet Excel report: Summary, Product Summary, IMEI Mismatch.

## Inputs
- --cleaned-csv     : output of clean_stock_quant.py (already has an Abnormality column)
- --opname-xlsx     : the physical count workbook. Must contain a
                       "Stock <CITY> <DATE> - IMEI" sheet and a
                       "Stock <CITY> <DATE> - Non IMEI" sheet for the resolved date.
- --masterfile-xlsx : product master data (a "Category" sheet with columns
                       Product, Category, Price in IDR, New Group - July 25).
                       Only products present in this master file are considered —
                       everything else is out of scope for this report.
- --city            : e.g. "SBY", "JKT"
- --date            : optional 6-digit override (e.g. "260730"); if omitted it's
                       inferred from the --cleaned-csv filename.

## Key design decisions (locked in with the business owner on 2026-08-11)

1. ERP "on-hand" rows that are flagged `ERP Duplicate line` by clean_stock_quant.py
   are EXCLUDED from the ERP side of the IMEI match. Rationale, confirmed against
   real SBY 260730 data: these are lines where the same lot has Quantity > 0 in two
   places at once (e.g. still shown in `SBY/Stock` *and* already in
   `Partners/Customers`) — a physical count independently confirms these are phantom
   ERP bookings (the physical count finds 0 of them), so counting them as "on hand"
   would make the ERP-vs-physical mismatch look artificially large.
2. IMEI matching is done by **Lot/Serial Number alone** (not Product+Lot), so a unit
   that ended up recorded under a different product name on one side still counts as
   a match — but gets flagged as "Product Mismatch" rather than silently ignored.
3. Stock Opname Accuracy has two independent measures, both computed for "All
   products" and "A+B only" (New Group A/B; Group C excluded from the A+B cut):
     a. Qty-based:      1 - (sum|Qty_ERP - Qty_Physical| / sum(Qty_ERP))
     b. Product-count-based, IMEI-tracked products only: how many IMEI-tracked
        products have ALL their IMEIs matched ("cocok") vs at least one mismatch
        ("tidak cocok").
"""

import argparse
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

import csv
import openpyxl
from openpyxl.styles import Font, Alignment

# Some log/report strings below use an em dash (—). On Windows, stdout defaults
# to the console's codepage (e.g. cp1252) rather than UTF-8, which raises
# UnicodeEncodeError on that character — force UTF-8 explicitly.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def parse_args():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--cleaned-csv", required=True, help="Output of clean_stock_quant.py")
    p.add_argument("--opname-xlsx", required=True, help="Physical Stock Opname workbook")
    p.add_argument("--masterfile-xlsx", required=True, help="Inventory master file (Category sheet)")
    p.add_argument("--city", required=True, help="e.g. SBY, JKT")
    p.add_argument("--date", default=None, help="6-digit date token e.g. 260730; inferred from --cleaned-csv filename if omitted")
    p.add_argument("--output", default=None)
    return p.parse_args()


def infer_date(cleaned_csv_path, override):
    if override:
        return override
    m = re.search(r"(\d{6})", Path(cleaned_csv_path).stem)
    if not m:
        raise ValueError(f"Could not infer a 6-digit date from cleaned CSV filename: {cleaned_csv_path}. Pass --date explicitly.")
    return m.group(1)


def load_masterfile(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Category"]
    header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    col_idx = {name: i for i, name in enumerate(header)}
    master = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        product = row[col_idx["Product"]]
        if not product:
            continue
        master[product] = {
            "price": row[col_idx["Price in IDR"]] or 0,
            "category": row[col_idx["Category"]],
            "new_group": row[col_idx["New Group - July 25"]],
        }
    return master


def load_cleaned_erp(path, city, whitelist):
    """Returns:
      erp_lot_rows: lot -> ERP row dict (qty>0, NOT flagged ERP Duplicate line)
      qty_by_product: product -> sum(Quantity) across all scoped rows (raw, used for non-IMEI products)
      dup_excluded_count: how many rows were excluded due to ERP Duplicate line
    """
    location = f"{city}/Stock"
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    scoped = [r for r in rows if r["Location"] == location and r["Product"] in whitelist]

    erp_lot_rows = {}
    qty_by_product = defaultdict(float)
    lot_tracked_products = set()  # products that have >=1 lot-bearing row, REGARDLESS of exclusion
    dup_excluded_count = 0
    for r in scoped:
        try:
            qty = float(r["Quantity"])
        except ValueError:
            qty = 0.0
        qty_by_product[r["Product"]] += qty
        lot = r["Lot/Serial Number"].strip()
        if lot:
            lot_tracked_products.add(r["Product"])
        if lot and qty > 0:
            if "ERP Duplicate line" in r["Abnormality"]:
                dup_excluded_count += 1
                continue
            erp_lot_rows[lot] = r

    return erp_lot_rows, qty_by_product, dup_excluded_count, lot_tracked_products


def load_physical_imei(wb, sheet_name, whitelist):
    ws = wb[sheet_name]
    physical_lot_rows = {}
    scan_counts = Counter()
    for row in ws.iter_rows(min_row=2, values_only=True):
        product, sku, lokasi, box_ke, lot, nomor_box = (tuple(row) + (None,) * 6)[:6]
        if product is None or product not in whitelist:
            continue
        if lot is None or str(lot).strip() == "":
            continue
        lot = str(lot).strip()
        scan_counts[lot] += 1
        physical_lot_rows[lot] = {
            "Product": product, "SKU": sku, "Lokasi Gudang": lokasi,
            "Box Ke-": box_ke, "Nomor Box": nomor_box,
        }
    duplicate_scans = {lot: n for lot, n in scan_counts.items() if n > 1}
    return physical_lot_rows, duplicate_scans


def load_physical_non_imei(wb, sheet_name, whitelist):
    ws = wb[sheet_name]
    qty_by_product = defaultdict(float)
    for row in ws.iter_rows(min_row=2, values_only=True):
        product, sku, lokasi, box_ke, qty, nomor_box = (tuple(row) + (None,) * 6)[:6]
        if product is None or product not in whitelist:
            continue
        try:
            qty_by_product[product] += float(qty or 0)
        except (TypeError, ValueError):
            pass
    return qty_by_product


def build_product_summary(whitelist, master, erp_lot_rows, erp_qty_by_product,
                           physical_lot_rows, physical_qty_nonimei, matched_lots, lot_tracked_products):
    """One row per whitelisted product that has any ERP or physical activity.

    A product counts as "IMEI-tracked" if it EVER shows up with a lot number on
    either side — even if every single one of its ERP lots got excluded as
    ERP Duplicate line (that should net out to a 0-vs-0 match, not fall back to
    the raw/uncorrected quantity sum, which would silently re-introduce the very
    phantom units the exclusion was meant to remove)."""
    imei_products = lot_tracked_products | {r["Product"] for r in physical_lot_rows.values()}

    rows = []
    for product in sorted(whitelist):
        info = master[product]
        if product in imei_products:
            qty_erp = sum(1 for r in erp_lot_rows.values() if r["Product"] == product)
            qty_phys = sum(1 for r in physical_lot_rows.values() if r["Product"] == product)
            relevant_lots = {lot for lot, r in erp_lot_rows.items() if r["Product"] == product} | \
                            {lot for lot, r in physical_lot_rows.items() if r["Product"] == product}
            matched_for_product = len(relevant_lots & matched_lots)
            match_rate = (matched_for_product / len(relevant_lots)) if relevant_lots else None
        else:
            qty_erp = erp_qty_by_product.get(product, 0)
            qty_phys = physical_qty_nonimei.get(product, 0)
            match_rate = None

        if qty_erp == 0 and qty_phys == 0:
            continue  # nothing to report for a product with zero activity on both sides

        difference = qty_erp - qty_phys
        rows.append({
            "Product": product,
            "Price in IDR": info["price"],
            "Qty_ERP": qty_erp,
            "Qty_Physical": qty_phys,
            "Difference": difference,
            "Discrepancy Value (IDR)": difference * (info["price"] or 0),
            "IMEI Match Rate (%)": match_rate,
            "Category": info["category"],
            "New Group - July 25": info["new_group"],
        })
    return rows


def compute_accuracy(product_rows, groups=None):
    rows = [r for r in product_rows if groups is None or r["New Group - July 25"] in groups]
    total_qty_erp = sum(r["Qty_ERP"] for r in rows)
    total_abs_diff = sum(abs(r["Difference"]) for r in rows)
    if total_qty_erp == 0:
        return None
    return 1 - (total_abs_diff / total_qty_erp)


def write_report(output_path, city, date, summary_ctx, product_rows, mismatch_rows):
    wb = openpyxl.Workbook()
    bold = Font(bold=True)

    # ---- Summary sheet ----
    ws = wb.active
    ws.title = "Summary"
    r = 1
    ws.cell(r, 1, f"Stock Opname Summary — {city} {date}").font = Font(bold=True, size=14)
    r += 3
    ws.cell(r, 1, "IMEI Overview").font = bold
    r += 1
    for label, key in [
        ("Total IMEIs in ERP", "total_erp"),
        ("Total IMEIs in Physical", "total_physical"),
        ("Matched IMEIs", "matched"),
        ("Unmatched IMEIs in ERP", "erp_only"),
        ("Unmatched IMEIs in Physical", "physical_only"),
        ("IMEI Match Rate (%)", "match_rate"),
    ]:
        ws.cell(r, 1, label)
        ws.cell(r, 2, summary_ctx[key])
        r += 1
    r += 2
    ws.cell(r, 1, "Stock Opname Accuracy").font = bold
    r += 1
    ws.cell(r, 1, "Stock Opname Accuracy (All)"); ws.cell(r, 2, summary_ctx["accuracy_all"]); r += 1
    ws.cell(r, 1, "Stock Opname Accuracy (A+B)"); ws.cell(r, 2, summary_ctx["accuracy_ab"]); r += 1
    r += 2
    ws.cell(r, 1, "IMEI-tracked Products — Match Count").font = bold
    r += 1
    ws.cell(r, 1, "Products fully matched (cocok)"); ws.cell(r, 2, summary_ctx["products_match"]); r += 1
    ws.cell(r, 1, "Products with >=1 mismatch (tidak cocok)"); ws.cell(r, 2, summary_ctx["products_mismatch"]); r += 1
    r += 2
    ws.cell(r, 1, "Discrepancy Overview (IDR)").font = bold
    r += 1
    ws.cell(r, 1, "Discrepancy Value (All)"); ws.cell(r, 2, summary_ctx["discrepancy_all"]); r += 1
    ws.cell(r, 1, "Discrepancy Value (A+B)"); ws.cell(r, 2, summary_ctx["discrepancy_ab"]); r += 1
    r += 2
    ws.cell(r, 1, "Summary per Group (A–C)").font = bold
    r += 1
    for i, h in enumerate(["New Group - July 25", "Products_Count", "Total_Qty", "Total_Discrepancy_Qty", "Total_Rupiah"]):
        ws.cell(r, i + 1, h).font = bold
    r += 1
    for group in ["A", "B", "C"]:
        rows_g = [p for p in product_rows if p["New Group - July 25"] == group]
        ws.cell(r, 1, group)
        ws.cell(r, 2, len(rows_g))
        ws.cell(r, 3, sum(p["Qty_ERP"] for p in rows_g))
        ws.cell(r, 4, sum(abs(p["Difference"]) for p in rows_g))
        ws.cell(r, 5, sum(p["Discrepancy Value (IDR)"] for p in rows_g))
        r += 1
    for col, width in zip("ABCDE", [32, 16, 14, 20, 16]):
        ws.column_dimensions[col].width = width

    # ---- Product Summary sheet ----
    ws2 = wb.create_sheet("Product Summary")
    headers = ["Product", "Price in IDR", "Qty_ERP", "Qty_Physical", "Difference",
               "Discrepancy Value (IDR)", "IMEI Match Rate (%)", "Category", "New Group - July 25"]
    for i, h in enumerate(headers):
        ws2.cell(1, i + 1, h).font = bold
    for i, p in enumerate(product_rows):
        for j, h in enumerate(headers):
            ws2.cell(i + 2, j + 1, p[h])
    ws2.column_dimensions["A"].width = 40

    # ---- IMEI Mismatch sheet ----
    ws3 = wb.create_sheet("IMEI Mismatch")
    mismatch_headers = [
        "Lot/Serial Number", "Match Status",
        "Product_ERP", "Location", "Quantity", "Owner", "Lot Created Date", "Abnormality",
        "Product_Physical", "SKU", "Lokasi Gudang", "Box Ke-", "Nomor Box",
    ]
    for i, h in enumerate(mismatch_headers):
        ws3.cell(1, i + 1, h).font = bold
    for i, row in enumerate(mismatch_rows):
        for j, h in enumerate(mismatch_headers):
            ws3.cell(i + 2, j + 1, row.get(h))
    ws3.column_dimensions["A"].width = 20
    ws3.column_dimensions["C"].width = 32
    ws3.column_dimensions["I"].width = 32

    wb.save(output_path)


def main():
    args = parse_args()
    date = infer_date(args.cleaned_csv, args.date)
    city = args.city

    master = load_masterfile(args.masterfile_xlsx)
    whitelist = set(master.keys())

    erp_lot_rows, erp_qty_by_product, dup_excluded_count, lot_tracked_products = load_cleaned_erp(args.cleaned_csv, city, whitelist)

    wb_op = openpyxl.load_workbook(args.opname_xlsx, data_only=True)
    imei_sheet = f"Stock {city} {date} - IMEI"
    non_imei_sheet = f"Stock {city} {date} - Non IMEI"
    if imei_sheet not in wb_op.sheetnames or non_imei_sheet not in wb_op.sheetnames:
        print(f"ERROR: expected sheets {imei_sheet!r} / {non_imei_sheet!r} not found. "
              f"Available sheets: {wb_op.sheetnames}", file=sys.stderr)
        sys.exit(1)

    physical_lot_rows, duplicate_scans = load_physical_imei(wb_op, imei_sheet, whitelist)
    physical_qty_nonimei = load_physical_non_imei(wb_op, non_imei_sheet, whitelist)

    erp_lots = set(erp_lot_rows.keys())
    physical_lots = set(physical_lot_rows.keys())
    matched_lots = erp_lots & physical_lots
    erp_only_lots = erp_lots - physical_lots
    physical_only_lots = physical_lots - erp_lots
    product_mismatch_lots = {lot for lot in matched_lots if erp_lot_rows[lot]["Product"] != physical_lot_rows[lot]["Product"]}

    union_lots = erp_lots | physical_lots
    match_rate = (len(matched_lots) / len(union_lots)) if union_lots else None

    product_rows = build_product_summary(
        whitelist, master, erp_lot_rows, erp_qty_by_product,
        physical_lot_rows, physical_qty_nonimei, matched_lots, lot_tracked_products,
    )

    imei_tracked_rows = [p for p in product_rows if p["IMEI Match Rate (%)"] is not None]
    products_match = sum(1 for p in imei_tracked_rows if p["IMEI Match Rate (%)"] == 1)
    products_mismatch = sum(1 for p in imei_tracked_rows if p["IMEI Match Rate (%)"] < 1)

    summary_ctx = {
        "total_erp": len(erp_lots),
        "total_physical": len(physical_lots),
        "matched": len(matched_lots),
        "erp_only": len(erp_only_lots),
        "physical_only": len(physical_only_lots),
        "match_rate": match_rate,
        "accuracy_all": compute_accuracy(product_rows),
        "accuracy_ab": compute_accuracy(product_rows, groups={"A", "B"}),
        "products_match": products_match,
        "products_mismatch": products_mismatch,
        "discrepancy_all": sum(p["Discrepancy Value (IDR)"] for p in product_rows),
        "discrepancy_ab": sum(p["Discrepancy Value (IDR)"] for p in product_rows if p["New Group - July 25"] in {"A", "B"}),
    }

    mismatch_rows = []
    for lot in sorted(product_mismatch_lots):
        e, ph = erp_lot_rows[lot], physical_lot_rows[lot]
        mismatch_rows.append({
            "Lot/Serial Number": lot, "Match Status": "Product Mismatch",
            "Product_ERP": e["Product"], "Location": e["Location"], "Quantity": e["Quantity"],
            "Owner": e["Owner"], "Lot Created Date": e["Lot Created Date"], "Abnormality": e["Abnormality"],
            "Product_Physical": ph["Product"], "SKU": ph["SKU"], "Lokasi Gudang": ph["Lokasi Gudang"],
            "Box Ke-": ph["Box Ke-"], "Nomor Box": ph["Nomor Box"],
        })
    for lot in sorted(erp_only_lots):
        e = erp_lot_rows[lot]
        mismatch_rows.append({
            "Lot/Serial Number": lot, "Match Status": "Missing in Physical",
            "Product_ERP": e["Product"], "Location": e["Location"], "Quantity": e["Quantity"],
            "Owner": e["Owner"], "Lot Created Date": e["Lot Created Date"], "Abnormality": e["Abnormality"],
        })
    for lot in sorted(physical_only_lots):
        ph = physical_lot_rows[lot]
        mismatch_rows.append({
            "Lot/Serial Number": lot, "Match Status": "Missing in ERP",
            "Product_Physical": ph["Product"], "SKU": ph["SKU"], "Lokasi Gudang": ph["Lokasi Gudang"],
            "Box Ke-": ph["Box Ke-"], "Nomor Box": ph["Nomor Box"],
        })

    output_path = args.output or str(Path(args.opname_xlsx).with_name(f"Stock Opname Report {city} {date}.xlsx"))
    write_report(output_path, city, date, summary_ctx, product_rows, mismatch_rows)

    print(f"City: {city}  Date: {date}")
    print(f"ERP rows excluded (ERP Duplicate line): {dup_excluded_count}")
    print(f"Total IMEIs in ERP: {summary_ctx['total_erp']}")
    print(f"Total IMEIs in Physical: {summary_ctx['total_physical']}")
    print(f"Matched: {summary_ctx['matched']}  ERP-only: {summary_ctx['erp_only']}  Physical-only: {summary_ctx['physical_only']}")
    print(f"Product Mismatch (same lot, different product name): {len(product_mismatch_lots)}")
    print(f"Physical-side duplicate scans: {len(duplicate_scans)}")
    print(f"IMEI Match Rate: {summary_ctx['match_rate']}")
    print(f"Stock Opname Accuracy (All): {summary_ctx['accuracy_all']}")
    print(f"Stock Opname Accuracy (A+B): {summary_ctx['accuracy_ab']}")
    print(f"IMEI-tracked products — Match: {products_match}  Mismatch: {products_mismatch}")
    print(f"Discrepancy Value (All): {summary_ctx['discrepancy_all']:,.0f} IDR")
    print(f"Discrepancy Value (A+B): {summary_ctx['discrepancy_ab']:,.0f} IDR")
    print(f"Output written to: {output_path}")
    if duplicate_scans:
        print("NOTE: physical-side duplicate scans found (same lot scanned more than once) — data entry issue, not reflected above:")
        for lot, n in list(duplicate_scans.items())[:10]:
            print(f"   {lot}: scanned {n}x")


if __name__ == "__main__":
    main()
