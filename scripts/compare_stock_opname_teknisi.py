#!/usr/bin/env python3
"""
compare_stock_opname_teknisi.py

Reconciles ERP inventory against a physical Stock Opname count done by
technicians (as opposed to compare_stock_opname.py, which reconciles a
per-city warehouse count).

Adapted from a technician-provided reference script ("stock opname teknisi",
currently at v6). Behavioural differences from that reference, made
deliberately while porting into this repo:

  - ERP cleansing is NOT done by this script. It expects an already-cleaned
    ERP CSV produced by `clean_stock_quant.py --location-suffix ""` (the
    empty suffix keeps every location, not just */Stock — this script needs
    to know where an item currently sits ANYWHERE in the ledger, e.g.
    /Customer or /Teknisi, not just on-hand warehouse stock). Rows with any
    non-empty "Abnormality" value are excluded from the ERP lookup, same as
    the reference script excluded rows with Has Abnormality=True.
  - All file paths are explicit CLI arguments instead of an auto-detected
    OneDrive folder.
  - Stock Movement and WebSMS device-list inputs accept either .csv/.xlsx
    (movement) or .xlsx (WebSMS) — extension is auto-detected. xlsx movement
    reads use the 'calamine' engine (matching the reference script as of v6)
    — far faster than openpyxl on the 300k+-row exports this file routinely is.
  - Fixed a bug still present in the reference script as of v6: a Non-IMEI
    row with a blank "SO Location" referenced undefined variables
    (erp_qty/status) and would crash. It now reports
    "Inaccuracy — Missing SO Location", mirroring the equivalent guard
    already present on the IMEI side.

IMEI (serialized) rows: use the LATEST DONE move on or after SO Date (a
same-day move counts as evidence — dates are compared date-only, so same-day
is the finest resolution available); no "Validate" status — only
OK/Inaccuracy. When an ERP serial has duplicate rows and no "Last Movement
Date" to break the tie, the row with Quantity > 0 wins over zero-quantity
"ghost" rows at old locations.
Non-IMEI (non-serialized) rows: compare per (Technician x Location x Product);
skip '/Customer' entirely.

Usage:
    python compare_stock_opname_teknisi.py \\
        --erp-csv "01 Stock ERP.csv" \\
        --movement "02 Stock Movement.xlsx" \\
        --east "03 East Stock Opname Teknisi.xlsx" \\
        --west "04 West Stock Opname Teknisi.xlsx" \\
        --masterfile "Inventory Masterfile.xlsx" \\
        --device-id "Device ID.xlsx" --device-sg "Device SG.xlsx" \\
        --output "Stock Opname Teknisi_260812.xlsx"
"""

import argparse
import re
import sys
import warnings
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd

try:
    import openpyxl  # noqa: F401
except Exception:
    pass

# Log messages below use unicode arrows/emoji (→, ✅, ⚠️, ❌). On Windows, stdout
# defaults to the console's codepage (e.g. cp1252) rather than UTF-8, which
# raises UnicodeEncodeError on those characters — force UTF-8 explicitly.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


# =============================
# CONFIG
# =============================
STRICT_SCOPE = True  # True: restrict products using Masterfile['Include for Stock Opname Teknisi']


# =============================
# LIGHTWEIGHT LOGGER
# =============================
def log(msg: str):
    now = datetime.now().strftime("%H:%M:%S")
    print(f"[{now}] {msg}", flush=True)


# =============================
# UTILITIES
# =============================
def norm_text(s):
    """Normalize text: handle NaN, strip, collapse slashes to '/', lowercase, collapse whitespace."""
    if pd.isna(s):
        return ""
    s = str(s).strip()
    s = re.sub(r"[\\/]+", "/", s)
    s = re.sub(r"\s+", " ", s)
    return s.lower()


def norm_owner(s):
    """Normalize Owner/Technician names (remove company prefix and lowercase)."""
    base = norm_text(s)
    return re.sub(r"^pt\s+otto\s+menara\s+globalindo,?\s*", "", base)


def date_only(x):
    """Parse any date/time-like value and return a date-only Timestamp.
    Robust to strings like '20 October 2025 07:00:49 (UTC+08:00)'."""
    if pd.isna(x):
        return pd.NaT
    s = str(x).strip()
    s = re.sub(r"\(.*?\)", "", s).strip()
    dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
    return pd.NaT if pd.isna(dt) else pd.to_datetime(dt).normalize()


def make_unique_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure df has unique column names by suffixing duplicates with .1, .2, ..."""
    seen = {}
    new_cols = []
    for c in map(str, df.columns):
        base = c.strip()
        n = seen.get(base, 0)
        new_cols.append(base if n == 0 else f"{base}.{n}")
        seen[base] = n + 1
    df = df.copy()
    df.columns = new_cols
    return df


# =============================
# SCOPE LOADER (Masterfile)
# =============================
def build_scope(masterfile: Path):
    """Read Inventory Masterfile.xlsx -> sheet 'Category' and build a set of
    product codes in scope. If 'Include for Stock Opname Teknisi' column is
    present and STRICT_SCOPE=True, filter by that flag."""
    log("Loading Masterfile scope…")
    cat = pd.read_excel(masterfile, sheet_name="Category")
    if "Product" not in cat.columns:
        raise RuntimeError("[MASTERFILE] 'Category' sheet missing 'Product' column.")
    prod_upper = cat["Product"].astype(str).str.strip().str.upper()
    if "Include for Stock Opname Teknisi" in cat.columns and STRICT_SCOPE:
        mask = cat["Include for Stock Opname Teknisi"].fillna(False) == True  # noqa: E712
        scope = set(prod_upper[mask].tolist())
        log(f"→ STRICT_SCOPE active with 'Include for Stock Opname Teknisi' ({len(scope)} products).")
    else:
        if STRICT_SCOPE:
            warnings.warn("[MASTERFILE] 'Include for Stock Opname Teknisi' not found — defaulting to all products in-scope.")
        scope = set(prod_upper.tolist())
        log(f"→ STRICT_SCOPE defaulted to ALL products ({len(scope)}).")
    return scope


# =============================
# LOADERS
# =============================
def load_erp_quant(erp_csv: Path, in_scope: set) -> pd.DataFrame:
    """Load the pre-cleaned ERP CSV (from clean_stock_quant.py --location-suffix ""),
    drop rows with any Abnormality, normalize, and apply scope filter."""
    if not erp_csv.exists():
        raise FileNotFoundError(f"Missing cleaned ERP CSV: {erp_csv}")
    log(f"Loading cleaned ERP quant → {erp_csv}")
    df = pd.read_csv(erp_csv, dtype={"Lot/Serial Number": str})
    if "Abnormality" not in df.columns:
        raise RuntimeError(
            f"[ERP] '{erp_csv}' has no 'Abnormality' column — is this the output of "
            "clean_stock_quant.py?"
        )
    df["Abnormality"] = df["Abnormality"].fillna("")
    before = len(df)
    df = df[df["Abnormality"].str.strip() == ""].copy()
    log(f"→ Dropped {before - len(df)} row(s) flagged with an abnormality; {len(df)} clean row(s) remain.")

    df["Product"] = df["Product"].astype(str).str.strip().str.upper()
    df["Location"] = df["Location"].map(norm_text)
    df["Owner"] = df["Owner"].map(norm_text)
    df["owner_norm"] = df["Owner"].map(norm_owner)
    df["quantity_num"] = pd.to_numeric(df.get("Quantity", 0), errors="coerce").fillna(0.0)
    if STRICT_SCOPE:
        df = df[df["Product"].isin(in_scope)].copy()
        log(f"→ ERP rows after scope filter: {len(df)}")
    else:
        log(f"→ ERP rows loaded: {len(df)}")
    return df


def _read_movement_table(movement_path: Path) -> pd.DataFrame:
    """Read the raw Stock Movement export, csv or xlsx, sheet 'Stock Move Line'
    if present (else the first sheet). xlsx is read with the 'calamine' engine —
    far faster than the default openpyxl engine on large exports (this file is
    routinely 300k+ rows; calamine cuts load time from minutes to seconds)."""
    suffix = movement_path.suffix.lower()
    if suffix == ".xlsx":
        wb_sheets = pd.ExcelFile(movement_path, engine="calamine").sheet_names
        sheet = "Stock Move Line" if "Stock Move Line" in wb_sheets else wb_sheets[0]
        return pd.read_excel(movement_path, sheet_name=sheet, dtype={"Lot/Serial Number": str}, engine="calamine")
    return pd.read_csv(movement_path, dtype={"Lot/Serial Number": str}, low_memory=False)


def load_moves(movement_path: Path) -> pd.DataFrame:
    """Load Stock Movement export, keep only DONE moves, normalize date/dest/owner fields."""
    if not movement_path.exists():
        raise FileNotFoundError(f"Missing Stock Movement file: {movement_path}")

    log(f"Loading Stock Movement → {movement_path}")
    is_xlsx = movement_path.suffix.lower() == ".xlsx"
    mv = _read_movement_table(movement_path)

    mv = mv.rename(columns={
        "To": "Destination Location",
        "Dest": "Destination Location",
        "Dest Location": "Destination Location",
        "Date Done": "Date",
        # The xlsx export uses "Destination Owner" / "Origin Owner" instead of
        # the older csv export's "Dest Owner" / "Source Owner" — map them so
        # the owner-norm logic below finds the columns either way.
        "Destination Owner": "Dest Owner",
        "Origin Owner": "Source Owner",
    })

    mv["Lot/Serial Number"] = mv["Lot/Serial Number"].astype(str).str.strip()
    mv["Status"] = mv["Status"].astype(str).str.strip().str.lower()
    mv = mv[mv["Status"].eq("done")].copy()

    # The xlsx export's Date column is already ISO (YYYY-MM-DD HH:MM:SS), so
    # dayfirst is irrelevant there (and triggers a pandas warning if passed);
    # older csv exports use DD/MM/YYYY, where dayfirst=True is required.
    if is_xlsx:
        mv["Date"] = pd.to_datetime(mv["Date"], errors="coerce").dt.normalize()
    else:
        mv["Date"] = pd.to_datetime(mv["Date"], dayfirst=True, errors="coerce").dt.normalize()
    mv["Destination Location"] = mv["Destination Location"].map(norm_text)

    if "Dest Owner" in mv.columns:
        mv["dest_owner_norm"] = mv["Dest Owner"].map(norm_owner)
    else:
        mv["dest_owner_norm"] = ""

    if "Source Owner" in mv.columns:
        mv["source_owner_norm"] = mv["Source Owner"].map(norm_owner)
    else:
        mv["source_owner_norm"] = ""

    log(f"→ Movement rows (done): {len(mv)}")
    return mv


# -----------------------------
# Excel helpers — detect header row robustly
# -----------------------------
HEADER_CANDIDATES = {"product", "imei", "lot/serial number", "lot/serial", "serial", "qty", "quantity", "so location", "location"}


def detect_header_row_xl(ws) -> int:
    max_row = min(getattr(ws, "max_row", 10), 10)
    max_col = min(getattr(ws, "max_column", 12), 12)
    for r in range(1, max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, max_col + 1)]
        normed = [norm_text(v) for v in vals]
        hits = sum(1 for v in normed if v in HEADER_CANDIDATES)
        if hits >= 3:
            return r
    return 1


def load_teknisi_workbook(xlsx: Path, area_label: str) -> pd.DataFrame:
    """Load a single teknisi workbook: read SO Date (B1), Technician (B2), detect
    header, normalize Product/Qty/Serial/SO Location, return a flat DataFrame."""
    if not xlsx.exists():
        raise FileNotFoundError(f"Missing teknisi workbook: {xlsx}")
    import openpyxl
    SKIP_SHEETS = {"index", "summary", "debug", "scope", "sheet1"}
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    all_rows = []
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        name = ws.title.strip().lower()
        if name in SKIP_SHEETS:
            continue
        opname_date = date_only(ws["B1"].value)
        technician = ws["B2"].value
        hdr_row = detect_header_row_xl(ws)
        data = ws.iter_rows(min_row=hdr_row, values_only=True)
        cols = [str(c or "").strip() for c in next(data)]
        recs = [list(row) for row in data]
        df = pd.DataFrame(recs, columns=cols)
        df = make_unique_columns(df)

        def pick(*names):
            for name in names:
                if name in df.columns:
                    return name
            lower_map = {str(c).strip().lower(): c for c in df.columns}
            for name in names:
                key = name.lower()
                if key in lower_map:
                    return lower_map[key]
            return None

        col_prod = pick("Product", "Product Name")
        col_qty = pick("Qty", "Quantity")
        col_serial = pick("Lot/Serial Number", "Lot/Serial", "Serial", "IMEI")
        col_loc = pick("SO Location", "Location")

        if col_prod is None and col_serial is None and col_qty is None:
            tech_name = str(technician or "").strip()
            if tech_name:
                log(f"⚠️ Skipped technician '{tech_name}' — no usable data in sheet '{sheet}'")
            else:
                log(f"⚠️ Skipped unnamed technician sheet '{sheet}' (empty or template)")
            continue

        df["Product"] = df[col_prod].astype(str).str.strip().str.upper() if col_prod else ""
        df["Qty"] = pd.to_numeric(df[col_qty], errors="coerce").fillna(0) if col_qty else 0
        if col_serial:
            ser = df[col_serial].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
            df["Lot/Serial Number"] = ser
        else:
            df["Lot/Serial Number"] = ""
        df["SO Location"] = df[col_loc].map(lambda x: re.sub(r"[\\/]+", "/", str(x).strip()) if pd.notna(x) else "") if col_loc else ""

        df["technician"] = str(technician or "").strip()
        df["technician_norm"] = df["technician"].map(norm_owner)
        df["opname_date"] = opname_date
        df["technician_sheet"] = f"{xlsx.name}::{sheet}"
        df["area"] = area_label
        all_rows.append(df)

    if not all_rows:
        return pd.DataFrame(columns=["Product", "Qty", "Lot/Serial Number", "SO Location", "technician", "technician_norm", "opname_date", "technician_sheet", "area"])
    out = pd.concat(all_rows, ignore_index=True)
    out = out[(out["Product"].astype(str).str.strip() != "")]
    out["Qty"] = pd.to_numeric(out["Qty"], errors="coerce").fillna(0)
    out["SO Location Norm"] = out["SO Location"].map(norm_text)
    return out


def load_teknisi_all(east_path: Path, west_path: Path, in_scope: set):
    """Load East & West workbooks, merge, then split into IMEI vs Non-IMEI."""
    east = load_teknisi_workbook(east_path, "EAST")
    log(f"✅ East file loaded → {len(east):,} rows")

    west = load_teknisi_workbook(west_path, "WEST")
    log(f"✅ West file loaded → {len(west):,} rows")

    if not isinstance(east.index, pd.RangeIndex) or not east.index.is_unique:
        east = east.reset_index(drop=True)
    if not isinstance(west.index, pd.RangeIndex) or not west.index.is_unique:
        west = west.reset_index(drop=True)

    east = make_unique_columns(east)
    west = make_unique_columns(west)

    all_rows = pd.concat([east, west], ignore_index=True)

    all_rows["SO Location Norm"] = all_rows["SO Location"].map(norm_text)
    all_rows = all_rows[all_rows["SO Location Norm"].str.contains("/teknisi", na=False)].copy()

    dupe_cols = [c for c in all_rows.columns if list(all_rows.columns).count(c) > 1]
    if dupe_cols:
        log(f"[WARN] Duplicate column names after concat: {sorted(set(dupe_cols))[:10]}")
        all_rows = make_unique_columns(all_rows)

    all_rows["Product"] = all_rows["Product"].astype(str).str.strip().str.upper()
    if STRICT_SCOPE:
        all_rows = all_rows[all_rows["Product"].isin(in_scope)].copy()

    ser = all_rows["Lot/Serial Number"].astype(str).fillna("").str.strip()
    no_serial = ser.eq("") | ser.str.lower().eq("nan") | ser.isin(["-", "0"])
    is_imei = ~no_serial

    imei_df = all_rows[is_imei].copy()
    non_df = all_rows[~is_imei].copy()

    imei_df.reset_index(drop=True, inplace=True)
    non_df.reset_index(drop=True, inplace=True)
    imei_df["Qty"] = pd.to_numeric(imei_df["Qty"], errors="coerce").fillna(0)
    non_df["Qty"] = pd.to_numeric(non_df["Qty"], errors="coerce").fillna(0)
    imei_df["SO Location Norm"] = imei_df["SO Location"].map(norm_text)
    non_df["SO Location Norm"] = non_df["SO Location"].map(norm_text)

    log(f"→ Teknisi rows: IMEI={len(imei_df)}, Non-IMEI={len(non_df)} (total={len(all_rows)})")
    return imei_df, non_df


def load_websms(device_id: Optional[Path], device_sg: Optional[Path]) -> set:
    """Load WebSMS device lists (Source of Truth for 'installed at customer').
    Both files use engine='calamine' because their style XML doesn't parse
    with openpyxl (invalid aRGB color values) — requires python-calamine."""
    files = [f for f in (device_id, device_sg) if f is not None]
    imei_list = []
    log("Loading WebSMS data (Source of Truth)...")

    for f in files:
        if f.exists():
            try:
                df = pd.read_excel(f, sheet_name="Perangkat", dtype=str, engine="calamine", header=2)
                col_imei = None
                for c in df.columns:
                    if str(c).strip().upper() == "IMEI":
                        col_imei = c
                        break
                if col_imei:
                    clean_imeis = (
                        df[col_imei]
                        .dropna().astype(str).str.strip()
                        .str.replace(r"\.0$", "", regex=True)
                        .tolist()
                    )
                    imei_list.extend(clean_imeis)
                    log(f"  + Loaded {len(clean_imeis)} IMEIs from {f.name} (Sheet: Perangkat)")
                else:
                    log(f"  ⚠️ Warning: Kolom 'IMEI' tidak ditemukan di {f.name}. Kolom yang terbaca: {list(df.columns)}")
            except Exception as e:
                log(f"  ⚠️ Gagal baca {f.name}: {e}")
        else:
            log(f"  ⚠️ File WebSMS tidak ditemukan: {f.name}")

    unique = set(imei_list)
    log(f"→ Total WebSMS Installed IMEIs: {len(unique)}")
    return unique


# =============================
# INDEX BUILDERS
# =============================
def build_indexes_imei(erp: pd.DataFrame, moves_done: pd.DataFrame):
    idx = {}

    erp_ser = erp.dropna(subset=["Lot/Serial Number"]).copy()
    erp_ser["serial_clean"] = erp_ser["Lot/Serial Number"].astype(str).str.strip()
    if "Last Movement Date" in erp_ser.columns:
        erp_ser["Last Movement Date"] = pd.to_datetime(erp_ser["Last Movement Date"], errors="coerce")
        erp_unique = (
            erp_ser.sort_values(["serial_clean", "Last Movement Date"])
                   .drop_duplicates("serial_clean", keep="last")
        )
    else:
        # No 'Last Movement Date' to break ties (the normal case — that column
        # isn't in clean_stock_quant.py's output). The ERP ledger leaves behind
        # zero-quantity "ghost" rows at a serial's old locations, so among
        # duplicate rows for the same serial, the one with Quantity > 0 is the
        # item's actual current location. Sorting only by serial_clean is a
        # no-op for tied rows, so without this tie-break, drop_duplicates
        # would pick an arbitrary row — sometimes a Qty=0 ghost — instead of
        # the real one.
        erp_ser["_has_qty"] = (erp_ser["quantity_num"] > 0).astype(int)
        erp_unique = (
            erp_ser.sort_values(["serial_clean", "_has_qty"])
                   .drop_duplicates("serial_clean", keep="last")
        )

    idx["erp_by_serial"] = {
        r.serial_clean: {"Location": r.Location, "owner_norm": r.owner_norm, "Product": r.Product}
        for r in erp_unique.itertuples(index=False)
    }

    md = moves_done.copy()
    if "Date" not in md.columns and "Date Done" in md.columns:
        md["Date"] = md["Date Done"]
    md["serial_clean"] = md["Lot/Serial Number"].astype(str).str.strip()
    md["to_norm"] = md["Destination Location"].map(norm_text)

    latest = (
        md.sort_values("Date")
          .dropna(subset=["Date"])
          .groupby(["serial_clean", "to_norm"], as_index=False)["Date"].max()
    )
    idx["latest_move_to_dest"] = {(r.serial_clean, r.to_norm): r.Date for r in latest.itertuples(index=False)}

    md_tek = md[md["to_norm"].astype(str).str.contains("teknisi", na=False)].copy()
    idx["moves_to_teknisi"] = md_tek

    log(
        f"→ Indexes built: ERP serials={len(idx['erp_by_serial'])}, "
        f"latest_move_to_dest={len(idx['latest_move_to_dest'])}, "
        f"moves_to_teknisi={len(idx['moves_to_teknisi'])}"
    )
    return idx


def build_indexes_nonimei(erp: pd.DataFrame):
    out = {}
    sum_no_owner = (
        erp.groupby(["Product", "Location"], dropna=False)["quantity_num"].sum().reset_index().rename(columns={"quantity_num": "erp_qty"})
    )
    sum_with_owner = (
        erp.groupby(["Product", "Location", "owner_norm"], dropna=False)["quantity_num"].sum().reset_index().rename(columns={"quantity_num": "erp_qty"})
    )
    out.update(dict(sum_no_owner=sum_no_owner, sum_with_owner=sum_with_owner))
    log("→ Non-IMEI indexes built.")
    return out


# =============================
# DECISION ENGINES
# =============================
def imei_status(row: pd.Series, idx: dict, websms_set: set) -> str:
    """
    Movement-first + Qty-aware logic (single-serial row):
      1) If latest DONE move ON OR AFTER SO Date -> current ERP location -> OK — Move after opname
      2) Else if ERP location == SO Location:
           - Qty == 0 -> Inaccuracy — ERP still shows stock at technician (Qty=0)
           - Qty > 0  -> OK — ERP match
      3) Else (ERP location != SO Location) and no post-SO move:
           - Qty == 0 -> Inaccuracy — No post-SO move found (ERP at {erp_loc})
           - Qty > 0  -> Inaccuracy — Location mismatch (found in {erp_loc})
      Guards:
        - Missing/invalid serial -> Inaccuracy — Missing IMEI when expected
        - Not in ERP -> Inaccuracy — Missing in ERP
        - Missing SO Location -> Inaccuracy — Missing SO Location
    """
    ser_raw = str(row.get("Lot/Serial Number", "")).strip()
    ser = re.sub(r"\.0$", "", ser_raw)

    if ser == "" or ser.lower() == "nan" or ser in {"-", "0"}:
        return "Inaccuracy — Missing IMEI when expected"

    so_loc = row.get("SO Location", "")
    so_loc_norm = norm_text(so_loc)
    so_date = date_only(row.get("opname_date"))

    if so_loc_norm == "":
        return "Inaccuracy — Missing SO Location"

    qty_val = row.get("Qty", None)
    try:
        qty_zero = (pd.notna(qty_val) and float(qty_val) == 0.0)
    except Exception:
        qty_zero = False

    if ser in websms_set:
        return "OK — Installed at Customer (WebSMS)"

    erp_info = idx["erp_by_serial"].get(ser)
    if erp_info is None:
        return "Inaccuracy — Missing in ERP"

    erp_loc = str(erp_info.get("Location", ""))
    erp_loc_norm = norm_text(erp_loc)

    latest_map = idx.get("latest_move_to_dest", {})
    latest_dt = latest_map.get((ser, erp_loc_norm), pd.NaT)

    # Both the opname date and movement date are date-only (no time-of-day is
    # ever captured from the technician's B1 cell), so a same-day move is the
    # finest resolution we can compare at. Using >= (not strict >) lets a
    # same-day move count as evidence — e.g. the technician counts Qty=0
    # because they installed the unit at a customer earlier that same day;
    # the movement and the report describe the same event, just with no way
    # to order them within the day.
    has_post_so_move_to_erp = (
        pd.notna(latest_dt)
        and pd.notna(so_date)
        and pd.to_datetime(latest_dt).normalize() >= so_date
    )
    if has_post_so_move_to_erp:
        return "OK — Move after opname"

    if erp_loc_norm == so_loc_norm:
        if qty_zero:
            return "Inaccuracy — ERP still shows stock at technician (Qty=0)"
        return "OK — ERP match"

    if qty_zero:
        return f"Inaccuracy — No post-SO move found (ERP at {erp_loc})"

    if "/customer" in erp_loc_norm:
        return "Inaccuracy — Check WO (already /Customer, no post-SO move)"

    return f"Inaccuracy — Location mismatch (found in {erp_loc})"


def nonimei_status_with_loc(row: pd.Series, idx: dict):
    """
    Decide status for Non-IMEI using pair-wise comparison (Tech x Loc x Product).
    - Skip '/Customer' entirely.
    - If SO Location blank -> Inaccuracy — Missing SO Location (caller handles this;
      not reached from here, kept as a defensive fallback).
    - Else compare Qty vs ERP qty at the same pair.
    Returns (status, erp_qty) or None to indicate 'skipped'.
    """
    prod = str(row.get("Product", ""))
    so = row.get("SO Location", "")
    so_norm = norm_text(so)
    qty = float(row.get("Qty", 0) or 0)

    if "/customer" in so_norm:
        return None

    if so_norm == "":
        return ("Inaccuracy — Missing SO Location", 0.0)

    use_owner = isinstance(so, str) and any(k in so_norm for k in ["teknisi", "return", "refurbishment"])
    sum_no_owner = idx["sum_no_owner"]
    sum_with_owner = idx["sum_with_owner"]
    tech_norm = norm_owner(row.get("technician"))

    if use_owner:
        m = sum_with_owner[(sum_with_owner["Product"].eq(prod)) & (sum_with_owner["Location"].eq(so_norm)) & (sum_with_owner["owner_norm"].eq(tech_norm))]
    else:
        m = sum_no_owner[(sum_no_owner["Product"].eq(prod)) & (sum_no_owner["Location"].eq(so_norm))]

    erp_qty = float(m["erp_qty"].iloc[0]) if not m.empty else 0.0

    if abs(erp_qty - qty) < 1e-9:
        return ("OK — Non-IMEI", erp_qty)
    return (f"Inaccuracy — Qty mismatch at {so}: opname={qty}, ERP={erp_qty}", erp_qty)


# =============================
# KPI
# =============================
def imei_kpi(df: pd.DataFrame) -> pd.DataFrame:
    total = len(df)

    ok_match = (df["status"] == "OK — ERP match").sum()
    ok_move = (df["status"] == "OK — Move after opname").sum()
    ok_websms = (df["status"] == "OK — Installed at Customer (WebSMS)").sum()

    missing = (df["status"] == "Inaccuracy — Missing in ERP").sum()
    mism_loc = df["status"].str.startswith("Inaccuracy — Location mismatch").sum()
    check_wo = df["status"].str.startswith("Inaccuracy — Check WO").sum()
    abnormal = (df["status"] == "Inaccuracy — Check and resolve abnormality").sum()
    missing_imei = (df["status"] == "Inaccuracy — Missing IMEI when expected").sum()

    total_ok = ok_match + ok_move + ok_websms
    accuracy_pct = (total_ok / total) if total > 0 else 0.0

    return pd.DataFrame([{
        "Total": total,
        "OK (ERP match)": ok_match,
        "OK (Move after opname)": ok_move,
        "OK (WebSMS)": ok_websms,
        "Check WO": check_wo,
        "Location mismatch": mism_loc,
        "Missing in ERP": missing,
        "Missing IMEI Input": missing_imei,
        "Check & resolve abnormality": abnormal,
        "Accuracy %": accuracy_pct,
    }])


def _attach_area_to_summary(summary_df: pd.DataFrame, src_rows: pd.DataFrame) -> pd.DataFrame:
    if summary_df.empty or src_rows.empty or "technician" not in src_rows.columns:
        return summary_df.copy()
    area_map = (
        src_rows.groupby("technician")["area"]
        .agg(lambda s: ", ".join(sorted(set(x for x in s if pd.notna(x)))))
        .to_dict()
    )
    out = summary_df.copy()
    out["Area"] = out["technician"].map(area_map).fillna("")
    return out


# =============================
# EXCEL WRITER
# =============================
def export_excel(out_path: Path,
                  detail_imei: pd.DataFrame,
                  detail_non: pd.DataFrame,
                  imei_overall: pd.DataFrame,
                  imei_bytech: pd.DataFrame,
                  non_bytech: pd.DataFrame,
                  scope_debug: Optional[pd.DataFrame] = None) -> None:
    try:
        area_src = []
        if not detail_imei.empty and {"technician", "area"}.issubset(detail_imei.columns):
            area_src.append(detail_imei[["technician", "area"]])
        if not detail_non.empty and {"technician", "area"}.issubset(detail_non.columns):
            area_src.append(detail_non[["technician", "area"]])

        area_map = {}
        if area_src:
            area_df = pd.concat(area_src, ignore_index=True)
            area_df["technician"] = area_df["technician"].astype(str).str.strip()
            area_df["area"] = area_df["area"].astype(str).str.strip()
            area_map = (
                area_df.dropna(subset=["technician"])
                       .groupby("technician")["area"]
                       .agg(lambda s: ", ".join(sorted({x for x in s if x and x != "nan"})))
                       .to_dict()
            )

        if not imei_bytech.empty and "technician" in imei_bytech.columns:
            imei_bytech = imei_bytech.copy()
            imei_bytech["Area"] = imei_bytech["technician"].map(area_map).fillna("")
        if not non_bytech.empty and "technician" in non_bytech.columns:
            non_bytech = non_bytech.copy()
            non_bytech["Area"] = non_bytech["technician"].map(area_map).fillna("")
    except Exception as e:
        print(f"[WARN] Could not attach Area to summaries: {e}")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as writer:
        (detail_imei if not detail_imei.empty else pd.DataFrame(
            columns=["technician_sheet", "technician", "opname_date", "Product", "Lot/Serial Number", "SO Location", "Qty", "status"])
        ).to_excel(writer, sheet_name="Detail (IMEI)", index=False)

        (detail_non if not detail_non.empty else pd.DataFrame(
            columns=["technician_sheet", "technician", "Product", "SO Location", "Qty", "ERP qty at SO", "status"])
        ).to_excel(writer, sheet_name="Detail (Non-IMEI)", index=False)

        imei_overall.to_excel(writer, sheet_name="Summary IMEI Overall", index=False)
        imei_bytech.to_excel(writer, sheet_name="Summary IMEI by Tech", index=False)
        non_bytech.to_excel(writer, sheet_name="Summary Non-IMEI by Tech", index=False)

        if scope_debug is not None and not scope_debug.empty:
            scope_debug.to_excel(writer, sheet_name="Debug — Scope", index=False)

        for sn in ["Detail (IMEI)", "Detail (Non-IMEI)", "Summary IMEI Overall",
                   "Summary IMEI by Tech", "Summary Non-IMEI by Tech"]:
            try:
                writer.sheets[sn].freeze_panes(1, 0)
            except KeyError:
                pass


# =============================
# CLI
# =============================
def parse_args():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--erp-csv", required=True, help="Cleaned ERP CSV from clean_stock_quant.py --location-suffix \"\"")
    p.add_argument("--movement", required=True, help="Stock Movement export (.csv or .xlsx)")
    p.add_argument("--east", required=True, help="'03 East Stock Opname Teknisi.xlsx'")
    p.add_argument("--west", required=True, help="'04 West Stock Opname Teknisi.xlsx'")
    p.add_argument("--masterfile", required=True, help="Inventory Masterfile.xlsx")
    p.add_argument("--device-id", default=None, help="Device ID.xlsx (WebSMS, optional)")
    p.add_argument("--device-sg", default=None, help="Device SG.xlsx (WebSMS, optional)")
    p.add_argument("--output", required=True, help="Path to write the report .xlsx")
    return p.parse_args()


# =============================
# MAIN
# =============================
def main():
    args = parse_args()

    erp_csv = Path(args.erp_csv)
    movement_path = Path(args.movement)
    east_path = Path(args.east)
    west_path = Path(args.west)
    master_file = Path(args.masterfile)
    device_id = Path(args.device_id) if args.device_id else None
    device_sg = Path(args.device_sg) if args.device_sg else None
    out_xlsx = Path(args.output)

    log(f"ERP CSV: {erp_csv}")
    log(f"Movement: {movement_path}")
    log(f"East: {east_path}")
    log(f"West: {west_path}")
    log(f"Masterfile: {master_file}")
    log(f"Output: {out_xlsx}")

    in_scope = build_scope(master_file)

    erp = load_erp_quant(erp_csv, in_scope)
    moves = load_moves(movement_path)
    websms_imeis = load_websms(device_id, device_sg)

    log("Loading teknisi workbooks (East & West)…")
    imei_rows, non_rows = load_teknisi_all(east_path, west_path, in_scope)

    log("Building indexes (IMEI)…")
    idx_imei = build_indexes_imei(erp, moves)
    log("Building indexes (Non-IMEI)…")
    idx_non = build_indexes_nonimei(erp)

    log("Running IMEI decisions…")
    imei_out_rows = []
    if not imei_rows.empty:
        for _, r in imei_rows.iterrows():
            status = imei_status(r, idx_imei, websms_imeis)
            imei_out_rows.append({
                "technician_sheet": r.get("technician_sheet"),
                "technician": r.get("technician"),
                "opname_date": r.get("opname_date"),
                "Product": r.get("Product"),
                "Lot/Serial Number": r.get("Lot/Serial Number"),
                "SO Location": r.get("SO Location"),
                "Qty": r.get("Qty", 0),
                "status": status,
                "area": r.get("area"),
            })

    detail_imei = pd.DataFrame(imei_out_rows)
    log(f"→ IMEI decisions done. Rows: {len(detail_imei)}")

    log("Computing IMEI KPIs…")
    if not detail_imei.empty:
        imei_overall = imei_kpi(detail_imei)
        imei_bytech = (
            detail_imei.groupby("technician", dropna=False)
            .apply(imei_kpi).reset_index(level=1, drop=True).reset_index()
            .rename(columns={"index": "technician"})
        )
    else:
        imei_overall = pd.DataFrame([{
            "Total": 0, "OK (ERP match)": 0, "OK (Move after opname)": 0, "OK (WebSMS)": 0,
            "Check WO": 0, "Location mismatch": 0, "Missing in ERP": 0, "Missing IMEI Input": 0,
            "Check & resolve abnormality": 0, "Accuracy %": 0.0,
        }])
        imei_bytech = pd.DataFrame(columns=[
            "technician", "Total", "OK (ERP match)", "OK (Move after opname)", "OK (WebSMS)",
            "Check WO", "Location mismatch", "Missing in ERP", "Missing IMEI Input",
            "Check & resolve abnormality", "Accuracy %",
        ])
    log("→ IMEI KPIs ready.")

    log("Running Non-IMEI decisions…")
    non_out_rows = []
    if not non_rows.empty:
        for _, r in non_rows.iterrows():
            res = nonimei_status_with_loc(r, idx_non)
            if res is None:
                continue
            status, erp_qty = res
            non_out_rows.append({
                "technician_sheet": r.get("technician_sheet"),
                "technician": r.get("technician"),
                "Product": r.get("Product"),
                "SO Location": r.get("SO Location"),
                "Qty": r.get("Qty", 0),
                "ERP qty at SO": erp_qty,
                "status": status,
                "area": r.get("area"),
            })

    detail_non = pd.DataFrame(non_out_rows)
    log(f"→ Non-IMEI decisions done. Rows: {len(detail_non)}")

    log("Computing Non-IMEI KPIs…")
    if not detail_non.empty:
        by = detail_non.groupby("technician", dropna=False).agg(
            ok=("status", lambda s: (s == "OK — Non-IMEI").sum()),
            total=("status", "count")
        ).reset_index()
        by["% Accuracy"] = by.apply(lambda r: (r["ok"] / r["total"]) if r["total"] else 0.0, axis=1)
    else:
        by = pd.DataFrame(columns=["technician", "ok", "total", "% Accuracy"])
    log("→ Non-IMEI KPIs ready.")

    log(f"Exporting Excel → {out_xlsx}")
    scope_debug = pd.DataFrame({"Product in scope": sorted(list(in_scope))}) if STRICT_SCOPE else None
    export_excel(out_xlsx, detail_imei, detail_non, imei_overall, imei_bytech, by, scope_debug=scope_debug)

    log(f"Report generated: {out_xlsx}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\n❌ Fatal error:\n", e)
        sys.exit(1)
