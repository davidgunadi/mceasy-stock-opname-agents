#!/usr/bin/env python3
"""
odoo_client.py

Scaffold for pulling Price and Category directly from Odoo, to eventually
replace the Inventory Masterfile as the source of those two fields for
scripts/compare_stock_opname.py (and, later, compare_stock_opname_teknisi.py).

STATUS: scaffold only. Nothing calls this yet — the Inventory Masterfile
remains the live source of Price/Category until this is validated against
real Odoo data and explicitly wired into the comparator script(s). Not a
functional pipeline change by itself; no version bump for this file alone.

## Credentials

Never hardcode credentials here. Reads them from environment variables:
  ODOO_URL       the JSON-RPC endpoint itself, e.g.
                 "https://phoenix.mceasy.cloud/jsonrpc/" -- not a bare host,
                 and not the /xmlrpc/2/* path pair some Odoo docs show.
  ODOO_DB        the Odoo database name
  ODOO_USERNAME  login used to authenticate (this instance uses a numeric
                 user id, not an email)
  ODOO_API_KEY   an API key (Settings > Users & Companies > Users > your user
                 > Account Security > API Keys) -- NOT your login password

Set these in your own shell/profile (or a local, gitignored .env you load
yourself) before running. This script will refuse to run with any of them
missing rather than silently falling back to something insecure.

## Confirmed against the real Odoo instance (2026-09-22)

  - Model: `product.template`.
  - Price: NOT a field on product.template -- it's the MINIMUM `price` across
    that product's `product.supplierinfo` records (vendor pricelist lines),
    i.e. the cheapest vendor price currently on file. A product with no
    supplierinfo lines at all has no price this way (see `fetch_product_info`
    for how that's surfaced).
  - Category: derived FROM that same price (a price-based tier: C/B/A), not
    read from Odoo's `categ_id` field. Bands given 2026-09-22 as a
    placeholder -- **confirmed changeable, not final** -- see
    `category_from_price`.

## Product lookup key: canonical_name(), confirmed 2026-09-22

Odoo's own `name` field does NOT carry the "[1011] " internal-reference
prefix the ERP/physical Stock Opname data (and the old Inventory Masterfile)
use for products that have one -- that prefix comes from a separate field,
`default_code`. `canonical_name()` reconstructs it: `"[1011] GPS WANWAY EV02"`
from `name="GPS WANWAY EV02"` + `default_code="1011"`; a product with no
default_code just uses its plain name.

This matters a lot: of category 34's 481 products, 50 have a default_code.
Using Odoo's bare `name` as the whitelist entry for those 50 means NONE of
them ever match an ERP/physical row (which use the bracketed form) --
confirmed this caused a real regression (0/19 physical IMEI-sheet products
matched, vs 2152/2152 under the old masterfile) before canonical_name() was
introduced. Every whitelist-facing function here (`fetch_products_in_category`,
`fetch_in_scope_products`) returns/uses canonical_name, not bare `name`.

## Scope (replaces the Inventory Masterfile whitelist)

The Inventory Masterfile's "Category" sheet used to BE the whitelist (only
products listed there were in scope). That's flipped: the whitelist is now
"every Odoo product.template with `categ_id` in CATEGORY_IDS`, MINUS anything
named in a separate exclusion workbook (e.g. "Stock Opname Exclude Item.xlsx"
-- single sheet, single "Product" column of names to drop). See
`fetch_in_scope_products`. `CATEGORY_IDS = [34]`, confirmed 2026-09-22 --
not yet independently re-verified as "the" right category if more than
stock-opname-relevant items also happen to sit in categ_id 34.

Usage (standalone smoke test once ODOO_* env vars are set, e.g. via a local
.env file -- see `.env` in the repo root, loaded automatically if
`python-dotenv` is installed):
    python scripts/odoo_client.py --product "GPS UNIT A" --product "GPS UNIT B"
    python scripts/odoo_client.py --scope --exclusion-xlsx "path/to/Stock Opname Exclude Item.xlsx"
"""

import argparse
import json
import os
import re
import sys
import urllib.error
import urllib.request

import openpyxl

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

PRODUCT_MODEL = "product.template"
SUPPLIERINFO_MODEL = "product.supplierinfo"
CATEGORY_IDS = [34]

REQUIRED_ENV_VARS = ("ODOO_URL", "ODOO_DB", "ODOO_USERNAME", "ODOO_API_KEY")


class OdooConfigError(RuntimeError):
    pass


def load_config():
    values = {name: os.environ.get(name) for name in REQUIRED_ENV_VARS}
    missing = [name for name, val in values.items() if not val]
    if missing:
        raise OdooConfigError(
            f"Missing required environment variable(s): {', '.join(missing)}. "
            "Set these in your own shell/profile, or fill in the repo's .env "
            "file -- never pass real credentials as command-line arguments "
            "or hardcode them here."
        )
    return values["ODOO_URL"], values["ODOO_DB"], values["ODOO_USERNAME"], values["ODOO_API_KEY"]


def _jsonrpc_call(url, service, method, args):
    """POST a single JSON-RPC 2.0 call to Odoo's unified /jsonrpc endpoint
    (this instance's ODOO_URL points directly at it, e.g.
    "https://phoenix.mceasy.cloud/jsonrpc/") -- as opposed to the separate
    /xmlrpc/2/common + /xmlrpc/2/object endpoints Odoo also exposes."""
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {"service": service, "method": method, "args": args},
        "id": 1,
    }
    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            body = json.loads(resp.read().decode("utf-8"))
    except urllib.error.URLError as e:
        raise OdooConfigError(f"Could not reach {url}: {e}")
    if body.get("error"):
        raise OdooConfigError(f"Odoo JSON-RPC error calling {service}.{method}: {body['error']}")
    return body.get("result")


class JsonRpcModels:
    """Thin execute_kw wrapper matching xmlrpc.client.ServerProxy's shape, so
    the rest of this file reads the same regardless of transport."""

    def __init__(self, url):
        self.url = url

    def execute_kw(self, db, uid, api_key, model, method, args, kwargs=None):
        return _jsonrpc_call(self.url, "object", "execute_kw", [db, uid, api_key, model, method, args, kwargs or {}])


def connect():
    """This instance's ODOO_USERNAME is already a numeric user id (not a
    login/email) -- confirmed 2026-09-22 that common.authenticate rejects it
    as a login (returns `false`, no error), but calling execute_kw directly
    with it as uid + the API key as password works. So there's no separate
    login step here: ODOO_USERNAME IS the uid."""
    url, db, username, api_key = load_config()
    try:
        uid = int(username)
    except ValueError:
        raise OdooConfigError(f"ODOO_USERNAME must be a numeric user id for this instance, got: {username!r}")
    return db, uid, api_key, JsonRpcModels(url)


def category_from_price(price):
    """Price-based tier (IDR): C <=150,000, B 150,001-600,000, A >=600,001.
    Placeholder bands given 2026-09-22 by the business owner -- explicitly
    confirmed changeable, not final. Returns None for price=None (nothing to
    classify)."""
    if price is None:
        return None
    if price <= 150000:
        return "C"
    if price <= 600000:
        return "B"
    return "A"


# Confirmed 2026-09-22 against real supplierinfo data: prices aren't all IDR
# -- found live USD lines (e.g. "DASHCAM HOWEN MC30-01 H", 70.0 USD) and one
# CNY line ("Bundle HIKVISION 24 Channel MVR AE-MH1212", 17571.0 CNY). Anything
# not in this map is deliberately NOT guessed at -- converting it wrong would
# silently corrupt the min-price comparison, so those lines are excluded from
# the min instead (see fetch_product_info's skipped-line warning).
CURRENCY_TO_IDR_RATE = {
    "IDR": 1,
    "USD": 16000,
    "CNY": 2500,
}


BRACKET_PREFIX_RE = re.compile(r"^\[[^\]]*\]\s*")


def canonical_name(record):
    """Reconstructs the "[code] Name" convention the ERP/physical Stock
    Opname data (and the old Inventory Masterfile) use for products that
    carry an internal reference code -- Odoo's `default_code`. A product
    with no default_code uses its plain `name` unchanged. See the module
    docstring's "Product lookup key" section for why this matters."""
    code = record.get("default_code")
    return f"[{code}] {record['name']}" if code else record["name"]


def _min_price_by_tmpl_id(db, uid, api_key, models, tmpl_ids):
    """Shared by fetch_price_by_tmpl_ids and fetch_product_info: min
    product.supplierinfo price per template id, converted to IDR via
    CURRENCY_TO_IDR_RATE. A line in a currency not in that map is skipped
    (with a stderr warning) rather than treated as IDR. An id with no usable
    supplierinfo line at all is simply absent from the result (not 0 -- 0
    would misleadingly look like a real "free" price)."""
    tmpl_ids = list(set(tmpl_ids))
    if not tmpl_ids:
        return {}
    supplier_lines = models.execute_kw(
        db, uid, api_key, SUPPLIERINFO_MODEL, "search_read",
        [[("product_tmpl_id", "in", tmpl_ids)]],
        {"fields": ["product_tmpl_id", "price", "currency_id"]},
    )
    min_price = {}
    for line in supplier_lines:
        tmpl_id = line["product_tmpl_id"][0] if isinstance(line["product_tmpl_id"], (list, tuple)) else line["product_tmpl_id"]
        price = line.get("price")
        if price is None:
            continue
        currency = line.get("currency_id")
        currency_name = currency[1] if isinstance(currency, (list, tuple)) else currency
        rate = CURRENCY_TO_IDR_RATE.get(currency_name)
        if rate is None:
            print(
                f"WARNING: skipping a product.template id={tmpl_id} supplierinfo line "
                f"priced in unconvertible currency {currency_name!r} (price={price}) -- "
                "add it to CURRENCY_TO_IDR_RATE if this should count.",
                file=sys.stderr,
            )
            continue
        price_idr = price * rate
        if tmpl_id not in min_price or price_idr < min_price[tmpl_id]:
            min_price[tmpl_id] = price_idr
    return min_price


def fetch_price_by_tmpl_ids(tmpl_ids):
    """tmpl_ids: iterable of product.template ids. Returns {tmpl_id: price_idr}
    -- see _min_price_by_tmpl_id for how price is computed."""
    db, uid, api_key, models = connect()
    return _min_price_by_tmpl_id(db, uid, api_key, models, tmpl_ids)


def fetch_product_info(product_names):
    """product_names: names to look up -- either Odoo's plain `name` or the
    "[code] Name" canonical form (see canonical_name()); both work, matched
    by stripping any bracket prefix before searching Odoo's `name` field.

    Returns {input_name: {"price": float | None, "category": str | None}},
    keyed by exactly the strings passed in `product_names` (so a canonical
    bracket-form input gets that same string back as its key). A name with
    no product.template match at all is absent from the result entirely."""
    db, uid, api_key, models = connect()

    search_name_of = {n: BRACKET_PREFIX_RE.sub("", n).strip() for n in product_names}
    search_names = list(set(search_name_of.values()))
    if not search_names:
        return {}

    templates = models.execute_kw(
        db, uid, api_key, PRODUCT_MODEL, "search_read",
        [[("name", "in", search_names)]], {"fields": ["name"]},
    )
    by_search_name = {t["name"]: t for t in templates}
    price_by_tmpl_id = _min_price_by_tmpl_id(db, uid, api_key, models, [t["id"] for t in templates])

    result = {}
    for original, search_name in search_name_of.items():
        t = by_search_name.get(search_name)
        if t is None:
            continue
        price = price_by_tmpl_id.get(t["id"])
        result[original] = {"price": price, "category": category_from_price(price)}
    return result


def fetch_products_in_category(category_ids=CATEGORY_IDS):
    """All product.template rows whose categ_id is in category_ids. Returns a
    list of {"id": int, "name": str, "default_code": str | False,
    "canonical_name": str} dicts -- canonical_name is the whitelist basis
    (see canonical_name())."""
    db, uid, api_key, models = connect()
    records = models.execute_kw(
        db, uid, api_key, PRODUCT_MODEL, "search_read",
        [[("categ_id", "in", category_ids)]], {"fields": ["name", "default_code"]},
    )
    for r in records:
        r["canonical_name"] = canonical_name(r)
    return records


def load_exclusion_names(path):
    """Reads the exclusion workbook (single sheet, single "Product" column of
    names to drop from scope) and returns a set of names normalized the same
    way build_product_lookup() in compare_stock_opname.py normalizes product
    names elsewhere in this pipeline (stripped, lowercased) -- so membership
    checks are case-insensitive and whitespace-tolerant.

    Some rows carry an internal-reference prefix like "[1001] GPS CONCOX
    GT06N" -- stripped here so membership checks compare against the same
    bracket-stripped form used when matching category products (see
    fetch_in_scope_products)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    rows = ws.iter_rows(min_row=2, max_col=1, values_only=True)
    names = set()
    for (name,) in rows:
        if name is None or not str(name).strip():
            continue
        cleaned = BRACKET_PREFIX_RE.sub("", str(name)).strip().lower()
        names.add(cleaned)
    return names


def fetch_in_scope_products(exclusion_xlsx_path, category_ids=CATEGORY_IDS, require_default_code=False):
    """The whitelist: every category_ids product (by canonical_name, so
    bracket-coded devices keep their "[code] Name" form -- see the module
    docstring), minus anything named in the exclusion workbook. Returns the
    full product records (not just names) so a caller can look up price by
    id afterward without a second name-based round trip.

    require_default_code=True restricts to products that carry an internal
    reference code (bracket-coded devices only) -- used by
    /stock-opname-teknisi, which per the business owner (2026-09-22) only
    tracks serialized/internal-reference devices at the technician level,
    not bulk/accessory items. /stock-opname (per-city) leaves this False."""
    category_products = fetch_products_in_category(category_ids)
    if require_default_code:
        category_products = [p for p in category_products if p.get("default_code")]
    excluded = load_exclusion_names(exclusion_xlsx_path)
    return [
        p for p in category_products
        if BRACKET_PREFIX_RE.sub("", p["canonical_name"]).strip().lower() not in excluded
    ]


def parse_args():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--product", action="append", help="Product name to look up price/category for (repeatable)")
    p.add_argument("--scope", action="store_true", help="Instead of --product lookups, print the new in-scope product list (category filter minus exclusion sheet)")
    p.add_argument("--exclusion-xlsx", help="Path to the exclusion workbook (required with --scope)")
    args = p.parse_args()
    if args.scope and not args.exclusion_xlsx:
        p.error("--scope requires --exclusion-xlsx")
    if not args.scope and not args.product:
        p.error("either --product (repeatable) or --scope --exclusion-xlsx is required")
    return args


def main():
    args = parse_args()
    try:
        if args.scope:
            category_products = fetch_products_in_category()
            excluded = load_exclusion_names(args.exclusion_xlsx)
            in_scope = fetch_in_scope_products(args.exclusion_xlsx)
            excluded_matched = {
                p["canonical_name"] for p in category_products
                if BRACKET_PREFIX_RE.sub("", p["canonical_name"]).strip().lower() in excluded
            }
            print(f"Category {CATEGORY_IDS} products in Odoo: {len(category_products)}")
            print(f"Exclusion sheet names: {len(excluded)}")
            print(f"  of which matched a category product: {len(excluded_matched)}")
            print(f"In-scope products (category minus exclusion): {len(in_scope)}")
            return
        info = fetch_product_info(args.product)
    except OdooConfigError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)

    for name in args.product:
        if name not in info:
            print(f"{name}: NOT FOUND in Odoo ({PRODUCT_MODEL})")
            continue
        price = info[name]["price"]
        price_str = price if price is not None else "(no supplierinfo lines)"
        print(f"{name}: price={price_str}  category={info[name]['category']}")


if __name__ == "__main__":
    main()
