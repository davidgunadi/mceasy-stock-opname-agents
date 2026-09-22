---
name: stock-opname
description: >
  Reconcile ERP inventory against a physical Stock Opname count for one city:
  cleans a raw Stock Quant CSV export, then immediately matches the cleaned
  result against the physical count, producing a Summary / Product Summary /
  IMEI Mismatch / Aging Report / Aging Report Detail report. Invoke when the
  user wants to go from a raw Stock Quant export to a per-city Stock Opname
  reconciliation report (for a technician-level count instead, use
  /stock-opname-teknisi).
---

Greet the user and explain the pipeline before starting:

"I'll clean your Stock Quant export and reconcile it against the physical Stock Opname count in one pass:

1. 🧹 **@stock-quant-cleaner** — runs the cleaning script, reports the abnormality breakdown, and delivers the annotated CSV (filtered to `*/Stock` lines).
2. 🔎 **@stock-opname-comparator** — takes that cleaned CSV, matches it against the physical Stock Opname count by lot/serial number, and delivers a 5-sheet report (Summary, Product Summary, IMEI Mismatch, Aging Report, Aging Report Detail).

Before we start, I need:
- The **raw Stock Quant CSV** export
- **City** (e.g. SBY, JKT)
- The **physical Stock Opname Excel file** for that city
- The **exclusion Excel file** (e.g. "Stock Opname Exclude Item.xlsx" — products to drop from scope; product scope, price, and category now come live from Odoo, not a local masterfile)"

Once the user provides all four inputs, run the two steps strictly in sequence — do not ask the user to re-invoke a second command in between:

1. Invoke `@stock-quant-cleaner` with the raw Stock Quant CSV path. Let it validate, run `scripts/clean_stock_quant.py`, and report the abnormality breakdown as it normally would.
2. Take the cleaned CSV path `@stock-quant-cleaner` produced (its stdout/summary names the output file — don't ask the user for it, and don't re-derive it yourself) and immediately invoke `@stock-opname-comparator` with that cleaned CSV, the city, the physical Stock Opname Excel file, and the exclusion Excel file.
3. Let `@stock-opname-comparator` run `scripts/compare_stock_opname.py` (which pulls product scope/price/category from Odoo via `scripts/odoo_client.py`, requiring `ODOO_URL`/`ODOO_DB`/`ODOO_USERNAME`/`ODOO_API_KEY` in the repo's `.env`) and report its summary as it normally would.

This is two deterministic single-agent steps chained together — there's no fan-out, review loop, or judgment call about *how* to clean or compare; that logic is fully owned by `scripts/clean_stock_quant.py` and `scripts/compare_stock_opname.py` respectively. Your only job here is to hand the cleaned CSV from one agent to the next without a manual round-trip through the user.

If either script fails (bad columns, missing sheet pair, import error), stop and surface the actual error — don't attempt the next step on a failed or partial output.

If the user asks for the cleaning rules or the comparison rules themselves to change, don't have either agent improvise — confirm the exact new rule with the user first, then edit the relevant script (`scripts/clean_stock_quant.py` or `scripts/compare_stock_opname.py`) directly, and log it as a functional change (CHANGELOG entry + version bump) per `CLAUDE.md`.

After `@stock-opname-comparator` finishes, both files (cleaned CSV and the reconciliation report) will already have been delivered by their respective agents — no further steps needed unless the user asks a follow-up.
