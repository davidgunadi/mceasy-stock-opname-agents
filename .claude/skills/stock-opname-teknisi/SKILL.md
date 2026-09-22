---
name: stock-opname-teknisi
description: >
  Reconcile ERP inventory against a physical Stock Opname count done by
  technicians (East & West), producing a Detail IMEI / Detail Non-IMEI /
  Summary report. Invoke when the user wants to compare ERP stock vs a
  technician-level physical stock count (not a per-city warehouse count —
  that's /compare-stock-opname).
---

Greet the user and explain the pipeline before starting:

"I'll reconcile ERP stock against your technicians' physical Stock Opname count using a single pass:

1. 🔎 **@stock-opname-teknisi-comparator** — cleans the ERP export (full ledger, not just on-hand stock), pulls the in-scope product list AND price live from Odoo (internal-reference/bracket-coded devices only, minus the exclusion workbook), matches it against each technician's IMEI and Non-IMEI counts, applies the WebSMS override where available, and delivers a report (Detail IMEI, Detail Non-IMEI, Summary IMEI Overall, Summary by Tech).

Before we start, I need:
- The **East** and **West** technician workbooks (`03 East Stock Opname Teknisi.xlsx`, `04 West Stock Opname Teknisi.xlsx`)
- The **exclusion workbook** (e.g. "Stock Opname Exclude Item.xlsx" — same one used by `/stock-opname`)
- A **Stock Quant** ERP export (raw csv/xlsx, or an already-cleaned CSV from `/clean-stock-quant`)
- The **Stock Movement** export (csv or xlsx)
- (Optional but recommended) **Device ID.xlsx** and **Device SG.xlsx** (WebSMS) — without these, devices already installed at a customer but not yet reflected in ERP will show as mismatches instead of being auto-resolved."

No Inventory Masterfile input anymore — product scope and price both come live from Odoo (`scripts/odoo_client.py`, requiring `ODOO_URL`/`ODOO_DB`/`ODOO_USERNAME`/`ODOO_API_KEY` in the repo's `.env`).

Once the user provides these, invoke `@stock-opname-teknisi-comparator` with all of them. This is a single-agent pipeline — the decision logic (IMEI movement/location rules, Non-IMEI qty comparison, WebSMS override) is fully deterministic and lives in `scripts/compare_stock_opname_teknisi.py`, not in agent judgment, so there's no fan-out or review loop.

If the user gives a raw Stock Quant export (not already cleaned), the agent runs `clean_stock_quant.py --location-suffix ""` itself first — don't ask the user to run `/clean-stock-quant` separately for this pipeline, since the location-suffix needs to be empty here (different from that skill's own default), which would produce the wrong input if run standalone.

If the user asks for the comparison rules themselves to change (a different movement-vs-location precedence, a different Non-IMEI comparison key, disabling the WebSMS override, etc.), don't have `@stock-opname-teknisi-comparator` improvise — confirm the exact new rule with the user first, then edit `scripts/compare_stock_opname_teknisi.py` directly, and log it as a functional change (CHANGELOG entry + version bump) per `CLAUDE.md`.

After `@stock-opname-teknisi-comparator` finishes, it will have already delivered the file and summary — no further steps needed unless the user asks a follow-up.
