# [Your Team Name] Claude Agents

This project uses Claude Code's multi-agent system to automate [describe what your pipeline does].

---

## Agents

Agents live in `.claude/agents/`. Each agent is a markdown file with a YAML frontmatter header.

| Agent | Model | Role |
|---|---|---|
| `stock-quant-cleaner` | sonnet | Runs `scripts/clean_stock_quant.py` over a Stock Quant CSV export, reports the abnormality breakdown, and delivers the annotated `*/Stock`-only output |
| `stock-opname-comparator` | sonnet | Runs `scripts/compare_stock_opname.py`: reconciles cleaned ERP stock (per city) against a physical Stock Opname count, and delivers a Summary/Product Summary/IMEI Mismatch/Aging Report/Aging Report Detail report |
| `stock-opname-teknisi-comparator` | sonnet | Runs `scripts/compare_stock_opname_teknisi.py`: reconciles the full ERP ledger against a technician-level physical Stock Opname count (East & West), applies the WebSMS installed-at-customer override, and delivers a Detail IMEI/Detail Non-IMEI/Summary report |

Add or remove rows as you define your own agents.

---

## Skills (Slash Commands)

Skills live in `.claude/skills/`. Each skill is a folder with a `SKILL.md` file inside.

| Command | What it triggers |
|---|---|
| `/stock-opname` | Chains `@stock-quant-cleaner` → `@stock-opname-comparator` in one pass: cleans a raw Stock Quant CSV export, then immediately reconciles the result against a physical Stock Opname count for one city |
| `/stock-opname-teknisi` | Runs `@stock-opname-teknisi-comparator`: reconciles ERP stock against a technician-level physical Stock Opname count (East & West) |

Type the command in Claude Code to start the pipeline.

---

## Pipeline Order

Define the order your agents run. Example:

```
User input
    ↓
@researcher
    ↓
@writer
    ↓
@reviewer
    ↓
⏸ Pause — human reviews
    ↓
@formatter
    ↓
Output file saved to ./outputs/
```

**`/stock-opname-teknisi` pipeline** (single agent, fully sequential — needs its own
ERP cleansing pass, run with the location filter disabled, before comparing):

```
User provides: East + West teknisi workbooks + exclusion xlsx (same one
                /stock-opname uses) + Stock Quant export (raw, or
                already-cleaned by @stock-quant-cleaner) + Stock Movement
                export + (optional) WebSMS Device ID/SG files
                (no Inventory Masterfile — fully replaced by Odoo, 2026-09-22)
    ↓
@stock-opname-teknisi-comparator
    ├─ if given a raw Stock Quant export: run scripts/clean_stock_quant.py
    │  --location-suffix "" (full ledger, not just */Stock — this pipeline
    │  needs to know where an item currently sits anywhere, e.g. /Customer
    │  or /Teknisi)
    ├─ pull product scope AND price live from Odoo (scripts/odoo_client.py,
    │  require_default_code=True — internal-reference/bracket-coded devices
    │  only, minus the exclusion xlsx)
    ├─ run scripts/compare_stock_opname_teknisi.py (one pass: drop any
    │  abnormality-flagged ERP row, split teknisi rows into IMEI vs
    │  Non-IMEI, apply movement-then-location IMEI rules and qty-based
    │  Non-IMEI rules, apply the WebSMS override — each stage depends on
    │  the last, so this is one script run, not split further)
    └─ report per-technician accuracy + deliver the 4-sheet report
       (Detail IMEI, Detail Non-IMEI, Summary IMEI Overall, Summary by Tech)
    ↓
Output delivered directly to the user (saved next to the movement/opname files,
"Stock Opname Teknisi_<date>.xlsx")
```

**`/stock-opname` pipeline** (two agents, strictly sequential — this skill exists to
remove the manual hand-off between cleaning and comparing; neither agent's own
logic changes):

```
User provides: raw Stock Quant CSV + city + physical Stock Opname xlsx +
                exclusion xlsx (e.g. "Stock Opname Exclude Item.xlsx")
    ↓
@stock-quant-cleaner
    ├─ validate input columns
    ├─ run scripts/clean_stock_quant.py
    └─ report abnormality breakdown + deliver annotated CSV
    ↓ (cleaned CSV path passed straight through, no re-asking the user)
@stock-opname-comparator
    ├─ pull product scope + price live from Odoo (scripts/odoo_client.py,
    │  requires ODOO_URL/ODOO_DB/ODOO_USERNAME/ODOO_API_KEY -- see the repo's
    │  .env), minus anything named in the exclusion xlsx
    ├─ run scripts/compare_stock_opname.py against the cleaned CSV + city +
    │  opname xlsx + exclusion xlsx
    └─ report Summary + deliver the 5-sheet report (Summary, Product Summary,
       IMEI Mismatch, Aging Report, Aging Report Detail)
    ↓
Both outputs delivered directly to the user (each saved next to its respective
input file, per the exception below)
```

---

## Rules

- All output files must be saved to `./outputs/` — never write to the repo root.
  Exception: `/stock-opname` and `/stock-opname-teknisi` intentionally save
  their output next to the user's input file (not `./outputs/`), since the
  input itself lives outside the repo (e.g. the user's Downloads folder) —
  this was an explicit decision, not an oversight.
- Any script invoked by an agent while running must live in `./scripts/` — never write scripts to the repo root
- Never commit real credentials. `.env` (Odoo API credentials for
  `scripts/odoo_client.py`) is gitignored — if it's ever missing, ask the
  user to fill it in themselves rather than requesting the values in chat.
- [Add your own rules here]

---

## Versioning

This repo uses [CHANGELOG.md](CHANGELOG.md) (Keep a Changelog format) and Semantic
Versioning (`MAJOR.MINOR.PATCH`).

**What counts as a functional change** (requires a version bump — this is a
template repo, so "functional" means it changes what happens when someone
clones and runs it):

- `.claude/agents/*.md` — adding, removing, or changing an agent's behavior, model, or tools
- `.claude/skills/**/SKILL.md` — adding, removing, or changing a slash command's pipeline
- `.claude/settings.json` — permission or hook changes
- Root `CLAUDE.md` — changes to the pipeline order or rules (the sections above this one)

**Not a functional change** (no bump needed): README prose/formatting, `.gitignore`,
comments, typo fixes in docs, or anything else that doesn't change what an
agent/skill actually does when run.

**Version bump rules:**
- **MAJOR** — breaking or workflow-changing (e.g. renamed/removed agent or skill, changed pipeline order)
- **MINOR** — new capability (new agent, new skill, new tool granted)
- **PATCH** — fix or tweak to existing agent/skill behavior

**On every commit with a functional change**, proactively (without being asked):
1. Add an entry under `## [Unreleased]` in [CHANGELOG.md](CHANGELOG.md) (Added/Changed/Fixed/Removed), then move it under a new version heading dated with the commit date.
2. Bump the version in `CHANGELOG.md` and in the `**Version:**` line in `README.md` — keep both in sync.

---

## Context

- [Describe your team, product, or domain here so agents have grounding]
- [Add any terminology, personas, or constraints that agents should know]

### Stock Quant cleaning domain

- The company that should own everything sitting in the warehouse is exactly
  **"PT Otto Menara Globalindo"**. Any other `Owner` value on a `*/Stock` line is a
  mismatch.
- Warehouse (on-hand) locations follow the pattern `<Warehouse>/Stock` (e.g.
  `SBY/Stock`, `JKT/Stock`). Other locations (Customer, Return, Teknisi,
  Refurbishment, Transit, Partners/*, etc.) are history/movement locations, not
  on-hand stock.
- Stock Quant exports from the ERP are a **historical ledger**, not a
  point-in-time snapshot — the same lot/serial number legitimately reappears
  across many rows as a unit moves through Vendor → Customer → Refurbishment →
  Teknisi → Stock over its lifetime, usually with `Quantity = 0` on the older
  rows. Don't treat "same lot appears on another line" as inherently abnormal —
  see `scripts/clean_stock_quant.py`'s docstring for the exact duplicate rule
  that accounts for this.
- Full abnormality rule definitions live in `scripts/clean_stock_quant.py`
  (used by `@stock-quant-cleaner`, invoked via `/stock-opname` or
  `/stock-opname-teknisi`). Changing any rule there is a functional change —
  bump the version per the Versioning section below.
- `/stock-opname`'s second step (`@stock-opname-comparator`) reconciles the
  cleaned ERP output against a physical Stock Opname count for one city. Any
  ERP row carrying a non-empty `Abnormality` flag (Owner Mismatch,
  Negative Qty, Imei Length Difference, Alphanumeric IMEI, Has Special
  Character, Missing Imei, ERP Duplicate line — not just duplicates) is
  excluded from the ERP "on-hand" side before matching, for both IMEI-lot
  matching and the non-IMEI qty sum — confirmed against real SBY 260826 data
  that a narrower duplicate-only exclusion let other abnormalities (e.g. a
  Negative Qty/Owner Mismatch line owned by a courier company) distort the
  non-IMEI qty comparison. Full logic lives in
  `scripts/compare_stock_opname.py` (requires `openpyxl`, see
  `scripts/requirements.txt`).
- **Product scope, price, and category no longer come from a local Inventory
  Masterfile file** (replaced 2026-09-22). `scripts/odoo_client.py` pulls all
  three live from Odoo (JSON-RPC, see its docstring for the endpoint/auth
  shape and required `ODOO_URL`/`ODOO_DB`/`ODOO_USERNAME`/`ODOO_API_KEY` env
  vars — a local, gitignored `.env` in the repo root holds these):
    - **Scope (whitelist)**: every `product.template` with `categ_id` in
      `odoo_client.CATEGORY_IDS` (currently `[34]`), minus anything named in
      the user-provided exclusion workbook (single sheet, single "Product"
      column — e.g. "Stock Opname Exclude Item.xlsx"). Products are matched
      by `odoo_client.canonical_name()` — Odoo's plain `name` plus a
      `"[code] "` prefix reconstructed from `default_code` for the ~50
      category-34 products that have one. This matters: Odoo's bare `name`
      never carries that bracket, but the ERP/physical Stock Opname data (and
      the old masterfile) always did for those products — using the bare name
      as the whitelist entry silently drops every one of them from matching.
    - **Price**: the minimum `product.supplierinfo` price per product,
      converted to IDR via `odoo_client.CURRENCY_TO_IDR_RATE` (`USD` ×16000,
      `CNY` ×2500 — confirmed both currencies appear in real supplierinfo
      data; any other currency is skipped with a warning rather than treated
      as IDR). A product with no usable supplierinfo price defaults to price
      `0`.
    - **Category / `New Group - July 25`**: both get the same value, a
      price-tier from `odoo_client.category_from_price` — `C` ≤150,000,
      `B` 150,001–600,000, `A` ≥600,001 IDR (including the price-`0` default
      above, which lands in `C`). **Placeholder bands given by the business
      owner on 2026-09-22 — explicitly confirmed changeable, not final.**
- `@stock-opname-comparator`'s report also includes an `Aging Report` /
  `Aging Report Detail` pair: whole-month age from each ERP row's `Lot
  Created Date` to the opname date, bucketed into `< 3 Mo / 3-6 Mo / 6-9 Mo /
  9-12 Mo / > 12 Mo`, built from the same scoped, non-abnormal rows used
  everywhere else in the report. Ported from a business-owner-provided
  reference notebook (`v2_stock_opname.ipynb`) where this logic existed but
  was never wired up — the app was silently missing this until it was ported
  into `scripts/compare_stock_opname.py` on 2026-09-22.

### Stock Opname Teknisi domain

- `/stock-opname-teknisi` reconciles ERP stock against a **technician-level**
  physical count (as opposed to `/stock-opname`'s per-city warehouse count).
  Ported from a business-owner-provided reference script ("stock opname
  teknisi v3"); full logic lives in `scripts/compare_stock_opname_teknisi.py`.
- **Product scope AND price now both come from Odoo** (changed 2026-09-22,
  same `scripts/odoo_client.py` as `/stock-opname` — the Inventory
  Masterfile is no longer read by this pipeline at all). Scope is
  `odoo_client.fetch_in_scope_products(..., require_default_code=True)` —
  only products carrying an internal reference code (bracket-coded devices,
  e.g. `[1011] GPS WANWAY EV02`), minus the same exclusion workbook
  `/stock-opname` uses. Per the business owner: this report only tracks
  serialized/internal-reference devices at the technician level, not
  bulk/accessory items — confirmed against real 260826 data (25 in-scope
  products, all 2252 resulting Detail rows are IMEI rows, zero Non-IMEI, and
  every row resolved a price with none falling back to 0). Price is the
  minimum `product.supplierinfo` price per product, currency-converted, same
  as `/stock-opname`. Switching price sources changed real numbers (Total
  Discrepancy Value moved from 238,938,943 to 266,484,470 IDR on the same
  260826 data) — expected, since Odoo's live supplierinfo prices genuinely
  differ from the old masterfile's static "Price in IDR" column, not a bug.
- Unlike `@stock-quant-cleaner`'s default use in `/stock-opname`, this pipeline needs
  `clean_stock_quant.py --location-suffix ""` (every location kept, not just
  `*/Stock`) — it needs to know where an item currently sits *anywhere* in the
  ledger (e.g. still at `/Customer` or `/Teknisi`). `--owner-check-suffix`
  stays at its default (`/Stock`) regardless — Owner Mismatch is scoped to
  on-hand warehouse locations on purpose, so it doesn't fire on every
  legitimately customer/technician-owned historical row once the location
  filter is opened up. (Empirically: with the filter open on real data, Owner
  Mismatch went from 224k false positives down to 96 real ones once this
  scoping was applied — don't remove it.)
- Any ERP row with a non-empty `Abnormality` (from `clean_stock_quant.py`) is
  excluded from the ERP lookup entirely, matching the reference script's
  `Has Abnormality` filter.
- Two fixed technician workbooks: `03 East Stock Opname Teknisi.xlsx` and
  `04 West Stock Opname Teknisi.xlsx`. Each sheet = one technician (`B1` =
  opname date, `B2` = technician name). Rows split into **IMEI** (has a
  Lot/Serial Number) vs **Non-IMEI** (qty-only), validated with different rules
  — see `scripts/compare_stock_opname_teknisi.py`'s docstring for the exact
  IMEI (movement-then-location) and Non-IMEI (qty-per Tech×Location×Product)
  decision trees.
- **WebSMS override**: `Device ID.xlsx` / `Device SG.xlsx` (sheet `Perangkat`,
  column `IMEI`) is a separate source of truth for "installed at customer" —
  an IMEI found there overrides an otherwise-mismatched status to
  `OK — Installed at Customer (WebSMS)`. Optional input; requires
  `python-calamine` to read (their style XML doesn't parse with plain
  `openpyxl`).
- A blank `SO Location` on a Non-IMEI row reports
  `Inaccuracy — Missing SO Location` (mirrors the IMEI side's equivalent
  guard) — the original reference script crashed on this case instead.
