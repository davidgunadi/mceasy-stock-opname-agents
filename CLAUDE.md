# [Your Team Name] Claude Agents

This project uses Claude Code's multi-agent system to automate [describe what your pipeline does].

---

## Agents

Agents live in `.claude/agents/`. Each agent is a markdown file with a YAML frontmatter header.

| Agent | Model | Role |
|---|---|---|
| `stock-quant-cleaner` | sonnet | Runs `scripts/clean_stock_quant.py` over a Stock Quant CSV export, reports the abnormality breakdown, and delivers the annotated `*/Stock`-only output |
| `stock-opname-comparator` | sonnet | Runs `scripts/compare_stock_opname.py`: reconciles cleaned ERP stock (per city) against a physical Stock Opname count, and delivers a Summary/Product Summary/IMEI Mismatch report |
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
User provides: East + West teknisi workbooks + Inventory Masterfile +
                Stock Quant export (raw, or already-cleaned by @stock-quant-cleaner) +
                Stock Movement export + (optional) WebSMS Device ID/SG files
    ↓
@stock-opname-teknisi-comparator
    ├─ if given a raw Stock Quant export: run scripts/clean_stock_quant.py
    │  --location-suffix "" (full ledger, not just */Stock — this pipeline
    │  needs to know where an item currently sits anywhere, e.g. /Customer
    │  or /Teknisi)
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
                Inventory Masterfile xlsx
    ↓
@stock-quant-cleaner
    ├─ validate input columns
    ├─ run scripts/clean_stock_quant.py
    └─ report abnormality breakdown + deliver annotated CSV
    ↓ (cleaned CSV path passed straight through, no re-asking the user)
@stock-opname-comparator
    ├─ run scripts/compare_stock_opname.py against the cleaned CSV + city +
    │  opname xlsx + masterfile xlsx
    └─ report Summary + deliver the 3-sheet report (Summary, Product Summary,
       IMEI Mismatch)
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
  cleaned ERP output against a physical Stock Opname count for one city. Only
  products listed in the Inventory Masterfile (`Category` sheet) are in
  scope. ERP rows flagged `ERP Duplicate line` are excluded from the ERP
  "on-hand" side before matching — confirmed against real data that these are
  phantom double-bookings a physical count won't find. Full logic lives in
  `scripts/compare_stock_opname.py` (requires `openpyxl`, see
  `scripts/requirements.txt`).

### Stock Opname Teknisi domain

- `/stock-opname-teknisi` reconciles ERP stock against a **technician-level**
  physical count (as opposed to `/stock-opname`'s per-city warehouse count).
  Ported from a business-owner-provided reference script ("stock opname
  teknisi v3"); full logic lives in `scripts/compare_stock_opname_teknisi.py`.
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
