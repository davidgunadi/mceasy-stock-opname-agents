# Changelog

All notable changes to this repo are documented in this file.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and versioning follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [2.0.1] - 2026-08-26

### Changed

- `scripts/compare_stock_opname.py`: the ERP-exclusion rule used by
  `@stock-opname-comparator` now excludes any ERP row carrying **any**
  non-empty `Abnormality` flag (Owner Mismatch, Negative Qty, Imei Length
  Difference, Alphanumeric IMEI, Has Special Character, Missing Imei, ERP
  Duplicate line) from both the IMEI-lot match and the non-IMEI qty sum —
  previously only `ERP Duplicate line` rows were excluded, and only from the
  IMEI-lot match. Confirmed against real SBY 260826 data: a Negative
  Qty/Owner Mismatch line (owned by a courier company, not the warehouse) was
  inflating the non-IMEI qty comparison for one product into a false ~41.3M
  IDR discrepancy that vanished once flagged rows were excluded.

### Fixed

- `scripts/compare_stock_opname.py`: product-name matching against the
  masterfile whitelist is now case-insensitive (`build_product_lookup`).
  Previously an exact, case-sensitive match silently dropped any ERP or
  physical row whose product name differed only in casing from the
  masterfile out of scope entirely, reporting it as ERP/physical qty 0
  instead of comparing it. Confirmed against real SBY 260826 data: the ERP
  export's `FUSE 2A` vs the masterfile's `Fuse 2A` produced a false -26,000
  IDR discrepancy (reported as ERP qty 0 vs physical 13) even though the row
  had no Abnormality flag and its actual quantity matched the physical count
  exactly.

## [2.0.0] - 2026-08-20

### Removed

- `/clean-stock-quant` skill — was redundant once `/clean-and-compare-stock-opname`
  existed; the `@stock-quant-cleaner` agent it invoked is unchanged and now runs
  only via `/stock-opname`.
- `/compare-stock-opname` skill — same reason; `@stock-opname-comparator` is
  unchanged and now runs only via `/stock-opname`. Users who only want to
  compare an already-cleaned CSV must now go through the combined pipeline
  (there is no standalone comparison-only entry point anymore).
- `/example-skill` — inert template placeholder, not a real pipeline; removed
  as dead scaffold now that the repo has real skills.
- `example-agent` — inert template agent, not a real agent; removed as dead
  scaffold for the same reason. The repo no longer ships a starter
  agent/skill template — use one of the real agents/skills as a reference
  instead when authoring a new one.

### Changed

- `/clean-and-compare-stock-opname` renamed to `/stock-opname` — now the sole
  entry point chaining `@stock-quant-cleaner` → `@stock-opname-comparator`.
  No change to either agent's own logic; this is a rename plus removal of the
  two skills it superseded.
- `CLAUDE.md`: updated the skills table, Pipeline Order, Rules exception
  list, and domain context sections to reflect `/stock-opname` as the only
  per-city reconciliation entry point, and dropped the `/example-skill` row.
- `README.md`: "How to use" and the folder-structure example now point at
  `/stock-opname` and `stock-quant-cleaner.md` instead of the removed
  `/example-skill` and `example-agent.md`.
- `CLAUDE.md`: dropped the `example-agent` row from the agents table and its
  "see example-agent.md for the template" pointer.

## [1.5.0] - 2026-08-19

### Added

- `/clean-and-compare-stock-opname` skill: chains `/clean-stock-quant` →
  `/compare-stock-opname` in one pass, so a raw Stock Quant CSV export goes
  straight to a Stock Opname reconciliation report without the user manually
  re-invoking a second command in between. No changes to either underlying
  agent or script — this skill only removes the manual hand-off (passing the
  cleaned CSV path straight from `@stock-quant-cleaner` to
  `@stock-opname-comparator`).
- `CLAUDE.md`: documented the `/clean-and-compare-stock-opname` pipeline and
  added it to the `./outputs/` rule's exception list (it also saves next to
  the user's input files, like the two skills it chains).

## [1.4.1] - 2026-08-13

### Changed

- `scripts/compare_stock_opname_teknisi.py`: ported three fixes from the
  business owner's updated reference script ("stock opname teknisi v6")
  onto the already-ported version:
  - Stock Movement `.xlsx` reads now use the `calamine` engine (matching v6)
    instead of the default `openpyxl` engine — on the ~700k-row export this
    file routinely is, that cut total pipeline runtime from ~5 minutes to
    ~47 seconds. `.csv` movement exports are unaffected.
  - `build_indexes_imei`: when an ERP serial has duplicate rows and no "Last
    Movement Date" to break the tie, the row with `Quantity > 0` now wins
    over zero-quantity "ghost" rows left behind at old locations — previously
    an arbitrary duplicate could win, sometimes a ghost row.
  - `imei_status`: the "move on/after opname date" check now uses `>=`
    instead of `>` — a same-day move now counts as evidence (both dates are
    date-only, so same-day is the finest resolution available).
  - Re-validated end to end against real `260812` data: overall IMEI accuracy
    moved from 84.67% to 85.66%, and `Location mismatch` rows dropped from 21
    to 1 — most of what the old tie-break/strict-`>` logic misclassified as
    mismatches were actually same-day-move or ghost-row artifacts.
- `scripts/requirements.txt`: `python-calamine` is now also required for
  `.xlsx` Stock Movement inputs, not just the WebSMS device-list files.

## [1.4.0] - 2026-08-12

### Added

- `/stock-opname-teknisi` skill and `stock-opname-teknisi-comparator` agent
  (sonnet): reconciles the full ERP ledger against a technician-level physical
  Stock Opname count (fixed `03 East` / `04 West` workbooks), applies a WebSMS
  "installed at customer" override (`Device ID.xlsx` / `Device SG.xlsx`), and
  delivers a Detail IMEI / Detail Non-IMEI / Summary-by-Tech report. Ported
  from a business-owner-provided reference script ("stock opname teknisi v3").
- `scripts/compare_stock_opname_teknisi.py`: the deterministic engine behind
  it. IMEI rows use movement-then-location rules (latest DONE move after the
  opname date takes precedence, then ERP-location-vs-SO-location, qty-aware);
  Non-IMEI rows compare qty per Technician×Location×Product, skipping
  `/Customer`. Fixed a bug present in the reference script: a Non-IMEI row
  with a blank `SO Location` referenced undefined variables and crashed — it
  now reports `Inaccuracy — Missing SO Location`, mirroring the equivalent
  IMEI-side guard.
- `clean_stock_quant.py`: added `.xlsx` input support (auto-detected by
  extension, alongside the existing `.csv` path) — the ERP's raw export can be
  downloaded as either. Added a separate `--owner-check-suffix` flag (default
  `/Stock`, independent of `--location-suffix`) so `/stock-opname-teknisi` can
  request the full, unfiltered ledger (`--location-suffix ""`) without Owner
  Mismatch firing on every legitimately customer/technician-owned historical
  row — validated on real data: 224k false positives down to 96 real ones
  once this scoping was in place.
- `scripts/requirements.txt`: added `xlsxwriter`, `pandas`, and
  `python-calamine` (the last only needed for `/stock-opname-teknisi`'s
  WebSMS device-list files, whose style XML doesn't parse with plain
  `openpyxl`).
- `CLAUDE.md`: documented the `/stock-opname-teknisi` pipeline and its
  "Stock Opname Teknisi domain" context section.

### Fixed

- `scripts/compare_stock_opname.py`: added a UTF-8 stdout/stderr reconfigure
  guard — some of its report strings use an em dash, which crashed with
  `UnicodeEncodeError` on Windows consoles defaulting to a non-UTF-8 codepage
  (e.g. cp1252). No behavior change to the comparison logic itself.

## [1.3.0] - 2026-08-11

### Added

- `/compare-stock-opname` skill and `stock-opname-comparator` agent (sonnet):
  reconciles a `/clean-stock-quant` output against a physical Stock Opname
  count for one city, producing a Summary / Product Summary / IMEI Mismatch
  report.
- `scripts/compare_stock_opname.py`: the deterministic engine behind it.
  Matches by Lot/Serial Number alone (flagging cross-product matches as
  `Product Mismatch` instead of two separate misses), excludes ERP rows
  flagged `ERP Duplicate line` from the ERP "on-hand" side before matching
  (validated against real SBY 260730 data — these are phantom double-bookings
  a physical count won't find), and reports two independent accuracy measures
  (qty-based, and IMEI-tracked-product-count-based) for "All" and "A+B only"
  scope.
- `scripts/requirements.txt`: first external dependency for `scripts/`
  (`openpyxl`, needed to read/write `.xlsx` files for this skill).
- `CLAUDE.md`: documented the `/compare-stock-opname` pipeline and its
  ERP-Duplicate-line exclusion rule under "Stock Quant cleaning domain".

## [1.2.0] - 2026-08-10

### Added

- `/clean-stock-quant` skill and `stock-quant-cleaner` agent (sonnet): cleans an
  ERP Stock Quant CSV export and flags data-quality abnormalities (`ERP
  Duplicate line`, `Has Special Character`, `Imei Length Difference`,
  `Alphanumeric IMEI`, `Negative Qty`, `Owner Mismatch`, `Missing Imei`),
  producing an annotated CSV filtered to `*/Stock` locations.
- `scripts/clean_stock_quant.py`: the deterministic engine backing the skill
  above. Encodes the duplicate-detection rule (positive-quantity conflict OR
  exact Product+Lot+Location repeat — historical zero-qty ledger lines from
  the item moving through other locations are not flagged) and the per-product
  numeric-vs-letter lot convention used by the `Alphanumeric IMEI` check.
- `CLAUDE.md`: added a "Stock Quant cleaning domain" context section (valid
  owner, `*/Stock` location convention, historical-ledger caveat) and an
  explicit exception to the "outputs go to `./outputs/`" rule for this skill
  (it saves next to the user's input file instead, since the input lives
  outside the repo).

## [1.1.0] - 2026-07-07

### Added

- `./scripts/` folder convention: any script invoked by an agent while running must live there, not in the repo root. Referenced in `CLAUDE.md` rules and `README.md`'s folder structure.

## [1.0.0] - 2026-07-06

### Added

- Initial multi-agent template scaffold: `.claude/agents/example-agent.md` and `.claude/skills/example-skill/` as starting points for custom agents and slash commands.
- `.claude/settings.json` for tool permissions.
- Root `CLAUDE.md` defining the agent/skill/pipeline structure and repo rules (e.g. outputs go to `./outputs/`).
- `README.md` with setup instructions (CLI and GitHub Desktop) and an explanation of how agents, skills, and `CLAUDE.md` work together.

### Instruction to init a new repo

Set up changelog-based versioning for this repo:

Create CHANGELOG.md in the repo root using the Keep a Changelog format, versioned per SemVer (MAJOR.MINOR.PATCH — MAJOR for breaking/workflow-changing changes, MINOR for new capabilities, PATCH for fixes/tweaks). Seed it with a [1.0.0] (or check if package.json/similar already has a version — use that instead) entry summarizing the current state from recent git history, dated today.
Add a **Version:** x.y.z — see [CHANGELOG.md](CHANGELOG.md) line near the top of README.md, matching the changelog's latest version.
Add a "Versioning" section to this repo's CLAUDE.md (create one if it doesn't exist) that:
Defines what counts as a functional change for this specific repo (be concrete — e.g. "source code, build config, CLI behavior" vs. "docs, fixtures, sample data" — adapt to what this repo actually contains).
Instructs that any commit making a functional change gets a CHANGELOG.md entry (Added/Changed/Fixed/Removed) and a version bump in both CHANGELOG.md and the README header, done proactively as part of the commit — not just when asked.
Ask me what counts as a "functional change" here if it's not obvious from the repo's structure, rather than guessing.
