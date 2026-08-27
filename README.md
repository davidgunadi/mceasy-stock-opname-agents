# McEasy Stock Opname Claude Agents

**Version:** 2.0.1 — see [CHANGELOG.md](CHANGELOG.md)

Claude Code agents and skills that automate McEasy's stock reconciliation workflow: cleaning raw ERP "Stock Quant" exports for data-quality issues, then reconciling them against physical Stock Opname counts — either per-city (`/stock-opname`) or per-technician (`/stock-opname-teknisi`). Built for ops/inventory teams who need to catch ERP-vs-physical stock mismatches without manual spreadsheet reconciliation.

---

## Background

### The problem

McEasy's ERP tracks stock through a **"Stock Quant"** export — a historical
ledger, not a snapshot. The same item (by lot/serial number) legitimately
appears on many rows as it moves through its lifecycle: Vendor → Customer →
Refurbishment → Teknisi → Warehouse, usually leaving zero-quantity "ghost"
rows behind at each stop. Before you can trust what the ERP says is on-hand,
you first have to catch data-quality issues in that ledger: duplicate
bookings, wrong owners, malformed IMEIs, missing serials.

Once the ERP data is clean, the real question is: **does the ERP match
physical reality?** A "Stock Opname" is a physical count — someone walks the
warehouse (or a technician's own stock) and counts what's actually there.
This repo automates comparing that physical count against the cleaned ERP
data and reports where they disagree.

### Two reconciliation modes

| | `/stock-opname` | `/stock-opname-teknisi` |
|---|---|---|
| Physical count is done by | Warehouse staff, per city | Field technicians, individually |
| Scope | On-hand stock only (`<City>/Stock`) | Full ledger — includes stock currently checked out to a customer or technician |
| Special handling | Excludes ERP "duplicate line" phantom bookings | Uses stock movement history + WebSMS device records to explain apparent mismatches (e.g. a technician installed a unit at a customer's site the same day) |
| Output | Summary / Product Summary / IMEI Mismatch | Detail IMEI / Detail Non-IMEI / Summary by Technician |

### Key terms

- **Stock Quant** — the raw ERP export of every stock lot/serial and where it
  currently sits.
- **IMEI / Lot-Serial Number** — a device's unique serial; used to track
  individually serialized items across locations.
- **Non-IMEI** — items counted by quantity only (not individually
  serialized).
- **Abnormality** — a data-quality flag added by `clean_stock_quant.py`
  (duplicate line, special character, IMEI length mismatch, alphanumeric
  IMEI, negative qty, owner mismatch, missing IMEI) — see the script's
  docstring for exact rules.
- **Teknisi** — Indonesian for "technician"; stock checked out to a field
  technician is tracked separately from warehouse stock.
- **WebSMS override** — a separate source of truth (device management
  system) used to confirm a unit is genuinely installed at a customer,
  overriding an otherwise-flagged mismatch.

---

## Prerequisites

Must Have:

- [Claude Desktop](https://claude.com/download)
- [Git](https://git-scm.com/install/)
- [Python](https://www.python.org/downloads/)
- [Node](https://nodejs.org/en/download)
- Claude Pro or Max account

- Optional:

- [GitHub Desktop](https://desktop.github.com/download/)
- [VSCode](https://code.visualstudio.com/download)

---

## Setup

### Option A: Using the terminal (CLI)

1. Clone the repo

```bash
git clone https://github.com/[your-org]/[your-repo].git
```

1. Open the folder in Claude Code

2. That's it — agents and skills load automatically from `.claude/`

### Option B: Using GitHub Desktop

1. Open [GitHub Desktop](https://desktop.github.com/)
2. Click **File → Clone Repository**
3. Go to the **URL** tab and paste the repo URL:
   - On the GitHub repo page, click the green **\< \> Code** button
   - Make sure **HTTPS** is selected
   - Click the copy icon next to the URL
   - Paste it into GitHub Desktop
4. Choose a local path and click **Clone**
5. Once cloned, open the folder in Claude Code:
   - In GitHub Desktop, click **Repository → Open in…** and select your Claude Code editor, or
   - Open Claude Code manually and use **File → Open Folder** to navigate to the cloned folder

> **Note:** The `.claude/` folder is hidden by default. On Mac press `Cmd + Shift + .` to show hidden files in Finder.

---

## How to use

Type the slash command in Claude Code:

```
/stock-opname
```

- **Input:** a raw Stock Quant CSV export, the city, a physical Stock Opname
  xlsx, and the Inventory Masterfile xlsx.
- **Output:** the cleaned/annotated Stock Quant CSV, plus a 3-sheet Excel
  report (Summary, Product Summary, IMEI Mismatch) — both saved next to your
  input files.

or, for a technician-level physical count instead of a per-city one:

```
/stock-opname-teknisi
```

- **Input:** East + West technician Stock Opname workbooks, the Inventory
  Masterfile, a Stock Quant export (raw or already cleaned), the Stock
  Movement export, and optionally the WebSMS Device ID/SG files.
- **Output:** a 4-sheet Excel report (Detail IMEI, Detail Non-IMEI, Summary
  IMEI Overall, Summary by Tech), saved next to your movement/opname files.

Claude will greet you, explain the pipeline, and ask for the inputs it needs.

---

## Folder structure

```
your-repo/
├── CLAUDE.md                          # Project instructions — agents read this for context
├── README.md                          # This file
├── .gitignore
├── outputs/                           # All generated files go here
│   └── .gitkeep
├── scripts/                           # Scripts invoked by agents while running
│   └── .gitkeep
└── .claude/
    ├── settings.json                  # Permissions (e.g. allow WebSearch)
    ├── agents/                        # One .md file per agent
    │   ├── stock-quant-cleaner.md
    │   ├── stock-opname-comparator.md
    │   └── stock-opname-teknisi-comparator.md
    └── skills/                        # One folder per slash command
        ├── stock-opname/
        │   └── SKILL.md
        └── stock-opname-teknisi/
            └── SKILL.md
```

---

## How agents work

Each file in `.claude/agents/` defines one agent. The frontmatter controls how Claude uses it:

```markdown
---
name: agent-name # How you invoke it: @agent-name
description: ... # Claude reads this to decide when to use it
model: sonnet # sonnet (fast/cheap) or opus (deep research)
tools: Read, Write # What tools the agent can use
---

[System prompt — the agent's instructions]
```

**Model guidance:**

- Use `opus` for research agents that need deep reasoning and web search
- Use `sonnet` for writing, reviewing, and formatting agents

**Tools you can grant:**

- `Read`, `Write`, `Edit` — file access
- `WebSearch`, `WebFetch` — internet access
- `Bash` — shell commands (use carefully)

---

## How skills work

Each folder in `.claude/skills/` is a slash command. The folder name becomes the command.

```
.claude/skills/my-pipeline/SKILL.md  →  /my-pipeline
```

The `SKILL.md` file is a prompt that tells Claude how to orchestrate the pipeline — which agents to invoke, in what order, and where to pause for human input.

---

## How CLAUDE.md works

`CLAUDE.md` in the repo root is always loaded as project context. Use it to:

- Define the pipeline order and rules
- Give agents shared background knowledge about your domain, product, or team
- Set hard rules (e.g. "always save to `./outputs/`", "scripts go in `./scripts/`", "write in English only")

Agents read `CLAUDE.md` automatically — you don't need to repeat context in every agent file.

---

## Questions

[Add your team contact here]
