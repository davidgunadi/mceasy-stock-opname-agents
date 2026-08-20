# McEasy Stock Opname Claude Agents

**Version:** 2.0.0 — see [CHANGELOG.md](CHANGELOG.md)

Claude Code agents and skills that automate McEasy's stock reconciliation workflow: cleaning raw ERP "Stock Quant" exports for data-quality issues, then reconciling them against physical Stock Opname counts — either per-city (`/stock-opname`) or per-technician (`/stock-opname-teknisi`). Built for ops/inventory teams who need to catch ERP-vs-physical stock mismatches without manual spreadsheet reconciliation.

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
