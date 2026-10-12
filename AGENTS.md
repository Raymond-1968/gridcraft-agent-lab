# AGENTS.md

## Scope

Isolated laboratory for driving the installed GridCraft desktop spreadsheet through its
documented machine interfaces: `gridcraft-cli`, the MCP server (`gridcraft-cli mcp`) and the
desktop control channel (`gridcraft --control PORT`).

Work only inside this repository. Do not modify other repositories, the GridCraft installation,
global GridCraft settings, user environment variables or global agent/MCP configuration.

## Operating rules

1. Discover before executing: confirm command ids and parameters with `gridcraft-cli commands`,
   `engine.commands` or MCP `list_commands` for the installed build. Never invent command ids.
2. Workbook creation, edits and calculation happen inside GridCraft via its own commands.
   Scripts only send requests, read replies and compare them with expectations. Do not write
   or compute workbook contents with third-party libraries.
3. Use the isolated lab window: `GRIDCRAFT_NO_PREFS=1`, control port `7991` on `127.0.0.1`
   (see `scripts/start_lab_window.ps1`). Do not persist UI preferences
   (`view.theme`, `view.darkMode`, `app.language.*`, `ui.set`) — they target `ui.json`.
4. Record every request and reply (JSONL transcript), plus `engine.journal`, screenshots and a
   `PASS` / `FAIL` / `BLOCKED` verdict for each stage. On failure, name the exact stage and error.
   Preserve failed runs as diagnostic records.
5. Do not update, reinstall or patch GridCraft.
6. Do not expand scope autonomously; each stage is a separate assignment.

## Data policy (public repository)

- Never use real financial data. Tests use synthetic values only.
- Workbooks (`*.xlsx`, `*.csv`, …), `workbooks/`, `inbox/` (incoming data / поступления),
  `tmp/` and local settings are git-ignored. Check `git status` before every commit.
- Logs and screenshots committed as evidence must contain synthetic data only.

## Git

Commits are allowed for lab work. Do not push without explicit human authorization.
