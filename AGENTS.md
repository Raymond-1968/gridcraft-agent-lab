# AGENTS.md

## Scope

Isolated laboratory for driving the installed GridCraft desktop spreadsheet through its
documented machine interfaces: `gridcraft-cli`, the MCP server (`gridcraft-cli mcp`) and the
desktop control channel (`gridcraft --control PORT`).

Work only inside this repository. Do not modify other repositories, the GridCraft installation,
global GridCraft settings, user environment variables or global agent/MCP configuration.

## Operating rules

1. Discover before executing: confirm command ids and parameters from the verified local
   catalog or a narrow MCP `list_commands` search (see "Command discovery"). Never invent command ids.
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

## Command discovery (build-bound catalog, selective disclosure)

- The local catalog `catalog/gridcraft-<version>/` is the reference for command ids and
  parameters. It is valid only for the build recorded in its `manifest.json`
  (version and SHA-256 of `gridcraft.exe` and `gridcraft-cli.exe`, plus hashes of the catalog files).
- Before relying on it, run `python scripts/catalog.py verify`. On any mismatch (new build,
  changed binary, edited catalog file) the catalog is invalid: stop, capture a new catalog in a
  new run folder and a new `catalog/gridcraft-<version>/` with its own manifest. Never edit
  catalog files by hand.
- Disclose commands selectively. In an MCP session, never call `list_commands` without a
  narrow `search` and never read the full `gridcraft://commands` resource: the full catalog
  (~96 KB) does not fit into a Claude Code session's tool output (Stage 02). Search for the exact
  ids you need (e.g. `cell.set`, `home.bold`, `file.save`) and use the returned params verbatim.
  Locally, use `python scripts/catalog.py search <term>`.
- Raising output limits (`MAX_MCP_OUTPUT_TOKENS` etc.) is not a substitute for selective discovery.

## Data policy (public repository)

- Never use real financial data. Tests use synthetic values only.
- Workbooks (`*.xlsx`, `*.csv`, …), `workbooks/`, `inbox/` (incoming data / поступления),
  `tmp/` and local settings are git-ignored. Check `git status` before every commit.
- Logs and screenshots committed as evidence must contain synthetic data only.

## Git

Commits are allowed for lab work. Do not push without explicit human authorization.
