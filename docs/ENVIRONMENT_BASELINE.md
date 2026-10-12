# Environment baseline

Recorded 2026-10-12. Nothing was installed, updated or reinstalled.

## Installed GridCraft

| Item | Value |
|---|---|
| Install directory | `C:\Program Files\GridCraft` |
| Uninstall entry | `GridCraft` 0.6.1, publisher Learning Machines LLC |
| Desktop app | `C:\Program Files\GridCraft\gridcraft.exe` — ProductVersion/FileVersion 0.6.1 |
| Desktop SHA-256 | `D9886BB105CD667CE5FA7B87B25F32BDEBD753CC0128898ACBAC7342EC009118` |
| CLI | `C:\Program Files\GridCraft\gridcraft-cli.exe` — `gridcraft-cli 0.6.1` (`--version`) |
| CLI SHA-256 | `D3010E92811B51BC6A316168103B6F9DF315F0CB74D2136D03A40B4106974B37` |
| Binaries timestamp | 2026-10-12 16:35:06 |
| On `PATH` | No — scripts and `.mcp.json` use the full path |
| MCP serverInfo | `{"name":"gridcraft","title":"GridCraft","version":"0.6.1"}` |
| Graphics | DirectX 12, NVIDIA GeForce RTX 4060 Ti |
| OS | Windows 11 Pro 10.0.26300; Python 3.12.0 (stdlib only) |

## Upstream documentation checked

Repository `https://github.com/storytold/gridcraft`, `main` at
`fc45b7f6d8ab79ef20f74e3c8347a2fd731e6af5` (2026-10-12T15:11:15Z). Latest release `v0.6.1`
(2026-10-12T17:07:30Z) — matches the installed build.

Read: `README.md`, `docs/control-protocol.md`, `docs/mcp.md`, `docs/cli.md`, `AGENTS.md`.

Relevant facts used by this lab:

- `gridcraft --control PORT` (or `GRIDCRAFT_CONTROL_PORT`) opens a newline-delimited JSON
  control channel on `127.0.0.1:PORT`; methods include `engine.execute`, `engine.commands`,
  `engine.journal`, `document.inspect`, `sheet.read`, `cell.get`, `ui.*`, `ui.screenshot`,
  `app.open` / `app.save` / `app.quit`.
- `gridcraft-cli mcp --connect PORT` bridges MCP (JSON-RPC 2.0 on stdio, protocol 2025-06-18)
  to a running window; `gridcraft-cli send PORT METHOD [json]` sends one control request.
- `GRIDCRAFT_NO_PREFS` (set) runs without saved settings and logs to stderr only.
- The GUI binary has no `--help`: it starts a window (observed, see Stage 01 report).

## Isolation choices

- Lab window: `GRIDCRAFT_NO_PREFS=1` set on the child process only; working directory `tmp/`.
- Control port `7991` (not the upstream example `7979`), loopback only.
- MCP registered in the project-scoped `.mcp.json` only; no `claude mcp add` (global/user scope).

## Command catalog (installed 0.6.1)

| Source | Count | File |
|---|---|---|
| `gridcraft-cli commands --json` (engine) | 339 | `runs/STAGE_01/catalog/commands.json`, `commands.txt` |
| `engine.commands` from the live window | 352 | `runs/STAGE_01/catalog/commands_live_window.json` |
| MCP `tools/list` | 21 tools | `runs/STAGE_01/catalog/mcp_tools.json` |
| `gridcraft-cli functions --json` | 513 functions | `runs/STAGE_01/catalog/functions.json` |

Window-only commands (13): `app.language.english`, `app.language.japanese`, `app.language.set`,
`ui.backstage`, `ui.dialog`, `ui.printSettings`, `ui.printSubmit`, `ui.ribbonTab`,
`view.animations`, `view.collapseRibbon`, `view.darkMode`, `view.interfaceScale`, `view.theme`.
