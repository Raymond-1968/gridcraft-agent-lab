# GridCraft Agent Lab

Isolated Windows laboratory for driving the GridCraft desktop spreadsheet from an AI agent
through GridCraft's own interfaces: `gridcraft-cli`, MCP (`gridcraft-cli mcp --connect`) and the
desktop control channel.

Stage 01 (connection and lifecycle smoke test) verdict: **PASS** — see
[runs/STAGE_01/REPORT.md](runs/STAGE_01/REPORT.md).

Stage 02 (Claude Code session with the project MCP server) verdict: **FAIL at step 1** — the session
drove GridCraft through MCP (17 + 25 = 42, save, screenshot all passed), but the full
`list_commands` reply (~96 KB) does not fit into the session's tool output. See
[runs/STAGE_02/REPORT.md](runs/STAGE_02/REPORT.md).

Stage 02 correction 01 (selective command discovery) verdict: **PASS** — narrow
`list_commands` searches fit in the session; `cell.set`, `home.bold`, `file.save` executed and
verified. See [runs/STAGE_02_CORRECTION_01/REPORT.md](runs/STAGE_02_CORRECTION_01/REPORT.md).

## Layout

| Path | Contents |
|---|---|
| [AGENTS.md](AGENTS.md) | Rules for agents working in this lab |
| [docs/ENVIRONMENT_BASELINE.md](docs/ENVIRONMENT_BASELINE.md) | Installed paths, versions, hashes, upstream docs |
| [scripts/start_lab_window.ps1](scripts/start_lab_window.ps1) | Starts an isolated lab window with a control channel |
| [scripts/stage01_smoke.py](scripts/stage01_smoke.py) | Stage 01 test: window + MCP bridge + checks (stdlib Python) |
| [catalog/gridcraft-0.6.1/](catalog/gridcraft-0.6.1/manifest.json) | Build-bound command catalog with version/hash manifest |
| [scripts/catalog.py](scripts/catalog.py) | `verify` the catalog against the installed build; `search <term>` selectively |
| [.mcp.json](.mcp.json) | Project-scoped MCP server `gridcraft-lab` → `gridcraft-cli mcp --connect 7991` |
| `runs/STAGE_01/` | Command catalog, transcripts, engine journal, screenshots, results, report |
| `workbooks/`, `inbox/`, `tmp/` | Git-ignored: workbooks, incoming data, scratch |

## Quick start

```powershell
# 1. isolated window (no saved prefs, control channel on 127.0.0.1:7991)
powershell -File scripts/start_lab_window.ps1
# 2. one control-channel request
& "C:\Program Files\GridCraft\gridcraft-cli.exe" send 7991 document.inspect
# 3. agent: open Claude Code in this folder and approve the project MCP server "gridcraft-lab"
```

Re-run the Stage 01 test (it launches and closes its own windows on port 7991):

```powershell
python scripts/stage01_smoke.py
```

## Data policy

This repository is public. Only synthetic test values are used. Workbooks, incoming data and
local settings are excluded by [.gitignore](.gitignore).
