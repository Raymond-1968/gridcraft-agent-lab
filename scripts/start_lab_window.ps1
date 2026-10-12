# Starts an isolated GridCraft lab window with a local control channel.
#
# Isolation:
#   - GRIDCRAFT_NO_PREFS=1 is set for this process only: no saved UI settings are
#     read or written, and the app logs to stderr only (redirected into tmp/).
#   - The control channel binds to 127.0.0.1 on a lab-specific port (default 7991),
#     not the upstream example port 7979, to avoid clashing with other projects.
#   - Nothing global is modified (no user env vars, no global MCP config).
#
# Usage:
#   powershell -File scripts/start_lab_window.ps1 [-Port 7991] [-Workbook path.xlsx]
param(
    [int]$Port = 7991,
    [string]$Workbook = "",
    [string]$Exe = "C:\Program Files\GridCraft\gridcraft.exe"
)

$ErrorActionPreference = "Stop"
$lab = Split-Path -Parent $PSScriptRoot
$tmp = Join-Path $lab "tmp"
New-Item -ItemType Directory -Force $tmp | Out-Null

$env:GRIDCRAFT_NO_PREFS = "1"   # process-scoped; inherited by the child only
$argsList = @("--control", "$Port")
if ($Workbook) { $argsList += $Workbook }

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$p = Start-Process -FilePath $Exe -ArgumentList $argsList -WorkingDirectory $tmp -PassThru `
    -RedirectStandardError (Join-Path $tmp "gridcraft_window_$stamp.stderr.log") `
    -RedirectStandardOutput (Join-Path $tmp "gridcraft_window_$stamp.stdout.log")
$p.Id | Set-Content -Encoding ascii (Join-Path $tmp "lab_window.pid")
Write-Output "GridCraft lab window PID=$($p.Id) control=127.0.0.1:$Port"
Write-Output "MCP bridge: `"C:\Program Files\GridCraft\gridcraft-cli.exe`" mcp --connect $Port"
