Stage 05, session 2 — functional checks of the shared-obligation balance (simplified copy) in GridCraft.
Lab: C:\repos\gridcraft-agent-lab. The GridCraft lab window (port 7991) has the prototype open:
C:\repos\gridcraft-agent-lab\workbooks\shared_obligation_balance.xlsx (built in session 1).
Use the MCP tools of server "gridcraft-lab" for every workbook action; Read only for plan files.
Never call list_commands without a narrow search. Fictional data only.

0. Call inspect_workbook; confirm the active workbook path is the prototype above. If not, stop: FAIL.
1. Read C:\repos\gridcraft-agent-lab\runs\STAGE_05\checks_plan.jsonl. Execute every line in order:
   call MCP tool "tool" with "arguments" exactly as written.
   - "expect": compare the read_range values with expect cell by cell. Numbers match if
     |actual - expected| < 0.005. An expected "" matches an empty string or null.
   - "expect_no_errors": the range must not contain any value that is an error (e.g. a string
     starting with "#" such as #VALUE!, #REF!, #N/A, #DIV/0!, #NAME?).
   - "expect_ok": the data.validate result field "ok" must equal expect_ok.
   Do not stop at a mismatch; record it and continue (later steps restore the base state).
   Stop only if a tool call itself returns an error, and report it exactly.
2. Snapshot: Read C:\repos\gridcraft-agent-lab\runs\STAGE_05\snapshot_cells.json. For each entry
   in order: execute_command sheet.activate {"sheet": sheet}, then get_cell {"cell": cell}.
   Compare style with the entry's "style": fill = style.fill.fg.Rgb, num_fmt = style.num_fmt,
   bold = style.font.bold. Report formula and value of every snapshot cell.
3. execute_command sheet.activate {"sheet": "Учёт"}; then inspect_workbook (report dirty and path).
   If dirty is true, execute_command file.save with path
   C:\repos\gridcraft-agent-lab\workbooks\shared_obligation_balance.xlsx and report it.

Do not quit the GridCraft window.

Final answer (plain text):
- one line per check id (L0, L1, ... L8): CHECK <id>: PASS|FAIL — <title(s)> — <mismatches with actual vs expected, or "all matched">
- SNAPSHOT: <n>/<total> style matches; list each mismatch
- STATE: <dirty, path>
- last line exactly: SESSION2_VERDICT: PASS|FAIL — <reason>
