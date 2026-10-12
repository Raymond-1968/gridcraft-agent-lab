You are running Stage 02 of the GridCraft agent lab (C:\repos\gridcraft-agent-lab).
A GridCraft lab window is already running with its control channel on port 7991.
Use ONLY the MCP tools of the server "gridcraft-lab" (mcp__gridcraft-lab__*). Do not try any other tool.
Synthetic test values only. Do not create any income form.

If the gridcraft-lab MCP tools are not available in this session, stop immediately and answer:
STAGE02_VERDICT: BLOCKED — <exact reason, including which tools you do see>.

Steps:
1. Call list_commands (no arguments) to get the command catalog; report the number of commands.
   Then call inspect_workbook and report the workbook state (title, path, dirty, sheets, used range).
2. Call new_workbook (no sample) to create an empty test workbook; confirm with inspect_workbook that it is empty.
3. set_cell A1 = 17, set_cell A2 = 25, set_cell A3 = "=SUM(A1:A2)".
4. get_cell A3 and read_range A1:A3. The computed value of A3 must be 42.
5. save_workbook with path C:\repos\gridcraft-agent-lab\workbooks\stage02_claude_session_test.xlsx
   then screenshot with path C:\repos\gridcraft-agent-lab\runs\STAGE_02\screenshots\stage02_result.png
   then inspect_workbook once more (path and dirty flag).

Do not quit the GridCraft window.

Final answer format (plain text):
- one line per step: STEP <n>: PASS|FAIL — <tool names called> — <key values from the raw tool results>
- last line exactly: STAGE02_VERDICT: PASS|FAIL|BLOCKED — <short reason>
