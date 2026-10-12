You are running Stage 02 correction 01 of the GridCraft agent lab (C:\repos\gridcraft-agent-lab):
selective command discovery. A GridCraft lab window is already running (control port 7991).
Use ONLY the MCP tools of the server "gridcraft-lab" (mcp__gridcraft-lab__*). Synthetic data only.

Hard rules:
- NEVER call list_commands without a search argument. Never request the full command catalog.
- Use exactly these narrow searches: "cell.set", "home.bold", "file.save".
- Run the commands through execute_command with the command ids and parameter names exactly as
  returned by list_commands. Do not guess ids or parameters.
- If the gridcraft-lab tools are not available, stop and answer:
  CORRECTION_VERDICT: BLOCKED — <exact reason>.

Steps:
1. Call list_commands with search "cell.set", then "home.bold", then "file.save".
2. For each of the three commands cell.set, home.bold and file.save, quote verbatim from the
   tool result in your context: id, label, params doc. If any of them is missing from the
   result or the result was not shown in full in your context, mark step 2 FAIL.
   For "file.save" also list any other ids that the search returned.
3. Call new_workbook (no sample). Then, with execute_command only:
   - cell.set: put the synthetic value 123.45 into B2;
   - home.bold: make B2 bold;
   - file.save: save a copy to
     C:\repos\gridcraft-agent-lab\workbooks\stage02_correction01_synthetic_copy.xlsx
     (if the discovered params show file.save cannot take a path, use the save-as command
     returned by the same "file.save" search, and say so).
4. Verify through MCP: get_cell B2 (value 123.45, style.font.bold true) and inspect_workbook
   (path = the saved copy, dirty false). Then screenshot to
   C:\repos\gridcraft-agent-lab\runs\STAGE_02_CORRECTION_01\screenshots\result.png

Do not quit the GridCraft window.

Final answer format (plain text):
- one block per step: STEP <n>: PASS|FAIL — <tools called> — <key values quoted from raw tool results>
- last line exactly: CORRECTION_VERDICT: PASS|FAIL|BLOCKED — <short reason>
