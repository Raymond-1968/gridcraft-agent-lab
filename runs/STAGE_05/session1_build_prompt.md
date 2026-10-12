Stage 05, session 1 — make the simplified copy (one cumulative balance) in GridCraft.
Lab: C:\repos\gridcraft-agent-lab. A fresh GridCraft lab window is running (control port 7991).
All workbook changes must be made by GridCraft through the MCP tools of server "gridcraft-lab".
You may use Read only to read the plan file. Fictional data only.
Never call list_commands without a narrow search.

1. Call inspect_workbook and report the active workbook.
2. Read C:\repos\gridcraft-agent-lab\runs\STAGE_05\build_plan.jsonl (one JSON object per line:
   {"step", "command", "params"}). Step 1 opens the Stage 04 workbook, step 2 saves it under the new
   name C:\repos\gridcraft-agent-lab\workbooks\shared_obligation_balance.xlsx — every later change
   goes to the copy. Never save to shared_obligation_tracker.xlsx.
3. Execute EVERY step in order with execute_command, passing "command" and "params" exactly as written
   (no edits, no merging, no skipping). If a step returns an error, stop and report the step number,
   command and exact error text.
4. After step 2, call inspect_workbook once and confirm the active path is the new file; if it is not,
   stop: FAIL.
5. After the last step: execute_command file.save {"path": "C:\repos\gridcraft-agent-lab\workbooks\shared_obligation_balance.xlsx"}.
6. inspect_workbook (sheets, used ranges, freeze, validations, conditional formats, dirty, path).
7. read_range sheet "Учёт" A4:E10 and A13:E20.
8. Screenshots: execute_command sheet.activate {"sheet": "Учёт"}, then screenshot to
   C:\repos\gridcraft-agent-lab\runs\STAGE_05\screenshots\build_1_uchet.png; execute_command
   sheet.activate {"sheet": "Настройки"}, then screenshot to ...\build_2_settings.png. Finally activate "Учёт"
   and execute_command selection.set {"range": "A20"}.
Do not quit the GridCraft window.

Final answer (plain text):
BUILD: <executed steps>/<total steps>, errors: <none | step n: error>
COPY: <path after step 2>
SAVE: <file.save result>
STATE: <key facts from inspect_workbook>
VALUES: <Учёт B4:B10, E4:E6 and the entry rows as read>
SCREENSHOTS: <paths>
SESSION1_VERDICT: PASS|FAIL — <reason>
