Stage 03, session 1 — build the income tracker prototype in GridCraft.
Lab: C:\repos\gridcraft-agent-lab. A fresh GridCraft lab window is running (control port 7991).
All workbook changes must be made by GridCraft through the MCP tools of server "gridcraft-lab".
You may use Read only to read the plan file. Fictional data only.
Never call list_commands without a narrow search.

1. Call inspect_workbook and confirm the active workbook is an empty, unsaved book.
2. Read C:\repos\gridcraft-agent-lab\runs\STAGE_03\build_plan.jsonl (one JSON object per line:
   {"step", "command", "params"}).
3. Execute EVERY step in order with execute_command, passing "command" and "params" exactly as
   written (no edits, no merging, no skipping). If a step returns an error, stop and report the
   step number, command and exact error text.
4. Save with execute_command file.save, params {"path": "C:\\repos\\gridcraft-agent-lab\\workbooks\\income_tracker_prototype.xlsx"}.
5. Call inspect_workbook (report sheet names, used ranges, freeze, validations, dirty, path).
6. Read back with read_range: sheet "Сводка" range A4:B10 and A12:H15; sheet "Поступления" A5:F10.
7. For each sheet in order "Сводка", "Источники", "Поступления": execute_command sheet.activate
   {"sheet": name}, then screenshot to
   C:\repos\gridcraft-agent-lab\runs\STAGE_03\screenshots\build_<n>_<sheet-name-in-latin: summary|sources|payments>.png
   Finally activate "Сводка" again.

Do not quit the GridCraft window.

Final answer (plain text):
BUILD: <executed steps>/<total steps>, errors: <none | step n: error>
SAVE: <file.save result>
STATE: <key facts from inspect_workbook>
VALUES: <Сводка B8, B9, B10 and rows A, B as read>
SCREENSHOTS: <paths>
SESSION1_VERDICT: PASS|FAIL — <reason>
