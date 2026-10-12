Stage 05, session 1b — apply the post-review fix to the shared-obligation balance.
Lab: C:\repos\gridcraft-agent-lab. The GridCraft lab window (port 7991) has
C:\repos\gridcraft-agent-lab\workbooks\shared_obligation_balance.xlsx open.
Use the MCP tools of server "gridcraft-lab" for every workbook action; Read only for the plan file.
Never call list_commands without a narrow search.

1. inspect_workbook: confirm the path above and dirty=false. If not, stop: FAIL.
2. (the fix plan activates the sheet itself).
3. Read C:\repos\gridcraft-agent-lab\runs\STAGE_05\fix_plan.jsonl and execute every step in order
   with execute_command, "command" and "params" exactly as written. Stop on the first error and report it.
4. inspect_workbook: report conditional formats per sheet.
5. execute_command file.save {"path": "C:\repos\gridcraft-agent-lab\workbooks\shared_obligation_balance.xlsx"}.
6. screenshot to C:\repos\gridcraft-agent-lab\runs\STAGE_05\screenshots\fix_uchet.png
Do not quit the window.

Final answer: FIX: <executed>/<total>, errors; SAVE: <result>; STATE: <conditional formats, dirty>;
last line exactly: SESSION1B_VERDICT: PASS|FAIL — <reason>
