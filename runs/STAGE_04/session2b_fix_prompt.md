Stage 04, session 2b — apply the post-review fix to the shared-obligation tracker.
Lab: C:\repos\gridcraft-agent-lab. The GridCraft lab window (port 7991) has
C:\repos\gridcraft-agent-lab\workbooks\shared_obligation_tracker.xlsx open.
Use the MCP tools of server "gridcraft-lab" for every workbook action; Read only for the plan file.
Never call list_commands without a narrow search.

1. inspect_workbook: confirm the path above and dirty=false. If not, stop: FAIL.
2. execute_command sheet.activate {"sheet": "Сводка"}.
3. Read C:\repos\gridcraft-agent-lab\runs\STAGE_04\fix_plan.jsonl and execute every step in order
   with execute_command, "command" and "params" exactly as written. Stop on the first error and report it.
4. inspect_workbook: report conditional formats per sheet.
5. execute_command file.save {"path": "C:\repos\gridcraft-agent-lab\workbooks\shared_obligation_tracker.xlsx"}.
6. screenshot to C:\repos\gridcraft-agent-lab\runs\STAGE_04\screenshots\fix_summary.png
Do not quit the window.

Final answer: FIX: <executed>/<total>, errors; SAVE: <result>; STATE: <conditional formats, dirty>;
last line exactly: SESSION2B_VERDICT: PASS|FAIL — <reason>
