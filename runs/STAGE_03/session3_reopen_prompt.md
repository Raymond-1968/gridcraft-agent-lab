Stage 03, session 3 — reopen verification of the income tracker prototype.
Lab: C:\repos\gridcraft-agent-lab. The previous GridCraft window was closed (app.quit) after saving.
A new lab window (port 7991) was started with an empty book.
Use the MCP tools of server "gridcraft-lab" for every workbook action; Read only for plan files.
Never call list_commands without a narrow search.

1. open_workbook {"path": "C:\\repos\\gridcraft-agent-lab\\workbooks\\income_tracker_prototype.xlsx"}.
2. inspect_workbook: report sheet names and order, used ranges, freeze, validations, conditional
   formats per sheet, dirty, path.
3. Values: Read C:\repos\gridcraft-agent-lab\runs\STAGE_03\checks_plan.jsonl and execute ONLY the
   lines whose "check" is "C0" (all are read_range with "expect"). Compare as follows: numbers
   match if |actual - expected| < 0.005; "" matches empty or null.
4. Snapshot: Read C:\repos\gridcraft-agent-lab\runs\STAGE_03\snapshot_cells.json. For each entry in
   order: execute_command sheet.activate {"sheet": sheet}, then get_cell {"cell": cell}. Compare style
   with the entry: fill = style.fill.fg.Rgb, num_fmt = style.num_fmt, bold = style.font.bold.
   Report formula and value of every snapshot cell.
5. No-error scan: read_range sheet "Сводка" A4:H33 and sheet "Поступления" F6:F505; no error values.
6. Validation still active: execute_command sheet.activate {"sheet":"Поступления"}, then
   execute_command data.validate {"cell":"B6","input":"ZZ"} (expect ok=false) and
   {"cell":"B6","input":"A"} (expect ok=true).
7. Screenshots of the reopened book: for "Сводка" (summary), "Источники" (sources),
   "Поступления" (payments): execute_command sheet.activate {"sheet": name}, execute_command
   selection.set {"range": "A1"}, then screenshot to
   C:\repos\gridcraft-agent-lab\runs\STAGE_03\screenshots\reopen_<n>_<summary|sources|payments>.png
   Finally activate "Сводка" and selection.set {"range":"B4"}.
   Do not save (nothing was changed). Do not quit the window.

Final answer (plain text):
OPEN: ... / STATE: ... / VALUES C0: PASS|FAIL with mismatches / SNAPSHOT: n/total, mismatches /
NO_ERRORS: PASS|FAIL / VALIDATION: PASS|FAIL / SCREENSHOTS: paths
last line exactly: SESSION3_VERDICT: PASS|FAIL — <reason>
