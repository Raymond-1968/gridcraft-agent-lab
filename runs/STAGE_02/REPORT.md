# Stage 02 — подключение Claude Code к GridCraft через MCP

**Итог: FAIL на шаге 1** (каталог команд получен, но не читается внутри сессии).
Подключение MCP и шаги 2–5 — **PASS**: A3 = 42, книга сохранена, скриншот получен.

Прогон 2026-10-12, 23:13–23:15 (UTC+3). Синтетические данные. Форма доходов не создавалась.

## Как проверялось

- Отдельная сессия Claude Code 2.1.289 (бинарь расширения VS Code), запущенная в
  `C:\repos\gridcraft-agent-lab` командой
  `claude -p <session_prompt.md> --mcp-config .mcp.json --strict-mcp-config --tools "" --allowedTools mcp__gridcraft-lab --output-format stream-json --verbose`.
- Сервер — локальный `.mcp.json` (`gridcraft-cli mcp --connect 7991`), лабораторное окно на порту 7991
  (`scripts/start_lab_window.ps1`, `GRIDCRAFT_NO_PREFS=1`).
- Встроенные инструменты сессии отключены: ей были доступны только 21 инструмент `mcp__gridcraft-lab__*`.
  Ни `stage01_smoke.py`, ни другой самописный MCP-клиент не использовались.
- В init-событии сессии: `gridcraft-lab` — `status: connected`.

## Результаты (попытка 2, `attempt_02/`)

| Шаг | Инструменты сессии | Результат |
|---|---|---|
| 1. Каталог и состояние книги | `list_commands`, `inspect_workbook` | **FAIL**: ответ `list_commands` (96 КБ) Claude Code сохранил в файл вместо контекста; прочитать его только MCP-инструментами сессия не может. `inspect_workbook` — OK |
| 2. Пустая книга | `new_workbook`, `inspect_workbook` | PASS: `Book2`, 0 ячеек, undo пуст |
| 3. 17, 25, сумма | `set_cell` ×3 | PASS: A1=17, A2=25, A3=`=SUM(A1:A2)` |
| 4. Результат | `get_cell`, `read_range` | PASS: A3 value 42, `[[17],[25],[42]]` |
| 5. Сохранение и скриншот | `save_workbook`, `screenshot`, `inspect_workbook` | PASS: `workbooks/stage02_claude_session_test_attempt02.xlsx` (4048 байт), скриншот 1440×900, `dirty: false` |

Сессия сама вынесла вердикт `STAGE02_VERDICT: FAIL` (итоговый текст — в `claude_session.stream.jsonl`).

Проверки после прогона (не подменяют проверку, `post_checks.txt`): `gridcraft-cli cat` сохранённой
книги показывает `17, 25, =SUM(A1:A2)` и значения `17, 25, 42`; окно закрыто `app.quit`, процессов не осталось.
Сохранённые сессией ответы `list_commands` содержат полный каталог — 352 команды (обе попытки).

## Попытки

1. `attempt_01/` — лимит по умолчанию: ответ 86 394 символа превысил лимит MCP-вывода
   («exceeds maximum allowed tokens»), сохранён в файл. Шаги 2–5 PASS.
   Скриншот перенесён из `runs/STAGE_02/screenshots/` в `attempt_01/screenshots/` после прогона,
   поэтому в журнале указан прежний путь. Книга: `workbooks/stage02_claude_session_test.xlsx`.
2. `attempt_02/` — `MAX_MCP_OUTPUT_TOKENS=60000` только для процесса сессии. Ответ (96,1 КБ) снова
   вынесен в файл («Output too large … Full output saved to …») — отдельный порог размера. Шаги 2–5 PASS.

Третью попытку подбором настроек не делал.

## Вывод и вариант решения (не проверен)

Агент через MCP полностью управляет GridCraft, но полный каталог одним вызовом в сессию не помещается.
Рабочий путь для следующих стадий — `list_commands` с `search` (например `cell.`, `home.`, `file.`),
каждая выборка по размеру укладывается в лимит. Это не проверялось на этой стадии.

## Файлы

- `launcher.txt` — запуск окна; `post_checks.txt` — проверки после прогона и закрытие окна.
- `attempt_0N/session_prompt.md` — задание сессии; `claude_session.stream.jsonl` — полный поток
  (вызовы MCP, ответы, итог); `catalog/` — ответ `list_commands`; `screenshots/stage02_result.png`.
- Книги `.xlsx` в git не входят.
