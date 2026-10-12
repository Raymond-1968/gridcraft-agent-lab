# Stage 02 correction 01 — выборочное раскрытие команд

**Итог: PASS.** Прогон 2026-10-12 23:28 (UTC+3). Синтетические данные. Исторический отчёт
[Stage 02](../STAGE_02/REPORT.md) не изменялся.

## Условия

- Отдельная сессия Claude Code 2.1.289 в `C:\repos\gridcraft-agent-lab`:
  `claude -p <session_prompt.md> --mcp-config .mcp.json --strict-mcp-config --tools "" --allowedTools mcp__gridcraft-lab --output-format stream-json --verbose`.
- Лимиты вывода по умолчанию (`MAX_MCP_OUTPUT_TOKENS` не задан). Полный каталог не запрашивался:
  все три вызова `list_commands` — с `search`.
- Лабораторное окно `scripts/start_lab_window.ps1` (порт 7991, `GRIDCRAFT_NO_PREFS=1`); сервер `gridcraft-lab` — `connected`.

## Результаты

| Шаг | Вызовы MCP | Результат |
|---|---|---|
| 1. Узкие поиски | `list_commands` `cell.set` / `home.bold` / `file.save` | PASS: ответы 210 / 185 / 645 символов, все прямо в контексте (ничего не вынесено в файл) |
| 2. Описания в контексте | — | PASS: `cell.set` «Enter Cell» `{cell?: "B2", input: "text, number or =formula", sheet?, array?: bool}`; `home.bold` «Bold» `{range?, on?: bool}`; `file.save` «Save» `{path?} (xlsx by default; .csv/.tsv/.json/.html/.pdf by extension)`. Поиск `file.save` вернул также `file.saveAs` и `file.saveBytes` |
| 3. Выполнение | `new_workbook`; `execute_command` `cell.set` B2=123.45, `home.bold` B2, `file.save` с путём | PASS: копия `workbooks/stage02_correction01_synthetic_copy.xlsx` (4043 байта) |
| 4. Проверка через MCP | `get_cell` B2, `inspect_workbook`, `screenshot` | PASS: value 123.45 (number), `style.font.bold: true`; path — сохранённая копия, `dirty: false`, undo `["Bold","Enter Cell"]` |

Итог сессии: `CORRECTION_VERDICT: PASS`.

Проверки после прогона (`post_checks.txt`): `gridcraft-cli cat` копии показывает 123.45 в B2;
окно закрыто через `app.quit`, процессов не осталось.

## Закреплено

- `catalog/gridcraft-0.6.1/` — локальный каталог (352 команды окна, 339 CLI, 21 MCP-инструмент,
  513 функций), скопирован из `runs/STAGE_01/catalog`; `manifest.json` с версией 0.6.1 и SHA-256
  обоих бинарей и всех файлов каталога. Хэши бинарей перепроверены — совпадают со Stage 01.
- `scripts/catalog.py verify` / `search <term>` — проверка привязки к билду и локальный выборочный поиск
  (поиск отказывает, если проверка не прошла).
- `.gitattributes`: `catalog/** -text`, чтобы git не менял байты файлов, привязанных к хэшам.
- `AGENTS.md`: раздел «Command discovery» — каталог действителен только для билда из манифеста;
  `list_commands` только с узким `search`; повышение лимитов не заменяет выборочное раскрытие.

## Файлы

- `session_prompt.md` — задание сессии; `claude_session.stream.jsonl` — полный поток сессии.
- `mcp_calls.jsonl` — извлечённые запросы и ответы MCP (10 вызовов).
- `screenshots/result.png`; `launcher.txt`; `post_checks.txt`; `session_meta.txt`.
