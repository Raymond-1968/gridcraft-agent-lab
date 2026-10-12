# Stage 01 — подготовка лаборатории и проверка подключения

**Итог: PASS** — 33 из 33 проверок (`results.json`), прогон 2026-10-12 ≈23:07 по местному времени.

Форма учёта доходов на этой стадии не создавалась. Использовались только синтетические числа.

## Среда

- GridCraft 0.6.1: `C:\Program Files\GridCraft\gridcraft.exe`; CLI: `gridcraft-cli.exe` 0.6.1.
  Совпадает с последним релизом upstream `v0.6.1`. Ничего не обновлялось и не переустанавливалось.
  Пути, хэши и проверенная документация — в [docs/ENVIRONMENT_BASELINE.md](../../docs/ENVIRONMENT_BASELINE.md).
- Лабораторное окно: `gridcraft.exe --control 7991`, переменная `GRIDCRAFT_NO_PREFS=1` только для
  дочернего процесса (настройки не читаются и не сохраняются), рабочая папка `tmp/`.
- Агент: `gridcraft-cli mcp --connect 7991` (MCP JSON-RPC 2.0 по stdio, протокол 2025-06-18),
  проектный `.mcp.json`; глобальные настройки MCP не менялись.

## Что проверено

| Этап | Проверка | Результат |
|---|---|---|
| 1 | Окно запущено, control channel слушает `127.0.0.1:7991`; `document.inspect`, `engine.commands` (352 команды) | PASS |
| 2 | MCP `initialize` (serverInfo GridCraft 0.6.1), `tools/list` (21 инструмент), стартовая книга пуста | PASS |
| 3 | `set_cell` A1=12.5, A2=30, A3=`=SUM(A1:A2)`; GridCraft вычислил A3=42.5; `read_range` A1:A3 | PASS |
| 4 | `format_range` A3: полужирный, заливка `#FFF2CC`, формат `0.00` → отображается `42.50` | PASS |
| 5 | A2=100 → A3=112.5; `undo` → 42.5; `redo` → 112.5; `undo` → 42.5; курсив `home.italic` + undo/redo | PASS |
| 6 | `save_workbook` → `workbooks/stage01_smoke_test.xlsx` (4134 байта); скриншот; `engine.journal`; `app.quit` (exit 0) | PASS |
| 7 | Новый процесс окна, `open_workbook`: формула, значение 42.5, `42.50`, полужирный + курсив, заливка сохранились; скриншот | PASS |

Дополнительно вручную: `scripts/start_lab_window.ps1` + `gridcraft-cli send 7991 document.inspect` /
`app.quit` — работает ([logs/launcher_and_cli_send_check.txt](logs/launcher_and_cli_send_check.txt)).

Все расчёты и изменения книги выполнил GridCraft: `engine.journal` содержит `cell.set`,
`home.bold`, `home.fillColor`, `home.numberFormat`, `home.italic` и т. д.
Скрипт только отправляет запросы и сравнивает ответы.

## Материалы

- Каталог команд: [catalog/commands.txt](catalog/commands.txt), `commands.json` (339, CLI),
  `commands_live_window.json` (352, окно), `mcp_tools.json` (21), `functions.json` (513).
- Журнал запросов и ответов (MCP + control channel + проверки): [logs/transcript.jsonl](logs/transcript.jsonl).
- Журнал движка: [logs/engine_journal_session1.json](logs/engine_journal_session1.json).
- Скриншот результата: [screenshots/stage01_result_reopened.png](screenshots/stage01_result_reopened.png)
  (после повторного открытия), а также `stage01_session1_after_save.png`.
- Сама книга `.xlsx` в git не входит (`.gitignore`).

## Замечания

1. `gridcraft.exe --help` не поддерживается: при разведке он запустил обычное окно (не
   изолированное, ~8 с, закрыто мной). Этот запуск создал `%APPDATA%\GridCraft\logs\gridcraft.log`;
   файл настроек `ui.json` не создавался. Папку я не удалял — она вне лаборатории.
2. Первая попытка проверить лаунчер зависла из-за моей обвязки теста (захват вывода через
   конвейер, унаследованный окном), а не из-за GridCraft; окно и зависший дочерний PowerShell
   этой попытки я остановил. Повторная проверка прошла.
3. 13 команд доступны только в окне (язык, тема, печать, диалоги, ленточные `ui.*`); команды
   темы и языка сохраняют настройки, поэтому в лаборатории их не используем.
