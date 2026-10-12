"""Generates the Stage 03 build plan: GridCraft commands (JSONL) for the income tracker prototype.

The plan is a schema only: every line is a GridCraft command id with params taken from the
verified local catalog (catalog/gridcraft-0.6.1). Nothing here computes workbook values;
GridCraft evaluates all formulas. Output: runs/STAGE_03/build_plan.jsonl

Each line: {"command": "<id>", "params": {...}} - runnable via MCP `execute_command`
(and, for dry runs only, `gridcraft-cli run --script`).
"""

import json
from pathlib import Path

LAB = Path(__file__).resolve().parents[1]
OUT = LAB / "runs" / "STAGE_03" / "build_plan.jsonl"
BOOK = LAB / "workbooks" / "income_tracker_prototype.xlsx"

SRC, PAY, SUM = "Источники", "Поступления", "Сводка"
Q_SRC, Q_PAY, Q_SUM = f"'{SRC}'", f"'{PAY}'", f"'{SUM}'"

SRC_FIRST, SRC_LAST = 8, 27          # 20 source rows
PAY_FIRST, PAY_LAST = 6, 505         # 500 payment rows
SUM_FIRST = 13                       # summary rows mirror source rows
SUM_LAST = SUM_FIRST + (SRC_LAST - SRC_FIRST)   # 32
SUM_TOTAL = SUM_LAST + 1                        # 33

INPUT_FILL = "#FFF2CC"     # yellow: user input
CALC_FILL = "#EDEDED"      # grey: formulas
HEAD_FILL = "#DDEBF7"      # blue: headers
MONEY = "#,##0.00"
DATE = "dd.mm.yyyy"
MONTH = "mm.yyyy"

plan = []


def cmd(command, **params):
    plan.append({"command": command, "params": params})


def fill(rng, formula):
    # range.fill shifts relative references from the ACTIVE cell of the current selection,
    # so select the target range first (observed in GridCraft 0.6.1).
    cmd("selection.set", range=rng, active=rng.split(":")[0])
    cmd("range.fill", range=rng, input=formula)


def header(rng):
    cmd("home.bold", range=rng, on=True)
    cmd("home.fillColor", range=rng, color=HEAD_FILL)
    cmd("home.wrapText", range=rng, on=True)
    cmd("home.alignMiddle", range=rng)
    cmd("home.borders", range=rng, preset="all", style="thin")


def title(cell):
    cmd("home.bold", range=cell, on=True)
    cmd("home.fontSize", range=cell, size=14)


# ---------------------------------------------------------------- sheets
cmd("sheet.rename", sheet=0, name=SUM)
cmd("home.insertSheet", name=SRC)
cmd("home.insertSheet", name=PAY)

# ---------------------------------------------------------------- Источники
cmd("sheet.activate", sheet=SRC)
cmd("range.setValues", range="A1", values=[
    ["Источники доходов и ежемесячные обязательства"],
    ["Жёлтые ячейки — ввод, серые — формулы. Месяц вводите датой первого дня в виде ГГГГ-ММ-ДД, например 2026-09-01."],
])
cmd("range.setValues", range="A4", values=[
    ["Начало учёта (месяц)", "2026-09-01", None,
     "Правило прототипа: начисление за месяц возникает целиком 1-го числа месяца."],
    [None, None, None,
     "Начальный долг — долг за период до начала учёта; он не начисляется повторно."],
    [None, None, None,
     "Просрочка не рассчитывается: сроки оплаты не заданы. Месяцы до начала учёта не начисляются."],
])
cmd("range.setValues", range=f"A{SRC_FIRST - 1}", values=[[
    "ID", "Название", "Сумма в месяц, EUR", "Первый месяц начисления",
    "Последний месяц начисления (пусто = продолжается)", "Начальный долг на начало учёта, EUR", "Примечание"]])
cmd("range.setValues", range=f"A{SRC_FIRST}", values=[
    ["A", "Источник A (вымышленный)", 300, "2026-09-01", None, 100, "контрольный пример"],
    ["B", "Источник B (вымышленный)", 200, "2026-10-01", None, 0, "контрольный пример"],
])
title("A1")
header(f"A{SRC_FIRST - 1}:G{SRC_FIRST - 1}")
cmd("home.bold", range="A4", on=True)
cmd("home.fillColor", range="B4", color=INPUT_FILL)
cmd("home.borders", range="B4", preset="outside", style="thin")
cmd("home.numberFormat", range="B4", code=MONTH)
cmd("home.fillColor", range=f"A{SRC_FIRST}:G{SRC_LAST}", color=INPUT_FILL)
cmd("home.borders", range=f"A{SRC_FIRST}:G{SRC_LAST}", preset="all", style="thin")
cmd("home.numberFormat", range=f"C{SRC_FIRST}:C{SRC_LAST}", code=MONEY)
cmd("home.numberFormat", range=f"F{SRC_FIRST}:F{SRC_LAST}", code=MONEY)
cmd("home.numberFormat", range=f"D{SRC_FIRST}:E{SRC_LAST}", code=MONTH)
cmd("home.columnWidth", cols="A:A", chars=8)
cmd("home.columnWidth", cols="B:B", chars=28)
cmd("home.columnWidth", cols="C:C", chars=14)
cmd("home.columnWidth", cols="D:E", chars=18)
cmd("home.columnWidth", cols="F:F", chars=18)
cmd("home.columnWidth", cols="G:G", chars=30)
cmd("data.validation", range="B4", type="date", operator="between",
    formula1="=DATE(2000,1,1)", formula2="=DATE(2100,12,31)",
    errorStyle="stop", errorTitle="Дата", errorMessage="Введите дату в виде ГГГГ-ММ-ДД, например 2026-09-01.")
cmd("data.validation", range=f"C{SRC_FIRST}:C{SRC_LAST}", type="decimal", operator="greaterOrEqual",
    formula1="0", errorStyle="stop", errorTitle="Сумма", errorMessage="Сумма в месяц — число не меньше 0.")
cmd("data.validation", range=f"D{SRC_FIRST}:E{SRC_LAST}", type="date", operator="between",
    formula1="=DATE(2000,1,1)", formula2="=DATE(2100,12,31)",
    errorStyle="stop", errorTitle="Месяц", errorMessage="Введите дату первого дня месяца в виде ГГГГ-ММ-ДД.")
cmd("data.validation", range=f"F{SRC_FIRST}:F{SRC_LAST}", type="decimal", operator="greaterOrEqual",
    formula1="0", errorStyle="stop", errorTitle="Начальный долг", errorMessage="Начальный долг — число не меньше 0.")
cmd("view.freezePanes", cell=f"A{SRC_FIRST}")

# ---------------------------------------------------------------- Поступления
cmd("sheet.activate", sheet=PAY)
cmd("range.setValues", range="A1", values=[
    ["Поступления: одна выплата — одна строка"],
    [f"Вместимость: 500 строк ({PAY_FIRST}–{PAY_LAST}). Дата — в виде ГГГГ-ММ-ДД, например 2026-10-15. ID — из списка (лист «Источники»)."],
    ["В итог входят выплаты с датой от начала учёта до «Расчёт на дату» (лист «Сводка»). «За какой период» — справочно, на расчёт не влияет."],
])
cmd("range.setValues", range=f"A{PAY_FIRST - 1}", values=[[
    "Дата поступления", "ID источника", "Сумма, EUR", "За какой период (справочно)", "Примечание", "Статус строки"]])
cmd("range.setValues", range=f"A{PAY_FIRST}", values=[
    ["2026-09-15", "A", 150, "сентябрь 2026", "контрольный пример"],
    ["2026-10-10", "A", 200, "октябрь 2026", "контрольный пример"],
    ["2026-10-20", "B", 250, "октябрь 2026", "контрольный пример"],
    ["2026-11-05", "A", 100, "октябрь 2026", "после даты расчёта — исключается"],
])
r = PAY_FIRST
status = (
    f'=IF(AND(A{r}="",B{r}="",C{r}=""),"",'
    f'IF(NOT(ISNUMBER(A{r})),"ошибка: дата",'
    f'IF(COUNTIF({Q_SRC}!$A${SRC_FIRST}:$A${SRC_LAST},B{r})=0,"ошибка: ID",'
    f'IF(NOT(ISNUMBER(C{r})),"ошибка: сумма",'
    f'IF(A{r}<{Q_SRC}!$B$4,"до начала учёта",'
    f'IF(A{r}>{Q_SUM}!$B$4,"после даты расчёта","учтено"))))))'
)
fill(f"F{PAY_FIRST}:F{PAY_LAST}", status)
title("A1")
header(f"A{PAY_FIRST - 1}:F{PAY_FIRST - 1}")
cmd("home.fillColor", range=f"A{PAY_FIRST}:E{PAY_LAST}", color=INPUT_FILL)
cmd("home.fillColor", range=f"F{PAY_FIRST}:F{PAY_LAST}", color=CALC_FILL)
cmd("home.borders", range=f"A{PAY_FIRST}:F{PAY_LAST}", preset="all", style="thin")
cmd("home.numberFormat", range=f"A{PAY_FIRST}:A{PAY_LAST}", code=DATE)
cmd("home.numberFormat", range=f"C{PAY_FIRST}:C{PAY_LAST}", code=MONEY)
cmd("home.columnWidth", cols="A:A", chars=16)
cmd("home.columnWidth", cols="B:B", chars=12)
cmd("home.columnWidth", cols="C:C", chars=14)
cmd("home.columnWidth", cols="D:D", chars=22)
cmd("home.columnWidth", cols="E:E", chars=34)
cmd("home.columnWidth", cols="F:F", chars=20)
cmd("data.validation", range=f"A{PAY_FIRST}:A{PAY_LAST}", type="date", operator="between",
    formula1=f"={Q_SRC}!$B$4", formula2="=DATE(2100,12,31)",
    errorStyle="stop", errorTitle="Дата",
    errorMessage="Дата в виде ГГГГ-ММ-ДД, не раньше начала учёта (лист «Источники»).")
cmd("data.validation", range=f"B{PAY_FIRST}:B{PAY_LAST}", type="list",
    formula1=f"={Q_SRC}!$A${SRC_FIRST}:$A${SRC_LAST}", dropdown=True,
    errorStyle="stop", errorTitle="ID источника", errorMessage="Выберите ID из списка (лист «Источники»).")
cmd("data.validation", range=f"C{PAY_FIRST}:C{PAY_LAST}", type="decimal", operator="greater",
    formula1="0", errorStyle="stop", errorTitle="Сумма", errorMessage="Сумма — положительное число.")
cmd("home.conditionalFormat", range=f"F{PAY_FIRST}:F{PAY_LAST}",
    rule={"type": "beginsWith", "text": "ошибка", "preset": "lightRedFillDarkRedText"})
cmd("view.freezePanes", cell=f"A{PAY_FIRST}")

# ---------------------------------------------------------------- Сводка
cmd("sheet.activate", sheet=SUM)
cmd("range.setValues", range="A1", values=[
    ["Учёт регулярных доходов — сводка"],
    ["Прототип. Вымышленные данные. Валюта — EUR. Жёлтые ячейки — ввод, серые — формулы."],
])
cmd("range.setValues", range="A4", values=[
    ["Расчёт на дату", "2026-10-31"],
    ["Начало учёта", f"={Q_SRC}!B4"],
    ["Учтено поступлений, шт.", f'=COUNTIF({Q_PAY}!F{PAY_FIRST}:F{PAY_LAST},"учтено")'],
    ["Строк с ошибками, шт.", f'=COUNTIF({Q_PAY}!F{PAY_FIRST}:F{PAY_LAST},"ошибка*")'],
    ["Общий долг (должны пользователю)", f"=F{SUM_TOTAL}"],
    ["Общий аванс / переплата", f"=G{SUM_TOTAL}"],
    ["Чистый баланс (долг − аванс)", "=B8-B9"],
])
cmd("range.setValues", range=f"A{SUM_FIRST - 1}", values=[[
    "ID", "Источник", "Начальный долг", "Начислено за период", "Получено",
    "Остаток долга", "Аванс / переплата", "Месяцев начислено"]])
s = SUM_FIRST
src_row = SRC_FIRST
months = (
    f'=IF($A{s}="","",IF({Q_SRC}!D{src_row}="",0,LET('
    f'accStart,{Q_SRC}!$B$4,firstM,{Q_SRC}!D{src_row},lastM,{Q_SRC}!E{src_row},asOf,$B$4,'
    f'fromM,MAX(DATE(YEAR(accStart),MONTH(accStart),1),DATE(YEAR(firstM),MONTH(firstM),1)),'
    f'toM,IF(lastM="",DATE(YEAR(asOf),MONTH(asOf),1),MIN(DATE(YEAR(asOf),MONTH(asOf),1),DATE(YEAR(lastM),MONTH(lastM),1))),'
    f'MAX(0,(YEAR(toM)-YEAR(fromM))*12+MONTH(toM)-MONTH(fromM)+1))))'
)
received = (
    f'=IF($A{s}="","",SUMIFS({Q_PAY}!$C${PAY_FIRST}:$C${PAY_LAST},'
    f'{Q_PAY}!$B${PAY_FIRST}:$B${PAY_LAST},$A{s},'
    f'{Q_PAY}!$A${PAY_FIRST}:$A${PAY_LAST},">="&$B$5,'
    f'{Q_PAY}!$A${PAY_FIRST}:$A${PAY_LAST},"<="&$B$4))'
)
rows = SUM_LAST
fill(f"A{s}:A{rows}", f'=IF({Q_SRC}!A{src_row}="","",{Q_SRC}!A{src_row})')
fill(f"B{s}:B{rows}", f'=IF($A{s}="","",{Q_SRC}!B{src_row})')
fill(f"C{s}:C{rows}", f'=IF($A{s}="","",N({Q_SRC}!F{src_row}))')
fill(f"H{s}:H{rows}", months)
fill(f"D{s}:D{rows}", f'=IF($A{s}="","",H{s}*N({Q_SRC}!C{src_row}))')
fill(f"E{s}:E{rows}", received)
fill(f"F{s}:F{rows}", f'=IF($A{s}="","",MAX(0,C{s}+D{s}-E{s}))')
fill(f"G{s}:G{rows}", f'=IF($A{s}="","",MAX(0,E{s}-C{s}-D{s}))')
cmd("range.setValues", range=f"A{SUM_TOTAL}", values=[[
    "Итого", None, f"=SUM(C{s}:C{rows})", f"=SUM(D{s}:D{rows})", f"=SUM(E{s}:E{rows})",
    f"=SUM(F{s}:F{rows})", f"=SUM(G{s}:G{rows})", None]])
cmd("range.setValues", range="J4", values=[
    ["Как пользоваться"],
    ["1. «Источники»: начало учёта и по строке на источник (ID, сумма в месяц, месяцы, начальный долг)."],
    ["2. «Поступления»: каждая выплата — отдельная строка (дата ГГГГ-ММ-ДД, ID из списка, сумма)."],
    ["3. Здесь задайте «Расчёт на дату» — итоги пересчитаются автоматически."],
    ["Остаток долга > 0 — источник должен пользователю; переплата показана отдельно как аванс."],
    ["Аванс одного источника не погашает долг другого. Выплаты после даты расчёта не учитываются."],
    ["Вместимость: 20 источников, 500 строк поступлений. Строки с ошибками отмечены на «Поступлениях»."],
])
title("A1")
cmd("home.bold", range="J4", on=True)
cmd("home.bold", range="A4:A10", on=True)
cmd("home.fillColor", range="B4", color=INPUT_FILL)
cmd("home.borders", range="B4", preset="thickOutside")
cmd("home.numberFormat", range="B4:B5", code=DATE)
cmd("home.fillColor", range="B5:B10", color=CALC_FILL)
cmd("home.numberFormat", range="B8:B10", code=MONEY)
cmd("home.borders", range="A4:B10", preset="all", style="thin")
header(f"A{SUM_FIRST - 1}:H{SUM_FIRST - 1}")
cmd("home.fillColor", range=f"A{s}:H{SUM_TOTAL}", color=CALC_FILL)
cmd("home.borders", range=f"A{s}:H{SUM_TOTAL}", preset="all", style="thin")
cmd("home.numberFormat", range=f"C{s}:G{SUM_TOTAL}", code=MONEY)
cmd("home.bold", range=f"A{SUM_TOTAL}:H{SUM_TOTAL}", on=True)
cmd("home.borders", range=f"A{SUM_TOTAL}:H{SUM_TOTAL}", preset="topDoubleBottom")
cmd("home.columnWidth", cols="A:A", chars=34)
cmd("home.columnWidth", cols="B:B", chars=26)
cmd("home.columnWidth", cols="C:H", chars=15)
cmd("home.columnWidth", cols="I:I", chars=3)
cmd("home.columnWidth", cols="J:J", chars=95)
cmd("data.validation", range="B4", type="date", operator="between",
    formula1="=DATE(2000,1,1)", formula2="=DATE(2100,12,31)",
    errorStyle="stop", errorTitle="Расчёт на дату", errorMessage="Введите дату в виде ГГГГ-ММ-ДД, например 2026-10-31.")
cmd("view.freezePanes", cell=f"A{SUM_FIRST}")
cmd("selection.set", range="B4")

if __name__ == "__main__":
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        for i, step in enumerate(plan, 1):
            f.write(json.dumps({"step": i, **step}, ensure_ascii=False) + "\n")
    print(f"{len(plan)} steps -> {OUT.relative_to(LAB)}")
