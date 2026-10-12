"""Stage 04 build plan: monthly shared-obligation tracker (payers A and B) as GridCraft commands.

Business model (user clarification 2026-10-13):
  - one obligation per settlement month (default 1100 EUR, per-month exceptions on «Настройки»);
  - payment is made for a past month; every payment row names the month it pays for;
  - expected B for a month = obligation - A payments for that month (not below 0);
  - remainder = obligation - A - B for that month; negative = overpayment;
  - months in the summary: from the accounting start month to the month before «Расчёт на дату»;
  - a payment counts only if its date is not after «Расчёт на дату»;
  - overpayments offset debts of other months in the overall total (one saldo).

Schema only: every line is a GridCraft command id with params from the verified catalog.
GridCraft evaluates every formula. Output: runs/STAGE_04/build_plan.jsonl
"""

import json
from pathlib import Path

LAB = Path(__file__).resolve().parents[1]
OUT = LAB / "runs" / "STAGE_04" / "build_plan.jsonl"

SUM, SET, PAY = "Сводка", "Настройки", "Поступления"

PAY_FIRST, PAY_LAST = 6, 505        # 500 payment rows
EXC_FIRST, EXC_LAST = 19, 42        # 24 obligation exceptions
M_FIRST, M_LAST = 12, 71            # 60 months
M_TOTAL = M_LAST + 1                # 72

INPUT_FILL, CALC_FILL, HEAD_FILL = "#FFF2CC", "#EDEDED", "#DDEBF7"
MONEY, DATE, MONTH = "#,##0.00", "dd.mm.yyyy", "mm.yyyy"

# Text-date parsing: accepts дд.мм.гггг and гггг-мм-дд, rejects impossible dates (31.02).
def parse_date(ref):
    return (f'IF(ISNUMBER({ref}),{ref},LET(t,TRIM({ref}),iso,MID(t,5,1)="-",'
            f'dd,IF(iso,VALUE(MID(t,9,2)),VALUE(LEFT(t,2))),mm,IF(iso,VALUE(MID(t,6,2)),VALUE(MID(t,4,2))),'
            f'yy,IF(iso,VALUE(LEFT(t,4)),VALUE(MID(t,7,4))),dt,DATE(yy,mm,dd),'
            f'IF(AND(LEN(t)=10,DAY(dt)=dd,MONTH(dt)=mm),dt,"ошибка")))')


# Month parsing: мм.гггг -> first day of month.
def parse_month(ref):
    return (f'IF(ISNUMBER({ref}),DATE(YEAR({ref}),MONTH({ref}),1),LET(t,TRIM({ref}),mm,VALUE(LEFT(t,2)),yy,VALUE(RIGHT(t,4)),'
            f'IF(AND(LEN(t)=7,MID(t,3,1)=".",mm>=1,mm<=12),DATE(yy,mm,1),"ошибка")))')


# GridCraft 0.6.1 evaluates validation on the value as it would be parsed (01.11.2026 -> a number),
# even in Text-formatted cells that then store the typed text. So numbers must pass here;
# the recognition columns and the row status catch anything the parser rejects.
def date_rule(ref):
    return (f'=OR({ref}="",ISNUMBER({ref}),AND(LEN({ref})=10,MID({ref},3,1)=".",MID({ref},6,1)="."),'
            f'AND(LEN({ref})=10,MID({ref},5,1)="-",MID({ref},8,1)="-"))')


def month_rule(ref):
    return f'=OR({ref}="",ISNUMBER({ref}),AND(LEN({ref})=7,MID({ref},3,1)="."))'


plan = []


def cmd(command, **params):
    plan.append({"command": command, "params": params})


def fill(rng, formula):
    # range.fill shifts relative references from the ACTIVE cell: select the target first.
    cmd("selection.set", range=rng, active=rng.split(":")[0])
    cmd("range.fill", range=rng, input=formula)


def header(rng, rows, height=32):
    cmd("home.bold", range=rng, on=True)
    cmd("home.fillColor", range=rng, color=HEAD_FILL)
    cmd("home.wrapText", range=rng, on=True)
    cmd("home.alignMiddle", range=rng)
    cmd("home.borders", range=rng, preset="all", style="thin")
    cmd("home.rowHeight", rows=rows, height=height)


def title(cell):
    cmd("home.bold", range=cell, on=True)
    cmd("home.fontSize", range=cell, size=14)


def text_input(rng):
    cmd("home.numberFormat", range=rng, format="Text")
    cmd("home.fillColor", range=rng, color=INPUT_FILL)


DATE_MSG = "Дата в виде дд.мм.гггг (например 15.10.2026) или гггг-мм-дд."
MONTH_MSG = "Месяц в виде мм.гггг, например 09.2026."

# ================================================================ sheets
cmd("sheet.rename", sheet=0, name=SUM)
cmd("home.insertSheet", name=SET)
cmd("home.insertSheet", name=PAY)

# ================================================================ Настройки
cmd("sheet.activate", sheet=SET)
text_input("B4")
text_input(f"A{EXC_FIRST}:A{EXC_LAST}")
cmd("range.setValues", range="A1", values=[
    ["Настройки: общее обязательство и плательщики"],
    ["Жёлтые ячейки — ввод, серые — формулы. Месяц вводите как мм.гггг (09.2026)."],
])
cmd("range.setValues", range="A4", values=[
    ["Начало учёта (мм.гггг)", "09.2026", f"={parse_month('B4')}"],
    ["Обязательство в месяц по умолчанию, EUR", 1100, None],
])
cmd("range.setValues", range="A7", values=[
    ["ID плательщика", "Плательщик", "Роль"],
    ["A", "Плательщик A (вымышленный)", "платит первым, сумма не фиксирована"],
    ["B", "Плательщик B (вымышленный)", "ожидается: обязательство месяца − платежи A за этот месяц"],
])
cmd("range.setValues", range="A11", values=[
    ["Правила расчёта"],
    ["1. За каждый месяц с начала учёта возникает одно общее обязательство (по умолчанию из B5, исключения — ниже)."],
    ["2. Оплата производится за прошлый месяц: месяц входит в расчёт, когда «Расчёт на дату» позже этого месяца."],
    ["3. Платёж относится к месяцу, указанному в строке поступления; частичные платежи складываются."],
    ["4. Ожидается от B = обязательство месяца − платежи A за месяц (не меньше 0). Остаток = обязательство − A − B."],
    ["5. Отрицательный остаток — переплата; в общем итоге переплаты зачитываются в долги других месяцев."],
])
cmd("range.setValues", range=f"A{EXC_FIRST - 1}", values=[[
    "Исключение: месяц (мм.гггг)", "Обязательство за месяц, EUR", "Месяц (распознан)", "Примечание"]])
fill(f"C{EXC_FIRST}:C{EXC_LAST}", f'=IF(A{EXC_FIRST}="","",{parse_month(f"A{EXC_FIRST}")})')
title("A1")
cmd("home.bold", range="A4:A5", on=True)
cmd("home.bold", range="A11", on=True)
header("A7:C7", "7:7", 20)
header(f"A{EXC_FIRST - 1}:D{EXC_FIRST - 1}", f"{EXC_FIRST - 1}:{EXC_FIRST - 1}")
cmd("home.fillColor", range="B5", color=INPUT_FILL)
cmd("home.fillColor", range="A8:C9", color=INPUT_FILL)
cmd("home.fillColor", range="C4", color=CALC_FILL)
cmd("home.borders", range="A4:C5", preset="all", style="thin")
cmd("home.borders", range="A8:C9", preset="all", style="thin")
cmd("home.numberFormat", range="C4", code=MONTH)
cmd("home.numberFormat", range="B5", code=MONEY)
cmd("home.fillColor", range=f"B{EXC_FIRST}:B{EXC_LAST}", color=INPUT_FILL)
cmd("home.fillColor", range=f"D{EXC_FIRST}:D{EXC_LAST}", color=INPUT_FILL)
cmd("home.fillColor", range=f"C{EXC_FIRST}:C{EXC_LAST}", color=CALC_FILL)
cmd("home.numberFormat", range=f"B{EXC_FIRST}:B{EXC_LAST}", code=MONEY)
cmd("home.numberFormat", range=f"C{EXC_FIRST}:C{EXC_LAST}", code=MONTH)
cmd("home.borders", range=f"A{EXC_FIRST}:D{EXC_LAST}", preset="all", style="thin")
cmd("home.columnWidth", cols="A:A", chars=40)
cmd("home.columnWidth", cols="B:B", chars=28)
cmd("home.columnWidth", cols="C:C", chars=58)
cmd("home.columnWidth", cols="D:D", chars=30)
cmd("data.validation", range="B4", type="custom", formula1=month_rule("B4"),
    errorStyle="stop", errorTitle="Начало учёта", errorMessage=MONTH_MSG)
cmd("data.validation", range="B5", type="decimal", operator="greaterOrEqual", formula1="0",
    errorStyle="stop", errorTitle="Обязательство", errorMessage="Сумма — число не меньше 0.")
cmd("data.validation", range=f"A{EXC_FIRST}:A{EXC_LAST}", type="custom",
    formula1=month_rule(f"A{EXC_FIRST}"), errorStyle="stop", errorTitle="Месяц", errorMessage=MONTH_MSG)
cmd("data.validation", range=f"B{EXC_FIRST}:B{EXC_LAST}", type="decimal", operator="greaterOrEqual",
    formula1="0", errorStyle="stop", errorTitle="Обязательство", errorMessage="Сумма — число не меньше 0.")
cmd("home.conditionalFormat", range=f"C4:C{EXC_LAST}",
    rule={"type": "beginsWith", "text": "ошибка", "preset": "lightRedFillDarkRedText"})

# ================================================================ Поступления
cmd("sheet.activate", sheet=PAY)
text_input(f"A{PAY_FIRST}:A{PAY_LAST}")
text_input(f"D{PAY_FIRST}:D{PAY_LAST}")
cmd("range.setValues", range="A1", values=[
    ["Поступления: одна выплата — одна строка"],
    [f"Вместимость: 500 строк ({PAY_FIRST}–{PAY_LAST}). Дата — дд.мм.гггг, месяц оплаты — мм.гггг (обязательно), плательщик — из списка."],
    ["Платёж учитывается за указанный месяц, если его дата не позже «Расчёт на дату» и месяц уже к оплате. Серые колонки — проверка ввода."],
])
cmd("range.setValues", range=f"A{PAY_FIRST - 1}", values=[[
    "Дата (дд.мм.гггг)", "Плательщик", "Сумма, EUR", "За месяц (мм.гггг)", "Примечание",
    "Дата (распознана)", "Месяц (распознан)", "Статус строки"]])
cmd("range.setValues", range=f"A{PAY_FIRST}", values=[
    ["05.10.2026", "A", 600, "09.2026", "контрольный пример"],
    ["12.10.2026", "B", 300, "09.2026", "контрольный пример: частичный платёж"],
    ["20.10.2026", "B", 150, "09.2026", "контрольный пример: частичный платёж"],
    ["03.11.2026", "A", 700, "10.2026", "контрольный пример"],
    ["10.11.2026", "B", 420, "10.2026", "контрольный пример: переплата 20"],
    ["14.11.2026", "A", 500, "11.2026", "ноябрь ещё не к оплате на 15.11"],
    ["16.11.2026", "B", 50, "09.2026", "после даты расчёта 15.11"],
])
r = PAY_FIRST
fill(f"F{r}:F{PAY_LAST}", f'=IF(A{r}="","",IFERROR({parse_date(f"A{r}")},"ошибка"))')
fill(f"G{r}:G{PAY_LAST}", f'=IF(D{r}="","",IFERROR({parse_month(f"D{r}")},"ошибка"))')
status = (
    f'=IF(AND(A{r}="",B{r}="",C{r}="",D{r}=""),"",'
    f'IF(NOT(ISNUMBER(F{r})),"ошибка: дата",'
    f"IF(COUNTIF('{SET}'!$A$8:$A$9,B{r})=0,\"ошибка: плательщик\","
    f'IF(NOT(ISNUMBER(C{r})),"ошибка: сумма",'
    f'IF(NOT(ISNUMBER(G{r})),"ошибка: месяц оплаты",'
    f"IF(F{r}>'{SUM}'!$C$4,\"после даты расчёта\","
    f"IF(G{r}<'{SET}'!$C$4,\"месяц до начала учёта\","
    f"IF(G{r}>'{SUM}'!$B$5,\"месяц ещё не к оплате\",\"учтено\"))))))))"
)
fill(f"H{r}:H{PAY_LAST}", status)
title("A1")
header(f"A{PAY_FIRST - 1}:H{PAY_FIRST - 1}", f"{PAY_FIRST - 1}:{PAY_FIRST - 1}")
cmd("home.fillColor", range=f"B{PAY_FIRST}:C{PAY_LAST}", color=INPUT_FILL)
cmd("home.fillColor", range=f"E{PAY_FIRST}:E{PAY_LAST}", color=INPUT_FILL)
cmd("home.fillColor", range=f"F{PAY_FIRST}:H{PAY_LAST}", color=CALC_FILL)
cmd("home.borders", range=f"A{PAY_FIRST}:H{PAY_LAST}", preset="all", style="thin")
cmd("home.numberFormat", range=f"C{PAY_FIRST}:C{PAY_LAST}", code=MONEY)
cmd("home.numberFormat", range=f"F{PAY_FIRST}:F{PAY_LAST}", code=DATE)
cmd("home.numberFormat", range=f"G{PAY_FIRST}:G{PAY_LAST}", code=MONTH)
cmd("home.alignRight", range=f"A{PAY_FIRST}:A{PAY_LAST}")
cmd("home.alignRight", range=f"D{PAY_FIRST}:D{PAY_LAST}")
cmd("home.alignCenter", range=f"B{PAY_FIRST}:B{PAY_LAST}")
for col, w in (("A", 17), ("B", 12), ("C", 13), ("D", 17), ("E", 38), ("F", 15), ("G", 15), ("H", 24)):
    cmd("home.columnWidth", cols=f"{col}:{col}", chars=w)
cmd("data.validation", range=f"A{PAY_FIRST}:A{PAY_LAST}", type="custom", formula1=date_rule(f"A{PAY_FIRST}"),
    errorStyle="stop", errorTitle="Дата", errorMessage=DATE_MSG)
cmd("data.validation", range=f"B{PAY_FIRST}:B{PAY_LAST}", type="list", formula1=f"='{SET}'!$A$8:$A$9",
    dropdown=True, errorStyle="stop", errorTitle="Плательщик", errorMessage="Выберите плательщика из списка (A или B).")
cmd("data.validation", range=f"C{PAY_FIRST}:C{PAY_LAST}", type="decimal", operator="greater", formula1="0",
    errorStyle="stop", errorTitle="Сумма", errorMessage="Сумма — положительное число.")
cmd("data.validation", range=f"D{PAY_FIRST}:D{PAY_LAST}", type="custom", formula1=month_rule(f"D{PAY_FIRST}"),
    errorStyle="stop", errorTitle="Месяц оплаты", errorMessage=MONTH_MSG)
cmd("home.conditionalFormat", range=f"F{PAY_FIRST}:H{PAY_LAST}",
    rule={"type": "beginsWith", "text": "ошибка", "preset": "lightRedFillDarkRedText"})
cmd("view.freezePanes", cell=f"A{PAY_FIRST}")

# ================================================================ Сводка
cmd("sheet.activate", sheet=SUM)
text_input("B4")
cmd("range.setValues", range="A1", values=[
    ["Общее обязательство: помесячная сводка"],
    ["Вымышленные данные, EUR. Жёлтые ячейки — ввод, серые — формулы. Правила — на листе «Настройки»."],
    ["Введите «Расчёт на дату» (дд.мм.гггг); платежи — на листе «Поступления»; месяц к оплате — каждый месяц до месяца даты расчёта."],
])
cmd("range.setValues", range="A4", values=[
    ["Расчёт на дату (дд.мм.гггг)", "15.11.2026", f"=IFERROR({parse_date('B4')},\"ошибка\")"],
    ["Последний месяц к оплате", f'=IF(ISNUMBER(C4),EDATE(DATE(YEAR(C4),MONTH(C4),1),-1),"ошибка")', None],
    ["Начало учёта", f"='{SET}'!C4", None],
    ["Обязательство по умолчанию", f"='{SET}'!B5", None],
    ["Учтено платежей, шт.", f"=COUNTIF('{PAY}'!H{PAY_FIRST}:H{PAY_LAST},\"учтено\")", None],
    ["Строк с ошибками, шт.", f"=COUNTIF('{PAY}'!H{PAY_FIRST}:H{PAY_LAST},\"ошибка*\")", None],
])
cmd("range.setValues", range="E4", values=[
    ["Обязательства за месяцы к оплате", f"=B{M_TOTAL}"],
    ["Получено за эти месяцы (A + B)", f"=C{M_TOTAL}+E{M_TOTAL}"],
    ["Сальдо с зачётом (+ долг / − переплата)", "=F4-F5"],
    ["Общий долг к получению", "=MAX(0,F6)"],
    ["Общая переплата", "=MAX(0,-F6)"],
    ["Предоплата за месяцы, ещё не к оплате", f"=SUMIFS('{PAY}'!C{PAY_FIRST}:C{PAY_LAST},'{PAY}'!H{PAY_FIRST}:H{PAY_LAST},\"месяц ещё не к оплате\")"],
])
cmd("range.setValues", range=f"A{M_FIRST - 1}", values=[[
    "Месяц", "Обязательство", "Получено A", "Ожидается B", "Получено B",
    "Остаток долга", "Переплата", "Сальдо нарастающим итогом"]])
m = M_FIRST
cmd("cell.set", cell=f"A{m}", input='=IF(AND(ISNUMBER($B$6),ISNUMBER($B$5)),IF($B$6<=$B$5,$B$6,""),"")')
fill(f"A{m + 1}:A{M_LAST}", f'=IF(ISNUMBER(A{m}),IF(EDATE(A{m},1)<=$B$5,EDATE(A{m},1),""),"")')
fill(f"B{m}:B{M_LAST}",
     f"=IF($A{m}=\"\",\"\",IFERROR(INDEX('{SET}'!$B${EXC_FIRST}:$B${EXC_LAST},MATCH($A{m},'{SET}'!$C${EXC_FIRST}:$C${EXC_LAST},0)),$B$7))")
received = ("=IF($A{m}=\"\",\"\",SUMIFS('" + PAY + "'!$C$6:$C$505,'" + PAY + "'!$B$6:$B$505,'" + SET +
            "'!$A${p},'" + PAY + "'!$G$6:$G$505,$A{m},'" + PAY + "'!$H$6:$H$505,\"учтено\"))")
fill(f"C{m}:C{M_LAST}", received.format(m=m, p=8))
fill(f"D{m}:D{M_LAST}", f'=IF($A{m}="","",MAX(0,B{m}-C{m}))')
fill(f"E{m}:E{M_LAST}", received.format(m=m, p=9))
fill(f"F{m}:F{M_LAST}", f'=IF($A{m}="","",MAX(0,B{m}-C{m}-E{m}))')
fill(f"G{m}:G{M_LAST}", f'=IF($A{m}="","",MAX(0,C{m}+E{m}-B{m}))')
cmd("cell.set", cell=f"H{m}", input=f'=IF($A{m}="","",B{m}-C{m}-E{m})')
fill(f"H{m + 1}:H{M_LAST}", f'=IF($A{m + 1}="","",H{m}+B{m + 1}-C{m + 1}-E{m + 1})')
cmd("range.setValues", range=f"A{M_TOTAL}", values=[[
    "Итого", f"=SUM(B{m}:B{M_LAST})", f"=SUM(C{m}:C{M_LAST})", f"=SUM(D{m}:D{M_LAST})",
    f"=SUM(E{m}:E{M_LAST})", f"=SUM(F{m}:F{M_LAST})", f"=SUM(G{m}:G{M_LAST})", "=F6"]])
title("A1")
cmd("home.bold", range="A4:A9", on=True)
cmd("home.bold", range="E4:E9", on=True)
cmd("home.bold", range="F7:F8", on=True)
cmd("home.borders", range="B4", preset="thickOutside")
cmd("home.fillColor", range="C4:C9", color=CALC_FILL)
cmd("home.fillColor", range="B5:B9", color=CALC_FILL)
cmd("home.fillColor", range="F4:F9", color=CALC_FILL)
cmd("home.numberFormat", range="C4", code=DATE)
cmd("home.numberFormat", range="B5:B6", code=MONTH)
cmd("home.numberFormat", range="B7", code=MONEY)
cmd("home.numberFormat", range="F4:F9", code=MONEY)
cmd("home.borders", range="A4:C9", preset="all", style="thin")
cmd("home.borders", range="E4:F9", preset="all", style="thin")
header(f"A{M_FIRST - 1}:H{M_FIRST - 1}", f"{M_FIRST - 1}:{M_FIRST - 1}")
cmd("home.fillColor", range=f"A{m}:H{M_TOTAL}", color=CALC_FILL)
cmd("home.borders", range=f"A{m}:H{M_TOTAL}", preset="all", style="thin")
cmd("home.numberFormat", range=f"A{m}:A{M_LAST}", code=MONTH)
cmd("home.numberFormat", range=f"B{m}:H{M_TOTAL}", code=MONEY)
cmd("home.alignLeft", range=f"A{m}:A{M_LAST}")
cmd("home.bold", range=f"A{M_TOTAL}:H{M_TOTAL}", on=True)
cmd("home.borders", range=f"A{M_TOTAL}:H{M_TOTAL}", preset="topDoubleBottom")
cmd("home.conditionalFormat", range=f"F{m}:F{M_LAST}",
    rule={"type": "cellIs", "operator": "greater", "value": "0", "preset": "lightRedFillDarkRedText"})
cmd("home.conditionalFormat", range=f"G{m}:G{M_LAST}",
    rule={"type": "cellIs", "operator": "greater", "value": "0", "preset": "greenFill"})
cmd("home.conditionalFormat", range="C4:C5",
    rule={"type": "beginsWith", "text": "ошибка", "preset": "lightRedFillDarkRedText"})
for col, w in (("A", 30), ("B", 14), ("C", 14), ("D", 14), ("E", 38), ("F", 16), ("G", 14), ("H", 16)):
    cmd("home.columnWidth", cols=f"{col}:{col}", chars=w)
cmd("data.validation", range="B4", type="custom", formula1=date_rule("B4"),
    errorStyle="stop", errorTitle="Расчёт на дату", errorMessage=DATE_MSG)
cmd("view.freezePanes", cell=f"A{M_FIRST}")
cmd("selection.set", range="B4")

if __name__ == "__main__":
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        for i, step in enumerate(plan, 1):
            f.write(json.dumps({"step": i, **step}, ensure_ascii=False) + "\n")
    print(f"{len(plan)} steps -> {OUT.relative_to(LAB)}")


# ================================================================ post-review fix (applied after the build)
# Visual review of the build screenshots: cellIs ">0" also matched the blank-text cells of months
# that are not due (text compares greater than numbers), painting empty rows; the H11 header was
# clipped. Applied to the built workbook as a separate plan so build_plan.jsonl stays as executed.
FIX_OUT = LAB / "runs" / "STAGE_04" / "fix_plan.jsonl"
fix = []
for rng, first, preset in ((f"F{M_FIRST}:F{M_LAST}", f"F{M_FIRST}", "lightRedFillDarkRedText"),
                           (f"G{M_FIRST}:G{M_LAST}", f"G{M_FIRST}", "greenFill")):
    fix.append({"command": "home.clearRules", "params": {"range": rng}})
    fix.append({"command": "home.conditionalFormat", "params": {
        "range": rng, "rule": {"type": "expression", "formula": f"=AND(ISNUMBER({first}),{first}>0)", "preset": preset}}})
fix.append({"command": "home.columnWidth", "params": {"cols": "H:H", "chars": 20}})
fix.append({"command": "selection.set", "params": {"range": "B4"}})

if __name__ == "__main__":
    with open(FIX_OUT, "w", encoding="utf-8") as f:
        for i, step in enumerate(fix, 1):
            f.write(json.dumps({"step": i, **step}, ensure_ascii=False) + "\n")
    print(f"{len(fix)} fix steps -> {FIX_OUT.relative_to(LAB)}")
