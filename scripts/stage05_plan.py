"""Stage 05 plan: simplified copy with one cumulative balance (user clarification 2026-10-13).

- Copy: open the Stage 04 workbook and Save As a new file; the Stage 04 file is not modified.
- Model: accumulated obligation = months due (accounting start .. month before «Расчёт на дату»)
  x default obligation, with per-month exceptions from «Настройки». Every payment of A or B dated
  from the accounting start to «Расчёт на дату» is subtracted, whatever month it was meant for.
  Remainder > 0 = still to receive; < 0 = overpayment. No payment month, no monthly allocation.
- One working sheet «Учёт»: summary on top, entry table below (date / payer / amount / note /
  «Проверка» shown only for problems). Recognition helpers live in hidden columns H:I.
- Dates are typed as дд.мм.гггг (or гггг-мм-дд) into a Text column; Ctrl+; (GridCraft's static
  date stamp) writes ММ/ДД/ГГГГ text into such a cell, which the parser also accepts.

Schema only; GridCraft evaluates every formula. Output: runs/STAGE_05/build_plan.jsonl
"""

import json
from pathlib import Path

LAB = Path(__file__).resolve().parents[1]
OUT = LAB / "runs" / "STAGE_05" / "build_plan.jsonl"
SRC_BOOK = str(LAB / "workbooks" / "shared_obligation_tracker.xlsx")
NEW_BOOK = str(LAB / "workbooks" / "shared_obligation_balance.xlsx")

WS, SET = "Учёт", "Настройки"
Q = f"'{SET}'"
FIRST, LAST = 14, 513               # 500 entry rows
EXC = "$B$19:$B$42", "$C$19:$C$42"   # exceptions on «Настройки» (amount, recognised month)

INPUT_FILL, CALC_FILL, HEAD_FILL = "#FFF2CC", "#EDEDED", "#DDEBF7"
MONEY = "#,##0.00"


def V(x):
    return f"IFERROR(VALUE({x}),0)"


def parse_date(ref):
    """дд.мм.гггг, гггг-мм-дд, or ММ/ДД/ГГГГ (GridCraft Ctrl+; in a Text cell) -> date, else "ошибка".
    LET names must not look like cell addresses (p1, m1 ...)."""
    return (f'IF(ISNUMBER({ref}),{ref},LET(t,TRIM({ref}),sl,ISNUMBER(FIND("/",t)),iso,MID(t,5,1)="-",'
            f'pa,IFERROR(FIND("/",t),0),pb,IFERROR(FIND("/",t,pa+1),0),'
            f'dd,IF(sl,{V("MID(t,pa+1,pb-pa-1)")},IF(iso,{V("MID(t,9,2)")},{V("LEFT(t,2)")})),'
            f'mm,IF(sl,{V("LEFT(t,pa-1)")},IF(iso,{V("MID(t,6,2)")},{V("MID(t,4,2)")})),'
            f'yy,IF(sl,{V("MID(t,pb+1,4)")},IF(iso,{V("LEFT(t,4)")},{V("MID(t,7,4)")})),'
            f'ok,IF(sl,AND(pb>0,LEN(t)-pb=4),AND(LEN(t)=10,OR(iso,AND(MID(t,3,1)=".",MID(t,6,1)=".")))),'
            f'dt,IFERROR(DATE(yy,mm,dd),0),'
            f'IF(AND(ok,dd>=1,mm>=1,mm<=12,DAY(dt)=dd,MONTH(dt)=mm),dt,"ошибка")))')


def date_rule(ref):
    # Validation sees typed dates as GridCraft would parse them (numbers), so numbers pass here.
    return (f'=OR({ref}="",ISNUMBER({ref}),AND(LEN({ref})=10,MID({ref},3,1)=".",MID({ref},6,1)="."),'
            f'AND(LEN({ref})=10,MID({ref},5,1)="-",MID({ref},8,1)="-"))')


plan = []


def cmd(command, **params):
    plan.append({"command": command, "params": params})


def fill(rng, formula):
    cmd("selection.set", range=rng, active=rng.split(":")[0])
    cmd("range.fill", range=rng, input=formula)


DATE_MSG = "Дата: дд.мм.гггг (например 15.10.2026) или Ctrl+; — сегодняшняя дата."

# ---------------------------------------------------------------- copy
cmd("file.open", path=SRC_BOOK)
cmd("file.saveAs", path=NEW_BOOK)
cmd("home.deleteSheet", sheet="Поступления")
cmd("home.deleteSheet", sheet="Сводка")
cmd("home.insertSheet", name=WS, before=0)

# ---------------------------------------------------------------- Настройки: rules for the new model
cmd("sheet.activate", sheet=SET)
cmd("range.setValues", range="C8", values=[["платит, сумма не фиксирована"], ["доплачивает остаток общего обязательства"]])
cmd("range.setValues", range="A12", values=[
    ["1. За каждый месяц с начала учёта возникает обязательство (по умолчанию из B5, исключения — ниже)."],
    ["2. Месяц входит в накопленное обязательство, когда «Расчёт на дату» позже этого месяца (оплата за прошлый месяц)."],
    ["3. Все поступления A и B с датой от начала учёта до «Расчёт на дату» вычитаются из накопленного обязательства."],
    ["4. Месяц оплаты не указывается; переплата уменьшает общий остаток."],
    ["5. Ожидается от B = накопленное обязательство − поступления A."],
])
cmd("selection.set", range="B4")

# ---------------------------------------------------------------- Учёт
cmd("sheet.activate", sheet=WS)
cmd("home.numberFormat", range="B4", format="Text")
cmd("home.numberFormat", range=f"A{FIRST}:A{LAST}", format="Text")
cmd("range.setValues", range="A1", values=[
    ["Учёт поступлений: общий баланс"],
    ["Вымышленные данные, EUR. Жёлтое — ввод. Правила и сумма обязательства — на листе «Настройки»."],
])
cmd("range.setValues", range="A4", values=[
    ["Расчёт на дату (дд.мм.гггг)", "15.11.2026", '=IF(ISNUMBER(H4),"","Ошибка: введите дату как дд.мм.гггг")'],
    ["Месяцев к оплате", '=IF(AND(ISNUMBER(H5),ISNUMBER(H6)),MAX(0,(YEAR(H5)-YEAR(H6))*12+MONTH(H5)-MONTH(H6)+1),0)',
     '=IF(B5>0,"с "&TEXT(H6,"mm.yyyy")&" по "&TEXT(H5,"mm.yyyy"),"нет месяцев к оплате")'],
    ["Накопленное обязательство",
     f'=IF(B5=0,0,B5*H7+SUMIFS({Q}!{EXC[0]},{Q}!{EXC[1]},">="&H6,{Q}!{EXC[1]},"<="&H5,{Q}!{EXC[0]},"<>")'
     f'-H7*COUNTIFS({Q}!{EXC[1]},">="&H6,{Q}!{EXC[1]},"<="&H5,{Q}!{EXC[0]},"<>"))', None],
    ["Получено от A", f'=SUMIFS($C${FIRST}:$C${LAST},$B${FIRST}:$B${LAST},{Q}!$A$8,$I${FIRST}:$I${LAST},"учтено")', None],
    ["Получено от B", f'=SUMIFS($C${FIRST}:$C${LAST},$B${FIRST}:$B${LAST},{Q}!$A$9,$I${FIRST}:$I${LAST},"учтено")', None],
    ["Остаток к получению", "=MAX(0,B6-B7-B8)", None],
    ["Переплата", "=MAX(0,B7+B8-B6)", None],
])
cmd("range.setValues", range="D4", values=[
    ["Ожидается от B всего (обязательство − A)", "=MAX(0,B6-B7)"],
    ["Не учтено: дата позже даты расчёта, EUR", f'=SUMIFS($C${FIRST}:$C${LAST},$I${FIRST}:$I${LAST},"Не учтено: дата позже*")'],
    ["Строк с ошибками", f'=COUNTIF($I${FIRST}:$I${LAST},"Ошибка*")'],
    ['=IF(E6>0,"Есть ошибки — см. колонку «Проверка»","")', None],
])
cmd("range.setValues", range="A11", values=[
    ["Новая запись — в первую пустую строку: дата (Ctrl+; — сегодня, или дд.мм.гггг), плательщик A/B, сумма, примечание по желанию."],
])
cmd("range.setValues", range=f"A{FIRST - 1}", values=[["Дата получения", "Плательщик", "Сумма, EUR", "Примечание", "Проверка"]])
cmd("range.setValues", range=f"A{FIRST}", values=[
    ["05.10.2026", "A", 600, None],
    ["12.10.2026", "B", 300, "частичный платёж"],
    ["20.10.2026", "B", 150, None],
    ["03.11.2026", "A", 700, None],
    ["10.11.2026", "B", 420, None],
    ["16.11.2026", "B", 50, "контрольный пример: позже даты расчёта"],
])
# hidden helpers: H4 as-of, H5 last month due, H6 accounting start, H7 default obligation
cmd("range.setValues", range="H3", values=[
    ["служебные"],
    [f'=IF(B4="","",IFERROR({parse_date("B4")},"ошибка"))'],
    ['=IF(ISNUMBER(H4),EDATE(DATE(YEAR(H4),MONTH(H4),1),-1),"")'],
    [f"={Q}!C4"],
    [f"={Q}!B5"],
])
r = FIRST
fill(f"H{r}:H{LAST}", f'=IF(A{r}="","",IFERROR({parse_date(f"A{r}")},"ошибка"))')
status = (
    f'=IF(AND(A{r}="",B{r}="",C{r}=""),"",'
    f'IF(NOT(ISNUMBER(H{r})),"Ошибка: дата не распознана — дд.мм.гггг или Ctrl+;",'
    f'IF(COUNTIF({Q}!$A$8:$A$9,B{r})=0,"Ошибка: выберите плательщика A или B",'
    f'IF(NOT(ISNUMBER(C{r})),"Ошибка: сумма должна быть числом",'
    f'IF(C{r}<=0,"Ошибка: сумма должна быть больше 0",'
    f'IF(H{r}>$H$4,"Не учтено: дата позже даты расчёта",'
    f'IF(H{r}<$H$6,"Не учтено: дата раньше начала учёта","учтено")))))))'
)
fill(f"I{r}:I{LAST}", status)
fill(f"E{r}:E{LAST}", f'=IF(OR(I{r}="",I{r}="учтено"),"",I{r})')

# formatting
cmd("home.bold", range="A1", on=True)
cmd("home.fontSize", range="A1", size=14)
cmd("home.bold", range="A4:A10", on=True)
cmd("home.bold", range="B9:B10", on=True)
cmd("home.bold", range="D4:D6", on=True)
cmd("home.fillColor", range="B4", color=INPUT_FILL)
cmd("home.borders", range="B4", preset="thickOutside")
cmd("home.fillColor", range="B5:B10", color=CALC_FILL)
cmd("home.fillColor", range="E4:E6", color=CALC_FILL)
cmd("home.borders", range="A4:B10", preset="all", style="thin")
cmd("home.borders", range="D4:E6", preset="all", style="thin")
cmd("home.numberFormat", range="B6:B10", code=MONEY)
cmd("home.numberFormat", range="E4:E5", code=MONEY)
cmd("home.fontColor", range="C4", color="#C00000")
cmd("home.fontColor", range="D7", color="#C00000")
cmd("home.bold", range="D7", on=True)
cmd("home.italic", range="A11", on=True)
cmd("home.bold", range=f"A{FIRST - 1}:E{FIRST - 1}", on=True)
cmd("home.fillColor", range=f"A{FIRST - 1}:E{FIRST - 1}", color=HEAD_FILL)
cmd("home.borders", range=f"A{FIRST - 1}:E{FIRST - 1}", preset="all", style="thin")
cmd("home.rowHeight", rows=f"{FIRST - 1}:{FIRST - 1}", height=20)
cmd("home.fillColor", range=f"A{FIRST}:D{LAST}", color=INPUT_FILL)
cmd("home.borders", range=f"A{FIRST}:E{LAST}", preset="all", style="thin")
cmd("home.numberFormat", range=f"C{FIRST}:C{LAST}", code=MONEY)
cmd("home.numberFormat", range="H4:H6", code="dd.mm.yyyy")
cmd("home.numberFormat", range=f"H{FIRST}:H{LAST}", code="dd.mm.yyyy")
cmd("home.alignRight", range=f"A{FIRST}:A{LAST}")
cmd("home.alignCenter", range=f"B{FIRST}:B{LAST}")
cmd("home.fontColor", range=f"E{FIRST}:E{LAST}", color="#C00000")
cmd("home.conditionalFormat", range=f"E{FIRST}:E{LAST}",
    rule={"type": "beginsWith", "text": "Ошибка", "preset": "lightRedFillDarkRedText"})
cmd("home.conditionalFormat", range="B9",
    rule={"type": "expression", "formula": "=B9>0", "preset": "lightRedFillDarkRedText"})
cmd("home.conditionalFormat", range="B10",
    rule={"type": "expression", "formula": "=B10>0", "preset": "greenFill"})
for col, w in (("A", 28), ("B", 14), ("C", 14), ("D", 42), ("E", 50)):
    cmd("home.columnWidth", cols=f"{col}:{col}", chars=w)
cmd("data.validation", range="B4", type="custom", formula1=date_rule("B4"),
    errorStyle="stop", errorTitle="Расчёт на дату", errorMessage="Дата в виде дд.мм.гггг, например 15.11.2026.")
cmd("data.validation", range=f"A{FIRST}:A{LAST}", type="custom", formula1=date_rule(f"A{FIRST}"),
    errorStyle="stop", errorTitle="Дата получения", errorMessage=DATE_MSG)
cmd("data.validation", range=f"B{FIRST}:B{LAST}", type="list", formula1=f"={Q}!$A$8:$A$9", dropdown=True,
    errorStyle="stop", errorTitle="Плательщик", errorMessage="Выберите плательщика из списка (A или B).")
cmd("data.validation", range=f"C{FIRST}:C{LAST}", type="decimal", operator="greater", formula1="0",
    errorStyle="stop", errorTitle="Сумма", errorMessage="Сумма — положительное число.")
cmd("home.hideColumns", cols="H:I")
cmd("view.freezePanes", cell=f"A{FIRST}")
cmd("selection.set", range=f"A{FIRST + 6}")

if __name__ == "__main__":
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        for i, step in enumerate(plan, 1):
            f.write(json.dumps({"step": i, **step}, ensure_ascii=False) + "\n")
    print(f"{len(plan)} steps -> {OUT.relative_to(LAB)}")


# ================================================================ post-review fix (applied after the build)
# Build screenshot: the period text in C5 («с 09.2026 по 10.2026») was clipped by D5.
FIX_OUT = LAB / "runs" / "STAGE_05" / "fix_plan.jsonl"
fix = [{"command": "sheet.activate", "params": {"sheet": WS}},
       {"command": "home.columnWidth", "params": {"cols": "C:C", "chars": 21}},
       {"command": "selection.set", "params": {"range": f"A{FIRST + 6}"}}]

if __name__ == "__main__":
    with open(FIX_OUT, "w", encoding="utf-8") as f:
        for i, step in enumerate(fix, 1):
            f.write(json.dumps({"step": i, **step}, ensure_ascii=False) + "\n")
    print(f"{len(fix)} fix steps -> {FIX_OUT.relative_to(LAB)}")
