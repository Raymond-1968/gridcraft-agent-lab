"""Stage 05 check plan with independently computed expectations (external arithmetic).

Model: accumulated obligation over months due (start .. month before as-of) with exceptions;
every A/B payment dated from start to as-of is subtracted; remainder / overpayment overall.
Output: runs/STAGE_05/checks_plan.jsonl, expected_model.json, snapshot_cells.json
"""

import datetime as dt
import json
from pathlib import Path

LAB = Path(__file__).resolve().parents[1]
RUN = LAB / "runs" / "STAGE_05"
BOOK = str(LAB / "workbooks" / "shared_obligation_balance.xlsx")
D = dt.date
START = D(2026, 9, 1)
DEFAULT = 1100.0
WS = "Учёт"


def serial(d):
    return float((d - D(1899, 12, 30)).days)


def parse_date(t):
    t = (t or "").strip()
    try:
        if "/" in t:
            m, d, y = t.split("/")
            return D(int(y), int(m), int(d)) if len(y) == 4 else None
        if len(t) == 10 and t[4] == "-" and t[7] == "-":
            return D(int(t[:4]), int(t[5:7]), int(t[8:]))
        if len(t) == 10 and t[2] == "." and t[5] == ".":
            return D(int(t[6:]), int(t[3:5]), int(t[:2]))
    except ValueError:
        return None
    return None


def status(date_t, payer, amount, asof):
    if not any([date_t, payer, amount not in (None, "")]):
        return ""
    d = parse_date(date_t)
    if d is None:
        return "Ошибка: дата не распознана — дд.мм.гггг или Ctrl+;"
    if payer not in ("A", "B"):
        return "Ошибка: выберите плательщика A или B"
    if not isinstance(amount, (int, float)):
        return "Ошибка: сумма должна быть числом"
    if amount <= 0:
        return "Ошибка: сумма должна быть больше 0"
    if d > asof:
        return "Не учтено: дата позже даты расчёта"
    if d < START:
        return "Не учтено: дата раньше начала учёта"
    return "учтено"


def model(asof_t, payments, exceptions=None):
    exceptions = exceptions or {}
    asof = parse_date(asof_t)
    last = D(asof.year, asof.month, 1)
    last = D(last.year - (last.month == 1), (last.month - 2) % 12 + 1, 1)
    n = max(0, (last.year - START.year) * 12 + last.month - START.month + 1)
    months = [D(START.year + (START.month - 1 + k) // 12, (START.month - 1 + k) % 12 + 1, 1) for k in range(n)]
    obligation = sum(exceptions.get(m, DEFAULT) for m in months)
    st = [status(*p, asof) for p in payments]
    a = sum(p[2] for p, s in zip(payments, st) if s == "учтено" and p[1] == "A")
    b = sum(p[2] for p, s in zip(payments, st) if s == "учтено" and p[1] == "B")
    late = sum(p[2] for p, s in zip(payments, st) if s.startswith("Не учтено: дата позже"))
    return {"months": float(n), "obligation": obligation, "A": a, "B": b,
            "remaining": max(0.0, obligation - a - b), "over": max(0.0, a + b - obligation),
            "expectedB": max(0.0, obligation - a), "late": late,
            "errors": float(sum(s.startswith("Ошибка") for s in st)),
            "check_column": ["" if s in ("", "учтено") else s for s in st]}


BASE = [("05.10.2026", "A", 600.0), ("12.10.2026", "B", 300.0), ("20.10.2026", "B", 150.0),
        ("03.11.2026", "A", 700.0), ("10.11.2026", "B", 420.0), ("16.11.2026", "B", 50.0)]
ASOF = "15.11.2026"
plan, expected = [], {}


def act(check, title, tool, **arguments):
    plan.append({"check": check, "title": title, "tool": tool, "arguments": arguments})


def expect(check, title, sheet, rng, values):
    plan.append({"check": check, "title": title, "tool": "read_range",
                 "arguments": {"sheet": sheet, "range": rng}, "expect": values})


def set_cell(check, title, cell, value, sheet=WS):
    act(check, title, "execute_command", command="cell.set", params={"sheet": sheet, "cell": cell, "input": value})


def set_row(check, title, row, values):
    act(check, title, "execute_command", command="range.setValues",
        params={"range": f"{WS}!A{row}", "values": [list(values)]})


def clear(check, title, rng):
    act(check, title, "execute_command", command="edit.clearContents", params={"range": rng})


def expect_all(check, asof, payments, rows=None, **kw):
    r = model(asof, payments, **kw)
    expected[check] = r
    expect(check, "summary B5:B10 (months, obligation, A, B, remaining, overpayment)", WS, "B5:B10",
           [[r["months"]], [r["obligation"]], [r["A"]], [r["B"]], [r["remaining"]], [r["over"]]])
    expect(check, "E4:E6 (expected from B, late payments, error rows)", WS, "E4:E6",
           [[r["expectedB"]], [r["late"]], [r["errors"]]])
    rows = rows or len(payments)
    expect(check, f"«Проверка» column E14:E{13 + rows}", WS, f"E14:E{13 + rows}",
           [[c] for c in (r["check_column"] + [""] * rows)[:rows]])
    return r


expect_all("L0", ASOF, BASE, rows=8)
expect("L0", "as-of error message C4 is blank", WS, "C4", [[""]])
expect("L0", "months text C5", WS, "C5", [["с 09.2026 по 10.2026"]])

set_row("L1", "add 14.11.2026 A 500 (overpayment reduces the balance)", 20, ("14.11.2026", "A", 500))
expect_all("L1", ASOF, BASE + [("14.11.2026", "A", 500.0)])
clear("L1r", "clear row 20", f"{WS}!A20:D20")

set_cell("L2", "as-of 30.11.2026", "B4", "30.11.2026")
expect_all("L2", "30.11.2026", BASE)
set_cell("L2b", "as-of 01.12.2026 (November due)", "B4", "01.12.2026")
expect_all("L2b", "01.12.2026", BASE)
set_cell("L2r", "revert as-of 15.11.2026", "B4", "15.11.2026")
expect_all("L2r", ASOF, BASE)

act("L3", "exception 10.2026 = 1000", "execute_command", command="range.setValues",
    params={"range": "Настройки!A19", "values": [["10.2026", 1000]]})
expect_all("L3", ASOF, BASE, exceptions={D(2026, 10, 1): 1000.0})
clear("L3r", "clear exception", "Настройки!A19:B19")
expect_all("L3r", ASOF, BASE)

for cid, title, row in (
        ("L4a", "ISO date 2026-11-01 B 10", ("2026-11-01", "B", 10)),
        ("L4b", "Ctrl+; text form 11/01/2026 B 10", ("11/01/2026", "B", 10)),
        ("L4c", "impossible date 31.02.2026", ("31.02.2026", "B", 10)),
        ("L4d", "date before accounting start 15.08.2026", ("15.08.2026", "B", 10)),
        ("L4e", "payer C", ("01.11.2026", "C", 10)),
        ("L4f", "amount 'abc'", ("01.11.2026", "B", "abc")),
        ("L4g", "amount -5", ("01.11.2026", "B", -5))):
    set_row(cid, title, 20, row)
    amount = row[2] if isinstance(row[2], str) else float(row[2])
    expect_all(cid, ASOF, BASE + [(row[0], row[1], amount)])
clear("L4r", "clear row 20", f"{WS}!A20:D20")
expect_all("L4r", ASOF, BASE)

expect("L5", "empty rows: «Проверка» E20:E22 blank", WS, "E20:E22", [[""], [""], [""]])
for sheet, rng in ((WS, "A1:I513"), ("Настройки", "A4:D42")):
    plan.append({"check": "L5", "title": f"no error values in {sheet}!{rng}", "tool": "read_range",
                 "arguments": {"sheet": sheet, "range": rng}, "expect_no_errors": True})

set_row("L6", "row 513 (500th): 01.11.2026 A 5", 513, ("01.11.2026", "A", 5))
expect_all("L6", ASOF, BASE + [("01.11.2026", "A", 5.0)])
expect("L6", "«Проверка» E513 blank (counted)", WS, "E513", [[""]])
clear("L6r", "clear row 513", f"{WS}!A513:D513")
expect_all("L6r", ASOF, BASE)

act("L7", f"activate {WS}", "execute_command", command="sheet.activate", params={"sheet": WS})
for cell, value, ok in (("A14", "15.10.2026", True), ("A14", "2026-10-15", True), ("A14", "10/12/2026", True),
                        ("A14", "15/10/2026", False), ("A513", "01.11.2026", True), ("B14", "A", True),
                        ("B14", "C", False), ("C14", "-5", False), ("B4", "15.11.2026", True),
                        ("B4", "15/11/2026", False)):
    plan.append({"check": "L7", "title": f"data.validate {cell}={value!r} expected ok={ok}", "tool": "execute_command",
                 "arguments": {"command": "data.validate", "params": {"cell": cell, "input": value}}, "expect_ok": ok})

act("L8", "select first empty entry row", "execute_command", command="selection.set", params={"range": "A20"})
act("L8", "save", "execute_command", command="file.save", params={"path": BOOK})

Y, G, H = 0xFFF2CC, 0xEDEDED, 0xDDEBF7
SNAPSHOT = [
    {"sheet": WS, "cell": "B4", "style": {"fill": Y, "num_fmt": "@"}},
    {"sheet": WS, "cell": "B6", "style": {"fill": G, "num_fmt": "#,##0.00"}},
    {"sheet": WS, "cell": "B7", "style": {"fill": G, "num_fmt": "#,##0.00"}},
    {"sheet": WS, "cell": "B9", "style": {"fill": G, "num_fmt": "#,##0.00", "bold": True}},
    {"sheet": WS, "cell": "B10", "style": {"fill": G, "num_fmt": "#,##0.00", "bold": True}},
    {"sheet": WS, "cell": "E4", "style": {"fill": G, "num_fmt": "#,##0.00"}},
    {"sheet": WS, "cell": "E6", "style": {"fill": G}},
    {"sheet": WS, "cell": "A13", "style": {"fill": H, "bold": True}},
    {"sheet": WS, "cell": "A14", "style": {"fill": Y, "num_fmt": "@"}},
    {"sheet": WS, "cell": "C14", "style": {"fill": Y, "num_fmt": "#,##0.00"}},
    {"sheet": WS, "cell": "E19", "style": {}},
    {"sheet": WS, "cell": "H4", "style": {"num_fmt": "dd.mm.yyyy"}},
    {"sheet": WS, "cell": "H14", "style": {"num_fmt": "dd.mm.yyyy"}},
    {"sheet": WS, "cell": "I14", "style": {}},
    {"sheet": WS, "cell": "A513", "style": {"fill": Y, "num_fmt": "@"}},
    {"sheet": WS, "cell": "I513", "style": {}},
    {"sheet": "Настройки", "cell": "B4", "style": {"fill": Y, "num_fmt": "@"}},
    {"sheet": "Настройки", "cell": "C4", "style": {"fill": G, "num_fmt": "mm.yyyy"}},
    {"sheet": "Настройки", "cell": "B5", "style": {"fill": Y, "num_fmt": "#,##0.00"}},
    {"sheet": "Настройки", "cell": "A12", "style": {}},
    {"sheet": "Настройки", "cell": "C19", "style": {"fill": G, "num_fmt": "mm.yyyy"}},
]

if __name__ == "__main__":
    RUN.mkdir(parents=True, exist_ok=True)
    with open(RUN / "checks_plan.jsonl", "w", encoding="utf-8") as f:
        for i, step in enumerate(plan, 1):
            f.write(json.dumps({"step": i, **step}, ensure_ascii=False) + "\n")
    (RUN / "expected_model.json").write_text(json.dumps(expected, ensure_ascii=False, indent=1), encoding="utf-8")
    (RUN / "snapshot_cells.json").write_text(json.dumps(SNAPSHOT, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{len(plan)} steps; L0 {dict((k, v) for k, v in expected['L0'].items() if k != 'check_column')}")
