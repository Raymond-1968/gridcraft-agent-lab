"""Stage 04 check plan with independently computed expectations (external arithmetic).

Model (user clarification 2026-10-13): one obligation per month (default 1100, per-month
exceptions); months due = accounting start .. month before the as-of date; a payment counts for
the month it names if its date <= as-of; expected B = max(0, obligation - A); remainder
= obligation - A - B per month; overpayments offset debts in the overall saldo.

Output: runs/STAGE_04/checks_plan.jsonl, expected_model.json, snapshot_cells.json
"""

import datetime as dt
import json
from pathlib import Path

LAB = Path(__file__).resolve().parents[1]
RUN = LAB / "runs" / "STAGE_04"
BOOK = str(LAB / "workbooks" / "shared_obligation_tracker.xlsx")
D = dt.date


def month_add(m, k):
    y, mo = divmod(m.month - 1 + k, 12)
    return D(m.year + y, mo + 1, 1)


def parse_date(t):
    t = (t or "").strip()
    try:
        if len(t) == 10 and t[2] == "." and t[5] == ".":
            return D(int(t[6:]), int(t[3:5]), int(t[:2]))
        if len(t) == 10 and t[4] == "-" and t[7] == "-":
            return D(int(t[:4]), int(t[5:7]), int(t[8:]))
    except ValueError:
        return None
    return None


def parse_month(t):
    t = (t or "").strip()
    try:
        if len(t) == 7 and t[2] == "." and 1 <= int(t[:2]) <= 12:
            return D(int(t[3:]), int(t[:2]), 1)
    except ValueError:
        return None
    return None


def model(asof_text, payments, start="09.2026", default=1100.0, exceptions=None):
    exceptions = exceptions or {}
    asof, start_m = parse_date(asof_text), parse_month(start)
    last_due = month_add(D(asof.year, asof.month, 1), -1)
    months = []
    m = start_m
    while m <= last_due and len(months) < 60:
        months.append(m)
        m = month_add(m, 1)
    statuses, rec = [], {}
    for date_t, payer, amount, month_t in payments:
        if not any([date_t, payer, amount, month_t]):
            statuses.append("")
            continue
        d, pm = parse_date(date_t), parse_month(month_t)
        if d is None:
            s = "ошибка: дата"
        elif payer not in ("A", "B"):
            s = "ошибка: плательщик"
        elif not isinstance(amount, (int, float)):
            s = "ошибка: сумма"
        elif pm is None:
            s = "ошибка: месяц оплаты"
        elif d > asof:
            s = "после даты расчёта"
        elif pm < start_m:
            s = "месяц до начала учёта"
        elif pm > last_due:
            s = "месяц ещё не к оплате"
        else:
            s = "учтено"
            rec[(pm, payer)] = rec.get((pm, payer), 0.0) + amount
        statuses.append(s)
    rows, saldo = [], 0.0
    for m in months:
        obl = exceptions.get(m, default)
        a, b = rec.get((m, "A"), 0.0), rec.get((m, "B"), 0.0)
        saldo += obl - a - b
        rows.append({"month": m.isoformat(), "obligation": obl, "A": a, "expB": max(0.0, obl - a), "B": b,
                     "debt": max(0.0, obl - a - b), "over": max(0.0, a + b - obl), "saldo": saldo})
    tot_obl = sum(r["obligation"] for r in rows)
    tot_rec = sum(r["A"] + r["B"] for r in rows)
    prepay = sum(p[2] for p, s in zip(payments, statuses) if s == "месяц ещё не к оплате")
    return {"rows": rows, "statuses": statuses,
            "totals": {"obligations": tot_obl, "received": tot_rec, "saldo": tot_obl - tot_rec,
                       "debt": max(0.0, tot_obl - tot_rec), "over": max(0.0, tot_rec - tot_obl),
                       "prepay": prepay},
            "counted": statuses.count("учтено"), "errors": sum(s.startswith("ошибка") for s in statuses),
            "last_due": last_due.isoformat()}


BASE = [("05.10.2026", "A", 600.0, "09.2026"), ("12.10.2026", "B", 300.0, "09.2026"),
        ("20.10.2026", "B", 150.0, "09.2026"), ("03.11.2026", "A", 700.0, "10.2026"),
        ("10.11.2026", "B", 420.0, "10.2026"), ("14.11.2026", "A", 500.0, "11.2026"),
        ("16.11.2026", "B", 50.0, "09.2026")]
ASOF = "15.11.2026"
EMPTY = ("", "", "", "")

plan, expected = [], {}


def act(check, title, tool, **arguments):
    plan.append({"check": check, "title": title, "tool": tool, "arguments": arguments})


def expect(check, title, sheet, rng, values):
    plan.append({"check": check, "title": title, "tool": "read_range",
                 "arguments": {"sheet": sheet, "range": rng}, "expect": values})


def set_cell(check, title, sheet, cell, value):
    act(check, title, "execute_command", command="cell.set", params={"sheet": sheet, "cell": cell, "input": value})


def set_row(check, title, row, values):
    act(check, title, "execute_command", command="range.setValues",
        params={"range": f"Поступления!A{row}", "values": [list(values)]})


def clear(check, title, rng):
    act(check, title, "execute_command", command="edit.clearContents", params={"range": rng})


def expect_all(check, asof, payments, rows_to_show=3, **kw):
    r = model(asof, payments, **kw)
    expected[check] = r
    t = r["totals"]
    expect(check, "totals F4:F9 (obligations, received, saldo, debt, overpayment, prepayment)", "Сводка", "F4:F9",
           [[t["obligations"]], [t["received"]], [t["saldo"]], [t["debt"]], [t["over"]], [t["prepay"]]])
    expect(check, "counted / error rows B8:B9", "Сводка", "B8:B9", [[float(r["counted"])], [float(r["errors"])]])
    vals = []
    for i in range(rows_to_show):
        if i < len(r["rows"]):
            x = r["rows"][i]
            vals.append([x["obligation"], x["A"], x["expB"], x["B"], x["debt"], x["over"], x["saldo"]])
        else:
            vals.append([""] * 7)
    expect(check, f"monthly rows 12..{11 + rows_to_show} B:H", "Сводка", f"B12:H{11 + rows_to_show}", vals)
    return r


# K0 control example
r0 = expect_all("K0", ASOF, BASE)
expect("K0", "last month due B5 (10.2026) as serial", "Сводка", "B5", [[float((D(2026, 10, 1) - D(1899, 12, 30)).days)]])
expect("K0", "statuses rows 6-12", "Поступления", "H6:H12", [[s] for s in r0["statuses"]])

# K1 partial payments sum, overpayment offsets
set_row("K1", "add 12.11.2026 B 30 for 10.2026 in row 13", 13, ("12.11.2026", "B", 30, "10.2026"))
expect_all("K1", ASOF, BASE + [("12.11.2026", "B", 30.0, "10.2026")])
clear("K1r", "clear row 13", "Поступления!A13:E13")

# K2 expected B follows A
set_cell("K2", "A payment for 09.2026: 600 -> 800", "Поступления", "C6", "800")
expect_all("K2", ASOF, [("05.10.2026", "A", 800.0, "09.2026")] + BASE[1:])
set_cell("K2r", "revert to 600", "Поступления", "C6", "600")
expect_all("K2r", ASOF, BASE)

# K3 / K4 as-of date moves
set_cell("K3", "as-of 30.11.2026", "Сводка", "B4", "30.11.2026")
expect_all("K3", "30.11.2026", BASE)
set_cell("K4", "as-of 01.12.2026 (November becomes due)", "Сводка", "B4", "01.12.2026")
expect_all("K4", "01.12.2026", BASE)
set_cell("K4r", "revert as-of 15.11.2026", "Сводка", "B4", "15.11.2026")
expect_all("K4r", ASOF, BASE)

# K5 per-month obligation exception
act("K5", "exception 10.2026 = 1000 on Настройки", "execute_command", command="range.setValues",
    params={"range": "Настройки!A19", "values": [["10.2026", 1000]]})
expect_all("K5", ASOF, BASE, exceptions={D(2026, 10, 1): 1000.0})
clear("K5r", "clear exception", "Настройки!A19:B19")
expect_all("K5r", ASOF, BASE)

# K6 A pays more than the obligation
set_row("K6", "add 04.11.2026 A 600 for 10.2026", 13, ("04.11.2026", "A", 600, "10.2026"))
expect_all("K6", ASOF, BASE + [("04.11.2026", "A", 600.0, "10.2026")])
clear("K6r", "clear row 13", "Поступления!A13:E13")

# K7 date formats and invalid input in rows
set_row("K7", "ISO date 2026-11-01 B 10 for 09.2026", 13, ("2026-11-01", "B", 10, "09.2026"))
expect_all("K7", ASOF, BASE + [("2026-11-01", "B", 10.0, "09.2026")])
expect("K7", "recognised date F13 = 01.11.2026", "Поступления", "F13", [[float((D(2026, 11, 1) - D(1899, 12, 30)).days)]])
set_cell("K7b", "impossible date 31.02.2026", "Поступления", "A13", "31.02.2026")
expect_all("K7b", ASOF, BASE + [("31.02.2026", "B", 10.0, "09.2026")])
expect("K7b", "status row 13", "Поступления", "H13", [["ошибка: дата"]])
set_cell("K7c", "valid date, payment month missing", "Поступления", "A13", "01.11.2026")
clear("K7c", "clear month D13", "Поступления!D13")
expect_all("K7c", ASOF, BASE + [("01.11.2026", "B", 10.0, "")])
expect("K7c", "status row 13", "Поступления", "H13", [["ошибка: месяц оплаты"]])
# K7d: inputs that pass validation (GridCraft reads them as numbers) are caught by the row status
set_row("K7d", "date 1.9.26 (passes validation as a number, stored as text)", 13, ("1.9.26", "B", 10, "09.2026"))
expect("K7d", "status row 13", "Поступления", "H13", [["ошибка: дата"]])
set_row("K7d", "month 10/2026 (passes validation as a number, stored as text)", 13, ("01.11.2026", "B", 10, "10/2026"))
expect("K7d", "status row 13", "Поступления", "H13", [["ошибка: месяц оплаты"]])
expect_all("K7d", ASOF, BASE + [("01.11.2026", "B", 10.0, "10/2026")])
clear("K7r", "clear row 13", "Поступления!A13:E13")

# K8 month before accounting start, K9 unknown payer
set_row("K8", "01.11.2026 A 100 for 08.2026", 13, ("01.11.2026", "A", 100, "08.2026"))
expect_all("K8", ASOF, BASE + [("01.11.2026", "A", 100.0, "08.2026")])
expect("K8", "status row 13", "Поступления", "H13", [["месяц до начала учёта"]])
set_cell("K9", "payer C (programmatic entry bypasses the list)", "Поступления", "B13", "C")
expect_all("K9", ASOF, BASE + [("01.11.2026", "C", 100.0, "08.2026")])
expect("K9", "status row 13", "Поступления", "H13", [["ошибка: плательщик"]])
clear("K9r", "clear row 13", "Поступления!A13:E13")
expect_all("K9r", ASOF, BASE)

# K10 empty rows and error scan
expect("K10", "statuses of empty rows 13-15 are blank", "Поступления", "H13:H15", [[""], [""], [""]])
expect("K10", "months after the last due month are blank", "Сводка", "A14:A16", [[""], [""], [""]])
for sheet, rng in (("Сводка", "A4:H72"), ("Поступления", "F6:H505"), ("Настройки", "A4:D42")):
    plan.append({"check": "K10", "title": f"no error values in {sheet}!{rng}", "tool": "read_range",
                 "arguments": {"sheet": sheet, "range": rng}, "expect_no_errors": True})

# K11 500th payment row
set_row("K11", "row 505 (500th): 01.11.2026 B 50 for 09.2026", 505, ("01.11.2026", "B", 50, "09.2026"))
expect_all("K11", ASOF, BASE + [("01.11.2026", "B", 50.0, "09.2026")])
expect("K11", "status row 505", "Поступления", "H505", [["учтено"]])
clear("K11r", "clear row 505", "Поступления!A505:E505")
expect_all("K11r", ASOF, BASE)

# K12 validation rules
for sheet, cells in (("Поступления", [("A6", "15.10.2026", True), ("A6", "2026-10-15", True), ("A6", "1.9.26", True),
                                       ("A6", "15/10/2026", False), ("B6", "A", True), ("B6", "C", False),
                                       ("C6", "-5", False), ("D6", "10.2026", True), ("D6", "10/2026", True),
                                       ("D6", "2026-10", False), ("A505", "01.11.2026", True), ("A505", "1/11/26", True)]),
                     ("Сводка", [("B4", "15.11.2026", True), ("B4", "15/11/2026", False)]),
                     ("Настройки", [("B4", "09.2026", True), ("A19", "10.2026", True), ("A19", "окт", False),
                                    ("B5", "-1", False)])):
    act("K12", f"activate {sheet}", "execute_command", command="sheet.activate", params={"sheet": sheet})
    for cell, value, ok in cells:
        plan.append({"check": "K12", "title": f"data.validate {sheet}!{cell}={value!r} expected ok={ok}",
                     "tool": "execute_command",
                     "arguments": {"command": "data.validate", "params": {"cell": cell, "input": value}},
                     "expect_ok": ok})

# K13 save
act("K13", "activate Сводка", "execute_command", command="sheet.activate", params={"sheet": "Сводка"})
act("K13", "select B4", "execute_command", command="selection.set", params={"range": "B4"})
act("K13", "save", "execute_command", command="file.save", params={"path": BOOK})

Y, G, H = 0xFFF2CC, 0xEDEDED, 0xDDEBF7
MONEY, DATE_F, MONTH_F, TEXT = "#,##0.00", "dd.mm.yyyy", "mm.yyyy", "@"
SNAPSHOT = [
    {"sheet": "Сводка", "cell": "B4", "style": {"fill": Y, "num_fmt": TEXT}},
    {"sheet": "Сводка", "cell": "C4", "style": {"fill": G, "num_fmt": DATE_F}},
    {"sheet": "Сводка", "cell": "B5", "style": {"fill": G, "num_fmt": MONTH_F}},
    {"sheet": "Сводка", "cell": "F6", "style": {"fill": G, "num_fmt": MONEY}},
    {"sheet": "Сводка", "cell": "F7", "style": {"fill": G, "num_fmt": MONEY, "bold": True}},
    {"sheet": "Сводка", "cell": "F9", "style": {"fill": G, "num_fmt": MONEY}},
    {"sheet": "Сводка", "cell": "A11", "style": {"fill": H, "bold": True}},
    {"sheet": "Сводка", "cell": "A12", "style": {"fill": G, "num_fmt": MONTH_F}},
    {"sheet": "Сводка", "cell": "A13", "style": {"fill": G, "num_fmt": MONTH_F}},
    {"sheet": "Сводка", "cell": "B12", "style": {"fill": G, "num_fmt": MONEY}},
    {"sheet": "Сводка", "cell": "C12", "style": {"fill": G, "num_fmt": MONEY}},
    {"sheet": "Сводка", "cell": "D12", "style": {"fill": G, "num_fmt": MONEY}},
    {"sheet": "Сводка", "cell": "E13", "style": {"fill": G, "num_fmt": MONEY}},
    {"sheet": "Сводка", "cell": "F12", "style": {"fill": G, "num_fmt": MONEY}},
    {"sheet": "Сводка", "cell": "G13", "style": {"fill": G, "num_fmt": MONEY}},
    {"sheet": "Сводка", "cell": "H13", "style": {"fill": G, "num_fmt": MONEY}},
    {"sheet": "Сводка", "cell": "H72", "style": {"bold": True, "num_fmt": MONEY}},
    {"sheet": "Настройки", "cell": "B4", "style": {"fill": Y, "num_fmt": TEXT}},
    {"sheet": "Настройки", "cell": "C4", "style": {"fill": G, "num_fmt": MONTH_F}},
    {"sheet": "Настройки", "cell": "B5", "style": {"fill": Y, "num_fmt": MONEY}},
    {"sheet": "Настройки", "cell": "A18", "style": {"fill": H, "bold": True}},
    {"sheet": "Настройки", "cell": "A19", "style": {"fill": Y, "num_fmt": TEXT}},
    {"sheet": "Настройки", "cell": "C19", "style": {"fill": G, "num_fmt": MONTH_F}},
    {"sheet": "Поступления", "cell": "A5", "style": {"fill": H, "bold": True}},
    {"sheet": "Поступления", "cell": "A6", "style": {"fill": Y, "num_fmt": TEXT}},
    {"sheet": "Поступления", "cell": "C6", "style": {"fill": Y, "num_fmt": MONEY}},
    {"sheet": "Поступления", "cell": "D6", "style": {"fill": Y, "num_fmt": TEXT}},
    {"sheet": "Поступления", "cell": "F6", "style": {"fill": G, "num_fmt": DATE_F}},
    {"sheet": "Поступления", "cell": "G6", "style": {"fill": G, "num_fmt": MONTH_F}},
    {"sheet": "Поступления", "cell": "H6", "style": {"fill": G}},
    {"sheet": "Поступления", "cell": "A505", "style": {"fill": Y, "num_fmt": TEXT}},
    {"sheet": "Поступления", "cell": "H505", "style": {"fill": G}},
]

if __name__ == "__main__":
    RUN.mkdir(parents=True, exist_ok=True)
    with open(RUN / "checks_plan.jsonl", "w", encoding="utf-8") as f:
        for i, step in enumerate(plan, 1):
            f.write(json.dumps({"step": i, **step}, ensure_ascii=False) + "\n")
    (RUN / "expected_model.json").write_text(json.dumps(expected, ensure_ascii=False, indent=1, default=str),
                                             encoding="utf-8")
    (RUN / "snapshot_cells.json").write_text(json.dumps(SNAPSHOT, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{len(plan)} steps; K0 {expected['K0']['totals']}; K4 {expected['K4']['totals']}")
