"""Stage 03 check plan with independently computed expectations.

The expected numbers come from a small external arithmetic model below (plain Python,
not GridCraft). The workbook's own results are produced by GridCraft and read back
through MCP during the check session; this file only states what they should be.

Output: runs/STAGE_03/checks_plan.jsonl, runs/STAGE_03/expected_model.json
Each line: {"check": id, "title": ..., "tool": <MCP tool>, "arguments": {...}, "expect"?: [[...]]}
"""

import datetime as dt
import json
from pathlib import Path

LAB = Path(__file__).resolve().parents[1]
RUN = LAB / "runs" / "STAGE_03"
BOOK = str(LAB / "workbooks" / "income_tracker_prototype.xlsx")

D = dt.date
ACC_START = D(2026, 9, 1)


def months(first, last, asof):
    """Whole-month accrual on the 1st, from max(accounting start, first) to min(asof month, last)."""
    if first is None:
        return 0
    m0 = max(D(ACC_START.year, ACC_START.month, 1), D(first.year, first.month, 1))
    m1 = D(asof.year, asof.month, 1)
    if last is not None:
        m1 = min(m1, D(last.year, last.month, 1))
    return max(0, (m1.year - m0.year) * 12 + m1.month - m0.month + 1)


def model(sources, payments, asof):
    rows, debt, adv = {}, 0.0, 0.0
    for sid, monthly, first, last, initial in sources:
        acc = months(first, last, asof) * monthly
        rec = sum(a for d, s, a in payments if s == sid and ACC_START <= d <= asof)
        net = initial + acc - rec
        rows[sid] = {"initial": initial, "accrued": acc, "received": rec,
                     "debt": max(0.0, net), "advance": max(0.0, -net), "months": months(first, last, asof)}
        debt += max(0.0, net)
        adv += max(0.0, -net)
    return rows, {"debt": debt, "advance": adv, "net": debt - adv}


BASE_SOURCES = [("A", 300.0, D(2026, 9, 1), None, 100.0), ("B", 200.0, D(2026, 10, 1), None, 0.0)]
BASE_PAYMENTS = [(D(2026, 9, 15), "A", 150.0), (D(2026, 10, 10), "A", 200.0),
                 (D(2026, 10, 20), "B", 250.0), (D(2026, 11, 5), "A", 100.0)]
OCT31, NOV30, DEC31 = D(2026, 10, 31), D(2026, 11, 30), D(2026, 12, 31)
SRC_C = ("C", 150.0, D(2026, 12, 1), None, 0.0)          # future start
SRC_D = ("D", 100.0, D(2026, 8, 1), D(2026, 9, 1), 0.0)  # started before accounting, last month Sep

plan, expected = [], {}


def act(check, title, tool, **arguments):
    plan.append({"check": check, "title": title, "tool": tool, "arguments": arguments})


def expect(check, title, sheet, rng, values):
    plan.append({"check": check, "title": title, "tool": "read_range",
                 "arguments": {"sheet": sheet, "range": rng}, "expect": values})


def expect_summary(check, sources, payments, asof, ids):
    rows, tot = model(sources, payments, asof)
    expected[check] = {"asof": asof.isoformat(), "rows": rows, "totals": tot}
    expect(check, "totals: debt, advance, net", "Сводка", "B8:B10",
           [[tot["debt"]], [tot["advance"]], [tot["net"]]])
    for i, sid in enumerate(ids):
        r = rows[sid]
        row = 13 + i
        expect(check, f"source {sid}: initial, accrued, received, debt, advance, months", "Сводка",
               f"C{row}:H{row}",
               [[r["initial"], r["accrued"], r["received"], r["debt"], r["advance"], float(r["months"])]])


def set_cell(sheet, cell, value, check, title):
    act(check, title, "execute_command", command="cell.set", params={"sheet": sheet, "cell": cell, "input": value})


# C0 control example
expect_summary("C0", BASE_SOURCES, BASE_PAYMENTS, OCT31, ["A", "B"])
expect("C0", "payment statuses rows 6-9", "Поступления", "F6:F9",
       [["учтено"], ["учтено"], ["учтено"], ["после даты расчёта"]])

# C1 change A payment 200 -> 220
set_cell("Поступления", "C7", "220", "C1", "A payment 10.10 200 -> 220")
p = [(d, s, 220.0 if (d, s) == (D(2026, 10, 10), "A") else a) for d, s, a in BASE_PAYMENTS]
expect_summary("C1", BASE_SOURCES, p, OCT31, ["A", "B"])

# C2 revert 200
set_cell("Поступления", "C7", "200", "C2", "revert A payment to 200")
expect_summary("C2", BASE_SOURCES, BASE_PAYMENTS, OCT31, ["A", "B"])

# C3 as-of 30.11.2026
set_cell("Сводка", "B4", "2026-11-30", "C3", "as-of date 2026-11-30")
expect_summary("C3", BASE_SOURCES, BASE_PAYMENTS, NOV30, ["A", "B"])

# C4 revert as-of
set_cell("Сводка", "B4", "2026-10-31", "C4", "revert as-of date 2026-10-31")
expect_summary("C4", BASE_SOURCES, BASE_PAYMENTS, OCT31, ["A", "B"])

# C5 future-start source C, C6 source D with last month
act("C5", "add source C (future start 2026-12) and D (2026-08..2026-09)", "execute_command",
    command="range.setValues", params={"range": "Источники!A10", "values": [
        ["C", "Источник C (будущий старт)", 150, "2026-12-01", None, 0, "проверка C5"],
        ["D", "Источник D (окончен)", 100, "2026-08-01", "2026-09-01", 0, "проверка C6"]]})
srcs = BASE_SOURCES + [SRC_C, SRC_D]
expect_summary("C5", srcs, BASE_PAYMENTS, OCT31, ["A", "B", "C", "D"])
set_cell("Сводка", "B4", "2026-12-31", "C6", "as-of 2026-12-31: C starts, D stays at 1 month")
expect_summary("C6", srcs, BASE_PAYMENTS, DEC31, ["A", "B", "C", "D"])
set_cell("Сводка", "B4", "2026-10-31", "C6r", "revert as-of 2026-10-31")
act("C6r", "remove test sources C, D", "execute_command",
    command="edit.clearContents", params={"range": "Источники!A10:G11"})
expect_summary("C6r", BASE_SOURCES, BASE_PAYMENTS, OCT31, ["A", "B"])

# C7 empty rows: no errors anywhere
expect("C7", "error-row counter", "Сводка", "B7", [[0.0]])
expect("C7", "statuses of empty rows 10-12 are blank", "Поступления", "F10:F12", [[""], [""], [""]])
plan.append({"check": "C7", "title": "no error values on Сводка A4:H33 and Поступления F6:F505",
             "tool": "read_range", "arguments": {"sheet": "Сводка", "range": "A4:H33"},
             "expect_no_errors": True})
plan.append({"check": "C7", "title": "no error values in status column", "tool": "read_range",
             "arguments": {"sheet": "Поступления", "range": "F6:F505"}, "expect_no_errors": True})

# C8 payment in the 500th row (row 505)
act("C8", "payment in row 505 (500th row): 2026-10-25 B 10", "execute_command",
    command="range.setValues", params={"range": "Поступления!A505",
                                        "values": [["2026-10-25", "B", 10, "октябрь 2026", "проверка строки 500"]]})
p8 = BASE_PAYMENTS + [(D(2026, 10, 25), "B", 10.0)]
expect_summary("C8", BASE_SOURCES, p8, OCT31, ["A", "B"])
expect("C8", "row 505 status", "Поступления", "F505", [["учтено"]])
expect("C8", "counted payments", "Сводка", "B6", [[4.0]])
act("C8r", "clear row 505", "execute_command", command="edit.clearContents", params={"range": "Поступления!A505:E505"})
expect_summary("C8r", BASE_SOURCES, BASE_PAYMENTS, OCT31, ["A", "B"])
expect("C8r", "counted payments", "Сводка", "B6", [[3.0]])

# C9 validation rules (checked without entering values)
act("C9", "activate Поступления", "execute_command", command="sheet.activate", params={"sheet": "Поступления"})
for cell, value, ok in [("B6", "A", True), ("B6", "ZZ", False), ("A6", "2026-10-15", True),
                        ("A6", "05.11.2026", False), ("A6", "15.09.2026", False), ("C6", "-5", False)]:
    plan.append({"check": "C9", "title": f"data.validate {cell}={value!r} expected ok={ok}",
                 "tool": "execute_command",
                 "arguments": {"command": "data.validate", "params": {"cell": cell, "input": value}},
                 "expect_ok": ok})

# C10 save
act("C10", "activate Сводка, select B4", "execute_command", command="sheet.activate", params={"sheet": "Сводка"})
act("C10", "select B4", "execute_command", command="selection.set", params={"range": "B4"})
act("C10", "save", "execute_command", command="file.save", params={"path": BOOK})

# Key cells snapshotted with get_cell before close and after reopen.
# "style" lists the formatting the build plan applied (fill RGB, number format, bold).
Y, G, H = 0xFFF2CC, 0xEDEDED, 0xDDEBF7
MONEY, DATE_F, MONTH_F = "#,##0.00", "dd.mm.yyyy", "mm.yyyy"
SNAPSHOT = [
    {"sheet": "Сводка", "cell": "B4", "style": {"fill": Y, "num_fmt": DATE_F}},
    {"sheet": "Сводка", "cell": "B5", "style": {"fill": G, "num_fmt": DATE_F}},
    {"sheet": "Сводка", "cell": "B8", "style": {"fill": G, "num_fmt": MONEY}},
    {"sheet": "Сводка", "cell": "B10", "style": {"fill": G, "num_fmt": MONEY}},
    {"sheet": "Сводка", "cell": "A12", "style": {"fill": H, "bold": True}},
    {"sheet": "Сводка", "cell": "A13", "style": {"fill": G}},
    {"sheet": "Сводка", "cell": "D13", "style": {"fill": G, "num_fmt": MONEY}},
    {"sheet": "Сводка", "cell": "E13", "style": {"fill": G, "num_fmt": MONEY}},
    {"sheet": "Сводка", "cell": "F13", "style": {"fill": G, "num_fmt": MONEY}},
    {"sheet": "Сводка", "cell": "G14", "style": {"fill": G, "num_fmt": MONEY}},
    {"sheet": "Сводка", "cell": "H13", "style": {"fill": G}},
    {"sheet": "Сводка", "cell": "F33", "style": {"bold": True, "num_fmt": MONEY}},
    {"sheet": "Источники", "cell": "B4", "style": {"fill": Y, "num_fmt": MONTH_F}},
    {"sheet": "Источники", "cell": "A7", "style": {"fill": H, "bold": True}},
    {"sheet": "Источники", "cell": "C8", "style": {"fill": Y, "num_fmt": MONEY}},
    {"sheet": "Источники", "cell": "D8", "style": {"fill": Y, "num_fmt": MONTH_F}},
    {"sheet": "Поступления", "cell": "A5", "style": {"fill": H, "bold": True}},
    {"sheet": "Поступления", "cell": "A6", "style": {"fill": Y, "num_fmt": DATE_F}},
    {"sheet": "Поступления", "cell": "C6", "style": {"fill": Y, "num_fmt": MONEY}},
    {"sheet": "Поступления", "cell": "F6", "style": {"fill": G}},
    {"sheet": "Поступления", "cell": "F505", "style": {"fill": G}},
]

if __name__ == "__main__":
    RUN.mkdir(parents=True, exist_ok=True)
    (RUN / "snapshot_cells.json").write_text(json.dumps(SNAPSHOT, ensure_ascii=False, indent=1), encoding="utf-8")
    with open(RUN / "checks_plan.jsonl", "w", encoding="utf-8") as f:
        for i, step in enumerate(plan, 1):
            f.write(json.dumps({"step": i, **step}, ensure_ascii=False) + "\n")
    (RUN / "expected_model.json").write_text(json.dumps(expected, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{len(plan)} steps; C0 totals {expected['C0']['totals']}; C3 totals {expected['C3']['totals']}")
