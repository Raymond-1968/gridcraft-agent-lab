"""Stage 05 UI check through the GridCraft control channel (`gridcraft-cli send`).

Real keyboard input in the running lab window on sheet «Учёт»:
  - Ctrl+; (GridCraft static date stamp) in a new entry row, then payer and amount;
    the stamp must be a literal (no formula) and must not change after recalculation;
  - a manual entry дд.мм.гггг with day <= 12;
  - an invalid date that validation must stop;
  - cleanup of the test rows. Every request/reply is logged.
GridCraft dates come from UTC (observed); the log records local and UTC dates.
"""

import argparse
import datetime as dt
import json
import subprocess
import sys
import time
from pathlib import Path

CLI = r"C:\Program Files\GridCraft\gridcraft-cli.exe"
LAB = Path(__file__).resolve().parents[1]
RUN = LAB / "runs" / "STAGE_05"


def serial(d):
    return float((d - dt.date(1899, 12, 30)).days)


class Ctl:
    def __init__(self, port, log):
        self.port, self.log = port, log

    def send(self, method, params=None):
        args = [CLI, "send", str(self.port), method] + ([json.dumps(params, ensure_ascii=False)] if params is not None else [])
        p = subprocess.run(args, capture_output=True, text=True, encoding="utf-8")
        out = p.stdout.strip()
        try:
            res = json.loads(out) if out else None
        except json.JSONDecodeError:
            res = out
        self.log.write(json.dumps({"ts": dt.datetime.now().isoformat(timespec="milliseconds"), "method": method,
                                   "params": params, "exit": p.returncode, "reply": res,
                                   "stderr": p.stderr.strip() or None}, ensure_ascii=False) + "\n")
        self.log.flush()
        return res

    def key_in(self, cell, keys):
        self.send("ui.cell.click", {"cell": cell})
        for k in keys:
            self.send("ui.text", {"text": k}) if isinstance(k, str) else self.send("ui.key", k)
        self.send("ui.key", {"key": "Enter"})
        time.sleep(0.6)
        return self.send("ui.inspect")

    def get(self, cell):
        return self.send("cell.get", {"cell": cell}) or {}

    def read(self, rng):
        return (self.send("sheet.read", {"sheet": "Учёт", "range": rng}) or {}).get("values")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=7991)
    port = ap.parse_args().port
    rows = []

    def check(name, ok, detail):
        rows.append({"check": name, "status": "PASS" if ok else "FAIL", "detail": detail})
        print(f"[{'PASS' if ok else 'FAIL'}] {name} -- {detail}")

    utc_today = dt.datetime.now(dt.timezone.utc).date()
    local_today = dt.datetime.now().date()
    with open(RUN / "ui_typing_log.jsonl", "w", encoding="utf-8") as log:
        c = Ctl(port, log)
        c.send("sheet.activate", {"sheet": "Учёт"})
        base = c.read("B9:B10")
        check("start state: remaining 30, overpayment 0", base == [[30.0], [0.0]], f"B9:B10={base}")

        ui = c.key_in("A20", [{"key": ";", "ctrl": True}])
        a20 = c.get("A20")
        check("Ctrl+; wrote a literal (no formula)", a20.get("formula") is None and a20.get("type") == "text",
              f"input={a20.get('input')!r} type={a20.get('type')} formula={a20.get('formula')}")
        h20 = c.get("H20")
        check("stamp recognised as GridCraft's (UTC) date",
              h20.get("value") == serial(utc_today),
              f"H20={h20.get('value')} UTC today={utc_today} ({serial(utc_today)}) local today={local_today}")
        check("note: stamp equals local date?", True,
              f"{'same' if utc_today == local_today else 'DIFFERENT: UTC date ' + str(utc_today) + ' vs local ' + str(local_today)}")
        c.key_in("B20", ["A"])
        ui = c.key_in("C20", ["25"])
        e20 = c.get("E20")
        b9 = c.read("B9")
        check("stamped payment A 25 counted; «Проверка» blank", e20.get("value") in ("", None) and b9 == [[5.0]],
              f"E20={e20.get('value')!r} B9={b9} message={(ui or {}).get('message')}")

        c.send("formulas.calculateNow")
        time.sleep(0.3)
        a20b, h20b = c.get("A20"), c.get("H20")
        check("stamp unchanged after recalculation", a20b.get("input") == a20.get("input") and h20b.get("value") == h20.get("value"),
              f"A20 {a20.get('input')!r} -> {a20b.get('input')!r}; H20 {h20.get('value')} -> {h20b.get('value')}")

        ui = c.key_in("A21", ["01.11.2026"])
        check("typed 01.11.2026 accepted", (ui or {}).get("message") is None, f"message={(ui or {}).get('message')}")
        c.key_in("B21", ["B"])
        c.key_in("C21", ["5"])
        h21, e21, b9 = c.get("H21"), c.get("E21"), c.read("B9:B10")
        check("01.11.2026 recognised as 1 November, counted; balance 0", h21.get("value") == serial(dt.date(2026, 11, 1))
              and e21.get("value") in ("", None) and b9 == [[0.0], [0.0]],
              f"H21={h21.get('value')} E21={e21.get('value')!r} B9:B10={b9}")

        ui = c.key_in("A22", ["15/10/2026"])
        msg = json.dumps((ui or {}).get("message"), ensure_ascii=False)
        check("typed 15/10/2026 stopped by validation", "дд.мм.гггг" in msg, msg)
        c.send("ui.edit.cancel")
        c.send("ui.key", {"key": "Escape"})
        time.sleep(0.4)
        check("A22 stays empty", c.get("A22").get("type") == "empty", "")

        c.send("engine.execute", {"command": "edit.clearContents", "params": {"range": "Учёт!A20:D22"}})
        end = c.read("B9:B10")
        check("cleanup: test rows cleared, remaining back to 30", end == [[30.0], [0.0]], f"B9:B10={end}")
        c.send("selection.set", {"range": "A20"})

    verdict = "PASS" if all(r["status"] == "PASS" for r in rows) else "FAIL"
    (RUN / "ui_typing_results.json").write_text(json.dumps(
        {"verdict": verdict, "utc_today": str(utc_today), "local_today": str(local_today), "checks": rows},
        ensure_ascii=False, indent=1), encoding="utf-8")
    print("UI TYPING VERDICT:", verdict)
    return 0 if verdict == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
