"""Stage 04 UI typing check through the GridCraft control channel (`gridcraft-cli send`).

Simulates a user typing into the running lab window (real pointer/keyboard input handled by
GridCraft's UI): a payment with a day <= 12 in дд.мм.гггг form, and an invalid date that the
validation must stop. Then removes the test row. Every request/reply is logged.

Usage: python scripts/stage04_ui_typing.py [--port 7991]
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
LOG = LAB / "runs" / "STAGE_04" / "ui_typing_log.jsonl"


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

    def type_cell(self, cell, text):
        self.send("ui.cell.click", {"cell": cell})
        self.send("ui.text", {"text": text})
        self.send("ui.key", {"key": "Enter"})
        time.sleep(0.6)
        return self.send("ui.inspect")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=7991)
    port = ap.parse_args().port
    rows = []

    def check(name, ok, detail):
        rows.append({"check": name, "status": "PASS" if ok else "FAIL", "detail": detail})
        print(f"[{'PASS' if ok else 'FAIL'}] {name} -- {detail}")

    with open(LOG, "w", encoding="utf-8") as log:
        c = Ctl(port, log)
        c.send("sheet.activate", {"sheet": "Поступления"})
        for cell, text in (("A13", "01.11.2026"), ("B13", "B"), ("C13", "10"), ("D13", "09.2026")):
            ui = c.type_cell(cell, text)
            msg = (ui or {}).get("message")
            check(f"typed {cell}={text!r} accepted (no validation message, editor closed)",
                  msg is None and (ui or {}).get("editor") is None, f"message={msg} editor={(ui or {}).get('editor')}")
        a13 = c.send("cell.get", {"cell": "A13"}) or {}
        check("A13 stored as typed text", a13.get("type") == "text" and a13.get("value") == "01.11.2026",
              f"type={a13.get('type')} value={a13.get('value')!r}")
        f13 = c.send("cell.get", {"cell": "F13"}) or {}
        check("F13 recognised as 01.11.2026 (not 11 January)", f13.get("value") == serial(dt.date(2026, 11, 1)),
              f"value={f13.get('value')} text={f13.get('text')!r}")
        h13 = c.send("cell.get", {"cell": "H13"}) or {}
        check("H13 status 'учтено'", h13.get("value") == "учтено", f"value={h13.get('value')!r}")
        s = c.send("sheet.read", {"sheet": "Сводка", "range": "F6"}) or {}
        check("Сводка F6 saldo 30 -> 20 after typed B 10 for 09.2026", s.get("values") == [[20.0]], json.dumps(s, ensure_ascii=False))

        ui = c.type_cell("A14", "15/10/2026")
        msg = json.dumps((ui or {}).get("message"), ensure_ascii=False)
        check("typed A14='15/10/2026' stopped by validation", "дд.мм.гггг" in msg, msg)
        c.send("ui.edit.cancel")
        c.send("ui.key", {"key": "Escape"})
        time.sleep(0.4)
        a14 = c.send("cell.get", {"cell": "A14"}) or {}
        check("A14 stays empty", a14.get("type") == "empty", f"type={a14.get('type')} value={a14.get('value')!r}")

        c.send("engine.execute", {"command": "edit.clearContents", "params": {"range": "Поступления!A13:E13"}})
        s = c.send("sheet.read", {"sheet": "Сводка", "range": "F6"}) or {}
        h13 = c.send("cell.get", {"cell": "H13"}) or {}
        check("cleanup: row 13 cleared, saldo back to 30", s.get("values") == [[30.0]] and h13.get("value") in ("", None),
              f"F6={s.get('values')} H13={h13.get('value')!r}")
        c.send("sheet.activate", {"sheet": "Сводка"})

    verdict = "PASS" if all(r["status"] == "PASS" for r in rows) else "FAIL"
    (LAB / "runs" / "STAGE_04" / "ui_typing_results.json").write_text(
        json.dumps({"verdict": verdict, "checks": rows}, ensure_ascii=False, indent=1), encoding="utf-8")
    print("UI TYPING VERDICT:", verdict)
    return 0 if verdict == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
