"""Independent evaluation of Stage 03 session transcripts.

Reads the Claude Code stream-json transcripts of sessions 1-3, extracts every MCP call
with its raw GridCraft reply (runs/STAGE_03/sessions/sessionN_mcp_calls.jsonl), and re-checks:
  - session 1: every build_plan step executed in order without errors;
  - session 2: every checks_plan expectation against the raw read_range / data.validate replies;
  - sessions 2 vs 3: snapshot cells (formula, value, fill, number format, bold) before close
    and after reopen, plus the style expected by the build plan;
  - session 3: C0 values after reopen.
Writes runs/STAGE_03/evaluation.json. Does not touch the workbook.
"""

import json
import sys
from pathlib import Path

LAB = Path(__file__).resolve().parents[1]
RUN = LAB / "runs" / "STAGE_03"
SES = RUN / "sessions"
REOPEN_CHECK = "C0"   # check id whose expectations are re-read after reopen


def load_calls(name):
    calls, pending = [], {}
    for line in open(SES / f"{name}.stream.jsonl", encoding="utf-8"):
        e = json.loads(line)
        if e.get("type") == "assistant":
            for c in e["message"]["content"]:
                if c["type"] == "tool_use":
                    pending[c["id"]] = c
        elif e.get("type") == "user":
            for c in e["message"]["content"]:
                if isinstance(c, dict) and c.get("type") == "tool_result" and c["tool_use_id"] in pending:
                    u = pending.pop(c["tool_use_id"])
                    cc = c["content"]
                    text = cc if isinstance(cc, str) else "".join(x.get("text", "") for x in cc)
                    try:
                        res = json.loads(text)
                    except json.JSONDecodeError:
                        res = text
                    calls.append({"tool": u["name"].split("__")[-1], "arguments": u["input"],
                                  "is_error": bool(c.get("is_error")), "result": res})
    with open(SES / f"{name}_mcp_calls.jsonl", "w", encoding="utf-8") as f:
        for c in calls:
            f.write(json.dumps(c, ensure_ascii=False) + "\n")
    return calls


def same(a, e):
    if e == "":
        return a in ("", None)
    if isinstance(e, (int, float)) and not isinstance(e, bool):
        return isinstance(a, (int, float)) and abs(a - e) < 0.005
    return a == e


def is_err(v):
    return isinstance(v, str) and v.startswith("#")


def match_reads(calls, steps):
    """Pair plan steps with session calls in order (same tool + same arguments).

    Sessions sometimes send a few independent calls as one batch in a different order, so a
    read-only step not found forward is also looked up among the last few unused calls.
    """
    out, i, used = [], 0, set()

    def same_call(c, st):
        return c["tool"] == st["tool"] and c["arguments"] == st["arguments"]

    for st in steps:
        j = i
        while j < len(calls) and (j in used or not same_call(calls[j], st)):
            j += 1
        if j == len(calls) and st["tool"] == "read_range":
            j = next((k for k in range(max(0, i - 5), i) if k not in used and same_call(calls[k], st)), len(calls))
        if j == len(calls):
            out.append((st, None))
        else:
            out.append((st, calls[j]))
            used.add(j)
            i = max(i, j + 1)
    return out


def snapshot(calls):
    snap, sheet = {}, None
    for c in calls:
        a = c["arguments"]
        if c["tool"] == "execute_command" and a.get("command") == "sheet.activate":
            sheet = a["params"]["sheet"]
        elif c["tool"] == "get_cell" and isinstance(c["result"], dict):
            r = c["result"]
            st = r.get("style", {})
            fg = (st.get("fill") or {}).get("fg")
            snap[f"{sheet}!{a['cell']}"] = {
                "formula": r.get("formula"), "value": r.get("value"), "text": r.get("text"),
                "fill": fg.get("Rgb") if isinstance(fg, dict) else fg,
                "num_fmt": st.get("num_fmt"), "bold": (st.get("font") or {}).get("bold")}
    return snap


def main():
    ev = {"checks": [], "summary": {}}
    plan = [json.loads(l) for l in open(RUN / "build_plan.jsonl", encoding="utf-8")]
    checks = [json.loads(l) for l in open(RUN / "checks_plan.jsonl", encoding="utf-8")]
    snap_cells = json.loads((RUN / "snapshot_cells.json").read_text(encoding="utf-8"))

    def rec(group, name, ok, detail=""):
        ev["checks"].append({"group": group, "check": name, "status": "PASS" if ok else "FAIL", "detail": detail})

    # session 1: build
    s1 = load_calls("session1_build")
    ex = [c for c in s1 if c["tool"] == "execute_command"]
    built = match_reads(ex, [{"tool": "execute_command",
                              "arguments": {"command": p["command"], "params": p["params"]}} for p in plan])
    missing = [st["arguments"]["command"] for st, c in built if c is None]
    errors = [(c["arguments"]["command"], c["result"]) for st, c in built if c and c["is_error"]]
    rec("build", f"all {len(plan)} plan steps executed verbatim via execute_command",
        not missing and not errors, f"missing={missing[:5]} errors={errors[:3]}")
    saves = [c for c in s1 if c["arguments"].get("command") == "file.save" and not c["is_error"]]
    rec("build", "file.save income_tracker_prototype.xlsx", bool(saves),
        json.dumps(saves[-1]["result"], ensure_ascii=False) if saves else "no save")

    # session 2: checks
    s2 = load_calls("session2_checks")
    for st, c in match_reads(s2, [s for s in checks]):
        if not any(k in st for k in ("expect", "expect_no_errors", "expect_ok")):
            if c is None or c["is_error"]:
                rec(st["check"], f"step {st['step']} action: {st['title']}", False,
                    "not executed" if c is None else str(c["result"]))
            continue
        if c is None:
            rec(st["check"], f"step {st['step']}: {st['title']}", False, "not executed")
            continue
        res = c["result"]
        if "expect" in st:
            vals = res.get("values") if isinstance(res, dict) else None
            mism = []
            for i, row in enumerate(st["expect"]):
                for j, e in enumerate(row):
                    a = vals[i][j] if vals and i < len(vals) and j < len(vals[i]) else None
                    if not same(a, e):
                        mism.append(f"[{i},{j}] got {a!r} expected {e!r}")
            rec(st["check"], f"step {st['step']}: {st['title']} ({st['arguments']['sheet']}!{st['arguments']['range']})",
                not mism, "; ".join(mism) or json.dumps(vals, ensure_ascii=False))
        elif st.get("expect_no_errors"):
            vals = res.get("values") or []
            errs = [v for row in vals for v in row if is_err(v)]
            rec(st["check"], f"step {st['step']}: {st['title']}", not errs, f"errors={errs[:5]}")
        else:
            ok = isinstance(res, dict) and res.get("ok") == st["expect_ok"]
            rec(st["check"], f"step {st['step']}: {st['title']}", ok, json.dumps(res, ensure_ascii=False))

    # sessions 2 vs 3: snapshot + expected style
    s3 = load_calls("session3_reopen")
    before, after = snapshot(s2), snapshot(s3)
    for sc in snap_cells:
        key = f"{sc['sheet']}!{sc['cell']}"
        b, a = before.get(key), after.get(key)
        rec("reopen", f"{key} identical before close / after reopen", b is not None and a == b,
            f"before={b} after={a}")
        if a:
            bad = [f"{k}: {a.get(k)!r} != {v!r}" for k, v in sc["style"].items() if a.get(k) != v]
            rec("format", f"{key} style after reopen {sc['style']}", not bad, "; ".join(bad))
    c0 = [s for s in checks if s["check"] == REOPEN_CHECK]
    for st, c in match_reads(s3, c0):
        vals = c["result"].get("values") if c and isinstance(c["result"], dict) else None
        mism = [f"[{i},{j}] got {(vals[i][j] if vals else None)!r} expected {e!r}"
                for i, row in enumerate(st["expect"]) for j, e in enumerate(row)
                if not (vals and same(vals[i][j], e))]
        rec("reopen", f"{REOPEN_CHECK} after reopen: {st['title']} ({st['arguments']['sheet']}!{st['arguments']['range']})",
            not mism, "; ".join(mism) or json.dumps(vals, ensure_ascii=False))

    total = len(ev["checks"])
    failed = [c for c in ev["checks"] if c["status"] == "FAIL"]
    ev["summary"] = {"total": total, "passed": total - len(failed), "failed": len(failed),
                     "verdict": "PASS" if not failed else "FAIL"}
    (RUN / "evaluation.json").write_text(json.dumps(ev, ensure_ascii=False, indent=1), encoding="utf-8")
    groups = {}
    for c in ev["checks"]:
        g = groups.setdefault(c["group"], [0, 0])
        g[0 if c["status"] == "PASS" else 1] += 1
    for g, (p, f) in groups.items():
        print(f"{g:8s} PASS {p:3d}  FAIL {f:3d}")
    for c in failed:
        print("FAIL", c["group"], c["check"], "--", c["detail"][:300])
    print("EVALUATION:", ev["summary"])
    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())
