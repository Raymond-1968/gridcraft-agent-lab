"""Stage 01 smoke test: isolated GridCraft window + control channel + MCP bridge.

Every workbook operation is performed by GridCraft itself through
`gridcraft-cli mcp --connect PORT` (MCP JSON-RPC over stdio) or through the
desktop control channel (newline-delimited JSON on 127.0.0.1:PORT).
This script only sends requests, reads replies and compares them with
expected values. It never writes workbook files itself.

Synthetic test values only. No real financial data.

Usage:  python scripts/stage01_smoke.py [--port 7991]
"""

import argparse
import datetime as dt
import json
import os
import socket
import subprocess
import sys
import time
from pathlib import Path

GC_EXE = Path(r"C:\Program Files\GridCraft\gridcraft.exe")
CLI_EXE = Path(r"C:\Program Files\GridCraft\gridcraft-cli.exe")

LAB = Path(__file__).resolve().parents[1]
RUN = LAB / "runs" / "STAGE_01"
LOGS = RUN / "logs"
SHOTS = RUN / "screenshots"
WORKBOOKS = LAB / "workbooks"           # git-ignored
TMP = LAB / "tmp"                       # git-ignored
BOOK = WORKBOOKS / "stage01_smoke_test.xlsx"

A1, A2 = 12.5, 30          # synthetic inputs
SUM_EXPECTED = 42.5
A2_ALT, SUM_ALT = 100, 112.5


def now():
    return dt.datetime.now().isoformat(timespec="milliseconds")


class Transcript:
    def __init__(self, path):
        self.f = open(path, "w", encoding="utf-8")

    def write(self, **rec):
        rec = {"ts": now(), **rec}
        self.f.write(json.dumps(rec, ensure_ascii=False) + "\n")
        self.f.flush()


class Control:
    """Direct desktop control channel client (same wire format as `gridcraft-cli send`)."""

    def __init__(self, port, log):
        self.port, self.log, self.n = port, log, 1000

    def call(self, method, params=None):
        self.n += 1
        req = {"id": self.n, "method": method}
        if params is not None:
            req["params"] = params
        self.log.write(channel="control", dir="->", msg=req)
        with socket.create_connection(("127.0.0.1", self.port), timeout=30) as s:
            s.sendall((json.dumps(req) + "\n").encode())
            buf = b""
            while not buf.endswith(b"\n"):
                chunk = s.recv(65536)
                if not chunk:
                    break
                buf += chunk
        resp = json.loads(buf.decode("utf-8"))
        self.log.write(channel="control", dir="<-", msg=resp)
        return resp


class Mcp:
    """Minimal MCP stdio client for `gridcraft-cli mcp --connect PORT`."""

    def __init__(self, port, log, stderr_path):
        self.log, self.n = log, 0
        self.err = open(stderr_path, "w", encoding="utf-8")
        self.p = subprocess.Popen(
            [str(CLI_EXE), "mcp", "--connect", str(port)],
            stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=self.err,
            text=True, encoding="utf-8", bufsize=1)

    def _send(self, msg):
        self.log.write(channel="mcp", dir="->", msg=msg)
        self.p.stdin.write(json.dumps(msg) + "\n")
        self.p.stdin.flush()

    def request(self, method, params=None):
        self.n += 1
        msg = {"jsonrpc": "2.0", "id": self.n, "method": method}
        if params is not None:
            msg["params"] = params
        self._send(msg)
        line = self.p.stdout.readline()
        if not line:
            raise RuntimeError(f"MCP server closed stdout (exit={self.p.poll()})")
        resp = json.loads(line)
        self.log.write(channel="mcp", dir="<-", msg=resp)
        return resp

    def notify(self, method, params=None):
        msg = {"jsonrpc": "2.0", "method": method}
        if params is not None:
            msg["params"] = params
        self._send(msg)

    def tool(self, name, **args):
        resp = self.request("tools/call", {"name": name, "arguments": args})
        if "error" in resp:
            raise RuntimeError(f"{name}: JSON-RPC error {resp['error']}")
        res = resp["result"]
        text = "".join(c.get("text", "") for c in res.get("content", []))
        if res.get("isError"):
            raise RuntimeError(f"{name}: tool error: {text}")
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            return text

    def close(self):
        try:
            self.p.stdin.close()
            self.p.wait(timeout=10)
        except Exception:
            self.p.kill()
        self.err.close()


def wait_port(port, timeout=30):
    end = time.time() + timeout
    while time.time() < end:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=1):
                return True
        except OSError:
            time.sleep(0.3)
    return False


def port_free(port):
    with socket.socket() as s:
        return s.connect_ex(("127.0.0.1", port)) != 0


def launch_window(port, tag):
    env = dict(os.environ, GRIDCRAFT_NO_PREFS="1")
    err = open(LOGS / f"gridcraft_window_{tag}.stderr.log", "w", encoding="utf-8")
    p = subprocess.Popen([str(GC_EXE), "--control", str(port)], cwd=TMP, env=env,
                         stdout=err, stderr=subprocess.STDOUT)
    if not wait_port(port):
        p.kill()
        raise RuntimeError(f"control port {port} did not open within 30 s")
    return p, err


class Steps:
    def __init__(self, log):
        self.log, self.rows, self.failed = log, [], False

    def check(self, stage, name, ok, detail):
        status = "PASS" if ok else "FAIL"
        self.rows.append({"stage": stage, "check": name, "status": status, "detail": detail})
        self.log.write(channel="check", stage=stage, check=name, status=status, detail=detail)
        print(f"[{status}] {stage}: {name} -- {detail}")
        if not ok:
            self.failed = True
        return ok


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def cell_value(info):
    """Pick the computed value out of a get_cell result."""
    if isinstance(info, dict):
        for k in ("value", "result"):
            if k in info:
                v = info[k]
                if isinstance(v, dict) and "value" in v:
                    v = v["value"]
                return v
    return info


def has_flag(info, key):
    """True when style.font.<key> (bold/italic) is set in a get_cell result."""
    try:
        return info["style"]["font"][key] is True
    except (KeyError, TypeError):
        return False


def shown(info):
    """Displayed text of a get_cell result."""
    return info.get("text") if isinstance(info, dict) else None


def fill_rgb(info):
    try:
        return info["style"]["fill"]["fg"]["Rgb"]
    except (KeyError, TypeError):
        return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=7991)
    port = ap.parse_args().port

    for d in (LOGS, SHOTS, WORKBOOKS, TMP):
        d.mkdir(parents=True, exist_ok=True)
    if BOOK.exists():
        BOOK.unlink()  # lab-owned synthetic test book from an earlier run

    log = Transcript(LOGS / "transcript.jsonl")
    st = Steps(log)
    stage = "0-preconditions"
    win = mcp = None
    try:
        st.check(stage, "control port free before launch", port_free(port), f"127.0.0.1:{port}")

        # ---- 1. launch isolated window ------------------------------------------------
        stage = "1-launch-window"
        win, win_err = launch_window(port, "session1")
        st.check(stage, "window started, control channel listening", True, f"pid={win.pid}")
        ctl = Control(port, log)
        r = ctl.call("document.inspect")
        st.check(stage, "control channel document.inspect", r.get("ok") is True,
                 json.dumps(r.get("result", r.get("error")), ensure_ascii=False)[:300])
        r = ctl.call("engine.commands")
        if r.get("ok"):
            (RUN / "catalog" / "commands_live_window.json").write_text(
                json.dumps(r["result"], ensure_ascii=False, indent=1), encoding="utf-8")
        st.check(stage, "control channel engine.commands", r.get("ok") is True,
                 f"{len(r.get('result') or [])} commands from live window")

        # ---- 2. MCP bridge --------------------------------------------------------------
        stage = "2-mcp-connect"
        mcp = Mcp(port, log, LOGS / "mcp_bridge_session1.stderr.log")
        init = mcp.request("initialize", {"protocolVersion": "2025-06-18", "capabilities": {},
                                          "clientInfo": {"name": "gridcraft-agent-lab", "version": "stage01"}})
        mcp.notify("notifications/initialized")
        st.check(stage, "initialize", "result" in init,
                 json.dumps(init.get("result", {}).get("serverInfo"), ensure_ascii=False))
        tl = mcp.request("tools/list")
        tools = [t["name"] for t in tl.get("result", {}).get("tools", [])]
        (RUN / "catalog" / "mcp_tools.json").write_text(
            json.dumps(tl.get("result"), ensure_ascii=False, indent=1), encoding="utf-8")
        st.check(stage, "tools/list", bool(tools), ", ".join(tools))

        wb = mcp.tool("inspect_workbook")
        st.check(stage, "inspect_workbook on start-up book", True,
                 json.dumps(wb, ensure_ascii=False)[:300])
        used = mcp.tool("read_range")
        st.check(stage, "start-up book is empty", not any(
            any(c not in (None, "") for c in row) for row in (used.get("values") or [])) if isinstance(used, dict) else False,
                 json.dumps(used, ensure_ascii=False)[:200])

        # ---- 3. values + formula -------------------------------------------------------------
        stage = "3-write-and-calc"
        mcp.tool("set_cell", cell="A1", input=str(A1))
        mcp.tool("set_cell", cell="A2", input=str(A2))
        mcp.tool("set_cell", cell="A3", input="=SUM(A1:A2)")
        c3 = mcp.tool("get_cell", cell="A3")
        st.check(stage, "A3 formula stored", c3.get("formula") == "=SUM(A1:A2)", json.dumps(c3, ensure_ascii=False)[:300])
        st.check(stage, "A3 computed by GridCraft = 42.5", num(cell_value(c3)) == SUM_EXPECTED,
                 f"value={cell_value(c3)!r}")
        rr = mcp.tool("read_range", range="A1:A3")
        st.check(stage, "read_range A1:A3", rr.get("values") == [[A1], [A2], [SUM_EXPECTED]], json.dumps(rr, ensure_ascii=False)[:200])

        # ---- 4. formatting -------------------------------------------------------------
        stage = "4-format"
        mcp.tool("format_range", range="A3", bold=True, fillColor="#FFF2CC", numberFormat="0.00")
        f3 = mcp.tool("get_cell", cell="A3")
        st.check(stage, "A3 bold", has_flag(f3, "bold"), json.dumps(f3, ensure_ascii=False)[:400])
        st.check(stage, "A3 displayed as 42.50", shown(f3) == "42.50", f"text={shown(f3)!r} (format 0.00)")
        st.check(stage, "A3 fill #FFF2CC", fill_rgb(f3) == 0xFFF2CC, f"fill.fg.Rgb={fill_rgb(f3)}")

        # ---- 5. undo / redo ---------------------------------------------------------------
        stage = "5-undo-redo"
        mcp.tool("set_cell", cell="A2", input=str(A2_ALT))
        v = num(cell_value(mcp.tool("get_cell", cell="A3")))
        st.check(stage, "edit A2=100 -> A3=112.5", v == SUM_ALT, f"A3={v}")
        mcp.tool("undo")
        v = num(cell_value(mcp.tool("get_cell", cell="A3")))
        st.check(stage, "undo -> A3=42.5", v == SUM_EXPECTED, f"A3={v}")
        mcp.tool("redo")
        v = num(cell_value(mcp.tool("get_cell", cell="A3")))
        st.check(stage, "redo -> A3=112.5", v == SUM_ALT, f"A3={v}")
        mcp.tool("undo")
        v = num(cell_value(mcp.tool("get_cell", cell="A3")))
        st.check(stage, "undo again -> A3=42.5", v == SUM_EXPECTED, f"A3={v}")
        # formatting undo/redo with a single command
        mcp.tool("execute_command", command="home.italic", params={"range": "A3"})
        st.check(stage, "italic applied", has_flag(mcp.tool("get_cell", cell="A3"), "italic"), "home.italic")
        mcp.tool("undo")
        st.check(stage, "undo removes italic", not has_flag(mcp.tool("get_cell", cell="A3"), "italic"), "")
        mcp.tool("redo")
        st.check(stage, "redo restores italic", has_flag(mcp.tool("get_cell", cell="A3"), "italic"), "")

        # ---- 6. save -------------------------------------------------------------------------
        stage = "6-save"
        sv = mcp.tool("save_workbook", path=str(BOOK))
        st.check(stage, "save_workbook", BOOK.exists() and BOOK.stat().st_size > 0,
                 f"{BOOK.relative_to(LAB)} size={BOOK.stat().st_size if BOOK.exists() else 0} reply={json.dumps(sv, ensure_ascii=False)[:150]}")
        shot1 = SHOTS / "stage01_session1_after_save.png"
        mcp.tool("screenshot", path=str(shot1))
        st.check(stage, "screenshot after save", shot1.exists() and shot1.stat().st_size > 0, str(shot1.relative_to(LAB)))
        j = ctl.call("engine.journal")
        (LOGS / "engine_journal_session1.json").write_text(
            json.dumps(j.get("result"), ensure_ascii=False, indent=1), encoding="utf-8")
        st.check(stage, "engine.journal captured", j.get("ok") is True, "logs/engine_journal_session1.json")

        mcp.close(); mcp = None
        q = ctl.call("app.quit")
        try:
            win.wait(timeout=20)
            st.check(stage, "window quit via app.quit", True, f"reply={q} exit={win.returncode}")
        except subprocess.TimeoutExpired:
            st.check(stage, "window quit via app.quit", False, f"reply={q}; still running after 20 s")
            win.kill()
        win_err.close(); win = None

        # ---- 7. reopen in a fresh window process ----------------------------------------------
        stage = "7-reopen"
        win, win_err = launch_window(port, "session2")
        mcp = Mcp(port, log, LOGS / "mcp_bridge_session2.stderr.log")
        mcp.request("initialize", {"protocolVersion": "2025-06-18", "capabilities": {},
                                   "clientInfo": {"name": "gridcraft-agent-lab", "version": "stage01"}})
        mcp.notify("notifications/initialized")
        op = mcp.tool("open_workbook", path=str(BOOK))
        st.check(stage, "open_workbook in fresh process", True, json.dumps(op, ensure_ascii=False)[:200])
        r3 = mcp.tool("get_cell", cell="A3")
        st.check(stage, "reopened A3 formula", r3.get("formula") == "=SUM(A1:A2)", json.dumps(r3, ensure_ascii=False)[:400])
        st.check(stage, "reopened A3 value = 42.5", num(cell_value(r3)) == SUM_EXPECTED, f"value={cell_value(r3)!r}")
        st.check(stage, "reopened A3 bold + italic", has_flag(r3, "bold") and has_flag(r3, "italic"), "")
        st.check(stage, "reopened A3 displayed 42.50", shown(r3) == "42.50", f"text={shown(r3)!r}")
        st.check(stage, "reopened A3 fill #FFF2CC", fill_rgb(r3) == 0xFFF2CC, f"fill.fg.Rgb={fill_rgb(r3)}")
        rr = mcp.tool("read_range", range="A1:A3")
        st.check(stage, "reopened values A1:A3", rr.get("values") == [[A1], [A2], [SUM_EXPECTED]],
                 json.dumps(rr, ensure_ascii=False)[:200])
        mcp.tool("select", range="A3")
        shot2 = SHOTS / "stage01_result_reopened.png"
        mcp.tool("screenshot", path=str(shot2))
        st.check(stage, "screenshot of reopened result", shot2.exists() and shot2.stat().st_size > 0,
                 str(shot2.relative_to(LAB)))
        mcp.close(); mcp = None
        ctl.call("app.quit")
        try:
            win.wait(timeout=20)
        except subprocess.TimeoutExpired:
            win.kill()
        win_err.close(); win = None
    except Exception as e:  # record the exact stage and error
        st.check(stage, "unexpected error", False, f"{type(e).__name__}: {e}")
    finally:
        if mcp:
            mcp.close()
        if win and win.poll() is None:
            win.kill()

    verdict = "FAIL" if st.failed else "PASS"
    (RUN / "results.json").write_text(json.dumps(
        {"verdict": verdict, "port": port, "finished": now(), "checks": st.rows},
        ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\nSTAGE 01 VERDICT: {verdict}")
    return 0 if verdict == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
