"""Local GridCraft command catalog bound to the installed build.

    python scripts/catalog.py verify            # installed binaries + catalog files vs manifest
    python scripts/catalog.py search cell.set   # selective lookup (id / label / ribbon path)

`search` refuses to run when `verify` fails: the catalog is valid only for the
exact build whose hashes are recorded in the manifest.
"""

import hashlib
import json
import sys
from pathlib import Path

LAB = Path(__file__).resolve().parents[1]
CATALOG = LAB / "catalog" / "gridcraft-0.6.1"
MANIFEST = json.loads((CATALOG / "manifest.json").read_text(encoding="utf-8"))


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def verify():
    ok = True
    for path, expected in MANIFEST["binaries"].items():
        actual = sha256(path) if Path(path).exists() else "MISSING"
        good = actual == expected
        ok &= good
        print(f"{'OK  ' if good else 'FAIL'} binary  {path}  {actual}")
    for name, meta in MANIFEST["files"].items():
        actual = sha256(CATALOG / name)
        good = actual == meta["sha256"]
        ok &= good
        print(f"{'OK  ' if good else 'FAIL'} catalog {name}  {actual}")
    print(f"catalog GridCraft {MANIFEST['version']}: {'VALID' if ok else 'INVALID - regenerate in a new run'}")
    return ok


def search(term):
    t = term.lower()
    commands = json.loads((CATALOG / "commands_window.json").read_text(encoding="utf-8"))
    hits = [c for c in commands
            if t in c["id"].lower() or t in c.get("label", "").lower()
            or t in " ".join(c.get("menu") or []).lower()]
    print(json.dumps(hits, ensure_ascii=False, indent=1))
    return hits


if __name__ == "__main__":
    if len(sys.argv) >= 2 and sys.argv[1] == "verify":
        sys.exit(0 if verify() else 1)
    if len(sys.argv) == 3 and sys.argv[1] == "search":
        if not verify():
            sys.exit(1)
        sys.exit(0 if search(sys.argv[2]) else 2)
    print(__doc__)
    sys.exit(2)
