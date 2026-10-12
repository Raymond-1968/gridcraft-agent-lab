"""Independent evaluation of Stage 05 session transcripts (same method as Stage 03).

Reuses scripts/stage03_evaluate.py with the Stage 05 run folder: build_plan, checks_plan
(L0..L8), snapshot_cells, sessions/session{1,2,3}_*.stream.jsonl -> runs/STAGE_05/evaluation.json.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import stage03_evaluate as ev  # noqa: E402

ev.RUN = ev.LAB / "runs" / "STAGE_05"
ev.SES = ev.RUN / "sessions"
ev.REOPEN_CHECK = "L0"

if __name__ == "__main__":
    sys.exit(ev.main())
