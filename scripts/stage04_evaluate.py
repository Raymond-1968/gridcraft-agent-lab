"""Independent evaluation of Stage 04 session transcripts (same method as Stage 03).

Reuses scripts/stage03_evaluate.py with the Stage 04 run folder: build_plan, checks_plan
(K0..K13), snapshot_cells, sessions/session{1,2,3}_*.stream.jsonl -> runs/STAGE_04/evaluation.json.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import stage03_evaluate as ev  # noqa: E402

ev.RUN = ev.LAB / "runs" / "STAGE_04"
ev.SES = ev.RUN / "sessions"
ev.REOPEN_CHECK = "K0"

if __name__ == "__main__":
    sys.exit(ev.main())
