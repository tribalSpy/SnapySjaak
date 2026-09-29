"""
Step 7: close the loop -- pull newly-flagged needs_review references into
the labeling queue automatically.

Every night, shelf-poller-app already marks a shelf_count row as
needs_review when the live model's confidence is low. That's exactly the
"hard case" data/collect.py's priority tier already knows how to fetch --
this script just chains collect.py (priority only, no general sampling)
and labeling/import_tasks.py so those photos land in Label Studio without
someone remembering to run both by hand.

It does not retrain anything itself -- once there's enough freshly-labeled
data to be worth it, run training/prepare_dataset.py + training/train.py
manually (retraining on every single new label would be wasteful and can
overfit to whatever just happened to be reviewed that week).

Usage:
    python run_feedback_cycle.py --days-back 14
"""
from __future__ import annotations

import argparse
import subprocess
import sys
from datetime import date, timedelta
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIR = SCRIPT_DIR.parent / "data"
LABELING_DIR = SCRIPT_DIR.parent / "labeling"


def main():
    parser = argparse.ArgumentParser(description="Step 7: pull newly-flagged needs_review photos into labeling")
    parser.add_argument("--days-back", type=int, default=14, help="How far back to look for needs_review references")
    parser.add_argument("--data-config", default=str(DATA_DIR / "config.json"))
    parser.add_argument("--labeling-config", default=str(LABELING_DIR / "config.json"))
    args = parser.parse_args()

    date_to = date.today().isoformat()
    date_from = (date.today() - timedelta(days=args.days_back)).isoformat()

    print(f"Pulling needs_review references from {date_from} to {date_to}...")
    subprocess.run(
        [sys.executable, str(DATA_DIR / "collect.py"), "--config", args.data_config,
         "--from-date", date_from, "--to-date", date_to, "--count", "0"],
        check=True,
    )

    print("\nImporting any newly-collected photos into Label Studio...")
    subprocess.run(
        [sys.executable, str(LABELING_DIR / "import_tasks.py"), "--config", args.labeling_config],
        check=True,
    )

    print(
        "\nFeedback cycle done. Label the new photos in Label Studio, then once you've "
        "accumulated enough new labels to be worth it, re-run:\n"
        "  training/prepare_dataset.py  (rebuild the YOLO dataset)\n"
        "  training/train.py            (retrain)\n"
        "  evaluation/evaluate.py       (check it actually improved)\n"
        "  training/promote_model.py    (register the new version)"
    )


if __name__ == "__main__":
    main()
