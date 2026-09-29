"""
Step 2: pull finished annotations back out of Label Studio.

Downloads Label Studio's native JSON export (one entry per task, each with
its image path and the shelf_level boxes drawn on it) to
labeling/exports/export_<timestamp>.json. Step 3 (dataset prep, not built
yet) will convert this into YOLO-format label files + train/val splits --
this script's only job is getting a raw, dated snapshot of what's been
labeled so far.

Usage:
    python export_annotations.py --config config.json
"""
from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path

import requests

from _auth import get_access_token

SCRIPT_DIR = Path(__file__).resolve().parent
EXPORTS_DIR = SCRIPT_DIR / "exports"


def load_config(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def main():
    parser = argparse.ArgumentParser(description="Step 2: export Label Studio annotations")
    parser.add_argument("--config", default=str(SCRIPT_DIR / "config.json"))
    parser.add_argument("--out", default=None, help="Output path (default: exports/export_<timestamp>.json)")
    args = parser.parse_args()

    config = load_config(Path(args.config))
    project_id = config.get("project_id")
    if not project_id:
        raise SystemExit("config.json has no project_id yet -- run import_tasks.py first.")

    access_token = get_access_token(config)
    url = f'{config["label_studio_url"].rstrip("/")}/api/projects/{project_id}/export'
    response = requests.get(
        url,
        headers={"Authorization": f"Bearer {access_token}"},
        params={"exportType": "JSON", "download_all_tasks": "false"},
        timeout=300,
    )
    if response.status_code != 200:
        raise SystemExit(
            f"Export request failed ({response.status_code}): {response.text[:500]}\n"
            "If this Label Studio version has moved export behind the async "
            "'Create Export' workflow, use the Export button in the project "
            "UI (Export > JSON) and save the file into labeling/exports/ instead."
        )
    tasks = response.json()

    EXPORTS_DIR.mkdir(parents=True, exist_ok=True)
    out_path = Path(args.out) if args.out else EXPORTS_DIR / f'export_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json'
    out_path.write_text(json.dumps(tasks, indent=2), encoding="utf-8")

    annotated = sum(1 for task in tasks if task.get("annotations"))
    print(f"Exported {len(tasks)} task(s), {annotated} with at least one annotation, to {out_path}")


if __name__ == "__main__":
    main()
