"""
Step 2: import collected photos into Label Studio as labeling tasks.

Reads data/dataset/manifest.csv (written by data/collect.py) and pushes any
row not already imported to a Label Studio project, as a task pointing at
the photo via Label Studio's local-files serving (so photos are never
uploaded/copied -- Label Studio just needs LOCAL_FILES_DOCUMENT_ROOT set to
the same dataset_dir when the server starts).

Creates the project on first run (using labeling_config.xml) and writes the
resulting project_id back into config.json, so reruns reuse it.

Usage:
    python import_tasks.py --config config.json
"""
from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path

import requests

from _auth import api_headers, get_access_token

SCRIPT_DIR = Path(__file__).resolve().parent
STATE_PATH = SCRIPT_DIR / "imported_tasks.json"


def load_config(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def save_config(path: Path, config: dict):
    path.write_text(json.dumps(config, indent=2) + "\n", encoding="utf-8")


def load_state() -> set[str]:
    if not STATE_PATH.exists():
        return set()
    return set(json.loads(STATE_PATH.read_text(encoding="utf-8")))


def save_state(imported: set[str]):
    STATE_PATH.write_text(json.dumps(sorted(imported), indent=2) + "\n", encoding="utf-8")


def ensure_project(config: dict, config_path: Path, access_token: str) -> int:
    if config.get("project_id"):
        return config["project_id"]
    label_config = (SCRIPT_DIR / "labeling_config.xml").read_text(encoding="utf-8")
    response = requests.post(
        f'{config["label_studio_url"].rstrip("/")}/api/projects',
        headers=api_headers(access_token),
        json={"title": config.get("project_title", "Shelf Count"), "label_config": label_config},
        timeout=30,
    )
    response.raise_for_status()
    project_id = response.json()["id"]
    config["project_id"] = project_id
    save_config(config_path, config)
    print(f"Created Label Studio project {project_id} (saved to {config_path})")
    return project_id


def load_manifest_rows(dataset_dir: Path) -> list[dict]:
    manifest_path = dataset_dir / "manifest.csv"
    if not manifest_path.exists():
        return []
    with open(manifest_path, newline="", encoding="utf-8") as handle:
        return list(csv.DictReader(handle))


def build_task(row: dict) -> dict:
    local_path = row["local_path"].replace("\\", "/")
    return {
        "data": {
            "image": f"/data/local-files/?d={local_path}",
            "customer_reference": row.get("customer_reference", ""),
            "date": row.get("date", ""),
            "priority": row.get("priority", ""),
        }
    }


def import_tasks(config: dict, project_id: int, tasks: list[dict], access_token: str):
    url = f'{config["label_studio_url"].rstrip("/")}/api/projects/{project_id}/import'
    batch_size = 200
    for start in range(0, len(tasks), batch_size):
        batch = tasks[start:start + batch_size]
        response = requests.post(url, headers=api_headers(access_token), json=batch, timeout=120)
        response.raise_for_status()
        result = response.json()
        task_count = result.get("task_count")
        if task_count is not None and task_count != len(batch):
            print(f"  WARNING: sent {len(batch)} task(s) but Label Studio reports task_count={task_count}")
            print(f"  Full response: {json.dumps(result, indent=2)[:2000]}")
        else:
            print(f"  imported {start + len(batch)}/{len(tasks)} (response: {json.dumps(result)[:300]})")


def main():
    parser = argparse.ArgumentParser(description="Step 2: import dataset photos into Label Studio")
    parser.add_argument("--config", default=str(SCRIPT_DIR / "config.json"))
    args = parser.parse_args()

    config_path = Path(args.config)
    config = load_config(config_path)
    dataset_dir = (SCRIPT_DIR / config["dataset_dir"]).resolve()

    rows = load_manifest_rows(dataset_dir)
    if not rows:
        print(f"No rows found in {dataset_dir / 'manifest.csv'} -- run data/collect.py first.")
        return

    imported = load_state()
    new_rows = [row for row in rows if row["local_path"] not in imported]
    if not new_rows:
        print("Nothing new to import -- every manifest row has already been sent to Label Studio.")
        return

    access_token = get_access_token(config)

    project_id = ensure_project(config, config_path, access_token)
    print(f"Importing {len(new_rows)} new task(s) into project {project_id}...")
    import_tasks(config, project_id, [build_task(row) for row in new_rows], access_token)

    imported.update(row["local_path"] for row in new_rows)
    save_state(imported)
    print(f"\nDone. {len(new_rows)} task(s) imported, {len(imported)} total tracked as imported.")


if __name__ == "__main__":
    main()
