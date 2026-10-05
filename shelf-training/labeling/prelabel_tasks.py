"""
Pre-labeling: run a trained shelf model on every Label Studio task that has
no annotation yet, and push its boxes back as *predictions*. Opening such a
task in Label Studio then shows the model's boxes already drawn -- correct
them (move/resize/delete/add) and Submit, instead of drawing every box.

Predictions are never training data by themselves: only what you Submit
becomes an annotation, and only annotations are exported/trained on.

Usage:
    python prelabel_tasks.py                      (the active model in models/registry.json)
    python prelabel_tasks.py --run shelf_level_20261005_110950
    python prelabel_tasks.py --conf 0.3 --limit 50
    python prelabel_tasks.py --overwrite          (also tasks that already got predictions)
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import requests

from _auth import api_headers, get_access_token

SCRIPT_DIR = Path(__file__).resolve().parent
TRAINING_DIR = SCRIPT_DIR.parent / "training"
MODELS_DIR = SCRIPT_DIR.parent / "models"
sys.path.insert(0, str(TRAINING_DIR))
from prepare_dataset import CLASS_NAMES  # noqa: E402 -- same class order the model was trained with

# labeling_config.xml: <RectangleLabels name="label" toName="image">
FROM_NAME = "label"
TO_NAME = "image"


def load_config(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def resolve_model(run: str | None) -> tuple[Path, str, int]:
    """(weights, version name, imgsz) for --run, or else the active model."""
    if run:
        weights = TRAINING_DIR / "runs" / run / "weights" / "best.pt"
        summary_path = TRAINING_DIR / "runs" / run / "summary.json"
        imgsz = 0
        if summary_path.exists():
            imgsz = int(json.loads(summary_path.read_text(encoding="utf-8")).get("imgsz") or 0)
        if not weights.exists():
            raise SystemExit(f"{weights} not found -- has this run finished training?")
        return weights, run, imgsz or 640
    registry_path = MODELS_DIR / "registry.json"
    if not registry_path.exists():
        raise SystemExit("No promoted model yet -- pass --run <training run name>.")
    registry = json.loads(registry_path.read_text(encoding="utf-8"))
    active = next((entry for entry in registry if entry.get("active")), None)
    if not active:
        raise SystemExit("No active model in registry.json -- pass --run <training run name>.")
    return MODELS_DIR / active["weights_path"], active["version"], int(active.get("imgsz") or 640)


def list_tasks(base_url: str, project_id: int, access_token: str) -> list[dict]:
    tasks = []
    page = 1
    while True:
        response = requests.get(
            f"{base_url}/api/tasks",
            headers=api_headers(access_token),
            params={"project": project_id, "page": page, "page_size": 100},
            timeout=60,
        )
        if response.status_code == 404:
            break  # past the last page
        response.raise_for_status()
        payload = response.json()
        batch = payload.get("tasks", payload) if isinstance(payload, dict) else payload
        if not batch:
            break
        tasks.extend(batch)
        total = payload.get("total") if isinstance(payload, dict) else None
        if total is not None and len(tasks) >= total:
            break
        page += 1
    return tasks


def local_image_path(task: dict, dataset_dir: Path) -> Path | None:
    url = str((task.get("data") or {}).get("image") or "")
    relative = parse_qs(urlparse(url).query).get("d", [""])[0]
    if not relative:
        return None
    return dataset_dir / relative


def to_label_studio_result(result, min_shelf_height_pct: float) -> list[dict]:
    boxes = result.boxes
    if boxes is None or len(boxes) == 0:
        return []
    height, width = result.orig_shape[:2]
    items = []
    for xyxy, class_id, confidence in zip(boxes.xyxy.tolist(), boxes.cls.tolist(), boxes.conf.tolist()):
        index = int(class_id)
        if not 0 <= index < len(CLASS_NAMES):
            continue
        x1, y1, x2, y2 = xyxy
        box = {
            "x": max(0.0, x1 / width * 100),
            "y": max(0.0, y1 / height * 100),
            "width": min(100.0, (x2 - x1) / width * 100),
            "height": min(100.0, (y2 - y1) / height * 100),
        }
        # Optional: widen hairline shelf boxes to the agreed, thicker style
        # (shelf board + front edge), centred on what the model found.
        if CLASS_NAMES[index] == "shelf_level" and min_shelf_height_pct and box["height"] < min_shelf_height_pct:
            grow = min_shelf_height_pct - box["height"]
            box["y"] = max(0.0, box["y"] - grow / 2)
            box["height"] = min_shelf_height_pct
        items.append({
            "from_name": FROM_NAME,
            "to_name": TO_NAME,
            "type": "rectanglelabels",
            "original_width": width,
            "original_height": height,
            "image_rotation": 0,
            "value": {**box, "rotation": 0, "rectanglelabels": [CLASS_NAMES[index]]},
            "score": round(float(confidence), 3),
        })
    return items


def main():
    parser = argparse.ArgumentParser(description="Pre-label unlabeled Label Studio tasks with a trained model")
    parser.add_argument("--config", default=str(SCRIPT_DIR / "config.json"))
    parser.add_argument("--run", default=None, help="Training run to use (default: the active model)")
    parser.add_argument("--conf", type=float, default=0.25, help="Only boxes the model is at least this sure of")
    parser.add_argument("--limit", type=int, default=0, help="Stop after this many tasks (0 = all)")
    parser.add_argument("--overwrite", action="store_true", help="Also tasks that already have predictions")
    parser.add_argument("--min-shelf-height", type=float, default=0.0,
                        help="Make predicted shelf boxes at least this tall, in %% of the photo height (0 = as predicted)")
    args = parser.parse_args()

    config = load_config(Path(args.config))
    if not config.get("project_id"):
        raise SystemExit("No project_id in config.json -- run import_tasks.bat first.")
    base_url = config["label_studio_url"].rstrip("/")
    dataset_dir = (SCRIPT_DIR / config["dataset_dir"]).resolve()
    weights, version, imgsz = resolve_model(args.run)

    from ultralytics import YOLO  # lazily: torch is only needed once there's work to do

    access_token = get_access_token(config)
    tasks = list_tasks(base_url, config["project_id"], access_token)
    todo = [
        task for task in tasks
        if not task.get("total_annotations") and (args.overwrite or not task.get("total_predictions"))
    ]
    if args.limit:
        todo = todo[:args.limit]
    print(f"{len(tasks)} task(s) in project {config['project_id']}; pre-labeling {len(todo)} unlabeled one(s) "
          f"with {version} (imgsz={imgsz}, conf={args.conf})")
    if not todo:
        return

    model = YOLO(str(weights))
    done = 0
    for task in todo:
        image_path = local_image_path(task, dataset_dir)
        if not image_path or not image_path.exists():
            print(f"  task {task.get('id')}: photo not found ({image_path}) -- skipped")
            continue
        result = model.predict(str(image_path), conf=args.conf, imgsz=imgsz, verbose=False)[0]
        items = to_label_studio_result(result, args.min_shelf_height)
        scores = [item["score"] for item in items]
        body = {
            "task": task["id"],
            "model_version": version,
            "score": round(sum(scores) / len(scores), 3) if scores else 0.0,
            "result": items,
        }
        response = requests.post(f"{base_url}/api/predictions", headers=api_headers(access_token), json=body, timeout=60)
        if response.status_code == 401:
            # Label Studio's access tokens are short-lived (minutes); a long
            # run outlives one -- exchange the personal token again and retry.
            access_token = get_access_token(config)
            response = requests.post(f"{base_url}/api/predictions", headers=api_headers(access_token), json=body, timeout=60)
        if response.status_code >= 400:
            print(f"  task {task['id']}: Label Studio refused the prediction ({response.status_code}): {response.text[:300]}")
            continue
        done += 1
        shelves = sum(1 for item in items if item["value"]["rectanglelabels"] == ["shelf_level"])
        print(f"  task {task['id']}: {shelves} shelf box(es), {len(items) - shelves} extension box(es)")
    print(f"\nDone. Pre-labeled {done} task(s). Open them in Label Studio, correct the boxes and Submit.")


if __name__ == "__main__":
    main()
