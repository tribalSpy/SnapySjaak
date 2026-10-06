"""
Inference helper for the active shelf-level model.

Used by shelf-poller-app in SIDE-BY-SIDE mode: the live nightly count still
comes from the Ollama vision-language model; this model's counts are sent
along next to it (shown as "Trained model" on the Shelf count page) so its
accuracy can be judged on real nights before anything switches over. The
poller calls this file with the training venv's python (it never imports
ultralytics itself), and silently skips it while no model is promoted.

Counts are per class -- shelf_level and extension are unrelated quantities
(see training/prepare_dataset.py CLASS_NAMES and evaluation/evaluate.py),
never one combined box count.

Usage as a library:
    from infer import count_photo
    count_photo("photo.jpg")  # {"shelf_level": 4, "extension": 0, "confidence": 0.83}

CLI, one photo (quick manual check):
    python infer.py photo.jpg
CLI, several photos as one JSON document (what shelf-poller-app runs):
    python infer.py --batch photo1.jpg photo2.jpg ...
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

MODELS_DIR = Path(__file__).resolve().parent
REGISTRY_PATH = MODELS_DIR / "registry.json"
sys.path.insert(0, str(MODELS_DIR.parent / "training"))

_model_cache = {}


def class_names() -> list[str]:
    try:
        from prepare_dataset import CLASS_NAMES  # single source of truth for class order
        return list(CLASS_NAMES)
    except Exception:  # noqa: BLE001 -- standalone copy of models/ without training/
        return ["shelf_level", "extension"]


def active_entry() -> dict:
    if not REGISTRY_PATH.exists():
        raise RuntimeError(f"{REGISTRY_PATH} doesn't exist -- run training/promote_model.py first.")
    registry = json.loads(REGISTRY_PATH.read_text(encoding="utf-8"))
    active = next((entry for entry in registry if entry.get("active")), None)
    if not active:
        raise RuntimeError("No active model in registry.json -- run training/promote_model.py first.")
    return active


def active_weights_path() -> Path:
    return MODELS_DIR / active_entry()["weights_path"]


def load_model():
    from ultralytics import YOLO  # imported lazily -- registry-only callers don't need torch

    weights_path = active_weights_path()
    if weights_path not in _model_cache:
        _model_cache[weights_path] = YOLO(str(weights_path))
    return _model_cache[weights_path]


def count_photo(image_path, conf: float = 0.25) -> dict:
    """Runs the active model on one photo: a count per class plus a
    confidence (mean of all box confidences, 0.0 if nothing detected)."""
    names = class_names()
    # Same photo size the active model was trained at (registry, from
    # promote_model.py); older entries without it fall back to the model's own default.
    imgsz = active_entry().get("imgsz")
    predict_args = {"conf": conf, "verbose": False}
    if imgsz:
        predict_args["imgsz"] = int(imgsz)
    from postprocess import boxes_from_result, clean_boxes, count_classes

    result = load_model().predict(str(image_path), **predict_args)[0]
    # Same clean-up as evaluation and pre-labeling (models/postprocess.py):
    # drops a shelf drawn twice a few pixels apart, and too-narrow shelves.
    boxes = clean_boxes(boxes_from_result(result, names), result.orig_shape[0], image_width=result.orig_shape[1])
    counts = count_classes(boxes, names)
    if not boxes:
        return {**counts, "confidence": 0.0}
    return {**counts, "confidence": sum(box["conf"] for box in boxes) / len(boxes)}


# Backwards-compatible name from the first version of this helper.
def count_shelf_levels(image_path, conf: float = 0.25) -> dict:
    result = count_photo(image_path, conf)
    return {"shelf_count": result.get("shelf_level", 0), "confidence": result["confidence"]}


def count_batch(image_paths) -> dict:
    entry = active_entry()
    photos = []
    for image_path in image_paths:
        try:
            photos.append({"file": Path(image_path).name, **count_photo(image_path)})
        except Exception as error:  # noqa: BLE001 -- one unreadable photo must not drop the rest
            photos.append({"file": Path(image_path).name, "error": str(error)})
    return {
        "model_version": entry.get("version") or entry.get("name") or entry.get("weights_path", ""),
        "photos": photos,
    }


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "--batch":
        print(json.dumps(count_batch(sys.argv[2:])))
    elif len(sys.argv) == 2:
        print(json.dumps(count_photo(sys.argv[1]), indent=2))
    else:
        raise SystemExit("Usage: python infer.py <image_path>  |  python infer.py --batch <image> [<image> ...]")
