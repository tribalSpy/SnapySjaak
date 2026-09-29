"""
Standalone inference helper for the active shelf-level model.

Deliberately not imported by shelf-poller-app -- the live nightly pipeline
still counts via the Ollama vision-language model, exactly as it did
before this training pipeline existed. This module exists so that
switching the live counting job over to the trained YOLO detector is a
later, separate, explicit decision (comparing accuracy via evaluate.py
first), not something this script forces to happen just by existing.

Usage as a library:
    from infer import count_shelf_levels
    result = count_shelf_levels("photo.jpg")  # {"shelf_count": 4, "confidence": 0.83}

Usage as a CLI (quick manual check):
    python infer.py photo.jpg
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

MODELS_DIR = Path(__file__).resolve().parent
REGISTRY_PATH = MODELS_DIR / "registry.json"

_model_cache = {}


def active_weights_path() -> Path:
    if not REGISTRY_PATH.exists():
        raise RuntimeError(f"{REGISTRY_PATH} doesn't exist -- run training/promote_model.py first.")
    registry = json.loads(REGISTRY_PATH.read_text(encoding="utf-8"))
    active = next((entry for entry in registry if entry.get("active")), None)
    if not active:
        raise RuntimeError("No active model in registry.json -- run training/promote_model.py first.")
    return MODELS_DIR / active["weights_path"]


def count_shelf_levels(image_path, conf: float = 0.25) -> dict:
    """Runs the active model on one photo, returns a count + a confidence
    (the mean of the individual box confidences, or 0.0 if nothing detected)."""
    from ultralytics import YOLO  # imported lazily -- callers that only need the registry helper shouldn't need torch installed

    weights_path = active_weights_path()
    if weights_path not in _model_cache:
        _model_cache[weights_path] = YOLO(str(weights_path))
    model = _model_cache[weights_path]

    results = model.predict(str(image_path), conf=conf, verbose=False)
    boxes = results[0].boxes
    if boxes is None or len(boxes) == 0:
        return {"shelf_count": 0, "confidence": 0.0}
    confidences = boxes.conf.tolist()
    return {"shelf_count": len(confidences), "confidence": sum(confidences) / len(confidences)}


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("Usage: python infer.py <image_path>")
    print(json.dumps(count_shelf_levels(sys.argv[1]), indent=2))
