"""
Step 6: register a trained run as a versioned model.

Copies a run's best.pt into models/ under a stable filename and records it
in models/registry.json (version, metrics, when it was trained/promoted,
which one is currently "active"). registry.json itself is small and
tracked in git so model history survives even though the .pt weight files
themselves are not (see models/README.md) -- they're moved between
machines by hand.

This does NOT wire the model into shelf-poller-app's live job handling.
That's a separate, deliberate decision (see models/README.md) since it
changes what a running production job actually does -- this script only
gets a model ready to be pointed at.

Usage:
    python promote_model.py --run shelf_level_20260910_120000
"""
from __future__ import annotations

import argparse
import json
import shutil
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
RUNS_DIR = SCRIPT_DIR / "runs"
MODELS_DIR = SCRIPT_DIR.parent / "models"
REGISTRY_PATH = MODELS_DIR / "registry.json"


def load_registry() -> list[dict]:
    if not REGISTRY_PATH.exists():
        return []
    return json.loads(REGISTRY_PATH.read_text(encoding="utf-8"))


def save_registry(registry: list[dict]):
    REGISTRY_PATH.write_text(json.dumps(registry, indent=2), encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description="Step 6: register a trained run in models/registry.json")
    parser.add_argument("--run", required=True, help="Run name, as printed by train.py")
    parser.add_argument("--no-activate", action="store_true", help="Register without making it the active model")
    args = parser.parse_args()

    run_dir = RUNS_DIR / args.run
    summary_path = run_dir / "summary.json"
    weights_path = run_dir / "weights" / "best.pt"
    if not summary_path.exists() or not weights_path.exists():
        raise SystemExit(f"{run_dir} doesn't look like a finished training run (missing summary.json or weights/best.pt).")

    summary = json.loads(summary_path.read_text(encoding="utf-8"))
    MODELS_DIR.mkdir(parents=True, exist_ok=True)
    dest_filename = f"{args.run}.pt"
    shutil.copyfile(weights_path, MODELS_DIR / dest_filename)

    registry = load_registry()
    registry = [entry for entry in registry if entry["version"] != args.run]
    activate = not args.no_activate
    if activate:
        for entry in registry:
            entry["active"] = False
    registry.append({
        "version": args.run,
        "weights_path": dest_filename,
        "base_model": summary.get("base_model"),
        # Photo size it was trained at -- inference must use the same size
        # (thin shelf edges get lost when a 1280/1600 model runs at 640).
        "imgsz": summary.get("imgsz"),
        "map50": summary.get("map50"),
        "map50_95": summary.get("map50_95"),
        "precision": summary.get("precision"),
        "recall": summary.get("recall"),
        "trained_at": summary.get("trained_at"),
        "promoted_at": datetime.now().isoformat(timespec="seconds"),
        "active": activate,
    })
    registry.sort(key=lambda entry: entry["promoted_at"])
    save_registry(registry)

    print(f"Registered {args.run} -> {MODELS_DIR / dest_filename}")
    print(f"  mAP50={summary.get('map50', 0):.3f}  active={activate}")
    print(f"  Registry: {REGISTRY_PATH}")


if __name__ == "__main__":
    main()
