"""
Step 4: train the shelf-level YOLO detector on the dataset prepare_dataset.py
built.

Thin wrapper around Ultralytics -- all the actual training logic lives in
that library. This script's job is just: point it at data.yaml, run with
this project's config, and leave a dated summary behind (metrics + which
base model/dataset produced it) for promote_model.py (Step 6) and
evaluate.py (Step 5) to read.

Usage:
    python train.py --config config.json --data dataset/data.yaml
"""
from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path

from ultralytics import YOLO

SCRIPT_DIR = Path(__file__).resolve().parent
RUNS_DIR = SCRIPT_DIR / "runs"


def load_config(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def main():
    parser = argparse.ArgumentParser(description="Step 4: train the shelf-level detector")
    parser.add_argument("--config", default=str(SCRIPT_DIR / "config.json"))
    parser.add_argument("--data", default=str(SCRIPT_DIR / "dataset" / "data.yaml"))
    parser.add_argument("--name", default=None, help="Run name (default: shelf_level_<timestamp>)")
    args = parser.parse_args()

    config = load_config(Path(args.config))
    data_yaml = Path(args.data)
    if not data_yaml.exists():
        raise SystemExit(f"{data_yaml} not found -- run prepare_dataset.py first.")

    run_name = args.name or f'shelf_level_{datetime.now().strftime("%Y%m%d_%H%M%S")}'

    # Any other key in config.json (workers, patience, cos_lr, mosaic,
    # close_mosaic, ...) is passed straight to Ultralytics -- these used to
    # be silently ignored. Unknown names are rejected by Ultralytics itself.
    core_keys = {"base_model", "epochs", "imgsz", "batch", "device", "seed"}
    extra_args = {key: value for key, value in config.items() if key not in core_keys and not key.startswith("_")}
    if extra_args:
        print(f"Extra training settings from config.json: {extra_args}")

    model = YOLO(config["base_model"])
    model.train(
        data=str(data_yaml),
        epochs=config["epochs"],
        imgsz=config["imgsz"],
        batch=config["batch"],
        device=config["device"],
        seed=config["seed"],
        project=str(RUNS_DIR),
        name=run_name,
        **extra_args,
    )

    run_dir = RUNS_DIR / run_name
    best_weights = run_dir / "weights" / "best.pt"
    if not best_weights.exists():
        raise SystemExit(f"Training finished but {best_weights} is missing -- check the Ultralytics log above.")

    print(f"\nValidating {best_weights}...")
    trained = YOLO(str(best_weights))
    metrics = trained.val(data=str(data_yaml), imgsz=config["imgsz"], device=config["device"])

    summary = {
        "run_name": run_name,
        "base_model": config["base_model"],
        "epochs": config["epochs"],
        "imgsz": config["imgsz"],
        "weights_path": str(best_weights),
        "map50": float(metrics.box.map50),
        "map50_95": float(metrics.box.map),
        "precision": float(metrics.box.mp),
        "recall": float(metrics.box.mr),
        "trained_at": datetime.now().isoformat(timespec="seconds"),
    }
    (run_dir / "summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")

    print(f"\nDone. Run: {run_dir}")
    print(f"  mAP50={summary['map50']:.3f}  mAP50-95={summary['map50_95']:.3f}  precision={summary['precision']:.3f}  recall={summary['recall']:.3f}")
    print(f"  Next: python evaluate.py --run {run_name}   (counting-accuracy report)")
    print(f"        python promote_model.py --run {run_name}   (register it in models/, once you're happy with it)")


if __name__ == "__main__":
    main()
