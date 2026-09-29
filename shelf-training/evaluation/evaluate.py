"""
Step 5: counting-accuracy report for a trained run.

box mAP (what Ultralytics reports during training) says how well the model
localizes shelf levels -- it doesn't directly say how often the resulting
*count* is right, which is what the nightly pipeline actually needs. This
re-runs the val split through the model and compares predicted box count
against the ground-truth box count per image.

Usage:
    python evaluate.py --run shelf_level_20260910_120000
"""
from __future__ import annotations

import argparse
import csv
import json
import statistics
from datetime import datetime
from pathlib import Path

from ultralytics import YOLO

SCRIPT_DIR = Path(__file__).resolve().parent
TRAINING_DIR = SCRIPT_DIR.parent / "training"
REPORTS_DIR = SCRIPT_DIR / "reports"


def ground_truth_count(label_path: Path) -> int:
    if not label_path.exists():
        return 0
    text = label_path.read_text(encoding="utf-8").strip()
    return len(text.splitlines()) if text else 0


def main():
    parser = argparse.ArgumentParser(description="Step 5: counting-accuracy report")
    parser.add_argument("--run", required=True, help="Run name, as printed by train.py")
    parser.add_argument("--dataset", default=str(TRAINING_DIR / "dataset"), help="Dataset dir from prepare_dataset.py")
    parser.add_argument("--conf", type=float, default=0.25)
    args = parser.parse_args()

    weights_path = TRAINING_DIR / "runs" / args.run / "weights" / "best.pt"
    if not weights_path.exists():
        raise SystemExit(f"{weights_path} not found -- has this run finished training?")

    dataset_dir = Path(args.dataset)
    val_images_dir = dataset_dir / "images" / "val"
    val_labels_dir = dataset_dir / "labels" / "val"
    image_paths = sorted([p for p in val_images_dir.glob("*") if p.suffix.lower() in (".jpg", ".jpeg", ".png")])
    if not image_paths:
        raise SystemExit(f"No val images found in {val_images_dir}")

    model = YOLO(str(weights_path))
    rows = []
    for image_path in image_paths:
        gt_count = ground_truth_count(val_labels_dir / (image_path.stem + ".txt"))
        results = model.predict(str(image_path), conf=args.conf, verbose=False)
        boxes = results[0].boxes
        pred_count = 0 if boxes is None else len(boxes)
        # Filenames are "<source_folder>__<original_name>" (see prepare_dataset.py)
        source_folder = image_path.stem.split("__", 1)[0]
        rows.append({
            "image": image_path.name,
            "source_folder": source_folder,
            "gt_count": gt_count,
            "pred_count": pred_count,
            "abs_diff": abs(gt_count - pred_count),
            "exact_match": gt_count == pred_count,
        })

    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = REPORTS_DIR / f"eval_{args.run}_{timestamp}.csv"
    with open(csv_path, "w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=["image", "source_folder", "gt_count", "pred_count", "abs_diff", "exact_match"])
        writer.writeheader()
        writer.writerows(rows)

    abs_diffs = [row["abs_diff"] for row in rows]
    summary = {
        "run": args.run,
        "conf_threshold": args.conf,
        "image_count": len(rows),
        "exact_match_rate": sum(row["exact_match"] for row in rows) / len(rows),
        "mean_absolute_error": statistics.mean(abs_diffs),
        "rmse": statistics.mean(d ** 2 for d in abs_diffs) ** 0.5,
        "max_abs_diff": max(abs_diffs),
        "evaluated_at": datetime.now().isoformat(timespec="seconds"),
    }
    summary_path = REPORTS_DIR / f"eval_{args.run}_{timestamp}_summary.json"
    summary_path.write_text(json.dumps(summary, indent=2), encoding="utf-8")

    print(f"Evaluated {len(rows)} val image(s):")
    print(f"  exact match rate: {summary['exact_match_rate']:.1%}")
    print(f"  mean absolute error: {summary['mean_absolute_error']:.2f} shelf levels")
    print(f"  rmse: {summary['rmse']:.2f}, max diff: {summary['max_abs_diff']}")
    print(f"\nPer-image report: {csv_path}\nSummary: {summary_path}")


if __name__ == "__main__":
    main()
