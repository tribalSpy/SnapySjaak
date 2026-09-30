"""
Step 5: counting-accuracy report for a trained run.

box mAP (what Ultralytics reports during training) says how well the model
localizes shelf levels and extensions -- it doesn't directly say how often
the resulting *count* is right, which is what the nightly pipeline actually
needs. This re-runs the val split through the model and compares predicted
box counts against ground-truth box counts per image, per class (shelf_level
and extension are unrelated quantities -- a combined "total boxes" number
would not mean anything useful, so each gets its own accuracy numbers).

Usage:
    python evaluate.py --run shelf_level_20260910_120000
"""
from __future__ import annotations

import argparse
import csv
import json
import statistics
import sys
from datetime import datetime
from pathlib import Path

from ultralytics import YOLO

SCRIPT_DIR = Path(__file__).resolve().parent
TRAINING_DIR = SCRIPT_DIR.parent / "training"
REPORTS_DIR = SCRIPT_DIR / "reports"

sys.path.insert(0, str(TRAINING_DIR))
from prepare_dataset import CLASS_NAMES  # noqa: E402 -- single source of truth for class order


def ground_truth_counts(label_path: Path) -> dict[str, int]:
    counts = {name: 0 for name in CLASS_NAMES}
    if not label_path.exists():
        return counts
    text = label_path.read_text(encoding="utf-8").strip()
    for line in text.splitlines():
        class_id = int(line.split()[0])
        counts[CLASS_NAMES[class_id]] += 1
    return counts


def predicted_counts(boxes) -> dict[str, int]:
    counts = {name: 0 for name in CLASS_NAMES}
    if boxes is None:
        return counts
    for class_id in boxes.cls.tolist():
        counts[CLASS_NAMES[int(class_id)]] += 1
    return counts


def class_summary(rows: list[dict], class_name: str) -> dict:
    abs_diffs = [row[f"{class_name}_abs_diff"] for row in rows]
    return {
        "exact_match_rate": sum(row[f"{class_name}_exact_match"] for row in rows) / len(rows),
        "mean_absolute_error": statistics.mean(abs_diffs),
        "rmse": statistics.mean(d ** 2 for d in abs_diffs) ** 0.5,
        "max_abs_diff": max(abs_diffs),
    }


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
        gt_counts = ground_truth_counts(val_labels_dir / (image_path.stem + ".txt"))
        results = model.predict(str(image_path), conf=args.conf, verbose=False)
        pred_counts = predicted_counts(results[0].boxes)
        # Filenames are "<source_folder>__<original_name>" (see prepare_dataset.py)
        source_folder = image_path.stem.split("__", 1)[0]
        row = {"image": image_path.name, "source_folder": source_folder}
        for class_name in CLASS_NAMES:
            gt = gt_counts[class_name]
            pred = pred_counts[class_name]
            row[f"{class_name}_gt"] = gt
            row[f"{class_name}_pred"] = pred
            row[f"{class_name}_abs_diff"] = abs(gt - pred)
            row[f"{class_name}_exact_match"] = gt == pred
        rows.append(row)

    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    fieldnames = ["image", "source_folder"] + [
        f"{class_name}_{suffix}" for class_name in CLASS_NAMES for suffix in ("gt", "pred", "abs_diff", "exact_match")
    ]
    csv_path = REPORTS_DIR / f"eval_{args.run}_{timestamp}.csv"
    with open(csv_path, "w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    summary = {
        "run": args.run,
        "conf_threshold": args.conf,
        "image_count": len(rows),
        "evaluated_at": datetime.now().isoformat(timespec="seconds"),
        "by_class": {class_name: class_summary(rows, class_name) for class_name in CLASS_NAMES},
    }
    summary_path = REPORTS_DIR / f"eval_{args.run}_{timestamp}_summary.json"
    summary_path.write_text(json.dumps(summary, indent=2), encoding="utf-8")

    print(f"Evaluated {len(rows)} val image(s):")
    for class_name in CLASS_NAMES:
        stats = summary["by_class"][class_name]
        print(f"  {class_name}: exact match {stats['exact_match_rate']:.1%}, MAE {stats['mean_absolute_error']:.2f}, "
              f"rmse {stats['rmse']:.2f}, max diff {stats['max_abs_diff']}")
    print(f"\nPer-image report: {csv_path}\nSummary: {summary_path}")


if __name__ == "__main__":
    main()
