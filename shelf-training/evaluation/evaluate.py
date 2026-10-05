"""
Step 5: counting-accuracy report for a trained run.

box mAP (what Ultralytics reports during training) says how well the model
localizes shelf levels and extensions -- it doesn't directly say how often
the resulting *count* is right, which is what the nightly pipeline actually
needs. This re-runs the val split through the model and compares predicted
box counts against ground-truth box counts per image, per class (shelf_level
and extension are unrelated quantities -- a combined "total boxes" number
would not mean anything useful, so each gets its own accuracy numbers).

Naming note: a "shelf_level" box marks every visible horizontal tier,
including the trolley's bottom/base tier -- which is a level but never a
countable shelf (confirmed always true, no trolley exception). So the
shelf_level class here is really a LEVEL count; the actual shelf count is
derived below as shelf_level_count - 1, floored at 0 for an empty/undetected
photo.

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
sys.path.insert(0, str(TRAINING_DIR.parent / "models"))
from postprocess import boxes_from_result, clean_boxes, count_classes  # noqa: E402 -- same clean-up as the live count


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


def class_summary(rows: list[dict], gt_field: str, pred_field: str, abs_diff_field: str, exact_match_field: str) -> dict:
    abs_diffs = [row[abs_diff_field] for row in rows]
    return {
        "exact_match_rate": sum(row[exact_match_field] for row in rows) / len(rows),
        "mean_absolute_error": statistics.mean(abs_diffs),
        "rmse": statistics.mean(d ** 2 for d in abs_diffs) ** 0.5,
        "max_abs_diff": max(abs_diffs),
    }


def add_derived_shelf_count(row: dict):
    """shelf_level counts levels (every horizontal tier, including the
    trolley's own base) -- the actual shelf count excludes that always-non-
    shelf bottom tier."""
    row["derived_shelf_count_gt"] = max(0, row["shelf_level_gt"] - 1)
    row["derived_shelf_count_pred"] = max(0, row["shelf_level_pred"] - 1)
    row["derived_shelf_count_abs_diff"] = abs(row["derived_shelf_count_gt"] - row["derived_shelf_count_pred"])
    row["derived_shelf_count_exact_match"] = row["derived_shelf_count_gt"] == row["derived_shelf_count_pred"]


def main():
    parser = argparse.ArgumentParser(description="Step 5: counting-accuracy report")
    parser.add_argument("--run", required=True, help="Run name, as printed by train.py")
    parser.add_argument("--dataset", default=str(TRAINING_DIR / "dataset"), help="Dataset dir from prepare_dataset.py")
    parser.add_argument("--conf", type=float, default=0.25)
    parser.add_argument("--imgsz", type=int, default=0, help="Photo size (default: the size the run was trained at)")
    parser.add_argument("--raw", action="store_true", help="Count the model's raw boxes, without the clean-up (models/postprocess.py)")
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

    # Evaluate at the size the run was trained at (summary.json), not
    # Ultralytics' 640 default.
    imgsz = args.imgsz
    if not imgsz:
        summary_path = TRAINING_DIR / "runs" / args.run / "summary.json"
        if summary_path.exists():
            imgsz = int(json.loads(summary_path.read_text(encoding="utf-8")).get("imgsz") or 0)
    imgsz = imgsz or 640
    print(f"Evaluating at imgsz={imgsz}, {'raw boxes' if args.raw else 'with clean-up (models/postprocess.py)'}")

    model = YOLO(str(weights_path))
    rows = []
    for image_path in image_paths:
        gt_counts = ground_truth_counts(val_labels_dir / (image_path.stem + ".txt"))
        results = model.predict(str(image_path), conf=args.conf, imgsz=imgsz, verbose=False)
        if args.raw:
            pred_counts = predicted_counts(results[0].boxes)
        else:
            # Same clean-up the live count and pre-labeling use.
            cleaned = clean_boxes(boxes_from_result(results[0], CLASS_NAMES), results[0].orig_shape[0])
            pred_counts = count_classes(cleaned, CLASS_NAMES)
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
        add_derived_shelf_count(row)
        rows.append(row)

    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_fields = list(CLASS_NAMES) + ["derived_shelf_count"]
    fieldnames = ["image", "source_folder"] + [
        f"{field}_{suffix}" for field in report_fields for suffix in ("gt", "pred", "abs_diff", "exact_match")
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
        "by_class": {
            field: class_summary(rows, f"{field}_gt", f"{field}_pred", f"{field}_abs_diff", f"{field}_exact_match")
            for field in report_fields
        },
    }
    summary_path = REPORTS_DIR / f"eval_{args.run}_{timestamp}_summary.json"
    summary_path.write_text(json.dumps(summary, indent=2), encoding="utf-8")

    print(f"Evaluated {len(rows)} val image(s):")
    for class_name in report_fields:
        stats = summary["by_class"][class_name]
        print(f"  {class_name}: exact match {stats['exact_match_rate']:.1%}, MAE {stats['mean_absolute_error']:.2f}, "
              f"rmse {stats['rmse']:.2f}, max diff {stats['max_abs_diff']}")
    print(f"\nPer-image report: {csv_path}\nSummary: {summary_path}")


if __name__ == "__main__":
    main()
