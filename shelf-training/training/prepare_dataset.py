"""
Step 3: turn Label Studio exports into a YOLO-format dataset.

Reads every labeling/exports/export_*.json (later files win for a given
task id, so re-running after more labeling always uses the latest box
positions), keeps only tasks with a real, non-skipped annotation, and
writes:

    training/dataset/images/{train,val}/<photo>.jpg
    training/dataset/labels/{train,val}/<photo>.txt   (YOLO format, see CLASS_NAMES below)
    training/dataset/data.yaml

The train/val split is stratified per customer_reference (same idea as
data/collect.py's stratified_sample) so one heavily-photographed customer
can't dominate the validation set, and a customer with only one labeled
photo keeps it in train rather than starving val of a near-useless single
example.

Usage:
    python prepare_dataset.py --raw-dataset-dir ../data/dataset --exports-dir ../labeling/exports
"""
from __future__ import annotations

import argparse
import json
import math
import random
import shutil
import sys
from collections import defaultdict
from pathlib import Path
from urllib.parse import parse_qs, urlparse

SCRIPT_DIR = Path(__file__).resolve().parent
# Must match labeling_config.xml's <Label> values exactly -- index here
# becomes the YOLO class id written into each label file.
CLASS_NAMES = ["shelf_level", "extension"]


def load_exports(exports_dir: Path) -> dict:
    """Later export files (sorted by filename, which is timestamped) win for
    a given Label Studio task id."""
    tasks_by_id = {}
    for export_path in sorted(exports_dir.glob("export_*.json")):
        tasks = json.loads(export_path.read_text(encoding="utf-8"))
        for task in tasks:
            tasks_by_id[task["id"]] = task
    return tasks_by_id


def image_relative_path(task: dict) -> str | None:
    """Task data.image looks like '/data/local-files/?d=raw/<folder>/<file>.jpg'."""
    image_url = task.get("data", {}).get("image", "")
    query = parse_qs(urlparse(image_url).query)
    values = query.get("d")
    return values[0] if values else None


def first_usable_annotation(task: dict) -> dict | None:
    for annotation in task.get("annotations", []):
        if not annotation.get("was_cancelled"):
            return annotation
    return None


def boxes_from_annotation(annotation: dict) -> list[tuple[int, float, float, float, float]]:
    """Label Studio gives x/y/width/height as percentages of the image, top-left
    origin. YOLO wants class + center x/y + width/height, all fractions of
    the image (0-1) -- the conversion is a straight /100 plus a center shift."""
    boxes = []
    for result in annotation.get("result", []):
        if result.get("type") != "rectanglelabels":
            continue
        value = result.get("value", {})
        labels = value.get("rectanglelabels") or []
        class_name = next((label for label in labels if label in CLASS_NAMES), None)
        if class_name is None:
            continue
        class_id = CLASS_NAMES.index(class_name)
        x = value["x"] / 100.0
        y = value["y"] / 100.0
        width = value["width"] / 100.0
        height = value["height"] / 100.0
        boxes.append((class_id, x + width / 2, y + height / 2, width, height))
    return boxes


def stratified_split(rows: list[dict], val_fraction: float, seed: int) -> tuple[list[dict], list[dict]]:
    rng = random.Random(seed)
    by_customer = defaultdict(list)
    for row in rows:
        by_customer[row["customer_reference"]].append(row)

    train, val = [], []
    for group in by_customer.values():
        rng.shuffle(group)
        if len(group) < 2:
            train.extend(group)
            continue
        val_count = max(1, math.floor(len(group) * val_fraction))
        val.extend(group[:val_count])
        train.extend(group[val_count:])
    return train, val


def write_split(rows: list[dict], split_name: str, out_dir: Path):
    images_dir = out_dir / "images" / split_name
    labels_dir = out_dir / "labels" / split_name
    images_dir.mkdir(parents=True, exist_ok=True)
    labels_dir.mkdir(parents=True, exist_ok=True)

    for row in rows:
        # Prefix with the source folder name -- two different customer/date
        # folders can otherwise hand back identically-named camera files
        # (e.g. "IMG_0007.jpg") and silently collide in the flat YOLO layout.
        unique_name = f"{row['image_path'].parent.name}__{row['image_path'].name}"
        dest_image = images_dir / unique_name
        shutil.copyfile(row["image_path"], dest_image)
        label_lines = [
            f"{cls} {cx:.6f} {cy:.6f} {w:.6f} {h:.6f}"
            for cls, cx, cy, w, h in row["boxes"]
        ]
        (labels_dir / (dest_image.stem + ".txt")).write_text(
            "\n".join(label_lines) + ("\n" if label_lines else ""), encoding="utf-8",
        )


def write_data_yaml(out_dir: Path):
    lines = [
        f"path: {out_dir.resolve().as_posix()}",
        "train: images/train",
        "val: images/val",
        "names:",
    ]
    for index, name in enumerate(CLASS_NAMES):
        lines.append(f"  {index}: {name}")
    (out_dir / "data.yaml").write_text("\n".join(lines) + "\n", encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description="Step 3: build a YOLO dataset from Label Studio exports")
    parser.add_argument("--raw-dataset-dir", default=str(SCRIPT_DIR / ".." / "data" / "dataset"))
    parser.add_argument("--exports-dir", default=str(SCRIPT_DIR / ".." / "labeling" / "exports"))
    parser.add_argument("--out", default=str(SCRIPT_DIR / "dataset"))
    parser.add_argument("--val-fraction", type=float, default=0.15)
    parser.add_argument("--seed", type=int, default=42)
    args = parser.parse_args()

    raw_dataset_dir = Path(args.raw_dataset_dir).resolve()
    exports_dir = Path(args.exports_dir).resolve()
    out_dir = Path(args.out).resolve()

    if not exports_dir.exists() or not any(exports_dir.glob("export_*.json")):
        print(f"No export_*.json files found in {exports_dir} -- run labeling/export_annotations.py first.")
        sys.exit(1)

    tasks_by_id = load_exports(exports_dir)
    print(f"Loaded {len(tasks_by_id)} task(s) from {exports_dir}")

    rows = []
    skipped_no_annotation = 0
    skipped_missing_image = 0
    empty_boxes = 0
    for task in tasks_by_id.values():
        annotation = first_usable_annotation(task)
        if annotation is None:
            skipped_no_annotation += 1
            continue
        relative_path = image_relative_path(task)
        if not relative_path:
            skipped_missing_image += 1
            continue
        image_path = raw_dataset_dir / relative_path
        if not image_path.exists():
            print(f"  missing image on disk, skipping: {image_path}")
            skipped_missing_image += 1
            continue
        boxes = boxes_from_annotation(annotation)
        if not boxes:
            empty_boxes += 1
        rows.append({
            "image_path": image_path,
            "boxes": boxes,
            "customer_reference": task.get("data", {}).get("customer_reference") or "unknown",
        })

    print(f"{len(rows)} labeled photo(s) usable ({empty_boxes} with zero boxes -- a photo with no visible shelf level, kept as a negative example)")
    print(f"Skipped: {skipped_no_annotation} not yet annotated/skipped in Label Studio, {skipped_missing_image} with a missing image file")

    if not rows:
        print("Nothing to write -- label some photos in Label Studio first.")
        sys.exit(1)

    train_rows, val_rows = stratified_split(rows, args.val_fraction, args.seed)
    print(f"Split: {len(train_rows)} train / {len(val_rows)} val")

    if out_dir.exists():
        shutil.rmtree(out_dir)
    write_split(train_rows, "train", out_dir)
    write_split(val_rows, "val", out_dir)
    write_data_yaml(out_dir)

    print(f"\nDone. Dataset written to {out_dir}\n  {out_dir / 'data.yaml'} is what training/train.py points Ultralytics at.")


if __name__ == "__main__":
    main()
