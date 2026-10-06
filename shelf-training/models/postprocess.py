"""
Clean-up of the model's raw boxes, using what's physically true of a trolley
(the model doesn't know this):

2. Two shelves can never sit only a few pixels apart. The model sometimes
   draws the same shelf twice, the second copy slightly lower -- and
   YOLO's own duplicate removal (NMS) misses it, because two hairline boxes
   a few pixels apart barely overlap in area. Of shelf boxes whose vertical
   centres are closer than `merge_gap_pct` of the photo height and that
   overlap sideways, only the most confident one is kept.
1. There is only one shelf size. A shelf box much narrower than the other
   shelves in the same photo (< `min_width_ratio` of their median width) is
   an artifact, not a shelf.

Used by models/infer.py (live side-by-side count), evaluation/evaluate.py
and labeling/prelabel_tasks.py, so all three clean up identically.
"""
from __future__ import annotations

from statistics import median

SHELF_CLASS = "shelf_level"
EXTENSION_CLASS = "extension"
# Real shelves are always well apart -- at least 5-10 cm on the trolley
# (user-confirmed), ~3.5% of a 1920 px tall photo is ~67 px, about 7-8 cm.
DEFAULT_MERGE_GAP_PCT = 3.5   # % of photo height
# A pole carries at most one extension: extension boxes whose horizontal
# centres are this close (% of photo width) are on the same pole.
DEFAULT_POLE_GAP_PCT = 4.0
# Extensions always sit on TOP of a pole, never lower down (user-confirmed):
# an extension box's middle must be in the upper part of the trolley --
# above this fraction of the way from the highest to the lowest shelf found,
# or (no shelves to go by) in the upper part of the photo.
DEFAULT_EXTENSION_MAX_TROLLEY_FRACTION = 0.5
DEFAULT_EXTENSION_MAX_PHOTO_FRACTION = 0.6
DEFAULT_MIN_WIDTH_RATIO = 0.6  # of the median shelf width in the same photo


def boxes_from_result(result, class_names) -> list[dict]:
    """Ultralytics result -> [{"cls", "x1", "y1", "x2", "y2", "conf"}]."""
    boxes = result.boxes
    if boxes is None or len(boxes) == 0:
        return []
    out = []
    for xyxy, class_id, confidence in zip(boxes.xyxy.tolist(), boxes.cls.tolist(), boxes.conf.tolist()):
        index = int(class_id)
        if 0 <= index < len(class_names):
            x1, y1, x2, y2 = xyxy
            out.append({"cls": class_names[index], "x1": x1, "y1": y1, "x2": x2, "y2": y2, "conf": float(confidence)})
    return out


def _horizontal_overlap(a: dict, b: dict) -> float:
    overlap = min(a["x2"], b["x2"]) - max(a["x1"], b["x1"])
    narrower = min(a["x2"] - a["x1"], b["x2"] - b["x1"])
    return overlap / narrower if narrower > 0 and overlap > 0 else 0.0


def clean_extensions(boxes: list[dict], image_width: float, pole_gap_pct: float = DEFAULT_POLE_GAP_PCT) -> list[dict]:
    """One extension per pole: of extension boxes on the same pole (centres
    within pole_gap_pct of the photo width -- stacked or overlapping copies),
    only the most confident one is kept."""
    gap = image_width * pole_gap_pct / 100.0
    kept: list[dict] = []
    for box in sorted(boxes, key=lambda item: item["conf"], reverse=True):
        centre = (box["x1"] + box["x2"]) / 2
        if not any(abs(centre - (other["x1"] + other["x2"]) / 2) < gap for other in kept):
            kept.append(box)
    kept.sort(key=lambda item: item["x1"])
    return kept


def keep_top_extensions(extensions: list[dict], shelves: list[dict], image_height: float,
                        trolley_fraction: float = DEFAULT_EXTENSION_MAX_TROLLEY_FRACTION,
                        photo_fraction: float = DEFAULT_EXTENSION_MAX_PHOTO_FRACTION) -> list[dict]:
    """Drops "extensions" in the lower part of the trolley -- they only ever
    sit on top of a pole."""
    if len(shelves) >= 2:
        top = min((box["y1"] + box["y2"]) / 2 for box in shelves)
        bottom = max((box["y1"] + box["y2"]) / 2 for box in shelves)
        limit = top + (bottom - top) * trolley_fraction
    else:
        limit = image_height * photo_fraction
    return [box for box in extensions if (box["y1"] + box["y2"]) / 2 <= limit]


def clean_boxes(boxes: list[dict], image_height: float,
                merge_gap_pct: float = DEFAULT_MERGE_GAP_PCT,
                min_width_ratio: float = DEFAULT_MIN_WIDTH_RATIO,
                image_width: float | None = None,
                pole_gap_pct: float = DEFAULT_POLE_GAP_PCT) -> list[dict]:
    shelves = sorted((box for box in boxes if box["cls"] == SHELF_CLASS), key=lambda box: box["conf"], reverse=True)
    extensions = [box for box in boxes if box["cls"] == EXTENSION_CLASS]
    others = [box for box in boxes if box["cls"] not in (SHELF_CLASS, EXTENSION_CLASS)]
    extensions = keep_top_extensions(extensions, shelves, image_height)
    if image_width:
        extensions = clean_extensions(extensions, image_width, pole_gap_pct)

    # Too-narrow boxes go first: a small artifact (sometimes more confident
    # than the real shelf at the same height) must not block that shelf in
    # the duplicate check below and then be dropped itself, losing both.
    # Judged against the median width of all shelf boxes in the photo, only
    # when there's something to compare with.
    if min_width_ratio and len(shelves) >= 2:
        typical = median(box["x2"] - box["x1"] for box in shelves)
        shelves = [box for box in shelves if (box["x2"] - box["x1"]) >= typical * min_width_ratio]

    # Near-duplicates: same height (within the gap) + overlapping sideways;
    # the most confident one is kept.
    gap = image_height * merge_gap_pct / 100.0
    kept: list[dict] = []
    for box in shelves:
        centre = (box["y1"] + box["y2"]) / 2
        duplicate = any(
            abs(centre - (other["y1"] + other["y2"]) / 2) < gap and _horizontal_overlap(box, other) > 0.5
            for other in kept
        )
        if not duplicate:
            kept.append(box)

    kept.sort(key=lambda box: box["y1"])
    return kept + extensions + others


def count_classes(boxes: list[dict], class_names) -> dict:
    counts = {name: 0 for name in class_names}
    for box in boxes:
        counts[box["cls"]] = counts.get(box["cls"], 0) + 1
    return counts
