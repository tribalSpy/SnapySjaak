"""Re-read every recorded screenshot with the current screen recognition and
rebuild data/map.json from it -- so improvements to how screens are
recognised also apply to sessions recorded before them.

The old map is kept as data/map.before-rebuild.json and the old
screenshots move to data/screens_old/.
"""
from __future__ import annotations

import json
import shutil

from PIL import Image

import recorder

DATA_DIR = recorder.DATA_DIR


def main():
    old = recorder.load_map()
    if not old["screens"]:
        print("Nothing recorded yet.")
        return
    (DATA_DIR / "map.before-rebuild.json").write_text(json.dumps(old, indent=2, ensure_ascii=False), encoding="utf-8")
    old_dir = DATA_DIR / "screens_old"
    if old_dir.exists():
        shutil.rmtree(old_dir)
    recorder.SCREENS_DIR.rename(old_dir)

    new = {"screens": {}, "transitions": {}}
    mapping = {}
    for old_id, entry in old["screens"].items():
        path = old_dir / f"{old_id}.png"
        if not path.exists():
            continue
        image = Image.open(path)
        screen_id, title, _selected, text, layout = recorder.describe_screen(image)
        mapping[old_id] = screen_id
        merged = recorder.remember_screen(new, screen_id, title, text, layout, image)
        merged["seen"] += max(0, int(entry.get("seen", 1)) - 1)
        merged["first_seen"] = min(merged.get("first_seen", ""), entry.get("first_seen", "")) or merged.get("first_seen", "")
        print(f"{old_id[:50]:50s} -> {screen_id}  {title}")

    for source, keys in old["transitions"].items():
        for key, targets in keys.items():
            for target, count in targets.items():
                if source not in mapping or target not in mapping:
                    continue
                bucket = new["transitions"].setdefault(mapping[source], {}).setdefault(key, {})
                bucket[mapping[target]] = bucket.get(mapping[target], 0) + count

    recorder.save_map(new)
    print(f"\n{len(old['screens'])} recorded screens -> {len(new['screens'])} distinct screens. Run make_report.bat to see the map.")


if __name__ == "__main__":
    main()
