"""Reading the Uniware window: capture, wait until it stops changing, OCR.

Uniware draws a fixed-width text screen, so Windows' own OCR (built into
Windows 10/11, no extra install) reads it well. Two corrections on top:
- fixed-width fonts make OCR split words ("Ve rkoop"): words are re-joined
  by their real pixel gap against the font's character width;
- the selected menu item is drawn in blue: each line reports whether it is
  highlighted, so the map can show what's selected.
"""
from __future__ import annotations

import asyncio
import ctypes
import hashlib
import io
import re
import time
from ctypes import wintypes
from statistics import median

import mss
from PIL import Image
from winsdk.windows.graphics.imaging import BitmapDecoder
from winsdk.windows.media.ocr import OcrEngine
from winsdk.windows.storage.streams import DataWriter, InMemoryRandomAccessStream

user32 = ctypes.windll.user32

# Screen coordinates in real pixels (otherwise a scaled display shifts the
# window rectangle against the screenshot).
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:  # noqa: BLE001 -- older Windows
    user32.SetProcessDPIAware()


def foreground_window() -> int:
    return user32.GetForegroundWindow()


def window_title(hwnd: int) -> str:
    length = user32.GetWindowTextLengthW(hwnd)
    buffer = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buffer, length + 1)
    return buffer.value


def window_rect(hwnd: int):
    rect = wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    return rect.left, rect.top, rect.right, rect.bottom


def capture_window(hwnd: int) -> Image.Image:
    left, top, right, bottom = window_rect(hwnd)
    with mss.mss() as grabber:
        shot = grabber.grab({"left": left, "top": top, "width": max(1, right - left), "height": max(1, bottom - top)})
    return Image.frombytes("RGB", shot.size, shot.rgb)


def image_fingerprint(image: Image.Image) -> str:
    small = image.convert("L").resize((160, 90))
    return hashlib.md5(small.tobytes()).hexdigest()


def capture_when_settled(hwnd: int, timeout: float = 3.0, interval: float = 0.15) -> Image.Image:
    """Capture repeatedly until two captures in a row are identical (the app
    finished drawing), or the timeout passes."""
    deadline = time.time() + timeout
    previous = capture_window(hwnd)
    previous_print = image_fingerprint(previous)
    while time.time() < deadline:
        time.sleep(interval)
        current = capture_window(hwnd)
        current_print = image_fingerprint(current)
        if current_print == previous_print:
            return current
        previous, previous_print = current, current_print
    return previous


async def _ocr(image: Image.Image):
    scaled = image.convert("RGB").resize((image.width * 2, image.height * 2))
    buffer = io.BytesIO()
    scaled.save(buffer, format="PNG")
    stream = InMemoryRandomAccessStream()
    writer = DataWriter(stream)
    writer.write_bytes(buffer.getvalue())
    await writer.store_async()
    stream.seek(0)
    decoder = await BitmapDecoder.create_async(stream)
    bitmap = await decoder.get_software_bitmap_async()
    engine = OcrEngine.try_create_from_user_profile_languages()
    result = await engine.recognize_async(bitmap)
    words = []
    for line_index, line in enumerate(result.lines):
        for word in line.words:
            rect = word.bounding_rect
            words.append({
                "line": line_index,
                "text": word.text,
                "x": rect.x / 2, "y": rect.y / 2, "w": rect.width / 2, "h": rect.height / 2,
            })
    return words


def _is_highlighted(image: Image.Image, box) -> bool:
    """Blue text (Uniware's selected item) in this box?"""
    left, top, right, bottom = (int(value) for value in box)
    region = image.crop((max(0, left), max(0, top), min(image.width, right), min(image.height, bottom))).convert("RGB")
    pixels = list(region.getdata())
    if not pixels:
        return False
    blue = sum(1 for r, g, b in pixels if b > r + 60 and b > g + 40)
    dark = sum(1 for r, g, b in pixels if max(r, g, b) < 120)
    return blue > max(8, dark * 0.5)


def read_screen(image: Image.Image) -> list[dict]:
    """The screen's text lines, top to bottom: {y, x, text, highlighted}."""
    words = asyncio.run(_ocr(image))
    if not words:
        return []
    char_width = median(word["w"] / max(1, len(word["text"])) for word in words)
    lines = {}
    for word in words:
        lines.setdefault(word["line"], []).append(word)
    result = []
    for line_words in lines.values():
        line_words.sort(key=lambda word: word["x"])
        # Split a visual line into separate fields where there is a wide gap
        # (menus often have two columns), and re-join words OCR split apart.
        fields = []
        current = None
        for word in line_words:
            if current is None:
                current = dict(word)
                continue
            gap = word["x"] - (current["x"] + current["w"])
            if gap > char_width * 3:
                fields.append(current)
                current = dict(word)
            else:
                joiner = "" if gap < char_width * 0.45 else " "
                current["text"] += joiner + word["text"]
                current["w"] = word["x"] + word["w"] - current["x"]
                current["h"] = max(current["h"], word["h"])
        fields.append(current)
        for field in fields:
            box = (field["x"], field["y"], field["x"] + field["w"], field["y"] + field["h"])
            result.append({
                "y": round(field["y"]),
                "x": round(field["x"]),
                "text": field["text"].strip(),
                "highlighted": _is_highlighted(image, box),
            })
    # OCR sometimes reports one visual line as several lines ("Uitstaand" +
    # "transportfust", "Prg." + "000"): same height and a small gap -> one.
    # Rows first (lines within a few pixels of height belong together), then
    # left to right within a row.
    result.sort(key=lambda line: line["y"])
    row_y = None
    for line in result:
        if row_y is None or line["y"] - row_y > 6:
            row_y = line["y"]
        line["_row"] = row_y
    result.sort(key=lambda line: (line["_row"], line["x"]))
    merged = []
    for line in result:
        line.pop("_row", None)
        previous = next((item for item in reversed(merged) if abs(item["y"] - line["y"]) <= 4), None)
        if previous is not None and 0 <= line["x"] - previous["_right"] < char_width * 3:
            previous["text"] += " " + line["text"]
            previous["_right"] = line["x"] + len(line["text"]) * char_width
            previous["highlighted"] = previous["highlighted"] or line["highlighted"]
            continue
        merged.append({**line, "_right": line["x"] + len(line["text"]) * char_width})
    for line in merged:
        line.pop("_right", None)
    merged.sort(key=lambda line: (line["y"] // 8, line["x"]))
    return merged


DATE_TIME_RE = re.compile(r"\b\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}\b|\b\d{1,2}:\d{2}(:\d{2})?\b|\[\d+\]|\b\d+\.\d+\.\d+\b")


def screen_identity(lines: list[dict], image_height: int) -> tuple[str, str]:
    """(id, title) for a screen: the header (top ~15%) without dates, times
    and version numbers -- e.g. "Hoofdmenu | Prg. 000"."""
    header = [line for line in lines if line["y"] < image_height * 0.15]
    parts = []
    for line in header:
        text = DATE_TIME_RE.sub("", line["text"]).strip(" -|")
        if text and text.lower() not in ("x",) and "uniware" not in text.lower():
            parts.append(text)
    title = " | ".join(parts) or "(no header)"
    key = re.sub(r"[^a-z0-9]+", "-", title.lower()).strip("-") or "screen"
    return key[:80], title
