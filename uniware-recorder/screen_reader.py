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
from PIL import Image, ImageChops
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


def _print_window(hwnd: int):
    """The window's own content via PrintWindow -- correct even when another
    window (the recorder, the Start menu) lies on top of it. None if the
    window can't draw itself that way (then the screen area is used)."""
    left, top, right, bottom = window_rect(hwnd)
    width, height = right - left, bottom - top
    if width <= 0 or height <= 0:
        return None
    gdi32 = ctypes.windll.gdi32
    window_dc = user32.GetWindowDC(hwnd)
    memory_dc = gdi32.CreateCompatibleDC(window_dc)
    bitmap = gdi32.CreateCompatibleBitmap(window_dc, width, height)
    gdi32.SelectObject(memory_dc, bitmap)
    try:
        if not user32.PrintWindow(hwnd, memory_dc, 2):  # PW_RENDERFULLCONTENT
            return None

        class BITMAPINFOHEADER(ctypes.Structure):
            _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG), ("biHeight", wintypes.LONG),
                        ("biPlanes", wintypes.WORD), ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                        ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG), ("biYPelsPerMeter", wintypes.LONG),
                        ("biClrUsed", wintypes.DWORD), ("biClrImportant", wintypes.DWORD)]

        header = BITMAPINFOHEADER()
        header.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        header.biWidth, header.biHeight = width, -height  # top-down
        header.biPlanes, header.biBitCount = 1, 32
        buffer = ctypes.create_string_buffer(width * height * 4)
        gdi32.GetDIBits(memory_dc, bitmap, 0, height, buffer, ctypes.byref(header), 0)
        image = Image.frombuffer("RGBA", (width, height), buffer, "raw", "BGRA", 0, 1).convert("RGB")
        # A window that didn't actually draw gives an all-black image.
        return image if image.getbbox() else None
    finally:
        gdi32.DeleteObject(bitmap)
        gdi32.DeleteDC(memory_dc)
        user32.ReleaseDC(hwnd, window_dc)


def _crop_empty_margin(image: Image.Image) -> Image.Image:
    """With display scaling (or a maximised window on another screen)
    PrintWindow can return a bitmap twice the size of what the program
    drew, the rest pure black (never drawn) -- then the title / status bar
    aren't where the screen reading expects them. Keep just the drawn part."""
    drawn = image.getbbox()  # everything that isn't exactly black
    if not drawn:
        return image
    left, top, right, bottom = drawn
    if right - left > image.width * 0.92 and bottom - top > image.height * 0.92:
        return image
    return image.crop((0, 0, min(image.width, right + 2), min(image.height, bottom + 2)))


MIN_CAPTURE_WIDTH = 1700
NORMAL_CAPTURE_WIDTH = 1936  # the window size the map was recorded at


def _normal_size(image: Image.Image) -> Image.Image:
    """A smaller window has smaller text: "Prg. 000" then reads as "Prg. --"
    and the screen gets the wrong name. Enlarge it to the size the map was
    recorded at, so names and positions come out the same."""
    if image.width >= MIN_CAPTURE_WIDTH:
        return image
    scale = NORMAL_CAPTURE_WIDTH / image.width
    return image.resize((NORMAL_CAPTURE_WIDTH, round(image.height * scale)), Image.LANCZOS)


def capture_window(hwnd: int) -> Image.Image:
    image = _print_window(hwnd)
    if image is not None:
        return _normal_size(_crop_empty_margin(image))
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


def _cursor_mask(image: Image.Image) -> Image.Image:
    """Where the list cursor is: Uniware marks the selected row of a list
    with a teal background (0,153,153) -- menus use blue text instead."""
    red, green, blue = image.convert("RGB").split()
    return ImageChops.multiply(
        ImageChops.multiply(red.point(lambda value: 255 if value < 60 else 0), green.point(lambda value: 255 if value > 110 else 0)),
        blue.point(lambda value: 255 if value > 110 else 0),
    )


def _on_cursor(mask: Image.Image, box) -> bool:
    left, top, right, bottom = (int(value) for value in box)
    region = mask.crop((max(0, left), max(0, top), min(mask.width, right), min(mask.height, bottom)))
    histogram = region.histogram()
    total = sum(histogram)
    return bool(total) and histogram[255] > total * 0.3


def read_screen(image: Image.Image) -> list[dict]:
    """The screen's text lines, top to bottom: {y, x, text, highlighted}."""
    # Text on the teal cursor row reads badly ("COL C122" for "EC COL
    # EC1227"): read a copy with that background made grey.
    cursor = _cursor_mask(image)
    readable = image.convert("RGB").copy()
    readable.paste((192, 192, 192), mask=cursor)
    words = asyncio.run(_ocr(readable))
    for word in words:
        if _ZERO_WORD_RE.fullmatch(word["text"]):
            word["text"] = _fix_zeros(word["text"])
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
                current = {**word, "words": [word]}
                continue
            gap = word["x"] - (current["x"] + current["w"])
            if gap > char_width * 3:
                fields.append(current)
                current = {**word, "words": [word]}
            else:
                joiner = "" if gap < char_width * 0.45 else " "
                current["text"] += joiner + word["text"]
                current["w"] = word["x"] + word["w"] - current["x"]
                current["h"] = max(current["h"], word["h"])
                current["words"].append(word)
        fields.append(current)
        for field in fields:
            box = (field["x"], field["y"], field["x"] + field["w"], field["y"] + field["h"])
            result.append({
                "y": round(field["y"]),
                "x": round(field["x"]),
                "w": round(field["w"]),
                "text": field["text"].strip(),
                "highlighted": _is_highlighted(image, box),
                "cursor": _on_cursor(cursor, box),
                "words": field["words"],
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
            previous["w"] = max(previous.get("w", 0), line["x"] + line.get("w", 0) - previous["x"])
            previous["highlighted"] = previous["highlighted"] or line["highlighted"]
            previous["cursor"] = previous.get("cursor") or line.get("cursor")
            previous["words"] = previous["words"] + line["words"]
            continue
        merged.append({**line, "_right": line["x"] + len(line["text"]) * char_width})
    for line in merged:
        line.pop("_right", None)
    merged = _split_at_columns(image, merged, char_width, cursor)
    _read_missing_program_numbers(readable, merged, char_width)
    _read_gapped_numbers(readable, merged)
    for line in merged:
        line.pop("words", None)
    merged.sort(key=lambda line: (line["y"] // 8, line["x"]))
    return merged


# Uniware's terminal font draws zero with a dot/stroke inside; Windows' text
# recognition reads it as "Û" (or "Ø"). In a word that is otherwise digits
# ("ÛÛÛ", "2Û26"), that's a 0.
_ZERO_WORD_RE = re.compile(r"[0-9ÛØøOo]*[ÛØø][0-9ÛØøOo]*")


def _fix_zeros(text: str) -> str:
    return re.sub(r"[ÛØøOo]", "0", text)


def _read_missing_program_numbers(image: Image.Image, lines: list[dict], char_width: float):
    """Windows' text recognition sometimes drops a lone number such as the
    "000" in "Prg. 000" (the menu's screen name). Read the bit right of a
    bare "Prg." again, enlarged, and add the number to that line."""
    for line in lines:
        if not re.fullmatch(r"Prg\.?[^0-9A-Za-z]*", line["text"].strip(), re.IGNORECASE):  # "Prg." / "Prg. --"
            continue
        left = line["x"] + char_width * 3
        top = max(0, line["y"] - 6)
        box = (int(left), int(top), int(min(image.width, left + char_width * 8)), int(min(image.height, line["y"] + 30)))
        crop = image.crop(box)
        crop = crop.resize((crop.width * 3, crop.height * 3), Image.LANCZOS)
        words = asyncio.run(_ocr(crop))
        digits = "".join(re.sub(r"[^0-9]", "", _fix_zeros(word["text"]) if _ZERO_WORD_RE.fullmatch(word["text"]) else word["text"]) for word in words)
        if digits:
            line["text"] = f"Prg. {digits}"


def _read_gapped_numbers(image: Image.Image, lines: list[dict]):
    """Windows' text recognition sometimes loses a digit inside a number and
    leaves a gap: order 4841449 -> "4 41449". Read such a field again,
    enlarged; take it when the gaps came back as digits."""
    for line in lines:
        text = line["text"]
        if not re.fullmatch(r"\d+(?: \d+)+", text) or not 5 <= len(text) <= 10 or not line.get("w"):
            continue
        box = (int(max(0, line["x"] - 6)), int(max(0, line["y"] - 6)),
               int(min(image.width, line["x"] + line["w"] + 8)), int(min(image.height, line["y"] + 24)))
        crop = image.crop(box)
        crop = crop.resize((crop.width * 3, crop.height * 3), Image.LANCZOS)
        digits = "".join(re.sub(r"[^0-9]", "", word["text"]) for word in asyncio.run(_ocr(crop)))
        if len(digits) == len(text):
            line["text"] = digits


def _split_at_columns(image, lines: list[dict], char_width: float, cursor) -> list[dict]:
    """Menu levels sit side by side and the longest item of one level can end
    close to the next level ("Alternatieve vt_control" + "Factuur"), so OCR
    glues them into one line. A line holding a word that starts exactly where
    another line starts (the next level's column) is cut there, and each
    part gets its own highlight. (The font is fixed-width, so only starts of
    whole lines count -- not the second word of many items.)"""
    columns = sorted({line["x"] for line in lines})
    result = []
    for line in lines:
        words = sorted(line.get("words") or [], key=lambda word: word["x"])
        cuts = [index for index, word in enumerate(words) if index and abs(word["x"] - line["x"]) > 6
                and any(abs(word["x"] - column) <= 6 for column in columns)
                and word["x"] - (words[index - 1]["x"] + words[index - 1]["w"]) > char_width * 0.45]
        if not cuts or line["y"] < image.height * 0.15:  # never the title bar (screen names)
            result.append(line)
            continue
        for start, end in zip([0] + cuts, cuts + [len(words)]):
            part = words[start:end]
            left = part[0]["x"]
            right = max(word["x"] + word["w"] for word in part)
            top = min(word["y"] for word in part)
            bottom = max(word["y"] + word["h"] for word in part)
            text = ""
            for index, word in enumerate(part):
                gap = word["x"] - (part[index - 1]["x"] + part[index - 1]["w"]) if index else 0
                text += ("" if not index or gap < char_width * 0.45 else " ") + word["text"]
            result.append({"y": round(top), "x": round(left), "w": round(right - left), "text": text.strip(),
                           "highlighted": _is_highlighted(image, (left, top, right, bottom)),
                           "cursor": _on_cursor(cursor, (left, top, right, bottom))})
    return result


# The terminal program's own window parts (title bar, menu bar, status bar),
# not Uniware's screen.
CHROME_RE = re.compile(
    r"^(g\s*)?Connect\s*\(|Bestand\s+Edit\s+Beeld|Sessie\s*\d|Connect\s*mode|SlowScr|^Pauze$|ingelogd|^x$",
    re.IGNORECASE,
)


def strip_chrome(lines: list[dict]) -> list[dict]:
    return [line for line in lines if not CHROME_RE.search(line["text"])]


DATE_TIME_RE = re.compile(r"\b\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}\b|\b\d{1,2}:\d{2}(:\d{2})?\b|\[\d+\]|\b\d+\.\d+\.\d+\b")


PRG_RE = re.compile(r"Prg\.?\s*(\d{2,4})", re.IGNORECASE)
SCREEN_NUMBER_RE = re.compile(r"^(.*?)\s*\[?(\d{1,4})\]?$")


def screen_identity(lines: list[dict], image_height: int, image_width: int = 0) -> tuple[str, str]:
    """(id, title) for a screen.

    Uniware puts a stable number on every screen: menus show "Prg. 260" (top
    right) with the menu name under it -> "prg-260"; forms and lists show
    "<title> [38]" in their top bar -> "scr-38". Only if neither is found:
    the header text without dates/times/versions."""
    lines = strip_chrome(lines)
    width = image_width or max((line["x"] for line in lines), default=0) * 1.2 or 1
    top = [line for line in lines if line["y"] < image_height * 0.22]
    def title_below(line):
        below = [other for other in top if 8 < other["y"] - line["y"] < 50 and other["x"] > width * 0.45
                 and not re.search(r"Prg\.?", other["text"], re.IGNORECASE)]
        return below[0]["text"] if below else ""

    for line in top:
        match = PRG_RE.search(line["text"])
        if match:
            number = match.group(1)
            return f"prg-{number}", f"Prg. {number} {title_below(line) or ''}".strip()
    # "Prg." with its number read as a separate piece on the same row, or
    # (some screens) no number at all -- then the title names it.
    for line in top:
        if re.fullmatch(r"Prg\.?", line["text"].strip(), re.IGNORECASE):
            same_row = [other for other in top if abs(other["y"] - line["y"]) <= 6 and other["x"] > line["x"] and re.fullmatch(r"\d{2,4}", other["text"].strip())]
            title = title_below(line)
            if same_row:
                number = same_row[0]["text"].strip()
                return f"prg-{number}", f"Prg. {number} {title}".strip()
            if title:
                slug = re.sub(r"[^a-z0-9]+", "-", title.lower()).strip("-")[:40]
                return f"prg-{slug}", f"Prg. -- {title}"
    for line in top:
        match = SCREEN_NUMBER_RE.match(line["text"])
        if not match or line["x"] < width * 0.5:
            continue
        text, number = match.group(1).strip(" [|"), match.group(2)
        if not text:
            left = [other for other in top if abs(other["y"] - line["y"]) <= 6 and other["x"] < line["x"] and other["x"] > width * 0.45]
            text = left[-1]["text"] if left else ""
        if text and not DATE_TIME_RE.search(text):
            return f"scr-{number}", f"{text} [{number}]"
    header = [line for line in lines if line["y"] < image_height * 0.15]
    parts = []
    for line in header:
        text = DATE_TIME_RE.sub("", line["text"]).strip(" -|")
        if text and text.lower() not in ("x",) and "uniware" not in text.lower():
            parts.append(text)
    title = " | ".join(parts) or "(no header)"
    key = re.sub(r"[^a-z0-9]+", "-", title.lower()).strip("-") or "screen"
    return key[:80], title
