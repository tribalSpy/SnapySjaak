"""Named Uniware actions: record once, replay from the dashboard.

An action is a list of steps, each one key (combination) or typed text plus
the screen Uniware should show afterwards (its Prg / [number] id, see
screen_reader.screen_identity). Replay sends the same keys to the Uniware
window and checks after every step that the expected screen is there --
anything else stops the run, so it never keeps typing into a wrong screen.

Steps can capture the screen (its text, or every page via Page Down) into
outputs/, and typed text can be turned into a field ("{invoice}") whose
value is asked when the action is started.
"""
from __future__ import annotations

import csv
import ctypes
import difflib
import datetime
import json
import os
import queue
import re
import subprocess
import threading
import time
from ctypes import wintypes
from pathlib import Path

from pynput import keyboard

import recorder
import screen_reader

APP_DIR = Path(__file__).resolve().parent
ACTIONS_DIR = APP_DIR / "actions"
OUTPUTS_DIR = APP_DIR / "outputs"
user32 = ctypes.windll.user32

# The Uniware window: "Connect (C Uniware Computer Systems BV) (Session 1: ...)".
DEFAULT_WINDOW_TITLE = "Uniware Computer Systems"
TYPING_GAP_SECONDS = 0.7  # typed characters closer together than this are one step
PAGE_WAIT_SECONDS = 8  # Page Down: no change within this time = end of the list


def default_max_wait(recorded_seconds: float) -> int:
    """How long a run waits for a step's screen: plenty more than it took
    while recording (Uniware can be slow), at least 30 s."""
    return int(min(600, max(30, recorded_seconds * 4)))


def slugify(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")[:60] or "action"


def list_actions() -> list[dict]:
    ACTIONS_DIR.mkdir(parents=True, exist_ok=True)
    actions = []
    for path in sorted(ACTIONS_DIR.glob("*.json")):
        try:
            actions.append(json.loads(path.read_text(encoding="utf-8")))
        except (OSError, json.JSONDecodeError):
            continue
    return actions


def load_action(slug: str) -> dict | None:
    path = ACTIONS_DIR / f"{slug}.json"
    return json.loads(path.read_text(encoding="utf-8")) if path.exists() else None


def save_action(action: dict):
    ACTIONS_DIR.mkdir(parents=True, exist_ok=True)
    path = ACTIONS_DIR / f"{action['slug']}.json"
    temp = path.with_suffix(".tmp")
    temp.write_text(json.dumps(action, indent=2, ensure_ascii=False), encoding="utf-8")
    temp.replace(path)


def action_fields(action: dict) -> list[str]:
    fields = []
    for step in action.get("steps", []):
        for name in re.findall(r"\{([a-z0-9_]+)\}", (step.get("text") or "") + (step.get("pick") or "")):
            if name not in fields:
                fields.append(name)
    return fields


# ------------------------------------------------------------------ menus
# Uniware's menus are levels side by side: the main menu on the left, each
# Right arrow opens the blue (selected) item's submenu as the next column,
# Left goes back a level. The level you're in is the rightmost column with a
# blue item (a column further right without one is just a preview of the
# selected item's submenu). So a menu route is the blue item per level --
# ["Expeditie", "Inklaring bedrijf"] -- and replay steers there by reading
# the screen, whatever the menu looked like when it started.
MENU_KEYS = {"Down", "Up", "Left", "Right", "Page Down", "Page Up", "Home", "End"}


def menu_levels(layout: list) -> list[dict]:
    """The columns of a menu screen, left to right: {x, items, selected}
    (selected = index of the blue item, or None)."""
    levels = []
    for line in sorted(layout, key=lambda line: line["x"]):
        if levels and line["x"] - levels[-1]["x"] <= 40:
            levels[-1]["lines"].append(line)
        else:
            levels.append({"x": line["x"], "lines": [line]})
    result = []
    for level in levels:
        lines = sorted(level["lines"], key=lambda line: line["y"])
        selected = [index for index, line in enumerate(lines) if line.get("highlighted")]
        result.append({"x": level["x"], "items": [line["text"] for line in lines], "selected": selected[0] if selected else None})
    return result


# The main menu's items: a real menu screen always shows them as its first
# column (a program screen such as Prg. 423 Transport afspraken has a Prg
# number and a long first column too, and was taken for a menu).
MAIN_MENU_ITEMS = ["Offertes", "Bestellingen", "Inkoop", "Verkoop", "Voorraad", "Productie", "Expeditie", "Fust administratie",
                   "Communicatie", "Onderhoud gegevens", "Afdrukken gegevens", "Financieel", "Systeem onderhoud"]


def is_menu(screen_id: str, layout: list) -> bool:
    """A menu screen: a Prg number and the main menu as the first column."""
    levels = menu_levels(layout)
    if not (screen_id.startswith("prg-") and levels and len(levels[0]["items"]) >= 8 and levels[0]["selected"] is not None):
        return False
    known = sum(1 for item in levels[0]["items"] if any(same_item(item, name) for name in MAIN_MENU_ITEMS))
    return known >= 6


def menu_path(layout: list) -> list[str]:
    """The blue item of every level up to the one you're in."""
    levels = menu_levels(layout)
    active = max((index for index, level in enumerate(levels) if level["selected"] is not None), default=-1)
    return [levels[index]["items"][levels[index]["selected"]] if levels[index]["selected"] is not None else "" for index in range(active + 1)]


def _screen_number(screen_id: str):
    match = re.fullmatch(r"(prg|scr)-(\d+)", screen_id or "")
    return (match.group(1), match.group(2)) if match else None


def same_screen(expected_id: str, expected_title: str, seen_id: str, seen_title: str) -> bool:
    """Is Uniware on the expected screen? The screen name comes from text
    recognition, which can misread a letter or the closing bracket
    ("Customsclearance cornpany [86]" / "company [861]"): same number, or
    one digit more at the end, plus a nearly identical title is the same
    screen. Different numbers otherwise never are (Prg. 000 vs 420)."""
    if expected_id == seen_id:
        return True
    letters = lambda text: re.sub(r"[^a-z]+", "", re.sub(r"\[[^\]]*\]?|prg\.?\s*\S+", "", (text or "").lower()))
    similar = difflib.SequenceMatcher(None, letters(expected_title), letters(seen_title)).ratio() >= 0.85
    expected_number, seen_number = _screen_number(expected_id), _screen_number(seen_id)
    if expected_number and seen_number:
        if expected_number[0] != seen_number[0]:
            return False
        short, long = sorted((expected_number[1], seen_number[1]), key=len)
        return similar and long.startswith(short) and len(long) - len(short) <= 1
    if similar and letters(expected_title):
        return True
    # One of the two names was read badly while recording, e.g. "o |
    # Customsclearance lines | Info | Filter Settings | ..." (the whole top
    # of the screen) for "Customsclearance lines [85]": the shorter title
    # inside the longer one is the same screen.
    short, long = sorted((letters(expected_title), letters(seen_title)), key=len)
    return len(short) >= 10 and short in long


def _plain(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", text.lower())


def same_item(seen: str, wanted: str) -> bool:
    """Menu text as read now vs as recorded -- OCR may differ a little
    ("vt_control" / "vt control", "virtuee" / "virtueel")."""
    left, right = _plain(seen), _plain(wanted)
    if not left or not right:
        return False
    if left == right or (min(len(left), len(right)) >= 5 and (left.startswith(right) or right.startswith(left))):
        return True
    return difflib.SequenceMatcher(None, left, right).ratio() >= 0.85


# ------------------------------------------------------------------ list loop
def list_loop_keys(value: str) -> list[str]:
    """ "Delete, Right, Enter" -> ["Delete", "Right", "Enter"] (a comma inside
    a combo like "Ctrl+," is not supported)."""
    return [part.strip() for part in str(value or "").split(",") if part.strip()]


def normalize_list_loop(loop: dict) -> dict:
    """Settings of a list loop step (see ActionRun._run_list_loop)."""
    words = lambda value: [word.strip() for word in (value if isinstance(value, list) else str(value or "").split(",")) if str(word).strip()]
    key = lambda value, default: str(value or "").strip() or default
    for name, default in (("match_key", "Delete"), ("next_key", "Down")):
        # Several keys in a row, comma-separated ("Delete, Right, Enter":
        # Delete, then answer Uniware's confirmation) -- each must be a key.
        for part in list_loop_keys(key(loop.get(name), default)):
            try:
                parse_combo(part)
            except (AttributeError, KeyError, ValueError):
                raise RuntimeError(f"List loop: '{part}' is not a key (e.g. Delete, Right, Enter, Down).")
    try:
        max_rows = max(1, min(2000, int(float(loop.get("max_rows") or 500))))
    except (TypeError, ValueError):
        max_rows = 500
    return {
        "match": words(loop.get("match")),
        "keep": words(loop.get("keep")),
        "match_key": key(loop.get("match_key"), "Delete"),
        "next_key": key(loop.get("next_key"), "Down"),
        "max_rows": max_rows,
        "dry_run": bool(loop.get("dry_run")),
        # The key doesn't take the row out of the list (e.g. Delete only
        # clears its tour number): move on with the next key afterwards.
        "row_stays": bool(loop.get("row_stays")),
        # Only rows whose text contains this (e.g. "508") are looked at --
        # a row already cleared is passed by (safe to run again).
        "only_rows_with": str(loop.get("only_rows_with") or "").strip(),
    }


# ------------------------------------------------------------------ lists
ARROW_KEYS = {"Down", "Up", "Page Down", "Page Up"}


def screen_rows(layout: list) -> list[dict]:
    """Text lines grouped into screen rows (top to bottom): {text, y,
    highlighted}. OCR splits a list row into one line per column."""
    rows = []
    for line in sorted(layout, key=lambda line: line["y"]):
        if rows and abs(line["y"] - rows[-1]["y"]) <= 8:
            rows[-1]["parts"].append(line)
            rows[-1]["highlighted"] = rows[-1]["highlighted"] or line.get("highlighted", False)
            rows[-1]["cursor"] = rows[-1]["cursor"] or line.get("cursor", False)
        else:
            rows.append({"y": line["y"], "parts": [line], "highlighted": line.get("highlighted", False), "cursor": line.get("cursor", False)})
    # The selected row of a list is the teal cursor row -- the one teal
    # across the most fields (a page counter "(1) 1-15 of 15" is teal too);
    # only a screen without one falls back to blue text (blue is also used
    # for column headers).
    for row in rows:
        row["cursor_parts"] = sum(1 for part in row["parts"] if part.get("cursor"))
    most = max((row["cursor_parts"] for row in rows), default=0)
    return [{"y": row["y"], "highlighted": (row["cursor_parts"] == most) if most else row["highlighted"],
             "text": " ".join(part["text"] for part in sorted(row["parts"], key=lambda part: part["x"]))} for row in rows]


# Windows' text recognition reads Uniware's digits as letters when a number
# stands on its own: "508" -> "soa", "515" -> "SIS", "55" -> "SS". When
# looking for a number, read those letters as the digits they resemble.
_DIGIT_LOOKALIKES = str.maketrans({"S": "5", "s": "5", "O": "0", "o": "0", "Û": "0", "Ø": "0", "D": "0", "a": "8", "B": "8",
                                   "I": "1", "l": "1", "i": "1", "|": "1", "Z": "2", "z": "2", "G": "6", "b": "6", "q": "9", "g": "9"})


def _as_digits(text: str, target: str) -> str:
    return text.translate(_DIGIT_LOOKALIKES) if target.strip().isdigit() else text


def value_matches(text: str, target: str) -> bool:
    """Does `text` show `target` as a whole value ("508" matches "508 GBR"
    and an OCR'd "soa", not "1508" or "50,80")?"""
    pattern = re.compile(rf"(?<![\w.,]){re.escape(target.strip())}(?![\w.,]|,\d|\.\d)", re.IGNORECASE)
    return bool(pattern.search(text) or pattern.search(_as_digits(text, target)))


def find_row(rows: list[dict], target: str) -> int | None:
    """Index of the row showing `target` as a whole value; a row that starts
    with it wins."""
    matches = [index for index, row in enumerate(rows) if value_matches(row["text"], target)]
    starting = [index for index in matches if _as_digits(rows[index]["text"].strip(), target).lower().startswith(target.strip().lower())]
    return (starting or matches or [None])[0]


def _words_with_x(line: dict) -> list[tuple[float, float, str]]:
    """(left, right, word) of every word in a line -- by character position,
    the font is fixed-width."""
    width = line.get("w") or len(line["text"]) * 14
    char = width / max(1, len(line["text"]))
    return [(line["x"] + match.start() * char, line["x"] + match.end() * char, match.group())
            for match in re.finditer(r"\S+", line["text"])]


def read_column(layout: list, header: str) -> list[str]:
    """The values under a table column, e.g. "invoiceRes", one per row
    (empty cells left out). The column runs from the end of the header
    before it to the start of the header after it (headers can be narrower
    than their column, right-aligned like "order"). A header cut to the
    column width ("invoiceRe") is found too."""
    wanted = header.strip().lower()
    best = None
    for line in layout:
        for left, right, word in _words_with_x(line):
            token = word.lower()
            if token == wanted or (len(token) >= 5 and wanted.startswith(token)):
                best = (line["y"], left, right)
                break
        if best:
            break
    if best is None:
        return []
    head_y, head_left, head_right = best
    header_words = sorted(word for line in layout if abs(line["y"] - head_y) <= 8 for word in _words_with_x(line))
    before = [right for left, right, _ in header_words if right <= head_left + 1]
    after = [left for left, right, _ in header_words if left >= head_right - 1]
    left = max(before) + 4 if before else head_left
    right = min(after) - 4 if after else head_right
    rows = {}
    for line in layout:
        if line["y"] <= head_y + 8:
            continue
        for word_left, word_right, word in _words_with_x(line):
            middle = (word_left + word_right) / 2
            if left <= middle <= right:
                rows.setdefault(round(line["y"] / 8), []).append((word_left, word))
    return [" ".join(word for _, word in sorted(rows[key])) for key in sorted(rows)]


def clean_number(value: str) -> str:
    """An OCR'd number with look-alike letters ("483S236") -> digits, when
    that's what it is; otherwise unchanged."""
    fixed = value.translate(_DIGIT_LOOKALIKES).strip()
    return fixed if re.fullmatch(r"[0-9][0-9 ./-]*", fixed) and sum(ch.isdigit() for ch in value) >= len(value.strip()) / 2 else value.strip()


def _column_of(layout: list, anchor: dict) -> list[dict]:
    """The texts straight above / below `anchor` (overlapping it
    horizontally): the items of a popup list like the tours after F1."""
    left, right = anchor["x"], anchor["x"] + max(1, anchor.get("w") or len(anchor["text"]) * 14)
    column = []
    for line in layout:
        line_left, line_right = line["x"], line["x"] + max(1, line.get("w") or len(line["text"]) * 14)
        overlap = min(right, line_right) - max(left, line_left)
        if overlap >= 0.4 * min(right - left, line_right - line_left):
            column.append(line)
    items = []
    for line in sorted(column, key=lambda line: line["y"]):
        if items and abs(line["y"] - items[-1]["y"]) <= 8:
            items[-1]["text"] += " " + line["text"]
            items[-1]["highlighted"] = items[-1]["highlighted"] or bool(line.get("cursor"))
        else:
            items.append({"y": line["y"], "text": line["text"], "highlighted": bool(line.get("cursor"))})
    return items


def list_items(layout: list, target: str) -> list[dict]:
    """The list to move through, as rows {text, y, highlighted}. A popup
    list lies on top of a table (its own teal cell vs the table's teal row):
    then it's the column holding the value with a teal cell in it, or the
    column of a lone teal cell. Otherwise the screen's rows."""
    for line in layout:
        if value_matches(line["text"], target):
            column = _column_of(layout, line)
            if any(item["highlighted"] for item in column):
                return column
    cursor_rows = {}
    for line in layout:
        if line.get("cursor"):
            cursor_rows.setdefault(round(line["y"] / 8), []).append(line)
    lone = [lines[0] for lines in cursor_rows.values() if len(lines) == 1 and (lines[0].get("w") or 999) < 200]
    if lone and len(cursor_rows) > 1:
        return _column_of(layout, lone[0])
    return screen_rows(layout)


def row_key(row_text: str) -> str:
    """What to look for in a list row next time: its first number (the tour,
    e.g. "508"), else its first word."""
    number = re.search(r"(?<![\w.,])\d+(?![\w.,])", row_text)
    if number:
        return number.group()
    words = row_text.split()
    return words[0] if words else row_text


# ------------------------------------------------------------------ window
def find_uniware_window(title_part: str = DEFAULT_WINDOW_TITLE) -> int | None:
    found = []

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def callback(hwnd, _):
        if user32.IsWindowVisible(hwnd) and title_part.lower() in screen_reader.window_title(hwnd).lower():
            found.append(hwnd)
        return True

    user32.EnumWindows(callback, 0)
    return found[0] if found else None


def bring_to_front(hwnd: int) -> bool:
    if user32.IsIconic(hwnd):
        user32.ShowWindow(hwnd, 9)  # SW_RESTORE
    # Windows only lets a process take focus right after input; a harmless
    # Alt tap counts as that.
    controller = keyboard.Controller()
    controller.press(keyboard.Key.alt)
    controller.release(keyboard.Key.alt)
    user32.SetForegroundWindow(hwnd)
    time.sleep(0.4)
    return screen_reader.foreground_window() == hwnd


# ------------------------------------------------------------------ other windows
# Uniware's Ctrl+P (print) doesn't save a file where we want it: it opens the
# data in a new window of another program. The run waits for that window,
# finds the file it shows and copies it as the dump.
def window_pid(hwnd: int) -> int:
    pid = wintypes.DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return pid.value


def process_exe(pid: int) -> str:
    handle = ctypes.windll.kernel32.OpenProcess(0x1000, False, pid)  # PROCESS_QUERY_LIMITED_INFORMATION
    if not handle:
        return ""
    try:
        buffer = ctypes.create_unicode_buffer(1024)
        size = wintypes.DWORD(len(buffer))
        if ctypes.windll.kernel32.QueryFullProcessImageNameW(handle, 0, buffer, ctypes.byref(size)):
            return buffer.value
        return ""
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)


def visible_windows() -> dict:
    """hwnd -> (pid, title) of every visible top-level window with a title."""
    found = {}

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def callback(hwnd, _):
        if user32.IsWindowVisible(hwnd):
            title = screen_reader.window_title(hwnd)
            if title.strip():
                found[hwnd] = (window_pid(hwnd), title)
        return True

    user32.EnumWindows(callback, 0)
    return found


def process_command_line(pid: int) -> str:
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", f"(Get-CimInstance Win32_Process -Filter 'ProcessId={int(pid)}').CommandLine"],
            capture_output=True, text=True, timeout=15, creationflags=0x08000000,  # CREATE_NO_WINDOW
        )
        return result.stdout.strip()
    except (OSError, subprocess.SubprocessError):
        return ""


DUMP_SEARCH_DEPTH = 4


def _search_dirs(extra: list[str]) -> list[Path]:
    home = Path.home()
    dirs = [Path(path) for path in extra if path]
    for name in ("TEMP", "TMP", "OneDrive", "OneDriveCommercial", "OneDriveConsumer"):
        if os.environ.get(name):
            dirs.append(Path(os.environ[name]))
    dirs += [home / "Downloads", home / "Documents", home / "Desktop", home / "AppData" / "Local" / "Temp"]
    unique = []
    for folder in dirs:
        if folder.is_dir() and folder not in unique:
            unique.append(folder)
    return unique


def _find_named_file(name: str, since: float, folders: list[Path]) -> Path | None:
    wanted = name.lower()
    best = None
    for folder in folders:
        base_depth = len(folder.parts)
        for root, dirnames, filenames in os.walk(folder):
            if len(Path(root).parts) - base_depth >= DUMP_SEARCH_DEPTH:
                dirnames[:] = []
            for filename in filenames:
                if filename.lower() != wanted:
                    continue
                path = Path(root) / filename
                try:
                    modified = path.stat().st_mtime
                except OSError:
                    continue
                if modified >= since - 300 and (best is None or modified > best[0]):
                    best = (modified, path)
    return best[1] if best else None


def dump_file_of_window(pid: int, title: str, since: float, extra_dirs: list[str]) -> Path | None:
    """The file a just-opened window shows: from its program's command line
    (e.g. notepad.exe "C:\\...\\dump.filipe.csv"), else by the file name in
    its title ("dump.filipe....csv - Excel") searched in the usual folders."""
    command = process_command_line(pid)
    candidates = []
    for quoted, bare in re.findall(r'"([^"]+)"|(\S+)', command):
        text = quoted or bare
        path = Path(text)
        if path.suffix.lower() != ".exe" and path.is_file():
            candidates.append(path)
    if candidates:
        return max(candidates, key=lambda path: path.stat().st_mtime)
    name = re.sub(r"\s*\[[^\]]*\]\s*$", "", title.split(" - ")[0]).strip(" *")
    if not name:
        return None
    if Path(name).is_file():
        return Path(name)
    return _find_named_file(name, since, _search_dirs(extra_dirs))


# ------------------------------------------------------------------ keys
def parse_combo(combo: str):
    """"Shift+F5" -> ([Key.shift], Key.f5); "a" -> ([], "a")."""
    if len(combo) == 1:
        return [], combo
    parts = combo.split("+")
    if combo.endswith("+"):  # "Ctrl++": the key itself is "+"
        parts = parts[:-2] + ["+"]
    *modifier_names, base = parts
    modifiers = []
    for name in modifier_names:
        modifiers.append({"ctrl": keyboard.Key.ctrl, "alt": keyboard.Key.alt, "altgr": keyboard.Key.alt_gr, "shift": keyboard.Key.shift}[name.lower()])
    if len(base) == 1:
        if modifiers and base.isascii() and base.isalnum():
            # With Ctrl/Alt the physical key, not the character: pynput
            # types "P" as a Unicode character, which Ctrl doesn't combine
            # with -- Uniware never saw Ctrl+P / Ctrl+A.
            return modifiers, keyboard.KeyCode.from_vk(ord(base.upper()))
        return modifiers, base
    key_name = base.lower().replace(" ", "_")
    if key_name == "space":
        return modifiers, " "
    return modifiers, getattr(keyboard.Key, key_name)


def send_combo(controller: keyboard.Controller, combo: str):
    modifiers, base = parse_combo(combo)
    for modifier in modifiers:
        controller.press(modifier)
    controller.press(base)
    controller.release(base)
    for modifier in reversed(modifiers):
        controller.release(modifier)


# ------------------------------------------------------------------ recording
class ActionRecorder:
    """Records one named action from the user's own key presses in Uniware.
    Ctrl+Alt+C marks the current screen to be captured, Ctrl+Alt+Q stops."""

    def __init__(self, name: str, description: str = "", title_part: str = DEFAULT_WINDOW_TITLE):
        self.action = {
            "name": name.strip() or "New action",
            "slug": slugify(name),
            "description": description.strip(),
            "created": datetime.datetime.now().isoformat(timespec="seconds"),
            "start_screen": "",
            "steps": [],
        }
        self.title_part = title_part
        self.hwnd = None
        self.held = set()
        self.events: queue.Queue = queue.Queue()
        self.active = False
        self.status = ""
        self.selected = []
        self._listener = None
        self._thread = None
        self._open = None  # (step, pressed at) until its screen is known
        self.uniware_pid = 0
        self.known_windows: set = set()
        self.screen_id = ""
        self.layout: list = []
        self._typed: list[str] = []
        self._typed_at = 0.0
        self._last_printable_at = 0.0
        self._helper = recorder.Recorder.__new__(recorder.Recorder)  # for combo_name()
        self._helper.held = self.held

    def start(self):
        self.hwnd = find_uniware_window(self.title_part)
        if not self.hwnd:
            raise RuntimeError("Uniware window not found -- start Uniware (logged in) first.")
        bring_to_front(self.hwnd)
        image = screen_reader.capture_when_settled(self.hwnd)
        screen_id, title, selected, _, layout = recorder.describe_screen(image)
        self.screen_id, self.layout = screen_id, layout
        self.action["start_screen"] = screen_id
        self.action["start_title"] = title
        self.uniware_pid = window_pid(self.hwnd)
        self.known_windows = set(visible_windows())
        self.selected = selected
        self.active = True
        self.status = f"Recording -- start screen: {title}"
        self._listener = keyboard.Listener(on_press=self._on_press, on_release=self._on_release)
        self._listener.start()
        self._thread = threading.Thread(target=self._worker, daemon=True)
        self._thread.start()

    def stop(self) -> dict:
        self.events.put(("stop", None, time.time(), None))
        if self._thread:
            self._thread.join(timeout=10)
        if self._listener:
            self._listener.stop()
        self.active = False
        save_action(self.action)
        self.status = f"Saved: {self.action['name']} ({len(self.action['steps'])} steps)"
        return self.action

    def mark_capture(self):
        self.events.put(("capture", None, time.time(), None))

    def _on_press(self, key):
        if key in recorder.MODIFIERS:
            self.held.add(recorder.MODIFIERS[key])
            return
        vk = getattr(key, "vk", None)
        if {"Ctrl", "Alt"} <= self.held and vk in (0x43, 0x51):  # C, Q
            self.events.put(("capture" if vk == 0x43 else "stop", None, time.time(), None))
            return
        if not self.active or screen_reader.foreground_window() != self.hwnd:
            return
        name, printable = self._helper.combo_name(key)
        now = time.time()
        # The screen as it is when this key goes down is where the previous
        # step ended up: you only press the next key once Uniware has finished
        # loading, however long that took. (Inside a run of typed characters
        # no new picture is needed.)
        starts_something = not printable or now - self._last_printable_at > TYPING_GAP_SECONDS
        image = screen_reader.capture_window(self.hwnd) if starts_something else None
        self._last_printable_at = now if printable else 0.0
        self.events.put(("key", (name, printable), now, image))

    def _on_release(self, key):
        if key in recorder.MODIFIERS:
            self.held.discard(recorder.MODIFIERS[key])

    def _finish_open_step(self, image, at: float):
        """The open step (last key, or a run of typed text) ends on the screen
        in `image`; remember where it went and how long Uniware took."""
        screen_id, title, selected, _, layout = recorder.describe_screen(image)
        self.selected = selected
        self.screen_id = screen_id
        self.layout = layout
        if self._open is None:
            return
        step, started = self._open
        self._open = None
        waited = max(0.0, at - started)
        menu_before = step.pop("_menu_before", None)
        if menu_before and not is_menu(screen_id, layout):
            # An arrow that left the menu (Right into a program): the arrows
            # before it become one route by name, this key opens it.
            while self.action["steps"] and self.action["steps"][-1].get("key") in MENU_KEYS and self.action["steps"][-1].get("to_menu"):
                self.action["steps"].pop()
            step["menu"] = menu_before
        step.update({
            "to": screen_id, "to_title": title, "capture": False, "capture_pages": False,
            "recorded_wait": round(waited, 1),
            "to_menu": is_menu(screen_id, layout),
            "max_wait": default_max_wait(waited),
            "wait_text": "",
        })
        self.action["steps"].append(step)
        self.status = f"Step {len(self.action['steps'])}: {step.get('key') or 'type ' + repr(step['text'])} -> {title} (took {waited:.0f}s)"
        windows = visible_windows()
        opened = [(pid, window_title) for hwnd, (pid, window_title) in windows.items()
                  if hwnd not in self.known_windows and pid != self.uniware_pid]
        self.known_windows = set(windows)
        if opened and "key" in step:
            pid, window_title = opened[0]
            step["opens_window"] = {"title": window_title, "exe": Path(process_exe(pid)).name}
            self.status = f"Step {len(self.action['steps'])}: {step['key']} opened '{window_title}' -- the run takes that file as the dump"

    def _close_typing(self, at: float):
        if self._typed:
            typed = "".join(" " if name == "Space" else name for name in self._typed)
            self._open = ({"text": typed, "recorded_text": typed}, self._typed_at)
            self._typed = []

    def _worker(self):
        while True:
            kind, value, at, image = self.events.get()
            if kind in ("stop", "capture"):
                # Nothing is pressed now, so the screen you're looking at is
                # the finished one.
                self._close_typing(at)
                if self._open is not None:
                    self._finish_open_step(screen_reader.capture_when_settled(self.hwnd), at)
                if kind == "stop":
                    return
                if self.action["steps"]:
                    self.action["steps"][-1]["capture"] = True
                    self.status = f"Step {len(self.action['steps'])} will capture its screen"
                else:
                    self.action["capture_start"] = True
                    self.status = "The start screen will be captured"
                continue
            name, printable = value
            if image is not None:
                self._close_typing(at)
                self._finish_open_step(image, at)
            if printable:
                if not self._typed:
                    self._typed_at = at
                self._typed.append(name)
                continue
            step = {"key": name}
            if name in MENU_KEYS and is_menu(self.screen_id, self.layout):
                # Arrow on a menu: if it turns out to leave the menu (Right
                # on "Inklaring bedrijf" opens the program), it becomes the
                # route step -- see _finish_open_step.
                step["_menu_before"] = menu_path(self.layout)
            if name not in MENU_KEYS and is_menu(self.screen_id, self.layout):
                # A key that leaves (or acts in) a menu: the arrows that led
                # here become one route by name, e.g. Expeditie > Inklaring
                # bedrijf -- replay steers there from wherever the menu is.
                steps = self.action["steps"]
                while steps and steps[-1].get("key") in MENU_KEYS and steps[-1].get("to_menu"):
                    steps.pop()
                step["menu"] = menu_path(self.layout)
                self.status = f"Menu route: {' > '.join(step['menu'])}, then {name}"
            elif name == "Enter":
                # Arrowed through a list (e.g. the tours after F1) and then
                # Enter: the row's position changes from day to day, so
                # replay looks the row up by its value instead of pressing
                # Down the same number of times.
                steps = self.action["steps"]
                arrows = 0
                while steps and steps[-1].get("key") in ARROW_KEYS and steps[-1].get("to") == self.screen_id:
                    steps.pop()
                    arrows += 1
                highlighted = [row for row in screen_rows(self.layout) if row["highlighted"]]
                if arrows and highlighted:
                    step["pick"] = row_key(highlighted[0]["text"])
                    step["pick_row"] = highlighted[0]["text"]
                    self.status = f"List: will look for the row with '{step['pick']}' ({highlighted[0]['text'][:50]}), not count {arrows} key(s) -- change it under Edit if needed"
                # Plain Enter in a list (no arrows): on whatever row is
                # selected then -- its text (an order, a customer) differs
                # every day, so it's never looked for.
            self._open = (step, at)


# ------------------------------------------------------------------ replay
class ActionRun:
    """Replays an action. After every step it waits until Uniware really shows
    the recorded screen (and the step's "wait until" text) and has stopped
    changing -- up to the step's max wait -- and stops on anything else."""

    def __init__(self, action: dict, values: dict, title_part: str = DEFAULT_WINDOW_TITLE):
        self.action = action
        self.values = {key: str(value) for key, value in (values or {}).items()}
        self.title_part = title_part
        self.log: list[str] = []
        self.active = False
        self.stop_requested = False
        self.result = ""
        self.output_files: list[str] = []
        self.exported_files: list[str] = []
        self.captures: list[dict] = []
        self.started = datetime.datetime.now()

    def say(self, message: str):
        self.log.append(f"{datetime.datetime.now():%H:%M:%S} {message}")

    def _check(self, hwnd):
        if self.stop_requested:
            raise RuntimeError("Stopped by you.")
        if screen_reader.foreground_window() != hwnd:
            raise RuntimeError("Uniware lost focus (another window was clicked) -- stopped for safety.")

    def _read(self, hwnd, timeout: float = 3.0):
        image = screen_reader.capture_when_settled(hwnd, timeout=timeout)
        return (image, *recorder.describe_screen(image))

    def _wait_for(self, hwnd, step: dict, label: str):
        """Read the screen until it is the step's screen, contains its "wait
        until" text and two reads in a row are the same (finished loading),
        for at most the step's max wait. Returns the reading."""
        expected = step["to"]
        expected_title = step.get("to_title") or expected
        wait_text = (step.get("wait_text") or "").strip()
        max_wait = float(step.get("max_wait") or default_max_wait(float(step.get("recorded_wait") or 0)))
        started = time.time()
        next_note = started + 5
        previous_text = None
        while True:
            self._check(hwnd)
            image, screen_id, title, selected, text, layout = self._read(hwnd)
            joined = "\n".join(text).lower()
            if same_screen(expected, expected_title, screen_id, title) and wait_text.lower() in joined:
                if text == previous_text:
                    return screen_id, title, selected, text, layout
                previous_text = text
                continue  # read once more: is it still drawing?
            previous_text = None
            waited = time.time() - started
            if waited > max_wait:
                self._save_problem_screen(image)
                looking = f"'{expected_title}'" + (f" showing '{wait_text}'" if wait_text else "")
                raise RuntimeError(f"{label}: waited {max_wait:.0f}s for {looking}, but Uniware shows '{title}'. Stopped (screenshot in outputs).")
            if time.time() > next_note:
                self.say(f"{label}: Uniware still busy, waiting for {expected_title}" + (f" / '{wait_text}'" if wait_text else "") + f" ({waited:.0f}s of max {max_wait:.0f}s)")
                next_note = time.time() + 10
            time.sleep(0.5)

    def _read_column_value(self, hwnd, step: dict, label: str, layout: list):
        """Read a column (e.g. invoiceRes) off the screen into a field for
        the dump file name. The value can appear a moment after the key
        (Shift+F5 makes the invoice), so it waits -- up to the step's max
        wait -- until the column is filled and two separate screenshots
        read the same: a misread number must never end up in a file name
        that SnappySjaak takes as the invoice number."""
        column, field = step["read_column"], step["read_field"]
        max_wait = float(step.get("max_wait") or 60)
        started = time.time()
        next_note = started + 5
        previous = None
        while True:
            self._check(hwnd)
            values = [clean_number(value) for value in read_column(layout, column)]
            if values and values == previous:
                break
            previous = values or None
            waited = time.time() - started
            if waited > max_wait:
                if not values:
                    raise RuntimeError(f"{label}: column '{column}' stayed empty for {max_wait:.0f}s (or isn't on this screen) -- nothing to put in {{{field}}}.")
                raise RuntimeError(f"{label}: column '{column}' kept reading differently ({', '.join(dict.fromkeys(values))}). Stopped to be safe.")
            if time.time() > next_note:
                self.say(f"{label}: waiting for '{column}' to be filled ({waited:.0f}s of max {max_wait:.0f}s)")
                next_note = time.time() + 10
            time.sleep(0.5)
            _, _, _, _, _, layout = self._read(hwnd, timeout=1.5)
        distinct = list(dict.fromkeys(values))
        # SnappySjaak takes the LAST number in the dump's file name as the
        # invoice number (ukdocsInvoiceNumberFromFileName): it has to be one
        # value, digits only -- otherwise it would quietly pick the wrong one.
        if len(distinct) > 1:
            raise RuntimeError(f"{label}: '{column}' has {len(distinct)} different values ({', '.join(distinct)}) -- SnappySjaak expects one invoice number per file. Stopped.")
        if field == "invoiceres" and not distinct[0].isdigit():
            raise RuntimeError(f"{label}: '{column}' reads '{distinct[0]}', not just digits -- SnappySjaak would take only part of it as the invoice number. Stopped.")
        self.values[field] = distinct[0]
        self.say(f"{label}: {column} = {self.values[field]} -> {{{field}}}")

    def _take_window_dump(self, hwnd, windows_before: dict, step: dict, label: str):
        """The key opened the data in another program's window (Uniware's
        Ctrl+P): wait for that window, find its file, copy it as the dump
        (renamed, to the dump folder) and close the window."""
        uniware_pid = window_pid(hwnd)
        max_wait = float(step.get("max_wait") or 60)
        started = time.time()
        window = None
        while window is None:
            if self.stop_requested:
                raise RuntimeError("Stopped by you.")
            for other, (pid, title) in visible_windows().items():
                if other not in windows_before and pid != uniware_pid:
                    window = (other, pid, title)
                    break
            if window is None:
                if time.time() - started > max_wait:
                    raise RuntimeError(f"{label}: no new window opened within {max_wait:.0f}s after {step['key']}.")
                time.sleep(0.5)
        other, pid, title = window
        time.sleep(1.5)  # let the program put the file name in its title
        title = screen_reader.window_title(other) or title
        self.say(f"{label}: {step['key']} opened '{title}' ({Path(process_exe(pid)).name})")
        source = None
        extra = [self.action.get("uniware_dump_dir") or "", self.action.get("export_dir") or ""]
        while source is None:
            source = dump_file_of_window(pid, screen_reader.window_title(other) or title, self.started.timestamp(), extra)
            if source is None:
                if time.time() - started > max_wait + 30:
                    raise RuntimeError(f"{label}: could not find the file shown in '{title}'. Tell which folder it is in under Edit (Uniware saves its dump in).")
                time.sleep(1)
        previous_size = -1
        while True:  # still being written?
            size = source.stat().st_size
            if size > 0 and size == previous_size:
                break
            previous_size = size
            time.sleep(1)
        target_dir = Path((self.action.get("export_dir") or "").strip() or (self.action.get("uniware_dump_dir") or "").strip() or OUTPUTS_DIR)
        target_dir.mkdir(parents=True, exist_ok=True)
        target = target_dir / f"{self.export_name()}{source.suffix or '.txt'}"
        temp = target_dir / f"~{target.name}.tmp"
        temp.write_bytes(source.read_bytes())  # copy: the program may still hold it open
        temp.replace(target)
        self.exported_files.append(str(target))
        self.say(f"{label}: dump {source} -> {target}")
        if self.action.get("close_dump_window", True):
            user32.PostMessageW(other, 0x0010, 0, 0)  # WM_CLOSE
            time.sleep(1)

    def _back_to_menu(self, hwnd, controller, reading):
        """From inside a program, press the back key (Esc by default, Edit:
        "Back key") until a menu shows -- at most 8 times, and never once a
        menu is there (Esc on the main menu could close Uniware)."""
        back_key = self.action.get("back_key") or "Esc"
        screen_id, title, selected, text, layout = reading
        unchanged = 0
        for _ in range(8):
            if is_menu(screen_id, layout):
                return screen_id, title, selected, text, layout
            self._check(hwnd)
            self.say(f"Not in a menu ('{title}'): pressing {back_key}")
            send_combo(controller, back_key)
            _, new_id, title, selected, new_text, layout = self._read(hwnd)
            unchanged = unchanged + 1 if new_text == text else 0
            if unchanged >= 2:
                break
            screen_id, text = new_id, new_text
        raise RuntimeError(f"Could not get back to the menu from '{title}' with {back_key}. Go to the menu yourself and start again (or set the Back key under Edit).")

    def _go_menu(self, hwnd, controller, path: list, label: str):
        """Steer the menu to `path` (blue item per level) from wherever it is:
        Left back to the first level that's wrong, Up/Down within the level
        (counting rows from the blue item), Right into the next level --
        reading the screen after every move."""
        path = [item for item in path if item]
        if not path:
            return
        route = " > ".join(path)
        last_depth = None
        for _ in range(80):
            self._check(hwnd)
            image, screen_id, title, _, _, layout = self._read(hwnd, timeout=1.5)
            if not is_menu(screen_id, layout):
                self._save_problem_screen(image)
                raise RuntimeError(f"{label}: expected a menu on the way to {route}, but Uniware shows '{title}'. Stopped (screenshot in outputs).")
            levels = menu_levels(layout)
            active = max(index for index, level in enumerate(levels) if level["selected"] is not None)
            current = [levels[index]["items"][levels[index]["selected"]] if levels[index]["selected"] is not None else "" for index in range(active + 1)]
            wrong = next((index for index in range(min(active + 1, len(path))) if not same_item(current[index], path[index])), None)
            if wrong is None and active == len(path) - 1:
                self.say(f"{label}: menu at {route}")
                return
            if wrong is not None and wrong < active:
                self._press(controller, "Left", active - wrong)
            elif wrong is None and active > len(path) - 1:
                self._press(controller, "Left", active - (len(path) - 1))
            elif wrong is None:
                if last_depth is not None and last_depth == active:
                    self._save_problem_screen(image)
                    raise RuntimeError(f"{label}: Right on '{current[active]}' did not open its submenu (looking for {route}). Stopped (screenshot in outputs).")
                last_depth = active
                self._press(controller, "Right", 1)
                continue
            else:
                items = levels[active]["items"]
                target = next((index for index, item in enumerate(items) if same_item(item, path[active])), None)
                if target is None:
                    self._save_problem_screen(image)
                    raise RuntimeError(f"{label}: '{path[active]}' is not in this menu level ({', '.join(items[:8])}...). Stopped (screenshot in outputs).")
                distance = target - levels[active]["selected"]
                self._press(controller, "Down" if distance > 0 else "Up", abs(distance))
            last_depth = None
        raise RuntimeError(f"{label}: could not reach {route} in the menu. Stopped.")

    def _read_stable(self, hwnd):
        """Read until two screenshots in a row give the same text (the info
        panel next to a list updates a moment after the cursor moves)."""
        previous = None
        reading = None
        for _ in range(8):
            reading = self._read(hwnd, timeout=1.5)
            if previous is not None and reading[4] == previous:
                return reading
            previous = reading[4]
        return reading

    def _run_list_loop(self, hwnd, controller, loop: dict, label: str):
        """Go through a list row by row: when the screen shows one of the
        match words (e.g. the order type "Accessoire" next to the list), press
        the match key (Delete: out of the tour / Enter: into it), otherwise
        the next key (Down). Only acts when the type is clearly readable --
        one of the match or keep words, read the same twice; an unreadable
        row is skipped and logged. Stops at the end of the list (the cursor
        row no longer changes) or after max_rows. dry_run: only logs what it
        would do and moves on with the next key."""
        match_words = [word.strip() for word in loop.get("match", []) if word.strip()]
        keep_words = [word.strip() for word in loop.get("keep", []) if word.strip()]
        match_key = loop.get("match_key") or "Delete"
        next_key = loop.get("next_key") or "Down"
        dry_run = bool(loop.get("dry_run"))
        row_stays = bool(loop.get("row_stays"))
        only_rows_with = str(loop.get("only_rows_with") or "").strip()
        max_rows = max(1, min(2000, int(loop.get("max_rows") or 500)))
        if not match_words:
            raise RuntimeError(f"{label}: the list loop has no match words (Edit).")

        def found(words, text):
            for word in words:
                if re.search(rf"(?<![\w]){re.escape(word)}(?![\w])", text, re.IGNORECASE):
                    return word
            return None

        def cursor_row(layout):
            rows = screen_rows(layout)
            highlighted = [row["text"] for row in rows if row["highlighted"]]
            return highlighted[0] if highlighted else ""

        def same_row(left, right):
            return bool(left) and bool(right) and difflib.SequenceMatcher(None, left, right).ratio() >= 0.9

        counts = {"matched": 0, "kept": 0, "unreadable": 0}
        last_row, last_action = None, None
        handled = []
        for row_number in range(1, max_rows + 1):
            self._check(hwnd)
            _, _, _, _, text, layout = self._read_stable(hwnd)
            joined = "\n".join(text)
            row = cursor_row(layout)
            if last_action == "match" and same_row(row, last_row):
                raise RuntimeError(f"{label}: '{last_row[:50]}' is still there after {match_key} -- does Uniware ask for a confirmation? Stopped to be safe.")
            if last_action == "next" and same_row(row, last_row):
                break  # the cursor no longer moves: end of the list
            if last_action == "match" and any(same_row(row, done) for done in handled):
                # The last row was taken out: the cursor went back to one
                # already handled -- move on (and find the end) quietly.
                for part in list_loop_keys(next_key):
                    send_combo(controller, part)
                last_row, last_action = row, "next"
                continue
            handled.append(row)
            if only_rows_with and not value_matches(row, only_rows_with):
                self.say(f"{label}: {row[:60] or f'row {row_number}'}: no '{only_rows_with}' -> passed by")
                for part in list_loop_keys(next_key):
                    send_combo(controller, part)
                last_row, last_action = row, "next"
                time.sleep(0.3)
                continue
            matched = found(match_words, joined)
            kept = found(keep_words, joined) if not matched else None
            if not matched and not kept:
                # Read once more before giving up on this row.
                _, _, _, _, text, layout = self._read_stable(hwnd)
                joined = "\n".join(text)
                matched = found(match_words, joined)
                kept = found(keep_words, joined) if not matched else None
            short = row[:60] or f"row {row_number}"
            if matched and not dry_run and row_stays:
                counts["matched"] += 1
                self.say(f"{label}: {short}: {matched} -> {match_key}")
                for index, part in enumerate(list_loop_keys(match_key)):
                    if index:
                        self._read_stable(hwnd)
                        self._check(hwnd)
                    send_combo(controller, part)
                # The row stays: it must have changed (e.g. its tour number
                # cleared), then on to the next row.
                _, _, _, _, _, after_layout = self._read_stable(hwnd)
                after = cursor_row(after_layout)
                if after == row:
                    raise RuntimeError(f"{label}: '{row[:50]}' did not change after {match_key}. Stopped to be safe.")
                for part in list_loop_keys(next_key):
                    send_combo(controller, part)
                last_row, last_action = after, "next"
            elif matched and not dry_run:
                counts["matched"] += 1
                self.say(f"{label}: {short}: {matched} -> {match_key}")
                # The key(s) one after another, the screen settled in
                # between (a confirmation question has to be there before
                # it's answered).
                for index, part in enumerate(list_loop_keys(match_key)):
                    if index:
                        self._read_stable(hwnd)
                        self._check(hwnd)
                    send_combo(controller, part)
                last_row, last_action = row, "match"
            else:
                if matched:
                    counts["matched"] += 1
                    self.say(f"{label}: {short}: {matched} -> would press {match_key} (trial run)")
                elif kept:
                    counts["kept"] += 1
                    self.say(f"{label}: {short}: {kept} -> keep")
                else:
                    counts["unreadable"] += 1
                    self.say(f"{label}: {short}: type not readable -> skipped (check this one)")
                for part in list_loop_keys(next_key):
                    send_combo(controller, part)
                last_row, last_action = row, "next"
            time.sleep(0.3)
        else:
            self.say(f"{label}: stopped after {max_rows} rows (the most this loop may do).")
        self.say(f"{label}: list done -- {counts['matched']} {'would be ' if dry_run else ''}{match_key}, {counts['kept']} kept, {counts['unreadable']} not readable.")

    def _press(self, controller, key: str, times: int):
        for _ in range(times):
            send_combo(controller, key)
            time.sleep(0.05)

    def _choose_row(self, hwnd, controller, target: str, label: str):
        """Move the highlight in a list to the row showing `target`: read the
        list, count the rows between the highlight and that row, press
        Down/Up that often, read again to check. Rows not on screen: scroll
        down to the end of the list, then up."""
        if not target.strip():
            raise RuntimeError(f"{label}: no value to look for in the list (fill in the field).")
        direction, previous_text = "Down", None
        for _ in range(60):
            self._check(hwnd)
            image, _, _, _, text, layout = self._read(hwnd, timeout=1.5)
            rows = list_items(layout, target)
            highlighted = [index for index, row in enumerate(rows) if row["highlighted"]]
            if not highlighted:
                self._save_problem_screen(image)
                raise RuntimeError(f"{label}: can't see which row is selected in the list. Stopped (screenshot in outputs).")
            current = highlighted[0]
            found = find_row(rows, target)
            if found == current:
                self.say(f"{label}: '{target}' selected ({rows[current]['text'][:60]})")
                return
            if found is not None:
                distance = found - current
                self._press(controller, "Down" if distance > 0 else "Up", abs(distance))
                previous_text = None
                continue
            # Not on screen: move past the visible rows to scroll the list.
            if text == previous_text:
                if direction == "Up":
                    break
                direction = "Up"  # end of the list: search the part above
            previous_text = text
            # Past the last (first) visible row and about a screen further.
            to_edge = len(rows) - 1 - current if direction == "Down" else current
            self._press(controller, direction, to_edge + max(1, len(rows) - 1))
        self._save_problem_screen(screen_reader.capture_window(hwnd))
        raise RuntimeError(f"{label}: '{target}' is not in the list today. Stopped (screenshot in outputs).")

    def _save_problem_screen(self, image):
        OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
        name = f"{self.action['slug']}_{self.started:%Y%m%d_%H%M%S}_stopped.png"
        image.save(OUTPUTS_DIR / name)
        self.output_files.append(name)

    def _capture(self, hwnd, step_number: int, all_pages: bool, controller, reading):
        screen_id, title, _, text, layout = reading
        pages = [{"screen": screen_id, "title": title, "lines": layout}]
        previous = text
        while all_pages and len(pages) < 200:
            self._check(hwnd)
            send_combo(controller, "Page Down")
            # Wait until the next page is drawn (and stays the same); no
            # change at all within PAGE_WAIT_SECONDS = that was the last page.
            deadline = time.time() + PAGE_WAIT_SECONDS
            new_page = None
            while time.time() < deadline:
                self._check(hwnd)
                _, screen_id, title, _, text, layout = self._read(hwnd, timeout=1.5)
                if text != previous:
                    _, _, _, _, again, again_layout = self._read(hwnd, timeout=1.5)
                    if again == text:
                        new_page = (screen_id, title, text, again_layout)
                        break
                time.sleep(0.3)
            if new_page is None:
                break
            screen_id, title, previous, layout = new_page
            pages.append({"screen": screen_id, "title": title, "lines": layout})
        self.captures.append({"step": step_number, "pages": pages})
        self.say(f"Captured {len(pages)} page(s) of {title}")

    def run(self):
        self.active = True
        try:
            hwnd = find_uniware_window(self.title_part)
            if not hwnd:
                raise RuntimeError("Uniware window not found -- start Uniware (logged in) first.")
            if not bring_to_front(hwnd):
                raise RuntimeError("Could not bring Uniware to the front -- click it once and start again.")
            controller = keyboard.Controller()
            _, screen_id, title, selected, text, layout = self._read(hwnd)
            steps = self.action.get("steps", [])
            if steps and steps[0].get("menu"):
                # Starts with a menu route: any menu position is fine (the
                # route steers from there); from inside a program, go back
                # to the menu first.
                screen_id, title, selected, text, layout = self._back_to_menu(hwnd, controller, (screen_id, title, selected, text, layout))
            elif self.action.get("start_screen") and screen_id != self.action["start_screen"]:
                raise RuntimeError(f"Uniware is on '{title}', but this action starts on '{self.action.get('start_title') or self.action['start_screen']}'. Go there first.")
            self.say(f"Start: {title}")
            opens_window = any(step.get("opens_window") for step in steps)
            dump_before = None if opens_window else self._dump_files_now()
            if self.action.get("capture_start"):
                self._capture(hwnd, 0, False, controller, (screen_id, title, selected, text, layout))
            for number, step in enumerate(self.action.get("steps", []), start=1):
                self._check(hwnd)
                if step.get("loop"):
                    self._run_list_loop(hwnd, controller, step["loop"], f"Step {number}")
                    _, screen_id, title, selected, text, layout = self._read(hwnd)
                    continue
                if "text" in step:
                    typed = re.sub(r"\{([a-z0-9_]+)\}", lambda match: self.values.get(match.group(1), ""), step["text"])
                    controller.type(typed)
                    self.say(f"Step {number}: typed {typed!r}")
                else:
                    if step.get("menu"):
                        self._go_menu(hwnd, controller, step["menu"], f"Step {number}")
                    elif step.get("pick"):
                        target = re.sub(r"\{([a-z0-9_]+)\}", lambda match: self.values.get(match.group(1), ""), step["pick"])
                        self._choose_row(hwnd, controller, target, f"Step {number}")
                    elif step.get("enter_on"):
                        # Arrow to the recorded menu item first if the
                        # highlight is somewhere else.
                        for _ in range(30):
                            if selected and selected[0] == step["enter_on"]:
                                break
                            send_combo(controller, "Down")
                            _, _, _, selected, _, _ = self._read(hwnd, timeout=1.5)
                        else:
                            raise RuntimeError(f"Menu item '{step['enter_on']}' not found.")
                    windows_before = visible_windows() if step.get("opens_window") else None
                    send_combo(controller, step["key"])
                    if windows_before is not None:
                        self._take_window_dump(hwnd, windows_before, step, f"Step {number}")
                        if not bring_to_front(hwnd):
                            raise RuntimeError("Could not bring Uniware back to the front after the dump window.")
                    self.say(f"Step {number}: {step['key']}" + (f" on '{step['pick']}'" if step.get("pick") else f" on '{step['enter_on']}'" if step.get("enter_on") else ""))
                reading = self._wait_for(hwnd, step, f"Step {number}")
                screen_id, title, selected, text, layout = reading
                if step.get("read_column") and step.get("read_field"):
                    self._read_column_value(hwnd, step, f"Step {number}", layout)
                if step.get("capture"):
                    self._capture(hwnd, number, bool(step.get("capture_pages")), controller, reading)
            if dump_before is not None:
                self._collect_uniware_dump(dump_before)
            self.result = "done"
            self.say("Finished.")
        except Exception as error:  # noqa: BLE001 -- shown in the dashboard
            self.result = "failed"
            self.say(f"STOPPED: {error}")
        finally:
            try:
                self._write_outputs()
            except OSError as error:
                self.result = "failed"
                self.say(f"Could not write the output: {error}")
            self.active = False

    def export_name(self) -> str:
        """The dump file name from the action's template: {date}, {time},
        {action} and any field, e.g. "tour-508_{date}"."""
        template = self.action.get("export_name") or "{action}_{date}_{time}"
        names = {**self.values, "action": self.action["slug"], "date": f"{self.started:%Y%m%d}", "time": f"{self.started:%H%M%S}"}
        name = re.sub(r"\{([a-z0-9_]+)\}", lambda match: names.get(match.group(1), ""), template)
        return re.sub(r'[\\/:*?"<>|]+', "-", name).strip(" .") or self.action["slug"]

    # Uniware writes its own dump file (e.g. "dump.filipe.2026-10-14_16_57_30.csv")
    # into a folder; the run picks up the new one and gives it the action's name.
    def _dump_files_now(self) -> dict | None:
        folder = (self.action.get("uniware_dump_dir") or "").strip()
        if not folder:
            return None
        if not Path(folder).is_dir():
            raise RuntimeError(f"Uniware's dump folder '{folder}' doesn't exist (see Edit).")
        pattern = self.action.get("uniware_dump_pattern") or "dump.*.csv"
        return {path.name: path.stat().st_mtime for path in Path(folder).glob(pattern) if path.is_file()}

    def _collect_uniware_dump(self, before: dict):
        folder = Path(self.action["uniware_dump_dir"].strip())
        pattern = self.action.get("uniware_dump_pattern") or "dump.*.csv"
        max_wait = float(self.action.get("uniware_dump_wait") or 180)
        started = time.time()
        next_note = started + 10
        last_size, steady_since = None, None
        self.say(f"Waiting for Uniware's dump file ({pattern}) in {folder}")
        while True:
            if self.stop_requested:
                raise RuntimeError("Stopped by you.")
            new = [path for path in folder.glob(pattern)
                   if path.is_file() and (path.name not in before or path.stat().st_mtime > before[path.name])]
            if new:
                newest = max(new, key=lambda path: path.stat().st_mtime)
                size = newest.stat().st_size
                if size > 0 and size == last_size:
                    # Same size for 3 s: Uniware (and OneDrive) are done
                    # writing. Moving fails while it's still open -- retry.
                    if time.time() - steady_since >= 3 and self._move_dump(newest):
                        return
                else:
                    last_size, steady_since = size, time.time()
            waited = time.time() - started
            if waited > max_wait:
                raise RuntimeError(f"No new dump file ({pattern}) appeared in {folder} within {max_wait:.0f}s.")
            if time.time() > next_note:
                self.say(f"Still waiting for the dump file ({waited:.0f}s of max {max_wait:.0f}s)")
                next_note = time.time() + 15
            time.sleep(1)

    def _move_dump(self, source: Path) -> bool:
        target_dir = Path((self.action.get("export_dir") or "").strip() or source.parent)
        target_dir.mkdir(parents=True, exist_ok=True)
        target = target_dir / f"{self.export_name()}{source.suffix}"
        try:
            if target_dir.resolve().drive == source.resolve().drive:
                source.replace(target)
            else:
                temp = target_dir / f"~{target.name}.tmp"
                temp.write_bytes(source.read_bytes())
                temp.replace(target)
                source.unlink()
        except PermissionError:
            return False  # still being written
        self.exported_files.append(str(target))
        self.say(f"Dump file {source.name} -> {target}")
        return True

    def _write_outputs(self):
        if not self.captures:
            return
        OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
        stamp = f"{self.action['slug']}_{self.started:%Y%m%d_%H%M%S}"
        text_path = OUTPUTS_DIR / f"{stamp}.txt"
        csv_path = OUTPUTS_DIR / f"{stamp}.csv"
        with text_path.open("w", encoding="utf-8") as handle:
            for capture in self.captures:
                for index, page in enumerate(capture["pages"], start=1):
                    handle.write(f"=== step {capture['step']} -- {page['title']} -- page {index}\n")
                    for line in page["lines"]:
                        handle.write(f"{line['text']}\n")
        with csv_path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle, delimiter=";")
            writer.writerow(["step", "page", "screen", "y", "x", "text"])
            for capture in self.captures:
                for index, page in enumerate(capture["pages"], start=1):
                    for line in page["lines"]:
                        writer.writerow([capture["step"], index, page["screen"], line["y"], line["x"], line["text"]])
        (OUTPUTS_DIR / f"{stamp}.json").write_text(json.dumps(self.captures, indent=2, ensure_ascii=False), encoding="utf-8")
        self.output_files = [text_path.name, csv_path.name] + self.output_files
        export_dir = (self.action.get("export_dir") or "").strip()
        if not export_dir or self.action.get("uniware_dump_dir"):
            return  # with Uniware's own dump file, that file is the dump
        if self.result != "done":
            self.say("NOT put in the dump folder: the run did not finish (the partial data is only in outputs).")
            return
        target_dir = Path(export_dir)
        target_dir.mkdir(parents=True, exist_ok=True)
        base = self.export_name()
        for source, suffix in ((csv_path, ".csv"), (text_path, ".txt")):
            target = target_dir / f"{base}{suffix}"
            # Written under a temporary name first, so OneDrive never syncs a
            # half-written file.
            temp = target_dir / f"~{base}{suffix}.tmp"
            temp.write_bytes(source.read_bytes())
            temp.replace(target)
            self.exported_files.append(str(target))
        self.say(f"Dump written to {target_dir}: {base}.csv / .txt")
