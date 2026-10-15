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
import datetime
import json
import queue
import re
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
        else:
            rows.append({"y": line["y"], "parts": [line], "highlighted": line.get("highlighted", False)})
    return [{"y": row["y"], "highlighted": row["highlighted"],
             "text": " ".join(part["text"] for part in sorted(row["parts"], key=lambda part: part["x"]))} for row in rows]


def find_row(rows: list[dict], target: str) -> int | None:
    """Index of the row showing `target` as a whole value ("508" matches
    "508 GBR ..." but not "1508" or "50,80"); a row that starts with it wins."""
    pattern = re.compile(rf"(?<![\w.,]){re.escape(target.strip())}(?![\w.,]|,\d|\.\d)", re.IGNORECASE)
    matches = [index for index, row in enumerate(rows) if pattern.search(row["text"])]
    starting = [index for index in matches if rows[index]["text"].strip().lower().startswith(target.strip().lower())]
    return (starting or matches or [None])[0]


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
        step.update({
            "to": screen_id, "to_title": title, "capture": False, "capture_pages": False,
            "recorded_wait": round(waited, 1),
            "max_wait": default_max_wait(waited),
            "wait_text": "",
        })
        self.action["steps"].append(step)
        self.status = f"Step {len(self.action['steps'])}: {step.get('key') or 'type ' + repr(step['text'])} -> {title} (took {waited:.0f}s)"

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
            if name == "Enter":
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
                elif self.selected:
                    step["enter_on"] = self.selected[0]
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
            if screen_id == expected and wait_text.lower() in joined:
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
            rows = screen_rows(layout)
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
            if self.action.get("start_screen") and screen_id != self.action["start_screen"]:
                raise RuntimeError(f"Uniware is on '{title}', but this action starts on '{self.action.get('start_title') or self.action['start_screen']}'. Go there first.")
            self.say(f"Start: {title}")
            dump_before = self._dump_files_now()
            if self.action.get("capture_start"):
                self._capture(hwnd, 0, False, controller, (screen_id, title, selected, text, layout))
            for number, step in enumerate(self.action.get("steps", []), start=1):
                self._check(hwnd)
                if "text" in step:
                    typed = re.sub(r"\{([a-z0-9_]+)\}", lambda match: self.values.get(match.group(1), ""), step["text"])
                    controller.type(typed)
                    self.say(f"Step {number}: typed {typed!r}")
                else:
                    if step.get("pick"):
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
                    send_combo(controller, step["key"])
                    self.say(f"Step {number}: {step['key']}" + (f" on '{step['pick']}'" if step.get("pick") else f" on '{step['enter_on']}'" if step.get("enter_on") else ""))
                reading = self._wait_for(hwnd, step, f"Step {number}")
                screen_id, title, selected, text, layout = reading
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
