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
        for name in re.findall(r"\{([a-z0-9_]+)\}", step.get("text") or ""):
            if name not in fields:
                fields.append(name)
    return fields


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
        self._helper = recorder.Recorder.__new__(recorder.Recorder)  # for combo_name()
        self._helper.held = self.held

    def start(self):
        self.hwnd = find_uniware_window(self.title_part)
        if not self.hwnd:
            raise RuntimeError("Uniware window not found -- start Uniware (logged in) first.")
        bring_to_front(self.hwnd)
        image = screen_reader.capture_when_settled(self.hwnd)
        screen_id, title, selected, _, _ = recorder.describe_screen(image)
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
        self.events.put(("stop", None, time.time()))
        if self._thread:
            self._thread.join(timeout=10)
        if self._listener:
            self._listener.stop()
        self.active = False
        save_action(self.action)
        self.status = f"Saved: {self.action['name']} ({len(self.action['steps'])} steps)"
        return self.action

    def mark_capture(self):
        self.events.put(("capture", None, time.time()))

    def _on_press(self, key):
        if key in recorder.MODIFIERS:
            self.held.add(recorder.MODIFIERS[key])
            return
        vk = getattr(key, "vk", None)
        if {"Ctrl", "Alt"} <= self.held and vk in (0x43, 0x51):  # C, Q
            self.events.put(("capture" if vk == 0x43 else "stop", None, time.time()))
            return
        if not self.active or screen_reader.foreground_window() != self.hwnd:
            return
        name, printable = self._helper.combo_name(key)
        self.events.put(("key", (name, printable), time.time()))

    def _on_release(self, key):
        if key in recorder.MODIFIERS:
            self.held.discard(recorder.MODIFIERS[key])

    def _record(self, keys: list[tuple[str, bool]]):
        enter_on = self.selected[0] if [name for name, _ in keys] == ["Enter"] and self.selected else None
        image = screen_reader.capture_when_settled(self.hwnd)
        screen_id, title, selected, _, _ = recorder.describe_screen(image)
        if all(printable for _, printable in keys):
            typed = "".join(" " if name == "Space" else name for name, _ in keys)
            step = {"text": typed, "recorded_text": typed}
        else:
            step = {"key": keys[0][0]}
            if enter_on:
                step["enter_on"] = enter_on
        step.update({"to": screen_id, "to_title": title, "capture": False, "capture_pages": False})
        self.action["steps"].append(step)
        self.selected = selected
        self.status = f"Step {len(self.action['steps'])}: {step.get('key') or 'type ' + repr(step['text'])} -> {title}"

    def _worker(self):
        burst = []
        last_key_at = 0.0
        while True:
            try:
                kind, value, at = self.events.get(timeout=0.1)
            except queue.Empty:
                if burst and time.time() - last_key_at > 0.7:
                    self._record(burst)
                    burst = []
                continue
            if kind == "stop":
                if burst:
                    self._record(burst)
                return
            if kind == "capture":
                if burst:
                    self._record(burst)
                    burst = []
                if self.action["steps"]:
                    self.action["steps"][-1]["capture"] = True
                    self.status = f"Step {len(self.action['steps'])} will capture its screen"
                else:
                    self.action["capture_start"] = True
                    self.status = "The start screen will be captured"
                continue
            name, printable = value
            if printable:
                burst.append(value)
                last_key_at = at
                continue
            if burst:
                self._record(burst)
                burst = []
            self._record([value])


# ------------------------------------------------------------------ replay
class ActionRun:
    """Replays an action; check after every step, stop on anything unexpected."""

    def __init__(self, action: dict, values: dict, title_part: str = DEFAULT_WINDOW_TITLE):
        self.action = action
        self.values = {key: str(value) for key, value in (values or {}).items()}
        self.title_part = title_part
        self.log: list[str] = []
        self.active = False
        self.stop_requested = False
        self.result = ""
        self.output_files: list[str] = []
        self.captures: list[dict] = []

    def say(self, message: str):
        self.log.append(f"{datetime.datetime.now():%H:%M:%S} {message}")

    def _check_focus(self, hwnd):
        if screen_reader.foreground_window() != hwnd:
            raise RuntimeError("Uniware lost focus (another window was clicked) -- stopped for safety.")

    def _capture(self, hwnd, step_number: int, all_pages: bool, controller):
        pages = []
        previous = None
        for _ in range(60 if all_pages else 1):
            image = screen_reader.capture_when_settled(hwnd)
            screen_id, title, _, text, layout = recorder.describe_screen(image)
            if text == previous:
                break
            pages.append({"screen": screen_id, "title": title, "lines": layout})
            previous = text
            if not all_pages:
                break
            self._check_focus(hwnd)
            send_combo(controller, "Page Down")
        self.captures.append({"step": step_number, "pages": pages})
        self.say(f"Captured {len(pages)} page(s) of {pages[0]['title'] if pages else '?'}")

    def run(self):
        self.active = True
        try:
            hwnd = find_uniware_window(self.title_part)
            if not hwnd:
                raise RuntimeError("Uniware window not found -- start Uniware (logged in) first.")
            if not bring_to_front(hwnd):
                raise RuntimeError("Could not bring Uniware to the front -- click it once and start again.")
            controller = keyboard.Controller()
            image = screen_reader.capture_when_settled(hwnd)
            screen_id, title, selected, _, _ = recorder.describe_screen(image)
            if self.action.get("start_screen") and screen_id != self.action["start_screen"]:
                raise RuntimeError(f"Uniware is on '{title}', but this action starts on '{self.action.get('start_title') or self.action['start_screen']}'. Go there first.")
            self.say(f"Start: {title}")
            if self.action.get("capture_start"):
                self._capture(hwnd, 0, False, controller)
            for number, step in enumerate(self.action.get("steps", []), start=1):
                if self.stop_requested:
                    raise RuntimeError("Stopped by you.")
                self._check_focus(hwnd)
                if "text" in step:
                    text = re.sub(r"\{([a-z0-9_]+)\}", lambda match: self.values.get(match.group(1), ""), step["text"])
                    controller.type(text)
                    self.say(f"Step {number}: typed {text!r}")
                else:
                    if step.get("enter_on"):
                        # Arrow to the recorded menu item first if the
                        # highlight is somewhere else.
                        for _ in range(30):
                            if selected and selected[0] == step["enter_on"]:
                                break
                            send_combo(controller, "Down")
                            image = screen_reader.capture_when_settled(hwnd, timeout=1.5)
                            _, _, selected, _, _ = recorder.describe_screen(image)
                        else:
                            raise RuntimeError(f"Menu item '{step['enter_on']}' not found.")
                    send_combo(controller, step["key"])
                    self.say(f"Step {number}: {step['key']}" + (f" on '{step['enter_on']}'" if step.get("enter_on") else ""))
                image = screen_reader.capture_when_settled(hwnd)
                screen_id, title, selected, _, _ = recorder.describe_screen(image)
                if screen_id != step["to"]:
                    raise RuntimeError(f"Step {number}: expected '{step.get('to_title') or step['to']}', but Uniware shows '{title}'. Stopped.")
                if step.get("capture"):
                    self._capture(hwnd, number, bool(step.get("capture_pages")), controller)
            self.result = "done"
            self.say("Finished.")
        except Exception as error:  # noqa: BLE001 -- shown in the dashboard
            self.result = "failed"
            self.say(f"STOPPED: {error}")
        finally:
            self.active = False
            self._write_outputs()

    def _write_outputs(self):
        if not self.captures:
            return
        OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
        stamp = f"{self.action['slug']}_{datetime.datetime.now():%Y%m%d_%H%M%S}"
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
        self.output_files = [text_path.name, csv_path.name]
        self.say(f"Saved: outputs\\{text_path.name} and {csv_path.name}")
