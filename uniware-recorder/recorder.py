"""Uniware recorder -- map mode.

Watches only the Uniware window. For every key / key combination pressed
there (F-keys with Shift/Ctrl/Alt, letters, Enter, Esc, arrows...), it waits
until the screen stops changing, reads it (screen_reader.py) and records:

    screen before --key(s)--> screen after

into data/map.json (all sessions together), with a screenshot and the text
of every screen seen, plus a step-by-step log per session in data/sessions/.
make_report.bat turns map.json into data/map.html: every screen, what's on
it, and which keys lead where.

Hotkeys (anywhere):
    Ctrl+Alt+R   start recording the window that has focus now (Uniware)
    Ctrl+Alt+P   pause / resume (e.g. while typing a password)
    Ctrl+Alt+S   snapshot of the current screen without pressing a key
    Ctrl+Alt+Q   stop and save
"""
from __future__ import annotations

import datetime
import json
import queue
import re
import threading
import time
from pathlib import Path

from pynput import keyboard

import screen_reader

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
SCREENS_DIR = DATA_DIR / "screens"
SESSIONS_DIR = DATA_DIR / "sessions"
MAP_PATH = DATA_DIR / "map.json"
CONFIG_PATH = APP_DIR / "config.json"

DEFAULT_CONFIG = {
    # Typed text on a screen whose text contains one of these is stored as
    # "<hidden N chars>" instead of the characters themselves.
    "mask_text_on_screens_containing": ["wachtwoord", "password", "pincode", "pin code"],
    # Keys pressed within this many seconds of each other count as one step
    # (typing a number), the screen is read once after them.
    "burst_seconds": 0.7,
    # How long to wait for the screen to stop changing after a key.
    "settle_timeout_seconds": 3.0,
}

MODIFIERS = {
    keyboard.Key.shift: "Shift", keyboard.Key.shift_l: "Shift", keyboard.Key.shift_r: "Shift",
    keyboard.Key.ctrl: "Ctrl", keyboard.Key.ctrl_l: "Ctrl", keyboard.Key.ctrl_r: "Ctrl",
    keyboard.Key.alt: "Alt", keyboard.Key.alt_l: "Alt", keyboard.Key.alt_r: "Alt", keyboard.Key.alt_gr: "AltGr",
}
HOTKEY_VK = {0x52: "start", 0x50: "pause", 0x53: "snapshot", 0x51: "quit"}  # R P S Q


def load_config() -> dict:
    config = dict(DEFAULT_CONFIG)
    if CONFIG_PATH.exists():
        config.update(json.loads(CONFIG_PATH.read_text(encoding="utf-8-sig")))
    return config


def load_map() -> dict:
    if MAP_PATH.exists():
        return json.loads(MAP_PATH.read_text(encoding="utf-8"))
    return {"screens": {}, "transitions": {}}


def save_map(data: dict):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    temp = MAP_PATH.with_suffix(".tmp")
    temp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    temp.replace(MAP_PATH)


def now_iso() -> str:
    return datetime.datetime.now().isoformat(timespec="seconds")


def describe_screen(image):
    """(id, title, selected items, text lines, layout) of a captured screen,
    without the terminal program's own title/menu/status bars."""
    lines = screen_reader.strip_chrome(screen_reader.read_screen(image))
    screen_id, title = screen_reader.screen_identity(lines, image.height, image.width)
    header_limit = image.height * 0.15
    # The terminal program's status bar ("1,1 Sessie 1 ... UTF-8 ansi") is
    # the bottom ~45 px.
    footer_limit = image.height - 45
    body = [line for line in lines if header_limit <= line["y"] < footer_limit]
    selected = [line["text"] for line in body if line["highlighted"]]
    layout = [{"x": line["x"], "y": line["y"], "w": line.get("w", len(line["text"]) * 14), "text": line["text"], "highlighted": line["highlighted"], "cursor": line.get("cursor", False)} for line in body]
    return screen_id, title, selected, [line["text"] for line in body], layout


def remember_screen(data: dict, screen_id: str, title: str, text: list, layout: list, image=None):
    screens = data["screens"]
    entry = screens.get(screen_id)
    SCREENS_DIR.mkdir(parents=True, exist_ok=True)
    if entry is None:
        if image is not None:
            image.save(SCREENS_DIR / f"{screen_id}.png")
        entry = screens[screen_id] = {"title": title, "first_seen": now_iso(), "seen": 0, "texts": [], "layout": layout}
    # A fuller title wins ("Prg. 000" -> "Prg. 000 Hoofdmenu").
    if len(title) > len(entry.get("title", "")) and title.split(" ")[:2] == entry.get("title", "").split(" ")[:2]:
        entry["title"] = title
    entry["seen"] = entry.get("seen", 0) + 1
    entry["last_seen"] = now_iso()
    entry.setdefault("layout", layout)
    # Keep the first version, plus a few that differ (lists change).
    if text not in entry["texts"] and len(entry["texts"]) < 5:
        entry["texts"].append(text)
    return entry


class Recorder:
    def __init__(self):
        self.config = load_config()
        self.map = load_map()
        self.target_hwnd = None
        self.paused = False
        self.running = True
        self.held = set()
        self.events: queue.Queue = queue.Queue()
        self.current_screen = None  # id of the screen last seen
        self.current_selected = []  # its highlighted (selected) item(s)
        self.session_path = None

    # ------------------------------------------------------------ keyboard
    def combo_name(self, key) -> tuple[str, bool]:
        """("Shift+F5", printable?)"""
        if isinstance(key, keyboard.Key):
            base = key.name.upper() if re.fullmatch(r"f\d+", key.name) else key.name.replace("_", " ").title()
            printable = key == keyboard.Key.space
            if printable:
                base = "Space"
        else:
            char = key.char
            if char and char.isprintable() and not ({"Ctrl", "Alt"} & self.held):
                base, printable = char, True
            elif getattr(key, "vk", None):
                base, printable = chr(key.vk) if 0x30 <= key.vk <= 0x5A else f"vk{key.vk}", False
            else:
                base, printable = str(char), False
        modifiers = [name for name in ("Ctrl", "Alt", "AltGr", "Shift") if name in self.held]
        if printable and modifiers == ["Shift"]:
            modifiers = []  # Shift+a is just "A"
        return "+".join([*modifiers, base]), printable

    def on_press(self, key):
        if key in MODIFIERS:
            self.held.add(MODIFIERS[key])
            return
        vk = getattr(key, "vk", None)
        if {"Ctrl", "Alt"} <= self.held and vk in HOTKEY_VK:
            self.events.put(("hotkey", HOTKEY_VK[vk], time.time()))
            return
        if self.target_hwnd is None or self.paused or screen_reader.foreground_window() != self.target_hwnd:
            return
        name, printable = self.combo_name(key)
        self.events.put(("key", (name, printable), time.time()))

    def on_release(self, key):
        if key in MODIFIERS:
            self.held.discard(MODIFIERS[key])

    # ------------------------------------------------------------ screens
    def read_current(self):
        image = screen_reader.capture_when_settled(self.target_hwnd, timeout=self.config["settle_timeout_seconds"])
        screen_id, title, selected, text, layout = describe_screen(image)
        remember_screen(self.map, screen_id, title, text, layout, image)
        return screen_id, title, selected, text

    def masked(self, screen_text) -> bool:
        joined = " ".join(screen_text).lower()
        return any(word in joined for word in self.config["mask_text_on_screens_containing"])

    def record_step(self, keys: list[tuple[str, bool]]):
        before = self.current_screen
        before_text = self.map["screens"].get(before, {}).get("texts", [[]])[0] if before else []
        screen_id, title, selected, _ = self.read_current()
        if all(printable for _, printable in keys):
            typed = "".join(" " if name == "Space" else name for name, _ in keys)
            label = f'type <hidden {len(typed)} chars>' if self.masked(before_text) else f'type "{typed}"'
        else:
            label = " ".join(name for name, _ in keys)
            # Menus: arrow to an item, then Enter -- which item matters.
            if label == "Enter" and self.current_selected:
                label = f'Enter on "{self.current_selected[0]}"'
        if before:
            outgoing = self.map["transitions"].setdefault(before, {})
            targets = outgoing.setdefault(label, {})
            targets[screen_id] = targets.get(screen_id, 0) + 1
        self.current_screen = screen_id
        self.current_selected = selected
        self.log({"at": now_iso(), "from": before, "keys": label, "to": screen_id, "title": title, "selected": selected})
        arrow = f"[{self.map['screens'].get(before, {}).get('title', '-') if before else 'start'}] --{label}--> [{title}]"
        print(arrow + (f"   (selected: {', '.join(selected)})" if selected else ""))
        save_map(self.map)

    def log(self, entry: dict):
        if self.session_path:
            with self.session_path.open("a", encoding="utf-8") as handle:
                handle.write(json.dumps(entry, ensure_ascii=False) + "\n")

    # ------------------------------------------------------------ main loop
    def handle_hotkey(self, name: str):
        if name == "start":
            hwnd = screen_reader.foreground_window()
            if "uniware recorder" in screen_reader.window_title(hwnd).lower() or "python" in screen_reader.window_title(hwnd).lower():
                print("That's the recorder's own window -- click in the Uniware window first, then press Ctrl+Alt+R.")
                return
            self.target_hwnd = hwnd
            self.paused = False
            SESSIONS_DIR.mkdir(parents=True, exist_ok=True)
            self.session_path = SESSIONS_DIR / f"session_{datetime.datetime.now():%Y%m%d_%H%M%S}.jsonl"
            print(f'Recording window "{screen_reader.window_title(self.target_hwnd)}" -- work in Uniware as usual.')
            screen_id, title, selected, _ = self.read_current()
            self.current_screen = screen_id
            self.current_selected = selected
            self.log({"at": now_iso(), "start": screen_id, "title": title})
            print(f"Start screen: [{title}]")
            save_map(self.map)
        elif name == "pause":
            self.paused = not self.paused
            print("Paused -- keys are not recorded." if self.paused else "Recording again.")
        elif name == "snapshot" and self.target_hwnd:
            screen_id, title, selected, _ = self.read_current()
            self.current_screen = screen_id
            self.current_selected = selected
            print(f"Snapshot: [{title}]")
            save_map(self.map)
        elif name == "quit":
            self.running = False

    def worker(self):
        burst: list[tuple[str, bool]] = []
        last_key_at = 0.0
        while self.running:
            try:
                kind, value, at = self.events.get(timeout=0.1)
            except queue.Empty:
                if burst and time.time() - last_key_at >= self.config["burst_seconds"]:
                    try:
                        self.record_step(burst)
                    except Exception as error:  # noqa: BLE001 -- keep recording
                        print(f"Could not read the screen: {error}")
                    burst = []
                continue
            if kind == "hotkey":
                if burst:
                    self.record_step(burst)
                    burst = []
                self.handle_hotkey(value)
                continue
            name, printable = value
            try:
                if printable:
                    # Typing: collected, read once the typing pauses.
                    burst.append((name, printable))
                    last_key_at = at
                    continue
                # An F-key / Enter / Esc / arrow: typed text before it is its
                # own step first, then the key itself.
                if burst:
                    self.record_step(burst)
                    burst = []
                self.record_step([(name, printable)])
            except Exception as error:  # noqa: BLE001 -- keep recording
                print(f"Could not read the screen: {error}")
                burst = []

    def run(self):
        print(__doc__.split("Hotkeys (anywhere):")[1])
        print("Click the Uniware window, then press Ctrl+Alt+R to start.")
        thread = threading.Thread(target=self.worker, daemon=True)
        thread.start()
        with keyboard.Listener(on_press=self.on_press, on_release=self.on_release) as listener:
            while self.running:
                time.sleep(0.2)
            listener.stop()
        save_map(self.map)
        screens = len(self.map["screens"])
        moves = sum(len(keys) for keys in self.map["transitions"].values())
        print(f"Saved: {screens} screens, {moves} key moves in {MAP_PATH}. Run make_report.bat to see the map.")


if __name__ == "__main__":
    Recorder().run()
