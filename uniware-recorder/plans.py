"""Week plan: per weekday a list of actions in order, with what to do after
each one (if it fails: stop / go on / try once more; wait minutes before the
next), and optionally a start time at which the dashboard starts it by
itself. Every action has a fixed number (#1, #2, ...) to plan with.

plans.json (next to this file):
  {"numbers": {"<slug>": 3, ...},
   "days": {"mon": {"start": "07:30", "auto": false, "last_auto": "2026-10-19",
                    "items": [{"action": "<slug>", "on_fail": "stop",
                               "wait_minutes": 0, "values": {}}]}, ...}}
"""
from __future__ import annotations

import datetime
import json
import re
import threading
import time

import actions

PLANS_FILE = actions.APP_DIR / "plans.json"
DAYS = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]
DAY_NAMES = {"mon": "Monday", "tue": "Tuesday", "wed": "Wednesday", "thu": "Thursday", "fri": "Friday", "sat": "Saturday", "sun": "Sunday"}
ON_FAIL = ("stop", "continue", "retry")
_LOCK = threading.Lock()


def today_key(now: datetime.datetime | None = None) -> str:
    return DAYS[(now or datetime.datetime.now()).weekday()]


def _empty_day() -> dict:
    return {"start": "", "auto": False, "last_auto": "", "items": []}


def normalize_item(item: dict) -> dict:
    try:
        wait = max(0, min(600, int(float(item.get("wait_minutes") or 0))))
    except (TypeError, ValueError):
        wait = 0
    on_fail = item.get("on_fail") if item.get("on_fail") in ON_FAIL else "stop"
    values = item.get("values") if isinstance(item.get("values"), dict) else {}
    return {
        "action": str(item.get("action") or "").strip(),
        "on_fail": on_fail,
        "wait_minutes": wait,
        "values": {str(key): str(value) for key, value in values.items()},
    }


def normalize_day(day: dict | None) -> dict:
    day = day if isinstance(day, dict) else {}
    start = str(day.get("start") or "").strip()
    if start and not re.fullmatch(r"([01]?\d|2[0-3]):[0-5]\d", start):
        start = ""
    return {
        "start": start,
        "auto": bool(day.get("auto")) and bool(start),
        "last_auto": str(day.get("last_auto") or ""),
        "items": [normalize_item(item) for item in day.get("items") or [] if isinstance(item, dict) and item.get("action")],
    }


def load_plans() -> dict:
    """The plan file, with a number for every action (new ones get the next
    free number; numbers of deleted actions are not reused)."""
    with _LOCK:
        data = {}
        if PLANS_FILE.exists():
            try:
                data = json.loads(PLANS_FILE.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                data = {}
        numbers = {str(slug): int(number) for slug, number in (data.get("numbers") or {}).items() if str(number).isdigit()}
        changed = False
        for action in sorted(actions.list_actions(), key=lambda action: action.get("created") or ""):
            if action["slug"] not in numbers:
                numbers[action["slug"]] = max(numbers.values(), default=0) + 1
                changed = True
        plans = {
            "numbers": numbers,
            "days": {day: normalize_day((data.get("days") or {}).get(day)) for day in DAYS},
        }
        if changed or not PLANS_FILE.exists():
            _write(plans)
        return plans


def _write(plans: dict):
    temp = PLANS_FILE.with_suffix(".tmp")
    temp.write_text(json.dumps(plans, indent=2, ensure_ascii=False), encoding="utf-8")
    temp.replace(PLANS_FILE)


def save_plans(plans: dict):
    with _LOCK:
        _write(plans)


def set_day(day: str, settings: dict) -> dict:
    if day not in DAYS:
        raise RuntimeError(f"Unknown day '{day}'.")
    plans = load_plans()
    known = {action["slug"] for action in actions.list_actions()}
    day_plan = normalize_day({**settings, "last_auto": plans["days"][day].get("last_auto", "")})
    missing = [item["action"] for item in day_plan["items"] if item["action"] not in known]
    if missing:
        raise RuntimeError(f"Unknown action(s): {', '.join(missing)}.")
    plans["days"][day] = day_plan
    save_plans(plans)
    return plans


def set_number(slug: str, number: int) -> dict:
    plans = load_plans()
    if slug not in plans["numbers"]:
        raise RuntimeError("Action not found.")
    if number < 1 or number > 999:
        raise RuntimeError("A number from 1 to 999.")
    other = next((other for other, value in plans["numbers"].items() if value == number and other != slug), None)
    if other:
        # Swap: the other action takes this one's old number.
        plans["numbers"][other] = plans["numbers"][slug]
    plans["numbers"][slug] = number
    save_plans(plans)
    return plans


class PlanRun:
    """Runs one day's plan: the actions one after another (each as a normal
    ActionRun, shown in "Last run"), with the plan's own log."""

    def __init__(self, day: str, plan: dict, set_current_run, started_by: str = "you"):
        self.day = day
        self.items = plan["items"]
        self.set_current_run = set_current_run  # shows each action's run in the dashboard
        self.started_by = started_by
        self.active = False
        self.stop_requested = False
        self.result = ""
        self.index = 0
        self.log: list[str] = []
        self.results: list[dict] = []
        self.current_run = None

    def say(self, message: str):
        self.log.append(f"{datetime.datetime.now():%H:%M:%S} {message}")

    def _wait(self, minutes: int):
        until = time.time() + minutes * 60
        while time.time() < until:
            if self.stop_requested:
                raise RuntimeError("Stopped by you.")
            time.sleep(1)

    def run(self):
        self.active = True
        numbers = load_plans()["numbers"]
        try:
            self.say(f"{DAY_NAMES[self.day]} plan started by {self.started_by}: {len(self.items)} action(s).")
            for self.index, item in enumerate(self.items):
                action = actions.load_action(item["action"])
                label = f"{self.index + 1}/{len(self.items)} #{numbers.get(item['action'], '?')}"
                if not action:
                    self.say(f"{label} '{item['action']}' no longer exists -- skipped.")
                    self.results.append({"action": item["action"], "result": "missing"})
                    continue
                tries = 2 if item["on_fail"] == "retry" else 1
                result = ""
                for attempt in range(1, tries + 1):
                    if self.stop_requested:
                        raise RuntimeError("Stopped by you.")
                    self.say(f"{label} {action['name']}" + (" (try 2)" if attempt > 1 else "") + " ...")
                    run = actions.ActionRun(action, dict(item.get("values") or {}))
                    self.current_run = run
                    self.set_current_run(run)
                    run.run()
                    result = run.result
                    last = run.log[-1][9:] if run.log else ""
                    self.say(f"{label} {action['name']}: {result}" + (f" -- {last}" if result != "done" else ""))
                    if result == "done" or any("Stopped by you" in line for line in run.log[-2:]):
                        break
                self.results.append({"action": item["action"], "result": result})
                if any("Stopped by you" in line for line in (self.current_run.log[-2:] if self.current_run else [])):
                    raise RuntimeError("Stopped by you.")
                if result != "done" and item["on_fail"] in ("stop", "retry"):
                    raise RuntimeError(f"'{action['name']}' failed -- plan stopped (the next actions did not run).")
                if item["wait_minutes"] and self.index < len(self.items) - 1:
                    self.say(f"Waiting {item['wait_minutes']} min before the next action.")
                    self._wait(item["wait_minutes"])
            failed = [entry for entry in self.results if entry["result"] != "done"]
            self.result = "done" if not failed else "done with failures"
            self.say(f"Plan finished: {len(self.results) - len(failed)} done" + (f", {len(failed)} failed/skipped." if failed else "."))
        except Exception as error:  # noqa: BLE001 -- shown in the dashboard
            self.result = "stopped" if "Stopped by you" in str(error) else "failed"
            self.say(f"PLAN STOPPED: {error}")
        finally:
            self.active = False

    def status(self) -> dict:
        return {
            "active": self.active,
            "day": self.day,
            "index": self.index,
            "total": len(self.items),
            "result": self.result,
            "log": self.log[-40:],
            "started_by": self.started_by,
        }


def due_auto_day(now: datetime.datetime | None = None) -> str | None:
    """Today's plan when it starts by itself and its start time has come
    (and it didn't start by itself today yet) -- within 2 hours of it, so a
    PC switched on in the afternoon doesn't run the morning plan."""
    now = now or datetime.datetime.now()
    day = today_key(now)
    plan = load_plans()["days"][day]
    if not plan["auto"] or not plan["items"] or plan["last_auto"] == now.date().isoformat():
        return None
    hour, minute = (int(part) for part in plan["start"].split(":"))
    start = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
    return day if start <= now <= start + datetime.timedelta(hours=2) else None


def mark_auto_started(day: str, now: datetime.datetime | None = None):
    plans = load_plans()
    plans["days"][day]["last_auto"] = (now or datetime.datetime.now()).date().isoformat()
    save_plans(plans)
