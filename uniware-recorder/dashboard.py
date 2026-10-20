"""Uniware actions dashboard -- http://127.0.0.1:8765 (this PC only).

Lists the named actions (actions/*.json) with a Start button, records new
ones, lets typed text become a field and steps capture their screen, and
shows each run's progress and output files. Runs only on this PC; nothing
is sent anywhere.
"""
from __future__ import annotations

import json
import re
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse

from pynput import keyboard

import actions
import plans

PORT = 8765
STATE = {"recorder": None, "run": None, "plan": None, "message": ""}
LOCK = threading.Lock()


def busy() -> bool:
    return any(STATE[name] and STATE[name].active for name in ("run", "recorder", "plan"))


def set_current_run(run):
    STATE["run"] = run


def start_plan(day: str, started_by: str):
    plan = plans.load_plans()["days"][day]
    if not plan["items"]:
        raise RuntimeError(f"The {plans.DAY_NAMES[day]} plan has no actions.")
    plan_run = plans.PlanRun(day, plan, set_current_run, started_by)
    STATE["plan"] = plan_run
    threading.Thread(target=plan_run.run, daemon=True).start()
    STATE["message"] = f"Running the {plans.DAY_NAMES[day]} plan ({len(plan['items'])} actions) -- don't touch keyboard or mouse until it's done (Ctrl+Alt+Q stops)."


def plan_scheduler():
    """Starts today's plan by itself at its start time (when that is on)."""
    while True:
        time.sleep(20)
        try:
            day = plans.due_auto_day()
            if day:
                with LOCK:
                    if not busy():
                        plans.mark_auto_started(day)
                        start_plan(day, "its start time")
        except Exception as error:  # noqa: BLE001 -- never stop the scheduler
            STATE["message"] = f"Week plan could not start: {error}"


def state_payload() -> dict:
    recorder = STATE["recorder"]
    run = STATE["run"]
    plan_run = STATE["plan"]
    return {
        "plans": plans.load_plans(),
        "today": plans.today_key(),
        "plan_run": plan_run.status() if plan_run else None,
        "actions": [
            {**action, "fields": actions.action_fields(action)}
            for action in actions.list_actions()
        ],
        "day_lists": actions.list_day_lists(),
        "recording": {
            "active": bool(recorder and recorder.active),
            "name": recorder.action["name"] if recorder else "",
            "status": recorder.status if recorder else "",
            "steps": len(recorder.action["steps"]) if recorder else 0,
        },
        "run": {
            "active": bool(run and run.active),
            "name": run.action["name"] if run else "",
            "result": run.result if run else "",
            "log": run.log[-60:] if run else [],
            "outputs": run.output_files if run else [],
            "exported": run.exported_files if run else [],
        },
        "message": STATE["message"],
    }


def stop_hotkey_listener():
    """Ctrl+Alt+Q stops a running action from anywhere."""
    held = set()

    def on_press(key):
        if key in (keyboard.Key.ctrl, keyboard.Key.ctrl_l, keyboard.Key.ctrl_r):
            held.add("ctrl")
        elif key in (keyboard.Key.alt, keyboard.Key.alt_l, keyboard.Key.alt_r, keyboard.Key.alt_gr):
            held.add("alt")
        elif getattr(key, "vk", None) == 0x51 and {"ctrl", "alt"} <= held:
            if STATE["plan"] and STATE["plan"].active:
                STATE["plan"].stop_requested = True
            if STATE["run"] and STATE["run"].active:
                STATE["run"].stop_requested = True

    def on_release(key):
        if key in (keyboard.Key.ctrl, keyboard.Key.ctrl_l, keyboard.Key.ctrl_r):
            held.discard("ctrl")
        elif key in (keyboard.Key.alt, keyboard.Key.alt_l, keyboard.Key.alt_r, keyboard.Key.alt_gr):
            held.discard("alt")

    keyboard.Listener(on_press=on_press, on_release=on_release, daemon=True).start()


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def send(self, status: int, body, content_type="application/json"):
        data = body if isinstance(body, bytes) else (json.dumps(body, ensure_ascii=False) if content_type == "application/json" else body).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", f"{content_type}; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def body(self) -> dict:
        length = int(self.headers.get("Content-Length") or 0)
        return json.loads(self.rfile.read(length) or b"{}") if length else {}

    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/":
            self.send(200, PAGE, "text/html")
        elif path == "/api/state":
            self.send(200, state_payload())
        elif path.startswith("/outputs/"):
            name = Path(unquote(path[len("/outputs/"):])).name
            file = actions.OUTPUTS_DIR / name
            if not file.exists():
                self.send(404, {"error": "not found"})
                return
            kind = "text/csv" if name.endswith(".csv") else "text/plain"
            self.send(200, file.read_bytes(), kind)
        else:
            self.send(404, {"error": "not found"})

    def do_POST(self):
        path = urlparse(self.path).path
        try:
            body = self.body()
            with LOCK:
                if path == "/api/record/start":
                    if busy():
                        raise RuntimeError("Something is already running.")
                    name = str(body.get("name") or "").strip()
                    if not name:
                        raise RuntimeError("Give the action a name first.")
                    recorder = actions.ActionRecorder(name, str(body.get("description") or ""))
                    recorder.start()
                    STATE["recorder"] = recorder
                    STATE["message"] = f"Recording '{name}': do the action in Uniware now. Ctrl+Alt+C captures the current screen, Stop (or Ctrl+Alt+Q) when done."
                elif path == "/api/record/capture":
                    if STATE["recorder"] and STATE["recorder"].active:
                        STATE["recorder"].mark_capture()
                elif path == "/api/record/stop":
                    if STATE["recorder"] and STATE["recorder"].active:
                        action = STATE["recorder"].stop()
                        STATE["message"] = f"Saved '{action['name']}' with {len(action['steps'])} steps."
                elif path.startswith("/api/actions/") and path.endswith("/run"):
                    slug = path.split("/")[3]
                    if busy():
                        raise RuntimeError("Something is already running.")
                    action = actions.load_action(slug)
                    if not action:
                        raise RuntimeError("Action not found.")
                    run = actions.ActionRun(action, body.get("values") or {})
                    STATE["run"] = run
                    threading.Thread(target=run.run, daemon=True).start()
                    STATE["message"] = f"Running '{action['name']}' -- don't touch keyboard or mouse until it's done (Ctrl+Alt+Q stops)."
                elif path == "/api/run/stop":
                    if STATE["plan"] and STATE["plan"].active:
                        STATE["plan"].stop_requested = True
                    if STATE["run"] and STATE["run"].active:
                        STATE["run"].stop_requested = True
                elif path == "/api/plans/day":
                    day = str(body.get("day") or "")
                    plans.set_day(day, body)
                    STATE["message"] = f"{plans.DAY_NAMES.get(day, day)} plan saved."
                elif path == "/api/plans/number":
                    plans.set_number(str(body.get("slug") or ""), int(body.get("number") or 0))
                    STATE["message"] = "Number saved."
                elif path == "/api/plans/run":
                    if busy():
                        raise RuntimeError("Something is already running.")
                    day = str(body.get("day") or plans.today_key())
                    if day not in plans.DAYS:
                        raise RuntimeError("Unknown day.")
                    start_plan(day, "you")
                elif path.startswith("/api/actions/") and path.endswith("/update"):
                    slug = path.split("/")[3]
                    action = actions.load_action(slug)
                    if not action:
                        raise RuntimeError("Action not found.")
                    action["description"] = str(body.get("description", action.get("description", "")))
                    action["export_dir"] = str(body.get("export_dir", action.get("export_dir", ""))).strip().strip('"')
                    action["export_name"] = str(body.get("export_name", action.get("export_name", ""))).strip()
                    action["uniware_dump_dir"] = str(body.get("uniware_dump_dir", action.get("uniware_dump_dir", ""))).strip().strip('"')
                    action["uniware_dump_pattern"] = str(body.get("uniware_dump_pattern") or "dump.*.csv").strip()
                    try:
                        action["uniware_dump_wait"] = max(10, min(3600, int(float(body.get("uniware_dump_wait")))))
                    except (TypeError, ValueError):
                        action["uniware_dump_wait"] = 180
                    action["back_key"] = str(body.get("back_key") or action.get("back_key") or "Esc").strip()
                    if "close_dump_window" in body:
                        action["close_dump_window"] = bool(body.get("close_dump_window"))
                    if "join_dumps" in body:
                        action["join_dumps"] = bool(body.get("join_dumps"))
                    for index, change in enumerate(body.get("steps") or []):
                        if index < len(action["steps"]):
                            step = action["steps"][index]
                            if step.get("loop") and isinstance(change.get("loop"), dict):
                                step["loop"] = actions.normalize_list_loop(change["loop"])
                            if step.get("foreach") and isinstance(change.get("foreach"), dict):
                                step["foreach"] = actions.normalize_foreach(change["foreach"])
                            if step.get("classify") and isinstance(change.get("classify"), dict):
                                step["classify"] = actions.normalize_classify(change["classify"])
                            step["capture"] = bool(change.get("capture"))
                            if "key" in step and change.get("opens_window") is not None:
                                if change.get("opens_window"):
                                    step["opens_window"] = step.get("opens_window") or {"title": "", "exe": ""}
                                else:
                                    step.pop("opens_window", None)
                            step["capture_pages"] = bool(change.get("capture_pages"))
                            step["wait_text"] = str(change.get("wait_text") or "").strip()
                            try:
                                step["max_wait"] = max(5, min(1800, int(float(change.get("max_wait")))))
                            except (TypeError, ValueError):
                                step["max_wait"] = actions.default_max_wait(float(step.get("recorded_wait") or 0))
                            if step.get("key") == "Enter" and change.get("pick") is not None:
                                # Look the list row up by this value (e.g. a
                                # tour number) instead of its position.
                                pick = str(change["pick"]).strip()
                                if pick:
                                    step["pick"] = pick
                                    step.pop("enter_on", None)
                                else:
                                    step.pop("pick", None)
                                    step.pop("enter_on", None)
                            if "text" in step and change.get("field") is not None:
                                # What this step types: plain text as it is
                                # ("ltd" for a recorded "td"); a field in
                                # braces -- "{order}", "{date}", "#{pd_code}"
                                # -- filled in when it runs; empty: the
                                # recorded text.
                                value = str(change["field"])
                                if value.strip() == "":
                                    step["text"] = step.get("recorded_text", step["text"])
                                else:
                                    step["text"] = re.sub(r"\{([^{}]*)\}", lambda match: "{" + (re.sub(r"[^a-z0-9_]+", "_", match.group(1).lower()).strip("_") or "field") + "}", value)
                            if "key" in step and str(change.get("key") or "").strip() and change["key"].strip() != step["key"]:
                                # Another key for this step (e.g. Left instead of
                                # Right) without recording again -- must be a key
                                # the run can press.
                                new_key = change["key"].strip()
                                try:
                                    actions.parse_combo(new_key)
                                except (AttributeError, KeyError, ValueError):
                                    raise RuntimeError(f"Step {index + 1}: '{new_key}' is not a key (e.g. Left, Right, Enter, F6, Shift+F11, Ctrl+P).")
                                step["key"] = new_key
                            column = str(change.get("read_column") or "").strip()
                            field = re.sub(r"[^a-z0-9_]+", "_", str(change.get("read_field") or "").lower()).strip("_")
                            if column and field:
                                step["read_column"], step["read_field"] = column, field
                                step["read_row_with"] = str(change.get("read_row_with") or "").strip()
                            else:
                                step.pop("read_column", None)
                                step.pop("read_field", None)
                                step.pop("read_row_with", None)
                    removed = {index for index, change in enumerate(body.get("steps") or []) if change.get("remove")}
                    if removed:
                        action["steps"] = [step for index, step in enumerate(action["steps"]) if index not in removed]
                    actions.save_action(action)
                    STATE["message"] = f"Saved changes to '{action['name']}'."
                elif path.startswith("/api/actions/") and (path.endswith("/add-foreach") or path.endswith("/add-check")):
                    # "For each order of a day list" (repeats the steps after
                    # it) or a check step ("the screen shows SPICER -> 516").
                    slug = path.split("/")[3]
                    action = actions.load_action(slug)
                    if not action:
                        raise RuntimeError("Action not found.")
                    after = max(0, min(len(action["steps"]), int(body.get("after") or 0)))
                    if path.endswith("/add-foreach"):
                        step = {"foreach": actions.normalize_foreach({"list": body.get("list") or "", "count": 1}), "to": "", "to_title": "for each order"}
                    else:
                        step = {"classify": actions.normalize_classify({}), "to": "", "to_title": "check"}
                    step.update({"capture": False, "capture_pages": False})
                    action["steps"].insert(after, step)
                    actions.save_action(action)
                    STATE["message"] = f"Step {after + 1} added to '{action['name']}' -- fill it in, then Save."
                elif path.startswith("/api/daylists/") and path.endswith("/label"):
                    # A person sets / corrects an order's result in a day list.
                    data = actions.load_day_list(str(body.get("list") or ""))
                    entry = actions.day_list_item(data, str(body.get("key") or ""))
                    if entry is None:
                        raise RuntimeError("Order not in this day list.")
                    entry["label"] = str(body.get("label") or "").strip()
                    entry["warning"] = not entry["label"]
                    entry["note"] = (entry.get("note") or "") + (" -- set by hand" if "set by hand" not in (entry.get("note") or "") else "")
                    actions.save_day_list(data)
                    STATE["message"] = f"{entry['key']} set to {entry['label'] or '(none)'}."
                elif path.startswith("/api/daylists/") and path.endswith("/delete"):
                    name = str(body.get("list") or "")
                    (actions.DAYLISTS_DIR / f"{actions.day_list_slug(name)}.json").unlink(missing_ok=True)
                    STATE["message"] = f"Day list '{name}' deleted."
                elif path.startswith("/api/actions/") and path.endswith("/add-loop"):
                    # A list loop step (go through a list row by row, press a
                    # key depending on what the screen shows) after step N.
                    slug = path.split("/")[3]
                    action = actions.load_action(slug)
                    if not action:
                        raise RuntimeError("Action not found.")
                    after = max(0, min(len(action["steps"]), int(body.get("after") or 0)))
                    action["steps"].insert(after, {
                        "loop": actions.normalize_list_loop({"dry_run": True}),
                        "to": "", "to_title": "list loop", "capture": False, "capture_pages": False,
                    })
                    actions.save_action(action)
                    STATE["message"] = f"List loop added as step {after + 1} of '{action['name']}' (trial run on) -- fill in the words and keys, then Save."
                elif path.startswith("/api/actions/") and path.endswith("/copy"):
                    # A copy to change (e.g. tour 515 from tour 508) without
                    # recording it again.
                    slug = path.split("/")[3]
                    action = actions.load_action(slug)
                    if not action:
                        raise RuntimeError("Action not found.")
                    name = str(body.get("name") or "").strip()
                    if not name:
                        raise RuntimeError("Give the copy a name.")
                    new_slug = actions.slugify(name)
                    if (actions.ACTIONS_DIR / f"{new_slug}.json").exists():
                        raise RuntimeError(f"An action called '{name}' already exists.")
                    copy = json.loads(json.dumps(action))
                    copy.update({"name": name, "slug": new_slug, "copied_from": action["name"]})
                    actions.save_action(copy)
                    STATE["message"] = f"Copied '{action['name']}' to '{name}' -- change it under Edit."
                elif path.startswith("/api/actions/") and path.endswith("/delete"):
                    slug = path.split("/")[3]
                    (actions.ACTIONS_DIR / f"{slug}.json").unlink(missing_ok=True)
                    STATE["message"] = "Action deleted."
                else:
                    self.send(404, {"error": "not found"})
                    return
            self.send(200, state_payload())
        except Exception as error:  # noqa: BLE001 -- shown on the page
            self.send(400, {"error": str(error)})


PAGE = """<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Uniware actions</title>
<style>
  :root { --bg:#f6f7f9; --card:#fff; --line:#d6dbe1; --text:#17202a; --muted:#6b7280; --accent:#1f7a5a; --warn:#b45309; --bad:#b91c1c; }
  body { font-family: system-ui, sans-serif; margin: 0; padding: 16px; background: var(--bg); color: var(--text); }
  h1 { margin: 0 0 4px; } .muted { color: var(--muted); }
  .card { background: var(--card); border: 1px solid var(--line); border-radius: 8px; padding: 14px 16px; margin: 12px 0; }
  button { border: 1px solid #b8c2cc; border-radius: 8px; background: #fff; padding: 7px 12px; cursor: pointer; font: inherit; }
  button.primary { background: var(--accent); border-color: var(--accent); color: #fff; font-weight: 700; }
  button.danger { color: var(--bad); }
  input, textarea { font: inherit; padding: 6px 8px; border: 1px solid #b8c2cc; border-radius: 6px; }
  table { border-collapse: collapse; width: 100%; } td, th { border-bottom: 1px solid var(--line); padding: 6px; text-align: left; vertical-align: top; }
  .row { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }
  pre { background: #111827; color: #e5e7eb; padding: 10px; border-radius: 6px; max-height: 320px; overflow: auto; white-space: pre-wrap; }
  .msg { background: #fffbeb; border: 1px solid #f59e0b; border-radius: 8px; padding: 8px 12px; }
  .err { background: #fef2f2; border-color: var(--bad); }
  .tabs { display: flex; gap: 4px; flex-wrap: wrap; margin: 8px 0; }
  .tabs button.on { background: var(--accent); border-color: var(--accent); color: #fff; font-weight: 700; }
  .num { display: inline-block; min-width: 28px; padding: 1px 6px; border-radius: 6px; background: #e8f3ee; color: var(--accent); font-weight: 700; text-align: center; }
  select { font: inherit; padding: 5px 6px; border: 1px solid #b8c2cc; border-radius: 6px; }
</style></head><body>
<h1>Uniware actions</h1>
<div class="muted">Record an action once in Uniware, start it from here. Runs on this PC only.</div>
<div id="msg"></div>
<div class="card" id="recordCard"></div>
<div class="card" id="runCard"></div>
<div class="card" id="planCard"></div>
<div class="card"><h2>Actions</h2><div id="actions"></div></div>
<div class="card"><h2>Day lists</h2><div class="muted">Information one action collects for the next (e.g. the accessoire orders of tour 515 and whether they are SPICER). Set a result by hand before the assigning action runs.</div><div id="dayLists"></div></div>
<div class="card" id="editCard" style="display:none"></div>
<script>
let state = null, editing = null;
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
async function api(path, body) {
  const response = await fetch(path, { method: body === undefined ? "GET" : "POST", headers: {"Content-Type":"application/json"}, body: body === undefined ? undefined : JSON.stringify(body) });
  const payload = await response.json();
  if (!response.ok) { $("msg").innerHTML = `<div class="msg err">${esc(payload.error)}</div>`; return null; }
  state = payload; render(); return payload;
}
function render() {
  if (!state) return;
  $("msg").innerHTML = state.message ? `<div class="msg">${esc(state.message)}</div>` : "";
  if (!document.activeElement || !document.activeElement.closest("#dayLists")) renderDayLists();
  if (!document.activeElement || !document.activeElement.closest("#planCard")) renderPlan();
  const rec = state.recording;
  $("recordCard").innerHTML = rec.active
    ? `<h2>Recording: ${esc(rec.name)}</h2><p>${esc(rec.status)} -- ${rec.steps} step(s)</p>
       <div class="row"><button onclick="api('/api/record/capture',{})">Capture this screen (Ctrl+Alt+C)</button><button class="primary" onclick="api('/api/record/stop',{})">Stop and save</button></div>`
    : `<h2>Record a new action</h2>
       <div class="row"><input id="newName" placeholder="Name, e.g. UKdocs invoice data" size="36"><input id="newDesc" placeholder="What it does (optional)" size="50">
       <button class="primary" onclick="startRecording()">Record</button></div>
       <p class="muted">Put Uniware on the screen where the action starts (e.g. the main menu), then click Record. Do the action by hand; press Ctrl+Alt+C on each screen whose data you want, then Stop.</p>`;
  const run = state.run;
  $("runCard").style.display = run.name ? "" : "none";
  $("runCard").innerHTML = run.name ? `<h2>${run.active ? "Running" : "Last run"}: ${esc(run.name)} ${run.result ? `<span class="muted">(${esc(run.result)})</span>` : ""}</h2>
      ${run.active ? `<button class="danger" onclick="api('/api/run/stop',{})">Stop (Ctrl+Alt+Q)</button>` : ""}
      <pre>${esc(run.log.join("\\n"))}</pre>
      ${run.outputs.map((f) => `<a href="/outputs/${encodeURIComponent(f)}" target="_blank">${esc(f)}</a>`).join(" &middot; ")}
      ${run.exported.length ? `<p>Dump: ${run.exported.map(esc).join("<br>")}</p>` : ""}` : "";
  $("actions").innerHTML = state.actions.length ? `<table><tr><th>#</th><th>Action</th><th>Steps</th><th>Fields</th><th></th></tr>${state.actions.slice().sort((x, y) => (state.plans.numbers[x.slug] || 0) - (state.plans.numbers[y.slug] || 0)).map((a) => `
      <tr><td><input size="2" value="${esc(state.plans.numbers[a.slug] || "")}" title="This action's number in the week plan (another action with that number gets this one's)" onchange="setNumber('${a.slug}', this.value)"></td><td><strong>${esc(a.name)}</strong><div class="muted">${esc(a.description || "")}</div><div class="muted">starts on: ${a.steps.length && a.steps[0].menu ? `any menu position (goes back with ${esc(a.back_key || "Esc")} if needed)` : esc(a.start_title || a.start_screen)}</div>${a.uniware_dump_dir ? `<div class="muted">dump file &rarr; ${esc(a.export_name || "{action}_{date}_{time}")}.csv</div>` : a.export_dir ? `<div class="muted">dump to: ${esc(a.export_dir)} &rarr; ${esc(a.export_name || "{action}_{date}_{time}")}.csv</div>` : ""}</td>
      <td>${a.steps.length} (${a.steps.filter((s) => s.capture).length} capture)</td>
      <td>${a.fields.map((f) => `<label>${esc(f)} <input data-field="${esc(a.slug)}:${esc(f)}" size="14"></label>`).join("<br>") || "-"}</td>
      <td class="row"><button class="primary" onclick="runAction('${a.slug}')">Start</button><button onclick="edit('${a.slug}')">Edit</button><button onclick="copyAction('${a.slug}')">Copy</button><button class="danger" onclick="del('${a.slug}')">Delete</button></td></tr>`).join("")}</table>`
    : `<p class="muted">No actions yet -- record one above.</p>`;
}
const DAY_KEYS = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"];
const DAY_LABELS = { mon: "Monday", tue: "Tuesday", wed: "Wednesday", thu: "Thursday", fri: "Friday", sat: "Saturday", sun: "Sunday" };
let planDay = null;
function actionBySlug(slug) { return state.actions.find((a) => a.slug === slug); }
function actionLabel(slug) {
  const a = actionBySlug(slug);
  return `#${state.plans.numbers[slug] || "?"} ${a ? a.name : slug + " (deleted)"}`;
}
function dayDraft(day) { return JSON.parse(JSON.stringify(state.plans.days[day])); }
function saveDay(day, draft) { api("/api/plans/day", { day, ...draft }); }
function setNumber(slug, value) { api("/api/plans/number", { slug, number: Number(value) }); }
function planChange(day, change) { const draft = dayDraft(day); change(draft); saveDay(day, draft); }
function planMove(day, index, by) {
  planChange(day, (draft) => {
    const target = index + by;
    if (target < 0 || target >= draft.items.length) return;
    [draft.items[index], draft.items[target]] = [draft.items[target], draft.items[index]];
  });
}
function planAdd(day) {
  const slug = $("planAddAction").value;
  if (slug) planChange(day, (draft) => draft.items.push({ action: slug, on_fail: "stop", wait_minutes: 0, values: {} }));
}
function planSetNumbers(day) {
  const bySlug = Object.entries(state.plans.numbers);
  const wanted = $("planNumbers").value.split(/[^0-9]+/).filter(Boolean).map(Number);
  const unknown = wanted.filter((number) => !bySlug.some(([, value]) => value === number));
  if (unknown.length) { $("msg").innerHTML = `<div class="msg err">No action with number ${unknown.join(", ")}.</div>`; return; }
  planChange(day, (draft) => {
    const old = draft.items;
    draft.items = wanted.map((number) => {
      const slug = bySlug.find(([, value]) => value === number)[0];
      return old.find((item) => item.action === slug) || { action: slug, on_fail: "stop", wait_minutes: 0, values: {} };
    });
  });
}
function planCopyTo(day) {
  const target = $("planCopyTo").value;
  if (!target || target === day) return;
  if (state.plans.days[target].items.length && !confirm(`Replace the ${DAY_LABELS[target]} plan with this one?`)) return;
  const draft = dayDraft(day);
  draft.auto = false;  // a copied plan never starts by itself until you say so
  saveDay(target, draft);
  planDay = target;
}
function planRun(day) {
  if (!confirm(`Run the ${DAY_LABELS[day]} plan now (${state.plans.days[day].items.length} actions)? Don't touch keyboard or mouse until it's done.`)) return;
  api("/api/plans/run", { day });
}
function planSetValue(day, i, field, value) {
  planChange(day, (d) => { d.items[i].values = { ...(d.items[i].values || {}), [field]: value }; });
}
function renderPlan() {
  const plans = state.plans;
  const day = planDay || state.today;
  const plan = plans.days[day];
  const run = state.plan_run;
  const onFail = (item, i) => `<select onchange="planChange('${day}', (d) => { d.items[${i}].on_fail = this.value; })">
      <option value="stop" ${item.on_fail === "stop" ? "selected" : ""}>stop the plan</option>
      <option value="continue" ${item.on_fail === "continue" ? "selected" : ""}>go on with the next</option>
      <option value="retry" ${item.on_fail === "retry" ? "selected" : ""}>try once more, then stop</option></select>`;
  const fields = (item, i) => {
    const a = actionBySlug(item.action);
    return a && a.fields.length
      ? a.fields.map((field) => `<label>${esc(field)} <input size="8" value="${esc((item.values || {})[field] || "")}" data-plan-field="${esc(field)}" onchange="planSetValue('${day}', ${i}, this.dataset.planField, this.value)"></label>`).join("<br>")
      : `<span class="muted">-</span>`;
  };
  const options = state.actions.slice().sort((x, y) => (plans.numbers[x.slug] || 0) - (plans.numbers[y.slug] || 0))
    .map((a) => `<option value="${esc(a.slug)}">${esc(actionLabel(a.slug))}</option>`).join("");
  $("planCard").innerHTML = `<h2>Week plan</h2>
    <div class="muted">Per day the actions in your order (the numbers are in the Actions list). Each action ends on the main menu, so the next one starts from there.</div>
    <div class="tabs">${DAY_KEYS.map((key) => `<button class="${key === day ? "on" : ""}" onclick="planDay='${key}'; renderPlan()">${DAY_LABELS[key]}${key === state.today ? " (today)" : ""} <span class="muted">${plans.days[key].items.length || ""}</span>${plans.days[key].auto ? " &#9200;" : ""}</button>`).join("")}</div>
    ${run ? `<div class="card" style="margin:6px 0"><strong>${run.active ? "Running" : "Last plan"}: ${esc(DAY_LABELS[run.day])}</strong> -- ${run.active ? `action ${run.index + 1} of ${run.total}` : esc(run.result)} <span class="muted">(started by ${esc(run.started_by)})</span>
      ${run.active ? ` <button class="danger" onclick="api('/api/run/stop',{})">Stop the plan (Ctrl+Alt+Q)</button>` : ""}<pre>${esc(run.log.join(String.fromCharCode(10)))}</pre></div>` : ""}
    ${plan.items.length ? `<table><tr><th></th><th>Action</th><th>Fields</th><th>If it fails</th><th>Then wait</th><th></th></tr>
      ${plan.items.map((item, i) => `<tr><td>${i + 1}.</td><td><span class="num">#${esc(plans.numbers[item.action] || "?")}</span> ${esc((actionBySlug(item.action) || {}).name || item.action + " (deleted)")}</td>
        <td>${fields(item, i)}</td><td>${onFail(item, i)}</td>
        <td><input size="3" value="${esc(item.wait_minutes || 0)}" onchange="planChange('${day}', (d) => { d.items[${i}].wait_minutes = Number(this.value) || 0; })"> min</td>
        <td class="row"><button onclick="planMove('${day}', ${i}, -1)" ${i ? "" : "disabled"}>&uarr;</button><button onclick="planMove('${day}', ${i}, 1)" ${i < plan.items.length - 1 ? "" : "disabled"}>&darr;</button><button class="danger" onclick="planChange('${day}', (d) => d.items.splice(${i}, 1))">Remove</button></td></tr>`).join("")}</table>`
      : `<p class="muted">No actions on ${DAY_LABELS[day]} yet.</p>`}
    <div class="row" style="margin-top:8px"><label>Add <select id="planAddAction">${options}</select></label><button onclick="planAdd('${day}')">Add to ${DAY_LABELS[day]}</button>
      <span class="muted">or type the numbers in order:</span><input id="planNumbers" size="14" placeholder="e.g. 3, 7, 4" value="${esc(plan.items.map((item) => plans.numbers[item.action]).join(", "))}"><button onclick="planSetNumbers('${day}')">Set ${DAY_LABELS[day]}</button></div>
    <div class="row" style="margin-top:8px"><label>Start at <input type="time" value="${esc(plan.start)}" onchange="planChange('${day}', (d) => { d.start = this.value; if (!this.value) d.auto = false; })"></label>
      <label><input type="checkbox" ${plan.auto ? "checked" : ""} ${plan.start ? "" : "disabled"} onchange="planChange('${day}', (d) => { d.auto = this.checked; })"> start by itself at that time</label>
      <span class="muted">(this dashboard must be open and Uniware logged in on the main menu; it starts once a day, up to 2 hours after the time)</span></div>
    <div class="row" style="margin-top:8px"><button class="primary" onclick="planRun('${day}')" ${plan.items.length ? "" : "disabled"}>Run the ${DAY_LABELS[day]} plan now</button>
      <label>Copy this day to <select id="planCopyTo"><option value=""></option>${DAY_KEYS.filter((key) => key !== day).map((key) => `<option value="${key}">${DAY_LABELS[key]}</option>`).join("")}</select></label><button onclick="planCopyTo('${day}')">Copy</button></div>`;
}
function startRecording() { api("/api/record/start", { name: $("newName").value, description: $("newDesc").value }); }
function runAction(slug) {
  const values = {};
  document.querySelectorAll(`[data-field^="${slug}:"]`).forEach((input) => { values[input.dataset.field.split(":")[1]] = input.value; });
  if (!confirm("Start this action? Don't touch keyboard or mouse until it's done.")) return;
  api(`/api/actions/${slug}/run`, { values });
}
function copyAction(slug) {
  const a = state.actions.find((x) => x.slug === slug);
  const name = prompt("Name of the copy:", `${a.name} (copy)`);
  if (name) api(`/api/actions/${slug}/copy`, { name });
}
function del(slug) { if (confirm("Delete this action?")) api(`/api/actions/${slug}/delete`, {}); }
function edit(slug) {
  const a = state.actions.find((x) => x.slug === slug); editing = slug;
  $("editCard").style.display = "";
  $("editCard").innerHTML = `<h2>Edit: ${esc(a.name)}</h2>
    <p><input id="editDesc" size="70" value="${esc(a.description || "")}"></p>
    <p class="row"><label>Back key (from a program back to the menu) <input id="editBackKey" size="8" value="${esc(a.back_key || "Esc")}"></label></p>
    <h3>Dump file</h3>
    <p class="muted">Uniware's Ctrl+P opens the data in a new window: tick <strong>opens the dump in a new window</strong> on that step (done by itself when recorded). The run then waits for that window, finds the file it shows, copies it with the name below to the folder below, and closes the window.</p>
    <p class="row"><label><input type="checkbox" id="editCloseWindow" ${a.close_dump_window === false ? "" : "checked"}> close that window afterwards</label>
      <label><input type="checkbox" id="editJoinDumps" ${a.join_dumps ? "checked" : ""}> <strong>join all dumps of a run into one file</strong> (e.g. tour 508: two companies, two Ctrl+P) -- named at the end, so the file name can use a value read after the first dump</label></p>
    <p class="row"><label>Uniware saves its dump in <input id="editUniwareDir" size="60" placeholder="paste the folder path (e.g. the OneDrive folder)" value="${esc(a.uniware_dump_dir || "")}"></label>
      <label>Its name looks like <input id="editUniwarePattern" size="14" value="${esc(a.uniware_dump_pattern || "dump.*.csv")}"></label>
      <label>Max wait (s) <input id="editUniwareWait" size="4" value="${esc(a.uniware_dump_wait || 180)}"></label></p>
    <p class="row"><label>Rename it to <input id="editExportName" size="28" placeholder="{action}_{date}_{time}" value="${esc(a.export_name || "")}"></label>
      <label>and put it in <input id="editExportDir" size="50" placeholder="e.g. the OneDrive folder" value="${esc(a.export_dir || "")}"></label></p>
    <p class="muted">At the end of the run it waits for the new dump file Uniware writes (e.g. dump.filipe.2026-10-14_16_57_30.csv), waits until it is completely written, then renames it, e.g. <code>tour-508_{date}</code> gives tour-508_20261014.csv. You can use {date}, {time}, {action} and the fields. Same name again = replaced. No new file within Max wait = the run fails and the file is left alone.</p>
    <table><tr><th>#</th><th>Step</th><th>Goes to</th><th>Field (typed text) / list row</th><th>Wait until screen shows</th><th>Max wait (s)</th><th>Read column (into field)</th><th>Capture screen</th><th>All pages</th><th>Remove</th></tr>
    ${a.steps.map((s, i) => `<tr><td>${i + 1}</td>
      <td>${s.foreach ? foreachEditor(s.foreach, i) : s.classify ? classifyEditor(s.classify, i) : s.loop ? loopEditor(s.loop, i) : s.text !== undefined ? `type <code>${esc(s.text)}</code>` : `${s.menu ? `menu <strong>${s.menu.map(esc).join(" &rsaquo; ")}</strong>, then ` : ""}<input data-step-key="${i}" value="${esc(s.key)}" size="9" title="The key this step presses -- change it (e.g. Left instead of Right) without recording again">${s.enter_on ? ` on "${esc(s.enter_on)}"` : ""}${s.key ? `<div><label><input type="checkbox" data-step-window="${i}" ${s.opens_window ? "checked" : ""}> opens the dump in a new window</label>${s.opens_window && s.opens_window.title ? `<div class="muted">recorded: ${esc(s.opens_window.title)}${s.opens_window.exe ? ` (${esc(s.opens_window.exe)})` : ""}</div>` : ""}</div>` : ""}`}</td>
      <td>${esc(s.to_title || s.to)}${s.recorded_wait ? `<div class="muted">took ${esc(s.recorded_wait)}s when recorded</div>` : ""}</td>
      <td>${s.text !== undefined ? `<input data-step-field="${i}" placeholder="text, or {field}" title="What this step types. Plain text is typed as it is (e.g. ltd instead of the recorded td). A field in braces is filled in when it runs: {order} (asked at start / filled by 'for each'), {today} / {tomorrow} / {next_workday} (that date, written like the recorded one), also mixed: #{pd_code}. Empty: the recorded text." value="${esc(s.text !== (s.recorded_text ?? s.text) ? s.text : "")}" size="12"><div class="muted">recorded: ${esc(s.recorded_text || s.text)}</div>` : s.key === "Enter" ? `find row with <input data-step-pick="${i}" placeholder="empty = selected row" value="${esc(s.pick || "")}" size="12">${s.pick_row ? `<div class="muted">recorded row: ${esc(s.pick_row)}</div>` : ""}` : ""}</td>
      <td><input data-step-wait="${i}" size="16" placeholder="(optional) e.g. Totaal" value="${esc(s.wait_text || "")}"></td>
      <td><input data-step-max="${i}" size="4" value="${esc(s.max_wait || 30)}"></td>
      <td><input data-step-readcol="${i}" size="10" placeholder="e.g. invoiceRes" value="${esc(s.read_column || "")}"> &rarr; <input data-step-readfield="${i}" size="9" placeholder="invoiceres" value="${esc(s.read_field || "")}"><br><input data-step-readrow="${i}" size="10" placeholder="only row with e.g. LTD" title="Only the row(s) containing this text -- e.g. LTD when the screen has a row per company" value="${esc(s.read_row_with || "")}"></td>
      <td><input type="checkbox" data-step-capture="${i}" ${s.capture ? "checked" : ""}></td>
      <td><input type="checkbox" data-step-pages="${i}" ${s.capture_pages ? "checked" : ""}></td>
      <td><input type="checkbox" data-step-remove="${i}"></td></tr>`).join("")}
    </table><p class="row"><button class="primary" onclick="saveEdit()">Save</button><button onclick="$('editCard').style.display='none'">Close</button></p>
    <p class="row"><label>Add "for each order" after step <input id="foreachAfter" size="3" value="${a.steps.length}"></label><button onclick="addStep('foreach')">Add for each order</button>
      <label>Add a check step after step <input id="checkAfter" size="3" value="${a.steps.length}"></label><button onclick="addStep('check')">Add check</button></p>
    <p class="row"><label>Add a list loop after step <input id="loopAfter" size="3" value="${a.steps.length}"></label><button onclick="addLoop()">Add list loop</button>
      <span class="muted">e.g. go through the orders of a tour and Delete the ones that aren't flowers -- save your other changes first.</span></p>
    <p class="muted">Field: the typed text becomes a value you fill in when starting (leave empty to always type the recorded text). All pages: Page Down until the list ends. Key: change what a step presses (e.g. Left instead of Right) without recording again. Read column: after this step, read that column of the table on screen (e.g. invoiceRes) into a field -- use it in the dump file name as {invoiceres}; it is read twice and must agree, otherwise the run stops. Remove: drop the step. Find row with: in a list (e.g. the tours after F1) the run reads the list, finds the row with this value and moves the highlight there -- the position may differ every day. Arrowing in a list and pressing Enter while recording fills this in by itself (the first number of the row).<br>
    Waiting: after each step the run waits until Uniware shows that screen and has stopped changing (finished loading), up to Max wait -- then it stops with a screenshot. If data appears on the same screen later (e.g. after typing the tour), put a word that only shows once the data is there in "Wait until screen shows".</p>`;
}
function loopMode(loop) { return loop.collect_into ? "collect" : loop.use_list ? "daylist" : "words"; }
function showLoopMode(i, mode) {
  document.querySelectorAll(`[data-loop-show="${i}"]`).forEach((part) => {
    part.style.display = part.dataset.modes.split(" ").includes(mode) ? "" : "none";
  });
}
function loopEditor(loop, i) {
  const mode = loopMode(loop);
  const part = (modes, html) => `<div class="row" data-loop-show="${i}" data-modes="${modes}"${modes.split(" ").includes(mode) ? "" : ' style="display:none"'}>${html}</div>`;
  return `<div><strong>List loop</strong> -- goes through the list row by row</div>
    <div class="row"><label>What it does <select data-loop-mode="${i}" onchange="showLoopMode(${i}, this.value)">
      <option value="words" ${mode === "words" ? "selected" : ""}>press a key when the screen shows a word (e.g. Accessoire -> Delete)</option>
      <option value="collect" ${mode === "collect" ? "selected" : ""}>collect the orders into a day list (presses only Down)</option>
      <option value="daylist" ${mode === "daylist" ? "selected" : ""}>press a key for the orders with a result in a day list (e.g. 516 -> Enter)</option>
    </select></label></div>
    ${part("words", `<label>When the screen shows <input data-loop-match="${i}" size="22" placeholder="e.g. Accessoire, Planten" value="${esc((loop.match || []).join(", "))}"></label>
      <label>keep when it shows <input data-loop-keep="${i}" size="16" placeholder="e.g. Bloemen" value="${esc((loop.keep || []).join(", "))}"></label>`)}
    ${part("collect", `<label>Day list <input data-loop-collect="${i}" size="20" placeholder="e.g. accessoires vandaag" value="${esc(loop.collect_into || "")}"></label><span class="muted">emptied first, then every row is written down</span>`)}
    ${part("collect", `<label>Each row is known by <select data-loop-collectkey="${i}"><option value="order" ${loop.collect_key !== "lot" ? "selected" : ""}>its order number (7 digits after the customer code)</option><option value="lot" ${loop.collect_key === "lot" ? "selected" : ""}>its Connect lot (8 digits, 44...)</option></select></label>
      <label>also read from the screen <input data-loop-collectvalue="${i}" size="8" placeholder="e.g. EPP" value="${esc(loop.collect_value || "")}"></label><span class="muted">the number after this label (e.g. "EPP: [9100185445726]" in the bottom bar) for each row -- saved as a CSV (lot;epp) in the dump folder</span>`)}
    ${part("collect", `<label>per row open a popup with <input data-loop-popupkeys="${i}" size="10" placeholder="e.g. Shift+?" value="${esc(loop.row_popup_keys || "")}"></label>
      <label>and close it with <input data-loop-closekeys="${i}" size="6" placeholder="Esc" value="${esc(loop.row_close_keys || "")}"></label><span class="muted">every line the popup shows is kept (column "Info" in the CSV, and in the run's Ctrl+P dump next to its lot) -- slow, fine for a night action</span>`)}
    ${part("daylist", `<label>Day list <input data-loop-uselist="${i}" size="20" placeholder="e.g. accessoires vandaag" value="${esc(loop.use_list || "")}"></label><label>orders with result <input data-loop-uselabel="${i}" size="5" placeholder="516" value="${esc(loop.use_label || "")}"></label>`)}
    ${part("words daylist", `<label>press <input data-loop-key="${i}" size="18" placeholder="e.g. Delete, Right, Enter" value="${esc(loop.match_key || "Delete")}" title="Several keys one after another, separated by commas -- e.g. Delete, then Right and Enter to answer the confirmation"></label>`)}
    <div class="row"><label>to the next row press <input data-loop-next="${i}" size="7" value="${esc(loop.next_key || "Down")}"></label>
      <label>at most <input data-loop-max="${i}" size="4" value="${esc(loop.max_rows || 500)}"> rows</label></div>
    ${part("words daylist", `<label><input type="checkbox" data-loop-dry="${i}" ${loop.dry_run ? "checked" : ""}> <strong>trial run</strong> (only log what it would do)</label>`)}
    ${part("words daylist", `<label><input type="checkbox" data-loop-stays="${i}" ${loop.row_stays ? "checked" : ""}> the row stays in the list after the key (e.g. only its tour number is cleared)</label>`)}
    ${part("words daylist", `<label>only rows that contain <input data-loop-only="${i}" size="8" placeholder="e.g. 508" value="${esc(loop.only_rows_with || "")}"></label><span class="muted">other rows are passed by -- safe to run again</span>`)}
    <div class="muted">Stops at the end of the list (no new row after 3 tries), or when a row is still there after the key (a confirmation question?).</div>`;
}
function loopValues(i) {
  // Only the fields of the chosen kind of loop count; the others are cleared.
  const value = (name) => document.querySelector(`[data-loop-${name}="${i}"]`).value;
  const checked = (name) => document.querySelector(`[data-loop-${name}="${i}"]`).checked;
  const mode = value("mode");
  return {
    match: mode === "words" ? value("match").split(",") : [],
    keep: mode === "words" ? value("keep").split(",") : [],
    match_key: value("key"),
    next_key: value("next"),
    max_rows: value("max"),
    dry_run: mode !== "collect" && checked("dry"),
    row_stays: mode !== "collect" && checked("stays"),
    only_rows_with: mode !== "collect" ? value("only") : "",
    collect_into: mode === "collect" ? value("collect") : "",
    collect_key: mode === "collect" ? value("collectkey") : "order",
    collect_value: mode === "collect" ? value("collectvalue") : "",
    row_popup_keys: mode === "collect" ? value("popupkeys") : "",
    row_close_keys: mode === "collect" ? value("closekeys") : "",
    use_list: mode === "daylist" ? value("uselist") : "",
    use_label: mode === "daylist" ? value("uselabel") : "",
  };
}
function foreachEditor(settings, i) {
  return `<div><strong>For each order</strong> of a day list -- repeats the steps after it</div>
    <div class="row"><label>day list <input data-fe-list="${i}" size="18" placeholder="accessoires vandaag" value="${esc(settings.list || "")}"></label>
      <label>the next <input data-fe-count="${i}" size="3" value="${esc(settings.count || 1)}"> step(s)</label>
      <strong>= steps ${i + 2} to ${i + 1 + Number(settings.count || 1)}</strong>
      <label>order number as field <input data-fe-field="${i}" size="7" value="${esc(settings.field || "order")}"></label></div>
    <div class="row"><label><input type="checkbox" data-fe-open="${i}" ${settings.only_open !== false ? "checked" : ""}> only orders without a result yet (a stopped run continues where it was)</label></div>
    <div class="muted">In those steps, typed text {${esc(settings.field || "order")}} is the order number, e.g. Ctrl+F, then type {${esc(settings.field || "order")}}, Enter, F11, a check step, back.</div>`;
}
function classifyEditor(settings, i) {
  return `<div><strong>Check</strong> the current order's product lines (inside "for each order")</div>
    <div class="row"><label>a line shows <input data-cl-word="${i}" size="10" placeholder="SPICER" value="${esc(settings.word || "")}"></label>
      <label>-> result <input data-cl-if="${i}" size="5" placeholder="516" value="${esc(settings.label_if || "")}"></label>
      <label>lines without it -> <input data-cl-else="${i}" size="5" placeholder="515" value="${esc(settings.label_else || "")}"></label></div>
    <div class="muted">Both kinds of lines in one order -> warning "multiple types of accessoires" (no result, a person decides). Product lines = lines with a quantity like "1 x 50".</div>`;
}
function addStep(kind) {
  const after = Number($(kind === "foreach" ? "foreachAfter" : "checkAfter").value || 0);
  api(`/api/actions/${editing}/add-${kind === "foreach" ? "foreach" : "check"}`, { after }).then((payload) => { if (payload) edit(editing); });
}
function setDayLabel(list, key, label) { api(`/api/daylists/x/label`, { list, key, label }); }
function deleteDayList(list) { if (confirm(`Delete day list '${list}'?`)) api(`/api/daylists/x/delete`, { list }); }
function renderDayLists() {
  const lists = state.day_lists || [];
  $("dayLists").innerHTML = lists.length ? lists.map((list) => {
    const items = list.items || [];
    const counts = items.reduce((acc, item) => { const key = item.label || (item.warning ? "warning" : "open"); acc[key] = (acc[key] || 0) + 1; return acc; }, {});
    return `<h3>${esc(list.name)} <span class="muted">-- ${items.length} order(s): ${Object.entries(counts).map(([key, count]) => `${count} ${esc(key)}`).join(", ")}${list.updated ? ` -- ${esc(list.updated.replace("T", " "))}` : ""}</span> <button class="danger" onclick="deleteDayList('${esc(list.name)}')">Delete list</button></h3>
      <table><tr><th>Order / lot</th><th>Row</th><th>Type</th>${items.some((item) => item.value !== undefined) ? "<th>Value read</th>" : ""}<th>Result</th><th>Note</th><th>Done</th></tr>
      ${items.map((item) => `<tr${item.warning ? ' style="background:#fff7ed"' : ""}><td>${esc(item.key)}</td><td class="muted">${esc(item.row || "")}</td><td>${esc(item.type || "")}</td>${items.some((entry) => entry.value !== undefined) ? `<td>${esc(item.value || "-")}</td>` : ""}
        <td><input size="5" value="${esc(item.label || "")}" onchange="setDayLabel('${esc(list.name)}', '${esc(item.key)}', this.value)"></td>
        <td>${item.warning ? "<strong>" : ""}${esc(item.note || "")}${item.warning ? "</strong>" : ""}</td><td>${item.done ? esc(String(item.done).replace("T", " ")) : ""}</td></tr>`).join("")}</table>`;
  }).join("") : `<p class="muted">No day lists yet -- a list loop with "collect into day list" makes one.</p>`;
}
function addLoop() {
  const after = Number($("loopAfter").value || 0);
  api(`/api/actions/${editing}/add-loop`, { after }).then((payload) => { if (payload) edit(editing); });
}
function saveEdit() {
  const a = state.actions.find((x) => x.slug === editing);
  const steps = a.steps.map((s, i) => ({
    capture: document.querySelector(`[data-step-capture="${i}"]`).checked,
    capture_pages: document.querySelector(`[data-step-pages="${i}"]`).checked,
    wait_text: document.querySelector(`[data-step-wait="${i}"]`).value,
    max_wait: document.querySelector(`[data-step-max="${i}"]`).value,
    field: s.text !== undefined ? document.querySelector(`[data-step-field="${i}"]`).value : null,
    pick: s.key === "Enter" ? document.querySelector(`[data-step-pick="${i}"]`).value : null,
    opens_window: s.key ? document.querySelector(`[data-step-window="${i}"]`).checked : null,
    key: s.key ? document.querySelector(`[data-step-key="${i}"]`).value : null,
    read_column: document.querySelector(`[data-step-readcol="${i}"]`).value,
    read_field: document.querySelector(`[data-step-readfield="${i}"]`).value,
    read_row_with: document.querySelector(`[data-step-readrow="${i}"]`).value,
    remove: document.querySelector(`[data-step-remove="${i}"]`).checked,
    foreach: s.foreach ? {
      list: document.querySelector(`[data-fe-list="${i}"]`).value,
      count: document.querySelector(`[data-fe-count="${i}"]`).value,
      field: document.querySelector(`[data-fe-field="${i}"]`).value,
      only_open: document.querySelector(`[data-fe-open="${i}"]`).checked,
    } : null,
    classify: s.classify ? {
      word: document.querySelector(`[data-cl-word="${i}"]`).value,
      label_if: document.querySelector(`[data-cl-if="${i}"]`).value,
      label_else: document.querySelector(`[data-cl-else="${i}"]`).value,
    } : null,
    loop: s.loop ? loopValues(i) : null,
  }));
  // Show the saved steps right away: an Edit table that still showed the old
  // steps looked like nothing happened, and saving it again changed the
  // wrong (renumbered) steps.
  api(`/api/actions/${editing}/update`, { description: $("editDesc").value, export_dir: $("editExportDir").value, back_key: $("editBackKey").value, close_dump_window: $("editCloseWindow").checked, join_dumps: $("editJoinDumps").checked, export_name: $("editExportName").value, uniware_dump_dir: $("editUniwareDir").value, uniware_dump_pattern: $("editUniwarePattern").value, uniware_dump_wait: $("editUniwareWait").value, steps }).then((payload) => { if (payload) edit(editing); });
}
api("/api/state");
setInterval(() => { if (state && (state.recording.active || state.run.active || (state.plan_run && state.plan_run.active))) api("/api/state"); }, 1000);
// Now and then also when idle: a plan that started by itself shows up.
setInterval(() => { if (state && !document.activeElement.closest("input, select, textarea")) api("/api/state"); }, 15000);
</script></body></html>"""


class SingleServer(ThreadingHTTPServer):
    # Windows lets a second program share a port that allows reuse, and the
    # browser then talks to either one at random -- an old dashboard with
    # old code kept answering. Only one dashboard at a time.
    allow_reuse_address = False


def main():
    url = f"http://127.0.0.1:{PORT}"
    try:
        server = SingleServer(("127.0.0.1", PORT), Handler)
    except OSError:
        print(f"The dashboard is already running ({url}). Close its window first to restart it with the newest version.")
        webbrowser.open(url)
        input("Press Enter to close this window...")
        return
    stop_hotkey_listener()
    threading.Thread(target=plan_scheduler, daemon=True).start()
    print(f"Uniware actions dashboard: {url}  (close this window to stop it)")
    webbrowser.open(url)
    server.serve_forever()


if __name__ == "__main__":
    main()
