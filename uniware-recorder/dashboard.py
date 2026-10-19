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
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse

from pynput import keyboard

import actions

PORT = 8765
STATE = {"recorder": None, "run": None, "message": ""}
LOCK = threading.Lock()


def state_payload() -> dict:
    recorder = STATE["recorder"]
    run = STATE["run"]
    return {
        "actions": [
            {**action, "fields": actions.action_fields(action)}
            for action in actions.list_actions()
        ],
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
        elif getattr(key, "vk", None) == 0x51 and {"ctrl", "alt"} <= held and STATE["run"] and STATE["run"].active:
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
                    if (STATE["run"] and STATE["run"].active) or (STATE["recorder"] and STATE["recorder"].active):
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
                    if (STATE["run"] and STATE["run"].active) or (STATE["recorder"] and STATE["recorder"].active):
                        raise RuntimeError("Something is already running.")
                    action = actions.load_action(slug)
                    if not action:
                        raise RuntimeError("Action not found.")
                    run = actions.ActionRun(action, body.get("values") or {})
                    STATE["run"] = run
                    threading.Thread(target=run.run, daemon=True).start()
                    STATE["message"] = f"Running '{action['name']}' -- don't touch keyboard or mouse until it's done (Ctrl+Alt+Q stops)."
                elif path == "/api/run/stop":
                    if STATE["run"] and STATE["run"].active:
                        STATE["run"].stop_requested = True
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
                    for index, change in enumerate(body.get("steps") or []):
                        if index < len(action["steps"]):
                            step = action["steps"][index]
                            if step.get("loop") and isinstance(change.get("loop"), dict):
                                step["loop"] = actions.normalize_list_loop(change["loop"])
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
                                # A field name: typed text becomes "{field}",
                                # asked at start; empty: the recorded text.
                                field = re.sub(r"[^a-z0-9_]+", "_", str(change["field"]).lower()).strip("_")
                                step["text"] = f"{{{field}}}" if field else step.get("recorded_text", step["text"])
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
                            else:
                                step.pop("read_column", None)
                                step.pop("read_field", None)
                    removed = {index for index, change in enumerate(body.get("steps") or []) if change.get("remove")}
                    if removed:
                        action["steps"] = [step for index, step in enumerate(action["steps"]) if index not in removed]
                    actions.save_action(action)
                    STATE["message"] = f"Saved changes to '{action['name']}'."
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
</style></head><body>
<h1>Uniware actions</h1>
<div class="muted">Record an action once in Uniware, start it from here. Runs on this PC only.</div>
<div id="msg"></div>
<div class="card" id="recordCard"></div>
<div class="card" id="runCard"></div>
<div class="card"><h2>Actions</h2><div id="actions"></div></div>
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
  $("actions").innerHTML = state.actions.length ? `<table><tr><th>Action</th><th>Steps</th><th>Fields</th><th></th></tr>${state.actions.map((a) => `
      <tr><td><strong>${esc(a.name)}</strong><div class="muted">${esc(a.description || "")}</div><div class="muted">starts on: ${a.steps.length && a.steps[0].menu ? `any menu position (goes back with ${esc(a.back_key || "Esc")} if needed)` : esc(a.start_title || a.start_screen)}</div>${a.uniware_dump_dir ? `<div class="muted">dump file &rarr; ${esc(a.export_name || "{action}_{date}_{time}")}.csv</div>` : a.export_dir ? `<div class="muted">dump to: ${esc(a.export_dir)} &rarr; ${esc(a.export_name || "{action}_{date}_{time}")}.csv</div>` : ""}</td>
      <td>${a.steps.length} (${a.steps.filter((s) => s.capture).length} capture)</td>
      <td>${a.fields.map((f) => `<label>${esc(f)} <input data-field="${esc(a.slug)}:${esc(f)}" size="14"></label>`).join("<br>") || "-"}</td>
      <td class="row"><button class="primary" onclick="runAction('${a.slug}')">Start</button><button onclick="edit('${a.slug}')">Edit</button><button onclick="copyAction('${a.slug}')">Copy</button><button class="danger" onclick="del('${a.slug}')">Delete</button></td></tr>`).join("")}</table>`
    : `<p class="muted">No actions yet -- record one above.</p>`;
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
    <p class="row"><label><input type="checkbox" id="editCloseWindow" ${a.close_dump_window === false ? "" : "checked"}> close that window afterwards</label></p>
    <p class="row"><label>Uniware saves its dump in <input id="editUniwareDir" size="60" placeholder="paste the folder path (e.g. the OneDrive folder)" value="${esc(a.uniware_dump_dir || "")}"></label>
      <label>Its name looks like <input id="editUniwarePattern" size="14" value="${esc(a.uniware_dump_pattern || "dump.*.csv")}"></label>
      <label>Max wait (s) <input id="editUniwareWait" size="4" value="${esc(a.uniware_dump_wait || 180)}"></label></p>
    <p class="row"><label>Rename it to <input id="editExportName" size="28" placeholder="{action}_{date}_{time}" value="${esc(a.export_name || "")}"></label>
      <label>and put it in <input id="editExportDir" size="50" placeholder="e.g. the OneDrive folder" value="${esc(a.export_dir || "")}"></label></p>
    <p class="muted">At the end of the run it waits for the new dump file Uniware writes (e.g. dump.filipe.2026-10-14_16_57_30.csv), waits until it is completely written, then renames it, e.g. <code>tour-508_{date}</code> gives tour-508_20261014.csv. You can use {date}, {time}, {action} and the fields. Same name again = replaced. No new file within Max wait = the run fails and the file is left alone.</p>
    <table><tr><th>#</th><th>Step</th><th>Goes to</th><th>Field (typed text) / list row</th><th>Wait until screen shows</th><th>Max wait (s)</th><th>Read column (into field)</th><th>Capture screen</th><th>All pages</th><th>Remove</th></tr>
    ${a.steps.map((s, i) => `<tr><td>${i + 1}</td>
      <td>${s.loop ? loopEditor(s.loop, i) : s.text !== undefined ? `type <code>${esc(s.text)}</code>` : `${s.menu ? `menu <strong>${s.menu.map(esc).join(" &rsaquo; ")}</strong>, then ` : ""}<input data-step-key="${i}" value="${esc(s.key)}" size="9" title="The key this step presses -- change it (e.g. Left instead of Right) without recording again">${s.enter_on ? ` on "${esc(s.enter_on)}"` : ""}${s.key ? `<div><label><input type="checkbox" data-step-window="${i}" ${s.opens_window ? "checked" : ""}> opens the dump in a new window</label>${s.opens_window && s.opens_window.title ? `<div class="muted">recorded: ${esc(s.opens_window.title)}${s.opens_window.exe ? ` (${esc(s.opens_window.exe)})` : ""}</div>` : ""}</div>` : ""}`}</td>
      <td>${esc(s.to_title || s.to)}${s.recorded_wait ? `<div class="muted">took ${esc(s.recorded_wait)}s when recorded</div>` : ""}</td>
      <td>${s.text !== undefined ? `<input data-step-field="${i}" placeholder="e.g. invoice" value="${esc((s.text.match(/^\\{(.+)\\}$/) || [])[1] || "")}" size="12"><div class="muted">recorded: ${esc(s.recorded_text || s.text)}</div>` : s.key === "Enter" ? `find row with <input data-step-pick="${i}" placeholder="empty = selected row" value="${esc(s.pick || "")}" size="12">${s.pick_row ? `<div class="muted">recorded row: ${esc(s.pick_row)}</div>` : ""}` : ""}</td>
      <td><input data-step-wait="${i}" size="16" placeholder="(optional) e.g. Totaal" value="${esc(s.wait_text || "")}"></td>
      <td><input data-step-max="${i}" size="4" value="${esc(s.max_wait || 30)}"></td>
      <td><input data-step-readcol="${i}" size="10" placeholder="e.g. invoiceRes" value="${esc(s.read_column || "")}"> &rarr; <input data-step-readfield="${i}" size="9" placeholder="invoiceres" value="${esc(s.read_field || "")}"></td>
      <td><input type="checkbox" data-step-capture="${i}" ${s.capture ? "checked" : ""}></td>
      <td><input type="checkbox" data-step-pages="${i}" ${s.capture_pages ? "checked" : ""}></td>
      <td><input type="checkbox" data-step-remove="${i}"></td></tr>`).join("")}
    </table><p class="row"><button class="primary" onclick="saveEdit()">Save</button><button onclick="$('editCard').style.display='none'">Close</button></p>
    <p class="row"><label>Add a list loop after step <input id="loopAfter" size="3" value="${a.steps.length}"></label><button onclick="addLoop()">Add list loop</button>
      <span class="muted">e.g. go through the orders of a tour and Delete the ones that aren't flowers -- save your other changes first.</span></p>
    <p class="muted">Field: the typed text becomes a value you fill in when starting (leave empty to always type the recorded text). All pages: Page Down until the list ends. Key: change what a step presses (e.g. Left instead of Right) without recording again. Read column: after this step, read that column of the table on screen (e.g. invoiceRes) into a field -- use it in the dump file name as {invoiceres}; it is read twice and must agree, otherwise the run stops. Remove: drop the step. Find row with: in a list (e.g. the tours after F1) the run reads the list, finds the row with this value and moves the highlight there -- the position may differ every day. Arrowing in a list and pressing Enter while recording fills this in by itself (the first number of the row).<br>
    Waiting: after each step the run waits until Uniware shows that screen and has stopped changing (finished loading), up to Max wait -- then it stops with a screenshot. If data appears on the same screen later (e.g. after typing the tour), put a word that only shows once the data is there in "Wait until screen shows".</p>`;
}
function loopEditor(loop, i) {
  return `<div><strong>List loop</strong> -- goes through the list row by row</div>
    <div class="row"><label>When the screen shows <input data-loop-match="${i}" size="22" placeholder="e.g. Accessoire, Planten" value="${esc((loop.match || []).join(", "))}"></label>
      <label>press <input data-loop-key="${i}" size="18" placeholder="e.g. Delete, Right, Enter" value="${esc(loop.match_key || "Delete")}" title="Several keys one after another, separated by commas -- e.g. Delete, then Right and Enter to answer the confirmation"></label></div>
    <div class="row"><label>Keep when it shows <input data-loop-keep="${i}" size="16" placeholder="e.g. Bloemen" value="${esc((loop.keep || []).join(", "))}"></label>
      <label>otherwise press <input data-loop-next="${i}" size="7" value="${esc(loop.next_key || "Down")}"></label></div>
    <div class="row"><label>at most <input data-loop-max="${i}" size="4" value="${esc(loop.max_rows || 500)}"> rows</label>
      <label><input type="checkbox" data-loop-dry="${i}" ${loop.dry_run ? "checked" : ""}> trial run (only log what it would do)</label></div>
    <div class="row"><label><input type="checkbox" data-loop-stays="${i}" ${loop.row_stays ? "checked" : ""}> the row stays in the list after the key (e.g. only its tour number is cleared) -- then go on with the otherwise-key</label></div>
    <div class="row"><label>only rows that contain <input data-loop-only="${i}" size="8" placeholder="e.g. 508" value="${esc(loop.only_rows_with || "")}"></label><span class="muted">other rows (e.g. already cleared) are passed by -- safe to run again</span></div>
    <div class="muted">A row whose type can't be read clearly (none of these words) is skipped, never deleted. Stops at the end of the list, or when a row is still there after the key (a confirmation question?).</div>`;
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
    remove: document.querySelector(`[data-step-remove="${i}"]`).checked,
    loop: s.loop ? {
      match: document.querySelector(`[data-loop-match="${i}"]`).value.split(","),
      keep: document.querySelector(`[data-loop-keep="${i}"]`).value.split(","),
      match_key: document.querySelector(`[data-loop-key="${i}"]`).value,
      next_key: document.querySelector(`[data-loop-next="${i}"]`).value,
      max_rows: document.querySelector(`[data-loop-max="${i}"]`).value,
      dry_run: document.querySelector(`[data-loop-dry="${i}"]`).checked,
      row_stays: document.querySelector(`[data-loop-stays="${i}"]`).checked,
      only_rows_with: document.querySelector(`[data-loop-only="${i}"]`).value,
    } : null,
  }));
  // Show the saved steps right away: an Edit table that still showed the old
  // steps looked like nothing happened, and saving it again changed the
  // wrong (renumbered) steps.
  api(`/api/actions/${editing}/update`, { description: $("editDesc").value, export_dir: $("editExportDir").value, back_key: $("editBackKey").value, close_dump_window: $("editCloseWindow").checked, export_name: $("editExportName").value, uniware_dump_dir: $("editUniwareDir").value, uniware_dump_pattern: $("editUniwarePattern").value, uniware_dump_wait: $("editUniwareWait").value, steps }).then((payload) => { if (payload) edit(editing); });
}
api("/api/state");
setInterval(() => { if (state && (state.recording.active || state.run.active)) api("/api/state"); }, 1000);
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
    print(f"Uniware actions dashboard: {url}  (close this window to stop it)")
    webbrowser.open(url)
    server.serve_forever()


if __name__ == "__main__":
    main()
