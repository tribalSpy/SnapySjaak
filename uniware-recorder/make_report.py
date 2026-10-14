"""data/map.json -> data/map.html: every recorded Uniware screen, its text,
and which keys lead where (also as a diagram)."""
from __future__ import annotations

import html
import json
import re
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
MAP_PATH = DATA_DIR / "map.json"
REPORT_PATH = DATA_DIR / "map.html"


def mermaid_label(text: str) -> str:
    return text.replace('"', "'").replace("|", "/")[:60]


def shortest_paths(transitions: dict, start: str) -> dict:
    """screen id -> [(key, screen id), ...] from start, fewest key moves."""
    paths = {start: []}
    queue = [start]
    while queue:
        current = queue.pop(0)
        for key, targets in transitions.get(current, {}).items():
            for target in targets:
                if target not in paths:
                    paths[target] = paths[current] + [(key, target)]
                    queue.append(target)
    return paths


def menu_columns(entry: dict, main_items: set) -> list[list[str]]:
    """A menu screen's own options (the main-menu column left out), per
    column left to right."""
    layout = entry.get("layout") or []
    columns = {}
    for line in layout:
        if line["text"] in main_items:
            continue
        column = round(line["x"] / 60)
        columns.setdefault(column, []).append(line)
    return [[line["text"] for line in sorted(items, key=lambda item: item["y"])] for _, items in sorted(columns.items())]


def build_overview(screens: dict, transitions: dict) -> tuple[str, str]:
    """(html, markdown) menu overview: every menu with its options and the
    keys that get there from the main menu."""
    main = screens.get("prg-000")
    main_layout = (main or {}).get("layout") or []
    main_x = min((line["x"] for line in main_layout), default=0)
    main_items = {line["text"] for line in main_layout if abs(line["x"] - main_x) < 40}
    paths = shortest_paths(transitions, "prg-000") if main else {}
    html_parts = []
    md_parts = ["# Uniware menu overview", ""]
    if main_items:
        ordered = [line["text"] for line in sorted(main_layout, key=lambda line: line["y"]) if line["text"] in main_items]
        md_parts += ["## Hoofdmenu (Prg. 000)", "", *[f"- {item}" for item in ordered], ""]
        html_parts.append(f"<h3>Hoofdmenu (Prg. 000)</h3><ul>{''.join(f'<li>{html.escape(item)}</li>' for item in ordered)}</ul>")
    def sort_key(item):
        number = re.search(r"\d+", item[0])
        return (0 if item[0].startswith("prg-") else 1, int(number.group()) if number else 9999, item[0])
    for screen_id, entry in sorted(screens.items(), key=sort_key):
        if screen_id == "prg-000":
            continue
        route = " → ".join(f"{key} → {screens.get(target, {}).get('title', target)}" for key, target in paths.get(screen_id, [])) or "(route from the main menu not recorded)"
        columns = menu_columns(entry, main_items) if screen_id.startswith("prg-") else []
        md_parts += [f"## {entry['title']}", "", f"Route: {route}", ""]
        for column in columns:
            md_parts += [f"- {item}" for item in column] + [""]
        options = "".join(f"<ul>{''.join(f'<li>{html.escape(item)}</li>' for item in column)}</ul>" for column in columns)
        html_parts.append(f'<h3><a href="#{screen_id}">{html.escape(entry["title"])}</a></h3><p class="muted">Route: {html.escape(route)}</p><div class="row">{options}</div>')
    return "".join(html_parts), "\n".join(md_parts)


def main():
    data = json.loads(MAP_PATH.read_text(encoding="utf-8"))
    screens = data.get("screens", {})
    transitions = data.get("transitions", {})
    node_ids = {screen_id: f"s{index}" for index, screen_id in enumerate(screens)}

    graph = ["flowchart LR"]
    for screen_id, entry in screens.items():
        graph.append(f'  {node_ids[screen_id]}["{mermaid_label(entry["title"])}"]')
    for source, keys in transitions.items():
        for key, targets in keys.items():
            for target in targets:
                if source in node_ids and target in node_ids and source != target:
                    graph.append(f'  {node_ids[source]} -- "{mermaid_label(key)}" --> {node_ids[target]}')

    sections = []
    for screen_id, entry in sorted(screens.items(), key=lambda item: item[1]["title"].lower()):
        moves = []
        for key, targets in sorted(transitions.get(screen_id, {}).items()):
            for target, count in targets.items():
                target_title = screens.get(target, {}).get("title", target)
                same = " (stays on this screen)" if target == screen_id else ""
                moves.append(f'<li><kbd>{html.escape(key)}</kbd> &rarr; <a href="#{target}">{html.escape(target_title)}</a>{same} <span class="muted">x{count}</span></li>')
        text = "\n".join(entry.get("texts", [[]])[0])
        sections.append(f"""
<section id="{screen_id}">
  <h2>{html.escape(entry["title"])}</h2>
  <p class="muted">seen {entry.get("seen", 0)}x, first {entry.get("first_seen", "")}, last {entry.get("last_seen", "")}</p>
  <div class="row">
    <img src="screens/{screen_id}.png" alt="{html.escape(entry["title"])}">
    <div>
      <h3>Keys from here</h3>
      <ul>{"".join(moves) or "<li class='muted'>none recorded yet</li>"}</ul>
      <h3>On the screen</h3>
      <pre>{html.escape(text)}</pre>
    </div>
  </div>
</section>""")

    overview_html, overview_md = build_overview(screens, transitions)
    (DATA_DIR / "menu_overview.md").write_text(overview_md, encoding="utf-8")

    page = f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Uniware map</title>
<style>
  body {{ font-family: system-ui, sans-serif; margin: 16px; background: #f6f7f9; color: #17202a; }}
  section {{ background: #fff; border: 1px solid #d6dbe1; border-radius: 8px; padding: 12px 16px; margin: 12px 0; }}
  .row {{ display: flex; gap: 16px; flex-wrap: wrap; }}
  img {{ max-width: 640px; width: 100%; border: 1px solid #ccc; }}
  pre {{ background: #f3f4f6; padding: 8px; white-space: pre-wrap; }}
  kbd {{ background: #eef2ff; border: 1px solid #c7d2fe; border-radius: 4px; padding: 1px 6px; }}
  .muted {{ color: #6b7280; }}
</style></head><body>
<h1>Uniware map</h1>
<p class="muted">{len(screens)} screens, {sum(len(keys) for keys in transitions.values())} key moves recorded.</p>
<section><h2>Menu overview</h2>{overview_html}</section>
<section><h2>Diagram</h2><pre class="mermaid">{html.escape(chr(10).join(graph))}</pre></section>
{"".join(sections)}
<script type="module">
  import mermaid from "https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.esm.min.mjs";
  mermaid.initialize({{ startOnLoad: true }});
</script>
</body></html>"""
    REPORT_PATH.write_text(page, encoding="utf-8")
    print(f"Report written: {REPORT_PATH}")


if __name__ == "__main__":
    main()
