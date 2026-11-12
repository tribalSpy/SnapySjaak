"""The link to the SnappySjaak app, for actions that ask it for work or send
it files (e.g. the daily facturen for Finance Audit UKdocs).

The address and key come from snappysjaak.json next to this file (download
it in SnappySjaak: Finance Audit UKdocs -> "Recorder setup file", admin) or
from config.json ("snappysjaak_url", "snappysjaak_api_key"). Standard
library only.
"""
from __future__ import annotations

import base64
import json
import socket
import urllib.error
import urllib.request
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
TIMEOUT = 120
BATCH = 10  # PDFs per upload


def settings() -> dict:
    for name, url_key, key_key in (("snappysjaak.json", "server_url", "api_key"), ("config.json", "snappysjaak_url", "snappysjaak_api_key")):
        path = APP_DIR / name
        if not path.exists():
            continue
        try:
            data = json.loads(path.read_text(encoding="utf-8-sig"))
        except (OSError, json.JSONDecodeError):
            continue
        url = str(data.get(url_key) or "").strip().rstrip("/")
        key = str(data.get(key_key) or "").strip()
        if url and key:
            return {"url": url, "key": key}
    raise RuntimeError("No SnappySjaak link: put snappysjaak.json (SnappySjaak > Finance Audit UKdocs > Recorder setup file) in the uniware-recorder folder.")


def _request(method: str, path: str, body: dict | None = None) -> dict:
    link = settings()
    data = json.dumps(body).encode("utf-8") if body is not None else None
    request = urllib.request.Request(
        f"{link['url']}{path}",
        data=data,
        method=method,
        headers={"x-shadow-agent-key": link["key"], "content-type": "application/json", "accept": "application/json"},
    )
    try:
        with urllib.request.urlopen(request, timeout=TIMEOUT) as response:
            return json.loads(response.read().decode("utf-8") or "{}")
    except urllib.error.HTTPError as error:
        try:
            message = json.loads(error.read().decode("utf-8")).get("error") or ""
        except (ValueError, OSError):
            message = ""
        raise RuntimeError(f"SnappySjaak answered {error.code}: {message or error.reason}") from None
    except (urllib.error.URLError, socket.timeout, TimeoutError) as error:
        raise RuntimeError(f"SnappySjaak can't be reached ({link['url']}): {getattr(error, 'reason', error)}") from None


def needed_invoices(date: str = "") -> dict:
    """{"date": "2026-11-11", "invoices": [{"invoice_number", "customer", "category", ...}]}
    -- the invoice numbers of that day's sendings (default: the previous
    workday) that still have no factuur in Finance Audit."""
    query = f"?date={date}" if date else ""
    return _request("GET", f"/api/agent/finance-audit/needed-invoices{query}")


def upload_invoices(date: str, paths: list[Path], not_found: list[str], agent: str) -> dict:
    """Sends the PDFs (in batches) and the numbers Uniware didn't have; the
    app matches each PDF by the number inside it, as "Match factuur folder".
    Returns {"results": [...], "still_needed": [...]}."""
    results, still_needed = [], []
    batches = [paths[index:index + BATCH] for index in range(0, len(paths), BATCH)] or [[]]
    for number, batch in enumerate(batches):
        files = [{"file_name": path.name, "content_base64": base64.b64encode(path.read_bytes()).decode("ascii")} for path in batch]
        answer = _request("POST", "/api/agent/finance-audit/invoice-documents", {
            "date": date,
            "agent": agent,
            "files": files,
            # The numbers not found only once (with the first batch).
            "not_found": not_found if number == 0 else [],
        })
        results.extend(answer.get("results") or [])
        still_needed = answer.get("still_needed") or []
    return {"results": results, "still_needed": still_needed}
