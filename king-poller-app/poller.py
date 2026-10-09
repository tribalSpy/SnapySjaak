import base64
import json
import os
import socket
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

# Delivers King ERP exports from the shadow-app (Inkoop Controle > "Import
# naar King") into the folder King's scheduler reads, on the office network.
# Same pull model as shelf-poller-app: it only ever asks for, and can only
# ever be handed, "king_export" jobs (see claimNextLlmJob in shadow-app's
# server/db.js -- the generic poller never gets this type).
#
# Per job, in this order:
#   1. the invoice PDFs into pdf_dir (the folder King's archive XML points
#      at -- the same folder as the "King PDF folder" setting in the app,
#      seen from this PC),
#   2. the archive XML (KING_DIGITAAL_ARCHIEF) into import_dir,
#   3. wait until King has read it (the file disappears from import_dir),
#      up to archive_wait_minutes, so the archive items exist before
#   4. the journal XML (KING_JOURNAAL) is written -- King links each journal
#      line to its archive item through JR_ARCHIEFSTUK_EXTERN_ID.
# Every file is written as "<name>.part" and renamed when complete, so the
# scheduler never picks up a half-written file.

APP_DIR = Path(__file__).resolve().parent
CONFIG_PATH = APP_DIR / "config.json"
EXAMPLE_CONFIG_PATH = APP_DIR / "config.example.json"
# Written into the download from the app (server_url + api_key), so the
# installer only has to ask for the folders. config.json wins over it.
DEFAULTS_CONFIG_PATH = APP_DIR / "config.defaults.json"

JOB_TYPES = ["king_export"]


def load_json(path: Path):
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8-sig"))


def env_or_config(env_name: str, config: dict, key: str, default=""):
    value = os.getenv(env_name)
    if value is not None and str(value).strip() != "":
        return value
    return config.get(key, default)


def load_config():
    if CONFIG_PATH.exists() or DEFAULTS_CONFIG_PATH.exists():
        config = {**load_json(DEFAULTS_CONFIG_PATH), **load_json(CONFIG_PATH)}
    else:
        config = load_json(EXAMPLE_CONFIG_PATH)
    return {
        "server_url": str(env_or_config("KING_POLLER_SERVER_URL", config, "server_url", "")).rstrip("/"),
        "api_key": str(env_or_config("KING_POLLER_API_KEY", config, "api_key", "")).strip(),
        "agent_name": str(env_or_config("KING_POLLER_AGENT_NAME", config, "agent_name", "")).strip(),
        "pc_name": str(env_or_config("KING_POLLER_PC_NAME", config, "pc_name", socket.gethostname())).strip(),
        "version": str(env_or_config("KING_POLLER_VERSION", config, "version", "1.0.0")).strip(),
        "poll_interval_seconds": int(env_or_config("KING_POLLER_INTERVAL_SECONDS", config, "poll_interval_seconds", 30) or 30),
        "import_dir": str(env_or_config("KING_IMPORT_DIR", config, "import_dir", "")).strip(),
        "pdf_dir": str(env_or_config("KING_PDF_DIR", config, "pdf_dir", "")).strip(),
        "archive_wait_minutes": float(env_or_config("KING_ARCHIVE_WAIT_MINUTES", config, "archive_wait_minutes", 30) or 30),
        # The PDF folder as King's server sees it (shown in the app, which
        # offers it for the "King PDF folder" setting); empty = same as pdf_dir.
        "king_pdf_dir": str(env_or_config("KING_PDF_DIR_FOR_KING", config, "king_pdf_dir", "")).strip(),
    }


def api_request(config: dict, method: str, path: str, payload=None):
    data = None
    headers = {
        "Accept": "application/json",
        "x-shadow-agent-key": config["api_key"],
    }
    if payload is not None:
        data = json.dumps(payload).encode("utf-8")
        headers["Content-Type"] = "application/json"
    request = urllib.request.Request(
        url=f'{config["server_url"]}{path}',
        method=method,
        headers=headers,
        data=data,
    )
    try:
        with urllib.request.urlopen(request, timeout=180) as response:
            body = response.read().decode("utf-8")
            return json.loads(body) if body else {}
    except urllib.error.HTTPError as error:
        body = error.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {error.code}: {body}") from error
    except urllib.error.URLError as error:
        raise RuntimeError(f"Network error: {error}") from error


def safe_file_name(name: str) -> str:
    cleaned = "".join(ch for ch in str(name or "") if ch.isalnum() or ch in "._-")
    if not cleaned or cleaned.startswith("."):
        raise RuntimeError(f"Refusing unsafe file name: {name!r}")
    return cleaned


def write_atomic(folder: Path, name: str, content: bytes) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / safe_file_name(name)
    partial = target.with_name(target.name + ".part")
    partial.write_bytes(content)
    os.replace(partial, target)
    return target


def wait_until_gone(path: Path, minutes: float) -> bool:
    deadline = time.time() + max(0.0, minutes) * 60
    while time.time() < deadline:
        if not path.exists():
            return True
        time.sleep(15)
    return not path.exists()


# The XML names are fixed (king_journaal.xml / king_archief.xml), so the
# previous export's file may still be waiting for King's scheduler. Never
# overwrite it: wait until King has read it; still there after the wait ->
# fail (the server retries the job later).
def wait_for_free_name(folder: Path, name: str, minutes: float):
    target = folder / name
    if target.exists():
        print(f"  {name} from a previous export is still waiting for King -- waiting")
        if not wait_until_gone(target, minutes):
            raise RuntimeError(f"{name} is still in {folder} (King hasn't read the previous export yet); not overwriting it -- will retry")


def run_job(config: dict, job: dict):
    payload = job.get("payload_json") or {}
    import_dir = Path(config["import_dir"])
    pdf_dir = Path(config["pdf_dir"])
    written = []
    warnings = []

    for pdf in payload.get("pdfs") or []:
        content = base64.b64decode(str(pdf.get("content_base64") or ""))
        if not content:
            raise RuntimeError(f"Empty PDF in job: {pdf.get('file_name')}")
        written.append(str(write_atomic(pdf_dir, pdf.get("file_name"), content)))

    archief_name = safe_file_name(payload.get("archief_file_name") or "king_archief.xml")
    journaal_name = safe_file_name(payload.get("journaal_file_name") or "king_journaal.xml")
    wait_for_free_name(import_dir, archief_name, config["archive_wait_minutes"])
    wait_for_free_name(import_dir, journaal_name, config["archive_wait_minutes"])
    archief_path = write_atomic(import_dir, archief_name, str(payload.get("archief_xml") or "").encode("utf-8"))
    written.append(str(archief_path))
    print(f"  archive XML written: {archief_path} -- waiting for King to read it")
    if not wait_until_gone(archief_path, config["archive_wait_minutes"]):
        warnings.append(
            f"King had not read {archief_path.name} after {config['archive_wait_minutes']} min; "
            "journal written anyway -- check in King that the PDFs are attached."
        )

    journaal_path = write_atomic(import_dir, journaal_name, str(payload.get("journaal_xml") or "").encode("utf-8"))
    written.append(str(journaal_path))
    return {
        "delivered": True,
        "batch_id": payload.get("batch_id"),
        "written": written,
        "warnings": warnings,
    }


def heartbeat_payload(config: dict, status: str):
    return {
        "agent_name": config["agent_name"],
        "pc_name": config["pc_name"],
        "model_name": "",
        "version": config["version"],
        "status": status,
        "capabilities": JOB_TYPES,
        "meta": {
            "python": sys.version.split()[0],
            "platform": sys.platform,
            "import_dir": config["import_dir"],
            "pdf_dir": config["pdf_dir"],
            "king_pdf_dir": config["king_pdf_dir"] or config["pdf_dir"],
        },
    }


def validate_config(config: dict):
    missing = [key for key in ["server_url", "api_key", "agent_name", "import_dir", "pdf_dir"] if not str(config.get(key) or "").strip()]
    if missing:
        raise RuntimeError(f"Missing config values: {', '.join(missing)}")
    for key in ["import_dir", "pdf_dir"]:
        folder = Path(config[key])
        if not folder.exists():
            raise RuntimeError(f"{key} does not exist or is not reachable from this PC: {folder}")


# `python poller.py --check`: what the installer runs -- config complete,
# both folders writable from this PC, and the app accepts this PC (one
# heartbeat). Prints one line per check; exit code 0 only when all pass.
def run_check() -> int:
    ok = True
    try:
        config = load_config()
        validate_config(config)
        print("OK   settings complete")
    except Exception as error:
        print(f"FAIL settings: {error}")
        return 1
    for key, label in [("import_dir", "King import folder"), ("pdf_dir", "PDF folder")]:
        test_file = Path(config[key]) / f".king-poller-test-{os.getpid()}.tmp"
        try:
            test_file.write_text("test", encoding="utf-8")
            test_file.unlink()
            print(f"OK   {label} is writable: {config[key]}")
        except Exception as error:
            ok = False
            print(f"FAIL {label} is not writable from this PC ({config[key]}): {error}")
    try:
        api_request(config, "POST", "/api/llm/agent/heartbeat", heartbeat_payload(config, "online"))
        print(f"OK   connected to {config['server_url']}")
    except Exception as error:
        ok = False
        print(f"FAIL cannot reach the app at {config['server_url']}: {error}")
    return 0 if ok else 1


def main():
    if "--check" in sys.argv:
        sys.exit(run_check())
    config = load_config()
    validate_config(config)
    print(f'King poller starting for agent {config["agent_name"]} -> {config["server_url"]}')
    print(f'  import folder: {config["import_dir"]}  pdf folder: {config["pdf_dir"]}')
    while True:
        try:
            poll_response = api_request(
                config,
                "POST",
                "/api/llm/agent/poll",
                {**heartbeat_payload(config, "online"), "job_types": JOB_TYPES},
            )
            job = poll_response.get("job")
            if not job:
                time.sleep(max(5, int(config["poll_interval_seconds"])))
                continue

            job_id = str(job.get("id") or "").strip()
            print(f"Claimed job {job_id} ({job.get('job_type')})")
            base = {
                "agent_name": config["agent_name"],
                "pc_name": config["pc_name"],
                "model_name": "",
                "version": config["version"],
                "status": "idle",
                "capabilities": JOB_TYPES,
            }
            try:
                result_json = run_job(config, job)
                api_request(config, "POST", f"/api/llm/jobs/{job_id}/result", {**base, "result_json": result_json})
                print(f"Delivered job {job_id}")
            except Exception as error:
                # Retry allowed: a share that's briefly unreachable shouldn't
                # lose the export (the server stops after max_attempts).
                api_request(config, "POST", f"/api/llm/jobs/{job_id}/fail", {**base, "error_text": str(error), "allow_retry": True})
                print(f"Failed job {job_id}: {error}")
        except KeyboardInterrupt:
            print("Poller stopped by user")
            return
        except Exception as error:
            print(f"Poller loop error: {error}")
            time.sleep(max(10, int(config["poll_interval_seconds"])))


if __name__ == "__main__":
    main()
