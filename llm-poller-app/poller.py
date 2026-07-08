import base64
import json
import os
import socket
import sys
import tempfile
import time
import urllib.error
import urllib.request
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    import pythoncom
    import win32com.client
except ImportError:
    pythoncom = None
    win32com = None


APP_DIR = Path(__file__).resolve().parent
CONFIG_PATH = APP_DIR / "config.json"
EXAMPLE_CONFIG_PATH = APP_DIR / "config.example.json"


def load_json(path: Path):
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def env_or_config(env_name: str, config: dict, key: str, default=""):
    value = os.getenv(env_name)
    if value is not None and str(value).strip() != "":
        return value
    return config.get(key, default)


def load_config():
    config = load_json(CONFIG_PATH) if CONFIG_PATH.exists() else load_json(EXAMPLE_CONFIG_PATH)
    loaded = {
        "server_url": str(env_or_config("SHADOW_POLLER_SERVER_URL", config, "server_url", "")).rstrip("/"),
        "api_key": str(env_or_config("SHADOW_POLLER_API_KEY", config, "api_key", "")).strip(),
        "agent_name": str(env_or_config("SHADOW_POLLER_AGENT_NAME", config, "agent_name", "")).strip(),
        "pc_name": str(env_or_config("SHADOW_POLLER_PC_NAME", config, "pc_name", socket.gethostname())).strip(),
        "model_name": str(env_or_config("SHADOW_POLLER_MODEL_NAME", config, "model_name", "")).strip(),
        "version": str(env_or_config("SHADOW_POLLER_VERSION", config, "version", "1.0.0")).strip(),
        "poll_interval_seconds": int(env_or_config("SHADOW_POLLER_INTERVAL_SECONDS", config, "poll_interval_seconds", 10) or 10),
        "ollama_url": str(env_or_config("OLLAMA_URL", config, "ollama_url", "http://127.0.0.1:11434")).rstrip("/"),
        "capabilities": config.get("capabilities", ["ollama_chat"]),
    }
    if not isinstance(loaded["capabilities"], list):
        loaded["capabilities"] = ["ollama_chat"]
    return loaded


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
        with urllib.request.urlopen(request, timeout=120) as response:
            body = response.read().decode("utf-8")
            return json.loads(body) if body else {}
    except urllib.error.HTTPError as error:
        body = error.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {error.code}: {body}") from error
    except urllib.error.URLError as error:
        raise RuntimeError(f"Network error: {error}") from error


def ollama_chat(config: dict, payload: dict):
    model_name = str(payload.get("model") or config["model_name"]).strip()
    if not model_name:
        raise RuntimeError("No model configured for ollama_chat job")
    messages = prepare_ollama_messages(payload.get("messages", []), payload)
    request_body = {
        "model": model_name,
        "messages": messages,
        "stream": False,
    }
    if payload.get("format"):
        request_body["format"] = payload["format"]
    if isinstance(payload.get("options"), dict):
        request_body["options"] = payload["options"]
    if payload.get("keep_alive"):
        request_body["keep_alive"] = payload["keep_alive"]

    request = urllib.request.Request(
        url=f'{config["ollama_url"]}/api/chat',
        method="POST",
        headers={
            "Accept": "application/json",
            "Content-Type": "application/json",
        },
        data=json.dumps(request_body).encode("utf-8"),
    )
    try:
        with urllib.request.urlopen(request, timeout=600) as response:
            body = response.read().decode("utf-8")
            return json.loads(body) if body else {}
    except urllib.error.HTTPError as error:
        body = error.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Ollama HTTP {error.code}: {body}") from error
    except urllib.error.URLError as error:
        raise RuntimeError(f"Ollama network error: {error}") from error


def excel_to_pdf(payload: dict):
    if sys.platform != "win32":
        raise RuntimeError("Excel PDF export requires Windows")
    if win32com is None:
        raise RuntimeError("pywin32 is required for Excel PDF export")

    workbook_base64 = str(payload.get("workbook_content_base64") or "").strip()
    workbook_name = str(payload.get("workbook_name") or "invoice.xlsx").strip() or "invoice.xlsx"
    pdf_name = str(payload.get("pdf_name") or "").strip() or f'{Path(workbook_name).stem}.pdf'
    if not workbook_base64:
        raise RuntimeError("No workbook_content_base64 provided for Excel PDF export")

    temp_dir = Path(tempfile.mkdtemp(prefix="shadow-excel-pdf-"))
    workbook_path = temp_dir / Path(workbook_name).name
    pdf_path = temp_dir / Path(pdf_name).name
    workbook_path.write_bytes(base64.b64decode(workbook_base64))

    if pythoncom is not None:
        pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(str(workbook_path.resolve()))
        workbook.ExportAsFixedFormat(0, str(pdf_path.resolve()))
        workbook.Close(False)
        workbook = None
        excel.Quit()
        excel = None
        pdf_bytes = pdf_path.read_bytes()
        return {
            "file_name": pdf_path.name,
            "mime_type": "application/pdf",
            "content_base64": base64.b64encode(pdf_bytes).decode("ascii"),
            "source_storage_name": str(payload.get("source_storage_name") or "").strip(),
            "source_original_name": workbook_name,
            "category": str(payload.get("category") or "").strip(),
            "document_kind": "invoice",
        }
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        for candidate in [pdf_path, workbook_path]:
            try:
                if candidate.exists():
                    candidate.unlink()
            except Exception:
                pass
        try:
            temp_dir.rmdir()
        except Exception:
            pass


def render_pdf_to_png_base64_list(content_base64: str, max_pages: int = 2):
    if fitz is None:
        raise RuntimeError("PyMuPDF is required for PDF vision extraction")
    import base64

    pdf_bytes = base64.b64decode(content_base64)
    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    page_limit = max(1, int(max_pages or 1))
    for page_index in range(min(len(document), page_limit)):
        page = document.load_page(page_index)
        pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        images.append(base64.b64encode(pixmap.tobytes("png")).decode("ascii"))
    return images


def prepare_vision_images(payload: dict):
    vision_documents = payload.get("vision_documents")
    if not isinstance(vision_documents, list) or not vision_documents:
        return [], []

    prepared_images = []
    notes = []
    for index, document in enumerate(vision_documents, start=1):
        mime_type = str(document.get("mime_type") or "").strip().lower()
        name = str(document.get("name") or f"document-{index}").strip()
        content_base64 = str(document.get("content_base64") or "").strip()
        if not content_base64:
            notes.append(f"{name}: missing file content")
            continue
        if mime_type == "application/pdf":
            max_pages = int(document.get("max_pages") or 2)
            try:
                pdf_images = render_pdf_to_png_base64_list(content_base64, max_pages=max_pages)
                for image_base64 in pdf_images:
                    prepared_images.append(image_base64)
                notes.append(f"{name}: rendered {len(pdf_images)} PDF page image(s)")
            except Exception as error:
                notes.append(f"{name}: PDF vision render failed ({error})")
            continue
        if mime_type.startswith("image/"):
            prepared_images.append(content_base64)
            notes.append(f"{name}: attached directly as image")
            continue
        notes.append(f"{name}: unsupported vision mime type {mime_type or 'unknown'}")
    return prepared_images, notes


def prepare_ollama_messages(messages, payload: dict):
    safe_messages = list(messages) if isinstance(messages, list) else []
    vision_images, vision_notes = prepare_vision_images(payload)
    if not vision_images and not vision_notes:
        return safe_messages

    if safe_messages and isinstance(safe_messages[-1], dict) and str(safe_messages[-1].get("role") or "").strip() == "user":
        updated_last = dict(safe_messages[-1])
        existing_content = str(updated_last.get("content") or "").strip()
        note_text = ""
        if vision_notes:
            note_text = "\n\nVision attachment notes:\n- " + "\n- ".join(vision_notes)
        updated_last["content"] = f"{existing_content}{note_text}".strip()
        if vision_images:
            existing_images = updated_last.get("images")
            if isinstance(existing_images, list):
                updated_last["images"] = existing_images + vision_images
            else:
                updated_last["images"] = vision_images
        safe_messages[-1] = updated_last
        return safe_messages

    content = "Use the attached images for the temporary phyto PDF review."
    if vision_notes:
        content += "\n\nVision attachment notes:\n- " + "\n- ".join(vision_notes)
    safe_messages.append({
        "role": "user",
        "content": content,
        **({"images": vision_images} if vision_images else {}),
    })
    return safe_messages


def run_job(config: dict, job: dict):
    job_type = str(job.get("job_type") or "").strip()
    payload = job.get("payload_json") or {}
    if job_type in {"ollama_chat", "ukdocs_csi_audit"}:
        ollama_response = ollama_chat(config, payload)
        return {
            "job_type": job_type,
            "model": payload.get("model") or config["model_name"],
            "ollama_response": ollama_response,
        }
    if job_type == "excel_to_pdf":
        export_result = excel_to_pdf(payload)
        return {
            "job_type": job_type,
            "excel_pdf_result": export_result,
        }
    raise RuntimeError(f"Unsupported job type: {job_type}")


def heartbeat_payload(config: dict, status: str):
    return {
        "agent_name": config["agent_name"],
        "pc_name": config["pc_name"],
        "model_name": config["model_name"],
        "version": config["version"],
        "status": status,
        "capabilities": config["capabilities"],
        "meta": {
            "python": sys.version.split()[0],
            "platform": sys.platform,
        },
    }


def validate_config(config: dict):
    missing = []
    for key in ["server_url", "api_key", "agent_name"]:
        if not str(config.get(key) or "").strip():
            missing.append(key)
    if missing:
        raise RuntimeError(f"Missing config values: {', '.join(missing)}")


def main():
    config = load_config()
    validate_config(config)
    print(f'Poller starting for agent {config["agent_name"]} -> {config["server_url"]}')
    while True:
        try:
            api_request(config, "POST", "/api/llm/agent/heartbeat", heartbeat_payload(config, "online"))
            poll_response = api_request(config, "POST", "/api/llm/agent/poll", heartbeat_payload(config, "online"))
            job = poll_response.get("job")
            if not job:
                time.sleep(max(2, int(config["poll_interval_seconds"])))
                continue

            job_id = str(job.get("id") or "").strip()
            print(f"Claimed job {job_id} ({job.get('job_type')})")
            try:
                result_json = run_job(config, job)
                api_request(
                    config,
                    "POST",
                    f"/api/llm/jobs/{job_id}/result",
                    {
                        "agent_name": config["agent_name"],
                        "pc_name": config["pc_name"],
                        "model_name": config["model_name"],
                        "version": config["version"],
                        "status": "idle",
                        "capabilities": config["capabilities"],
                        "result_json": result_json,
                    },
                )
                print(f"Finished job {job_id}")
            except Exception as error:
                api_request(
                    config,
                    "POST",
                    f"/api/llm/jobs/{job_id}/fail",
                    {
                        "agent_name": config["agent_name"],
                        "pc_name": config["pc_name"],
                        "model_name": config["model_name"],
                        "version": config["version"],
                        "status": "idle",
                        "capabilities": config["capabilities"],
                        "error_text": str(error),
                        "allow_retry": False,
                    },
                )
                print(f"Failed job {job_id}: {error}")
        except KeyboardInterrupt:
            print("Poller stopped by user")
            return
        except Exception as error:
            print(f"Poller loop error: {error}")
            time.sleep(max(5, int(config["poll_interval_seconds"])))


if __name__ == "__main__":
    main()
