import base64
import json
import os
import socket
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
from pathlib import Path

# A dedicated poller for exactly one job type ("shelf_count") -- kept
# deliberately separate from llm-poller-app/poller.py rather than adding a
# branch there, so the invoice-PDF/CSI-audit poller stays completely
# untouched. This can run on the same PC as that poller (same Ollama
# instance, same GPU) as a second, independent process.
#
# It always sends job_types=["shelf_count"] on every poll -- the server
# only ever hands a "shelf_count" job to a poller that asks for it by name
# (see claimNextLlmJob/the /api/llm/agent/poll route in shadow-app), so this
# poller can never accidentally steal a job meant for the other one, and
# the other poller (which never asks for "shelf_count") can never
# accidentally claim and fail one of these either.

APP_DIR = Path(__file__).resolve().parent
CONFIG_PATH = APP_DIR / "config.json"
EXAMPLE_CONFIG_PATH = APP_DIR / "config.example.json"

JOB_TYPES = ["shelf_count"]


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
    return {
        "server_url": str(env_or_config("SHELF_POLLER_SERVER_URL", config, "server_url", "")).rstrip("/"),
        "api_key": str(env_or_config("SHELF_POLLER_API_KEY", config, "api_key", "")).strip(),
        "agent_name": str(env_or_config("SHELF_POLLER_AGENT_NAME", config, "agent_name", "")).strip(),
        "pc_name": str(env_or_config("SHELF_POLLER_PC_NAME", config, "pc_name", socket.gethostname())).strip(),
        "model_name": str(env_or_config("SHELF_POLLER_MODEL_NAME", config, "model_name", "")).strip(),
        "version": str(env_or_config("SHELF_POLLER_VERSION", config, "version", "1.0.0")).strip(),
        "poll_interval_seconds": int(env_or_config("SHELF_POLLER_INTERVAL_SECONDS", config, "poll_interval_seconds", 10) or 10),
        "ollama_url": str(env_or_config("OLLAMA_URL", config, "ollama_url", "http://127.0.0.1:11434")).rstrip("/"),
        # Optional side-by-side count with the trained YOLO model
        # (shelf-training/models/infer.py, run with the training venv's own
        # python so this poller never needs torch/ultralytics). Empty = off.
        "trained_model_python": str(env_or_config("SHELF_TRAINED_MODEL_PYTHON", config, "trained_model_python", "")).strip(),
        "trained_model_infer": str(env_or_config("SHELF_TRAINED_MODEL_INFER", config, "trained_model_infer", "")).strip(),
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
        with urllib.request.urlopen(request, timeout=120) as response:
            body = response.read().decode("utf-8")
            return json.loads(body) if body else {}
    except urllib.error.HTTPError as error:
        body = error.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {error.code}: {body}") from error
    except urllib.error.URLError as error:
        raise RuntimeError(f"Network error: {error}") from error


def prepare_vision_images(payload: dict):
    # Shelf-count photos are always plain camera images (jpg/png) straight
    # off the RFID scan portal, never PDFs, so this is simpler than the
    # other poller's version -- no PDF-page rendering needed.
    vision_documents = payload.get("vision_documents")
    if not isinstance(vision_documents, list) or not vision_documents:
        return [], []

    prepared_images = []
    notes = []
    for index, document in enumerate(vision_documents, start=1):
        mime_type = str(document.get("mime_type") or "").strip().lower()
        name = str(document.get("name") or f"photo-{index}").strip()
        content_base64 = str(document.get("content_base64") or "").strip()
        if not content_base64:
            notes.append(f"{name}: missing file content")
            continue
        if mime_type.startswith("image/"):
            prepared_images.append(content_base64)
            continue
        notes.append(f"{name}: unsupported mime type {mime_type or 'unknown'} (expected an image)")
    fitted, fit_notes = fit_images_to_context(prepared_images, payload)
    return fitted, notes + fit_notes


# Vision models spend roughly one token per 32x32 pixel block of a photo
# (Qwen-VL style): a 1080x1920 trolley photo is ~2,000 tokens. Confirmed
# real: a 20-photo reference asked for 41,469 tokens against Ollama's 32,768
# context and failed. All photos together must stay under this budget,
# leaving room for the prompt and the answer.
IMAGE_TOKEN_BUDGET = 24000
TOKENS_PER_PIXEL = 1 / (32 * 32)


def jpeg_or_png_size(data: bytes):
    """(width, height) from the file header, stdlib only; None if unknown."""
    if data[:8] == bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]) and len(data) >= 24:
        return int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big")
    if data[:2] != bytes([0xFF, 0xD8]):
        return None
    index = 2
    while index + 9 < len(data):
        if data[index] != 0xFF:
            index += 1
            continue
        marker = data[index + 1]
        if marker in (0xC0, 0xC1, 0xC2):
            return int.from_bytes(data[index + 7:index + 9], "big"), int.from_bytes(data[index + 5:index + 7], "big")
        length = int.from_bytes(data[index + 2:index + 4], "big")
        index += 2 + length
    return None


def fit_images_to_context(images_base64, payload: dict):
    """Keeps all photos together under IMAGE_TOKEN_BUDGET: scales every photo
    down by the same factor (Pillow), or -- without Pillow -- sends an even
    selection of the photos that fits, and says so in the prompt."""
    if not images_base64:
        return images_base64, []
    budget = int((payload.get("options") or {}).get("image_token_budget") or IMAGE_TOKEN_BUDGET)
    raw = [base64.b64decode(item) for item in images_base64]
    sizes = [jpeg_or_png_size(item) or (1080, 1920) for item in raw]
    tokens = [width * height * TOKENS_PER_PIXEL for width, height in sizes]
    total = sum(tokens)
    if total <= budget:
        return images_base64, []
    scale = (budget / total) ** 0.5
    try:
        from io import BytesIO
        from PIL import Image  # optional: pip install pillow
    except ImportError:
        keep = max(1, int(len(raw) * budget / total))
        step = len(raw) / keep
        chosen = sorted({int(i * step) for i in range(keep)})
        return [images_base64[i] for i in chosen], [
            f"Only {len(chosen)} of the {len(raw)} photos are attached (evenly spread) -- the rest did not fit; "
            "they show the same trolleys from other angles."
        ]
    resized = []
    for data, (width, height) in zip(raw, sizes):
        image = Image.open(BytesIO(data))
        image = image.convert("RGB")
        target = (max(32, int(width * scale)), max(32, int(height * scale)))
        image = image.resize(target, Image.LANCZOS)
        out = BytesIO()
        image.save(out, format="JPEG", quality=90)
        resized.append(base64.b64encode(out.getvalue()).decode("ascii"))
    return resized, []


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
            note_text = "\n\nPhoto attachment notes:\n- " + "\n- ".join(vision_notes)
        updated_last["content"] = f"{existing_content}{note_text}".strip()
        if vision_images:
            existing_images = updated_last.get("images")
            if isinstance(existing_images, list):
                updated_last["images"] = existing_images + vision_images
            else:
                updated_last["images"] = vision_images
        safe_messages[-1] = updated_last
        return safe_messages

    content = "Use the attached photos to count shelves and levels."
    if vision_notes:
        content += "\n\nPhoto attachment notes:\n- " + "\n- ".join(vision_notes)
    safe_messages.append({
        "role": "user",
        "content": content,
        **({"images": vision_images} if vision_images else {}),
    })
    return safe_messages


def ollama_chat(config: dict, payload: dict):
    model_name = str(payload.get("model") or config["model_name"]).strip()
    if not model_name:
        raise RuntimeError("No model configured for shelf_count job")
    messages = prepare_ollama_messages(payload.get("messages", []), payload)
    request_body = {
        "model": model_name,
        "messages": messages,
        "stream": False,
    }
    if payload.get("format"):
        request_body["format"] = payload["format"]
    if isinstance(payload.get("think"), bool):
        request_body["think"] = payload["think"]
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


def run_job(config: dict, job: dict):
    job_type = str(job.get("job_type") or "").strip()
    if job_type != "shelf_count":
        # Should never happen -- the server only ever hands this poller a
        # shelf_count job (see JOB_TYPES above) -- but fail loudly rather
        # than silently mishandling an unexpected type if that ever changes.
        raise RuntimeError(f"Unexpected job type for the shelf-count poller: {job_type}")
    payload = job.get("payload_json") or {}
    ollama_response = ollama_chat(config, payload)
    result = {
        "job_type": job_type,
        "model": payload.get("model") or config["model_name"],
        "ollama_response": ollama_response,
    }
    trained = run_trained_model(config, payload)
    if trained is not None:
        result["trained_model"] = trained
    return result


IMAGE_SUFFIXES = {"image/jpeg": ".jpg", "image/jpg": ".jpg", "image/png": ".png", "image/webp": ".webp"}


# Side-by-side count with the trained shelf model -- never allowed to fail
# the job: the Ollama count above stays the official one, this is only sent
# along for comparison (shown as "Trained model" on the Shelf count page).
def run_trained_model(config: dict, payload: dict):
    python_path = config.get("trained_model_python") or ""
    infer_path = config.get("trained_model_infer") or ""
    if not python_path or not infer_path:
        return None
    if not Path(python_path).exists() or not Path(infer_path).exists():
        return {"error": f"trained model not found (python: {python_path}, infer: {infer_path})"}
    documents = payload.get("vision_documents") or []
    with tempfile.TemporaryDirectory(prefix="shelf-trained-") as temp_dir:
        image_paths = []
        for index, document in enumerate(documents, start=1):
            mime_type = str(document.get("mime_type") or "").strip().lower()
            content = str(document.get("content_base64") or "").strip()
            if not content or not mime_type.startswith("image/"):
                continue
            image_path = Path(temp_dir) / f"photo-{index:02d}{IMAGE_SUFFIXES.get(mime_type, '.jpg')}"
            image_path.write_bytes(base64.b64decode(content))
            image_paths.append(str(image_path))
        if not image_paths:
            return {"error": "no photos to count"}
        try:
            completed = subprocess.run(
                [python_path, infer_path, "--batch", *image_paths],
                capture_output=True, text=True, timeout=600,
            )
        except Exception as error:  # noqa: BLE001
            return {"error": f"could not run trained model: {error}"}
        if completed.returncode != 0:
            message = (completed.stderr or completed.stdout or "").strip().splitlines()
            return {"error": message[-1] if message else f"exit code {completed.returncode}"}
        try:
            return json.loads(completed.stdout.strip().splitlines()[-1])
        except Exception as error:  # noqa: BLE001
            return {"error": f"unreadable trained-model output: {error}"}


def heartbeat_payload(config: dict, status: str):
    return {
        "agent_name": config["agent_name"],
        "pc_name": config["pc_name"],
        "model_name": config["model_name"],
        "version": config["version"],
        "status": status,
        "capabilities": JOB_TYPES,
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
    print(f'Shelf-count poller starting for agent {config["agent_name"]} -> {config["server_url"]}')
    while True:
        try:
            api_request(config, "POST", "/api/llm/agent/heartbeat", heartbeat_payload(config, "online"))
            poll_response = api_request(
                config,
                "POST",
                "/api/llm/agent/poll",
                {**heartbeat_payload(config, "online"), "job_types": JOB_TYPES},
            )
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
                        "capabilities": JOB_TYPES,
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
                        "capabilities": JOB_TYPES,
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
