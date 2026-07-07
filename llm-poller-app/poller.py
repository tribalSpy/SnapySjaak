import json
import os
import socket
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path


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
    request_body = {
        "model": model_name,
        "messages": payload.get("messages", []),
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


def run_job(config: dict, job: dict):
    job_type = str(job.get("job_type") or "").strip()
    payload = job.get("payload_json") or {}
    if job_type == "ollama_chat":
        ollama_response = ollama_chat(config, payload)
        return {
            "job_type": job_type,
            "model": payload.get("model") or config["model_name"],
            "ollama_response": ollama_response,
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
