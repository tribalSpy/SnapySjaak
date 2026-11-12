"""Media poller for the BEAST PC: makes the Campagnes photos and videos.

It claims media_* jobs from SnappySjaak (Render), does all the heavy work
here -- reading photos from Google Drive, Pillow + ffmpeg (NVENC on the
RTX), FR/EN captions with the local Ollama, uploads to Drive -- and reports
back only ids and texts.

  python poller.py                     run the poller (run_poller.bat)
  python poller.py --test-recipe zoom_in --photo a.jpg [--photo b.jpg] \
         [--title "Raphis"] [--variants 9x16,4x5] [--out out_dir]
                                       make one recipe locally: no server,
                                       no Drive (for trying looks)
"""
from __future__ import annotations

import argparse
import json
import os
import platform
import shutil
import sys
import tempfile
import time
import traceback
import urllib.error
import urllib.request
import uuid
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(APP_DIR))
# Captions have accents and emojis; a Windows console must not crash on them.
for stream in (sys.stdout, sys.stderr):
    try:
        stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:  # noqa: BLE001
        pass

import recipes  # noqa: E402
from captions import make_captions, unload_ollama  # noqa: E402

VERSION = "1.1.0"
BASE_JOB_TYPES = ["media_scan", "media_render"]


def ai_enabled(config: dict) -> bool:
    return bool((config.get("ai") or {}).get("enabled"))


def in_ai_window(config: dict, now=None) -> bool:
    """AI jobs run in the night window (GPU free from Ollama / training)
    unless sent as "now". Window "19:00-07:00"; empty = always."""
    window = str((config.get("ai") or {}).get("window") or "").strip()
    if not window or "-" not in window:
        return True
    now = now or time.localtime()
    minutes = now.tm_hour * 60 + now.tm_min
    try:
        start, end = [int(part.split(":")[0]) * 60 + int(part.split(":")[1]) for part in window.split("-")]
    except Exception:
        return True
    return start <= minutes < end if start < end else (minutes >= start or minutes < end)


def job_types(config: dict) -> list[str]:
    types = list(BASE_JOB_TYPES)
    if ai_enabled(config):
        types.append("media_ai_now")
        if in_ai_window(config):
            types.append("media_ai")
    return types


def load_json(path: Path):
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def load_config() -> dict:
    path = APP_DIR / "config.json"
    if not path.exists():
        path = APP_DIR / "config.example.json"
    config = load_json(path)
    config.setdefault("poll_interval_seconds", 10)
    if not config.get("work_dir"):
        config["work_dir"] = str(APP_DIR / "work")
    config.setdefault("capabilities", ["media_render", "media_scan"])
    config["server_url"] = os.environ.get("SHADOW_POLLER_SERVER_URL", config.get("server_url", "")).rstrip("/")
    config["api_key"] = os.environ.get("SHADOW_POLLER_API_KEY", config.get("api_key", ""))
    return config


def api_request(config: dict, method: str, path: str, payload=None, timeout: int = 120):
    data = json.dumps(payload).encode() if payload is not None else None
    request = urllib.request.Request(
        f"{config['server_url']}{path}", data=data, method=method,
        headers={"content-type": "application/json", "x-shadow-agent-key": config["api_key"]},
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            return json.loads(response.read().decode() or "{}")
    except urllib.error.HTTPError as error:
        raise RuntimeError(f"{method} {path} -> {error.code}: {error.read().decode(errors='replace')[:300]}") from error


def render_context(config: dict, payload: dict, stem: str) -> dict:
    return {
        "title": payload.get("title") or "",
        "theme_name": (payload.get("theme") or {}).get("name") or "",
        "brand": payload.get("brand") or {"name": config.get("brand_name", ""), "accent": "#1f7a5a"},
        "font_regular": config.get("font_regular"),
        "font_bold": config.get("font_bold"),
        "ffmpeg": config.get("ffmpeg_path") or None,
        "music_dir": config.get("music_dir"),
        "stem": stem,
    }


# ---------------------------------------------------------------------------
# jobs

def run_render(config: dict, payload: dict, use_ai: bool = False) -> dict:
    from drive import Drive

    started = time.time()
    drive = Drive(config)
    work = Path(config["work_dir"])
    cache = work / "photos"
    cache.mkdir(parents=True, exist_ok=True)
    job_dir = Path(tempfile.mkdtemp(prefix=f"item{payload.get('item_id')}_", dir=work))
    try:
        photos = []
        for asset in payload.get("assets") or []:
            target = cache / f"{asset['drive_id']}{Path(asset.get('name') or '').suffix or '.jpg'}"
            if not target.exists():
                print(f"  downloading {asset.get('name')}")
                drive.download(asset["drive_id"], target)
            photos.append(target)
        stem = f"{payload.get('recipe')}_{payload.get('item_id')}_{time.strftime('%Y%m%d_%H%M%S')}"
        context = render_context(config, payload, stem)
        if use_ai:
            import ai
            unload_ollama(config)  # the GPU for the AI model
            outputs = ai.make(payload["recipe"], config, photos, payload.get("articles") or [], payload.get("options") or {}, context, job_dir)
        else:
            outputs = recipes.make(payload["recipe"], photos, payload.get("articles") or [], payload.get("options") or {}, context, job_dir)
        output = payload.get("output") or {}
        folder_id = drive.ensure_path(output["folder_id"], output.get("subfolder") or [])
        uploaded = []
        for item in outputs:
            print(f"  uploading {item['path'].name}")
            meta = drive.upload(item["path"], folder_id, item["mime"], share=output.get("share", True))
            uploaded.append({**{k: v for k, v in item.items() if k != "path"}, "drive_id": meta["id"], "name": meta["name"],
                             "size": int(meta.get("size") or item["path"].stat().st_size)})
        captions = payload.get("captions") or {}
        texts = {"fr": "", "en": "", "hashtags": ""}
        if not captions.get("keep"):
            texts = make_captions(config, payload.get("articles") or [], payload.get("theme"), (payload.get("brand") or {}).get("name", ""),
                                  payload.get("title") or "", payload.get("recipe") or "")
        return {"outputs": uploaded, "caption_fr": texts["fr"], "caption_en": texts["en"], "hashtags": texts["hashtags"],
                "render_seconds": round(time.time() - started, 1), "codec": recipes.detect_codec(recipes.find_ffmpeg() if not config.get("ffmpeg_path") else config["ffmpeg_path"])}
    finally:
        shutil.rmtree(job_dir, ignore_errors=True)


def run_scan(config: dict, payload: dict) -> dict:
    """List the library folder, make small thumbnails (shared by link, in
    <output>/_thumbs) and report the photos to the server in chunks."""
    from drive import Drive, load_json as read_cache
    import io
    from PIL import Image, ImageOps

    drive = Drive(config)
    work = Path(config["work_dir"])
    work.mkdir(parents=True, exist_ok=True)
    cache_path = work / "thumb_cache.json"
    cache = read_cache(cache_path, {})
    thumbs_folder = drive.ensure_folder(payload["output_folder_id"], "_thumbs")
    scan_id = uuid.uuid4().hex
    batch, total = [], 0

    def send(complete: bool):
        nonlocal batch
        api_request(config, "POST", "/api/agent/media/assets", {"scan_id": scan_id, "assets": batch, "complete": complete}, timeout=300)
        batch = []

    for item in drive.walk_images(payload["library_folder_id"]):
        key = item["id"]
        fingerprint = item.get("md5Checksum") or item.get("modifiedTime") or ""
        cached = cache.get(key) or {}
        thumb_id = cached.get("thumb_id", "") if cached.get("fingerprint") == fingerprint else ""
        if not thumb_id:
            data = drive.download_thumbnail(item.get("thumbnailLink") or "")
            if data is None:
                full = work / "photos" / f"{key}{Path(item['name']).suffix or '.jpg'}"
                if not full.exists():
                    drive.download(key, full)
                data = full.read_bytes()
            try:
                image = ImageOps.exif_transpose(Image.open(io.BytesIO(data))).convert("RGB")
                image.thumbnail((480, 480))
                thumb_file = work / "thumb.jpg"
                image.save(thumb_file, quality=85)
                thumb_id = drive.upload(thumb_file, thumbs_folder, "image/jpeg", name=f"{key}.jpg", share=payload.get("share_outputs", True))["id"]
                cache[key] = {"fingerprint": fingerprint, "thumb_id": thumb_id}
                cache_path.write_text(json.dumps(cache), encoding="utf-8")
            except Exception as error:  # noqa: BLE001
                print(f"  no thumbnail for {item['name']}: {error}")
        meta = item.get("imageMediaMetadata") or {}
        width, height = meta.get("width"), meta.get("height")
        if meta.get("rotation") in (1, 3) and width and height:
            width, height = height, width
        taken = meta.get("time")
        if taken and len(taken) >= 19:
            taken = taken[:10].replace(":", "-") + "T" + taken[11:19]
        batch.append({
            "drive_id": key, "name": item["name"], "folder_path": item.get("folder_path", ""), "mime": item.get("mimeType", ""),
            "width": width, "height": height, "size": int(item.get("size") or 0) or None, "taken_at": taken or item.get("createdTime"),
            "thumb_id": thumb_id, "tags": [part for part in item.get("folder_path", "").split("/") if part],
        })
        total += 1
        if len(batch) >= 200:
            send(False)
    send(True)
    return {"photos": total, "scan_id": scan_id}


def run_job(config: dict, job: dict) -> dict:
    job_type = job.get("job_type")
    payload = job.get("payload_json") or {}
    if job_type == "media_render":
        return run_render(config, payload)
    if job_type in ("media_ai", "media_ai_now"):
        if not ai_enabled(config):
            raise RuntimeError("AI is not switched on on this PC (config ai.enabled)")
        return run_render(config, payload, use_ai=True)
    if job_type == "media_scan":
        return run_scan(config, payload)
    raise RuntimeError(f"Unsupported job type: {job_type}")


def heartbeat_payload(config: dict, status: str) -> dict:
    return {
        "agent_name": config["agent_name"],
        "pc_name": config.get("pc_name") or platform.node(),
        "model_name": config.get("caption_model") or "",
        "version": VERSION,
        "status": status,
        "capabilities": sorted(set(config["capabilities"]) | ({"media_ai"} if ai_enabled(config) else set())),
        "job_types": job_types(config),
        "meta": {"python": platform.python_version(), "platform": platform.platform(), "app": "media-poller-app"},
    }


def main_loop(config: dict) -> None:
    for key in ("server_url", "api_key", "agent_name"):
        if not config.get(key):
            raise SystemExit(f"config.json: {key} is missing")
    Path(config["work_dir"]).mkdir(parents=True, exist_ok=True)
    print(f"Media poller {VERSION} for {config['server_url']} as {config['agent_name']} (encoder: {recipes.detect_codec(config.get('ffmpeg_path') or recipes.find_ffmpeg())}, "
          f"AI: {'on, window ' + str((config.get('ai') or {}).get('window') or 'always') if ai_enabled(config) else 'off'})")
    while True:
        try:
            response = api_request(config, "POST", "/api/llm/agent/poll", heartbeat_payload(config, "online"))
            job = response.get("job")
            if not job:
                time.sleep(max(2, int(config["poll_interval_seconds"])))
                continue
            job_id = job["id"]
            print(f"Claimed job {job_id} ({job.get('job_type')})")
            try:
                result = run_job(config, job)
                api_request(config, "POST", f"/api/llm/jobs/{job_id}/result", {**heartbeat_payload(config, "idle"), "result_json": result})
                print(f"Done job {job_id}")
            except Exception as error:  # noqa: BLE001
                traceback.print_exc()
                api_request(config, "POST", f"/api/llm/jobs/{job_id}/fail", {**heartbeat_payload(config, "idle"), "error_text": str(error)[:2000], "allow_retry": False})
        except Exception as error:  # noqa: BLE001
            print(f"Poll error: {error}")
            time.sleep(15)


def test_recipe(config: dict, args) -> None:
    photos = [Path(photo) for photo in args.photo]
    for photo in photos:
        if not photo.exists():
            raise SystemExit(f"Photo not found: {photo}")
    out = Path(args.out or (APP_DIR / "test_output"))
    options = {"duration": args.duration, "variants": args.variants.split(","), "text_fr": args.text_fr, "text_en": args.text_en, "show_conditions": True}
    articles = []
    if args.title:
        articles = [{"name": args.title, "name_fr": args.text_fr or args.title, "name_en": args.text_en or args.title,
                     "conditions": {"length_cm": 55, "packages": 10, "per_package": 50, "origin": "Nederland / Pays-Bas"}}]
    context = render_context(config, {"title": args.title, "brand": {"name": config.get("brand_name", "Sjaak van der Vijver"), "accent": "#1f7a5a"}}, args.test_recipe)
    started = time.time()
    if args.ai_prompt:
        options["ai_prompt"] = args.ai_prompt
    if args.test_recipe.startswith("ai_"):
        import ai
        unload_ollama(config)
        outputs = ai.make(args.test_recipe, config, photos, articles, options, context, out)
    else:
        outputs = recipes.make(args.test_recipe, photos, articles, options, context, out)
    print(f"Made in {time.time() - started:.1f}s with {recipes.detect_codec(context['ffmpeg'] or recipes.find_ffmpeg())}:")
    for item in outputs:
        print(f"  {item['path']}  {item['width']}x{item['height']} {item.get('duration', '')}")
    if args.captions:
        print(json.dumps(make_captions(config, articles, {"name": "Test", "caption_prompt": "", "hashtags": "#flowers"}, context["brand"]["name"], args.title or "", args.test_recipe), ensure_ascii=False, indent=1))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--test-recipe")
    parser.add_argument("--photo", action="append", default=[])
    parser.add_argument("--title", default="")
    parser.add_argument("--text-fr", default="")
    parser.add_argument("--text-en", default="")
    parser.add_argument("--variants", default="9x16")
    parser.add_argument("--duration", type=float, default=8)
    parser.add_argument("--out")
    parser.add_argument("--captions", action="store_true")
    parser.add_argument("--ai-prompt", default="")
    cli = parser.parse_args()
    cfg = load_config()
    if cli.test_recipe:
        test_recipe(cfg, cli)
    else:
        main_loop(cfg)
