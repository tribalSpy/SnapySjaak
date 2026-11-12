"""The AI recipes, on the BEAST's RTX 5090 (Hugging Face diffusers).

  ai_flower_drop  Wan 2.2 image-to-video: flowers / petals falling around the
                  product, then ffmpeg slow motion (motion-interpolated 2x)
  ai_living       Wan 2.2 image-to-video: the photo gently comes alive
  ai_art          SDXL + ControlNet (canny): the photo redrawn as pencil
                  sketch, watercolour and oil painting (keeps its lines),
                  then a morph video photo -> art + the art stills

Models (all allow commercial use; downloaded once into the Hugging Face
cache, see `python ai.py --download`):
  video: Wan-AI/Wan2.2-TI2V-5B-Diffusers           (Apache-2.0)
  art:   stabilityai/stable-diffusion-xl-base-1.0   (CreativeML OpenRAIL++-M)
         diffusers/controlnet-canny-sdxl-1.0        (OpenRAIL++)
Flux-dev is deliberately not used: its licence is non-commercial.

The GPU is shared with Ollama: the poller unloads Ollama before an AI job;
every pipeline is freed again after the job.
"""
from __future__ import annotations

import gc
import subprocess
import sys
import time
from pathlib import Path

from PIL import Image

import recipes

DEFAULTS = {
    "video_model": "Wan-AI/Wan2.2-TI2V-5B-Diffusers",
    "video_width": 704,
    "video_height": 1280,
    "video_frames": 121,
    "video_fps": 24,
    "video_steps": 40,
    "video_guidance": 5.0,
    "art_model": "stabilityai/stable-diffusion-xl-base-1.0",
    "art_controlnet": "diffusers/controlnet-canny-sdxl-1.0",
    "art_steps": 30,
    "art_strength": 0.75,
    "cpu_offload": False,
}

NEGATIVE_VIDEO = ("blurry, low quality, distorted, deformed flowers, text, watermark, logo, jitter, flicker, "
                  "camera shake, extra pots, morphing objects, ugly, static")

ART_STYLES = {
    "sketch": "a detailed pencil sketch of {subject}, graphite drawing on white paper, fine hatching, botanical illustration",
    "watercolor": "a delicate watercolour painting of {subject}, soft washes, paper texture, botanical art",
    "oil": "an impressionist oil painting of {subject}, visible brush strokes, rich colours, canvas texture",
}


def settings(config: dict) -> dict:
    return {**DEFAULTS, **((config.get("ai") or {}))}


def free_gpu() -> None:
    gc.collect()
    try:
        import torch
        if torch.cuda.is_available():
            torch.cuda.empty_cache()
            torch.cuda.ipc_collect()
    except Exception:
        pass


def subject_of(articles: list[dict], title: str) -> str:
    article = articles[0] if articles else {}
    return article.get("name_en") or article.get("name") or title or "flowers"


# ---------------------------------------------------------------------------
# video: Wan 2.2 image -> video

def wan_video(config: dict, image: Image.Image, prompt: str, log=print) -> list[Image.Image]:
    import torch
    from diffusers import AutoencoderKLWan, WanImageToVideoPipeline

    ai = settings(config)
    log(f"  loading {ai['video_model']} ...")
    started = time.time()
    try:
        vae = AutoencoderKLWan.from_pretrained(ai["video_model"], subfolder="vae", torch_dtype=torch.float32)
        pipe = WanImageToVideoPipeline.from_pretrained(ai["video_model"], vae=vae, torch_dtype=torch.bfloat16)
    except Exception as error:
        raise RuntimeError(
            f"Could not load {ai['video_model']} as image-to-video ({str(error)[:300]}). Update diffusers "
            "(pip install -U diffusers), or set config ai.video_model to \"Wan-AI/Wan2.2-I2V-A14B-Diffusers\" "
            "with ai.cpu_offload true."
        ) from error
    if ai["cpu_offload"]:
        pipe.enable_model_cpu_offload()
    else:
        pipe.to("cuda")
    log(f"  loaded in {time.time() - started:.0f}s; generating {ai['video_frames']} frames ...")
    try:
        result = pipe(
            image=image.resize((ai["video_width"], ai["video_height"]), Image.LANCZOS),
            prompt=prompt,
            negative_prompt=NEGATIVE_VIDEO,
            width=ai["video_width"],
            height=ai["video_height"],
            num_frames=ai["video_frames"],
            num_inference_steps=ai["video_steps"],
            guidance_scale=ai["video_guidance"],
            generator=torch.Generator("cuda").manual_seed(int(time.time()) % 100000),
        )
        frames = result.frames[0]
        # diffusers gives PIL images or float arrays depending on version.
        if not isinstance(frames[0], Image.Image):
            import numpy as np
            frames = [Image.fromarray((np.clip(frame, 0, 1) * 255).astype("uint8")) for frame in frames]
        return frames
    finally:
        del pipe, vae
        free_gpu()


def frames_to_file(frames, path: Path, fps: int, ffmpeg: str, slow_motion: float, size) -> Path:
    """Encode the generated frames, motion-interpolated for slow motion and
    upscaled to the output size (lanczos + light sharpen)."""
    width, height = frames[0].size
    filters = []
    if slow_motion > 1:
        filters.append(f"minterpolate=fps={int(fps * slow_motion)}:mi_mode=mci:mc_mode=aobmc:me_mode=bidir:vsbmc=1")
        filters.append(f"setpts={slow_motion}*PTS")
    filters.append(f"scale={size[0]}:{size[1]}:flags=lanczos:force_original_aspect_ratio=increase,crop={size[0]}:{size[1]}")
    filters.append("unsharp=5:5:0.6")
    command = [ffmpeg, "-y", "-hide_banner", "-loglevel", "error", "-f", "rawvideo", "-pix_fmt", "rgb24", "-s", f"{width}x{height}",
               "-r", str(fps), "-i", "-", "-vf", ",".join(filters), "-r", str(fps), "-c:v", "libx264", "-crf", "12", "-preset", "fast",
               "-pix_fmt", "yuv420p", str(path)]
    process = subprocess.Popen(command, stdin=subprocess.PIPE, stderr=subprocess.PIPE)
    for frame in frames:
        process.stdin.write(frame.convert("RGB").tobytes())
    process.stdin.close()
    error = process.stderr.read().decode("utf-8", "replace")
    if process.wait() != 0:
        raise RuntimeError(f"ffmpeg (slow motion / upscale) failed: {error[-500:]}")
    return path


def read_frames(path: Path, size, ffmpeg: str):
    """Decode a video back into PIL frames (for the text overlays)."""
    width, height = size
    process = subprocess.Popen([ffmpeg, "-hide_banner", "-loglevel", "quiet", "-i", str(path), "-f", "rawvideo", "-pix_fmt", "rgb24", "-"],
                               stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
    frame_bytes = width * height * 3
    try:
        while True:
            chunk = process.stdout.read(frame_bytes)
            if len(chunk) < frame_bytes:
                break
            yield Image.frombytes("RGB", size, chunk)
    finally:
        # Also when the caller stops early (the cover frame).
        process.stdout.close()
        if process.poll() is None:
            process.kill()
        process.wait()


def count_frames(path: Path, ffmpeg: str) -> int:
    probe = subprocess.run([ffmpeg, "-hide_banner", "-i", str(path), "-map", "0:v:0", "-c", "copy", "-f", "null", "-"], capture_output=True, text=True)
    for line in reversed(probe.stderr.splitlines()):
        if "frame=" in line:
            try:
                return int(line.split("frame=")[1].split()[0])
            except Exception:
                break
    return 0


def video_recipe(recipe: str, config: dict, photos, articles, options, context, out_dir: Path, log=print) -> list[dict]:
    ai = settings(config)
    ffmpeg = context.get("ffmpeg") or recipes.find_ffmpeg()
    subject = subject_of(articles, context.get("title") or "")
    extra = (options.get("ai_prompt") or "").strip()
    if recipe == "ai_flower_drop":
        prompt = (f"{subject}, fresh flower petals and small blossoms slowly falling and drifting down around it, "
                  "gentle slow motion, soft natural light, clean studio background, shallow depth of field, cinematic, highly detailed")
        slow = 2.0
    else:
        prompt = (f"{subject}, leaves and petals moving very gently in a soft breeze, subtle natural motion, light slowly shifting, "
                  "camera almost still with a slow push in, photorealistic, highly detailed")
        slow = 1.0
    if extra:
        prompt = f"{prompt}, {extra}"
    # Generate vertical (9:16); the other formats are crops of it.
    gen_size = (ai["video_width"], ai["video_height"])
    staged, _ = recipes.stage_photo(recipes.load_photo(photos[0]), gen_size, options.get("fit") or "auto")
    frames = wan_video(config, staged, prompt, log)
    work = out_dir / "_ai"
    work.mkdir(parents=True, exist_ok=True)
    outputs = []
    variants = [v for v in (options.get("variants") or ["9x16"]) if v in recipes.SIZES] or ["9x16"]
    article = articles[0] if articles else None
    title_fr, title_en = recipes.article_titles(article, options, context)
    sub = recipes.conditions_line(article, options.get("show_price")) if article and options.get("show_conditions", True) else ""
    music = Path(context["music_dir"]) / options["music"] if options.get("music") and context.get("music_dir") else None
    for variant in variants:
        size = recipes.SIZES[variant]
        base = frames_to_file(frames, work / f"base_{variant}.mp4", ai["video_fps"], ffmpeg, slow, size)
        total = count_frames(base, ffmpeg)
        duration = total / ai["video_fps"]
        title = recipes.text_overlay(size, context, title_fr, title_en, sub)
        brand = recipes.brand_overlay(size, context)

        def composed(base=base, size=size, title=title, brand=brand, duration=duration):
            for index, frame in enumerate(read_frames(base, size, ffmpeg)):
                seconds = index / ai["video_fps"]
                frame = recipes.blend(frame, brand, 1.0)
                # The text comes in after the first second and stays.
                yield recipes.blend(frame, title, min(1.0, max(0.0, (seconds - 1.0) / 0.8)))

        stem = context.get("stem") or recipe
        path = out_dir / f"{stem}_{variant}.mp4"
        context_fps = {**context}
        recipes.write_video_fps(path, size, composed(), total, context_fps, music, ai["video_fps"])
        outputs.append({"path": path, "variant": variant, "kind": "video", "mime": "video/mp4", "width": size[0], "height": size[1], "duration": round(duration, 2)})
        cover = None
        for index, frame in enumerate(composed()):
            if index >= int(ai["video_fps"] * 2.2):
                cover = frame
                break
        if cover is not None:
            cover_path = out_dir / f"{stem}_{variant}_cover.jpg"
            cover.save(cover_path, quality=92)
            outputs.append({"path": cover_path, "variant": variant, "kind": "cover", "mime": "image/jpeg", "width": size[0], "height": size[1]})
    return outputs


# ---------------------------------------------------------------------------
# art: SDXL + ControlNet canny

def canny(image: Image.Image) -> Image.Image:
    import cv2
    import numpy as np
    array = np.array(image.convert("RGB"))
    edges = cv2.Canny(array, 80, 180)
    return Image.fromarray(np.stack([edges] * 3, axis=-1))


def art_images(config: dict, image: Image.Image, subject: str, styles, extra: str, log=print) -> dict:
    import torch
    from diffusers import AutoencoderKL, ControlNetModel, StableDiffusionXLControlNetImg2ImgPipeline

    ai = settings(config)
    log(f"  loading {ai['art_model']} + {ai['art_controlnet']} ...")
    controlnet = ControlNetModel.from_pretrained(ai["art_controlnet"], torch_dtype=torch.float16)
    vae = AutoencoderKL.from_pretrained("madebyollin/sdxl-vae-fp16-fix", torch_dtype=torch.float16)
    pipe = StableDiffusionXLControlNetImg2ImgPipeline.from_pretrained(ai["art_model"], controlnet=controlnet, vae=vae, torch_dtype=torch.float16, variant="fp16")
    if ai["cpu_offload"]:
        pipe.enable_model_cpu_offload()
    else:
        pipe.to("cuda")
    try:
        control = canny(image)
        results = {}
        for style in styles:
            prompt = ART_STYLES[style].format(subject=subject) + (f", {extra}" if extra else "")
            log(f"  drawing: {style}")
            result = pipe(
                prompt=prompt,
                negative_prompt="photo, photorealistic, blurry, text, watermark, logo, deformed",
                image=image,
                control_image=control,
                strength=ai["art_strength"],
                controlnet_conditioning_scale=0.65,
                num_inference_steps=ai["art_steps"],
                guidance_scale=7.0,
                generator=torch.Generator("cuda").manual_seed(1234),
            )
            results[style] = result.images[0]
        return results
    finally:
        del pipe, controlnet, vae
        free_gpu()


def art_recipe(config: dict, photos, articles, options, context, out_dir: Path, log=print) -> list[dict]:
    subject = subject_of(articles, context.get("title") or "")
    styles = [style for style in (options.get("art_styles") or ["sketch", "watercolor", "oil"]) if style in ART_STYLES] or ["sketch", "watercolor"]
    # SDXL works best near 1 megapixel: 832 x 1216 (about 9:16).
    gen_size = (832, 1216)
    staged, _ = recipes.stage_photo(recipes.load_photo(photos[0]), gen_size, options.get("fit") or "auto")
    staged = staged.resize(gen_size, Image.LANCZOS)
    arts = art_images(config, staged, subject, styles, (options.get("ai_prompt") or "").strip(), log)
    work = out_dir / "_ai"
    work.mkdir(parents=True, exist_ok=True)
    sequence = [work / "00_photo.png"]
    staged.save(sequence[0])
    for index, style in enumerate(styles, start=1):
        path = work / f"{index:02d}_{style}.png"
        arts[style].save(path)
        sequence.append(path)
    stem = context.get("stem") or "ai_art"
    # The morph video: photo -> each style, crossfades with a slow zoom.
    video_options = {**options, "fit": "cover", "duration": max(8.0, float(options.get("duration") or 10))}
    outputs = recipes.make("slideshow", sequence, articles, video_options, {**context, "stem": stem}, out_dir)
    # And the art stills themselves (4:5 posts).
    for index, style in enumerate(styles, start=1):
        still = recipes.photo_post(sequence[index], recipes.SIZES["4x5"], articles[0] if articles else None,
                                   {**options, "fit": "cover", "text_fr": options.get("text_fr", ""), "text_en": options.get("text_en", "")}, context)
        path = out_dir / f"{stem}_4x5_{style}.jpg"
        still.save(path, quality=93)
        outputs.append({"path": path, "variant": "4x5", "kind": "photo", "mime": "image/jpeg", "width": 1080, "height": 1350, "index": index})
    return outputs


def make(recipe: str, config: dict, photos, articles, options, context, out_dir: Path, log=print) -> list[dict]:
    out_dir.mkdir(parents=True, exist_ok=True)
    if recipe in ("ai_flower_drop", "ai_living"):
        return video_recipe(recipe, config, photos, articles, options, context, out_dir, log)
    if recipe == "ai_art":
        return art_recipe(config, photos, articles, options, context, out_dir, log)
    raise RuntimeError(f"Unknown AI recipe: {recipe}")


# ---------------------------------------------------------------------------

def check(config: dict) -> None:
    """Is the AI part installed and does the GPU work?"""
    import torch
    import diffusers
    print(f"torch {torch.__version__}, CUDA {torch.version.cuda}, diffusers {diffusers.__version__}")
    if not torch.cuda.is_available():
        raise SystemExit("CUDA is not available -- install the CUDA 12.8 build of torch (setup_media_pc.bat does).")
    props = torch.cuda.get_device_properties(0)
    print(f"GPU: {props.name}, {props.total_memory / 1024**3:.0f} GB, compute {props.major}.{props.minor}")
    x = torch.randn(2048, 2048, device="cuda", dtype=torch.bfloat16)
    print("matmul OK:", float((x @ x).float().abs().mean()) > 0)


def download(config: dict) -> None:
    """Fetch the models into the Hugging Face cache once (tens of GB)."""
    from huggingface_hub import snapshot_download
    ai = settings(config)
    for repo in (ai["video_model"], ai["art_model"], ai["art_controlnet"], "madebyollin/sdxl-vae-fp16-fix"):
        print(f"downloading {repo} ...")
        allow = None
        if repo == ai["art_model"]:
            # Only the diffusers weights (not the extra single-file checkpoints).
            allow = ["*.json", "*.txt", "*.fp16.safetensors", "tokenizer/*", "tokenizer_2/*", "scheduler/*"]
        snapshot_download(repo, allow_patterns=allow)
    print("models ready.")


if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import poller  # noqa: E402
    cfg = poller.load_config()
    if "--download" in sys.argv:
        download(cfg)
    else:
        check(cfg)
