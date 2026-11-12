"""The recipes: photos and videos made from HD photos.

Frames are drawn with Pillow (smooth sub-pixel crops, eased motion, text
cards) and piped into ffmpeg, which encodes them -- on the BEAST with the
RTX's NVENC encoder (h264_nvenc), elsewhere with libx264.

Every recipe takes:
  photos   -- local image paths (in order)
  articles -- [{name, name_fr, name_en, conditions{...}, offer_price}] (may be empty)
  options  -- {duration, variants ["9x16","4x5","1x1"], text_fr, text_en,
              show_price, show_conditions, music, points [{x,y}]}
  context  -- {title, theme_name, brand {name, accent}, font paths, ffmpeg...}
and returns a list of output files: {path, variant, kind, mime, width,
height, duration, index}.
"""
from __future__ import annotations

import math
import shutil
import subprocess
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont, ImageOps

FPS = 30
SIZES = {"9x16": (1080, 1920), "4x5": (1080, 1350), "1x1": (1080, 1080)}

VIDEO_RECIPES = {"zoom_in", "zoom_out", "pan", "detail_tour", "slideshow", "offer_reel"}
PHOTO_RECIPES = {"photo_post", "carousel", "catalogue_card"}


# ---------------------------------------------------------------------------
# helpers

def ease(t: float) -> float:
    """Ease in-out (smooth start and stop)."""
    t = max(0.0, min(1.0, t))
    return t * t * (3 - 2 * t)


def hex_rgb(value: str, fallback=(31, 122, 90)) -> tuple[int, int, int]:
    try:
        value = value.lstrip("#")
        return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))
    except Exception:
        return fallback


def find_font(context: dict, bold: bool = False) -> str | None:
    wanted = context.get("font_bold" if bold else "font_regular")
    candidates = [wanted] if wanted else []
    candidates += (
        ["C:/Windows/Fonts/segoeuib.ttf", "C:/Windows/Fonts/arialbd.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"]
        if bold else
        ["C:/Windows/Fonts/segoeui.ttf", "C:/Windows/Fonts/arial.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"]
    )
    for path in candidates:
        if path and Path(path).exists():
            return path
    return None


def font(context: dict, size: int, bold: bool = False):
    path = find_font(context, bold)
    return ImageFont.truetype(path, size) if path else ImageFont.load_default()


def load_photo(path: Path, max_side: int = 6000) -> Image.Image:
    image = Image.open(path)
    image = ImageOps.exif_transpose(image).convert("RGB")
    if max(image.size) > max_side:
        image.thumbnail((max_side, max_side), Image.LANCZOS)
    return image


def is_packshot(image: Image.Image) -> bool:
    """A product photo on a plain light background (the border is mostly
    near-white): shown whole instead of cropped."""
    small = image.resize((64, 64))
    pixels = [small.getpixel((x, y)) for x in range(64) for y in (0, 1, 62, 63)] + [small.getpixel((x, y)) for y in range(64) for x in (0, 1, 62, 63)]
    light = sum(1 for r, g, b in pixels if min(r, g, b) > 228)
    return light / len(pixels) > 0.6


def border_colour(image: Image.Image) -> tuple[int, int, int]:
    small = image.resize((32, 32))
    pixels = sorted([small.getpixel((x, y)) for x in range(32) for y in (0, 31)] + [small.getpixel((x, y)) for y in range(32) for x in (0, 31)], key=sum)
    return pixels[len(pixels) // 2]


def stage_photo(image: Image.Image, size, fit: str = "auto"):
    """The picture the motion runs over. "cover" = the photo itself (cropped
    to the output shape); "contain" = the whole photo on a canvas of the
    output shape in its own background colour (packshots). "auto" picks
    contain for packshots whose shape differs from the output's.
    Returns (staged image, map) -- map turns photo-relative points (0..1)
    into staged-relative points."""
    target = size[0] / size[1]
    shape = image.width / image.height
    if fit == "auto":
        fit = "contain" if is_packshot(image) and abs(shape - target) / target > 0.08 else "cover"
    if fit != "contain":
        return image, (lambda x, y: (x, y))
    margin = 1.08
    if shape > target:
        canvas_w = int(image.width * margin)
        canvas_h = int(canvas_w / target)
    else:
        canvas_h = int(image.height * margin)
        canvas_w = int(canvas_h * target)
    canvas = Image.new("RGB", (canvas_w, canvas_h), border_colour(image))
    ox = (canvas_w - image.width) // 2
    # Tall outputs: the product a bit above the middle, the text goes below.
    oy = int((canvas_h - image.height) * (0.36 if target < 0.7 else 0.12 if target < 0.95 else 0.5))
    canvas.paste(image, (ox, oy))
    staged_map = lambda x, y: ((ox + x * image.width) / canvas_w, (oy + y * image.height) / canvas_h)  # noqa: E731
    staged_map.contain = True
    return canvas, staged_map


def detail_points(image: Image.Image, count: int = 3) -> list[dict]:
    """The busiest parts of the photo (most edges) -- leaves, flowers,
    texture -- spread out, for the detail tour."""
    grid = 8
    small = image.convert("L").resize((grid * 24, grid * 24)).filter(ImageFilter.FIND_EDGES)
    cells = []
    for gx in range(1, grid - 1):
        for gy in range(1, grid - 1):
            box = small.crop((gx * 24, gy * 24, gx * 24 + 24, gy * 24 + 24))
            cells.append((sum(box.getdata()), (gx + 0.5) / grid, (gy + 0.5) / grid))
    cells.sort(reverse=True)
    chosen = []
    for _, x, y in cells:
        if all(abs(x - p["x"]) + abs(y - p["y"]) > 0.3 for p in chosen):
            chosen.append({"x": x, "y": y})
        if len(chosen) == count:
            break
    return chosen or [{"x": 0.5, "y": 0.5}]


def base_crop(image: Image.Image, size: tuple[int, int], focus=(0.5, 0.5)):
    """The largest crop of the output's shape, centred on the focus."""
    width, height = image.size
    target = size[0] / size[1]
    if width / height > target:
        crop_h = height
        crop_w = height * target
    else:
        crop_w = width
        crop_h = width / target
    cx = min(max(focus[0] * width, crop_w / 2), width - crop_w / 2)
    cy = min(max(focus[1] * height, crop_h / 2), height - crop_h / 2)
    return cx, cy, crop_w, crop_h


def prepared(image: Image.Image, size: tuple[int, int], min_scale: float) -> Image.Image:
    """Scale the photo down so the tightest crop still has full output detail
    (keeps every frame fast without losing sharpness)."""
    _, _, crop_w, _ = base_crop(image, size)
    needed = size[0] / max(min_scale, 0.05)
    factor = min(1.0, (needed * 1.05) / crop_w)
    if factor < 0.999:
        image = image.resize((max(1, int(image.width * factor)), max(1, int(image.height * factor))), Image.LANCZOS)
    return image


def frame_at(image: Image.Image, size, cx, cy, scale) -> Image.Image:
    """The output frame: a crop of (scale x the largest crop) around cx, cy."""
    _, _, crop_w, crop_h = base_crop(image, size)
    w, h = crop_w * scale, crop_h * scale
    cx = min(max(cx, w / 2), image.width - w / 2)
    cy = min(max(cy, h / 2), image.height - h / 2)
    box = (cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2)
    return image.resize(size, Image.BICUBIC, box=box)


def wrap(draw: ImageDraw.ImageDraw, text: str, face, width: int) -> list[str]:
    words, lines, line = str(text or "").split(), [], ""
    for word in words:
        trial = f"{line} {word}".strip()
        if draw.textlength(trial, font=face) <= width or not line:
            line = trial
        else:
            lines.append(line)
            line = word
    if line:
        lines.append(line)
    return lines


def conditions_line(article: dict, show_price: bool = False) -> str:
    c = article.get("conditions") or {}
    parts = []
    if c.get("length_cm"):
        parts.append(f"{c['length_cm']:g} cm")
    if c.get("pot_cm"):
        parts.append(f"Ø {c['pot_cm']:g} cm")
    if c.get("height_cm"):
        parts.append(f"H {c['height_cm']:g} cm")
    if c.get("packages") and c.get("per_package"):
        pieces = c.get("pieces") or c["packages"] * c["per_package"]
        parts.append(f"{c['packages']:g} × {c['per_package']:g} = {pieces:g}")
    elif c.get("per_package"):
        parts.append(f"{c['per_package']:g} / box")
    if c.get("origin"):
        parts.append(str(c["origin"]))
    if show_price and article.get("offer_price"):
        parts.append(f"€ {float(article['offer_price']):.2f}")
    return "  ·  ".join(parts)


def article_titles(article: dict | None, options: dict, context: dict) -> tuple[str, str]:
    """(French line, English line) for a text card."""
    fr = options.get("text_fr") or (article or {}).get("name_fr") or ""
    en = options.get("text_en") or (article or {}).get("name_en") or ""
    if not fr and not en:
        fr = (article or {}).get("name") or context.get("title") or context.get("theme_name") or ""
    return fr, en


# ---------------------------------------------------------------------------
# overlays (RGBA, drawn once per segment, blended per frame)

def text_overlay(size, context: dict, title_fr: str, title_en: str, sub: str = "", position: str = "bottom") -> Image.Image:
    width, height = size
    overlay = Image.new("RGBA", size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    accent = hex_rgb((context.get("brand") or {}).get("accent", "#1f7a5a"))
    big = font(context, int(width * 0.07), bold=True)
    mid = font(context, int(width * 0.045))
    small = font(context, int(width * 0.036))
    lines = []
    for text, face, colour in ((title_fr, big, (255, 255, 255, 255)), (title_en, mid, (235, 240, 236, 255)), (sub, small, (235, 240, 236, 235))):
        for line in wrap(draw, text, face, int(width * 0.84)):
            lines.append((line, face, colour))
    if not lines:
        return overlay
    line_heights = [face.getbbox("Ag")[3] + int(width * 0.018) for _, face, _ in lines]
    block = sum(line_heights)
    top = {"bottom": int(height * 0.68) - block // 2, "band": int(height * 0.865) - block // 2}.get(position, int(height * 0.16))
    # Soft dark gradient + a rounded panel behind the text so it reads on
    # any photo, white packshots too.
    gradient = Image.new("L", (1, 256))
    for y in range(256):
        gradient.putpixel((0, y), int(150 * (y / 255) ** 1.6))
    band_top = max(0, top - int(height * 0.12))
    band = gradient.resize((width, height - band_top)) if position in ("bottom", "band") else gradient.transpose(Image.FLIP_TOP_BOTTOM).resize((width, top + block + int(height * 0.1)))
    shade = Image.new("RGBA", band.size, (0, 0, 0, 255))
    shade.putalpha(band)
    overlay.alpha_composite(shade, (0, band_top if position in ("bottom", "band") else 0))
    draw = ImageDraw.Draw(overlay)
    x = int(width * 0.08)
    text_w = max(draw.textlength(line, font=face) for line, face, _ in lines)
    pad = int(width * 0.035)
    panel = Image.new("RGBA", size, (0, 0, 0, 0))
    ImageDraw.Draw(panel).rounded_rectangle([x - pad, top - int(width * 0.05), x + text_w + pad, top + block + pad // 2], radius=int(width * 0.025), fill=(12, 20, 16, 150))
    overlay.alpha_composite(panel)
    draw = ImageDraw.Draw(overlay)
    draw.rectangle([x, top - int(width * 0.03), x + int(width * 0.12), top - int(width * 0.022)], fill=accent + (255,))
    y = top
    for (line, face, colour), line_height in zip(lines, line_heights):
        draw.text((x, y), line, font=face, fill=colour)
        y += line_height
    return overlay


def brand_overlay(size, context: dict) -> Image.Image:
    width, height = size
    overlay = Image.new("RGBA", size, (0, 0, 0, 0))
    name = (context.get("brand") or {}).get("name") or ""
    if not name:
        return overlay
    draw = ImageDraw.Draw(overlay)
    face = font(context, int(width * 0.032), bold=True)
    pad = int(width * 0.05)
    text_w = draw.textlength(name, font=face)
    box_h = face.getbbox("Ag")[3] + int(width * 0.02)
    draw.rounded_rectangle([pad - 14, pad - 8, pad + text_w + 14, pad + box_h], radius=12, fill=(0, 0, 0, 90))
    draw.text((pad, pad), name, font=face, fill=(255, 255, 255, 235))
    return overlay


def blend(frame: Image.Image, overlay: Image.Image | None, alpha: float) -> Image.Image:
    if overlay is None or alpha <= 0:
        return frame
    if alpha >= 1:
        base = frame.convert("RGBA")
        base.alpha_composite(overlay)
        return base.convert("RGB")
    faded = overlay.copy()
    faded.putalpha(faded.getchannel("A").point(lambda value: int(value * alpha)))
    base = frame.convert("RGBA")
    base.alpha_composite(faded)
    return base.convert("RGB")


# ---------------------------------------------------------------------------
# ffmpeg

class Encoder:
    def __init__(self, context: dict):
        self.ffmpeg = context.get("ffmpeg") or find_ffmpeg()
        self.codec = context.get("video_codec") or detect_codec(self.ffmpeg)

    def open(self, path: Path, size, music: Path | None, duration: float, fps: int = FPS):
        width, height = size
        command = [self.ffmpeg, "-y", "-hide_banner", "-loglevel", "error",
                   "-f", "rawvideo", "-pix_fmt", "rgb24", "-s", f"{width}x{height}", "-r", str(fps), "-i", "-"]
        if music and music.exists():
            command += ["-i", str(music)]
            audio = ["-filter:a", f"afade=t=in:d=0.8,afade=t=out:st={max(0.0, duration - 1.2):.2f}:d=1.2", "-c:a", "aac", "-b:a", "160k"]
        else:
            # A silent track: some platforms handle a video without audio badly.
            command += ["-f", "lavfi", "-i", "anullsrc=r=44100:cl=stereo"]
            audio = ["-c:a", "aac", "-b:a", "96k"]
        command += ["-map", "0:v", "-map", "1:a", "-t", f"{duration:.2f}"]
        if self.codec == "h264_nvenc":
            command += ["-c:v", "h264_nvenc", "-preset", "p5", "-rc", "vbr", "-cq", "20", "-b:v", "0"]
        else:
            command += ["-c:v", "libx264", "-preset", "medium", "-crf", "19"]
        command += ["-pix_fmt", "yuv420p", *audio, "-shortest", "-movflags", "+faststart", str(path)]
        return subprocess.Popen(command, stdin=subprocess.PIPE, stderr=subprocess.PIPE)


def find_ffmpeg() -> str:
    found = shutil.which("ffmpeg")
    if found:
        return found
    try:
        import imageio_ffmpeg  # optional, for test PCs without ffmpeg
        return imageio_ffmpeg.get_ffmpeg_exe()
    except Exception as error:  # noqa: BLE001
        raise RuntimeError("ffmpeg not found: install it (see README) or set config ffmpeg_path") from error


_codec_cache: dict[str, str] = {}


def detect_codec(ffmpeg: str) -> str:
    """h264_nvenc when the GPU can really encode (a 1-second test), else libx264."""
    if ffmpeg in _codec_cache:
        return _codec_cache[ffmpeg]
    codec = "libx264"
    try:
        test = subprocess.run(
            [ffmpeg, "-hide_banner", "-loglevel", "error", "-f", "lavfi", "-i", "color=c=black:s=256x256:d=1", "-c:v", "h264_nvenc", "-f", "null", "-"],
            capture_output=True, timeout=30,
        )
        if test.returncode == 0:
            codec = "h264_nvenc"
    except Exception:
        pass
    _codec_cache[ffmpeg] = codec
    return codec


def write_video(path: Path, size, frames, total_frames: int, context: dict, music: Path | None, fps: int = FPS) -> None:
    encoder = Encoder(context)
    process = encoder.open(path, size, music, total_frames / fps, fps)
    try:
        for frame in frames:
            process.stdin.write(frame.tobytes())
        process.stdin.close()
    except BrokenPipeError:
        pass
    error = process.stderr.read().decode("utf-8", "replace")
    if process.wait() != 0:
        raise RuntimeError(f"ffmpeg failed: {error[-600:]}")


# ---------------------------------------------------------------------------
# motion plans: for frame t (0..1) -> (cx, cy, scale)

def plan_zoom(image, size, focus, start_scale, end_scale):
    cx, cy, _, _ = base_crop(image, size, focus)
    return lambda t: (cx, cy, start_scale + (end_scale - start_scale) * ease(t))


def plan_pan(image, size):
    width, height = image.size
    scale = 0.86
    _, _, crop_w, crop_h = base_crop(image, size)
    w, h = crop_w * scale, crop_h * scale
    if width - w > height - h:  # pan sideways
        y = height / 2
        return lambda t: (w / 2 + (width - w) * ease(t), y, scale)
    x = width / 2
    return lambda t: (x, h / 2 + (height - h) * ease(t), scale)


def plan_detail_tour(image, size, points):
    """Whole photo -> each detail point (close) -> whole photo again."""
    width, height = image.size
    points = points or [{"x": 0.5, "y": 0.42}, {"x": 0.32, "y": 0.62}, {"x": 0.68, "y": 0.35}]
    cx0, cy0, _, _ = base_crop(image, size)
    stops = [(cx0, cy0, 1.0)] + [(p["x"] * width, p["y"] * height, 0.42) for p in points] + [(cx0, cy0, 1.0)]
    legs = len(stops) - 1

    def at(t):
        position = min(legs - 1e-6, t * legs)
        leg = int(position)
        local = position - leg
        # Hold a little at each stop: move in the middle 70% of a leg.
        local = ease(min(1.0, max(0.0, (local - 0.15) / 0.7)))
        a, b = stops[leg], stops[leg + 1]
        return a[0] + (b[0] - a[0]) * local, a[1] + (b[1] - a[1]) * local, a[2] + (b[2] - a[2]) * local

    return at


# ---------------------------------------------------------------------------
# video recipes

def single_photo_video(recipe, photo: Path, size, duration, article, options, context):
    image, to_stage = stage_photo(load_photo(photo), size, options.get("fit") or "auto")
    points = [dict(zip(("x", "y"), to_stage(p["x"], p["y"]))) for p in (options.get("points") or [])]
    if recipe == "detail_tour" and not points:
        points = detail_points(image)
    focus = (points[0]["x"], points[0]["y"]) if points and recipe in ("zoom_in", "zoom_out") else (0.5, 0.5)
    # A whole packshot on its canvas is small: zoom further in.
    close = 0.62 if getattr(to_stage, "contain", False) else 0.8
    if getattr(to_stage, "contain", False) and recipe in ("zoom_in", "zoom_out") and not (options.get("points") or []):
        focus = to_stage(0.5, 0.5)
    min_scale = 0.42 if recipe == "detail_tour" else close - 0.02
    image = prepared(image, size, min_scale)
    if recipe == "zoom_in":
        plan = plan_zoom(image, size, focus, 1.0, close)
    elif recipe == "zoom_out":
        plan = plan_zoom(image, size, focus, close, 1.0)
    elif recipe == "pan":
        plan = plan_pan(image, size)
    else:
        plan = plan_detail_tour(image, size, points)
    title_fr, title_en = article_titles(article, options, context)
    sub = conditions_line(article, options.get("show_price")) if article and options.get("show_conditions", True) else ""
    title = text_overlay(size, context, title_fr, title_en)
    end_card = text_overlay(size, context, title_fr, title_en, sub) if sub else None
    brand = brand_overlay(size, context)
    total = int(duration * FPS)

    def frames():
        for index in range(total):
            t = index / max(1, total - 1)
            seconds = index / FPS
            frame = frame_at(image, size, *plan(t))
            frame = blend(frame, brand, 1.0)
            # Title fades in at 0.4 s, out before the end card.
            if end_card is not None and seconds > duration - 2.6:
                frame = blend(frame, end_card, min(1.0, (seconds - (duration - 2.6)) / 0.5))
            else:
                fade_in = min(1.0, max(0.0, (seconds - 0.4) / 0.6))
                fade_out = min(1.0, max(0.0, (duration - 0.3 - seconds) / 0.5)) if end_card is None else min(1.0, max(0.0, (duration - 2.7 - seconds) / 0.4))
                frame = blend(frame, title, min(fade_in, fade_out))
            yield frame

    return frames(), total


def multi_photo_video(recipe, photos, size, duration, articles, options, context):
    """Slideshow / offer reel: a slow zoom per photo, crossfades between."""
    count = len(photos)
    per = max(2.0, duration / count)
    fade = min(0.6, per / 4)
    total = int(per * count * FPS)
    segments = []
    for index, photo in enumerate(photos):
        image = prepared(stage_photo(load_photo(photo), size, options.get("fit") or "auto")[0], size, 0.82)
        zoom_in = index % 2 == 0
        plan = plan_zoom(image, size, (0.5, 0.5), 1.0 if zoom_in else 0.84, 0.84 if zoom_in else 1.0)
        article = articles[index] if index < len(articles) else (articles[0] if recipe == "slideshow" and len(articles) == 1 else None)
        if recipe == "offer_reel" or (article and index == 0) or (index == 0 and (options.get("text_fr") or options.get("text_en"))):
            title_fr, title_en = article_titles(article, options if index == 0 else {}, context)
            sub = conditions_line(article, options.get("show_price")) if article and options.get("show_conditions", True) else ""
            overlay = text_overlay(size, context, title_fr, title_en, sub)
        else:
            overlay = None
        segments.append((image, plan, overlay))
    brand = brand_overlay(size, context)
    seg_frames = int(per * FPS)

    def render(segment, local_index):
        image, plan, overlay = segment
        t = local_index / max(1, seg_frames - 1)
        frame = frame_at(image, size, *plan(t))
        seconds = local_index / FPS
        if overlay is not None:
            frame = blend(frame, overlay, min(1.0, max(0.0, (seconds - 0.3) / 0.5), max(0.0, (per - 0.2 - seconds) / 0.4)))
        return frame

    def frames():
        fade_frames = int(fade * FPS)
        for index in range(total):
            segment_index = min(count - 1, index // seg_frames)
            local = index - segment_index * seg_frames
            frame = render(segments[segment_index], local)
            # Crossfade into the next photo over the last fade frames.
            if segment_index < count - 1 and local >= seg_frames - fade_frames:
                nxt = render(segments[segment_index + 1], 0)
                frame = Image.blend(frame, nxt, (local - (seg_frames - fade_frames)) / max(1, fade_frames))
            yield blend(frame, brand, 1.0)

    return frames(), total


# ---------------------------------------------------------------------------
# photo recipes

def photo_post(photo: Path, size, article, options, context) -> Image.Image:
    original = load_photo(photo)
    title_fr, title_en = article_titles(article, options, context)
    sub = conditions_line(article, options.get("show_price")) if article and options.get("show_conditions", True) else ""
    fit = options.get("fit") or "auto"
    if fit == "contain" or (fit == "auto" and is_packshot(original)):
        # A packshot: the whole product in the upper part, the text in its
        # own band below (never over the product).
        width, height = size
        area = (width, int(height * 0.74))
        staged, _ = stage_photo(original, area, "contain")
        frame = Image.new("RGB", size, border_colour(original))
        frame.paste(staged.resize(area, Image.LANCZOS), (0, 0))
        frame = blend(frame, text_overlay(size, context, title_fr, title_en, sub, position="band"), 1.0)
        return blend(frame, brand_overlay(size, context), 1.0)
    image, to_stage = stage_photo(original, size, fit)
    points = [dict(zip(("x", "y"), to_stage(p["x"], p["y"]))) for p in (options.get("points") or [])]
    focus = (points[0]["x"], points[0]["y"]) if points else (0.5, 0.5)
    cx, cy, w, h = base_crop(image, size, focus)
    frame = image.resize(size, Image.LANCZOS, box=(cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2))
    frame = blend(frame, text_overlay(size, context, title_fr, title_en, sub), 1.0)
    return blend(frame, brand_overlay(size, context), 1.0)


def catalogue_card(photo: Path, size, article, options, context) -> Image.Image:
    """White card: photo on top, name FR / EN and the conditions below."""
    width, height = size
    accent = hex_rgb((context.get("brand") or {}).get("accent", "#1f7a5a"))
    card = Image.new("RGB", size, (255, 255, 255))
    photo_h = int(height * 0.64)
    image, _ = stage_photo(load_photo(photo), (width, photo_h), options.get("fit") or "auto")
    cx, cy, w, h = base_crop(image, (width, photo_h))
    card.paste(image.resize((width, photo_h), Image.LANCZOS, box=(cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2)), (0, 0))
    draw = ImageDraw.Draw(card)
    draw.rectangle([0, photo_h, width, photo_h + 10], fill=accent)
    x, y = int(width * 0.07), photo_h + int(height * 0.04)
    title_fr, title_en = article_titles(article, options, context)
    for text, face, colour in ((title_fr, font(context, int(width * 0.06), True), (23, 32, 42)), (title_en, font(context, int(width * 0.042)), (80, 92, 104))):
        for line in wrap(draw, text, face, int(width * 0.86)):
            draw.text((x, y), line, font=face, fill=colour)
            y += face.getbbox("Ag")[3] + 8
    y += 14
    c = (article or {}).get("conditions") or {}
    rows = []
    if c.get("length_cm"):
        rows.append(("Longueur / Length", f"{c['length_cm']:g} cm"))
    if c.get("pot_cm"):
        rows.append(("Pot", f"{c['pot_cm']:g} cm"))
    if c.get("height_cm"):
        rows.append(("Hauteur / Height", f"{c['height_cm']:g} cm"))
    if c.get("packages") and c.get("per_package"):
        rows.append(("Conditionnement / Packing", f"{c['packages']:g} × {c['per_package']:g} = {(c.get('pieces') or c['packages'] * c['per_package']):g}"))
    elif c.get("per_package"):
        rows.append(("Par carton / Per box", f"{c['per_package']:g}"))
    for key, label in (("quality", "Qualité / Quality"), ("colour", "Couleur / Colour"), ("origin", "Origine / Origin")):
        if c.get(key):
            rows.append((label, str(c[key])))
    if options.get("show_price") and (article or {}).get("offer_price"):
        rows.append(("Prix / Price", f"€ {float(article['offer_price']):.2f}"))
    label_face, value_face = font(context, int(width * 0.03)), font(context, int(width * 0.034), True)
    for label, value in rows[:6]:
        draw.text((x, y), label, font=label_face, fill=(91, 104, 117))
        draw.text((int(width * 0.55), y), value, font=value_face, fill=(23, 32, 42))
        y += int(width * 0.05)
    brand = (context.get("brand") or {}).get("name") or ""
    if brand:
        face = font(context, int(width * 0.03), True)
        draw.text((width - x - draw.textlength(brand, font=face), height - int(height * 0.05)), brand, font=face, fill=accent)
    return card


# ---------------------------------------------------------------------------
# entry point

def make(recipe: str, photos: list[Path], articles: list[dict], options: dict, context: dict, out_dir: Path) -> list[dict]:
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = context.get("stem") or recipe
    outputs = []
    if recipe in VIDEO_RECIPES:
        variants = [v for v in (options.get("variants") or ["9x16"]) if v in SIZES] or ["9x16"]
        duration = float(options.get("duration") or 8)
        music = Path(context["music_dir"]) / options["music"] if options.get("music") and context.get("music_dir") else None
        for variant in variants:
            size = SIZES[variant]
            if recipe in ("slideshow", "offer_reel"):
                frames, total = multi_photo_video(recipe, photos, size, duration, articles, options, context)
            else:
                frames, total = single_photo_video(recipe, photos[0], size, duration, articles[0] if articles else None, options, context)
            path = out_dir / f"{stem}_{variant}.mp4"
            write_video(path, size, frames, total, context, music)
            outputs.append({"path": path, "variant": variant, "kind": "video", "mime": "video/mp4", "width": size[0], "height": size[1], "duration": round(total / FPS, 2)})
            # A cover picture (the frame at 1 s) for the platforms.
            cover_frames, _ = (multi_photo_video(recipe, photos, size, duration, articles, options, context) if recipe in ("slideshow", "offer_reel")
                               else single_photo_video(recipe, photos[0], size, duration, articles[0] if articles else None, options, context))
            cover = None
            for index, frame in enumerate(cover_frames):
                if index >= FPS * 1.2:
                    cover = frame
                    break
            if cover is not None:
                cover_path = out_dir / f"{stem}_{variant}_cover.jpg"
                cover.save(cover_path, quality=92)
                outputs.append({"path": cover_path, "variant": variant, "kind": "cover", "mime": "image/jpeg", "width": size[0], "height": size[1]})
    elif recipe in PHOTO_RECIPES:
        variants = [v for v in (options.get("variants") or ["4x5"]) if v in SIZES] or ["4x5"]
        sources = photos if recipe == "carousel" else photos[:1]
        for variant in variants:
            size = SIZES[variant]
            for index, photo in enumerate(sources):
                article = articles[index] if index < len(articles) else (articles[0] if len(articles) == 1 else None)
                image = catalogue_card(photo, size, article, options, context) if recipe == "catalogue_card" else photo_post(photo, size, article, options if index == 0 else {**options, "text_fr": "", "text_en": ""}, context)
                path = out_dir / f"{stem}_{variant}_{index + 1:02d}.jpg"
                image.save(path, quality=93)
                outputs.append({"path": path, "variant": variant, "kind": "photo", "mime": "image/jpeg", "width": size[0], "height": size[1], "index": index + 1})
    else:
        raise RuntimeError(f"Unknown recipe: {recipe}")
    return outputs


def write_video_fps(path: Path, size, frames, total_frames: int, context: dict, music: Path | None, fps: int) -> None:
    write_video(path, size, frames, total_frames, context, music, fps)
