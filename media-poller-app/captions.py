"""French + English captions from the local Ollama model (free, on the
BEAST). Without Ollama a plain caption is made from the article names, so a
render never fails on its caption."""
from __future__ import annotations

import json
import urllib.request


def _conditions(article: dict) -> str:
    c = article.get("conditions") or {}
    bits = []
    if c.get("length_cm"):
        bits.append(f"length {c['length_cm']} cm")
    if c.get("pot_cm"):
        bits.append(f"pot {c['pot_cm']} cm")
    if c.get("height_cm"):
        bits.append(f"height {c['height_cm']} cm")
    if c.get("origin"):
        bits.append(f"origin {c['origin']}")
    if c.get("colour"):
        bits.append(f"colour {c['colour']}")
    return ", ".join(bits)


def fallback(articles: list[dict], theme: dict | None, brand: str, title: str) -> dict:
    names_fr = ", ".join(a.get("name_fr") or a.get("name") or "" for a in articles if a) or title
    names_en = ", ".join(a.get("name_en") or a.get("name") or "" for a in articles if a) or title
    theme_name = (theme or {}).get("name") or ""
    return {
        "fr": f"{names_fr} 🌿 {theme_name} — {brand}".strip(" —"),
        "en": f"{names_en} 🌿 {theme_name} — {brand}".strip(" —"),
        "hashtags": (theme or {}).get("hashtags") or "#flowers #fleurs",
    }


def make_captions(config: dict, articles: list[dict], theme: dict | None, brand: str, title: str, recipe: str) -> dict:
    model = config.get("caption_model") or config.get("model_name") or ""
    url = (config.get("ollama_url") or "http://127.0.0.1:11434").rstrip("/")
    if not model:
        return fallback(articles, theme, brand, title)
    article_lines = "\n".join(
        f"- {a.get('name') or ''} (FR: {a.get('name_fr') or '?'}, EN: {a.get('name_en') or '?'}; {_conditions(a)})" for a in articles if a
    ) or f"- {title}"
    prompt = (
        f"You write short social media captions for {brand}, a Dutch flower and plant wholesaler selling to florists and shops "
        f"(mostly in France).\nTheme: {(theme or {}).get('name') or 'general'}. Brief: {(theme or {}).get('caption_prompt') or ''}\n"
        f"Content type: {recipe.replace('_', ' ')}.\nProducts:\n{article_lines}\n\n"
        "Write one caption in French and one in English (each at most 3 short sentences, natural, no prices unless given, "
        "at most 2 emojis), plus 5-8 relevant hashtags mixing French and English. "
        'Answer only JSON: {"fr": "...", "en": "...", "hashtags": "#... #..."}'
    )
    body = json.dumps({"model": model, "stream": False, "format": "json", "messages": [{"role": "user", "content": prompt}],
                       "options": {"temperature": 0.7}}).encode()
    try:
        request = urllib.request.Request(f"{url}/api/chat", data=body, headers={"content-type": "application/json"})
        with urllib.request.urlopen(request, timeout=int(config.get("caption_timeout_seconds") or 180)) as response:
            reply = json.loads(response.read().decode())
        content = json.loads(reply.get("message", {}).get("content") or "{}")
        result = {key: str(content.get(key) or "").strip() for key in ("fr", "en", "hashtags")}
        if result["fr"] and result["en"]:
            if not result["hashtags"]:
                result["hashtags"] = (theme or {}).get("hashtags") or ""
            return result
    except Exception as error:  # noqa: BLE001
        print(f"  captions: Ollama not used ({str(error)[:120]})")
    return fallback(articles, theme, brand, title)


def unload_ollama(config: dict) -> None:
    """Free the GPU memory Ollama holds (before a big AI job)."""
    model = config.get("caption_model") or config.get("model_name") or ""
    url = (config.get("ollama_url") or "http://127.0.0.1:11434").rstrip("/")
    if not model:
        return
    try:
        body = json.dumps({"model": model, "keep_alive": 0}).encode()
        urllib.request.urlopen(urllib.request.Request(f"{url}/api/generate", data=body, headers={"content-type": "application/json"}), timeout=30).read()
    except Exception:
        pass
