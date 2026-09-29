"""Shared JWT-exchange helper for import_tasks.py and export_annotations.py.

Recent Label Studio versions issue JWT tokens: the "Personal Access Token"
shown in the UI (Account & Settings) is a long-lived *refresh* token, not a
static API key -- it has to be exchanged here for a short-lived access
token first (the standard SimpleJWT /api/token/refresh flow) before it can
be used in an Authorization header.
"""
from __future__ import annotations

import requests


def get_access_token(config: dict) -> str:
    try:
        response = requests.post(
            f'{config["label_studio_url"].rstrip("/")}/api/token/refresh',
            json={"refresh": config["api_token"]},
            timeout=30,
        )
        response.raise_for_status()
        return response.json()["access"]
    except requests.exceptions.HTTPError as error:
        raise SystemExit(
            f"Could not exchange api_token for an access token ({error}). "
            "If this Label Studio instance is old enough to use static legacy "
            "tokens instead of JWT, /api/token/refresh 404s -- see "
            "labeling/README.md's troubleshooting note."
        ) from error


def api_headers(access_token: str) -> dict:
    return {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
