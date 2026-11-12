"""Google Drive for the media poller: list the Social library, download
photos / thumbnails, upload what was made (resumable) and share it by link.

Auth (config.json "drive"):
  - "service_account_file": a service-account JSON. A service account has no
    storage of its own, so the OUTPUT folder must be in a Shared Drive (add
    the service account as Content manager) -- or use OAuth below.
  - or "oauth": {"client_id", "client_secret", "refresh_token"} of a Google
    user who owns / can write the folders.
"""
from __future__ import annotations

import io
import json
import time
import urllib.request
from pathlib import Path

from google.oauth2 import service_account
from google.oauth2.credentials import Credentials as UserCredentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

SCOPES = ["https://www.googleapis.com/auth/drive"]
FOLDER_MIME = "application/vnd.google-apps.folder"


class Drive:
    def __init__(self, config: dict):
        drive_config = config.get("drive") or {}
        oauth = drive_config.get("oauth") or {}
        if oauth.get("refresh_token"):
            self.credentials = UserCredentials(
                token=None,
                refresh_token=oauth["refresh_token"],
                client_id=oauth.get("client_id"),
                client_secret=oauth.get("client_secret"),
                token_uri="https://oauth2.googleapis.com/token",
                scopes=SCOPES,
            )
        else:
            path = drive_config.get("service_account_file") or ""
            if not path or not Path(path).exists():
                raise RuntimeError("config.json drive.service_account_file (or drive.oauth) is not set up")
            self.credentials = service_account.Credentials.from_service_account_file(path, scopes=SCOPES)
        self.service = build("drive", "v3", credentials=self.credentials, cache_discovery=False)
        self._folder_cache: dict[tuple[str, str], str] = {}

    # -- listing ------------------------------------------------------------
    def _list(self, query: str, fields: str) -> list[dict]:
        items, token = [], None
        while True:
            response = self.service.files().list(
                q=query, fields=f"nextPageToken, files({fields})", pageSize=1000, pageToken=token,
                supportsAllDrives=True, includeItemsFromAllDrives=True,
            ).execute()
            items.extend(response.get("files", []))
            token = response.get("nextPageToken")
            if not token:
                return items

    def walk_images(self, folder_id: str, path: tuple[str, ...] = ()):
        """Every image under the folder (recursively), with its folder path."""
        fields = "id, name, mimeType, size, md5Checksum, modifiedTime, createdTime, imageMediaMetadata(width, height, time, rotation), thumbnailLink"
        for item in self._list(f"'{folder_id}' in parents and trashed = false", fields):
            if item.get("mimeType") == FOLDER_MIME:
                # Folders starting with "_" (e.g. _thumbs) are ours, skipped.
                if not item["name"].startswith("_"):
                    yield from self.walk_images(item["id"], path + (item["name"],))
            elif str(item.get("mimeType", "")).startswith("image/"):
                yield {**item, "folder_path": "/".join(path)}

    # -- download -----------------------------------------------------------
    def download(self, file_id: str, target: Path) -> Path:
        target.parent.mkdir(parents=True, exist_ok=True)
        request = self.service.files().get_media(fileId=file_id, supportsAllDrives=True)
        with io.FileIO(str(target), "wb") as handle:
            downloader = MediaIoBaseDownload(handle, request, chunksize=8 * 1024 * 1024)
            done = False
            while not done:
                _, done = downloader.next_chunk()
        return target

    def download_thumbnail(self, thumbnail_link: str, size: int = 480) -> bytes | None:
        """Drive's own thumbnail (small, no full photo download needed)."""
        if not thumbnail_link:
            return None
        if not self.credentials.valid:
            self.credentials.refresh(Request())
        link = thumbnail_link.split("=s")[0] + f"=s{size}"
        request = urllib.request.Request(link, headers={"Authorization": f"Bearer {self.credentials.token}"})
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                return response.read()
        except Exception:
            return None

    # -- upload -------------------------------------------------------------
    def ensure_folder(self, parent_id: str, name: str) -> str:
        key = (parent_id, name)
        if key in self._folder_cache:
            return self._folder_cache[key]
        safe = name.replace("'", "\\'")
        found = self._list(f"'{parent_id}' in parents and name = '{safe}' and mimeType = '{FOLDER_MIME}' and trashed = false", "id")
        if found:
            folder_id = found[0]["id"]
        else:
            folder_id = self.service.files().create(
                body={"name": name, "mimeType": FOLDER_MIME, "parents": [parent_id]}, fields="id", supportsAllDrives=True,
            ).execute()["id"]
        self._folder_cache[key] = folder_id
        return folder_id

    def ensure_path(self, root_id: str, parts) -> str:
        folder_id = root_id
        for part in parts or []:
            if str(part).strip():
                folder_id = self.ensure_folder(folder_id, str(part).strip()[:100])
        return folder_id

    def upload(self, path: Path, folder_id: str, mime: str, name: str | None = None, share: bool = True) -> dict:
        media = MediaFileUpload(str(path), mimetype=mime, resumable=True, chunksize=16 * 1024 * 1024)
        request = self.service.files().create(
            body={"name": name or path.name, "parents": [folder_id]}, media_body=media, fields="id, name, size", supportsAllDrives=True,
        )
        response = None
        attempts = 0
        while response is None:
            try:
                _, response = request.next_chunk()
            except Exception:
                attempts += 1
                if attempts > 5:
                    raise
                time.sleep(2 * attempts)
        if share:
            self.share_by_link(response["id"])
        return response

    def share_by_link(self, file_id: str) -> bool:
        """Anyone with the link can view (for the app's previews). A Shared
        Drive's policy may forbid it -- then previews need a Google login."""
        try:
            self.service.permissions().create(
                fileId=file_id, body={"type": "anyone", "role": "reader"}, fields="id", supportsAllDrives=True,
            ).execute()
            return True
        except Exception as error:  # noqa: BLE001
            print(f"  (could not share {file_id} by link: {str(error)[:160]})")
            return False


def load_json(path: Path, fallback):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return fallback
