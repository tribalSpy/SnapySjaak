from __future__ import annotations

import json
import os
from pathlib import Path

from google.auth.exceptions import MalformedError
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import Resource, build

from src.models import ImageFile

DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
FOLDER_MIME_TYPE = "application/vnd.google-apps.folder"
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tif", ".tiff"}
QR_FILE_NAMES = {"qr.txt", "qr.json"}
SERVICE_ACCOUNT_JSON_ENV = "GOOGLE_SERVICE_ACCOUNT_JSON"


def _get_streamlit_secrets() -> dict:
    try:
        import streamlit as st
    except Exception:
        return {}

    try:
        return dict(st.secrets)
    except Exception:
        return {}


def _get_setting(name: str) -> str | None:
    value = os.getenv(name)
    if value:
        return value

    secrets = _get_streamlit_secrets()
    secret_value = secrets.get(name)
    if secret_value is None:
        return None
    return str(secret_value)


def get_setting(name: str) -> str | None:
    return _get_setting(name)


def _load_service_account_info() -> dict:
    credentials_json = _get_setting(SERVICE_ACCOUNT_JSON_ENV)
    if credentials_json:
        try:
            credentials_data = json.loads(credentials_json)
        except json.JSONDecodeError as exc:
            raise RuntimeError(
                f"{SERVICE_ACCOUNT_JSON_ENV} is not valid JSON"
            ) from exc
        return credentials_data

    credentials_path = _get_setting("GOOGLE_APPLICATION_CREDENTIALS")
    if not credentials_path:
        raise RuntimeError(
            "Set GOOGLE_APPLICATION_CREDENTIALS or GOOGLE_SERVICE_ACCOUNT_JSON"
        )
    if not Path(credentials_path).exists():
        raise RuntimeError(
            f"GOOGLE_APPLICATION_CREDENTIALS points to a missing file: {credentials_path}"
        )

    try:
        return json.loads(Path(credentials_path).read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise RuntimeError(
            f"GOOGLE_APPLICATION_CREDENTIALS is not valid JSON: {credentials_path}"
        ) from exc


def _build_credentials() -> Credentials:
    credentials_json = _get_setting(SERVICE_ACCOUNT_JSON_ENV)
    if credentials_json:
        try:
            credentials_info = json.loads(credentials_json)
            return Credentials.from_service_account_info(
                credentials_info,
                scopes=DRIVE_SCOPES,
            )
        except (MalformedError, ValueError, json.JSONDecodeError) as exc:
            raise RuntimeError(
                "GOOGLE_SERVICE_ACCOUNT_JSON must contain a valid Google service "
                "account JSON key."
            ) from exc

    credentials_path = _get_setting("GOOGLE_APPLICATION_CREDENTIALS")
    if not credentials_path:
        raise RuntimeError(
            "Set GOOGLE_APPLICATION_CREDENTIALS or GOOGLE_SERVICE_ACCOUNT_JSON"
        )

    try:
        return Credentials.from_service_account_file(
            credentials_path,
            scopes=DRIVE_SCOPES,
        )
    except (MalformedError, ValueError, OSError) as exc:
        raise RuntimeError(
            "GOOGLE_APPLICATION_CREDENTIALS must point to a valid Google service "
            "account JSON key file. The current file appears to be an OAuth client "
            "config or is otherwise malformed."
        ) from exc


class DriveService:
    def __init__(self, service: Resource) -> None:
        self._service = service

    @classmethod
    def from_service_account_env(cls) -> "DriveService":
        credentials = _build_credentials()
        service = build("drive", "v3", credentials=credentials, cache_discovery=False)
        return cls(service)

    def list_child_folders(self, folder_id: str) -> list[dict]:
        query = (
            f"'{folder_id}' in parents and trashed = false and "
            f"mimeType = '{FOLDER_MIME_TYPE}'"
        )
        return self._list_files(
            query=query,
            fields="files(id, name, mimeType, webViewLink)",
        )

    def list_files(self, folder_id: str) -> list[dict]:
        query = f"'{folder_id}' in parents and trashed = false"
        return self._list_files(
            query=query,
            fields="files(id, name, mimeType, webViewLink, size)",
        )

    def get_file(self, file_id: str) -> dict:
        return (
            self._service.files()
            .get(
                fileId=file_id,
                supportsAllDrives=True,
                fields="id, name, mimeType, parents, webViewLink",
            )
            .execute()
        )

    def list_image_files(self, folder_id: str) -> list[ImageFile]:
        image_files: list[ImageFile] = []
        for item in self.list_files(folder_id):
            if self.is_image_file(item):
                image_files.append(
                    ImageFile(
                        id=item["id"],
                        name=item["name"],
                        mime_type=item.get("mimeType", ""),
                        web_view_link=item.get("webViewLink"),
                        size=int(item["size"]) if item.get("size") else None,
                    )
                )
        return image_files

    def download_file_bytes(self, file_id: str) -> bytes:
        return (
            self._service.files()
            .get_media(fileId=file_id)
            .execute()
        )

    def extract_qr_info(self, folder_id: str) -> tuple[str, str | None]:
        files = self.list_files(folder_id)

        for item in files:
            name = item["name"].lower()
            if name in QR_FILE_NAMES:
                content = self.download_file_bytes(item["id"]).decode(
                    "utf-8", errors="replace"
                )
                if name.endswith(".json"):
                    try:
                        parsed = json.loads(content)
                        return json.dumps(parsed, indent=2, ensure_ascii=True), item["name"]
                    except json.JSONDecodeError:
                        return content.strip() or "No QR info found", item["name"]
                return content.strip() or "No QR info found", item["name"]

        for item in files:
            if "qr" in item["name"].lower():
                return item["name"], "filename"

        return "No QR info found", None

    @staticmethod
    def is_image_file(file_info: dict) -> bool:
        mime_type = file_info.get("mimeType", "")
        name = file_info.get("name", "")
        if mime_type.startswith("image/"):
            return True
        return Path(name).suffix.lower() in IMAGE_EXTENSIONS

    def _list_files(self, query: str, fields: str) -> list[dict]:
        items: list[dict] = []
        page_token: str | None = None

        while True:
            response = (
                self._service.files()
                .list(
                    q=query,
                    spaces="drive",
                    corpora="allDrives",
                    includeItemsFromAllDrives=True,
                    supportsAllDrives=True,
                    fields=f"nextPageToken, {fields}",
                    pageToken=page_token,
                    pageSize=1000,
                    orderBy="name",
                )
                .execute()
            )
            items.extend(response.get("files", []))
            page_token = response.get("nextPageToken")
            if not page_token:
                break

        return items


def validate_drive_configuration() -> str:
    root_folder_id = _get_setting("GOOGLE_DRIVE_ROOT_FOLDER_ID")

    if not root_folder_id:
        raise RuntimeError("GOOGLE_DRIVE_ROOT_FOLDER_ID is not set")

    credentials_data = _load_service_account_info()

    if credentials_data.get("type") != "service_account":
        raise RuntimeError(
            "Google Drive credentials must be a service account key JSON file. "
            "The current value is not a service account credential."
        )

    return root_folder_id
