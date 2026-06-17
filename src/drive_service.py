from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
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
DRIVE_ACCOUNT_NAMES_ENV = "GOOGLE_DRIVE_ACCOUNT_NAMES"
DEFAULT_DRIVE_ACCOUNT = "default"


@dataclass(frozen=True)
class DriveAccountConfig:
    account_name: str
    root_folder_id: str


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


def _normalize_account_name(account_name: str) -> str:
    normalized = re.sub(r"[^A-Za-z0-9]+", "_", (account_name or "").strip()).strip("_")
    return normalized.lower() or DEFAULT_DRIVE_ACCOUNT


def _account_suffix(account_name: str) -> str:
    normalized = _normalize_account_name(account_name)
    if normalized == DEFAULT_DRIVE_ACCOUNT:
        return ""
    return f"_{normalized.upper()}"


def _get_account_setting(
    name: str,
    account_name: str = DEFAULT_DRIVE_ACCOUNT,
    *,
    fallback_to_default: bool = False,
) -> str | None:
    suffix = _account_suffix(account_name)
    if suffix:
        value = _get_setting(f"{name}{suffix}")
        if value:
            return value
    if (
        _normalize_account_name(account_name) == DEFAULT_DRIVE_ACCOUNT
        or fallback_to_default
    ):
        return _get_setting(name)
    return None


def _configured_drive_account_names() -> list[str]:
    raw_names = _get_setting(DRIVE_ACCOUNT_NAMES_ENV) or ""
    names = [DEFAULT_DRIVE_ACCOUNT]
    seen = {DEFAULT_DRIVE_ACCOUNT}

    for raw_name in raw_names.split(","):
        normalized = _normalize_account_name(raw_name)
        if normalized in seen:
            continue
        seen.add(normalized)
        names.append(normalized)

    return names


def _load_service_account_info(account_name: str = DEFAULT_DRIVE_ACCOUNT) -> dict:
    credentials_json = _get_account_setting(
        SERVICE_ACCOUNT_JSON_ENV,
        account_name,
        fallback_to_default=True,
    )
    if credentials_json:
        try:
            credentials_data = json.loads(credentials_json)
        except json.JSONDecodeError as exc:
            raise RuntimeError(
                f"{SERVICE_ACCOUNT_JSON_ENV}{_account_suffix(account_name)} is not valid JSON"
            ) from exc
        return credentials_data

    credentials_path = _get_account_setting(
        "GOOGLE_APPLICATION_CREDENTIALS",
        account_name,
        fallback_to_default=True,
    )
    if not credentials_path:
        raise RuntimeError(
            f"Set GOOGLE_APPLICATION_CREDENTIALS{_account_suffix(account_name)} "
            f"or GOOGLE_SERVICE_ACCOUNT_JSON{_account_suffix(account_name)}"
        )
    if not Path(credentials_path).exists():
        raise RuntimeError(
            f"GOOGLE_APPLICATION_CREDENTIALS{_account_suffix(account_name)} points to "
            f"a missing file: {credentials_path}"
        )

    try:
        return json.loads(Path(credentials_path).read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise RuntimeError(
            f"GOOGLE_APPLICATION_CREDENTIALS{_account_suffix(account_name)} is not valid "
            f"JSON: {credentials_path}"
        ) from exc


def _build_credentials(account_name: str = DEFAULT_DRIVE_ACCOUNT) -> Credentials:
    credentials_json = _get_account_setting(
        SERVICE_ACCOUNT_JSON_ENV,
        account_name,
        fallback_to_default=True,
    )
    if credentials_json:
        try:
            credentials_info = json.loads(credentials_json)
            return Credentials.from_service_account_info(
                credentials_info,
                scopes=DRIVE_SCOPES,
            )
        except (MalformedError, ValueError, json.JSONDecodeError) as exc:
            raise RuntimeError(
                f"GOOGLE_SERVICE_ACCOUNT_JSON{_account_suffix(account_name)} must contain "
                "a valid Google service account JSON key."
            ) from exc

    credentials_path = _get_account_setting(
        "GOOGLE_APPLICATION_CREDENTIALS",
        account_name,
        fallback_to_default=True,
    )
    if not credentials_path:
        raise RuntimeError(
            f"Set GOOGLE_APPLICATION_CREDENTIALS{_account_suffix(account_name)} "
            f"or GOOGLE_SERVICE_ACCOUNT_JSON{_account_suffix(account_name)}"
        )

    try:
        return Credentials.from_service_account_file(
            credentials_path,
            scopes=DRIVE_SCOPES,
        )
    except (MalformedError, ValueError, OSError) as exc:
        raise RuntimeError(
            f"GOOGLE_APPLICATION_CREDENTIALS{_account_suffix(account_name)} must point "
            "to a valid Google service account JSON key file. The current file appears "
            "to be an OAuth client config or is otherwise malformed."
        ) from exc


class DriveService:
    def __init__(self, service: Resource) -> None:
        self._service = service

    @classmethod
    def from_service_account_env(
        cls,
        account_name: str = DEFAULT_DRIVE_ACCOUNT,
    ) -> "DriveService":
        credentials = _build_credentials(account_name)
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
        return self.list_image_files_from_items(self.list_files(folder_id))

    def download_file_bytes(self, file_id: str) -> bytes:
        return (
            self._service.files()
            .get_media(fileId=file_id)
            .execute()
        )

    def extract_qr_info(self, folder_id: str) -> tuple[str, str | None]:
        return self.extract_qr_info_from_items(self.list_files(folder_id))

    def list_run_folder_details(self, folder_id: str) -> tuple[list[ImageFile], tuple[str, str | None]]:
        files = self.list_files(folder_id)
        return self.list_image_files_from_items(files), self.extract_qr_info_from_items(files)

    def list_image_files_from_items(self, items: list[dict]) -> list[ImageFile]:
        image_files: list[ImageFile] = []
        for item in items:
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

    def extract_qr_info_from_items(self, files: list[dict]) -> tuple[str, str | None]:

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


def validate_drive_configuration(account_name: str = DEFAULT_DRIVE_ACCOUNT) -> str:
    root_folder_id = _get_account_setting("GOOGLE_DRIVE_ROOT_FOLDER_ID", account_name)

    if not root_folder_id:
        raise RuntimeError(
            f"GOOGLE_DRIVE_ROOT_FOLDER_ID{_account_suffix(account_name)} is not set"
        )

    credentials_data = _load_service_account_info(account_name)

    if credentials_data.get("type") != "service_account":
        raise RuntimeError(
            "Google Drive credentials must be a service account key JSON file. "
            "The current value is not a service account credential."
        )

    return root_folder_id


def list_drive_configurations() -> list[DriveAccountConfig]:
    configurations: list[DriveAccountConfig] = []

    for account_name in _configured_drive_account_names():
        root_folder_id = _get_account_setting("GOOGLE_DRIVE_ROOT_FOLDER_ID", account_name)
        if not root_folder_id:
            continue
        validated_root_folder_id = validate_drive_configuration(account_name)
        configurations.append(
            DriveAccountConfig(
                account_name=account_name,
                root_folder_id=validated_root_folder_id,
            )
        )

    if not configurations:
        raise RuntimeError("No Google Drive accounts are configured")

    return configurations
