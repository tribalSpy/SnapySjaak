from __future__ import annotations

import argparse
import base64
import json
import mimetypes
import os
import smtplib
import sys
import tempfile
from pathlib import Path
from email.message import EmailMessage

from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

load_dotenv(REPO_ROOT / ".env")

from src.drive_service import DEFAULT_DRIVE_ACCOUNT, DriveService  # noqa: E402

SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
DRIVE_FILE_SCOPES = ["https://www.googleapis.com/auth/drive"]
DRIVE_FOLDER_MIME_TYPE = "application/vnd.google-apps.folder"


def serialize_image(image) -> dict[str, object]:
    return {
        "id": image.id,
        "name": image.name,
        "mime_type": image.mime_type,
        "web_view_link": image.web_view_link,
        "size": image.size,
    }


def details() -> int:
    runs = json.loads(sys.stdin.read() or "[]")
    if isinstance(runs, dict):
        runs = [runs]
    hydrated_runs = []
    drive_services: dict[str, DriveService] = {}

    for run in runs:
        metadata = run.get("metadata", {}) if isinstance(run, dict) else {}
        account_name = str(metadata.get("drive_account", DEFAULT_DRIVE_ACCOUNT))
        if account_name not in drive_services:
            drive_services[account_name] = DriveService.from_service_account_env(
                account_name
            )
        images, (qr_info, qr_source) = drive_services[account_name].list_run_folder_details(
            run["folder_id"]
        )
        hydrated_runs.append(
            {
                **run,
                "images": [serialize_image(image) for image in images],
                "qr_info": qr_info,
                "qr_source": qr_source,
            }
        )

    sys.stdout.write(json.dumps(hydrated_runs, ensure_ascii=True))
    return 0


def image(file_id: str, account_name: str = DEFAULT_DRIVE_ACCOUNT) -> int:
    drive_service = DriveService.from_service_account_env(account_name)
    sys.stdout.buffer.write(drive_service.download_file_bytes(file_id))
    return 0


def _service_account_credentials(scopes: list[str] | None = None) -> Credentials:
    active_scopes = scopes or SHEETS_SCOPES
    credentials_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if credentials_json:
        return Credentials.from_service_account_info(
            json.loads(credentials_json),
            scopes=active_scopes,
        )

    credentials_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
    if credentials_path:
        return Credentials.from_service_account_file(
            credentials_path,
            scopes=active_scopes,
        )

    raise RuntimeError("Set GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_APPLICATION_CREDENTIALS")


def service_account_info() -> int:
    credentials = _service_account_credentials()
    sys.stdout.write(
        json.dumps(
            {
                "client_email": credentials.service_account_email,
                "project_id": getattr(credentials, "project_id", "") or "",
            },
            ensure_ascii=True,
        )
    )
    return 0


def sheets_read(spreadsheet_id: str, sheet_name: str) -> int:
    sheet_range = f"{sheet_name}!A:Z"
    service = build("sheets", "v4", credentials=_service_account_credentials(), cache_discovery=False)
    response = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=sheet_range)
        .execute()
    )
    sys.stdout.write(json.dumps({"values": response.get("values", [])}, ensure_ascii=True))
    return 0


def sheets_append(spreadsheet_id: str, sheet_name: str) -> int:
    payload = json.loads(sys.stdin.read() or "{}")
    row = payload.get("row") if isinstance(payload, dict) else []
    service = build("sheets", "v4", credentials=_service_account_credentials(), cache_discovery=False)
    response = (
        service.spreadsheets()
        .values()
        .append(
            spreadsheetId=spreadsheet_id,
            range=sheet_name,
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body={"values": [row]},
        )
        .execute()
    )
    sys.stdout.write(json.dumps(response, ensure_ascii=True))
    return 0


def _column_name(index: int) -> str:
    result = ""
    current = index
    while current > 0:
        current, remainder = divmod(current - 1, 26)
        result = chr(65 + remainder) + result
    return result



def _drive_service():
    return build("drive", "v3", credentials=_service_account_credentials(DRIVE_FILE_SCOPES), cache_discovery=False)


def _escape_drive_query(value: str) -> str:
    return value.replace("\\", "\\\\").replace("'", "\\'")


def find_or_create_drive_folder(service, parent_id: str, name: str) -> str:
    folder_name = str(name or "").strip()
    if not folder_name:
        raise RuntimeError("Folder name is required")

    query = (
        f"'{_escape_drive_query(parent_id)}' in parents and "
        f"mimeType = '{DRIVE_FOLDER_MIME_TYPE}' and "
        f"name = '{_escape_drive_query(folder_name)}' and trashed = false"
    )
    response = (
        service.files()
        .list(
            q=query,
            spaces="drive",
            fields="files(id, name)",
            pageSize=1,
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
        )
        .execute()
    )
    files = response.get("files", [])
    if files:
        return files[0]["id"]

    folder = (
        service.files()
        .create(
            body={
                "name": folder_name,
                "mimeType": DRIVE_FOLDER_MIME_TYPE,
                "parents": [parent_id],
            },
            fields="id",
            supportsAllDrives=True,
        )
        .execute()
    )
    return folder["id"]


def drive_upload_cmr() -> int:
    payload = json.loads(sys.stdin.read() or "{}")
    country_folder_id = str(payload.get("country_folder_id") or "").strip()
    folder_path = payload.get("folder_path") if isinstance(payload.get("folder_path"), list) else []
    filename = str(payload.get("filename") or "cmr-upload").strip()
    mime_type = str(payload.get("mime_type") or mimetypes.guess_type(filename)[0] or "application/octet-stream").strip()
    content_base64 = str(payload.get("content_base64") or "")

    if not country_folder_id:
        raise RuntimeError("Country folder ID is required")
    if not content_base64:
        raise RuntimeError("CMR file content is required")

    service = _drive_service()
    parent_id = country_folder_id
    for folder_name in folder_path:
        parent_id = find_or_create_drive_folder(service, parent_id, str(folder_name))

    file_bytes = base64.b64decode(content_base64)
    with tempfile.NamedTemporaryFile(delete=False) as temp_file:
        temp_file.write(file_bytes)
        temp_path = temp_file.name

    try:
        media = MediaFileUpload(temp_path, mimetype=mime_type, resumable=False)
        uploaded = (
            service.files()
            .create(
                body={"name": filename, "parents": [parent_id]},
                media_body=media,
                fields="id, name, webViewLink, webContentLink",
                supportsAllDrives=True,
            )
            .execute()
        )
    finally:
        Path(temp_path).unlink(missing_ok=True)

    sys.stdout.write(json.dumps(uploaded, ensure_ascii=True))
    return 0


def sheets_write_first_empty(spreadsheet_id: str, sheet_name: str) -> int:
    payload = json.loads(sys.stdin.read() or "{}")
    row = payload.get("row") if isinstance(payload, dict) else []
    if not isinstance(row, list) or not row:
        raise RuntimeError("No row data provided")

    last_column = _column_name(len(row))
    sheet_range = f"{sheet_name}!A:{last_column}"
    service = build("sheets", "v4", credentials=_service_account_credentials(), cache_discovery=False)
    existing = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=sheet_range)
        .execute()
    )
    values = existing.get("values", [])
    target_row = len(values) + 1 if values else 2

    for index, existing_row in enumerate(values[1:], start=2):
        if not any(str(cell or "").strip() for cell in existing_row):
            target_row = index
            break

    update_range = f"{sheet_name}!A{target_row}:{last_column}{target_row}"
    response = (
        service.spreadsheets()
        .values()
        .update(
            spreadsheetId=spreadsheet_id,
            range=update_range,
            valueInputOption="USER_ENTERED",
            body={"values": [row]},
        )
        .execute()
    )
    sys.stdout.write(json.dumps({"row_number": target_row, "response": response}, ensure_ascii=True))
    return 0


def email_send() -> int:
    payload = json.loads(sys.stdin.read() or "{}")
    recipients = payload.get("recipients") if isinstance(payload, dict) else []
    subject = str(payload.get("subject") or "")
    body = str(payload.get("body") or "")
    smtp = payload.get("smtp") if isinstance(payload, dict) and isinstance(payload.get("smtp"), dict) else {}

    if not recipients:
      raise RuntimeError("No recipients provided")

    smtp_host = str(smtp.get("host") or os.getenv("SMTP_HOST") or "").strip()
    smtp_port = int(smtp.get("port") or os.getenv("SMTP_PORT", "587"))
    smtp_username = str(smtp.get("username") or os.getenv("SMTP_USERNAME") or "").strip()
    smtp_password = str(smtp.get("password") or os.getenv("SMTP_PASSWORD") or "")
    smtp_from = str(smtp.get("from") or os.getenv("SMTP_FROM") or smtp_username).strip()
    smtp_starttls = smtp.get("starttls")
    if smtp_starttls is None:
        smtp_starttls = os.getenv("SMTP_STARTTLS", "1") != "0"
    else:
        smtp_starttls = False if str(smtp_starttls).lower() in {"0", "false", "no"} else True

    if not smtp_host or not smtp_from:
      raise RuntimeError("Set SMTP host and sender email in Settings, or configure SMTP_HOST and SMTP_FROM/SMTP_USERNAME")

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = smtp_from
    message["To"] = ", ".join(recipients)
    message.set_content(body)

    with smtplib.SMTP(smtp_host, smtp_port, timeout=30) as server:
        server.ehlo()
        if smtp_starttls:
            server.starttls()
            server.ehlo()
        if smtp_username:
            server.login(smtp_username, smtp_password or "")
        server.send_message(message)

    sys.stdout.write(json.dumps({"ok": True}, ensure_ascii=True))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(dest="command", required=True)
    subparsers.add_parser("details")
    image_parser = subparsers.add_parser("image")
    image_parser.add_argument("--account", default=DEFAULT_DRIVE_ACCOUNT)
    image_parser.add_argument("file_id")
    sheets_read_parser = subparsers.add_parser("sheets-read")
    sheets_read_parser.add_argument("--spreadsheet-id", required=True)
    sheets_read_parser.add_argument("--sheet-name", required=True)
    sheets_append_parser = subparsers.add_parser("sheets-append")
    sheets_append_parser.add_argument("--spreadsheet-id", required=True)
    sheets_append_parser.add_argument("--sheet-name", required=True)
    sheets_write_first_empty_parser = subparsers.add_parser("sheets-write-first-empty")
    sheets_write_first_empty_parser.add_argument("--spreadsheet-id", required=True)
    sheets_write_first_empty_parser.add_argument("--sheet-name", required=True)
    subparsers.add_parser("email-send")
    subparsers.add_parser("service-account-info")
    subparsers.add_parser("drive-upload-cmr")
    args = parser.parse_args()

    if args.command == "details":
        return details()
    if args.command == "image":
        return image(args.file_id, args.account)
    if args.command == "sheets-read":
        return sheets_read(args.spreadsheet_id, args.sheet_name)
    if args.command == "sheets-append":
        return sheets_append(args.spreadsheet_id, args.sheet_name)
    if args.command == "sheets-write-first-empty":
        return sheets_write_first_empty(args.spreadsheet_id, args.sheet_name)
    if args.command == "email-send":
        return email_send()
    if args.command == "service-account-info":
        return service_account_info()
    if args.command == "drive-upload-cmr":
        return drive_upload_cmr()

    return 1


if __name__ == "__main__":
    raise SystemExit(main())
