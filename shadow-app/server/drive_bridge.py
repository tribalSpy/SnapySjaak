from __future__ import annotations

import argparse
import json
import os
import smtplib
import sys
from pathlib import Path
from email.message import EmailMessage

from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

load_dotenv(REPO_ROOT / ".env")

from src.drive_service import DEFAULT_DRIVE_ACCOUNT, DriveService  # noqa: E402

SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


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


def _service_account_credentials() -> Credentials:
    credentials_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if credentials_json:
        return Credentials.from_service_account_info(
            json.loads(credentials_json),
            scopes=SHEETS_SCOPES,
        )

    credentials_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
    if credentials_path:
        return Credentials.from_service_account_file(
            credentials_path,
            scopes=SHEETS_SCOPES,
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
    subparsers.add_parser("email-send")
    subparsers.add_parser("service-account-info")
    args = parser.parse_args()

    if args.command == "details":
        return details()
    if args.command == "image":
        return image(args.file_id, args.account)
    if args.command == "sheets-read":
        return sheets_read(args.spreadsheet_id, args.sheet_name)
    if args.command == "sheets-append":
        return sheets_append(args.spreadsheet_id, args.sheet_name)
    if args.command == "email-send":
        return email_send()
    if args.command == "service-account-info":
        return service_account_info()

    return 1


if __name__ == "__main__":
    raise SystemExit(main())
