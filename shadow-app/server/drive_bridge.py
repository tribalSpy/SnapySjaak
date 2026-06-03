from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from dotenv import load_dotenv

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

load_dotenv(REPO_ROOT / ".env")

from src.drive_service import DEFAULT_DRIVE_ACCOUNT, DriveService  # noqa: E402


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


def main() -> int:
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(dest="command", required=True)
    subparsers.add_parser("details")
    image_parser = subparsers.add_parser("image")
    image_parser.add_argument("--account", default=DEFAULT_DRIVE_ACCOUNT)
    image_parser.add_argument("file_id")
    args = parser.parse_args()

    if args.command == "details":
        return details()
    if args.command == "image":
        return image(args.file_id, args.account)

    return 1


if __name__ == "__main__":
    raise SystemExit(main())
