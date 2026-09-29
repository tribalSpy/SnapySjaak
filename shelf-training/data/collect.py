"""
Step 1: dataset collection for the shelf-counting model.

Downloads a sample of trolley photos from Google Drive into a local dataset
folder, stratified across customers and dates, prioritizing photos
shadow-app already flagged "needs_review" (the hard cases worth labeling
first -- the "deviation" category doesn't exist yet, since expected_average
isn't populated until a later phase actually has something to derive it
from).

Talks to shadow-app over HTTP with the same poller API key rather than
connecting to Postgres directly -- this script never needs database
credentials, only Drive (read-only, service account) and the one shadow-app
endpoint built for it (GET /api/shelf-count/dataset-candidates).

Usage:
    python collect.py --config config.json --from-date 2026-09-01 --to-date 2026-09-29 --count 200
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import random
import sys
import urllib.error
import urllib.request
from collections import defaultdict
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
TRAINING_ROOT = SCRIPT_DIR.parent
REPO_ROOT = TRAINING_ROOT.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from dotenv import load_dotenv  # noqa: E402

load_dotenv(REPO_ROOT / ".env")  # GOOGLE_SERVICE_ACCOUNT_JSON / GOOGLE_DRIVE_ROOT_FOLDER_ID, same as drive_bridge.py

from src.drive_service import DriveService, DEFAULT_DRIVE_ACCOUNT  # noqa: E402
from src.parser import parse_run_folder_name  # noqa: E402

MANIFEST_FIELDS = [
    "file_id", "folder_name", "customer_reference", "date", "local_path",
    "sha256", "priority", "status", "confidence",
]


def load_config(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def api_request(config: dict, path: str) -> dict:
    request = urllib.request.Request(
        url=f'{config["server_url"].rstrip("/")}{path}',
        headers={"Accept": "application/json", "x-shadow-agent-key": config["api_key"]},
    )
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            body = response.read().decode("utf-8")
            return json.loads(body) if body else {}
    except urllib.error.HTTPError as error:
        body = error.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {error.code}: {body}") from error
    except urllib.error.URLError as error:
        raise RuntimeError(f"Network error: {error}") from error


def fetch_priority_candidates(config: dict, date_from: str, date_to: str) -> list[dict]:
    """References shadow-app already flagged needs_review -- the hard cases
    worth labeling first (per Step 1's own priority)."""
    payload = api_request(
        config,
        f"/api/shelf-count/dataset-candidates?from={date_from}&to={date_to}&status=needs_review",
    )
    return payload.get("rows", [])


def discover_general_folders(drive: DriveService, root_folder_id: str, date_from: str, date_to: str) -> list[dict]:
    """Every Drive folder under the root dated within range, regardless of
    whether shadow-app ever processed it -- the general stratified-sampling
    pool. Reuses parse_run_folder_name (the same parser the manual photo
    viewer already relies on) rather than re-deriving the customer_YYYYMMDD
    convention here."""
    start = datetime.strptime(date_from, "%Y-%m-%d").date()
    end = datetime.strptime(date_to, "%Y-%m-%d").date()
    folders = []
    for child in drive.list_child_folders(root_folder_id):
        name = str(child.get("name") or "").strip()
        try:
            parsed = parse_run_folder_name(name)
        except ValueError:
            continue
        if start <= parsed.run_date <= end:
            folders.append({
                "folder_id": child.get("id"),
                "name": name,
                "customer_code": parsed.customer_code,
                "run_date": parsed.run_date.isoformat(),
            })
    return folders


def stratified_sample(folders: list[dict], target_count: int) -> list[dict]:
    """Round-robins across customer groups so one busy customer can't fill
    the whole sample -- a simple, transparent stratification, not a fancy
    one. Shuffled within each group too, so it also spreads across dates/
    lighting conditions for a given customer rather than always picking
    that customer's earliest folders."""
    by_customer = defaultdict(list)
    for folder in folders:
        by_customer[folder["customer_code"]].append(folder)
    groups = list(by_customer.values())
    for group in groups:
        random.shuffle(group)
    random.shuffle(groups)

    selected = []
    while len(selected) < target_count and groups:
        for group in list(groups):
            if not group:
                groups.remove(group)
                continue
            selected.append(group.pop())
            if len(selected) >= target_count:
                break
    return selected[:target_count]


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def load_seen_hashes(manifest_path: Path) -> set[str]:
    if not manifest_path.exists():
        return set()
    with open(manifest_path, newline="", encoding="utf-8") as handle:
        return {row["sha256"] for row in csv.DictReader(handle) if row.get("sha256")}


def append_manifest(manifest_path: Path, rows: list[dict]):
    is_new = not manifest_path.exists()
    with open(manifest_path, "a", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=MANIFEST_FIELDS)
        if is_new:
            writer.writeheader()
        writer.writerows(rows)


def download_folder_photos(
    drive: DriveService, folder_id: str, folder_name: str, customer_reference: str,
    folder_date: str, dataset_dir: Path, seen_hashes: set[str], priority: str,
    status: str, confidence,
) -> list[dict]:
    out_dir = dataset_dir / "raw" / folder_name
    manifest_rows = []
    for file_info in drive.list_files(folder_id):
        mime_type = str(file_info.get("mimeType") or "")
        if not mime_type.startswith("image/"):
            continue
        file_id = file_info.get("id")
        name = file_info.get("name") or f"{file_id}.jpg"
        try:
            content = drive.download_file_bytes(file_id)
        except Exception as error:
            print(f"  skip {name}: download failed ({error})")
            continue
        digest = sha256_bytes(content)
        if digest in seen_hashes:
            print(f"  skip {name}: duplicate (already in dataset)")
            continue
        seen_hashes.add(digest)
        out_dir.mkdir(parents=True, exist_ok=True)
        local_path = out_dir / name
        local_path.write_bytes(content)
        manifest_rows.append({
            "file_id": file_id,
            "folder_name": folder_name,
            "customer_reference": customer_reference,
            "date": folder_date,
            "local_path": str(local_path.relative_to(dataset_dir)),
            "sha256": digest,
            "priority": priority,
            "status": status,
            "confidence": confidence if confidence is not None else "",
        })
    return manifest_rows


def main():
    parser = argparse.ArgumentParser(description="Step 1: collect a labeling dataset from Drive")
    parser.add_argument("--config", default=str(SCRIPT_DIR / "config.json"))
    parser.add_argument("--from-date", dest="date_from", required=True, help="YYYY-MM-DD")
    parser.add_argument("--to-date", dest="date_to", required=True, help="YYYY-MM-DD")
    parser.add_argument("--count", type=int, default=200, help="Total photos to collect (priority needs_review cases are downloaded in full first and don't count against this cap)")
    parser.add_argument("--customer-references", default="", help="Comma-separated list to restrict the GENERAL sample to (the priority tier is never restricted)")
    args = parser.parse_args()

    config = load_config(Path(args.config))
    dataset_dir = Path(config.get("dataset_dir") or (TRAINING_ROOT / "data" / "dataset"))
    dataset_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = dataset_dir / "manifest.csv"
    seen_hashes = load_seen_hashes(manifest_path)

    drive = DriveService.from_service_account_env(config.get("drive_account", DEFAULT_DRIVE_ACCOUNT))
    root_folder_id = config["drive_root_folder_id"]

    print(f"Fetching priority (needs_review) candidates from shadow-app for {args.date_from}..{args.date_to}...")
    priority_rows = fetch_priority_candidates(config, args.date_from, args.date_to)
    print(f"  {len(priority_rows)} needs_review reference(s) found")

    manifest_rows = []
    handled_folders = set()
    root_children = None
    for row in priority_rows:
        folder_name = row.get("drive_folder_name") or ""
        if not folder_name or folder_name in handled_folders:
            continue
        handled_folders.add(folder_name)
        if root_children is None:
            root_children = drive.list_child_folders(root_folder_id)
        match = next((f for f in root_children if f.get("name") == folder_name), None)
        if not match:
            print(f"  {folder_name}: not found on Drive anymore, skipping")
            continue
        print(f"  downloading priority folder {folder_name} (needs_review, confidence={row.get('confidence')})")
        manifest_rows.extend(download_folder_photos(
            drive, match["id"], folder_name, row.get("customer_reference", ""), row.get("nightly_run_date", ""),
            dataset_dir, seen_hashes, priority="needs_review", status=row.get("status", ""), confidence=row.get("confidence"),
        ))

    remaining = max(0, args.count - len(manifest_rows))
    if remaining:
        print(f"Discovering general folders on Drive for stratified sampling ({remaining} more photos targeted)...")
        general_folders = discover_general_folders(drive, root_folder_id, args.date_from, args.date_to)
        general_folders = [f for f in general_folders if f["name"] not in handled_folders]
        if args.customer_references.strip():
            wanted = {c.strip().upper() for c in args.customer_references.split(",") if c.strip()}
            general_folders = [f for f in general_folders if f["customer_code"].upper() in wanted]
        sampled = stratified_sample(general_folders, remaining)
        print(f"  sampled {len(sampled)} of {len(general_folders)} candidate folder(s)")
        for folder in sampled:
            print(f"  downloading {folder['name']}")
            manifest_rows.extend(download_folder_photos(
                drive, folder["folder_id"], folder["name"], folder["customer_code"], folder["run_date"],
                dataset_dir, seen_hashes, priority="general", status="", confidence=None,
            ))

    append_manifest(manifest_path, manifest_rows)
    print(f"\nDone. {len(manifest_rows)} new photo(s) added to {manifest_path}")


if __name__ == "__main__":
    main()
