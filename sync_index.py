from __future__ import annotations

import argparse
import json
import os
from datetime import date, datetime, timedelta
from pathlib import Path

from dotenv import load_dotenv

from src.drive_service import (
    DEFAULT_DRIVE_ACCOUNT,
    DriveService,
    get_setting,
    list_drive_configurations,
)
from src.local_archive import (
    load_local_archive_run_index,
    load_local_archive_run_index_for_date,
    regroup_direct_run_folders_by_date,
)
from src.models import ParseError, RunFolder
from src.parser import parse_run_folder_name

load_dotenv()

CACHE_DIR = Path(os.getenv("SNAPPYSJAAK_CACHE_DIR", ".cache")).expanduser()
RUN_DATA_CACHE_PATH = CACHE_DIR / "run_data.json"
RUN_DATA_CACHE_VERSION = "v2"
INDEX_SYNC_STATUS_PATH = CACHE_DIR / "index_sync_status.json"
INDEX_SYNC_STALE_MINUTES = 120


def _load_configured_google_runs(
    parse_errors: list[ParseError],
    archive_cutoff_date: date,
    maximum_date: date | None = None,
) -> list[RunFolder]:
    runs: list[RunFolder] = []
    for drive_config in list_drive_configurations():
        drive_service = DriveService.from_service_account_env(drive_config.account_name)
        runs.extend(
            _load_google_drive_run_index(
                drive_service=drive_service,
                root_folder_id=drive_config.root_folder_id,
                parse_errors=parse_errors,
                archive_cutoff_date=archive_cutoff_date,
                maximum_date=maximum_date,
                drive_account=drive_config.account_name,
            )
        )
    return runs


def _write_status(state: str, mode: str, **extra: object) -> None:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    payload: dict[str, object] = {
        "state": state,
        "mode": mode,
        "updated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
    }
    payload.update(extra)
    INDEX_SYNC_STATUS_PATH.write_text(
        json.dumps(payload, ensure_ascii=True, indent=2),
        encoding="utf-8",
    )


def _run_identity(run: RunFolder) -> tuple[str, str, date, str, str]:
    return (
        run.carrier.casefold(),
        run.customer_code.casefold(),
        run.run_date,
        (run.run_id or "").casefold(),
        run.folder_name.casefold(),
    )


def _merge_runs_prefer_google_drive(primary_runs: list[RunFolder], fallback_runs: list[RunFolder]) -> list[RunFolder]:
    merged: dict[tuple[str, str, date, str, str], RunFolder] = {
        _run_identity(run): run for run in fallback_runs
    }
    for run in primary_runs:
        merged[_run_identity(run)] = run
    return list(merged.values())


def _deduplicate_parse_errors(parse_errors: list[ParseError]) -> list[ParseError]:
    unique_errors: dict[tuple[str, str, str, str], ParseError] = {}
    for error in parse_errors:
        unique_errors[
            (
                error.folder_id,
                error.folder_name,
                error.carrier,
                error.reason,
            )
        ] = error
    return list(unique_errors.values())


def _parse_error_date(error: ParseError) -> date | None:
    try:
        return parse_run_folder_name(error.folder_name).run_date
    except ValueError:
        return None


def _serialize_runs_payload(
    runs: list[RunFolder],
    parse_errors: list[ParseError],
) -> dict[str, object]:
    return {
        "runs": [
            {
                "folder_id": run.folder_id,
                "folder_name": run.folder_name,
                "customer_code": run.customer_code,
                "run_date": run.run_date.isoformat(),
                "carrier": run.carrier,
                "run_id": run.run_id,
                "images": [],
                "qr_info": "No QR info found",
                "qr_source": None,
                "metadata": dict(run.metadata),
            }
            for run in runs
        ],
        "parse_errors": [
            {
                "folder_id": error.folder_id,
                "folder_name": error.folder_name,
                "carrier": error.carrier,
                "reason": error.reason,
            }
            for error in parse_errors
        ],
    }


def _deserialize_runs_payload(
    payload: dict[str, object],
) -> tuple[list[RunFolder], list[ParseError]]:
    runs = [
        RunFolder(
            folder_id=item["folder_id"],
            folder_name=item["folder_name"],
            customer_code=item["customer_code"],
            run_date=date.fromisoformat(item["run_date"]),
            carrier=item["carrier"],
            run_id=item.get("run_id"),
            metadata=dict(item.get("metadata", {})),
        )
        for item in payload.get("runs", [])
    ]
    parse_errors = [
        ParseError(
            folder_id=item["folder_id"],
            folder_name=item["folder_name"],
            carrier=item["carrier"],
            reason=item["reason"],
        )
        for item in payload.get("parse_errors", [])
    ]
    return runs, parse_errors


def _read_persisted_run_data() -> tuple[list[RunFolder], list[ParseError]] | None:
    if not RUN_DATA_CACHE_PATH.exists():
        return None

    try:
        payload = json.loads(RUN_DATA_CACHE_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, ValueError):
        return None

    if payload.get("cache_version") != RUN_DATA_CACHE_VERSION:
        return None

    try:
        return _deserialize_runs_payload(payload)
    except (KeyError, TypeError, ValueError):
        return None


def _write_persisted_run_data(runs: list[RunFolder], parse_errors: list[ParseError]) -> None:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    payload = _serialize_runs_payload(runs, parse_errors)
    payload["cache_version"] = RUN_DATA_CACHE_VERSION
    payload["generated_at"] = datetime.now().astimezone().isoformat(timespec="seconds")
    RUN_DATA_CACHE_PATH.write_text(
        json.dumps(payload, ensure_ascii=True, indent=2),
        encoding="utf-8",
    )


def _existing_run_count() -> int:
    cached_payload = _read_persisted_run_data()
    if cached_payload is None:
        return 0
    runs, _parse_errors = cached_payload
    return len(runs)


def _assert_rebuild_is_not_empty(runs: list[RunFolder]) -> None:
    existing_count = _existing_run_count()
    if existing_count > 0 and not runs:
        raise RuntimeError(
            "Rebuild found 0 runs while an existing saved index contains "
            f"{existing_count} runs. Keeping the existing index."
        )


def _assert_refresh_kept_selected_date(
    selected_date: date,
    before_runs: list[RunFolder],
    after_runs: list[RunFolder],
) -> None:
    before_count = sum(1 for run in before_runs if run.run_date == selected_date)
    after_count = sum(1 for run in after_runs if run.run_date == selected_date)
    if before_count > 0 and after_count == 0:
        raise RuntimeError(
            f"Refresh for {selected_date.isoformat()} found 0 runs while the "
            f"existing saved index contains {before_count} runs for that date. "
            "Keeping the existing index."
        )


def _load_google_drive_run_index(
    drive_service: DriveService,
    root_folder_id: str,
    parse_errors: list[ParseError],
    archive_cutoff_date: date,
    maximum_date: date | None = None,
    drive_account: str = DEFAULT_DRIVE_ACCOUNT,
) -> list[RunFolder]:
    runs: list[RunFolder] = []
    root_folder = drive_service.get_file(root_folder_id)
    child_folders = drive_service.list_child_folders(root_folder_id)

    direct_run_folders: list[dict] = []
    carrier_folders: list[dict] = []

    for child_folder in child_folders:
        try:
            parsed = parse_run_folder_name(child_folder["name"])
            if parsed.run_date >= archive_cutoff_date:
                direct_run_folders.append(child_folder)
        except ValueError:
            carrier_folders.append(child_folder)

    if direct_run_folders:
        runs.extend(
            _build_google_run_index_for_folders(
                run_folders=direct_run_folders,
                carrier_name=root_folder["name"],
                carrier_folder_id=root_folder_id,
                parse_errors=parse_errors,
                minimum_date=archive_cutoff_date,
                maximum_date=maximum_date,
                drive_account=drive_account,
            )
        )

    for carrier_folder in carrier_folders:
        runs.extend(
            _build_google_run_index_for_folders(
                run_folders=drive_service.list_child_folders(carrier_folder["id"]),
                carrier_name=carrier_folder["name"],
                carrier_folder_id=carrier_folder["id"],
                parse_errors=parse_errors,
                minimum_date=archive_cutoff_date,
                maximum_date=maximum_date,
                drive_account=drive_account,
            )
        )

    return runs


def _build_google_run_index_for_folders(
    run_folders: list[dict],
    carrier_name: str,
    carrier_folder_id: str,
    parse_errors: list[ParseError],
    minimum_date: date,
    maximum_date: date | None = None,
    drive_account: str = DEFAULT_DRIVE_ACCOUNT,
) -> list[RunFolder]:
    runs: list[RunFolder] = []

    for run_folder in run_folders:
        folder_name = run_folder["name"]
        try:
            parsed = parse_run_folder_name(folder_name)
        except ValueError as exc:
            parse_errors.append(
                ParseError(
                    folder_id=run_folder["id"],
                    folder_name=folder_name,
                    carrier=carrier_name,
                    reason=str(exc),
                )
            )
            continue

        if parsed.run_date < minimum_date:
            continue
        if maximum_date is not None and parsed.run_date > maximum_date:
            continue

        runs.append(
            RunFolder(
                folder_id=run_folder["id"],
                folder_name=folder_name,
                customer_code=parsed.customer_code,
                run_date=parsed.run_date,
                carrier=carrier_name,
                run_id=parsed.run_id,
                metadata={
                    "carrier_folder_id": carrier_folder_id,
                    "source": "google_drive",
                    "drive_account": drive_account,
                },
            )
        )

    return runs


def rebuild_index() -> tuple[list[RunFolder], list[ParseError]]:
    runs: list[RunFolder] = []
    parse_errors: list[ParseError] = []
    archive_cutoff_days = int(get_setting("LOCAL_ARCHIVE_AFTER_DAYS") or "7")
    archive_cutoff_date = date.today() - timedelta(days=archive_cutoff_days)

    try:
        runs.extend(
            _load_configured_google_runs(
                parse_errors=parse_errors,
                archive_cutoff_date=archive_cutoff_date,
            )
        )
    except RuntimeError:
        pass

    local_archive_root = get_setting("LOCAL_ARCHIVE_ROOT")
    if local_archive_root:
        archive_root = Path(local_archive_root)
        regroup_direct_run_folders_by_date(archive_root)
        local_runs = load_local_archive_run_index(
            archive_root=archive_root,
            parse_errors=parse_errors,
            archive_cutoff_date=None,
        )
        runs = _merge_runs_prefer_google_drive(runs, local_runs)

    runs.sort(
        key=lambda run: (
            run.run_date,
            run.customer_code.lower(),
            run.carrier.lower(),
            run.run_id or "",
        )
    )
    parse_errors = _deduplicate_parse_errors(parse_errors)
    return runs, parse_errors


def refresh_date_index(selected_date: date) -> tuple[list[RunFolder], list[ParseError]]:
    cached_payload = _read_persisted_run_data()
    existing_runs, existing_parse_errors = cached_payload if cached_payload is not None else ([], [])

    runs: list[RunFolder] = [
        run for run in existing_runs if run.run_date != selected_date
    ]
    parse_errors: list[ParseError] = [
        error for error in existing_parse_errors if _parse_error_date(error) != selected_date
    ]
    fresh_parse_errors: list[ParseError] = []

    try:
        google_runs = _load_configured_google_runs(
            parse_errors=fresh_parse_errors,
            archive_cutoff_date=selected_date,
            maximum_date=selected_date,
        )
        runs.extend(google_runs)
    except RuntimeError:
        pass

    local_archive_root = get_setting("LOCAL_ARCHIVE_ROOT")
    if local_archive_root:
        archive_root = Path(local_archive_root)
        regroup_direct_run_folders_by_date(archive_root)
        local_runs = load_local_archive_run_index_for_date(
            archive_root=archive_root,
            parse_errors=fresh_parse_errors,
            selected_date=selected_date,
        )
        runs = _merge_runs_prefer_google_drive(runs, local_runs)

    runs.sort(
        key=lambda run: (
            run.run_date,
            run.customer_code.lower(),
            run.carrier.lower(),
            run.run_id or "",
        )
    )
    parse_errors = _deduplicate_parse_errors(parse_errors + fresh_parse_errors)
    _assert_refresh_kept_selected_date(selected_date, existing_runs, runs)
    return runs, parse_errors


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--mode", choices=["rebuild", "refresh_date"], required=True)
    parser.add_argument("--date")
    args = parser.parse_args()

    started_at = datetime.now().astimezone().isoformat(timespec="seconds")
    _write_status("running", args.mode, started_at=started_at)

    try:
        if args.mode == "rebuild":
            runs, parse_errors = rebuild_index()
            _assert_rebuild_is_not_empty(runs)
        else:
            if not args.date:
                raise ValueError("--date is required for refresh_date mode")
            runs, parse_errors = refresh_date_index(date.fromisoformat(args.date))

        _write_persisted_run_data(runs, parse_errors)
        _write_status(
            "completed",
            args.mode,
            started_at=started_at,
            run_count=len(runs),
            parse_error_count=len(parse_errors),
        )
        return 0
    except Exception as exc:
        _write_status(
            "failed",
            args.mode,
            started_at=started_at,
            error=str(exc),
        )
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
