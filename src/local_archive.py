from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path

from src.models import ParseError, RunFolder
from src.parser import parse_run_folder_name

DATE_FOLDER_FORMAT = "%Y-%m-%d"


def date_folder_name(run_date: date) -> str:
    return run_date.strftime(DATE_FOLDER_FORMAT)


def parse_date_folder_name(folder_name: str) -> date | None:
    for date_format in (DATE_FOLDER_FORMAT, "%Y%m%d"):
        try:
            return datetime.strptime(folder_name, date_format).date()
        except ValueError:
            pass
    return None


def _is_run_folder(folder: Path) -> bool:
    try:
        parse_run_folder_name(folder.name)
    except ValueError:
        return False
    return True


def _single_carrier_folder(archive_root: Path) -> Path | None:
    carrier_folders: list[Path] = []

    for child_folder in [path for path in archive_root.iterdir() if path.is_dir()]:
        if parse_date_folder_name(child_folder.name) is not None:
            continue
        if _is_run_folder(child_folder):
            continue
        carrier_folders.append(child_folder)

    if len(carrier_folders) != 1:
        return None
    return carrier_folders[0]


def _move_folder_without_overwrite(source: Path, target: Path) -> bool:
    if target.exists():
        return False

    target.parent.mkdir(parents=True, exist_ok=True)
    source.rename(target)
    return True


def _regroup_run_folders_in_container(container: Path, target_container: Path | None = None) -> int:
    moved_count = 0
    destination_container = target_container or container
    child_folders = [path for path in container.iterdir() if path.is_dir()]

    for child_folder in child_folders:
        try:
            parsed = parse_run_folder_name(child_folder.name)
        except ValueError:
            continue

        target_parent = destination_container / date_folder_name(parsed.run_date)
        target_path = target_parent / child_folder.name
        if _move_folder_without_overwrite(child_folder, target_path):
            moved_count += 1

    return moved_count


def _move_root_date_folders_into_carrier(archive_root: Path, carrier_folder: Path) -> int:
    moved_count = 0
    root_date_folders = [
        path
        for path in archive_root.iterdir()
        if path.is_dir() and parse_date_folder_name(path.name) is not None
    ]

    for root_date_folder in root_date_folders:
        target_date_folder = carrier_folder / root_date_folder.name
        run_folders = [path for path in root_date_folder.iterdir() if path.is_dir()]

        for run_folder in run_folders:
            target_path = target_date_folder / run_folder.name
            if _move_folder_without_overwrite(run_folder, target_path):
                moved_count += 1

        try:
            root_date_folder.rmdir()
        except OSError:
            pass

    return moved_count


def regroup_direct_run_folders_by_date(archive_root: Path) -> int:
    if not archive_root.exists():
        return 0

    carrier_folder = _single_carrier_folder(archive_root)
    moved_count = 0
    if carrier_folder is not None:
        moved_count += _move_root_date_folders_into_carrier(archive_root, carrier_folder)
    moved_count += _regroup_run_folders_in_container(archive_root, carrier_folder)
    child_folders = [path for path in archive_root.iterdir() if path.is_dir()]

    for child_folder in child_folders:
        if parse_date_folder_name(child_folder.name) is not None:
            continue

        try:
            parse_run_folder_name(child_folder.name)
            continue
        except ValueError:
            moved_count += _regroup_run_folders_in_container(child_folder)

    return moved_count


def _date_is_in_range(
    folder_date: date,
    minimum_date: date | None,
    maximum_date: date | None,
) -> bool:
    if minimum_date is not None and folder_date < minimum_date:
        return False
    if maximum_date is not None and folder_date >= maximum_date:
        return False
    return True


def _build_local_run_index_for_container(
    container: Path,
    carrier_name: str,
    parse_errors: list[ParseError],
    minimum_date: date | None = None,
    maximum_date: date | None = None,
) -> list[RunFolder]:
    runs: list[RunFolder] = []
    direct_run_folders: list[Path] = []
    date_folders: list[Path] = []

    for child_folder in [path for path in container.iterdir() if path.is_dir()]:
        folder_date = parse_date_folder_name(child_folder.name)
        if folder_date is not None:
            if _date_is_in_range(folder_date, minimum_date, maximum_date):
                date_folders.append(child_folder)
            continue

        try:
            parsed = parse_run_folder_name(child_folder.name)
        except ValueError:
            continue

        if _date_is_in_range(parsed.run_date, minimum_date, maximum_date):
            direct_run_folders.append(child_folder)

    if direct_run_folders:
        runs.extend(
            build_local_run_index_for_folders(
                run_folders=direct_run_folders,
                carrier_name=carrier_name,
                carrier_folder_id=str(container),
                parse_errors=parse_errors,
                minimum_date=minimum_date,
                maximum_date=maximum_date,
            )
        )

    for date_folder in date_folders:
        run_folders = [path for path in date_folder.iterdir() if path.is_dir()]
        runs.extend(
            build_local_run_index_for_folders(
                run_folders=run_folders,
                carrier_name=carrier_name,
                carrier_folder_id=str(date_folder),
                parse_errors=parse_errors,
                minimum_date=minimum_date,
                maximum_date=maximum_date,
            )
        )

    return runs


def load_local_archive_run_index(
    archive_root: Path,
    parse_errors: list[ParseError],
    archive_cutoff_date: date | None,
) -> list[RunFolder]:
    runs: list[RunFolder] = []

    if not archive_root.exists():
        parse_errors.append(
            ParseError(
                folder_id=str(archive_root),
                folder_name=archive_root.name or str(archive_root),
                carrier="local-archive",
                reason="Local archive root path does not exist",
            )
        )
        return runs

    child_folders = [path for path in archive_root.iterdir() if path.is_dir()]
    carrier_folders: list[Path] = []

    for child_folder in child_folders:
        if parse_date_folder_name(child_folder.name) is not None:
            continue

        try:
            parse_run_folder_name(child_folder.name)
        except ValueError:
            carrier_folders.append(child_folder)

    runs.extend(
        _build_local_run_index_for_container(
            container=archive_root,
            carrier_name=archive_root.name,
            parse_errors=parse_errors,
            maximum_date=archive_cutoff_date,
        )
    )

    for carrier_folder in carrier_folders:
        runs.extend(
            _build_local_run_index_for_container(
                container=carrier_folder,
                carrier_name=carrier_folder.name,
                parse_errors=parse_errors,
                maximum_date=archive_cutoff_date,
            )
        )

    return runs


def load_local_archive_run_index_for_date(
    archive_root: Path,
    parse_errors: list[ParseError],
    selected_date: date,
) -> list[RunFolder]:
    runs: list[RunFolder] = []

    if not archive_root.exists():
        parse_errors.append(
            ParseError(
                folder_id=str(archive_root),
                folder_name=archive_root.name or str(archive_root),
                carrier="local-archive",
                reason="Local archive root path does not exist",
            )
        )
        return runs

    child_folders = [path for path in archive_root.iterdir() if path.is_dir()]
    carrier_folders: list[Path] = []

    for child_folder in child_folders:
        if parse_date_folder_name(child_folder.name) is not None:
            continue

        try:
            parse_run_folder_name(child_folder.name)
        except ValueError:
            carrier_folders.append(child_folder)

    runs.extend(
        _build_local_run_index_for_container(
            container=archive_root,
            carrier_name=archive_root.name,
            parse_errors=parse_errors,
            minimum_date=selected_date,
            maximum_date=selected_date + timedelta(days=1),
        )
    )

    for carrier_folder in carrier_folders:
        runs.extend(
            _build_local_run_index_for_container(
                container=carrier_folder,
                carrier_name=carrier_folder.name,
                parse_errors=parse_errors,
                minimum_date=selected_date,
                maximum_date=selected_date + timedelta(days=1),
            )
        )

    return runs


def build_local_run_index_for_folders(
    run_folders: list[Path],
    carrier_name: str,
    carrier_folder_id: str,
    parse_errors: list[ParseError],
    minimum_date: date | None = None,
    maximum_date: date | None = None,
) -> list[RunFolder]:
    runs: list[RunFolder] = []

    for run_folder in run_folders:
        try:
            parsed = parse_run_folder_name(run_folder.name)
        except ValueError as exc:
            parse_errors.append(
                ParseError(
                    folder_id=str(run_folder),
                    folder_name=run_folder.name,
                    carrier=carrier_name,
                    reason=str(exc),
                )
            )
            continue

        if minimum_date is not None and parsed.run_date < minimum_date:
            continue
        if maximum_date is not None and parsed.run_date >= maximum_date:
            continue

        runs.append(
            RunFolder(
                folder_id=str(run_folder),
                folder_name=run_folder.name,
                customer_code=parsed.customer_code,
                run_date=parsed.run_date,
                carrier=carrier_name,
                run_id=parsed.run_id,
                metadata={
                    "carrier_folder_id": carrier_folder_id,
                    "source": "local_archive",
                },
            )
        )

    return runs
