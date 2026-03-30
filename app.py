from __future__ import annotations

import os
import json
import hashlib
import mimetypes
import shutil
from collections import defaultdict
from datetime import date, timedelta
from io import BytesIO
from math import ceil
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv
from googleapiclient.errors import HttpError
from PIL import Image, ImageOps

from src.drive_service import DriveService, get_setting, validate_drive_configuration
from src.models import ImageFile, ParseError, RunFolder
from src.parser import parse_run_folder_name
from src.ui_helpers import (
    group_runs_by_customer,
    render_parse_errors,
)

load_dotenv()

st.set_page_config(
    page_title="Drive Run Dashboard",
    page_icon=":open_file_folder:",
    layout="wide",
)

st.markdown(
    """
    <style>
    .customer-card {
        border: 1px solid #d6dde8;
        border-radius: 14px;
        padding: 1rem 1.1rem;
        margin: 0.75rem 0 0.35rem 0;
        background: linear-gradient(180deg, #ffffff 0%, #f7faff 100%);
        box-shadow: 0 1px 2px rgba(16, 24, 40, 0.04);
    }
    .customer-card-title {
        font-size: 1.05rem;
        font-weight: 700;
        color: #10233f;
        margin-bottom: 0.2rem;
    }
    .customer-card-meta {
        font-size: 0.92rem;
        color: #516074;
    }
    .customer-card-open {
        border-left: 4px solid #2b6df3;
        background: linear-gradient(180deg, #f8fbff 0%, #eef5ff 100%);
    }
    .run-shell {
        border: 1px solid #e3e8f0;
        border-radius: 12px;
        padding: 1rem;
        margin: 0.5rem 0 1rem 0;
        background: #ffffff;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

CACHE_DIR = Path(".cache")
RUN_DATA_CACHE_PATH = CACHE_DIR / "run_data.json"
RUN_DATA_CACHE_VERSION = "v2"
IMAGE_CACHE_DIR = CACHE_DIR / "images"
IMAGE_CACHE_VERSION = "v4"
USB_CAMERA_AUTO_ROTATE_CARRIERS = {"CargoSnapArchive"}


@st.cache_resource(show_spinner=False)
def get_drive_service() -> DriveService:
    return DriveService.from_service_account_env()


@st.cache_data(ttl=300, show_spinner=False)
def load_run_data(_refresh_token: int) -> tuple[list[RunFolder], list[ParseError]]:
    cached_payload = _read_persisted_run_data()
    if cached_payload is not None:
        return cached_payload

    runs: list[RunFolder] = []
    parse_errors: list[ParseError] = []
    archive_cutoff_days = int(get_setting("LOCAL_ARCHIVE_AFTER_DAYS") or "7")
    archive_cutoff_date = date.today() - timedelta(days=archive_cutoff_days)
    google_runs: list[RunFolder] = []

    try:
        root_folder_id = validate_drive_configuration()
        drive_service = get_drive_service()
        google_runs = _load_google_drive_runs(
            drive_service=drive_service,
            root_folder_id=root_folder_id,
            parse_errors=parse_errors,
            archive_cutoff_date=archive_cutoff_date,
        )
        runs.extend(google_runs)
    except RuntimeError:
        # Allow local archive-only usage if Drive is not configured.
        pass

    local_archive_root = get_setting("LOCAL_ARCHIVE_ROOT")
    if local_archive_root:
        local_runs = _load_local_archive_runs(
            archive_root=Path(local_archive_root),
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
    _write_persisted_run_data(runs, parse_errors)
    return runs, parse_errors


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
        runs = [
            RunFolder(
                folder_id=item["folder_id"],
                folder_name=item["folder_name"],
                customer_code=item["customer_code"],
                run_date=date.fromisoformat(item["run_date"]),
                carrier=item["carrier"],
                run_id=item.get("run_id"),
                images=[
                    ImageFile(
                        id=image["id"],
                        name=image["name"],
                        mime_type=image["mime_type"],
                        web_view_link=image.get("web_view_link"),
                        size=image.get("size"),
                    )
                    for image in item.get("images", [])
                ],
                qr_info=item.get("qr_info", "No QR info found"),
                qr_source=item.get("qr_source"),
                metadata=item.get("metadata", {}),
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
    except (KeyError, TypeError, ValueError):
        return None

    return runs, parse_errors


def _write_persisted_run_data(runs: list[RunFolder], parse_errors: list[ParseError]) -> None:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    payload = {
        "cache_version": RUN_DATA_CACHE_VERSION,
        "runs": [
            {
                "folder_id": run.folder_id,
                "folder_name": run.folder_name,
                "customer_code": run.customer_code,
                "run_date": run.run_date.isoformat(),
                "carrier": run.carrier,
                "run_id": run.run_id,
                "images": [
                    {
                        "id": image.id,
                        "name": image.name,
                        "mime_type": image.mime_type,
                        "web_view_link": image.web_view_link,
                        "size": image.size,
                    }
                    for image in run.images
                ],
                "qr_info": run.qr_info,
                "qr_source": run.qr_source,
                "metadata": run.metadata,
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
    RUN_DATA_CACHE_PATH.write_text(
        json.dumps(payload, ensure_ascii=True, indent=2),
        encoding="utf-8",
    )


def _image_cache_path(run_folder_id: str, image_id: str, carrier: str) -> Path:
    safe_name = json.dumps([IMAGE_CACHE_VERSION, run_folder_id, image_id, carrier], ensure_ascii=True)
    cache_key = hashlib.sha256(safe_name.encode("utf-8")).hexdigest()
    return IMAGE_CACHE_DIR / f"{cache_key}.bin"


def clear_persisted_cache() -> None:
    if RUN_DATA_CACHE_PATH.exists():
        RUN_DATA_CACHE_PATH.unlink()
    if IMAGE_CACHE_DIR.exists():
        shutil.rmtree(IMAGE_CACHE_DIR)


def _auto_rotation_degrees(image: Image.Image, carrier: str) -> int:
    if carrier not in USB_CAMERA_AUTO_ROTATE_CARRIERS:
        return 0

    normalized = ImageOps.exif_transpose(image)
    width, height = normalized.size

    if height >= width:
        return 0

    return 90


def _transform_image_bytes(image_data: bytes, carrier: str) -> bytes:
    with Image.open(BytesIO(image_data)) as image:
        normalized = ImageOps.exif_transpose(image)
        total_rotation = _auto_rotation_degrees(image, carrier) % 360
        rotated = normalized.rotate(total_rotation, expand=True) if total_rotation else normalized.copy()
        output = BytesIO()
        save_format = image.format or "PNG"

        if save_format == "JPEG" and rotated.mode in ("RGBA", "LA"):
            rotated = rotated.convert("RGB")

        rotated.save(output, format=save_format)
        return output.getvalue()


def _load_google_drive_runs(
    drive_service: DriveService,
    root_folder_id: str,
    parse_errors: list[ParseError],
    archive_cutoff_date: date,
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
            _build_google_runs_for_folders(
                drive_service=drive_service,
                run_folders=direct_run_folders,
                carrier_name=root_folder["name"],
                carrier_folder_id=root_folder_id,
                parse_errors=parse_errors,
                minimum_date=archive_cutoff_date,
            )
        )

    for carrier_folder in carrier_folders:
        runs.extend(
            _build_google_runs_for_folders(
                drive_service=drive_service,
                run_folders=drive_service.list_child_folders(carrier_folder["id"]),
                carrier_name=carrier_folder["name"],
                carrier_folder_id=carrier_folder["id"],
                parse_errors=parse_errors,
                minimum_date=archive_cutoff_date,
            )
        )

    return runs


def _build_google_runs_for_folders(
    drive_service: DriveService,
    run_folders: list[dict],
    carrier_name: str,
    carrier_folder_id: str,
    parse_errors: list[ParseError],
    minimum_date: date,
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

        images = drive_service.list_image_files(run_folder["id"])
        qr_info, qr_source = drive_service.extract_qr_info(run_folder["id"])
        runs.append(
            RunFolder(
                folder_id=run_folder["id"],
                folder_name=folder_name,
                customer_code=parsed.customer_code,
                run_date=parsed.run_date,
                carrier=carrier_name,
                run_id=parsed.run_id,
                images=images,
                qr_info=qr_info,
                qr_source=qr_source,
                metadata={
                    "carrier_folder_id": carrier_folder_id,
                    "source": "google_drive",
                },
            )
        )

    return runs


def _load_local_archive_runs(
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
    direct_run_folders: list[Path] = []
    carrier_folders: list[Path] = []

    for child_folder in child_folders:
        try:
            parsed = parse_run_folder_name(child_folder.name)
            if archive_cutoff_date is None or parsed.run_date < archive_cutoff_date:
                direct_run_folders.append(child_folder)
        except ValueError:
            carrier_folders.append(child_folder)

    if direct_run_folders:
        runs.extend(
            _build_local_runs_for_folders(
                run_folders=direct_run_folders,
                carrier_name=archive_root.name,
                carrier_folder_id=str(archive_root),
                parse_errors=parse_errors,
                maximum_date=archive_cutoff_date,
            )
        )

    for carrier_folder in carrier_folders:
        run_folders = [path for path in carrier_folder.iterdir() if path.is_dir()]
        runs.extend(
            _build_local_runs_for_folders(
                run_folders=run_folders,
                carrier_name=carrier_folder.name,
                carrier_folder_id=str(carrier_folder),
                parse_errors=parse_errors,
                maximum_date=archive_cutoff_date,
            )
        )

    return runs


def _build_local_runs_for_folders(
    run_folders: list[Path],
    carrier_name: str,
    carrier_folder_id: str,
    parse_errors: list[ParseError],
    maximum_date: date | None,
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

        if maximum_date is not None and parsed.run_date >= maximum_date:
            continue

        files = [path for path in run_folder.iterdir() if path.is_file()]
        images = []
        for file_path in files:
            mime_type, _ = mimetypes.guess_type(file_path.name)
            suffix = file_path.suffix.lower()
            if (mime_type or "").startswith("image/") or suffix in {
                ".jpg",
                ".jpeg",
                ".png",
                ".gif",
                ".bmp",
                ".webp",
                ".tif",
                ".tiff",
            }:
                images.append(
                    {
                        "id": str(file_path),
                        "name": file_path.name,
                        "mime_type": mime_type or "application/octet-stream",
                        "size": file_path.stat().st_size,
                    }
                )

        qr_info, qr_source = _extract_local_qr_info(files)
        runs.append(
            RunFolder(
                folder_id=str(run_folder),
                folder_name=run_folder.name,
                customer_code=parsed.customer_code,
                run_date=parsed.run_date,
                carrier=carrier_name,
                run_id=parsed.run_id,
                images=[
                    ImageFile(
                        id=image["id"],
                        name=image["name"],
                        mime_type=image["mime_type"],
                        size=image["size"],
                    )
                    for image in images
                ],
                qr_info=qr_info,
                qr_source=qr_source,
                metadata={
                    "carrier_folder_id": carrier_folder_id,
                    "source": "local_archive",
                },
            )
        )

    return runs


def _extract_local_qr_info(files: list[Path]) -> tuple[str, str | None]:
    for file_path in files:
        lower_name = file_path.name.lower()
        if lower_name in {"qr.txt", "qr.json"}:
            content = file_path.read_text(encoding="utf-8", errors="replace")
            if lower_name.endswith(".json"):
                try:
                    parsed = json.loads(content)
                    return json.dumps(parsed, indent=2, ensure_ascii=True), file_path.name
                except json.JSONDecodeError:
                    return content.strip() or "No QR info found", file_path.name
            return content.strip() or "No QR info found", file_path.name

    for file_path in files:
        if "qr" in file_path.name.lower():
            return file_path.name, "filename"

    return "No QR info found", None


@st.cache_data(ttl=300, show_spinner=False)
def load_run_images(
    run_folder_id: str,
    carrier: str,
    image_ids: tuple[str, ...],
    source: str,
    _refresh_token: int,
) -> dict[str, bytes]:
    def read_or_cache_image(image_id: str) -> bytes:
        if source == "google_drive":
            cache_path = _image_cache_path(
                run_folder_id,
                image_id,
                carrier,
            )
            if cache_path.exists():
                return cache_path.read_bytes()

            drive_service = get_drive_service()
            image_data = _transform_image_bytes(
                drive_service.download_file_bytes(image_id),
                carrier,
            )
            IMAGE_CACHE_DIR.mkdir(parents=True, exist_ok=True)
            cache_path.write_bytes(image_data)
            return image_data

        return _transform_image_bytes(
            Path(image_id).read_bytes(),
            carrier,
        )

    return {
        image_id: read_or_cache_image(image_id)
        for image_id in image_ids
    }


def render_run_card(run: RunFolder, debug_mode: bool, refresh_token: int) -> None:
    has_qr_info = run.qr_info != "No QR info found"
    st.markdown(
        "\n".join(
            [item for item in [
                f"**Carrier:** {run.carrier}",
                f"**Run ID:** {run.run_id or 'N/A'}",
                f"**Folder:** `{run.folder_name}`",
                f"**QR Info:** `{run.qr_info}`" if has_qr_info and run.qr_source == "filename" else None,
                "**QR Info:**" if has_qr_info and run.qr_source != "filename" else None,
            ] if item]
        )
    )
    if has_qr_info and run.qr_source and run.qr_source != "filename":
        st.caption(f"Source: {run.qr_source}")
    if has_qr_info and run.qr_source != "filename":
        st.code(run.qr_info, language="json" if str(run.qr_source).endswith(".json") else "text")
    if debug_mode:
        st.caption(f"Folder ID: {run.folder_id}")
    image_bytes = load_run_images(
        run.folder_id,
        run.carrier,
        tuple(image.id for image in run.images),
        str(run.metadata.get("source", "google_drive")),
        refresh_token,
    )
    render_run_images(run, image_bytes, debug_mode)


def render_run_images(
    run: RunFolder,
    image_bytes: dict[str, bytes],
    debug_mode: bool,
) -> None:
    if not run.images:
        st.caption("No image files found in this run folder.")
        return

    column_count = min(3, max(1, len(run.images)))
    rows = ceil(len(run.images) / column_count)

    for row_index in range(rows):
        columns = st.columns(column_count)
        start = row_index * column_count
        end = start + column_count
        for column, image in zip(columns, run.images[start:end]):
            with column:
                try:
                    image_data = image_bytes[image.id]
                    st.image(image=image_data, caption=image.name)
                    if debug_mode:
                        st.caption("Auto rotation enabled")
                except Exception as exc:  # pragma: no cover - Streamlit rendering fallback
                    st.caption(f"{image.name} could not be rendered: {exc}")


def toggle_customer(customer_code: str) -> None:
    expanded_customers = set(st.session_state.get("expanded_customers", []))
    if customer_code in expanded_customers:
        expanded_customers.remove(customer_code)
    else:
        expanded_customers.add(customer_code)
    st.session_state["expanded_customers"] = sorted(expanded_customers)


def render_customer_header(customer_code: str, customer_runs: list[RunFolder], is_expanded: bool) -> None:
    total_images = sum(len(run.images) for run in customer_runs)
    carriers = sorted({run.carrier for run in customer_runs})
    card_class = "customer-card customer-card-open" if is_expanded else "customer-card"
    st.markdown(
        (
            f"<div class='{card_class}'>"
            f"<div class='customer-card-title'>{customer_code}</div>"
            f"<div class='customer-card-meta'>"
            f"{len(customer_runs)} runs | {total_images} images | "
            f"Carriers: {', '.join(carriers)}"
            f"</div>"
            f"</div>"
        ),
        unsafe_allow_html=True,
    )


def main() -> None:
    st.title("Sjaak vd Vijver Expedition Photo Dashboard")
    st.caption("Choose departure date and optionally filter by customer code")

    st.sidebar.header("Controls")
    debug_mode = st.sidebar.checkbox("Debug mode", value=False)
    refresh_requested = st.sidebar.button("Reload saved data", use_container_width=True)

    refresh_token = st.session_state.get("refresh_token", 0)
    if refresh_requested:
        refresh_token += 1
        st.session_state["refresh_token"] = refresh_token
        clear_persisted_cache()
        load_run_data.clear()
        load_run_images.clear()
        st.sidebar.caption("Saved cache cleared. Data will be reloaded from source.")
    elif RUN_DATA_CACHE_PATH.exists():
        st.sidebar.caption("Using saved cache for faster loading.")
    else:
        st.sidebar.caption("No saved cache yet. The next load will be stored locally.")

    try:
        with st.spinner("Scanning Google Drive and local archive folders..."):
            runs, parse_errors = load_run_data(refresh_token)
    except (RuntimeError, HttpError, OSError) as exc:
        st.error(f"Unable to load data from Google Drive: {exc}")
        st.stop()

    render_parse_errors(parse_errors, debug_mode)

    available_dates = sorted({run.run_date for run in runs})
    if not available_dates:
        st.warning("No valid run folders were found under the configured root folder.")
        return

    default_date = available_dates[-1]
    filter_col, search_col, metric_customer_col, metric_run_col, metric_image_col = st.columns(
        [2.2, 2.4, 1, 1, 1]
    )
    with filter_col:
        selected_date = st.date_input(
            "Filter by date",
            value=default_date,
            min_value=available_dates[0],
            max_value=available_dates[-1],
        )
    with search_col:
        search_term = st.text_input("Search customer code", placeholder="cust123")

    filtered_runs = [run for run in runs if run.run_date == selected_date]
    if search_term.strip():
        search_lower = search_term.strip().lower()
        filtered_runs = [
            run for run in filtered_runs if search_lower in run.customer_code.lower()
        ]

    customer_count, run_count, image_count = (
        len({run.customer_code for run in filtered_runs}),
        len(filtered_runs),
        sum(len(run.images) for run in filtered_runs),
    )
    metric_customer_col.metric("Customers", customer_count)
    metric_run_col.metric("Runs", run_count)
    metric_image_col.metric("Images", image_count)

    grouped_runs = group_runs_by_customer(filtered_runs)
    if not grouped_runs:
        st.info("No runs match the selected filters.")
        return

    for customer_code, customer_runs in grouped_runs.items():
        expanded_customers = set(st.session_state.get("expanded_customers", []))
        is_expanded = customer_code in expanded_customers
        header_col, action_col = st.columns([6, 1])
        with header_col:
            render_customer_header(customer_code, customer_runs, is_expanded)
        with action_col:
            st.write("")
            st.write("")
            st.button(
                "Collapse" if is_expanded else "Expand",
                key=f"toggle-{customer_code}",
                on_click=toggle_customer,
                args=(customer_code,),
                use_container_width=True,
            )
        if is_expanded:
            for run in customer_runs:
                with st.container():
                    st.markdown("<div class='run-shell'>", unsafe_allow_html=True)
                    st.subheader(f"{run.carrier} | {run.run_id or 'No run ID'}")
                    render_run_card(run, debug_mode, refresh_token)
                    st.markdown("</div>", unsafe_allow_html=True)

    if debug_mode:
        st.sidebar.subheader("Debug summary")
        st.sidebar.write(f"Root folder ID: `{get_setting('GOOGLE_DRIVE_ROOT_FOLDER_ID') or ''}`")
        st.sidebar.write(f"Local archive root: `{get_setting('LOCAL_ARCHIVE_ROOT') or ''}`")
        carriers_by_date: dict[date, set[str]] = defaultdict(set)
        for run in runs:
            carriers_by_date[run.run_date].add(run.carrier)
        st.sidebar.write(f"Dates found: {len(available_dates)}")
        st.sidebar.write(
            f"Carriers on selected date: {len(carriers_by_date.get(selected_date, set()))}"
        )


if __name__ == "__main__":
    main()
