from __future__ import annotations

from collections import defaultdict
from math import ceil

import streamlit as st

from src.models import ParseError, RunFolder


def summarize_runs(runs: list[RunFolder]) -> tuple[int, int, int]:
    customers = {run.customer_code for run in runs}
    images = sum(len(run.images) for run in runs)
    return len(customers), len(runs), images


def group_runs_by_customer(runs: list[RunFolder]) -> dict[str, list[RunFolder]]:
    grouped: dict[str, list[RunFolder]] = defaultdict(list)
    for run in runs:
        grouped[run.customer_code].append(run)

    for customer_runs in grouped.values():
        customer_runs.sort(key=lambda run: (run.carrier.lower(), run.run_id or ""))

    return dict(sorted(grouped.items(), key=lambda item: item[0].lower()))


def render_metrics(runs: list[RunFolder]) -> None:
    customer_count, run_count, image_count = summarize_runs(runs)
    col1, col2, col3 = st.columns(3)
    col1.metric("Customers", customer_count)
    col2.metric("Runs", run_count)
    col3.metric("Images", image_count)


def render_run_images(run: RunFolder, image_bytes: dict[str, bytes]) -> None:
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
            try:
                column.image(image=image_bytes[image.id], caption=image.name)
            except Exception as exc:  # pragma: no cover - Streamlit rendering fallback
                column.caption(f"{image.name} could not be rendered: {exc}")


def render_parse_errors(errors: list[ParseError], debug_mode: bool) -> None:
    st.sidebar.subheader("Skipped folders")
    st.sidebar.caption(f"{len(errors)} malformed or unsupported folders were skipped.")

    if not errors:
        st.sidebar.caption("No skipped folders.")
        return

    for error in errors:
        st.sidebar.write(f"`{error.folder_name}` ({error.carrier})")
        if debug_mode:
            st.sidebar.caption(error.reason)
