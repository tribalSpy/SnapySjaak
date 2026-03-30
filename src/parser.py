from __future__ import annotations

from datetime import datetime

from src.models import ParsedFolderName

def parse_run_folder_name(folder_name: str) -> ParsedFolderName:
    parts = folder_name.strip().split("_")
    date_index = next(
        (index for index, part in enumerate(parts) if len(part) == 8 and part.isdigit()),
        None,
    )

    if date_index is None or date_index == 0:
        raise ValueError(
            "Folder name must match customer_YYYYMMDD or customer_YYYYMMDD_runid"
        )

    customer_raw = "_".join(parts[:date_index]).strip()
    run_id = "_".join(parts[date_index + 1 :]).strip() or None

    if not customer_raw:
        raise ValueError(
            "Folder name must match customer_YYYYMMDD or customer_YYYYMMDD_runid"
        )

    try:
        run_date = datetime.strptime(parts[date_index], "%Y%m%d").date()
    except ValueError as exc:
        raise ValueError("Folder date is not a valid YYYYMMDD value") from exc

    return ParsedFolderName(
        customer_code=customer_raw.replace("_", "#"),
        run_date=run_date,
        run_id=run_id,
    )
