from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Any


@dataclass(frozen=True)
class ImageFile:
    id: str
    name: str
    mime_type: str
    web_view_link: str | None = None
    size: int | None = None


@dataclass(frozen=True)
class ParseError:
    folder_id: str
    folder_name: str
    carrier: str
    reason: str


@dataclass
class RunFolder:
    folder_id: str
    folder_name: str
    customer_code: str
    run_date: date
    carrier: str
    run_id: str | None
    images: list[ImageFile] = field(default_factory=list)
    qr_info: str = "No QR info found"
    qr_source: str | None = None
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class ParsedFolderName:
    customer_code: str
    run_date: date
    run_id: str | None


@dataclass(frozen=True)
class FolderScanResult:
    runs: list[RunFolder]
    errors: list[ParseError]
