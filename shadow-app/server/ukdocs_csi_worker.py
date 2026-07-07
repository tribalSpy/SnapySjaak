#!/usr/bin/env python3
import csv
import json
import sys
from io import StringIO
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    load_workbook = None

try:
    import xlrd
except ImportError:
    xlrd = None


def clean_text(value):
    return str(value or "").replace("\x00", "").strip()


def limit_lines(lines, max_lines=220):
    if len(lines) <= max_lines:
      return lines
    kept = lines[:max_lines]
    kept.append(f"... truncated, {len(lines) - max_lines} more lines")
    return kept


def extract_csv(path: Path):
    text = path.read_text(encoding="utf-8", errors="replace")
    reader = csv.reader(StringIO(text))
    rows = []
    for row in reader:
        cleaned = [clean_text(cell) for cell in row]
        if any(cleaned):
            rows.append("\t".join(cleaned))
    return {
        "content_type": "csv",
        "text": "\n".join(limit_lines(rows)),
        "line_count": len(rows),
    }


def extract_xlsx(path: Path):
    if load_workbook is None:
        raise RuntimeError("openpyxl is required to read .xlsx files")
    workbook = load_workbook(filename=path, data_only=True, read_only=True)
    lines = []
    sheet_names = []
    for worksheet in workbook.worksheets:
        sheet_names.append(worksheet.title)
        lines.append(f"[Sheet] {worksheet.title}")
        for row in worksheet.iter_rows(values_only=True):
            cleaned = [clean_text(cell) for cell in row]
            if any(cleaned):
                lines.append("\t".join(cleaned))
    return {
        "content_type": "xlsx",
        "sheet_names": sheet_names,
        "text": "\n".join(limit_lines(lines)),
        "line_count": len(lines),
    }


def extract_xls(path: Path):
    if xlrd is None:
        raise RuntimeError("xlrd is required to read .xls files")
    workbook = xlrd.open_workbook(path.as_posix())
    lines = []
    sheet_names = []
    for index in range(workbook.nsheets):
        sheet = workbook.sheet_by_index(index)
        sheet_names.append(sheet.name)
        lines.append(f"[Sheet] {sheet.name}")
        for row_index in range(sheet.nrows):
            row = [clean_text(sheet.cell_value(row_index, col_index)) for col_index in range(sheet.ncols)]
            if any(row):
                lines.append("\t".join(row))
    return {
        "content_type": "xls",
        "sheet_names": sheet_names,
        "text": "\n".join(limit_lines(lines)),
        "line_count": len(lines),
    }


def extract_file(entry):
    path = Path(entry.get("path") or "")
    suffix = path.suffix.lower()
    if not path.exists():
        return {
            "kind": clean_text(entry.get("kind")),
            "name": clean_text(entry.get("name")) or path.name,
            "mime_type": clean_text(entry.get("mime_type")),
            "content_type": "missing",
            "error": "File not found",
            "text": "",
        }

    base = {
        "kind": clean_text(entry.get("kind")),
        "name": clean_text(entry.get("name")) or path.name,
        "mime_type": clean_text(entry.get("mime_type")),
    }
    try:
        if suffix == ".csv":
            return {**base, **extract_csv(path)}
        if suffix == ".xlsx":
            return {**base, **extract_xlsx(path)}
        if suffix == ".xls":
            return {**base, **extract_xls(path)}
        if suffix == ".pdf":
            return {
                **base,
                "content_type": "pdf",
                "text": "",
                "line_count": 0,
                "note": "PDF kept for manual review in CSI first version",
            }
        return {
            **base,
            "content_type": suffix.lstrip(".") or "binary",
            "text": "",
            "line_count": 0,
            "note": "Unsupported file type for CSI extraction",
        }
    except Exception as error:
        return {
            **base,
            "content_type": suffix.lstrip(".") or "unknown",
            "text": "",
            "line_count": 0,
            "error": str(error),
        }


def main():
    command = sys.argv[1] if len(sys.argv) > 1 else ""
    payload = json.loads(sys.stdin.read() or "{}")
    if command != "extract":
        raise RuntimeError(f"Unknown command: {command}")
    files = payload.get("files") or []
    documents = [extract_file(entry) for entry in files]
    print(json.dumps({"documents": documents}, ensure_ascii=False))


if __name__ == "__main__":
    main()
