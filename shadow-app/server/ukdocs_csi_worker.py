#!/usr/bin/env python3
import csv
import json
import re
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

try:
    from pypdf import PdfReader
except ImportError:
    try:
        from PyPDF2 import PdfReader
    except ImportError:
        PdfReader = None


def clean_text(value):
    return str(value or "").replace("\x00", "").strip()


def limit_lines(lines, max_lines=220):
    if len(lines) <= max_lines:
      return lines
    kept = lines[:max_lines]
    kept.append(f"... truncated, {len(lines) - max_lines} more lines")
    return kept


def normalize_key(value):
    return re.sub(r"[^a-z0-9]+", " ", clean_text(value).lower()).strip()


def parse_number(value):
    text = clean_text(value).replace(",", ".")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def parse_int_like(value):
    number = parse_number(value)
    if number is None:
        return None
    return int(round(number))


def map_ipaffs_product(genus, commodity_code):
    genus_key = normalize_key(genus)
    if "chrysanthem" in genus_key:
        return "Flowers chrysanthemums"
    if "dianthus" in genus_key:
        return "Flowers carnation"
    if "gypsoph" in genus_key:
        return "Flowers (other fresh)"
    if "rosa" in genus_key:
        return "Flowers roses"
    if "solidago" in genus_key:
        return "Flowers (other fresh)"
    code = re.sub(r"\D+", "", clean_text(commodity_code))
    if code.startswith("603140") or code.startswith("060314"):
        return "Flowers chrysanthemums"
    if code.startswith("603120") or code.startswith("060312"):
        return "Flowers carnation"
    if code.startswith("603110") or code.startswith("060311"):
        return "Flowers roses"
    if code.startswith("603150") or code.startswith("060315"):
        return "Flowers lilies"
    if code.startswith("604209"):
        return "Flowers green"
    return "Flowers (other fresh)"


def parse_ipaffs_rows(rows):
    parsed_rows = []
    summary = {}
    for row in rows[1:]:
        if not any(clean_text(cell) for cell in row):
            continue
        commodity_code = clean_text(row[0] if len(row) > 0 else "")
        genus = clean_text(row[1] if len(row) > 1 else "")
        packages = parse_int_like(row[7] if len(row) > 7 else "")
        quantity = parse_int_like(row[9] if len(row) > 9 else "")
        unit = clean_text(row[10] if len(row) > 10 else "")
        weight = parse_number(row[11] if len(row) > 11 else "")
        product = map_ipaffs_product(genus, commodity_code)
        parsed_row = {
            "product": product,
            "commodity_code": commodity_code,
            "genus": genus,
            "packages": packages,
            "quantity": quantity,
            "unit": unit,
            "weight": weight,
        }
        parsed_rows.append(parsed_row)
        summary[product] = int(summary.get(product, 0) + (quantity or 0))
    return {
        "rows": parsed_rows,
        "product_totals": [{"product": product, "quantity": quantity} for product, quantity in summary.items()],
    }


def parse_export_sheet(workbook):
    worksheet = workbook.worksheets[0]
    rows = []
    summary = {}
    in_goods = False
    for row in worksheet.iter_rows(values_only=True):
        values = [clean_text(cell) for cell in row]
        if not in_goods and len(values) > 3 and values[0] == "Goods description" and values[3] == "Quantity":
            in_goods = True
            continue
        if not in_goods:
            continue
        description = values[0] if len(values) > 0 else ""
        commodity_code = values[1] if len(values) > 1 else ""
        quantity = parse_int_like(values[3] if len(values) > 3 else "")
        origin = values[7] if len(values) > 7 else ""
        if not description or not commodity_code or quantity is None:
            continue
        parsed_row = {
            "product": description,
            "origin": origin,
            "commodity_code": commodity_code,
            "quantity": quantity,
        }
        rows.append(parsed_row)
        key = f"{description} | {origin}".strip()
        summary[key] = int(summary.get(key, 0) + quantity)
    return {
        "rows": rows,
        "product_origin_totals": [{"product_origin": key, "quantity": quantity} for key, quantity in summary.items()],
    }


def parse_invoice_sheet(workbook):
    worksheet = workbook.worksheets[0]
    rows = []
    summary = {}
    meta = {
        "invoice_number": "",
        "date": "",
        "truck": "",
        "delivery_terms": "",
        "currency": "",
    }
    in_rows = False
    for row in worksheet.iter_rows(values_only=True):
        values = [clean_text(cell) for cell in row]
        label = values[1] if len(values) > 1 else ""
        if label == "Date :":
            meta["date"] = values[2] if len(values) > 2 else ""
        elif label == "Invoice nr :":
            meta["invoice_number"] = values[2] if len(values) > 2 else ""
        elif label == "Licence Truck :":
            meta["truck"] = values[2] if len(values) > 2 else ""
        elif label == "Delivery Terms :":
            meta["delivery_terms"] = values[2] if len(values) > 2 else ""
        if not in_rows and label == "classificationType TARIC":
            in_rows = True
            continue
        if not in_rows:
            continue
        commodity_code = values[1] if len(values) > 1 else ""
        description = values[2] if len(values) > 2 else ""
        origin = values[3] if len(values) > 3 else ""
        quantity = parse_int_like(values[4] if len(values) > 4 else "")
        if not description or not commodity_code or quantity is None:
            continue
        parsed_row = {
            "product": description,
            "origin": origin,
            "commodity_code": commodity_code,
            "quantity": quantity,
        }
        rows.append(parsed_row)
        key = f"{description} | {origin}".strip()
        summary[key] = int(summary.get(key, 0) + quantity)
    return {
        "meta": meta,
        "rows": rows,
        "product_origin_totals": [{"product_origin": key, "quantity": quantity} for key, quantity in summary.items()],
    }


def find_line_value(lines, label, max_lookahead=4):
    label_key = normalize_key(label)
    for index, line in enumerate(lines):
        current_key = normalize_key(line)
        if label_key and label_key in current_key:
            tail = line.split(label, 1)[-1].strip(" :") if label in line else ""
            if tail:
                return tail
            for next_index in range(index + 1, min(len(lines), index + 1 + max_lookahead)):
                candidate = clean_text(lines[next_index])
                if candidate and not candidate.startswith("[Page]"):
                    return candidate
    return ""


def parse_temp_phyto_pdf_text(text):
    lines = [clean_text(line) for line in text.splitlines() if clean_text(line)]
    if not lines:
        return {}

    parsed = {
        "document_state": "ok",
        "pcnu_number": "",
        "destination_country": "",
        "origin_country": "",
        "consignee": "",
        "product_lines": [],
        "total_quantity": None,
        "problems": [],
    }

    lowered = normalize_key(text)
    if "nog niet geactiveerd" in lowered or "not activated" in lowered:
        parsed["document_state"] = "not_activated"
        parsed["problems"].append("Temporary phyto document appears not activated")

    pcnu_match = re.search(r"PCNU\s+([A-Z0-9]+)", text, flags=re.IGNORECASE)
    if pcnu_match:
        parsed["pcnu_number"] = clean_text(pcnu_match.group(1))
    else:
        parsed["problems"].append("PCNU number not found")

    parsed["destination_country"] = find_line_value(lines, "to Plant Protection Organization(s) of")
    if not parsed["destination_country"]:
        parsed["problems"].append("Destination country not found")

    parsed["origin_country"] = find_line_value(lines, "Place of origin")
    if not parsed["origin_country"]:
        parsed["problems"].append("Place of origin not found")

    consignee_lines = []
    consignee_started = False
    for line in lines:
        if normalize_key("Declared name and address of consignee") in normalize_key(line):
            consignee_started = True
            continue
        if consignee_started:
            if re.match(r"^\d+\b", line):
                break
            consignee_lines.append(line)
    parsed["consignee"] = ", ".join(consignee_lines[:5]).strip(", ")

    total_match = re.search(r"TOTAL\s+([\d.,]+)\s+Pieces", text, flags=re.IGNORECASE)
    if total_match:
        parsed["total_quantity"] = parse_int_like(total_match.group(1))

    for raw_line in lines:
        line = clean_text(raw_line)
        if not re.match(r"^\d{3,4}\b", line):
            continue
        quantity_match = re.search(r"([\d.,]+)\s+Pieces\b", line, flags=re.IGNORECASE)
        packages_match = re.search(r"\b(\d+)\s+Box\b", line, flags=re.IGNORECASE)
        if not quantity_match:
            continue
        quantity = parse_int_like(quantity_match.group(1))
        packages = parse_int_like(packages_match.group(1)) if packages_match else None
        product_text = re.sub(r"^\d{3,4}\s*", "", line)
        product_text = re.sub(r"\s+\d+\s+Box\b.*$", "", product_text, flags=re.IGNORECASE)
        product_text = re.sub(r"\s+[\d.,]+\s+Pieces\b.*$", "", product_text, flags=re.IGNORECASE)
        parsed["product_lines"].append({
            "product": clean_text(product_text),
            "packages": packages,
            "quantity": quantity,
        })

    if not parsed["product_lines"] and parsed["document_state"] == "ok":
        parsed["problems"].append("No product lines extracted from temporary phyto PDF")

    return parsed


def extract_csv(path: Path):
    text = path.read_text(encoding="utf-8", errors="replace")
    reader = csv.reader(StringIO(text))
    raw_rows = list(reader)
    rows = []
    for row in raw_rows:
        cleaned = [clean_text(cell) for cell in row]
        if any(cleaned):
            rows.append("\t".join(cleaned))
    payload = {
        "content_type": "csv",
        "text": "\n".join(limit_lines(rows)),
        "line_count": len(rows),
    }
    if raw_rows:
        payload["parsed_data"] = parse_ipaffs_rows(raw_rows)
    return payload


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
    payload = {
        "content_type": "xlsx",
        "sheet_names": sheet_names,
        "text": "\n".join(limit_lines(lines)),
        "line_count": len(lines),
    }
    lower_name = path.name.lower()
    if "invoice " in lower_name:
        payload["parsed_data"] = parse_invoice_sheet(workbook)
    else:
        payload["parsed_data"] = parse_export_sheet(workbook)
    return payload


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


def extract_pdf(path: Path):
    if PdfReader is None:
        return {
            "content_type": "pdf",
            "text": "",
            "line_count": 0,
            "note": "PDF text extractor not installed, keep for manual review",
        }
    reader = PdfReader(str(path))
    lines = []
    for page_index, page in enumerate(reader.pages, start=1):
        page_text = clean_text(page.extract_text() or "")
        if not page_text:
            continue
        lines.append(f"[Page] {page_index}")
        for line in page_text.splitlines():
            cleaned = clean_text(line)
            if cleaned:
                lines.append(cleaned)
    payload = {
        "content_type": "pdf",
        "text": "\n".join(limit_lines(lines)),
        "line_count": len(lines),
    }
    lower_name = path.name.lower()
    if lower_name.endswith(".pdf"):
        payload["parsed_data"] = parse_temp_phyto_pdf_text("\n".join(lines))
    return payload


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
            return {**base, **extract_pdf(path)}
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
