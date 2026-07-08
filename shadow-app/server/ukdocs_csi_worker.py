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


def row_find_index(values, expected):
    target = normalize_key(expected)
    for index, value in enumerate(values):
        if normalize_key(value) == target:
            return index
    return None


def row_find_label_value(values, expected):
    target = normalize_key(expected)
    for index, value in enumerate(values):
        if normalize_key(value) != target:
            continue
        for next_index in range(index + 1, len(values)):
            candidate = clean_text(values[next_index])
            if candidate:
                return candidate
    return ""


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
    header = [normalize_key(cell) for cell in (rows[0] if rows else [])]
    start_index = 1 if header and any("commodity" in cell or "genus" in cell for cell in header) else 0
    for row in rows[start_index:]:
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


def parse_delimited_rows(text):
    sample = text[:4096]
    delimiters = ",;\t|"
    dialect = None
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=delimiters)
    except Exception:
        dialect = None

    if dialect is not None:
        reader = csv.reader(StringIO(text), dialect)
        rows = list(reader)
        if rows:
            return rows, dialect.delimiter

    fallback_delimiter = ","
    for candidate in [";", "\t", "|", ","]:
        lines = [line for line in text.splitlines() if line.strip()]
        if not lines:
            continue
        split_rows = [line.split(candidate) for line in lines]
        max_columns = max((len(row) for row in split_rows), default=0)
        if max_columns > 1:
            return split_rows, candidate

    reader = csv.reader(StringIO(text))
    return list(reader), fallback_delimiter


def parse_export_sheet(workbook):
    worksheet = workbook.worksheets[0]
    rows = []
    summary = {}
    column_indexes = None
    for row in worksheet.iter_rows(values_only=True):
        values = [clean_text(cell) for cell in row]
        if column_indexes is None:
            description_index = row_find_index(values, "Goods description")
            commodity_index = row_find_index(values, "Commodity code")
            quantity_index = row_find_index(values, "Quantity")
            origin_index = row_find_index(values, "oorsprong")
            if description_index is not None and commodity_index is not None and quantity_index is not None:
                column_indexes = {
                    "description": description_index,
                    "commodity_code": commodity_index,
                    "quantity": quantity_index,
                    "origin": origin_index,
                }
            continue
        if column_indexes is None:
            continue
        description = values[column_indexes["description"]] if column_indexes["description"] < len(values) else ""
        commodity_code = values[column_indexes["commodity_code"]] if column_indexes["commodity_code"] < len(values) else ""
        quantity = parse_int_like(values[column_indexes["quantity"]] if column_indexes["quantity"] < len(values) else "")
        origin_index = column_indexes.get("origin")
        origin = values[origin_index] if origin_index is not None and origin_index < len(values) else ""
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
    column_indexes = None
    for row in worksheet.iter_rows(values_only=True):
        values = [clean_text(cell) for cell in row]
        meta["date"] = meta["date"] or row_find_label_value(values, "Date :")
        meta["invoice_number"] = meta["invoice_number"] or row_find_label_value(values, "Invoice nr :")
        meta["truck"] = meta["truck"] or row_find_label_value(values, "Licence Truck :")
        meta["delivery_terms"] = meta["delivery_terms"] or row_find_label_value(values, "Delivery Terms :")
        meta["currency"] = meta["currency"] or row_find_label_value(values, "Currency of invoice")

        if column_indexes is None:
            commodity_index = row_find_index(values, "classificationType TARIC")
            description_index = row_find_index(values, "Goods description")
            origin_index = row_find_index(values, "Origin")
            quantity_index = row_find_index(values, "Quantity")
            if commodity_index is not None and description_index is not None and origin_index is not None and quantity_index is not None:
                column_indexes = {
                    "commodity_code": commodity_index,
                    "description": description_index,
                    "origin": origin_index,
                    "quantity": quantity_index,
                }
            continue
        if column_indexes is None:
            continue
        commodity_code = values[column_indexes["commodity_code"]] if column_indexes["commodity_code"] < len(values) else ""
        description = values[column_indexes["description"]] if column_indexes["description"] < len(values) else ""
        origin = values[column_indexes["origin"]] if column_indexes["origin"] < len(values) else ""
        quantity = parse_int_like(values[column_indexes["quantity"]] if column_indexes["quantity"] < len(values) else "")
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

    index = 0
    product_chunks = []
    while index < len(lines):
        line = clean_text(lines[index])
        if not re.match(r"^\d{3,4}\b", line):
            index += 1
            continue
        chunk = [line]
        lookahead = index + 1
        while lookahead < len(lines) and len(chunk) < 5:
            candidate = clean_text(lines[lookahead])
            if not candidate:
                lookahead += 1
                continue
            if re.match(r"^\d{3,4}\b", candidate):
                break
            if "text end" in normalize_key(candidate):
                break
            chunk.append(candidate)
            lookahead += 1
        product_chunks.append(" ".join(chunk))
        index = lookahead

    for chunk_text in product_chunks:
        quantity_match = re.search(r"([\d.,]+)\s+Pieces\b", chunk_text, flags=re.IGNORECASE)
        packages_match = re.search(r"\b(\d+)\s+Box\b", chunk_text, flags=re.IGNORECASE)
        quantity = parse_int_like(quantity_match.group(1)) if quantity_match else None
        packages = parse_int_like(packages_match.group(1)) if packages_match else None
        product_text = re.sub(r"^\d{3,4}\s*", "", chunk_text)
        product_text = re.sub(r"\s+\d+\s+Box\b.*$", "", product_text, flags=re.IGNORECASE)
        product_text = re.sub(r"\s+[\d.,]+\s+Pieces\b.*$", "", product_text, flags=re.IGNORECASE)
        cleaned_product = clean_text(product_text)
        if not cleaned_product:
            continue
        parsed["product_lines"].append({
            "product": cleaned_product,
            "packages": packages,
            "quantity": quantity,
        })

    if len(parsed["product_lines"]) == 1 and parsed["product_lines"][0].get("quantity") is None and parsed["total_quantity"] is not None:
        parsed["product_lines"][0]["quantity"] = parsed["total_quantity"]

    if not parsed["product_lines"] and parsed["document_state"] == "ok":
        parsed["problems"].append("No product lines extracted from temporary phyto PDF")

    return parsed


def extract_csv(path: Path):
    try:
        text = path.read_text(encoding="utf-8-sig", errors="replace")
    except Exception:
        text = path.read_text(encoding="utf-8", errors="replace")
    raw_rows, detected_delimiter = parse_delimited_rows(text)
    rows = []
    for row in raw_rows:
        cleaned = [clean_text(cell) for cell in row]
        if any(cleaned):
            rows.append("\t".join(cleaned))
    payload = {
        "content_type": "csv",
        "text": "\n".join(limit_lines(rows)),
        "line_count": len(rows),
        "delimiter": detected_delimiter,
    }
    if raw_rows:
        payload["parsed_data"] = {
            **parse_ipaffs_rows(raw_rows),
            "delimiter": detected_delimiter,
        }
    return payload


def extract_xlsx(path: Path, kind=""):
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
    normalized_kind = clean_text(kind)
    if normalized_kind == "generated_invoice" or ("invoice " in lower_name):
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


def extract_pdf(path: Path, kind=""):
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
    if clean_text(kind) == "temp_phyto":
        payload["parsed_data"] = parse_temp_phyto_pdf_text("\n".join(lines))
    return payload


def extract_file(entry):
    path = Path(entry.get("path") or "")
    suffix = path.suffix.lower()
    kind = clean_text(entry.get("kind"))
    if not path.exists():
        return {
            "kind": kind,
            "name": clean_text(entry.get("name")) or path.name,
            "mime_type": clean_text(entry.get("mime_type")),
            "content_type": "missing",
            "error": "File not found",
            "text": "",
        }

    base = {
        "kind": kind,
        "name": clean_text(entry.get("name")) or path.name,
        "mime_type": clean_text(entry.get("mime_type")),
    }
    try:
        if suffix == ".csv":
            return {**base, **extract_csv(path)}
        if suffix == ".xlsx":
            return {**base, **extract_xlsx(path, kind)}
        if suffix == ".xls":
            return {**base, **extract_xls(path)}
        if suffix == ".pdf":
            return {**base, **extract_pdf(path, kind)}
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
