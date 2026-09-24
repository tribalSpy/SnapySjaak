import argparse
import io
import json
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import extract_msg
from openpyxl import load_workbook


def clean_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_header(value):
    return re.sub(r"[^a-z0-9]+", "_", clean_text(value).lower()).strip("_")


def cell_value(value):
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return value


# The ERP "Screen" export is a legacy fixed-width report: a few decorative
# divider/title rows before the real header, which we find by looking for
# the one column every reconciliation needs -- PAV (the shared reference
# with FloraHolland's invoices).
def parse_erp(input_path: Path):
    workbook = load_workbook(filename=str(input_path), data_only=True, read_only=True)
    worksheet = workbook[workbook.sheetnames[0]]
    all_rows = [list(row) for row in worksheet.iter_rows(values_only=True)]

    header_row_index = None
    headers = []
    for idx, row in enumerate(all_rows):
        normalized = [normalize_header(value) for value in row]
        if "pav" in normalized:
            header_row_index = idx
            headers = normalized
            break

    if header_row_index is None:
        return {"rows": [], "error": "Could not find a header row containing PAV in this file"}

    rows = []
    for row in all_rows[header_row_index + 1:]:
        if not any(clean_text(value) for value in row):
            continue
        record = {}
        for col_index, header in enumerate(headers):
            if not header or col_index >= len(row):
                continue
            record[header] = cell_value(row[col_index])
        rows.append(record)

    return {"rows": rows}


TYPE_CODE_RE = re.compile(r"[._](FK|FC|HA)[._]", re.IGNORECASE)


# The XML-folder messages ("XCrossIndustryInvoice;049876.FC.2026.0181") don't
# spell out "Klokfactuur"/"Connect factuur"/"Handel aankopen" in the subject
# the way the PDF-folder messages do -- but both always carry the same
# FK/FC/HA type code from the invoice number, so that's the reliable signal.
def classify_subject(subject: str) -> str:
    normalized = (subject or "").lower()
    match = TYPE_CODE_RE.search(subject or "")
    if match:
        code = match.group(1).upper()
        return {"FK": "klokfactuur", "FC": "connect", "HA": "handel"}[code]
    if "klokfactuur" in normalized:
        return "klokfactuur"
    if "connect factuur" in normalized:
        return "connect"
    if "handel aankopen" in normalized:
        return "handel"
    return "other"


INVOICE_NUMBER_RE = re.compile(r"(\d{6}[._][A-Z]{2}[._]\d{4}[._]\d{3,4})", re.IGNORECASE)


def extract_invoice_number(subject: str):
    match = INVOICE_NUMBER_RE.search(subject or "")
    if not match:
        return ""
    return match.group(1).replace("_", ".").upper()


def find_xml_bytes(msg):
    for attachment in msg.attachments:
        name = (attachment.longFilename or attachment.shortFilename or "").lower()
        if not attachment.data:
            continue
        if name.endswith(".xml"):
            return attachment.data
        if name.endswith(".zip"):
            try:
                with zipfile.ZipFile(io.BytesIO(attachment.data)) as archive:
                    xml_names = [n for n in archive.namelist() if n.lower().endswith(".xml")]
                    if xml_names:
                        return archive.read(xml_names[0])
            except zipfile.BadZipFile:
                continue
    return None


def local_tag(tag: str) -> str:
    return tag.split("}")[-1] if "}" in tag else tag


def find_child_text(element, tag_name):
    for child in element.iter():
        if local_tag(child.tag) == tag_name:
            return (child.text or "").strip()
    return None


def find_reference(element, scheme_name):
    for ref in element.iter():
        if local_tag(ref.tag) != "ReferenceReferencedDocument":
            continue
        for child in ref.iter():
            if local_tag(child.tag) == "IssuerAssignedID" and child.attrib.get("schemeName") == scheme_name:
                return (child.text or "").strip()
    return None


def parse_number(text):
    if text is None:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def parse_invoice_xml(xml_bytes):
    root = ET.fromstring(xml_bytes)
    lines = []
    for item in root.iter():
        if local_tag(item.tag) != "InvoiceTradeLineItem":
            continue
        quantity = parse_number(find_child_text(item, "BilledQuantity"))
        description = find_child_text(item, "DescriptionText") or ""
        unit_price = parse_number(find_child_text(item, "ChargeAmount"))
        total = parse_number(find_child_text(item, "GrandTotalAmount"))
        reference_bt = find_reference(item, "BT") or ""
        lines.append({
            "description": description,
            "quantity": quantity,
            "unit_price": unit_price,
            "total": total,
            "reference_bt": reference_bt,
        })
    return lines


# Only Klokfactuur/Connect/Handel messages carry the CII XML this tool needs
# -- anything else (e.g. an AI2 Dagnota) is reported as skipped so the
# caller can show why a file in the upload wasn't used, rather than it
# silently disappearing.
def parse_veiling(input_path: Path):
    extract_dir = Path(tempfile.mkdtemp(prefix="inkoop-veiling-"))
    with zipfile.ZipFile(input_path) as archive:
        archive.extractall(extract_dir)

    invoices = []
    skipped = []
    for msg_path in sorted(extract_dir.rglob("*.msg")):
        msg = extract_msg.Message(str(msg_path))
        try:
            subject = msg.subject or ""
            kind = classify_subject(subject)
            if kind == "other":
                skipped.append({"file_name": msg_path.name, "reason": f"Unsupported message type (subject: {subject})"})
                continue

            xml_bytes = find_xml_bytes(msg)
            if not xml_bytes:
                skipped.append({"file_name": msg_path.name, "reason": "No invoice XML attachment found"})
                continue

            lines = parse_invoice_xml(xml_bytes)
            invoice_number = extract_invoice_number(subject)
            invoices.append({
                "file_name": msg_path.name,
                "type": kind,
                "invoice_number": invoice_number,
                "supplier_number": invoice_number.split(".")[0] if invoice_number else "",
                "subject": subject,
                "lines": lines,
            })
        except Exception as error:  # noqa: BLE001 -- report per-file, never abort the whole batch
            skipped.append({"file_name": msg_path.name, "reason": str(error)})
        finally:
            msg.close()

    return {"invoices": invoices, "skipped": skipped}


def main():
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(dest="command", required=True)

    erp_parser = subparsers.add_parser("parse-erp")
    erp_parser.add_argument("--input", required=True)

    veiling_parser = subparsers.add_parser("parse-veiling")
    veiling_parser.add_argument("--input", required=True)

    args = parser.parse_args()

    if args.command == "parse-erp":
        print(json.dumps(parse_erp(Path(args.input))))
    elif args.command == "parse-veiling":
        print(json.dumps(parse_veiling(Path(args.input))))


if __name__ == "__main__":
    main()
