import argparse
import base64
import csv
import io
import json
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import extract_msg
import pdfplumber
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


ERP_DASH_DATE_RE = re.compile(r"^(\d{2})-(\d{2})-(\d{2})$")


# The "Date"/"ITE.date" columns arrive as either a real datetime (the xlsx
# export) or a "DD-MM-YY" text string (the CSV export, e.g. "13-09-26") --
# both are normalized to a plain ISO date so erp_date/matching never has to
# care which export format a given upload happened to be.
def normalize_erp_date(value):
    if hasattr(value, "isoformat"):
        return value.isoformat()
    text = clean_text(value)
    match = ERP_DASH_DATE_RE.match(text)
    if match:
        day, month, year = match.groups()
        return f"20{year}-{month}-{day}"
    return text


ERP_DATE_COLUMNS = {"date", "ite_date"}


def find_erp_header(all_rows):
    for idx, row in enumerate(all_rows):
        normalized = [normalize_header(value) for value in row]
        if "pav" in normalized:
            return normalized, all_rows[idx + 1:]
    return None, []


def load_erp_xlsx_rows(input_path: Path):
    workbook = load_workbook(filename=str(input_path), data_only=True, read_only=True)
    worksheet = workbook[workbook.sheetnames[0]]
    return [list(row) for row in worksheet.iter_rows(values_only=True)]


def load_erp_csv_rows(input_path: Path):
    # The export uses ";" as the field delimiter -- values are Dutch-locale
    # formatted numbers/decimal text, not real comma-separated data.
    with open(input_path, encoding="utf-8-sig", errors="replace", newline="") as handle:
        return list(csv.reader(handle, delimiter=";"))


# The ERP "Screen" export is a legacy fixed-width report: a few decorative
# divider/title rows before the real header, which we find by looking for
# the one column every reconciliation needs -- PAV (the shared reference
# with FloraHolland's invoices). Exported either as .xlsx or, going forward,
# as a ";"-delimited .csv -- both a real day's worth of purchases can now
# span several calendar days per file, not just one.
def parse_erp(input_path: Path):
    is_csv = str(input_path).lower().endswith(".csv")
    all_rows = load_erp_csv_rows(input_path) if is_csv else load_erp_xlsx_rows(input_path)
    headers, data_rows = find_erp_header(all_rows)

    if headers is None:
        return {"rows": [], "error": "Could not find a header row containing PAV in this file"}

    rows = []
    for row in data_rows:
        if not any(clean_text(value) for value in row):
            continue
        record = {}
        for col_index, header in enumerate(headers):
            if not header or col_index >= len(row):
                continue
            value = row[col_index]
            record[header] = normalize_erp_date(value) if header in ERP_DATE_COLUMNS else cell_value(value)
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
    if "dagnota" in normalized:
        return "ai2"
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


def find_pdf_bytes(msg):
    for attachment in msg.attachments:
        name = (attachment.longFilename or attachment.shortFilename or "").lower()
        if attachment.data and name.endswith(".pdf"):
            return attachment.data
    return None


# AI2 sends two PDFs per email ("Dagnota" -- a daily settlement summary,
# confirmed useless for reconciliation, no quantity/price/GLN -- and
# "Productnota" -- the itemized purchase list this needs). Picks the
# specific attachment by name rather than just any PDF.
def find_named_pdf_bytes(msg, name_prefix: str):
    prefix = name_prefix.lower()
    for attachment in msg.attachments:
        name = (attachment.longFilename or attachment.shortFilename or "").lower()
        if attachment.data and name.endswith(".pdf") and name.startswith(prefix):
            return attachment.data
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


# FloraHolland's own VBN product classification, scoped to the line's own
# Product block (not the unrelated TypeCode inside ExtendedTradePrice) --
# 57 = a real flower/plant product, 67 = packaging/container/logistics
# material, 128 = an administrative roll-up (e.g. Transactieheffing).
# Confirmed against real invoice XML: only 57 lines ever have an ERP
# counterpart, so this is the reliable signal for what to even attempt to
# match, rather than inferring it from quantity/reference presence alone.
def find_product_type_code(item):
    for product in item.iter():
        if local_tag(product.tag) != "Product":
            continue
        return find_child_text(product, "TypeCode") or ""
    return ""


def find_supplier_info(item):
    for party in item.iter():
        if local_tag(party.tag) != "SupplierParty":
            continue
        gln = None
        fh_number = None
        name = None
        for child in party.iter():
            if local_tag(child.tag) == "PrimaryID" and gln is None:
                gln = (child.text or "").strip()
            elif local_tag(child.tag) == "AdditionalID" and child.attrib.get("schemeAgencyName") == "FH":
                fh_number = (child.text or "").strip()
            elif local_tag(child.tag) == "Name" and name is None:
                name = (child.text or "").strip()
        return {"gln": gln or "", "fh_number": fh_number or "", "name": name or ""}
    return {"gln": "", "fh_number": "", "name": ""}


# Header-level element (one per invoice, appears once before any line item)
# identifying which of our own buying entities this invoice was billed to --
# e.g. "Juniflor Flower Export BV". Distinct from SupplierParty, which is the
# grower/seller and is scoped per line item.
def find_invoicee_info(root):
    for party in root.iter():
        if local_tag(party.tag) != "InvoiceeParty":
            continue
        gln = None
        name = None
        for child in party.iter():
            if local_tag(child.tag) == "PrimaryID" and gln is None:
                gln = (child.text or "").strip()
            elif local_tag(child.tag) == "Name" and name is None:
                name = (child.text or "").strip()
        return {"gln": gln or "", "name": name or ""}
    return {"gln": "", "name": ""}


def parse_invoice_xml(xml_bytes):
    root = ET.fromstring(xml_bytes)
    invoicee = find_invoicee_info(root)
    # Header-level issue date, used as a fallback when a line has no date of
    # its own -- real dates from the XML instead of the date manually typed
    # into the upload form, which is what a calendar/day report needs to be
    # trustworthy.
    issue_date = find_child_text(root, "IssueDateTime") or ""
    lines = []
    for item in root.iter():
        if local_tag(item.tag) != "InvoiceTradeLineItem":
            continue
        quantity = parse_number(find_child_text(item, "BilledQuantity"))
        description = find_child_text(item, "DescriptionText") or ""
        unit_price = parse_number(find_child_text(item, "ChargeAmount"))
        total = parse_number(find_child_text(item, "GrandTotalAmount"))
        reference_bt = find_reference(item, "BT") or ""
        supplier = find_supplier_info(item)
        line_date = find_child_text(item, "LineDateTime") or issue_date
        lines.append({
            "description": description,
            "quantity": quantity,
            "unit_price": unit_price,
            "total": total,
            "reference_bt": reference_bt,
            "supplier_gln": supplier["gln"],
            "supplier_fh_number": supplier["fh_number"],
            "supplier_name": supplier["name"],
            "company_name": invoicee["name"],
            "line_date": line_date,
            "product_type_code": find_product_type_code(item),
        })
    return lines


# Master data names are terse ("DUTCH GREEN CENTRE"); names on real
# invoices/AI2 statements are verbose ("Dutch Green Centre B.V.", "Arie
# Verschoor B.V. en Zn. (verkoop)") -- strip the trailing sale/purchase
# parenthetical and common Dutch corporate suffixes before the existing
# punctuation-stripping, so both sides normalize to the same key. Confirmed
# against two real AI2 growers this session.
SUPPLIER_NAME_PARENTHETICAL_RE = re.compile(r"\s*\((?:verkoop|inkoop)\)\s*$", re.IGNORECASE)
# Allows for optional whitespace between letters (e.g. "B. V.") since
# _insert_camel_boundaries can split "B.V." that way when de-gluing
# PDF-extracted text -- without it, "b\.?v\.?" alone would only match a
# tight "BV"/"B.V." with no space, missing exactly the case it needs to
# handle here.
SUPPLIER_NAME_CORP_SUFFIX_RE = re.compile(r"\b(?:b\.?\s*v\.?|n\.?\s*v\.?|en\s+zn\.?|zn\.?)\b", re.IGNORECASE)


def normalize_supplier_name(name: str) -> str:
    text = _insert_camel_boundaries(name or "")
    text = SUPPLIER_NAME_PARENTHETICAL_RE.sub("", text)
    text = SUPPLIER_NAME_CORP_SUFFIX_RE.sub("", text)
    return re.sub(r"[^a-z0-9]+", "", text.lower())


# pdfplumber glues adjacent words together with no space whenever the PDF
# itself has none between them (confirmed on real AI2 output, e.g.
# "DutchGreenCentreB.V." for what prints as "Dutch Green Centre B.V.") --
# this reconstructs word boundaries wherever a lowercase letter (or a dot)
# is immediately followed by an uppercase one, a no-op on text that already
# has real spaces.
def _insert_camel_boundaries(text):
    return re.sub(r"(?<=[a-z0-9.])(?=[A-Z])", " ", text or "")


def parse_decimal_comma(text):
    cleaned = (text or "").strip()
    if not cleaned:
        return None
    if "," in cleaned and "." in cleaned:
        cleaned = cleaned.replace(".", "").replace(",", ".")
    elif "," in cleaned:
        cleaned = cleaned.replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return None


AI2_DATE_TOKEN_RE = re.compile(r"^(\d{2})-(\d{2})$")
AI2_HEADER_TITLE_RE = re.compile(r"Product invoice:\s*\d{1,2}\s+[A-Za-z]+\s+(\d{4})", re.IGNORECASE)
AI2_KLANTNUMMER_RE = re.compile(r"Klantnummer:\s*(\S+)")
AI2_KLANTNAAM_RE = re.compile(r"Klantnaam:\s*(.+?)\s*Btwnummer:", re.IGNORECASE)
# "Nota nummer" (e.g. "90985-4835") is printed inside the PDF body, not the
# email subject -- a stable, unique-per-statement id, used as the ledger's
# invoice_number for AI2 lines.
AI2_NOTANUMMER_RE = re.compile(r"Notanummer:\s*(\S+)")

# Column boundaries observed on the real AI2 "Productnota" PDF -- stable
# across both the "Debet" and "Credit" tables (identical header layout).
# Only the columns needed for matching against the ERP are extracted; the
# size-grade/quality columns and the Emballage/Deposit/Rent/Costs/Reference
# columns are display-only and not needed here.
AI2_COL_DATUM_MAX = 45
AI2_COL_REFERENTIE_MAX = 88
AI2_COL_DESCRIPTION_MAX = 200
AI2_COL_IGNORE_MAX = 310
AI2_COL_NUMNAME_MAX = 470
AI2_COL_AANTAL_MAX = 500
AI2_COL_APE_MAX = 525
AI2_COL_TOTAL_MAX = 560
AI2_COL_PRIJS_MAX = 585
AI2_COL_BEDRAG_MAX = 615


def parse_ai2_row(row_words, year, is_credit):
    datum = referentie = vbn_desc = num_name = total_pieces = prijs = bedrag = None
    for word in sorted(row_words, key=lambda w: w["x0"]):
        x0 = word["x0"]
        text = word["text"]
        if x0 < AI2_COL_DATUM_MAX:
            datum = text
        elif x0 < AI2_COL_REFERENTIE_MAX:
            referentie = text
        elif x0 < AI2_COL_DESCRIPTION_MAX:
            vbn_desc = text
        elif x0 < AI2_COL_IGNORE_MAX:
            continue
        elif x0 < AI2_COL_NUMNAME_MAX:
            num_name = text
        elif x0 < AI2_COL_AANTAL_MAX:
            continue
        elif x0 < AI2_COL_APE_MAX:
            continue
        elif x0 < AI2_COL_TOTAL_MAX:
            total_pieces = text
        elif x0 < AI2_COL_PRIJS_MAX:
            prijs = text
        elif x0 < AI2_COL_BEDRAG_MAX:
            bedrag = text
        # else: Emb./Dep/Rent/Costs/Order/Buyer reference columns -- not
        # needed for matching, intentionally ignored.

    date_match = AI2_DATE_TOKEN_RE.match(datum or "")
    if not date_match or not year:
        return None
    day, month = date_match.groups()
    iso_date = f"{year}-{month}-{day}"

    code_desc_match = re.match(r"^(\d+)(.*)$", vbn_desc or "")
    description = _insert_camel_boundaries(code_desc_match.group(2)).strip() if code_desc_match else (vbn_desc or "")

    code_name_match = re.match(r"^(\d+)(.*)$", num_name or "")
    supplier_fh_number = code_name_match.group(1) if code_name_match else ""
    supplier_name = _insert_camel_boundaries(code_name_match.group(2)).strip() if code_name_match else (num_name or "")
    if not supplier_name:
        return None

    quantity = parse_decimal_comma(total_pieces)
    unit_price = parse_decimal_comma(prijs)
    total = parse_decimal_comma(bedrag)
    if quantity is None or unit_price is None or total is None:
        return None
    # Defensive: reject anything that doesn't arithmetically add up rather
    # than silently including a mis-parsed row as real purchase data.
    if abs(quantity * unit_price - total) > 0.05:
        return None

    return {
        "description": description or (vbn_desc or ""),
        "quantity": quantity,
        "unit_price": -unit_price if is_credit else unit_price,
        "total": -total if is_credit else total,
        "reference_bt": "",
        "supplier_gln": "",
        "supplier_fh_number": supplier_fh_number,
        "supplier_name": supplier_name,
        "line_date": iso_date,
        "product_type_code": "57",
        "referentie": referentie or "",
        "is_credit": is_credit,
    }


# AI2's "Productnota" (product statement) PDF -- itemized broker-purchase
# lines, one row per grower/product, confirmed via real data to correspond
# exactly (same pieces/price/total, to the cent) to real ERP rows once the
# placeholder-PAV bug is fixed (see matchInkoopVeilingLines). Genuine
# selectable text, parsed positionally with pdfplumber since the PDF has no
# ruling lines for pdfplumber's own table-detection to key off.
def parse_ai2_productnota(pdf_bytes):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        first_page_text = pdf.pages[0].extract_text() or ""
        title_match = AI2_HEADER_TITLE_RE.search(first_page_text)
        year = int(title_match.group(1)) if title_match else None
        klantnummer_match = AI2_KLANTNUMMER_RE.search(first_page_text)
        klantnaam_match = AI2_KLANTNAAM_RE.search(first_page_text)
        notanummer_match = AI2_NOTANUMMER_RE.search(first_page_text)

        lines = []
        skipped_rows = 0
        in_table = False
        is_credit = False
        for page in pdf.pages:
            rows = {}
            for word in page.extract_words():
                key = round(word["top"])
                rows.setdefault(key, []).append(word)
            for key in sorted(rows.keys()):
                row_words = rows[key]
                row_text = "".join(w["text"] for w in sorted(row_words, key=lambda w: w["x0"]))
                if row_text.startswith("DatumReferentie"):
                    in_table = True
                    continue
                if row_text.startswith("Credit") and "Product" in row_text:
                    is_credit = True
                    continue
                if row_text.startswith("Financieeloverzicht") or row_text.startswith("Omschrijving"):
                    in_table = False
                    continue
                if not in_table or row_text.startswith("Total"):
                    continue
                parsed = parse_ai2_row(row_words, year, is_credit)
                if parsed:
                    lines.append(parsed)
                elif any(AI2_DATE_TOKEN_RE.match(w["text"]) for w in row_words if w["x0"] < AI2_COL_DATUM_MAX):
                    skipped_rows += 1

    return {
        "klantnummer": klantnummer_match.group(1).strip() if klantnummer_match else "",
        "klantnaam": _insert_camel_boundaries(klantnaam_match.group(1).strip()).strip() if klantnaam_match else "",
        "notanummer": notanummer_match.group(1).strip() if notanummer_match else "",
        "lines": lines,
        "skipped_rows": skipped_rows,
    }


# "Kwekercod" (present on both stamgegevens exports) is FloraHolland's own
# grower reference, one leading letter (a registration-series marker, e.g.
# "a"/"f"/"w"/"r"/"v"/"t" -- confirmed real, NOT decorative: two different
# growers can share the same digits under different letters) followed by up
# to 6 digits, e.g. "a898". Zero-padded, the digits alone equal the exact
# same 6-digit number invoices carry as the grower's <AdditionalID
# schemeAgencyName="FH"> (see find_supplier_info) -- confirmed against real
# invoice data (370/376 real GLN-verified supplier lines matched exactly).
# Since the letter isn't part of that number, only the digits are kept.
FH_KWEKERCODE_DIGITS_RE = re.compile(r"[0-9]")


def derive_fh_key_from_kwekercode(kwekercode):
    digits = "".join(FH_KWEKERCODE_DIGITS_RE.findall(kwekercode or ""))
    if not digits or len(digits) > 6:
        return None
    return digits.zfill(6)


# Master data dumps ("stamgegevens") -- semicolon-delimited, ~240 columns,
# but only Code/Naam/GLN kweke/Kwekercod matter here. Growers ("kwekers")
# and direct-trade partners ("leveranciers") are separate exports with
# different coverage, so both are read and merged (leveranciers entries win
# on conflict since they're the more specific, purchase-relevant list).
#
# GLN is the primary, exact join key, but not every row has one (some
# direct-trade partners are only recorded as an internal alias with a name,
# no GLN at all) -- for those, a normalized-name index is built too, used
# only as a *suggestion* on the "supplier not linked" screen, never to
# silently auto-match, since company names can collide or vary in ways a
# GLN can't. The Kwekercod-derived FH key sits in between: unlike a name, an
# exact digit match is trustworthy enough to auto-match on -- but confirmed
# real (see derive_fh_key_from_kwekercode): stripping the letter can still
# collide two unrelated growers who happen to share the same digits under a
# different letter (624 of 11038 real codes, ~5.7%), so any key with more
# than one distinct Code behind it is dropped entirely rather than guessed.
def load_supplier_csv(path: Path, source_label: str):
    by_gln = {}
    by_name = {}
    fh_candidates = {}
    with open(path, encoding="utf-8-sig", errors="replace", newline="") as handle:
        reader = csv.reader(handle, delimiter=";")
        header = next(reader, [])
        idx = {name: i for i, name in enumerate(header)}
        code_idx = idx.get("Code")
        naam_idx = idx.get("Naam")
        gln_idx = idx.get("GLN kweke")
        kwekercode_idx = idx.get("Kwekercod")
        if code_idx is None or naam_idx is None or gln_idx is None:
            return by_gln, by_name, fh_candidates
        for row in reader:
            if len(row) <= max(code_idx, naam_idx, gln_idx):
                continue
            code = row[code_idx].strip()
            name = row[naam_idx].strip()
            if not code:
                continue
            gln = row[gln_idx].strip()
            if gln:
                by_gln[gln] = {"code": code, "name": name, "source": source_label}
            normalized_name = normalize_supplier_name(name)
            if normalized_name:
                by_name[normalized_name] = {"code": code, "name": name, "source": source_label}
            if kwekercode_idx is not None and len(row) > kwekercode_idx:
                fh_key = derive_fh_key_from_kwekercode(row[kwekercode_idx].strip())
                if fh_key:
                    fh_candidates.setdefault(fh_key, []).append({"code": code, "name": name, "source": source_label})
    return by_gln, by_name, fh_candidates


def parse_suppliers(kwekers_path: Path, leveranciers_path: Path):
    by_gln = {}
    by_name = {}
    fh_candidates = {}
    for path, label in [(kwekers_path, "kwekers"), (leveranciers_path, "leveranciers")]:
        if not path:
            continue
        gln_entries, name_entries, fh_entries = load_supplier_csv(path, label)
        by_gln.update(gln_entries)
        by_name.update(name_entries)
        for fh_key, entries in fh_entries.items():
            fh_candidates.setdefault(fh_key, []).extend(entries)
    by_fh = {
        fh_key: entries[0]
        for fh_key, entries in fh_candidates.items()
        if len({entry["code"] for entry in entries}) == 1
    }
    return {"by_gln": by_gln, "by_name": by_name, "by_fh": by_fh}


# Klokfactuur/Connect/Handel messages carry the CII XML this tool parses;
# AI2 messages carry a "Productnota" PDF instead (its "Dagnota" zip/XML is a
# settlement summary with no itemized data, confirmed useless -- see
# parse_ai2_productnota's docstring-equivalent comment above). The zip you
# upload also has a second, PDF-only .msg per Klok/Connect/Handel invoice
# (for opening the original document later) -- that's not an error, so it's
# captured into `pdfs` rather than reported as a skip.
def parse_veiling(input_path: Path):
    extract_dir = Path(tempfile.mkdtemp(prefix="inkoop-veiling-"))
    with zipfile.ZipFile(input_path) as archive:
        archive.extractall(extract_dir)

    invoices = []
    skipped = []
    pdfs = {}
    for msg_path in sorted(extract_dir.rglob("*.msg")):
        msg = extract_msg.Message(str(msg_path))
        try:
            subject = msg.subject or ""
            kind = classify_subject(subject)
            if kind == "other":
                skipped.append({"file_name": msg_path.name, "reason": f"Unsupported message type (subject: {subject})"})
                continue

            if kind == "ai2":
                pdf_bytes = find_named_pdf_bytes(msg, "productnota")
                if not pdf_bytes:
                    skipped.append({"file_name": msg_path.name, "reason": "AI2 Dagnota: daily settlement summary, not itemized purchase data"})
                    continue
                parsed = parse_ai2_productnota(pdf_bytes)
                if not parsed["lines"]:
                    skipped.append({"file_name": msg_path.name, "reason": "AI2 Productnota: could not parse any line items"})
                    continue
                invoice_number = f"AI2-{parsed['klantnummer']}-{parsed['notanummer'] or msg_path.stem}"
                pdfs[invoice_number] = base64.b64encode(pdf_bytes).decode("ascii")
                invoices.append({
                    "file_name": msg_path.name,
                    "type": kind,
                    "invoice_number": invoice_number,
                    "supplier_number": f"AI2-{parsed['klantnummer']}",
                    "subject": subject,
                    "company_name_override": parsed["klantnaam"],
                    "lines": parsed["lines"],
                })
                continue

            xml_bytes = find_xml_bytes(msg)
            if not xml_bytes:
                pdf_bytes = find_pdf_bytes(msg)
                invoice_number = extract_invoice_number(subject)
                if pdf_bytes and invoice_number:
                    pdfs[invoice_number] = base64.b64encode(pdf_bytes).decode("ascii")
                else:
                    skipped.append({"file_name": msg_path.name, "reason": "No invoice XML or PDF attachment found"})
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

    return {"invoices": invoices, "skipped": skipped, "pdfs": pdfs}


def main():
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(dest="command", required=True)

    erp_parser = subparsers.add_parser("parse-erp")
    erp_parser.add_argument("--input", required=True)

    veiling_parser = subparsers.add_parser("parse-veiling")
    veiling_parser.add_argument("--input", required=True)

    suppliers_parser = subparsers.add_parser("parse-suppliers")
    suppliers_parser.add_argument("--kwekers", required=False, default="")
    suppliers_parser.add_argument("--leveranciers", required=False, default="")

    args = parser.parse_args()

    if args.command == "parse-erp":
        print(json.dumps(parse_erp(Path(args.input))))
    elif args.command == "parse-veiling":
        print(json.dumps(parse_veiling(Path(args.input))))
    elif args.command == "parse-suppliers":
        kwekers_path = Path(args.kwekers) if args.kwekers else None
        leveranciers_path = Path(args.leveranciers) if args.leveranciers else None
        print(json.dumps(parse_suppliers(kwekers_path, leveranciers_path)))


if __name__ == "__main__":
    main()
