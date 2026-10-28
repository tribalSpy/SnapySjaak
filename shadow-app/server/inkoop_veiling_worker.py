import argparse
import base64
import csv
import io
import collections
import json
import re
import shutil
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


# "20-10-26", and "20-10-2026" / "20/10/2026" / "20.10.2026" (newer exports).
ERP_DASH_DATE_RE = re.compile(r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4}|\d{2})$")


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
        return f"{year if len(year) == 4 else '20' + year}-{int(month):02d}-{int(day):02d}"
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
        # The Communications dump (Prt / Info / OrderNr, no PAV) is not a
        # Connect purchase export -- it only adds to the lots it names.
        if any("prt" in [clean_text(value).lower() for value in row] for row in all_rows[:30]):
            return {"rows": [], "error": "This looks like the Communications dump (a Prt column, no PAV) -- upload it under 'EPP list' instead: it only adds EPP / order number / GLN to the Connect lots, matched by Prt."}
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


# FK klok, FC Connect, HA handel aankoop, plus FD dienst and HV handel
# verkoop (finance-confirmed list of every FloraHolland invoice kind) -- the
# latter two carry no purchases to reconcile, but do have to go to King.
TYPE_CODE_RE = re.compile(r"[._](FK|FC|HA|FD|HV)[._]", re.IGNORECASE)


# The XML-folder messages ("XCrossIndustryInvoice;049876.FC.2026.0181") don't
# spell out "Klokfactuur"/"Connect factuur"/"Handel aankopen" in the subject
# the way the PDF-folder messages do -- but both always carry the same
# FK/FC/HA type code from the invoice number, so that's the reliable signal.
def classify_subject(subject: str) -> str:
    normalized = (subject or "").lower()
    match = TYPE_CODE_RE.search(subject or "")
    if match:
        code = match.group(1).upper()
        return {"FK": "klokfactuur", "FC": "connect", "HA": "handel", "FD": "dienst", "HV": "handel_verkoop"}[code]
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


def direct_child(element, tag_name):
    for child in element:
        if local_tag(child.tag) == tag_name:
            return child
    return None


def direct_path_text(element, *path):
    current = element
    for tag_name in path:
        if current is None:
            return ""
        current = direct_child(current, tag_name)
    return (current.text or "").strip() if current is not None else ""


def find_iv_references(item):
    refs = []
    for ref in item:
        if local_tag(ref.tag) != "ReferenceReferencedDocument":
            continue
        for child in ref:
            if local_tag(child.tag) == "IssuerAssignedID" and child.attrib.get("schemeName") == "IV":
                refs.append((child.text or "").strip())
    return refs


# Fields King's journal export needs (see buildKingJournal in index.js):
# summary lines (SetTriggerCode 0, no quantity) carry FloraHolland's own
# Prd Id (Product/ID, schemeAgency FH, e.g. 1431) and the BTW category
# (H/S/O/E); detail lines carry the marketplace (klok location / Connect /
# Handelsregeling) and the sales account (001 flowers, 002/003 plants).
def king_line_fields(item):
    product = direct_child(item, "Product")
    prd_id = ""
    if product is not None:
        for child in product:
            if local_tag(child.tag) == "ID" and child.attrib.get("schemeAgencyName") == "FH":
                prd_id = (child.text or "").strip()
                break
    iv_refs = find_iv_references(item)
    return {
        "line_id": direct_path_text(item, "ID"),
        "prd_id": prd_id,
        "vat_category": direct_path_text(item, "TotalCalculatedTax", "CategoryCode"),
        "trigger_code": direct_path_text(item, "AccountSalesAccountingAccount", "SetTriggerCode"),
        "account_id": direct_path_text(item, "AccountSalesAccountingAccount", "ID"),
        "summary_ref": iv_refs[0] if iv_refs else "",
        "marketplace_id": direct_path_text(item, "TradingTerms", "MarketPlace", "ID"),
        "marketplace_name": direct_path_text(item, "TradingTerms", "MarketPlace", "NameText"),
        "market_form_code": direct_path_text(item, "TradingTerms", "MarketFormCode"),
        # Handel Aankopen extras: the VBN product code, "Uw ref." (scheme
        # AAG), the delivery reference (scheme VN) and the supplier's note.
        "vbn_code": direct_path_text(item, "Product", "IndustryAssignedID"),
        "your_reference": find_reference(item, "AAG") or "",
        "delivery_reference": find_reference(item, "VN") or "",
        "note": direct_path_text(item, "AdditionalInformationTradeNote", "Content"),
        # More references per line: the order number (ON, Connect -- Uniware
        # shows it with 3 digits added), the brief number without its letter
        # (ACE: the last 6 digits of the Connect lot's EPP when it is numeric),
        # and the other transaction / logistics numbers.
        "order_number": find_reference(item, "ON") or "",
        "brief_base": find_reference(item, "ACE") or "",
        "ref_ss": find_reference(item, "SS") or "",
        "ref_adj": find_reference(item, "ADJ") or "",
    }


KLOK_LOCATION_RE = re.compile(r"\b(AALSMEER|NAALDWIJK|RIJNSBURG)\b\s+KLOK", re.IGNORECASE)


# Invoice-level totals for King: the BTW per category and the total incl.
# BTW. Confirmed on every sample file: per category the summary lines add up
# to BasisAmount, and the BasisAmounts plus BTW equal GrandTotalAmount.
# klok_location: an FK invoice always covers exactly one klok (all product
# lines share one MarketPlace) -- finance books its "Product aankopen" per
# location ("1015_Aalsmeer_klok").
def parse_invoice_header(xml_bytes):
    root = ET.fromstring(xml_bytes)
    invoice = None
    for element in root.iter():
        if local_tag(element.tag) == "CrossIndustryInvoice":
            invoice = element
            break
    if invoice is None:
        return {}
    vat_subtotals = []
    locations = {}
    for child in invoice:
        tag = local_tag(child.tag)
        if tag == "CategorySubtotalTax":
            vat_subtotals.append({
                "category": direct_path_text(child, "CategoryCode"),
                "rate": parse_number(direct_path_text(child, "CalculatedRate")),
                "basis": parse_number(direct_path_text(child, "BasisAmount")) or 0.0,
                "amount": parse_number(direct_path_text(child, "CalculatedAmount")) or 0.0,
            })
        elif tag == "InvoiceTradeLineItem":
            match = KLOK_LOCATION_RE.search(direct_path_text(child, "TradingTerms", "MarketPlace", "NameText"))
            if match:
                location = match.group(1).capitalize()
                locations[location] = locations.get(location, 0) + 1
    klok_location = max(locations, key=locations.get) if locations else ""
    return {
        "billing_id": direct_path_text(invoice, "BillingDocument", "IssuerAssignedID"),
        "issue_date": direct_path_text(invoice, "BillingDocument", "IssueDateTime")[:10],
        "type_code": direct_path_text(invoice, "BillingDocument", "TypeCode"),
        "currency": direct_path_text(invoice, "BillingDocument", "BillingCurrencyCode") or "EUR",
        "grand_total": parse_number(direct_path_text(invoice, "BillingMonetarySummation", "GrandTotalAmount")),
        "vat_subtotals": vat_subtotals,
        "klok_location": klok_location,
        # More than one location would mean the per-invoice rule doesn't hold
        # for this invoice -- surfaced as a warning in the King export.
        "klok_locations": sorted(locations),
        "invoicee_name": direct_path_text(invoice, "InvoiceeParty", "Name"),
        "invoicee_gln": direct_path_text(invoice, "InvoiceeParty", "PrimaryID"),
    }


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
            # Aantal (boxes) and Stuks (per box): 33 x 80 = 2640.
            "packages": parse_number(find_child_text(item, "PackageUnitQuantity")),
            "per_package": parse_number(find_child_text(item, "ProductUnitQuantity")),
            "unit_price": unit_price,
            "total": total,
            "reference_bt": reference_bt,
            "supplier_gln": supplier["gln"],
            "supplier_fh_number": supplier["fh_number"],
            "supplier_name": supplier["name"],
            "company_name": invoicee["name"],
            "line_date": line_date,
            "product_type_code": find_product_type_code(item),
            **king_line_fields(item),
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


AI2_NUMBER_RE = re.compile(r"-?[0-9][0-9.]*(,[0-9]+)?")


def parse_ai2_row(row_words, year, is_credit):
    # Lists, not single values -- the tighter x_tolerance above means a
    # column can now legitimately contain more than one pdfplumber word
    # (that's the whole point, for Description/Num.Name), so every column
    # collects every word inside its x-range and joins them with a single
    # space at the end, rather than the last word silently overwriting the
    # ones before it.
    columns = {"datum": [], "referentie": [], "vbn_desc": [], "num_name": [], "total_pieces": [], "prijs": [], "bedrag": []}
    for word in sorted(row_words, key=lambda w: w["x0"]):
        x0 = word["x0"]
        text = word["text"]
        if x0 < AI2_COL_DATUM_MAX:
            columns["datum"].append(text)
        elif x0 < AI2_COL_REFERENTIE_MAX:
            columns["referentie"].append(text)
        elif x0 < AI2_COL_DESCRIPTION_MAX:
            columns["vbn_desc"].append(text)
        elif x0 < AI2_COL_IGNORE_MAX:
            continue
        elif x0 < AI2_COL_NUMNAME_MAX:
            columns["num_name"].append(text)
        elif x0 < AI2_COL_AANTAL_MAX:
            continue
        elif x0 < AI2_COL_APE_MAX:
            continue
        elif x0 < AI2_COL_TOTAL_MAX:
            columns["total_pieces"].append(text)
        elif x0 < AI2_COL_PRIJS_MAX:
            columns["prijs"].append(text)
        elif x0 < AI2_COL_BEDRAG_MAX:
            columns["bedrag"].append(text)
        # else: Emb./Dep/Rent/Costs/Order/Buyer reference columns -- not
        # needed for matching, intentionally ignored.
    # Aantal / ApE / Total / Prijs / Bedrag are right-aligned numbers: a wide
    # price ("14,500") starts left of the Prijs boundary and fell into the
    # Total column, so the row didn't add up and was silently dropped (Mos
    # tillandsia, EUR 14,50). After the supplier name they always come in
    # this order -- take them by position in that run, not by x.
    numbers = [word["text"] for word in sorted(row_words, key=lambda w: w["x0"])
               if word["x0"] >= AI2_COL_NUMNAME_MAX - 5 and AI2_NUMBER_RE.fullmatch(word["text"])]
    if len(numbers) >= 5:
        columns["total_pieces"], columns["prijs"], columns["bedrag"] = [numbers[2]], [numbers[3]], [numbers[4]]
    datum = " ".join(columns["datum"]) or None
    referentie = " ".join(columns["referentie"]) or None
    vbn_desc = " ".join(columns["vbn_desc"]) or None
    num_name = " ".join(columns["num_name"]) or None
    total_pieces = " ".join(columns["total_pieces"]) or None
    prijs = " ".join(columns["prijs"]) or None
    bedrag = " ".join(columns["bedrag"]) or None

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
# The AI2 Dagnota XML (zipped in the same email as the Productnota PDF) is
# Florecom CII like FloraHolland's invoices, so parse_invoice_header reads its
# total and BTW as is. For King, its detail lines (trigger 1) are summed into
# the kingkoppel keys: products (type 57) -> AI2 per BTW category, emballage
# (type 67) -> AI2_PC huur / AI2_PE eenmalig / AI2_CT statiegeld (BTW 0).
# Same shape as an XML invoice's summary lines (trigger 0), so the King
# builder books them the same way. XML BTW category Z (0%) is kingkoppel "O".
def ai2_king_summary(xml_bytes):
    root = ET.fromstring(xml_bytes)
    totals = {}
    for item in root.iter():
        if local_tag(item.tag) != "InvoiceTradeLineItem":
            continue
        if direct_path_text(item, "AccountSalesAccountingAccount", "SetTriggerCode") != "1":
            continue
        type_code = direct_path_text(item, "Product", "TypeCode")
        description = direct_path_text(item, "Product", "DescriptionText").lower()
        vat = direct_path_text(item, "TotalCalculatedTax", "CategoryCode").upper() or "O"
        vat = "O" if vat in ("Z", "E") else vat
        amount = None
        for tag in ("ActualBillingMonetarySummation", "SpecifiedBilledDelivery"):
            node = direct_child(item, tag)
            if node is None:
                continue
            for child in node.iter():
                if local_tag(child.tag) in ("GrandTotalAmount", "LineTotalAmount", "ChargeAmount") and child.text:
                    amount = float(child.text)
                    break
            if amount is not None:
                break
        if amount is None:
            continue
        if type_code == "57":
            key = "AI2"
        elif "huur" in description:
            key, vat = "AI2_PC", vat if vat != "O" else "H"
        elif "eenmalig" in description:
            key, vat = "AI2_PE", vat if vat != "O" else "S"
        elif vat == "O" or "statiegeld" in description:
            key, vat = "AI2_CT", "O"
        else:
            key = "AI2_PC" if vat == "H" else "AI2_PE"
        totals[(key, vat)] = totals.get((key, vat), 0.0) + amount
    labels = {"AI2": "Product aankopen AI2", "AI2_CT": "Emballage statiegeld AI2", "AI2_PC": "Emballage huur AI2", "AI2_PE": "Emballage eenmalig AI2"}
    return [
        {
            "trigger_code": "0",
            "prd_id": key,
            "vat_category": vat,
            "total": round(amount, 2),
            "description": labels[key],
            "product_type_code": "57" if key == "AI2" else "67",
        }
        for (key, vat), amount in sorted(totals.items())
        if round(amount, 2) != 0
    ]


AI2_SUMMARY_ROWS = [
    # (text the row starts with, summary line description, BTW, kingkoppel key)
    ("Artikel totaal L", "Product aankopen AI2 (BTW laag)", "S", "AI2"),
    ("Artikel totaal H", "Product aankopen AI2 (BTW hoog)", "H", "AI2"),
    ("Emballage meermalig statiegeld totaal", "Emballage meermalig statiegeld", "O", "AI2_CT"),
    ("Emballage meermalig huur totaal", "Emballage meermalig huur", "H", "AI2_PC"),
    ("Emballage eenmalig totaal", "Emballage eenmalig", "S", "AI2_PE"),
]


def ai2_financial_summary(pdf, product_lines, year):
    """The Productnota's "Financieel overzicht veilingincasso": article
    totals per BTW and the emballage (statiegeld, huur, eenmalig) as summary
    lines without quantity -- the same shape as a FloraHolland invoice's
    "Product aankopen" / "Emballage" lines, so the Overview shows products
    AND emballage for AI2 too (it had no emballage at all). Also checks the
    product lines add up to the article totals."""
    text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    section = text.split("Financieel overzicht", 1)[1] if "Financieel overzicht" in text else ""
    summary = []
    article_total = None
    for line in section.splitlines():
        line = line.strip()
        for start, description, vat, king_key in AI2_SUMMARY_ROWS:
            # The PDF text glues words ("Artikeltotaal L", "Emballage-
            # meermaligstatiegeldtotaal"): compare without spaces.
            if not line.replace(" ", "").lower().startswith(start.replace(" ", "").lower()):
                continue
            amounts = [parse_decimal_comma(value) for value in re.findall(r"-?[0-9][0-9.]*,[0-9]{2}", line)]
            if not amounts or amounts[-1] is None:
                continue
            saldo = amounts[-1]  # Te incasseren - Te ontvangen
            if start.startswith("Artikel"):
                article_total = (article_total or 0.0) + saldo
            if round(saldo, 2) != 0:
                summary.append({
                    "description": description,
                    "quantity": None,
                    "unit_price": None,
                    "total": round(saldo, 2),
                    "reference_bt": "",
                    "supplier_gln": "",
                    "supplier_fh_number": "",
                    "supplier_name": "Ai2",
                    "line_date": product_lines[0]["line_date"] if product_lines else None,
                    "product_type_code": "57" if start.startswith("Artikel") else "67",
                    "vat_category": vat,
                    "prd_id": king_key,
                    "trigger_code": "0",
                    "is_credit": saldo < 0,
                })
            break
    lines_total = round(sum(line["total"] for line in product_lines), 2)
    check = {"lines_total": lines_total, "article_total": round(article_total, 2) if article_total is not None else None}
    return summary, check


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
            # Default x_tolerance (3) merges visually-distinct words in this
            # PDF's narrow columns into one glued token whenever the real
            # gap between them is smaller than that threshold (confirmed
            # real against a live AI2 PDF: "2188Corylusavcontorta100cmx40"
            # for what prints as "2188 Corylus av contorta 100cm x40") --
            # 1.0 was verified to split every such case correctly across a
            # real multi-page invoice with no over-splitting (checked every
            # resulting word for suspicious single-character fragments).
            for word in page.extract_words(x_tolerance=1.0):
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

        summary_lines, check = ai2_financial_summary(pdf, lines, year)

    return {
        "summary_lines": summary_lines,
        "check": check,
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
    # "Bedrijf" per code (VEILIN, FAC, INKOOP, VDV, LTD ...): INKOOP / VDV
    # codes are internal -- no purchase to match, not in the Connect value.
    bedrijf_by_code = {}
    with open(path, encoding="utf-8-sig", errors="replace", newline="") as handle:
        reader = csv.reader(handle, delimiter=";")
        header = next(reader, [])
        idx = {name: i for i, name in enumerate(header)}
        code_idx = idx.get("Code")
        naam_idx = idx.get("Naam")
        gln_idx = idx.get("GLN kweke")
        kwekercode_idx = idx.get("Kwekercod")
        bedrijf_idx = next((i for i, name in enumerate(header) if name.strip().lower() == "bedrijf"), None)
        if code_idx is None or naam_idx is None or gln_idx is None:
            return by_gln, by_name, fh_candidates, bedrijf_by_code
        for row in reader:
            if len(row) <= max(code_idx, naam_idx, gln_idx):
                continue
            code = row[code_idx].strip()
            name = row[naam_idx].strip()
            if not code:
                continue
            if bedrijf_idx is not None and len(row) > bedrijf_idx and row[bedrijf_idx].strip():
                bedrijf_by_code[code.upper()] = row[bedrijf_idx].strip().upper()
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
    return by_gln, by_name, fh_candidates, bedrijf_by_code


def parse_suppliers(kwekers_path: Path, leveranciers_path: Path):
    by_gln = {}
    by_name = {}
    fh_candidates = {}
    bedrijf = {}
    for path, label in [(kwekers_path, "kwekers"), (leveranciers_path, "leveranciers")]:
        if not path:
            continue
        gln_entries, name_entries, fh_entries, bedrijf_entries = load_supplier_csv(path, label)
        by_gln.update(gln_entries)
        by_name.update(name_entries)
        bedrijf.update(bedrijf_entries)
        for fh_key, entries in fh_entries.items():
            fh_candidates.setdefault(fh_key, []).extend(entries)
    by_fh = {
        fh_key: entries[0]
        for fh_key, entries in fh_candidates.items()
        if len({entry["code"] for entry in entries}) == 1
    }
    return {"by_gln": by_gln, "by_name": by_name, "by_fh": by_fh, "bedrijf": bedrijf}


def find_header_row_by_column(all_rows, column_name):
    target = normalize_header(column_name)
    for idx, row in enumerate(all_rows):
        normalized = [normalize_header(value) for value in row]
        if target in normalized:
            return normalized, all_rows[idx + 1:]
    return None, []


# "klant gegevens" (customer master data) is the same fixed-width "Screen"
# export shape as the ERP/supplier dumps, keyed by "Vkn" -- the customer
# code that appears verbatim as the destination on a dispatch dump's
# "Dispatched to : <code>" rows (see parse_dispatch_dump), confirmed real by
# cross-referencing both files. "Tgrp B"/"Tgrp P" (bloemen/planten trade
# group) both being "119" is FloraHolland's own signal for a "retour"
# (return) account -- confirmed real: every one of the 12 customers
# carrying it is literally named "retour bloemen/planten ...".
KLANT_RETOUR_TGRP = "119"


def parse_klant_gegevens(input_path: Path):
    is_csv = str(input_path).lower().endswith(".csv")
    all_rows = load_erp_csv_rows(input_path) if is_csv else load_erp_xlsx_rows(input_path)
    headers, data_rows = find_header_row_by_column(all_rows, "Vkn")
    if headers is None:
        return {"by_code": {}, "error": "Could not find a header row containing Vkn in this file"}

    idx = {name: i for i, name in enumerate(headers)}

    def get(row, key):
        col = idx.get(key)
        if col is None or col >= len(row):
            return ""
        return clean_text(row[col])

    by_code = {}
    for row in data_rows:
        if not any(clean_text(value) for value in row):
            continue
        code = get(row, "vkn")
        if not code:
            continue
        tgrp_b = get(row, "tgrp_b")
        tgrp_p = get(row, "tgrp_p")
        by_code[code] = {
            "name": get(row, "naam"),
            "land": get(row, "lan"),
            "gebied": get(row, "gebi"),
            "tgrp_b": tgrp_b,
            "tgrp_p": tgrp_p,
            "type": get(row, "type"),
            "bw": get(row, "bw").upper(),
            "is_retour": tgrp_b == KLANT_RETOUR_TGRP or tgrp_p == KLANT_RETOUR_TGRP,
        }
    return {"by_code": by_code}


# "kwekers met betaalwijze" / "leveranciers met betaalwijze" -- the same
# fixed-width "Screen" export shape as klant gegevens, keyed by "Vkn" (the
# same internal code as the stamgegevens' Code column and the ERP's Suppl.),
# complementing the stamgegevens (which carry the GLN/Kwekercod join keys)
# with "BW" (betaalwijze): A=AI2, B=bunches, D=direct, F=FOB lokale valuta,
# M=mix klant, N=niet berekenen, T=LTD direct, U=Uraltorg direct, V=via
# veiling. Confirmed real: every BW=B code is an internal bunch-production
# account ("Rose Mix 1", "Bunches 10 Kraft", ...), never an auction supplier.
def parse_betaalwijze(paths):
    by_code = {}
    for path, label in paths:
        if not path:
            continue
        all_rows = load_erp_csv_rows(path)
        headers, data_rows = find_header_row_by_column(all_rows, "Vkn")
        if headers is None:
            return {"by_code": {}, "error": f"Could not find a header row containing Vkn in the {label} file"}
        idx = {name: i for i, name in enumerate(headers)}

        def get(row, key):
            col = idx.get(key)
            if col is None or col >= len(row):
                return ""
            return clean_text(row[col])

        for row in data_rows:
            code = get(row, "vkn").upper()
            if not code:
                continue
            by_code[code] = {
                "name": get(row, "naam"),
                "bw": get(row, "bw").upper(),
                "type": get(row, "type"),
                "rel_nr": get(row, "relnr"),
                "h_bedr": get(row, "h_bedr"),
                "land": get(row, "lan"),
                "source": label,
            }
    return {"by_code": by_code}


DISPATCHED_TO_RE = re.compile(r"^Dispatched to\s*:\s*(.+)$", re.IGNORECASE)


# The ERP "Screen" export can now include, immediately after each purchase
# lot's own row, one or more extra rows recording where that lot's stock
# actually went -- Description reads "Dispatched to : <customer code>",
# every other purchase-only field left blank, only Deb.no./Inv.no./Pieces
# filled in. Confirmed real: summing a lot's dispatched pieces recovers its
# own Pieces value (with a small, expected shortfall for stock not yet
# fully distributed at export time); a lot with no dispatch rows at all is
# always an Accessoires (T=A) packaging/material row, never tracked to an
# individual customer. Paired purely by file order -- a dispatch row always
# immediately follows the lot row it belongs to.
def parse_dispatch_dump(input_path: Path):
    is_csv = str(input_path).lower().endswith(".csv")
    all_rows = load_erp_csv_rows(input_path) if is_csv else load_erp_xlsx_rows(input_path)
    headers, data_rows = find_erp_header(all_rows)
    if headers is None:
        return {"lots": [], "error": "Could not find a header row containing PAV in this file"}

    idx = {name: i for i, name in enumerate(headers)}

    def get(row, key):
        col = idx.get(key)
        if col is None or col >= len(row):
            return ""
        return clean_text(row[col])

    # A large enough export can legitimately repeat the same lot number --
    # confirmed real: a lot still open across several of the export's
    # covered days re-appears with a fresh batch of dispatch rows each time.
    # Merging every occurrence's dispatches into one record (instead of
    # keeping only the last) is what keeps this correct -- an earlier
    # occurrence's real destinations must never be silently dropped just
    # because the lot recurs later in the same file.
    lots_by_number = {}
    lots = []
    current = None
    for row in data_rows:
        if not any(clean_text(value) for value in row):
            continue
        description = get(row, "description")
        dispatch_match = DISPATCHED_TO_RE.match(description)
        if dispatch_match:
            if current is not None:
                current["dispatches"].append({
                    "code": dispatch_match.group(1).strip(),
                    "pieces": parse_number(get(row, "pieces")),
                    "deb_no": get(row, "deb_no"),
                    "inv_no": get(row, "inv_no"),
                })
            continue
        lot = get(row, "lot")
        date_text = get(row, "date")
        if not lot or not date_text:
            current = None
            continue
        existing = lots_by_number.get(lot)
        if existing is not None:
            existing.update({
                "pav": get(row, "pav"),
                "date": normalize_erp_date(date_text),
                "suppl": get(row, "suppl"),
                "description": description,
                "pieces": parse_number(get(row, "pieces")),
                "price": parse_number(get(row, "price")),
                "t_price": parse_number(get(row, "t_price")),
            })
            current = existing
            continue
        current = {
            "lot": lot,
            "pav": get(row, "pav"),
            "date": normalize_erp_date(date_text),
            "suppl": get(row, "suppl"),
            "description": description,
            "pieces": parse_number(get(row, "pieces")),
            "price": parse_number(get(row, "price")),
            "t_price": parse_number(get(row, "t_price")),
            "dispatches": [],
        }
        lots_by_number[lot] = current
        lots.append(current)
    return {"lots": lots}


# Klokfactuur/Connect/Handel messages carry the CII XML this tool parses;
# AI2 messages carry a "Productnota" PDF instead (its "Dagnota" zip/XML is a
# settlement summary with no itemized data, confirmed useless -- see
# parse_ai2_productnota's docstring-equivalent comment above). The zip you
# upload also has a second, PDF-only .msg per Klok/Connect/Handel invoice
# (for opening the original document later) -- that's not an error, so it's
# captured into `pdfs` rather than reported as a skip.
def parse_veiling(input_path: Path):
    # Unpacked into its own temp folder, removed again when done (it used
    # to stay behind for every upload and filled the server's /tmp).
    extract_dir = Path(tempfile.mkdtemp(prefix="inkoop-veiling-"))
    try:
        return _parse_veiling_from(input_path, extract_dir)
    finally:
        shutil.rmtree(extract_dir, ignore_errors=True)


def _parse_veiling_from(input_path: Path, extract_dir: Path):
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
                check = parsed.get("check") or {}
                if check.get("article_total") is not None and abs(check["article_total"] - check["lines_total"]) > 0.01:
                    skipped.append({"file_name": msg_path.name, "reason": f"{invoice_number}: product lines add up to {check['lines_total']:.2f}, the nota says {check['article_total']:.2f} -- {parsed.get('skipped_rows', 0)} row(s) could not be read, check this nota"})
                pdfs[invoice_number] = base64.b64encode(pdf_bytes).decode("ascii")
                # King export: total/BTW and the booking lines from the
                # Dagnota XML in the same email (absent = not exportable).
                ai2_header = None
                ai2_xml = find_xml_bytes(msg)
                if ai2_xml:
                    try:
                        ai2_header = parse_invoice_header(ai2_xml)
                        ai2_header["king_summary"] = ai2_king_summary(ai2_xml)
                    except Exception:  # noqa: BLE001 -- the compare still works without it
                        ai2_header = None
                invoices.append({
                    "file_name": msg_path.name,
                    "type": kind,
                    "invoice_number": invoice_number,
                    "supplier_number": f"AI2-{parsed['klantnummer']}",
                    "subject": subject,
                    "company_name_override": parsed["klantnaam"],
                    "lines": parsed["lines"] + parsed.get("summary_lines", []),
                    "header": ai2_header,
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
            header = parse_invoice_header(xml_bytes)
            # The invoice's own number (in its XML) first: a forward or a
            # correction email can name another invoice in its subject, and
            # then got filed under that number ("1057 (2)").
            invoice_number = extract_invoice_number(header.get("billing_id") or "") or extract_invoice_number(subject)
            # The same email can carry the invoice PDF next to its XML --
            # kept too, or the invoice link stays "not found".
            pdf_bytes = find_pdf_bytes(msg)
            if pdf_bytes and invoice_number:
                pdfs.setdefault(invoice_number, base64.b64encode(pdf_bytes).decode("ascii"))
            invoices.append({
                "file_name": msg_path.name,
                "type": kind,
                "invoice_number": invoice_number,
                "supplier_number": invoice_number.split(".")[0] if invoice_number else "",
                "subject": subject,
                "lines": lines,
                "header": header,
            })
        except Exception as error:  # noqa: BLE001 -- report per-file, never abort the whole batch
            skipped.append({"file_name": msg_path.name, "reason": str(error)})
        finally:
            msg.close()

    # Loose files next to the emails: invoice XMLs (also zipped, like the
    # ".xml.zip" FloraHolland attaches) and invoice PDFs, named by their
    # invoice number. Same result as the email they'd otherwise come in.
    for nested_zip in sorted(extract_dir.rglob("*.zip")):
        try:
            with zipfile.ZipFile(nested_zip) as archive:
                archive.extractall(nested_zip.parent / f"{nested_zip.stem}_unzipped")
        except zipfile.BadZipFile:
            skipped.append({"file_name": nested_zip.name, "reason": "Not a readable zip file"})
    seen_xml_invoices = {invoice["invoice_number"] for invoice in invoices if invoice.get("invoice_number")}
    for xml_path in sorted(extract_dir.rglob("*.xml")):
        try:
            xml_bytes = xml_path.read_bytes()
            header = parse_invoice_header(xml_bytes)
            invoice_number = extract_invoice_number(header.get("billing_id") or "") or extract_invoice_number(xml_path.name)
            kind = classify_subject(invoice_number) if invoice_number else "other"
            if kind == "other":
                skipped.append({"file_name": xml_path.name, "reason": "Not a FloraHolland invoice XML (no FK/FC/HA/FD/HV invoice number)"})
                continue
            if invoice_number in seen_xml_invoices:
                continue  # already read from its email in this upload
            seen_xml_invoices.add(invoice_number)
            invoices.append({
                "file_name": xml_path.name,
                "type": kind,
                "invoice_number": invoice_number,
                "supplier_number": invoice_number.split(".")[0],
                "subject": xml_path.name,
                "lines": parse_invoice_xml(xml_bytes),
                "header": header,
            })
        except Exception as error:  # noqa: BLE001 -- report per file
            skipped.append({"file_name": xml_path.name, "reason": str(error)})
    for pdf_path in sorted(extract_dir.rglob("*.pdf")):
        invoice_number = extract_invoice_number(pdf_path.name)
        if invoice_number:
            pdfs.setdefault(invoice_number, base64.b64encode(pdf_path.read_bytes()).decode("ascii"))
        else:
            skipped.append({"file_name": pdf_path.name, "reason": "PDF name has no invoice number (e.g. 049876.FK.2026.0458.pdf)"})

    return {"invoices": drop_duplicate_invoices(invoices, skipped), "skipped": skipped, "pdfs": pdfs}


def invoice_line_core(line):
    """What makes an invoice line a purchase: brief number, pieces, price,
    total -- not its date, file, description or extra XML fields (a resent
    invoice, or one read from another source, can differ there)."""
    def number(value):
        try:
            return round(float(value), 4)
        except (TypeError, ValueError):
            return None
    return (
        str(line.get("reference_bt") or "").strip().upper(),
        number(line.get("quantity")),
        number(line.get("unit_price")),
        number(line.get("total")),
    )


def drop_duplicate_invoices(invoices, skipped):
    """The same invoice sent twice (e.g. one AI2 Productnota in two emails, a
    resend) gives the same invoice number with the same purchases -- read it
    once; the database refused the double rows ("ON CONFLICT DO UPDATE
    command cannot affect row a second time"), and a kept copy made its lines
    "only in invoice" (the original took the lots). Compared by the lines'
    core (see invoice_line_core): the same lines, or all of them already in
    the other copy -> one invoice, the one with the most lines. Two invoices
    with the same number and really different purchases are both kept (the
    second gets " (2)"), never dropped."""
    kept = []
    by_number = {}  # number -> [(core lines Counter, index in kept)]
    for invoice in invoices:
        number = invoice.get("invoice_number") or ""
        if not number:
            kept.append(invoice)
            continue
        core = collections.Counter(invoice_line_core(line) for line in invoice.get("lines") or [])
        known = by_number.setdefault(number, [])
        duplicate = None
        for index, (other_core, kept_index) in enumerate(known):
            if not core - other_core:
                duplicate = ("skip", kept_index)  # nothing this copy adds
                break
            if not other_core - core:
                duplicate = ("replace", kept_index, index)  # this copy has it all and more
                break
        if duplicate and duplicate[0] == "skip":
            # The same purchases: keep the copy with the most descriptions.
            described = lambda candidate: sum(1 for line in candidate.get("lines") or [] if str(line.get("description") or "").strip())
            kept_index = duplicate[1]
            if not other_core - core and described(invoice) > described(kept[kept_index]):
                kept[kept_index] = {**invoice, "invoice_number": kept[kept_index]["invoice_number"]}
            skipped.append({"file_name": invoice.get("file_name", ""), "reason": f"Same invoice {number} as {kept[duplicate[1]].get('file_name', '')} (sent again) -- read once"})
            continue
        if duplicate:
            _, kept_index, known_index = duplicate
            skipped.append({"file_name": kept[kept_index].get("file_name", ""), "reason": f"Invoice {number} came again with more lines in {invoice.get('file_name', '')} -- that one read"})
            kept[kept_index] = {**invoice, "invoice_number": kept[kept_index]["invoice_number"]}
            known[known_index] = (core, kept_index)
            continue
        if known:
            copy = len(known) + 1
            kept.append({**invoice, "invoice_number": f"{number} ({copy})"})
            skipped.append({"file_name": invoice.get("file_name", ""), "reason": f"Invoice number {number} appears again with other purchases -- kept as {number} ({copy}), check it"})
        else:
            kept.append(invoice)
        known.append((core, len(kept) - 1))
    return kept


# The Communications dump with the night action's columns (Prt, EPP, Info)
# as the "lot;epp;info" lines the EPP list upload reads. Matched by Prt only.
def parse_lot_table(input_path: Path):
    is_csv = str(input_path).lower().endswith(".csv")
    all_rows = load_erp_csv_rows(input_path) if is_csv else load_erp_xlsx_rows(input_path)
    for index, row in enumerate(all_rows[:40]):
        headers = [clean_text(value).lower() for value in row]
        if "prt" not in headers:
            continue
        prt = headers.index("prt")
        epp = headers.index("epp") if "epp" in headers else None
        # The night action's popup lines: "Popup" (or an "Info" column after
        # Prt -- the dump's own Info, before Prt, is the brief number).
        info = headers.index("popup") if "popup" in headers else (headers.index("info", prt + 1) if "info" in headers[prt + 1:] else None)
        lines = ["lot;epp;info"]
        for data in all_rows[index + 1:]:
            lot = clean_text(data[prt] if prt < len(data) else "")
            lot = lot[:-2] if lot.endswith(".0") else lot
            if not lot.isdigit():
                continue
            value = clean_text(data[epp]) if epp is not None and epp < len(data) else ""
            value = value[:-2] if value.endswith(".0") else value
            text = clean_text(data[info]) if info is not None and info < len(data) else ""
            lines.append(f"{lot};{value};{text.replace(';', ',')}")
        return {"csv": "\n".join(lines), "rows": len(lines) - 1}
    return {"csv": "", "rows": 0, "error": "No Prt column found -- is this the Communications dump?"}


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

    klant_parser = subparsers.add_parser("parse-klant-gegevens")
    klant_parser.add_argument("--input", required=True)

    dispatch_parser = subparsers.add_parser("parse-dispatch-dump")
    dispatch_parser.add_argument("--input", required=True)

    lot_table_parser = subparsers.add_parser("parse-lot-table")
    lot_table_parser.add_argument("--input", required=True)

    betaalwijze_parser = subparsers.add_parser("parse-betaalwijze")
    betaalwijze_parser.add_argument("--kwekers", required=False, default="")
    betaalwijze_parser.add_argument("--leveranciers", required=False, default="")

    args = parser.parse_args()

    if args.command == "parse-erp":
        print(json.dumps(parse_erp(Path(args.input))))
    elif args.command == "parse-veiling":
        print(json.dumps(parse_veiling(Path(args.input))))
    elif args.command == "parse-suppliers":
        kwekers_path = Path(args.kwekers) if args.kwekers else None
        leveranciers_path = Path(args.leveranciers) if args.leveranciers else None
        print(json.dumps(parse_suppliers(kwekers_path, leveranciers_path)))
    elif args.command == "parse-klant-gegevens":
        print(json.dumps(parse_klant_gegevens(Path(args.input))))
    elif args.command == "parse-dispatch-dump":
        print(json.dumps(parse_dispatch_dump(Path(args.input))))
    elif args.command == "parse-lot-table":
        print(json.dumps(parse_lot_table(Path(args.input))))
    elif args.command == "parse-betaalwijze":
        # Kwekers first so the leveranciers entry wins on a shared code, same
        # precedence as parse_suppliers.
        print(json.dumps(parse_betaalwijze([
            (Path(args.kwekers) if args.kwekers else None, "kwekers"),
            (Path(args.leveranciers) if args.leveranciers else None, "leveranciers"),
        ])))


if __name__ == "__main__":
    main()
