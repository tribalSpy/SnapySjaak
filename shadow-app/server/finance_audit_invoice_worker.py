from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

try:
    from pypdf import PdfReader
except ImportError:
    try:
        from PyPDF2 import PdfReader
    except ImportError:
        PdfReader = None


def extract_pdf_text(path: Path) -> str:
    if PdfReader is None:
        raise RuntimeError("PDF text extractor not installed (pypdf/PyPDF2)")
    reader = PdfReader(str(path))
    pages = []
    for page in reader.pages:
        try:
            pages.append(page.extract_text() or "")
        except Exception:
            pages.append("")
    return "\n".join(pages)


# Dutch/EUR invoices print amounts with a comma decimal ("19424,00") instead
# of a period ("19424.00") -- normalize by treating whichever separator sits
# right before the final 2 digits as the decimal point, and stripping any
# others (thousands separators, in either convention).
def normalize_amount(text: str):
    match = re.search(r"[.,](\d{2})$", text.strip())
    if not match:
        return None
    integer_part = re.sub(r"[.,]", "", text.strip()[: match.start()])
    if not integer_part:
        return None
    try:
        return float(f"{integer_part}.{match.group(1)}")
    except ValueError:
        return None


# The supplier invoice always ends with a grand-total breakdown line right
# after a "Nett Gross Colli Pieces Amount" header -- that single line gives
# nett/gross weight, colli (boxes), pieces, and the total amount together,
# and its amount always matches the standalone "Total <currency> x.xx" line
# earlier in the document (used here only as a cross-check).
def parse_invoice_text(text: str) -> dict:
    result = {
        "invoice_number": "",
        "currency": "",
        "total_amount": None,
        "nett_kg": None,
        "gross_kg": None,
        "colli": None,
        "pieces": None,
        "error": "",
    }

    match = re.search(r"Invoice\s+\d+\s*/\s*(\d+)", text)
    if match:
        result["invoice_number"] = match.group(1)

    if "£" in text:
        result["currency"] = "GBP"
    elif "€" in text or "¬" in text:
        # "¬" shows up in place of "€" from some suppliers' embedded PDF
        # fonts -- pypdf/PyPDF2 extract the glyph at the wrong codepoint.
        result["currency"] = "EUR"

    # Up to a few non-digit characters (currency symbol and/or spacing,
    # possibly garbled) between "Total" and the amount itself.
    total_matches = re.findall(r"(?<!Sub )Total[^\d\n]{0,6}(\d[\d.,]*\d)", text)
    standalone_total = normalize_amount(total_matches[-1]) if total_matches else None

    summary_matches = re.findall(
        r"Nett\s+Gross\s+Colli\s+Pieces\s+Amount\s*\n\s*(\d+)\s*KG\s+(\d+)\s*KG\s+(\d+)\s+(\d+)\s+(\d[\d.,]*\d)",
        text,
    )
    if not summary_matches:
        result["error"] = "Could not find the Nett/Gross/Colli/Pieces/Amount summary row"
        result["total_amount"] = standalone_total
        return result

    nett, gross, colli, pieces, amount = summary_matches[-1]
    result["nett_kg"] = int(nett)
    result["gross_kg"] = int(gross)
    result["colli"] = int(colli)
    result["pieces"] = int(pieces)
    summary_amount = normalize_amount(amount)
    result["total_amount"] = standalone_total if standalone_total is not None else summary_amount
    if standalone_total is not None and summary_amount is not None and abs(summary_amount - standalone_total) > 0.01:
        result["error"] = f"Total {result['currency'] or ''} {standalone_total} does not match the summary row amount {summary_amount}"

    return result


def parse_invoice_file(input_path: Path) -> dict:
    text = extract_pdf_text(input_path)
    parsed = parse_invoice_text(text)
    parsed["ok"] = not parsed["error"] and parsed["total_amount"] is not None
    return parsed


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(dest="command", required=True)
    parse_parser = subparsers.add_parser("parse")
    parse_parser.add_argument("--input", required=True)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.command == "parse":
        try:
            result = parse_invoice_file(Path(args.input))
        except Exception as error:  # noqa: BLE001 -- surface any failure as a normal parse error
            result = {
                "invoice_number": "", "currency": "", "total_amount": None,
                "nett_kg": None, "gross_kg": None, "colli": None, "pieces": None,
                "ok": False, "error": str(error),
            }
        print(json.dumps(result, ensure_ascii=True))


if __name__ == "__main__":
    main()
