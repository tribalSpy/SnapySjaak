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
    elif "€" in text:
        result["currency"] = "EUR"

    total_matches = re.findall(r"(?<!Sub )Total\s*[£€]\s*([\d,]+\.\d{2})", text)
    standalone_total = float(total_matches[-1].replace(",", "")) if total_matches else None

    summary_matches = re.findall(
        r"Nett\s+Gross\s+Colli\s+Pieces\s+Amount\s*\n\s*(\d+)\s*KG\s+(\d+)\s*KG\s+(\d+)\s+(\d+)\s+([\d,]+\.\d{2})",
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
    summary_amount = float(amount.replace(",", ""))
    result["total_amount"] = standalone_total if standalone_total is not None else summary_amount
    if standalone_total is not None and abs(summary_amount - standalone_total) > 0.01:
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
