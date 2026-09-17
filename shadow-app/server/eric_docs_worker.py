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


def parse_number(text: str):
    cleaned = text.strip().replace(",", "")
    try:
        return float(cleaned)
    except ValueError:
        return None


# The temporary phytosanitary certificate (e-cert.nl style) always prints a
# "TOTAL <n> Pieces" summary line -- a far more reliable anchor than trying
# to match individual product lines across two differently-formatted
# documents (product names/ordering don't line up 1:1 between them).
def parse_phyto_total(text: str):
    match = re.search(r"TOTAL\s+([\d.,]+)\s+Pieces", text, re.IGNORECASE)
    return parse_number(match.group(1)) if match else None


# The NVWA "Inspectielijst" always ends with a "TOTAAL <n> Stuks" summary row.
def parse_inspection_total(text: str):
    match = re.search(r"TOTAAL\s+([\d.,]+)\s+Stuks", text, re.IGNORECASE)
    return parse_number(match.group(1)) if match else None


def compare_documents(phyto_path: Path, inspection_path: Path) -> dict:
    result = {
        "phyto_total": None,
        "inspection_total": None,
        "match": None,
        "ok": False,
        "error": "",
    }
    phyto_text = extract_pdf_text(phyto_path)
    inspection_text = extract_pdf_text(inspection_path)
    result["phyto_total"] = parse_phyto_total(phyto_text)
    result["inspection_total"] = parse_inspection_total(inspection_text)
    if result["phyto_total"] is None:
        result["error"] = 'Could not find a "TOTAL ... Pieces" line in the temporary phyto PDF'
        return result
    if result["inspection_total"] is None:
        result["error"] = 'Could not find a "TOTAAL ... Stuks" line in the inspection list PDF'
        return result
    result["match"] = abs(result["phyto_total"] - result["inspection_total"]) < 0.01
    result["ok"] = True
    return result


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(dest="command", required=True)
    compare_parser = subparsers.add_parser("compare")
    compare_parser.add_argument("--phyto", required=True)
    compare_parser.add_argument("--inspection", required=True)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.command == "compare":
        try:
            result = compare_documents(Path(args.phyto), Path(args.inspection))
        except Exception as error:  # noqa: BLE001 -- surface any failure as a normal check result
            result = {
                "phyto_total": None, "inspection_total": None, "match": None,
                "ok": False, "error": str(error),
            }
        print(json.dumps(result, ensure_ascii=True))


if __name__ == "__main__":
    main()
