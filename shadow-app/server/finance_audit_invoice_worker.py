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
# earlier in the document (used here as a cross-check on the PDF's own
# internal consistency, separate from the packaging question below).
def parse_invoice_text(text: str) -> dict:
    result = {
        "invoice_number": "",
        "currency": "",
        "total_amount": None,
        "grand_total": None,
        "packaging_cost": None,
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
    grand_total = normalize_amount(total_matches[-1]) if total_matches else None

    # Some suppliers charge packaging/embalage separately -- we never invoice
    # or export that, so the value that should match our own records is the
    # goods-only subtotal from just before the packaging charge, not the
    # final Total. The one reliable anchor across different suppliers'
    # layouts (a repeated grand-total "Sub total" after the packaging block,
    # or none at all) is the standalone "Package <cost>" line itself: take
    # the last "Sub total" appearing before it. No packaging charge line
    # found means nothing to exclude -- goods total is just the grand total.
    sub_total_matches = list(re.finditer(r"Sub total[^\d\n]{0,6}(\d[\d.,]*\d)", text))
    package_match = re.search(r"Package\s+(\d[\d.,]*\d)\s*$", text, re.MULTILINE)
    goods_total = grand_total
    packaging_cost = None
    if package_match:
        before_package = [m for m in sub_total_matches if m.start() < package_match.start()]
        if before_package:
            goods_total = normalize_amount(before_package[-1].group(1))
            packaging_cost = normalize_amount(package_match.group(1))

    summary_matches = re.findall(
        r"Nett\s+Gross\s+Colli\s+Pieces\s+Amount\s*\n\s*(\d+)\s*KG\s+(\d+)\s*KG\s+(\d+)\s+(\d+)\s+(\d[\d.,]*\d)",
        text,
    )
    if not summary_matches:
        result["error"] = "Could not find the Nett/Gross/Colli/Pieces/Amount summary row"
        result["total_amount"] = goods_total
        result["grand_total"] = grand_total
        return result

    nett, gross, colli, pieces, amount = summary_matches[-1]
    result["nett_kg"] = int(nett)
    result["gross_kg"] = int(gross)
    result["colli"] = int(colli)
    result["pieces"] = int(pieces)
    summary_amount = normalize_amount(amount)
    result["grand_total"] = grand_total if grand_total is not None else summary_amount
    result["total_amount"] = goods_total if goods_total is not None else summary_amount
    result["packaging_cost"] = packaging_cost
    if grand_total is not None and summary_amount is not None and abs(summary_amount - grand_total) > 0.01:
        result["error"] = f"Total {result['currency'] or ''} {grand_total} does not match the summary row amount {summary_amount}"

    return result


def parse_invoice_file(input_path: Path) -> dict:
    text = extract_pdf_text(input_path)
    parsed = parse_invoice_text(text)
    parsed["ok"] = not parsed["error"] and parsed["total_amount"] is not None
    return parsed


# ---------------------------------------------------------------------------
# Explain a mismatch: the factuur's lines against the lines of the data file
# uploaded in UKdocs (one row per line: goodsDescriptionText, quantityValue,
# netPriceValue, value, packages, originCountryCode). What's left after
# pairing is the difference.

NUMBER = r"\d[\d.,]*"
# Flowers (EUR): "20 Alstroemeria FL Char Blanca 75 70 whi NL 0,48 9,60"
PLAIN_LINE_RE = re.compile(rf"^(\d+)\s+(.+?)\s+(\d+[.,]\d{{2,3}})\s+({NUMBER}[.,]\d{{2}})\s*$")
# Item (GBP), boxes x per box = pieces: "6 x 5 = 30 Chico Premium Asteca pb
# 50 A1 30 MX MX 31 2.51 75.30", "50 x1 = 50 Bags Paper ... 0.60 30.00"
ITEM_LINE_RE = re.compile(rf"^(\d+)\s+x\s*\d+\s*=\s*(\d+)\s+(.+?)\s+(\d+[.,]\d{{2,3}})\s+({NUMBER}[.,]\d{{2}})\s*$")
# A bouquet's content: "24 x 3.00 = 72.00 Chr t boeket promo 1 80 50 NL NL 37 23.52"
PART_LINE_RE = re.compile(rf"^(\d+)\s+x\s*\d+[.,]\d+\s*=\s*(\d+)(?:[.,]\d+)?\s+(.+?)\s+({NUMBER}[.,]\d{{2}})\s*$")


def _decimal(text: str) -> float | None:
    """"0,48", "0.485", "1.234,50", "5960,80" -> a number."""
    value = normalize_amount(text)
    if value is not None:
        return value
    try:
        return float(text.replace(",", "."))
    except ValueError:
        return None


def parse_invoice_lines(text: str) -> list[dict]:
    """The goods lines, up to the last "Sub total" (the packaging and the
    statistics block after it are never in the export). A bouquet counts
    by its contents -- the data file lists those, not the bouquet. A plain
    or item line counts only when pieces x price gives its amount, so a
    header or address line never slips in."""
    cut = text.rfind("Sub total")
    goods = text[:cut] if cut >= 0 else text
    lines = []
    raw = [line.strip() for line in goods.splitlines()]
    for index, line in enumerate(raw):
        part = PART_LINE_RE.match(line)
        if part:
            pieces, description, amount = int(part.group(2)), part.group(3), _decimal(part.group(4))
            if amount is not None and pieces:
                lines.append({"pieces": pieces, "description": description.strip(), "price": round(amount / pieces, 3), "amount": round(amount, 2)})
            continue
        item = ITEM_LINE_RE.match(line)
        if item:
            # A bouquet: its contents follow on the next lines -- those count.
            following = next((later for later in raw[index + 1:] if later), "")
            if PART_LINE_RE.match(following):
                continue
            pieces, description, price, amount = int(item.group(2)), item.group(3), _decimal(item.group(4)), _decimal(item.group(5))
        else:
            plain = PLAIN_LINE_RE.match(line)
            if not plain:
                continue
            pieces, description, price, amount = int(plain.group(1)), plain.group(2), _decimal(plain.group(3)), _decimal(plain.group(4))
        if amount is None or price is None or abs(pieces * price - amount) > 0.02 + amount * 0.001:
            continue
        lines.append({"pieces": pieces, "description": description.strip(), "price": round(price, 3), "amount": round(amount, 2)})
    return lines


def read_export_lines(path: Path) -> list[dict]:
    import ukdocs_worker  # same folder; its xlsx reader needs no extra packages
    rows = ukdocs_worker.parse_xlsx_rows(path.read_bytes())
    if not rows:
        return []
    header = {str(name or "").strip().lower(): index for index, name in enumerate(rows[0])}

    def cell(row, name):
        index = header.get(name.lower())
        return row[index] if index is not None and index < len(row) else None

    def number(value):
        try:
            return float(str(value).replace(",", ".")) if value not in (None, "") else 0.0
        except ValueError:
            return 0.0

    lines = []
    for row in rows[1:]:
        if cell(row, "quantityValue") in (None, ""):
            continue
        lines.append({
            "pieces": int(round(number(cell(row, "quantityValue")))),
            "description": str(cell(row, "goodsDescriptionText") or cell(row, "vbnDescription") or "").strip(),
            "price": round(number(cell(row, "netPriceValue")), 3),
            "amount": round(number(cell(row, "value")), 2),
            "colli": number(cell(row, "packages")),
            "origin": str(cell(row, "originCountryCode") or "").strip(),
        })
    return lines


def _name_key(description: str) -> str:
    """The first word's first 5 letters: "Antirr." and "Antirrhinum" are the
    same product, the factuur adds length/grade/colour after the name."""
    word = re.sub(r"[^a-z]", "", (description or "").lower().split(" ")[0] if description else "")
    return word[:5]


def _same(a: float, b: float, cents: float = 0.005) -> bool:
    return abs(a - b) < cents


def _subset_with(items: list[dict], pieces: int, amount: float, limit: int = 10):
    """A few lines (same product) that together are pieces / amount -- the
    factuur splits a line the export has as one (10 + 60 = 70)."""
    from itertools import combinations
    items = items[:limit]
    for size in range(2, len(items) + 1):
        for combo in combinations(items, size):
            if sum(item["pieces"] for item in combo) == pieces and _same(sum(item["amount"] for item in combo), amount, 0.02):
                return list(combo)
    return None


def explain_difference(invoice_lines: list[dict], export_lines: list[dict]) -> dict:
    facts = [dict(line) for line in invoice_lines]
    exports = [dict(line) for line in export_lines]

    def pair(match):
        nonlocal facts, exports
        left = []
        for fact in facts:
            hit = next((line for line in exports if match(fact, line)), None)
            if hit is not None:
                exports.remove(hit)
            else:
                left.append(fact)
        facts = left

    # 1. The same line on both sides (same product first).
    pair(lambda f, e: _name_key(f["description"]) == _name_key(e["description"]) and f["pieces"] == e["pieces"] and _same(f["amount"], e["amount"]))
    pair(lambda f, e: f["pieces"] == e["pieces"] and _same(f["amount"], e["amount"]))
    # Most lines without their twin: the data file of another sending (picked
    # by mistake) -- say so instead of listing everything as a difference.
    same_lines = len(invoice_lines) - len(facts)
    if invoice_lines and same_lines < 0.5 * len(invoice_lines):
        return {
            "ok": False,
            "wrong_file": True,
            "error": f"This data file doesn't fit this factuur: only {same_lines} of its {len(invoice_lines)} lines are the same -- pick the right data file.",
            "text": "",
            "factuur": {"lines": len(invoice_lines), "pieces": sum(line["pieces"] for line in invoice_lines), "amount": round(sum(line["amount"] for line in invoice_lines), 2)},
            "export": {"lines": len(export_lines), "pieces": sum(line["pieces"] for line in export_lines), "amount": round(sum(line["amount"] for line in export_lines), 2)},
        }
    # 2. One side split over several lines of one product.
    for export in list(exports):
        same = [f for f in facts if _name_key(f["description"]) == _name_key(export["description"])]
        combo = _subset_with(same, export["pieces"], export["amount"]) if len(same) > 1 else None
        if combo:
            exports.remove(export)
            facts = [f for f in facts if not any(f is item for item in combo)]
    for fact in list(facts):
        same = [e for e in exports if _name_key(e["description"]) == _name_key(fact["description"])]
        combo = _subset_with(same, fact["pieces"], fact["amount"]) if len(same) > 1 else None
        if combo:
            facts.remove(fact)
            exports = [e for e in exports if not any(e is item for item in combo)]
    # 3. The same product, another price or other pieces.
    price_diff, pieces_diff = [], []
    for fact in list(facts):
        export = next((e for e in exports if _name_key(e["description"]) == _name_key(fact["description"]) and e["pieces"] == fact["pieces"]), None)
        if export is not None:
            price_diff.append({"factuur": fact, "export": export})
            facts.remove(fact)
            exports.remove(export)
    for fact in list(facts):
        export = next((e for e in exports if _name_key(e["description"]) == _name_key(fact["description"]) and _same(e["price"], fact["price"])), None)
        if export is not None:
            pieces_diff.append({"factuur": fact, "export": export})
            facts.remove(fact)
            exports.remove(export)

    def money(value):
        return f"{value:.2f}"

    def colli(line):
        value = line.get("colli") or 0
        return f", {value:g} colli" if value else ""

    parts = []
    if exports:
        parts.append("In export, not on factuur: " + "; ".join(
            f"{e['pieces']} {e['description']}{' (' + e['origin'] + ')' if e.get('origin') else ''} x {e['price']:g} = {money(e['amount'])}{colli(e)}" for e in exports))
    if facts:
        parts.append("On factuur, not in export: " + "; ".join(
            f"{f['pieces']} {f['description']} x {f['price']:g} = {money(f['amount'])}" for f in facts))
    if price_diff:
        parts.append("Other price: " + "; ".join(
            f"{d['export']['description']} {d['factuur']['pieces']} pcs factuur {d['factuur']['price']:g} / export {d['export']['price']:g} ({money(d['factuur']['amount'] - d['export']['amount'])})" for d in price_diff))
    if pieces_diff:
        parts.append("Other pieces: " + "; ".join(
            f"{d['export']['description']} factuur {d['factuur']['pieces']} / export {d['export']['pieces']} pcs ({money(d['factuur']['amount'] - d['export']['amount'])})" for d in pieces_diff))
    return {
        "factuur": {"lines": len(invoice_lines), "pieces": sum(line["pieces"] for line in invoice_lines), "amount": round(sum(line["amount"] for line in invoice_lines), 2)},
        "export": {"lines": len(export_lines), "pieces": sum(line["pieces"] for line in export_lines), "amount": round(sum(line["amount"] for line in export_lines), 2), "colli": sum(line.get("colli") or 0 for line in export_lines)},
        "only_export": exports,
        "only_factuur": facts,
        "price_diff": price_diff,
        "pieces_diff": pieces_diff,
        "text": ". ".join(parts) + ("." if parts else ""),
        "ok": True,
        "error": "",
    }


def explain_files(data_path: Path, pdf_paths: list[Path]) -> dict:
    invoice_lines = []
    for pdf in pdf_paths:
        invoice_lines += parse_invoice_lines(extract_pdf_text(pdf))
    export_lines = read_export_lines(data_path)
    if not invoice_lines:
        return {"ok": False, "error": "Could not read the factuur's lines", "text": ""}
    if not export_lines:
        return {"ok": False, "error": "Could not read the data file's lines", "text": ""}
    return explain_difference(invoice_lines, export_lines)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(dest="command", required=True)
    parse_parser = subparsers.add_parser("parse")
    parse_parser.add_argument("--input", required=True)
    explain_parser = subparsers.add_parser("explain")
    explain_parser.add_argument("--data", required=True)
    explain_parser.add_argument("--pdf", action="append", required=True)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.command == "parse":
        try:
            result = parse_invoice_file(Path(args.input))
        except Exception as error:  # noqa: BLE001 -- surface any failure as a normal parse error
            result = {
                "invoice_number": "", "currency": "", "total_amount": None,
                "grand_total": None, "packaging_cost": None,
                "nett_kg": None, "gross_kg": None, "colli": None, "pieces": None,
                "ok": False, "error": str(error),
            }
        print(json.dumps(result, ensure_ascii=True))
    elif args.command == "explain":
        try:
            result = explain_files(Path(args.data), [Path(item) for item in args.pdf])
        except Exception as error:  # noqa: BLE001
            result = {"ok": False, "error": str(error), "text": ""}
        print(json.dumps(result, ensure_ascii=True))


if __name__ == "__main__":
    main()
