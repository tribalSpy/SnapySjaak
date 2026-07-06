import argparse
import json
import re
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook


CARRIER_GROUPS = [
    {"sheetLabel": "Breewel", "siteKlant": "Breewel"},
    {"sheetLabel": "ML Express", "siteKlant": "ML Express Parijs"},
    {"sheetLabel": "De Wit", "siteKlant": "De Wit"},
    {"sheetLabel": "De Wit 2", "siteKlant": "De Wit 2"},
]

DATE_RE = re.compile(r"(\d{2})-(\d{2})-(\d{4})")


def clean_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def to_number(value):
    text = clean_text(value).replace(",", ".")
    if not text:
        return 0
    try:
        return int(float(text))
    except ValueError:
        return 0


def load_rows(input_path: Path):
    data = input_path.read_bytes()
    workbook = load_workbook(filename=BytesIO(data), data_only=True, read_only=True)
    worksheet = workbook["Overzicht"] if "Overzicht" in workbook.sheetnames else workbook[workbook.sheetnames[0]]
    return [list(row) for row in worksheet.iter_rows(values_only=True)]


def parse_rows(input_path: Path):
    rows = load_rows(input_path)
    group_start_col = {}
    col = 1
    for label in ["Breewel", "ML Express", "De Wit", "De Wit 2", "Totaal"]:
        group_start_col[label] = col
        col += 3

    parsed = []
    for row_index, row in enumerate(rows[2:], start=3):
        label = clean_text(row[0] if len(row) > 0 else "")
        if not label or label.lower().startswith("totaal"):
            continue
        match = DATE_RE.search(label)
        if not match:
            continue
        iso_date = f"{match.group(3)}-{match.group(2)}-{match.group(1)}"
        for group in CARRIER_GROUPS:
            start_col = group_start_col[group["sheetLabel"]]
            dc = to_number(row[start_col] if len(row) > start_col else 0)
            dcs = to_number(row[start_col + 1] if len(row) > start_col + 1 else 0)
            dco = to_number(row[start_col + 2] if len(row) > start_col + 2 else 0)
            if dc == 0 and dcs == 0 and dco == 0:
                continue
            parsed.append({
                "type": "OUT",
                "action_date": iso_date,
                "source_date_label": label,
                "source_row_number": row_index,
                "country": "FR",
                "customer_name": group["siteKlant"],
                "metrics": {
                    "dc": dc,
                    "cctag": 0,
                    "dcs": dcs,
                    "dco": dco,
                    "pal": 0,
                    "vk": 0,
                },
            })
    return {
        "sheet_name": "Overzicht",
        "records": parsed,
    }


def main():
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(dest="command", required=True)
    parse_parser = subparsers.add_parser("parse")
    parse_parser.add_argument("--input", required=True)
    args = parser.parse_args()

    if args.command == "parse":
        payload = parse_rows(Path(args.input))
        print(json.dumps(payload))


if __name__ == "__main__":
    main()
