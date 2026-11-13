import argparse
import copy
import json
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Font


ROW_BY_CODE = {
    "510": 19,
    "519": 22,
    "520": 25,
    "525": 28,
    "533": 31,
    "544": 34,
    "560": 37,
    "566": 40,
    "577": 43,
    "596": 46,
    "597": 49,
    "CC": 52,
    "CCO": 55,
    "CCS": 58,
    "VK": 61,
}
DATE_CELL = "C12"
CUSTOMER_CELL = "J12"
CODE_COLUMN = "B"
CUSTOM_ROW_START = 64
CUSTOM_ROW_STEP = 3
TOTAL_OK_COLUMN = "E"
TOTAL_BROKEN_COLUMN = "G"
EXPORTER_BLOCK_CELL = "B65"
SIGNATURE_BLOCK_CELL = "H65"
# The exporter and signature boxes under the code rows are merged cells
# (B65:E71 and H65:K71 in the template). Extra rows go above them, so the
# boxes move down 3 rows per extra row -- writing into a merged box fails
# ("'MergedCell' object attribute 'value' is read-only").
BLOCK_FIRST_ROW = 65
BLOCK_LAST_ROW = 71
BLOCK_RANGES = (("B", "E"), ("H", "K"))
LAST_CODE_ROW = 61
MAIN_FONT_SIZE = 16
CELL_FONT_SIZE = 20


def style_value_cell(cell):
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.font = Font(size=CELL_FONT_SIZE)


def style_label_cell(cell, bold=False):
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.font = Font(size=MAIN_FONT_SIZE, bold=bold)


def clean_text(value):
    if value is None:
        return ""
    return str(value).strip()


def to_count(value):
    text = clean_text(value).replace(",", ".")
    if not text:
        return 0
    try:
        return max(0, int(float(text)))
    except ValueError:
        raise ValueError(f"Invalid quantity: {value}")


def display_count(value):
    count = to_count(value)
    return count if count > 0 else None


def move_blocks_down(worksheet, shift):
    """Move the exporter and signature boxes `shift` rows down, keeping every
    cell's own border / font (openpyxl drops the inner borders on unmerge and
    re-merging copies only the top-left cell's border)."""
    if shift <= 0:
        return
    first_col = min(col for col, _ in BLOCK_RANGES)
    last_col = max(col for _, col in BLOCK_RANGES)
    saved = []
    for row in worksheet[f"{first_col}{BLOCK_FIRST_ROW}:{last_col}{BLOCK_LAST_ROW}"]:
        for cell in row:
            value = None if isinstance(cell, MergedCell) else cell.value
            saved.append((cell.row, cell.column, value, copy.copy(cell._style), cell.has_style))
    for start, end in BLOCK_RANGES:
        coord = f"{start}{BLOCK_FIRST_ROW}:{end}{BLOCK_LAST_ROW}"
        if coord in worksheet.merged_cells:
            worksheet.unmerge_cells(coord)
    # Empty the old place (plain cells again)...
    for row, column, _, _, _ in saved:
        worksheet._cells.pop((row, column), None)
    # ...merge the boxes at their new place, then put every cell's style back.
    for start, end in BLOCK_RANGES:
        worksheet.merge_cells(f"{start}{BLOCK_FIRST_ROW + shift}:{end}{BLOCK_LAST_ROW + shift}")
    for row, column, value, style, has_style in saved:
        target = worksheet.cell(row=row + shift, column=column)
        if value is not None and not isinstance(target, MergedCell):
            target.value = value
        if has_style:
            target._style = copy.copy(style)


def style_extra_row(worksheet, row_number):
    """An extra row looks like the standard code rows (same borders)."""
    for column in (CODE_COLUMN, TOTAL_OK_COLUMN, TOTAL_BROKEN_COLUMN):
        worksheet[f"{column}{row_number}"]._style = copy.copy(worksheet[f"{column}{LAST_CODE_ROW}"]._style)
    height = worksheet.row_dimensions[LAST_CODE_ROW].height
    if height:
        worksheet.row_dimensions[row_number].height = height


def load_payload(path: Path):
    payload = json.loads(path.read_text(encoding="utf-8"))
    customer_name = clean_text(payload.get("customer_name"))
    action_date = clean_text(payload.get("action_date"))
    rows = payload.get("rows") or []
    if not customer_name:
        raise ValueError("Customer is required")
    if not action_date:
        raise ValueError("Action date is required")
    if not isinstance(rows, list):
        raise ValueError("Rows must be a list")
    return {
        "customer_name": customer_name,
        "action_date": action_date,
        "exporter": payload.get("exporter") or {},
        "rows": rows,
    }


def generate_workbook(template_path: Path, payload_path: Path, output_path: Path):
    payload = load_payload(payload_path)
    workbook = load_workbook(template_path)
    worksheet = workbook[workbook.sheetnames[0]]

    worksheet[DATE_CELL] = payload["action_date"]
    worksheet[CUSTOMER_CELL] = payload["customer_name"]
    style_value_cell(worksheet[DATE_CELL])
    style_value_cell(worksheet[CUSTOMER_CELL])

    for code, row_number in ROW_BY_CODE.items():
        worksheet[f"{TOTAL_OK_COLUMN}{row_number}"] = None
        worksheet[f"{TOTAL_BROKEN_COLUMN}{row_number}"] = None
        style_label_cell(worksheet[f"{CODE_COLUMN}{row_number}"], bold=True)
        style_value_cell(worksheet[f"{TOTAL_OK_COLUMN}{row_number}"])
        style_value_cell(worksheet[f"{TOTAL_BROKEN_COLUMN}{row_number}"])

    custom_rows = []
    for row in payload["rows"]:
        code = clean_text(row.get("code")).upper()
        if not code:
            continue
        row_number = ROW_BY_CODE.get(code)
        if row_number is None:
            custom_rows.append({
                "code": code,
                "total_ok": display_count(row.get("total_ok")),
                "total_broken": display_count(row.get("total_broken")),
            })
            continue
        worksheet[f"{TOTAL_OK_COLUMN}{row_number}"] = display_count(row.get("total_ok"))
        worksheet[f"{TOTAL_BROKEN_COLUMN}{row_number}"] = display_count(row.get("total_broken"))

    # Room for the extra rows: the boxes move down 3 rows per extra row (the
    # same gap under the last row as under VK in the template).
    shift = len(custom_rows) * CUSTOM_ROW_STEP
    move_blocks_down(worksheet, shift)

    for index, row in enumerate(custom_rows):
        row_number = CUSTOM_ROW_START + (index * CUSTOM_ROW_STEP)
        style_extra_row(worksheet, row_number)
        worksheet[f"{CODE_COLUMN}{row_number}"] = row["code"]
        worksheet[f"{TOTAL_OK_COLUMN}{row_number}"] = row["total_ok"]
        worksheet[f"{TOTAL_BROKEN_COLUMN}{row_number}"] = row["total_broken"]
        style_value_cell(worksheet[f"{CODE_COLUMN}{row_number}"])
        style_value_cell(worksheet[f"{TOTAL_OK_COLUMN}{row_number}"])
        style_value_cell(worksheet[f"{TOTAL_BROKEN_COLUMN}{row_number}"])

    exporter_cell = worksheet[f"{EXPORTER_BLOCK_CELL[0]}{BLOCK_FIRST_ROW + shift}"]
    signature_cell = worksheet[f"{SIGNATURE_BLOCK_CELL[0]}{BLOCK_FIRST_ROW + shift}"]
    exporter_block = clean_text((payload.get("exporter") or {}).get("block"))
    if exporter_block:
        exporter_cell.value = exporter_block
        exporter_cell.alignment = Alignment(wrap_text=True, vertical="top")
        exporter_cell.font = Font(size=MAIN_FONT_SIZE)

    signature_cell.value = "Delivery signature"
    signature_cell.font = Font(size=MAIN_FONT_SIZE, bold=True)
    signature_cell.alignment = Alignment(vertical="top")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def main():
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(dest="command", required=True)
    generate_parser = subparsers.add_parser("generate")
    generate_parser.add_argument("--template", required=True)
    generate_parser.add_argument("--payload", required=True)
    generate_parser.add_argument("--output", required=True)
    args = parser.parse_args()

    if args.command == "generate":
        generate_workbook(Path(args.template), Path(args.payload), Path(args.output))
        print(json.dumps({"ok": True, "output": args.output}))


if __name__ == "__main__":
    main()
