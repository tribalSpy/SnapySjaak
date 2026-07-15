#!/usr/bin/env python3
import json
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET

import pandas as pd


PRODUCT_PARENT_XPATH = "./DEELZENDINGEN/DEELZENDING"
PRODUCT_DETAILS_TAG = "DEELZENDINGGEGEVENS"
TOTAL_TEXT_XPATH = "./Z_AANTALLEN/Z_TOTALEN_TEKST"
PCNU_XPATHS = [
    "./ZENDINGGEGEVENS/CFT_NUMMER",
    "./ZENDINGGEGEVENS/SPS_DOC_ID",
    "./ZENDINGGEGEVENS/CFT_NUMMER_BARCODE",
    "./ZENDINGGEGEVENS/CERTIFICAAT_ID",
]
DESTINATION_COUNTRY_XPATH = "./ZENDINGGEGEVENS/SPS_COT_CN_COUNTRYNAME"
ORIGIN_COUNTRY_XPATH = "./ZENDINGGEGEVENS/ORIGINE_ZENDING"
CONSIGNEE_NAME_XPATH = "./ZENDINGGEGEVENS/SPS_COT_CN_NAME"


class PhytoXmlValidationError(ValueError):
    def __init__(
        self,
        *,
        xml_filename: str,
        row_number: str = "",
        product_name: str = "",
        field: str,
        actual_value: Any,
        expected_condition: str,
    ) -> None:
        self.xml_filename = xml_filename
        self.row_number = row_number
        self.product_name = product_name
        self.field = field
        self.actual_value = actual_value
        self.expected_condition = expected_condition
        message = (
            f"{xml_filename} | row={row_number or '-'} | product={product_name or '-'} "
            f"| field={field} | actual={actual_value!r} | expected={expected_condition}"
        )
        super().__init__(message)


@dataclass(frozen=True)
class ParsedTotal:
    quantity: float | int | None
    unit: str | None
    raw_text: str


def clean_text(value: Any) -> str:
    return str(value or "").replace("\x00", "").strip()


def find_text(parent: ET.Element | None, tag: str) -> str:
    if parent is None:
        return ""
    node = parent.find(tag)
    return clean_text(node.text if node is not None else "")


def find_first_text(parent: ET.Element | None, tags: list[str]) -> str:
    for tag in tags:
        value = find_text(parent, tag)
        if value:
            return value
    return ""


def find_first_xpath_text(parent: ET.Element | None, xpaths: list[str]) -> str:
    if parent is None:
        return ""
    for xpath in xpaths:
        node = parent.find(xpath)
        value = clean_text(node.text if node is not None else "")
        if value:
            return value
    return ""


def normalize_pcnu_number(value: Any) -> str:
    text = clean_text(value)
    if not text:
        return ""
    if text.startswith("*") and text.endswith("*") and len(text) > 2:
        text = text[1:-1]
    return "".join(character for character in text if character.isalnum())


def parse_numeric(value: Any) -> float | int | None:
    text = clean_text(value).replace("\u00a0", "").replace(" ", "")
    if not text:
        return None
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            normalized_text = text.replace(".", "").replace(",", ".")
        else:
            normalized_text = text.replace(",", "")
    elif "," in text:
        parts = text.split(",")
        if len(parts) > 1 and all(part.isdigit() for part in parts) and all(len(part) == 3 for part in parts[1:]):
            normalized_text = "".join(parts)
        else:
            normalized_text = text.replace(",", ".")
    else:
        normalized_text = text
    try:
        number = float(normalized_text)
    except ValueError:
        return None
    if number.is_integer():
        return int(number)
    return number


def parse_total_text(total_text: str) -> ParsedTotal:
    cleaned = clean_text(total_text)
    if not cleaned:
        return ParsedTotal(quantity=None, unit=None, raw_text=cleaned)
    parts = cleaned.split()
    if not parts:
        return ParsedTotal(quantity=None, unit=None, raw_text=cleaned)
    quantity = parse_numeric(parts[0])
    unit = parts[1] if len(parts) > 1 else None
    return ParsedTotal(quantity=quantity, unit=unit, raw_text=cleaned)


def line_number_sort_key(value: Any) -> tuple[int, str]:
    text = clean_text(value)
    parsed = parse_numeric(text)
    if parsed is None:
        return (10**9, text)
    return (int(parsed), text)


def extract_product_record(parent_node: ET.Element) -> dict[str, Any]:
    details = parent_node.find(PRODUCT_DETAILS_TAG)
    product_name = find_first_text(
        details,
        ["DZ_PRODUCT_NAAM", "SPS_QPT_COMMON_NAME", "DZ_BOTNAAM_PRESENTATIE", "DZ_BOTANISCHE_NAAM", "SPS_QPT_SCIENTIFIC_NAME"],
    )
    botanical_name = find_first_text(
        details,
        ["DZ_BOTNAAM_PRESENTATIE", "DZ_BOTANISCHE_NAAM", "SPS_QPT_SCIENTIFIC_NAME", "DZ_PRODUCT_NAAM"],
    )
    record = {
        "line_number": find_text(details, "DZ_NUMMER"),
        "product_name": product_name or None,
        "botanical_name": botanical_name or None,
        "package_count": parse_numeric(find_first_text(details, ["DZ_AANTAL", "SPS_QPT_PHYSICAL_SPS_PACKAGE_ITEM_QUANTITY"])),
        "package_unit": find_first_text(details, ["DZ_HANDELSEENHEID"]),
        "declared_quantity": parse_numeric(
            find_first_text(
                details,
                ["DZ_NETTO_HOEVEELHEID_INT", "DZ_NETTO_HOEVEELHEID"],
            )
        ),
        "quantity_unit": find_first_text(details, ["DZ_EENHEID_PRESENTATIE", "DZ_EENHEID"]),
        "formatted_quantity": find_first_text(details, ["DZ_NETTO"]),
    }
    if not record["package_unit"]:
        record["package_unit"] = None
    if not record["quantity_unit"]:
        record["quantity_unit"] = None
    if not record["formatted_quantity"]:
        record["formatted_quantity"] = None
    return record


def parse_phyto_xml(xml_path: str | Path) -> dict[str, Any]:
    path = Path(xml_path)
    root = ET.parse(path).getroot()
    parent_nodes = root.findall(PRODUCT_PARENT_XPATH)
    rows = [extract_product_record(node) for node in parent_nodes]
    rows.sort(key=lambda row: line_number_sort_key(row.get("line_number")))
    dataframe = pd.DataFrame(
        rows,
        columns=[
            "line_number",
            "product_name",
            "botanical_name",
            "package_count",
            "package_unit",
            "declared_quantity",
            "quantity_unit",
            "formatted_quantity",
        ],
    )
    total = parse_total_text(find_text(root, TOTAL_TEXT_XPATH))
    pcnu_number = normalize_pcnu_number(find_first_xpath_text(root, PCNU_XPATHS))
    metadata = {
        "xml_filename": path.name,
        "product_parent_xpath": PRODUCT_PARENT_XPATH,
        "product_details_tag": PRODUCT_DETAILS_TAG,
        "product_parent_count": len(parent_nodes),
        "pcnu_number": pcnu_number,
        "destination_country": find_text(root, DESTINATION_COUNTRY_XPATH),
        "origin_country": find_text(root, ORIGIN_COUNTRY_XPATH),
        "consignee": find_text(root, CONSIGNEE_NAME_XPATH),
        "total_field": {
            "raw_text": total.raw_text,
            "quantity": total.quantity,
            "unit": total.unit,
        },
    }
    return {
        "rows": rows,
        "dataframe": dataframe,
        "metadata": metadata,
    }


def validate_phyto_xml(
    parsed: dict[str, Any],
    *,
    expected_product_count: int | None = None,
    expected_total_quantity: float | int | None = None,
    expected_total_unit: str | None = None,
    expected_rows: dict[str, dict[str, Any]] | None = None,
) -> dict[str, Any]:
    rows = parsed["rows"]
    dataframe: pd.DataFrame = parsed["dataframe"]
    metadata = parsed["metadata"]
    xml_filename = metadata["xml_filename"]
    parent_count = metadata["product_parent_count"]

    if len(rows) != parent_count:
        raise PhytoXmlValidationError(
            xml_filename=xml_filename,
            field="row_count",
            actual_value=len(rows),
            expected_condition=f"must equal repeating XML node count {parent_count}",
        )

    seen_line_numbers: set[str] = set()
    duplicate_count = 0
    missing_quantity_count = 0
    sample_assertions_passed = True
    total_package_count = 0
    total_declared_quantity = 0
    quantity_unit = ""

    for row in rows:
        line_number = clean_text(row.get("line_number"))
        product_name = clean_text(row.get("product_name"))
        declared_quantity = row.get("declared_quantity")
        package_count = row.get("package_count")
        row_quantity_unit = clean_text(row.get("quantity_unit"))

        if line_number in seen_line_numbers:
            duplicate_count += 1
            raise PhytoXmlValidationError(
                xml_filename=xml_filename,
                row_number=line_number,
                product_name=product_name,
                field="line_number",
                actual_value=line_number,
                expected_condition="must be unique",
            )
        seen_line_numbers.add(line_number)

        if not product_name:
            raise PhytoXmlValidationError(
                xml_filename=xml_filename,
                row_number=line_number,
                product_name=product_name,
                field="product_name",
                actual_value=product_name,
                expected_condition="must not be empty",
            )

        if declared_quantity is None:
            missing_quantity_count += 1
            raise PhytoXmlValidationError(
                xml_filename=xml_filename,
                row_number=line_number,
                product_name=product_name,
                field="declared_quantity",
                actual_value=declared_quantity,
                expected_condition="must not be null",
            )

        if package_count is not None:
            total_package_count += int(package_count)
        total_declared_quantity += int(declared_quantity) if float(declared_quantity).is_integer() else float(declared_quantity)
        if row_quantity_unit:
            quantity_unit = quantity_unit or row_quantity_unit
            if quantity_unit != row_quantity_unit:
                raise PhytoXmlValidationError(
                    xml_filename=xml_filename,
                    row_number=line_number,
                    product_name=product_name,
                    field="quantity_unit",
                    actual_value=row_quantity_unit,
                    expected_condition=f"must match certificate unit {quantity_unit}",
                )

    if expected_product_count is not None and len(rows) != expected_product_count:
        raise PhytoXmlValidationError(
            xml_filename=xml_filename,
            field="row_count",
            actual_value=len(rows),
            expected_condition=f"must equal {expected_product_count}",
        )

    total_field = metadata.get("total_field") or {}
    xml_total_quantity = total_field.get("quantity")
    xml_total_unit = clean_text(total_field.get("unit"))
    if xml_total_quantity is not None and total_declared_quantity != xml_total_quantity:
        raise PhytoXmlValidationError(
            xml_filename=xml_filename,
            field="declared_quantity_sum",
            actual_value=total_declared_quantity,
            expected_condition=f"must equal XML total {xml_total_quantity}",
        )

    if expected_total_quantity is not None and total_declared_quantity != expected_total_quantity:
        raise PhytoXmlValidationError(
            xml_filename=xml_filename,
            field="declared_quantity_sum",
            actual_value=total_declared_quantity,
            expected_condition=f"must equal {expected_total_quantity}",
        )

    if expected_total_unit is not None and quantity_unit != expected_total_unit:
        raise PhytoXmlValidationError(
            xml_filename=xml_filename,
            field="quantity_unit",
            actual_value=quantity_unit,
            expected_condition=f"must equal {expected_total_unit}",
        )

    if expected_rows:
        for line_number, assertions in expected_rows.items():
            match = dataframe.loc[dataframe["line_number"] == line_number]
            if match.empty:
                raise PhytoXmlValidationError(
                    xml_filename=xml_filename,
                    row_number=line_number,
                    field="line_number",
                    actual_value="missing",
                    expected_condition="sample row must exist",
                )
            row = match.iloc[0].to_dict()
            for field, expected_value in assertions.items():
                actual_value = row.get(field)
                if actual_value != expected_value:
                    sample_assertions_passed = False
                    raise PhytoXmlValidationError(
                        xml_filename=xml_filename,
                        row_number=line_number,
                        product_name=clean_text(row.get("product_name")),
                        field=field,
                        actual_value=actual_value,
                        expected_condition=f"must equal {expected_value!r}",
                    )

    first_row_number = clean_text(rows[0]["line_number"]) if rows else ""
    last_row_number = clean_text(rows[-1]["line_number"]) if rows else ""
    return {
        "extracted_row_count": len(rows),
        "first_row_number": first_row_number,
        "last_row_number": last_row_number,
        "total_package_count": total_package_count,
        "total_declared_quantity": total_declared_quantity,
        "quantity_unit": quantity_unit or xml_total_unit,
        "duplicate_count": duplicate_count,
        "missing_quantity_count": missing_quantity_count,
        "all_sample_assertions_passed": sample_assertions_passed,
    }


def export_phyto_xml(
    parsed: dict[str, Any],
    output_dir: str | Path,
    *,
    basename: str | None = None,
) -> dict[str, str]:
    dataframe: pd.DataFrame = parsed["dataframe"]
    metadata = parsed["metadata"]
    destination = Path(output_dir)
    destination.mkdir(parents=True, exist_ok=True)
    stem = basename or Path(metadata["xml_filename"]).stem
    csv_path = destination / f"{stem}.csv"
    json_path = destination / f"{stem}.json"
    excel_path = destination / f"{stem}.xlsx"

    dataframe.to_csv(csv_path, index=False, encoding="utf-8-sig")
    dataframe.to_json(json_path, orient="records", force_ascii=False, indent=2)
    dataframe.to_excel(excel_path, index=False)
    return {
        "csv": str(csv_path),
        "json": str(json_path),
        "excel": str(excel_path),
    }


def build_worker_payload(parsed: dict[str, Any]) -> dict[str, Any]:
    metadata = parsed["metadata"]
    product_lines = []
    total_quantity = 0
    for row in parsed["rows"]:
        quantity = row.get("declared_quantity")
        product_lines.append({
            "line_number": row.get("line_number"),
            "product": row.get("product_name"),
            "botanical_name": row.get("botanical_name"),
            "packages": row.get("package_count"),
            "package_unit": row.get("package_unit"),
            "quantity": quantity,
            "quantity_unit": row.get("quantity_unit"),
            "formatted_quantity": row.get("formatted_quantity"),
        })
        if quantity is not None:
            total_quantity += int(quantity) if float(quantity).is_integer() else float(quantity)
    total_field = metadata.get("total_field") or {}
    return {
        "document_state": "ok",
        "pcnu_number": metadata.get("pcnu_number") or "",
        "destination_country": metadata.get("destination_country") or "",
        "origin_country": metadata.get("origin_country") or "",
        "consignee": metadata.get("consignee") or "",
        "product_lines": product_lines,
        "total_quantity": total_quantity if total_quantity else None,
        "total_quantity_unit": clean_text(total_field.get("unit")),
        "xml_total_quantity": total_field.get("quantity"),
        "xml_total_raw_text": total_field.get("raw_text"),
        "problems": [],
        "source_format": "xml",
        "product_parent_xpath": PRODUCT_PARENT_XPATH,
    }


def build_validation_report(parsed: dict[str, Any], validation: dict[str, Any]) -> dict[str, Any]:
    total_field = parsed["metadata"].get("total_field") or {}
    return {
        "extracted_row_count": validation["extracted_row_count"],
        "first_row_number": validation["first_row_number"],
        "last_row_number": validation["last_row_number"],
        "total_package_count": validation["total_package_count"],
        "total_declared_quantity": validation["total_declared_quantity"],
        "quantity_unit": validation["quantity_unit"],
        "duplicate_count": validation["duplicate_count"],
        "missing_quantity_count": validation["missing_quantity_count"],
        "all_sample_assertions_passed": validation["all_sample_assertions_passed"],
        "xml_total_field": total_field,
    }


def main(argv: list[str] | None = None) -> int:
    args = list(argv or [])
    if not args:
        print("Usage: phyto_xml_parser.py <xml_path> [output_dir]")
        return 1
    xml_path = Path(args[0])
    output_dir = Path(args[1]) if len(args) > 1 else xml_path.parent / "phyto_xml_exports"

    parsed = parse_phyto_xml(xml_path)
    validation = validate_phyto_xml(
        parsed,
        expected_product_count=40,
        expected_total_quantity=1044,
        expected_total_unit="Pieces",
        expected_rows={
            "0001": {
                "product_name": "Aloe vera",
                "package_count": 5,
                "package_unit": "Box",
                "declared_quantity": 78,
                "quantity_unit": "Pieces",
            },
            "0026": {
                "product_name": "Kalanchoe blossfeldiana",
                "package_count": 13,
                "package_unit": "Box",
                "declared_quantity": 119,
                "quantity_unit": "Pieces",
            },
            "0034": {
                "product_name": "Rosa hybrid",
                "package_count": 10,
                "package_unit": "Box",
                "declared_quantity": 120,
                "quantity_unit": "Pieces",
            },
        },
    )
    exports = export_phyto_xml(parsed, output_dir)
    report = build_validation_report(parsed, validation)
    print(json.dumps({
        "product_parent_xpath": PRODUCT_PARENT_XPATH,
        "exports": exports,
        "validation_report": report,
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
