from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from phyto_xml_parser import (  # noqa: E402
    PRODUCT_PARENT_XPATH,
    build_validation_report,
    build_worker_payload,
    export_phyto_xml,
    parse_phyto_xml,
    validate_phyto_xml,
)


FIXTURE_XML = Path(r"c:\Users\vdvfi\Downloads\Fytosanitair certificaat model 1 (geslachten gegroepeerd).xml")


def test_parse_phyto_xml_fixture():
    parsed = parse_phyto_xml(FIXTURE_XML)

    assert parsed["metadata"]["product_parent_xpath"] == PRODUCT_PARENT_XPATH
    assert parsed["metadata"]["product_parent_count"] == 40

    frame = parsed["dataframe"]
    assert len(frame) == 40
    assert frame.iloc[0]["line_number"] == "0001"
    assert frame.iloc[-1]["line_number"] == "0040"


def test_validate_phyto_xml_fixture():
    parsed = parse_phyto_xml(FIXTURE_XML)
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

    assert validation["extracted_row_count"] == 40
    assert validation["first_row_number"] == "0001"
    assert validation["last_row_number"] == "0040"
    assert validation["total_package_count"] == 106
    assert validation["total_declared_quantity"] == 1044
    assert validation["quantity_unit"] == "Pieces"
    assert validation["duplicate_count"] == 0
    assert validation["missing_quantity_count"] == 0
    assert validation["all_sample_assertions_passed"] is True


def test_export_and_worker_payload(tmp_path):
    parsed = parse_phyto_xml(FIXTURE_XML)
    validation = validate_phyto_xml(
        parsed,
        expected_product_count=40,
        expected_total_quantity=1044,
        expected_total_unit="Pieces",
    )
    exported = export_phyto_xml(parsed, tmp_path, basename="fixture-output")

    assert Path(exported["csv"]).exists()
    assert Path(exported["json"]).exists()
    assert Path(exported["excel"]).exists()

    worker_payload = build_worker_payload(parsed)
    report = build_validation_report(parsed, validation)

    assert worker_payload["pcnu_number"] == "358305340"
    assert worker_payload["total_quantity"] == 1044
    assert worker_payload["total_quantity_unit"] == "Pieces"
    assert len(worker_payload["product_lines"]) == 40
    assert report["all_sample_assertions_passed"] is True
