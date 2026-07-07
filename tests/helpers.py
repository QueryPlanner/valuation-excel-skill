from __future__ import annotations

import copy
import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "assets" / "template.xlsx"
VALID_INPUTS = ROOT / "tests" / "fixtures" / "valid_inputs.json"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def load_valid_inputs() -> dict:
    return copy.deepcopy(json.loads(VALID_INPUTS.read_text(encoding="utf-8")))


def write_json(path: Path, value: dict) -> None:
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def analyst_evidence(metric: str, value, units: str = "decimal") -> dict:
    return {
        "metric": metric,
        "source_type": "analyst_assumption",
        "period": "Forecast",
        "calculation": "Test assumption",
        "value": value,
        "units": units,
        "rationale": "Fixture branch coverage.",
    }


def filing_evidence(metric: str, value, units: str = "USD millions") -> dict:
    return {
        "metric": metric,
        "source_type": "filing",
        "accession_number": "0000000001-26-000001",
        "source_url": "https://www.sec.gov/Archives/example-filing.htm",
        "sha256": "a" * 64,
        "period": "FY 2025",
        "section": "Test note",
        "reported_label": "Test label",
        "calculation": "Reported value",
        "value": value,
        "units": units,
    }


def replace_evidence(inputs: dict, metric: str, evidence: dict) -> None:
    inputs["source_evidence"] = [
        record for record in inputs["source_evidence"] if record.get("metric") != metric
    ]
    inputs["source_evidence"].append(evidence)


def inject_cached_values(
    source: Path,
    destination: Path,
    values: dict[str, dict[str, tuple[object, str | None]]],
) -> None:
    with zipfile.ZipFile(source, "r") as archive:
        workbook_xml = ElementTree.fromstring(archive.read("xl/workbook.xml"))
        relationships = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {
            relation.attrib["Id"]: relation.attrib["Target"]
            for relation in relationships.findall(f"{{{PKG_REL_NS}}}Relationship")
        }
        sheet_targets: dict[str, str] = {}
        sheets = workbook_xml.find(f"{{{MAIN_NS}}}sheets")
        assert sheets is not None
        for sheet in sheets:
            relation_id = sheet.attrib[f"{{{DOC_REL_NS}}}id"]
            target = targets[relation_id]
            if target.startswith("/"):
                target_path = target.lstrip("/")
            else:
                target_path = f"xl/{target.lstrip('/')}"
            sheet_targets[sheet.attrib["name"]] = target_path

        modified: dict[str, bytes] = {}
        for sheet_name, cell_values in values.items():
            target = sheet_targets[sheet_name]
            root = ElementTree.fromstring(archive.read(target))
            for coordinate, (value, cell_type) in cell_values.items():
                cell = root.find(f".//{{{MAIN_NS}}}c[@r='{coordinate}']")
                assert cell is not None, f"{sheet_name}!{coordinate} missing"
                value_element = cell.find(f"{{{MAIN_NS}}}v")
                if value_element is None:
                    value_element = ElementTree.SubElement(cell, f"{{{MAIN_NS}}}v")
                value_element.text = str(value)
                if cell_type:
                    cell.attrib["t"] = cell_type
                else:
                    cell.attrib.pop("t", None)
            modified[target] = ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)

        with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as output:
            for item in archive.infolist():
                output.writestr(item, modified.get(item.filename, archive.read(item.filename)))
