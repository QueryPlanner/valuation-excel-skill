from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import openpyxl
from common import atomic_write_json, file_sha256
from verify_workbook import verify_recalculated

WORKBOOK_XML = "xl/workbook.xml"


class LibreOfficeRecalculationError(RuntimeError):
    pass


def _resolve_executable(executable: str | Path | None = None) -> str:
    if executable:
        candidate = Path(executable).expanduser()
        if candidate.is_file():
            return str(candidate.resolve())
        resolved = shutil.which(str(executable))
        if resolved:
            return resolved
        raise LibreOfficeRecalculationError(f"LibreOffice executable was not found: {executable}.")

    for name in ("soffice", "libreoffice"):
        resolved = shutil.which(name)
        if resolved:
            return resolved
    raise LibreOfficeRecalculationError("LibreOffice is not installed or soffice is not available on PATH.")


def _set_xml_attribute(tag: str, name: str, value: str) -> str:
    pattern = rf'\b{re.escape(name)}="[^"]*"'
    replacement = f'{name}="{value}"'
    if re.search(pattern, tag):
        return re.sub(pattern, replacement, tag, count=1)
    return tag[:-2] + f" {replacement}/>"


def restore_calculation_flags(workbook_path: str | Path) -> None:
    path = Path(workbook_path)
    replacement_path = path.with_suffix(".metadata-restored.xlsx")

    with ZipFile(path, "r") as source:
        entries = [(item, source.read(item.filename)) for item in source.infolist()]

    workbook_entry = next(
        ((item, payload) for item, payload in entries if item.filename == WORKBOOK_XML),
        None,
    )
    if workbook_entry is None:
        raise LibreOfficeRecalculationError("Recalculated workbook does not contain xl/workbook.xml.")
    workbook_xml = workbook_entry[1].decode("utf-8")
    match = re.search(r"<calcPr\b[^>]*/>", workbook_xml)
    if not match:
        raise LibreOfficeRecalculationError("Recalculated workbook is missing xl/workbook.xml calcPr metadata.")
    tag = match.group(0)
    for name, value in (
        ("calcMode", "auto"),
        ("fullCalcOnLoad", "1"),
        ("forceFullCalc", "1"),
        ("iterate", "1"),
        ("iterateCount", "100"),
        ("iterateDelta", "0.0001"),
    ):
        tag = _set_xml_attribute(tag, name, value)
    updated_xml = (workbook_xml[: match.start()] + tag + workbook_xml[match.end() :]).encode("utf-8")

    with ZipFile(replacement_path, "w", compression=ZIP_DEFLATED) as destination:
        for item, payload in entries:
            if item.filename == WORKBOOK_XML:
                payload = updated_xml
            destination.writestr(item, payload)

    os.replace(replacement_path, path)


def _prepare_local_input(source: Path, destination: Path) -> None:
    workbook = openpyxl.load_workbook(source)
    if "Run Metadata" not in workbook.sheetnames:
        workbook.close()
        raise LibreOfficeRecalculationError("Workbook is missing the Run Metadata sheet.")
    workbook["Run Metadata"]["B2"] = "complete"
    workbook.calculation.calcMode = "auto"
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.iterate = True
    workbook.calculation.iterateCount = 100
    workbook.calculation.iterateDelta = 0.0001
    workbook.save(destination)
    workbook.close()


def recalculate_and_verify(
    *,
    excel_file_path: str | Path,
    normalized_inputs_path: str | Path,
    template_path: str | Path,
    output_directory: str | Path,
    executable: str | Path | None = None,
    timeout_seconds: float = 120.0,
) -> dict[str, object]:
    source = Path(excel_file_path)
    normalized = Path(normalized_inputs_path)
    template = Path(template_path)
    output = Path(output_directory)
    for path, label in (
        (source, "Workbook"),
        (normalized, "Normalized inputs"),
        (template, "Template"),
    ):
        if not path.is_file():
            raise LibreOfficeRecalculationError(f"{label} does not exist: {path}.")
    if timeout_seconds <= 0:
        raise LibreOfficeRecalculationError("timeout_seconds must be greater than zero.")

    output.mkdir(parents=True, exist_ok=True)
    snapshot = output / "valuation.recalculated.xlsx"
    verification_path = output / "recalculation-verification.json"
    receipt_path = output / "libreoffice-recalculation.json"
    for path in (snapshot, verification_path, receipt_path):
        if path.exists():
            raise FileExistsError(f"Refusing to overwrite existing artifact: {path}.")

    resolved_executable = _resolve_executable(executable)
    with tempfile.TemporaryDirectory(prefix=".libreoffice-", dir=output) as temporary:
        temporary_path = Path(temporary)
        local_input = temporary_path / "valuation.local-input.xlsx"
        converted_directory = temporary_path / "converted"
        profile_directory = temporary_path / "profile"
        converted_directory.mkdir()
        profile_directory.mkdir()
        _prepare_local_input(source, local_input)
        command = [
            resolved_executable,
            f"-env:UserInstallation={profile_directory.resolve().as_uri()}",
            "--headless",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(converted_directory),
            str(local_input),
        ]
        try:
            completed = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=timeout_seconds,
                check=False,
            )
        except subprocess.TimeoutExpired as error:
            raise LibreOfficeRecalculationError(
                f"LibreOffice recalculation exceeded {timeout_seconds:g} seconds."
            ) from error
        if completed.returncode != 0:
            detail = (completed.stderr or completed.stdout or "").strip()[-2000:]
            raise LibreOfficeRecalculationError(
                f"LibreOffice recalculation failed with exit code {completed.returncode}: {detail}"
            )

        converted = converted_directory / local_input.name
        if not converted.is_file() or converted.stat().st_size == 0:
            detail = (completed.stderr or completed.stdout or "").strip()[-2000:]
            raise LibreOfficeRecalculationError(f"LibreOffice did not create a recalculated XLSX snapshot. {detail}")
        os.replace(converted, snapshot)

    restore_calculation_flags(snapshot)
    verification = verify_recalculated(snapshot, normalized, template)
    atomic_write_json(verification_path, verification)
    receipt: dict[str, object] = {
        "status": "complete",
        "backend": "libreoffice",
        "executable": resolved_executable,
        "snapshot_path": str(snapshot.resolve()),
        "snapshot_sha256": file_sha256(snapshot),
        "verification_report": str(verification_path.resolve()),
        "outputs": verification["outputs"],
    }
    atomic_write_json(receipt_path, receipt)
    return receipt


def main() -> None:
    parser = argparse.ArgumentParser(description="Recalculate and verify a valuation workbook with LibreOffice")
    parser.add_argument("--file", required=True)
    parser.add_argument("--inputs", required=True)
    parser.add_argument("--template", required=True)
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--executable")
    parser.add_argument("--timeout-seconds", type=float, default=120.0)
    args = parser.parse_args()
    try:
        receipt = recalculate_and_verify(
            excel_file_path=args.file,
            normalized_inputs_path=args.inputs,
            template_path=args.template,
            output_directory=args.output_dir,
            executable=args.executable,
            timeout_seconds=args.timeout_seconds,
        )
    except (OSError, ValueError, LibreOfficeRecalculationError) as error:
        print(str(error), file=sys.stderr)
        raise SystemExit(1) from error
    print(json.dumps(receipt, indent=2))


if __name__ == "__main__":  # pragma: no cover - exercised through main()
    main()
