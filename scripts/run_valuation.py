from __future__ import annotations

import argparse
import json
import sys
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from common import atomic_write_json, file_sha256
from fill_excel import fill_valuation_excel
from recalculate_with_libreoffice import recalculate_and_verify
from upload_to_sheets import publish_and_verify
from validate_inputs import InputValidationError, normalize_and_validate_inputs
from verify_workbook import verify_precalculation
from workbook_contract import load_workbook_contract


def _new_run_id() -> str:
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    return f"{timestamp}-{uuid.uuid4().hex[:12]}"


def _load_json(path: str | Path) -> dict[str, Any]:
    with Path(path).open(encoding="utf-8") as source:
        value = json.load(source)
    if not isinstance(value, dict):
        raise ValueError(f"{path} must contain a JSON object.")
    return value


def run_valuation(
    *,
    inputs_path: str | Path,
    template_path: str | Path,
    output_root: str | Path,
    backend: str,
    allow_legacy: bool = False,
    google_folder_id: str | None = None,
    google_file_id: str | None = None,
    replace_google_file: bool = False,
    share_with: str | None = None,
    service_account_path: str | Path | None = None,
    token_path: str | Path | None = None,
    allow_adc: bool = True,
    libreoffice_executable: str | Path | None = None,
    libreoffice_timeout_seconds: float = 120.0,
    run_id: str | None = None,
) -> tuple[dict[str, Any], int]:
    if backend not in {"none", "libreoffice", "google"}:
        raise ValueError("backend must be 'none', 'libreoffice', or 'google'.")
    resolved_run_id = run_id or _new_run_id()
    run_directory = Path(output_root) / resolved_run_id
    run_directory.mkdir(parents=True, exist_ok=False)
    manifest_path = run_directory / "run.json"
    manifest: dict[str, Any] = {
        "run_id": resolved_run_id,
        "status": "running",
        "backend": backend,
        "started_at": datetime.now(timezone.utc).isoformat(),
        "input_path": str(Path(inputs_path).resolve()),
        "input_sha256": file_sha256(inputs_path),
        "template_path": str(Path(template_path).resolve()),
        "template_sha256": file_sha256(template_path),
        "steps": {
            "validation": {"status": "pending"},
            "workbook_build": {"status": "pending"},
            "precalculation_verification": {"status": "pending"},
            "recalculation_delivery": {"status": "pending"},
        },
        "artifacts": {},
        "errors": [],
    }
    atomic_write_json(manifest_path, manifest)

    try:
        raw_inputs = _load_json(inputs_path)
        contract = load_workbook_contract(template_path)
        normalized = normalize_and_validate_inputs(
            raw_inputs,
            contract,
            allow_legacy=allow_legacy,
        )
        normalized_path = run_directory / "inputs.normalized.json"
        atomic_write_json(normalized_path, normalized)
        manifest["steps"]["validation"] = {
            "status": "passed",
            "findings": normalized["validation_findings"],
        }
        manifest["artifacts"]["normalized_inputs"] = str(normalized_path.resolve())
        atomic_write_json(manifest_path, manifest)

        workbook_path = run_directory / "valuation.awaiting-recalculation.xlsx"
        build_receipt = fill_valuation_excel(
            normalized["company_context"]["company_name"],
            normalized_path,
            template_path,
            workbook_path,
            run_id=resolved_run_id,
        )
        manifest["steps"]["workbook_build"] = {"status": "passed", "receipt": build_receipt}
        manifest["artifacts"]["workbook"] = str(workbook_path.resolve())
        atomic_write_json(manifest_path, manifest)

        preflight = verify_precalculation(workbook_path, normalized_path, template_path)
        preflight_path = run_directory / "precalculation-verification.json"
        atomic_write_json(preflight_path, preflight)
        manifest["steps"]["precalculation_verification"] = {
            "status": "passed",
            "report": str(preflight_path.resolve()),
        }
        manifest["artifacts"]["precalculation_report"] = str(preflight_path.resolve())
        atomic_write_json(manifest_path, manifest)

        if backend == "none":
            manifest["status"] = "awaiting_recalculation"
            manifest["steps"]["recalculation_delivery"] = {
                "status": "not_run",
                "reason": "No calculation backend was selected.",
            }
            exit_code = 2
        elif backend == "libreoffice":
            recalculation = recalculate_and_verify(
                excel_file_path=workbook_path,
                normalized_inputs_path=normalized_path,
                template_path=template_path,
                output_directory=run_directory,
                executable=libreoffice_executable,
                timeout_seconds=libreoffice_timeout_seconds,
            )
            manifest["status"] = "complete"
            manifest["steps"]["recalculation_delivery"] = {
                "status": "passed",
                "receipt": recalculation,
            }
            manifest["artifacts"]["recalculated_workbook"] = recalculation["snapshot_path"]
            manifest["artifacts"]["recalculation_report"] = recalculation["verification_report"]
            exit_code = 0
        else:
            if not google_folder_id:
                raise ValueError("google_folder_id is required for the Google calculation backend.")
            delivery = publish_and_verify(
                company_name=normalized["company_context"]["company_name"],
                run_id=resolved_run_id,
                excel_file_path=workbook_path,
                normalized_inputs_path=normalized_path,
                template_path=template_path,
                output_directory=run_directory,
                folder_id=google_folder_id,
                file_id=google_file_id,
                replace=replace_google_file,
                share_with=share_with,
                service_account_path=service_account_path,
                token_path=token_path,
                allow_adc=allow_adc,
            )
            manifest["status"] = "complete"
            manifest["steps"]["recalculation_delivery"] = {
                "status": "passed",
                "receipt": delivery,
            }
            manifest["artifacts"]["recalculated_workbook"] = delivery["snapshot_path"]
            manifest["artifacts"]["google_sheet"] = delivery["web_view_link"]
            exit_code = 0
    except Exception as error:
        manifest["status"] = "failed"
        manifest["errors"].append(
            {
                "type": type(error).__name__,
                "message": str(error),
            }
        )
        for step in manifest["steps"].values():
            if step["status"] == "pending":
                step["status"] = "not_run"
        exit_code = 1

    manifest["finished_at"] = datetime.now(timezone.utc).isoformat()
    atomic_write_json(manifest_path, manifest)
    return manifest, exit_code


def main() -> None:
    parser = argparse.ArgumentParser(description="Run the reliable v2 valuation workflow")
    parser.add_argument("--inputs", required=True)
    parser.add_argument("--template", required=True)
    parser.add_argument("--output-root", required=True)
    parser.add_argument(
        "--backend",
        choices=("none", "libreoffice", "google"),
        default="libreoffice",
    )
    parser.add_argument("--allow-legacy", action="store_true")
    parser.add_argument("--google-folder-id")
    parser.add_argument("--google-file-id")
    parser.add_argument("--replace-google-file", action="store_true")
    parser.add_argument("--share-with")
    parser.add_argument("--service-account-path")
    parser.add_argument("--token-path")
    parser.add_argument("--no-adc", action="store_true")
    parser.add_argument("--libreoffice-executable")
    parser.add_argument("--libreoffice-timeout-seconds", type=float, default=120.0)
    parser.add_argument("--run-id")
    args = parser.parse_args()
    try:
        manifest, exit_code = run_valuation(
            inputs_path=args.inputs,
            template_path=args.template,
            output_root=args.output_root,
            backend=args.backend,
            allow_legacy=args.allow_legacy,
            google_folder_id=args.google_folder_id,
            google_file_id=args.google_file_id,
            replace_google_file=args.replace_google_file,
            share_with=args.share_with,
            service_account_path=args.service_account_path,
            token_path=args.token_path,
            allow_adc=not args.no_adc,
            libreoffice_executable=args.libreoffice_executable,
            libreoffice_timeout_seconds=args.libreoffice_timeout_seconds,
            run_id=args.run_id,
        )
    except (FileExistsError, InputValidationError, ValueError) as error:
        print(f"Unable to start valuation run: {error}", file=sys.stderr)
        raise SystemExit(1) from error
    print(json.dumps(manifest, indent=2))
    raise SystemExit(exit_code)


if __name__ == "__main__":  # pragma: no cover - exercised through main()
    main()
