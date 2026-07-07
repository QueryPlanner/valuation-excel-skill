from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from run_valuation import _load_json, _new_run_id, run_valuation

from tests.helpers import TEMPLATE, VALID_INPUTS, load_valid_inputs, write_json


class RunValuationTests(unittest.TestCase):
    def test_run_id_and_json_loader_helpers(self) -> None:
        self.assertRegex(
            _new_run_id(),
            r"^\d{8}T\d{6}Z-[0-9a-f]{12}$",
        )
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "array.json"
            path.write_text("[]", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "JSON object"):
                _load_json(path)

    def test_build_only_persists_incomplete_manifest(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            output_root = Path(temporary) / "runs"
            manifest, exit_code = run_valuation(
                inputs_path=VALID_INPUTS,
                template_path=TEMPLATE,
                output_root=output_root,
                backend="none",
                run_id="build-only",
            )
            self.assertEqual(exit_code, 2)
            self.assertEqual(manifest["status"], "awaiting_recalculation")
            self.assertEqual(
                manifest["steps"]["recalculation_delivery"]["status"],
                "not_run",
            )
            persisted = json.loads((output_root / "build-only" / "run.json").read_text(encoding="utf-8"))
            self.assertEqual(persisted["status"], "awaiting_recalculation")
            self.assertTrue(Path(persisted["artifacts"]["workbook"]).is_file())

    def test_google_backend_completes_only_with_delivery_receipt(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            output_root = Path(temporary) / "runs"
            fake_delivery = {
                "status": "complete",
                "snapshot_path": "/tmp/recalculated.xlsx",
                "web_view_link": "https://docs.google.com/spreadsheets/d/test",
            }
            with patch("run_valuation.publish_and_verify", return_value=fake_delivery) as publish:
                manifest, exit_code = run_valuation(
                    inputs_path=VALID_INPUTS,
                    template_path=TEMPLATE,
                    output_root=output_root,
                    backend="google",
                    google_folder_id="folder",
                    run_id="google-complete",
                )
            self.assertEqual(exit_code, 0)
            self.assertEqual(manifest["status"], "complete")
            self.assertEqual(
                manifest["artifacts"]["google_sheet"],
                fake_delivery["web_view_link"],
            )
            publish.assert_called_once()

    def test_libreoffice_backend_completes_only_after_verification(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            output_root = Path(temporary) / "runs"
            fake_receipt = {
                "status": "complete",
                "snapshot_path": "/tmp/valuation.recalculated.xlsx",
                "verification_report": "/tmp/recalculation-verification.json",
                "outputs": {"value_per_share": 12.0},
            }
            with patch(
                "run_valuation.recalculate_and_verify",
                return_value=fake_receipt,
            ) as recalculate:
                manifest, exit_code = run_valuation(
                    inputs_path=VALID_INPUTS,
                    template_path=TEMPLATE,
                    output_root=output_root,
                    backend="libreoffice",
                    libreoffice_executable="/Applications/LibreOffice.app/soffice",
                    libreoffice_timeout_seconds=45,
                    run_id="libreoffice-complete",
                )
            self.assertEqual(exit_code, 0)
            self.assertEqual(manifest["status"], "complete")
            self.assertEqual(
                manifest["artifacts"]["recalculated_workbook"],
                fake_receipt["snapshot_path"],
            )
            self.assertEqual(
                manifest["artifacts"]["recalculation_report"],
                fake_receipt["verification_report"],
            )
            recalculate.assert_called_once()
            self.assertEqual(recalculate.call_args.kwargs["timeout_seconds"], 45)

    def test_failure_is_persisted_and_duplicate_run_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary)
            bad_inputs = load_valid_inputs()
            bad_inputs["schema_version"] = "wrong"
            bad_path = directory / "bad.json"
            write_json(bad_path, bad_inputs)
            output_root = directory / "runs"
            manifest, exit_code = run_valuation(
                inputs_path=bad_path,
                template_path=TEMPLATE,
                output_root=output_root,
                backend="none",
                run_id="failed",
            )
            self.assertEqual(exit_code, 1)
            self.assertEqual(manifest["status"], "failed")
            self.assertEqual(manifest["errors"][0]["type"], "InputValidationError")

            with self.assertRaises(FileExistsError):
                run_valuation(
                    inputs_path=VALID_INPUTS,
                    template_path=TEMPLATE,
                    output_root=output_root,
                    backend="none",
                    run_id="failed",
                )

    def test_invalid_backend_and_missing_google_folder(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            with self.assertRaises(ValueError):
                run_valuation(
                    inputs_path=VALID_INPUTS,
                    template_path=TEMPLATE,
                    output_root=temporary,
                    backend="invalid",
                )
            manifest, exit_code = run_valuation(
                inputs_path=VALID_INPUTS,
                template_path=TEMPLATE,
                output_root=temporary,
                backend="google",
                run_id="missing-folder",
            )
            self.assertEqual(exit_code, 1)
            self.assertIn("google_folder_id", manifest["errors"][0]["message"])


if __name__ == "__main__":
    unittest.main()
