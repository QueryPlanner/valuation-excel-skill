from __future__ import annotations

import io
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import fill_excel
import get_financial_reports
import run_valuation
import upload_to_sheets
import validate_inputs
import verify_workbook
import workbook_contract
from validate_inputs import InputValidationError
from verify_workbook import WorkbookVerificationError

from tests.helpers import TEMPLATE, VALID_INPUTS


class CliEntrypointTests(unittest.TestCase):
    def test_workbook_contract_main(self) -> None:
        with (
            patch("sys.argv", ["workbook_contract.py", "--template", str(TEMPLATE)]),
            patch("sys.stdout", new_callable=io.StringIO) as stdout,
        ):
            workbook_contract.main()
        self.assertIn("template_sha256", stdout.getvalue())

    def test_validate_main_success_and_failure(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            output = Path(temporary) / "normalized.json"
            report = Path(temporary) / "report.json"
            argv = [
                "validate_inputs.py",
                "--inputs",
                str(VALID_INPUTS),
                "--template",
                str(TEMPLATE),
                "--output",
                str(output),
                "--report",
                str(report),
            ]
            with patch("sys.argv", argv), patch("sys.stdout", new_callable=io.StringIO):
                validate_inputs.main()
            self.assertTrue(output.is_file())
            self.assertEqual(json.loads(report.read_text())["status"], "valid")

            with (
                patch(
                    "validate_inputs.normalize_and_validate_inputs",
                    side_effect=InputValidationError(["bad"]),
                ),
                patch("sys.argv", argv),
                patch("sys.stderr", new_callable=io.StringIO),
                self.assertRaises(SystemExit) as context,
            ):
                validate_inputs.main()
            self.assertEqual(context.exception.code, 1)
            self.assertEqual(json.loads(report.read_text())["status"], "failed")

            no_report_argv = [
                "validate_inputs.py",
                "--inputs",
                str(VALID_INPUTS),
                "--template",
                str(TEMPLATE),
                "--output",
                str(output),
            ]
            with patch("sys.argv", no_report_argv), patch("sys.stdout", new_callable=io.StringIO):
                validate_inputs.main()
            with (
                patch(
                    "validate_inputs.normalize_and_validate_inputs",
                    side_effect=InputValidationError(["bad"]),
                ),
                patch("sys.argv", no_report_argv),
                patch("sys.stderr", new_callable=io.StringIO),
                self.assertRaises(SystemExit),
            ):
                validate_inputs.main()

    def test_fill_main_success_and_failure(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            output = Path(temporary) / "value.xlsx"
            receipt = Path(temporary) / "receipt.json"
            argv = [
                "fill_excel.py",
                "--company",
                "Example Foods",
                "--inputs",
                str(VALID_INPUTS),
                "--template",
                str(TEMPLATE),
                "--output",
                str(output),
                "--run-id",
                "cli",
                "--receipt",
                str(receipt),
            ]
            with patch("sys.argv", argv), patch("sys.stdout", new_callable=io.StringIO):
                fill_excel.main()
            self.assertTrue(output.is_file())
            self.assertTrue(receipt.is_file())

            with (
                patch(
                    "fill_excel.fill_valuation_excel",
                    side_effect=ValueError("bad"),
                ),
                patch("sys.argv", argv),
                patch("sys.stderr", new_callable=io.StringIO),
                self.assertRaises(SystemExit) as context,
            ):
                fill_excel.main()
            self.assertEqual(context.exception.code, 1)

    def test_verify_main_success_and_failure(self) -> None:
        success = {"stage": "precalculation", "status": "passed", "errors": []}
        argv = [
            "verify_workbook.py",
            "--file",
            "value.xlsx",
            "--inputs",
            "inputs.json",
            "--template",
            "template.xlsx",
            "--stage",
            "precalculation",
            "--output",
            "/tmp/verify-cli.json",
        ]
        with (
            patch("verify_workbook.verify_precalculation", return_value=success),
            patch("sys.argv", argv),
            patch("sys.stdout", new_callable=io.StringIO),
        ):
            verify_workbook.main()

        failure_report = {"stage": "recalculated", "status": "failed", "errors": ["bad"]}
        argv[argv.index("precalculation")] = "recalculated"
        with (
            patch(
                "verify_workbook.verify_recalculated",
                side_effect=WorkbookVerificationError(["bad"], failure_report),
            ),
            patch("sys.argv", argv),
            patch("sys.stderr", new_callable=io.StringIO),
            self.assertRaises(SystemExit) as context,
        ):
            verify_workbook.main()
        self.assertEqual(context.exception.code, 1)

    def test_google_main_auth_publish_and_failure(self) -> None:
        with (
            patch(
                "upload_to_sheets.credential_preflight",
                return_value={"status": "not_configured"},
            ),
            patch("sys.argv", ["upload_to_sheets.py", "auth-check"]),
            patch("sys.stdout", new_callable=io.StringIO),
        ):
            upload_to_sheets.main()

        publish_argv = [
            "upload_to_sheets.py",
            "publish",
            "--company",
            "Example",
            "--run-id",
            "run",
            "--file",
            "file.xlsx",
            "--inputs",
            "inputs.json",
            "--template",
            "template.xlsx",
            "--output-dir",
            "/tmp",
            "--folder-id",
            "folder",
        ]
        with (
            patch(
                "upload_to_sheets.publish_and_verify",
                return_value={"status": "complete"},
            ),
            patch("sys.argv", publish_argv),
            patch("sys.stdout", new_callable=io.StringIO),
        ):
            upload_to_sheets.main()

        with (
            patch(
                "upload_to_sheets.publish_and_verify",
                side_effect=RuntimeError("api failed"),
            ),
            patch("sys.argv", publish_argv),
            patch("sys.stderr", new_callable=io.StringIO),
            self.assertRaises(SystemExit) as context,
        ):
            upload_to_sheets.main()
        self.assertEqual(context.exception.code, 1)

    def test_sec_main_and_default_output(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            output = Path(temporary) / "manifest.json"
            argv = [
                "get_financial_reports.py",
                "--company",
                "ABC",
                "--output",
                str(output),
                "--sec-user-agent",
                "App test@example.com",
            ]
            with (
                patch(
                    "get_financial_reports.get_financial_reports",
                    return_value={"manifest_version": "2.0"},
                ),
                patch("sys.argv", argv),
                patch("sys.stdout", new_callable=io.StringIO),
            ):
                get_financial_reports.main()
            self.assertTrue(output.is_file())
            self.assertEqual(
                get_financial_reports._default_output_path("A B").name,
                "a_b_reports.v2.json",
            )

    def test_run_main_status_and_start_failure(self) -> None:
        argv = [
            "run_valuation.py",
            "--inputs",
            "inputs.json",
            "--template",
            "template.xlsx",
            "--output-root",
            "/tmp",
        ]
        with (
            patch(
                "run_valuation.run_valuation",
                return_value=({"status": "awaiting_recalculation"}, 2),
            ) as run,
            patch("sys.argv", argv),
            patch("sys.stdout", new_callable=io.StringIO),
            self.assertRaises(SystemExit) as context,
        ):
            run_valuation.main()
        self.assertEqual(context.exception.code, 2)
        self.assertEqual(run.call_args.kwargs["backend"], "libreoffice")

        with (
            patch(
                "run_valuation.run_valuation",
                side_effect=ValueError("bad"),
            ),
            patch("sys.argv", argv),
            patch("sys.stderr", new_callable=io.StringIO),
            self.assertRaises(SystemExit) as context,
        ):
            run_valuation.main()
        self.assertEqual(context.exception.code, 1)


if __name__ == "__main__":
    unittest.main()
