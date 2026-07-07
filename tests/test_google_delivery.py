from __future__ import annotations

import json
import os
import stat
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import upload_to_sheets
from common import atomic_write_json
from fill_excel import fill_valuation_excel
from upload_to_sheets import (
    GoogleDeliveryError,
    _load_credentials,
    credential_preflight,
    publish_and_verify,
)
from validate_inputs import normalize_and_validate_inputs
from workbook_contract import load_workbook_contract

from tests.helpers import TEMPLATE, inject_cached_values, load_valid_inputs


class Executable:
    def __init__(self, value=None, error: Exception | None = None):
        self.value = value
        self.error = error

    def execute(self):
        if self.error:
            raise self.error
        return self.value


class FakeFiles:
    def __init__(self, export_bytes: bytes):
        self.export_bytes = export_bytes
        self.created: list[dict] = []
        self.updated: list[dict] = []

    def create(self, **kwargs):
        self.created.append(kwargs)
        return Executable(
            {
                "id": "sheet-created",
                "webViewLink": "https://docs.google.com/spreadsheets/d/sheet-created",
                "name": kwargs["body"]["name"],
            }
        )

    def update(self, **kwargs):
        self.updated.append(kwargs)
        return Executable(
            {
                "id": kwargs["fileId"],
                "webViewLink": f"https://docs.google.com/spreadsheets/d/{kwargs['fileId']}",
                "name": "Existing",
            }
        )

    def export_media(self, **kwargs):
        return Executable(self.export_bytes)


class FakePermissions:
    def __init__(self):
        self.created: list[dict] = []

    def create(self, **kwargs):
        self.created.append(kwargs)
        return Executable({})


class FakeDrive:
    def __init__(self, export_bytes: bytes):
        self.file_resource = FakeFiles(export_bytes)
        self.permission_resource = FakePermissions()

    def files(self):
        return self.file_resource

    def permissions(self):
        return self.permission_resource


class FakeValues:
    def __init__(self):
        self.updates: list[dict] = []

    def update(self, **kwargs):
        self.updates.append(kwargs)
        return Executable({})


class FakeSpreadsheets:
    def __init__(self):
        self.batch_updates: list[dict] = []
        self.value_resource = FakeValues()

    def batchUpdate(self, **kwargs):
        self.batch_updates.append(kwargs)
        return Executable({})

    def values(self):
        return self.value_resource


class FakeSheets:
    def __init__(self):
        self.spreadsheet_resource = FakeSpreadsheets()

    def spreadsheets(self):
        return self.spreadsheet_resource


class GoogleDeliveryTests(unittest.TestCase):
    def _build_files(self, directory: Path) -> tuple[Path, Path, bytes]:
        contract = load_workbook_contract(TEMPLATE)
        normalized = normalize_and_validate_inputs(load_valid_inputs(), contract)
        inputs_path = directory / "normalized.json"
        atomic_write_json(inputs_path, normalized)
        populated = directory / "populated.xlsx"
        fill_valuation_excel(
            "Example Foods",
            inputs_path,
            TEMPLATE,
            populated,
            run_id="google-test",
        )
        recalculated = directory / "recalculated.xlsx"
        inject_cached_values(
            populated,
            recalculated,
            {
                "Cost of capital worksheet": {"B13": (0.08, None)},
                "Valuation output": {
                    "B24": (1500.0, None),
                    "B31": (1200.0, None),
                    "B33": (12.0, None),
                },
            },
        )
        return inputs_path, populated, recalculated.read_bytes()

    def test_create_configure_export_verify_and_receipt(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary)
            inputs_path, populated, export_bytes = self._build_files(directory)
            drive = FakeDrive(export_bytes)
            sheets = FakeSheets()
            receipt = publish_and_verify(
                company_name="Example Foods",
                run_id="google-test",
                excel_file_path=populated,
                normalized_inputs_path=inputs_path,
                template_path=TEMPLATE,
                output_directory=directory,
                folder_id="folder-1",
                share_with="reviewer@example.com",
                drive_service=drive,
                sheets_service=sheets,
                poll_delay_seconds=0,
            )
            self.assertEqual(receipt["status"], "complete")
            self.assertEqual(receipt["delivery_mode"], "create")
            create_body = drive.file_resource.created[0]["body"]
            self.assertEqual(create_body["parents"], ["folder-1"])
            self.assertTrue(sheets.spreadsheet_resource.batch_updates)
            self.assertEqual(len(sheets.spreadsheet_resource.value_resource.updates), 2)
            self.assertEqual(
                drive.permission_resource.created[0]["body"]["emailAddress"],
                "reviewer@example.com",
            )
            self.assertTrue(Path(receipt["snapshot_path"]).is_file())
            self.assertTrue((directory / "google-test.google-delivery.json").is_file())

    def test_replace_requires_explicit_file_id_and_flag(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary)
            inputs_path, populated, export_bytes = self._build_files(directory)
            drive = FakeDrive(export_bytes)
            sheets = FakeSheets()
            base = {
                "company_name": "Example Foods",
                "run_id": "google-test",
                "excel_file_path": populated,
                "normalized_inputs_path": inputs_path,
                "template_path": TEMPLATE,
                "output_directory": directory,
                "folder_id": "folder-1",
                "drive_service": drive,
                "sheets_service": sheets,
                "poll_delay_seconds": 0,
            }
            with self.assertRaises(GoogleDeliveryError):
                publish_and_verify(**base, file_id="existing")
            with self.assertRaises(GoogleDeliveryError):
                publish_and_verify(**base, replace=True)
            receipt = publish_and_verify(
                **base,
                file_id="existing",
                replace=True,
            )
            self.assertEqual(receipt["delivery_mode"], "replace")
            self.assertEqual(drive.file_resource.updated[0]["fileId"], "existing")

    def test_delivery_rejects_bad_arguments_and_empty_export(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary)
            inputs_path, populated, _ = self._build_files(directory)
            with self.assertRaises(GoogleDeliveryError):
                publish_and_verify(
                    company_name="Example",
                    run_id="run",
                    excel_file_path=directory / "missing.xlsx",
                    normalized_inputs_path=inputs_path,
                    template_path=TEMPLATE,
                    output_directory=directory,
                    folder_id="folder",
                    drive_service=FakeDrive(b""),
                    sheets_service=FakeSheets(),
                )
            with self.assertRaises(GoogleDeliveryError):
                publish_and_verify(
                    company_name="Example",
                    run_id="run",
                    excel_file_path=populated,
                    normalized_inputs_path=inputs_path,
                    template_path=TEMPLATE,
                    output_directory=directory,
                    folder_id="",
                    drive_service=FakeDrive(b""),
                    sheets_service=FakeSheets(),
                )
            with self.assertRaises(GoogleDeliveryError):
                publish_and_verify(
                    company_name="Example",
                    run_id="run",
                    excel_file_path=populated,
                    normalized_inputs_path=inputs_path,
                    template_path=TEMPLATE,
                    output_directory=directory,
                    folder_id="folder",
                    drive_service=FakeDrive(b""),
                    sheets_service=FakeSheets(),
                    poll_attempts=0,
                )
            with self.assertRaises(GoogleDeliveryError):
                publish_and_verify(
                    company_name="Example",
                    run_id="run",
                    excel_file_path=populated,
                    normalized_inputs_path=inputs_path,
                    template_path=TEMPLATE,
                    output_directory=directory,
                    folder_id="folder",
                    drive_service=FakeDrive(b""),
                    sheets_service=FakeSheets(),
                )

    def test_credential_preflight_without_logging_secrets(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary)
            token = directory / "token.json"
            token.write_text(json.dumps({"refresh_token": "secret"}), encoding="utf-8")
            receipt = credential_preflight(token_path=token)
            self.assertEqual(receipt["credential_type"], "oauth_token")
            self.assertNotIn("secret", json.dumps(receipt))

            bad = directory / "bad.json"
            bad.write_text("{}", encoding="utf-8")
            with self.assertRaises(GoogleDeliveryError):
                credential_preflight(token_path=bad)
            with self.assertRaises(GoogleDeliveryError):
                credential_preflight(service_account_path=directory / "missing.json")

            service = directory / "service.json"
            service.write_text(
                json.dumps(
                    {
                        "type": "service_account",
                        "client_email": "service@example.com",
                        "private_key": "private",
                        "token_uri": "https://oauth2.googleapis.com/token",
                    }
                ),
                encoding="utf-8",
            )
            with patch(
                "google.oauth2.service_account.Credentials.from_service_account_info",
                return_value=Mock(),
            ):
                receipt = credential_preflight(service_account_path=service)
            self.assertEqual(receipt["credential_type"], "service_account")

    def test_load_credentials_requires_configuration(self) -> None:
        with self.assertRaises(GoogleDeliveryError):
            _load_credentials(allow_adc=False)
        fake_credentials = Mock()
        with patch("google.auth.default", return_value=(fake_credentials, "project")):
            credentials, credential_type = _load_credentials(allow_adc=True)
        self.assertIs(credentials, fake_credentials)
        self.assertEqual(credential_type, "adc")

    def test_credential_preflight_rejects_each_invalid_file_shape(self) -> None:
        with tempfile.TemporaryDirectory() as temporary, patch.dict(os.environ, {}, clear=True):
            directory = Path(temporary)
            self.assertEqual(credential_preflight()["status"], "not_configured")

            missing_token = directory / "missing-token.json"
            with self.assertRaisesRegex(GoogleDeliveryError, "GOOGLE_TOKEN_PATH"):
                credential_preflight(token_path=missing_token)

            invalid_service = directory / "invalid-service.json"
            invalid_service.write_text("{", encoding="utf-8")
            with self.assertRaisesRegex(GoogleDeliveryError, "not valid JSON"):
                credential_preflight(service_account_path=invalid_service)

            invalid_token = directory / "invalid-token.json"
            invalid_token.write_text("{", encoding="utf-8")
            with self.assertRaisesRegex(GoogleDeliveryError, "not valid JSON"):
                credential_preflight(token_path=invalid_token)

            wrong_service = directory / "wrong-service.json"
            wrong_service.write_text(
                json.dumps(
                    {
                        "type": "authorized_user",
                        "client_email": "service@example.com",
                        "private_key": "private",
                        "token_uri": "https://oauth2.googleapis.com/token",
                    }
                ),
                encoding="utf-8",
            )
            with self.assertRaisesRegex(GoogleDeliveryError, "missing required"):
                credential_preflight(service_account_path=wrong_service)

            empty_token = directory / "empty-token.json"
            empty_token.write_text("{}", encoding="utf-8")
            with self.assertRaisesRegex(GoogleDeliveryError, "usable token"):
                credential_preflight(token_path=empty_token)

            unreadable = directory / "unreadable.json"
            unreadable.write_text("{}", encoding="utf-8")
            with patch.object(Path, "read_text", side_effect=OSError("unreadable")), self.assertRaisesRegex(
                GoogleDeliveryError, "not valid JSON"
            ):
                credential_preflight(service_account_path=unreadable)
            with patch.object(Path, "read_text", side_effect=OSError("unreadable")), self.assertRaisesRegex(
                GoogleDeliveryError, "not valid JSON"
            ):
                credential_preflight(token_path=unreadable)

    def test_credential_preflight_rejects_unparseable_service_key(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            service = Path(temporary) / "service.json"
            service.write_text(
                json.dumps(
                    {
                        "type": "service_account",
                        "client_email": "service@example.com",
                        "private_key": "not-a-private-key",
                        "token_uri": "https://oauth2.googleapis.com/token",
                    }
                ),
                encoding="utf-8",
            )
            with patch(
                "google.oauth2.service_account.Credentials.from_service_account_info",
                side_effect=ValueError("bad key"),
            ), self.assertRaisesRegex(GoogleDeliveryError, "invalid or revoked"):
                credential_preflight(service_account_path=service)

    def test_load_service_account_and_oauth_credentials(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary)
            service = directory / "service.json"
            service.write_text("{}", encoding="utf-8")
            service_credentials = Mock()
            with patch(
                "upload_to_sheets.credential_preflight",
                return_value={"status": "configured"},
            ), patch(
                "google.oauth2.service_account.Credentials.from_service_account_file",
                return_value=service_credentials,
            ) as from_file:
                credentials, credential_type = _load_credentials(
                    service_account_path=service,
                    allow_adc=False,
                )
            self.assertIs(credentials, service_credentials)
            self.assertEqual(credential_type, "service_account")
            from_file.assert_called_once()

            token = directory / "token.json"
            token.write_text("{}", encoding="utf-8")
            oauth_credentials = Mock()
            oauth_credentials.expired = True
            oauth_credentials.refresh_token = "refresh"
            oauth_credentials.valid = True
            with patch(
                "upload_to_sheets.credential_preflight",
                return_value={"status": "configured"},
            ), patch(
                "google.oauth2.credentials.Credentials.from_authorized_user_file",
                return_value=oauth_credentials,
            ), patch("google.auth.transport.requests.Request", return_value="request"):
                credentials, credential_type = _load_credentials(
                    token_path=token,
                    allow_adc=False,
                )
            self.assertIs(credentials, oauth_credentials)
            self.assertEqual(credential_type, "oauth_token")
            oauth_credentials.refresh.assert_called_once_with("request")

            oauth_credentials.expired = False
            oauth_credentials.valid = False
            with patch(
                "upload_to_sheets.credential_preflight",
                return_value={"status": "configured"},
            ), patch(
                "google.oauth2.credentials.Credentials.from_authorized_user_file",
                return_value=oauth_credentials,
            ), self.assertRaisesRegex(GoogleDeliveryError, "not valid"):
                _load_credentials(token_path=token, allow_adc=False)

    def test_load_credentials_adc_fallback_and_interactive_auth(self) -> None:
        with patch.dict(os.environ, {}, clear=True), patch(
            "google.auth.default",
            side_effect=RuntimeError("ADC unavailable"),
        ), self.assertRaisesRegex(GoogleDeliveryError, "No Google credentials"):
            _load_credentials(allow_adc=True)

        with patch.dict(os.environ, {}, clear=True), patch(
            "google.auth.default",
            return_value=(None, None),
        ), self.assertRaisesRegex(GoogleDeliveryError, "No Google credentials"):
            _load_credentials(allow_adc=True)

        with patch.dict(os.environ, {}, clear=True), self.assertRaisesRegex(
            GoogleDeliveryError, "GOOGLE_CLIENT_SECRET_PATH"
        ):
            _load_credentials(allow_adc=False, interactive_auth=True)

        with tempfile.TemporaryDirectory() as temporary, patch.dict(os.environ, {}, clear=True):
            directory = Path(temporary)
            secret = directory / "client.json"
            secret.write_text("{}", encoding="utf-8")
            token = directory / "new-token.json"
            interactive_credentials = Mock()
            interactive_credentials.to_json.return_value = '{"token":"generated"}'
            flow = Mock()
            flow.run_local_server.return_value = interactive_credentials
            with patch(
                "google_auth_oauthlib.flow.InstalledAppFlow.from_client_secrets_file",
                return_value=flow,
            ):
                credentials, credential_type = _load_credentials(
                    allow_adc=False,
                    interactive_auth=True,
                    client_secret_path=secret,
                    token_path=token,
                )
            self.assertIs(credentials, interactive_credentials)
            self.assertEqual(credential_type, "interactive")
            self.assertEqual(stat.S_IMODE(token.stat().st_mode), 0o600)
            flow.run_local_server.assert_called_once_with(port=0, open_browser=False)

    def test_publish_builds_missing_services_with_loaded_credentials(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary)
            inputs_path, populated, export_bytes = self._build_files(directory)
            drive = FakeDrive(export_bytes)
            sheets = FakeSheets()
            credentials = Mock()
            media = Mock()
            with patch(
                "upload_to_sheets._load_credentials",
                return_value=(credentials, "adc"),
            ), patch(
                "googleapiclient.discovery.build",
                return_value=drive,
            ) as build, patch(
                "googleapiclient.http.MediaFileUpload",
                return_value=media,
            ):
                receipt = publish_and_verify(
                    company_name="Example Foods",
                    run_id="credentials-drive",
                    excel_file_path=populated,
                    normalized_inputs_path=inputs_path,
                    template_path=TEMPLATE,
                    output_directory=directory,
                    folder_id="folder",
                    drive_service=None,
                    sheets_service=sheets,
                    poll_delay_seconds=0,
                )
            self.assertEqual(receipt["credential_type"], "adc")
            self.assertIs(drive.file_resource.created[0]["media_body"], media)
            build.assert_called_once_with(
                "drive",
                "v3",
                credentials=credentials,
                cache_discovery=False,
            )

            second_drive = FakeDrive(export_bytes)
            second_sheets = FakeSheets()
            with patch(
                "upload_to_sheets._load_credentials",
                return_value=(credentials, "service_account"),
            ), patch(
                "googleapiclient.discovery.build",
                return_value=second_sheets,
            ) as build, patch(
                "googleapiclient.http.MediaFileUpload",
                return_value=media,
            ):
                receipt = publish_and_verify(
                    company_name="Example Foods",
                    run_id="credentials-sheets",
                    excel_file_path=populated,
                    normalized_inputs_path=inputs_path,
                    template_path=TEMPLATE,
                    output_directory=directory,
                    folder_id="folder",
                    drive_service=second_drive,
                    sheets_service=None,
                    poll_delay_seconds=0,
                )
            self.assertEqual(receipt["credential_type"], "service_account")
            build.assert_called_once_with(
                "sheets",
                "v4",
                credentials=credentials,
                cache_discovery=False,
            )

    def test_publish_retries_verification_and_reports_final_failure(self) -> None:
        failed = upload_to_sheets.WorkbookVerificationError(
            ["not recalculated"],
            {"status": "failed"},
        )
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary)
            inputs_path, populated, export_bytes = self._build_files(directory)
            drive = FakeDrive(export_bytes)
            sheets = FakeSheets()
            with patch(
                "upload_to_sheets.verify_recalculated",
                side_effect=[failed, {"status": "passed"}],
            ), patch("upload_to_sheets.time.sleep") as sleep:
                receipt = publish_and_verify(
                    company_name="Example Foods",
                    run_id="retry",
                    excel_file_path=populated,
                    normalized_inputs_path=inputs_path,
                    template_path=TEMPLATE,
                    output_directory=directory,
                    folder_id="folder",
                    drive_service=drive,
                    sheets_service=sheets,
                    poll_attempts=2,
                    poll_delay_seconds=0.01,
                )
            self.assertEqual(receipt["status"], "complete")
            sleep.assert_called_once_with(0.01)

            with patch(
                "upload_to_sheets.verify_recalculated",
                side_effect=failed,
            ), self.assertRaisesRegex(GoogleDeliveryError, "did not pass verification"):
                publish_and_verify(
                    company_name="Example Foods",
                    run_id="failure",
                    excel_file_path=populated,
                    normalized_inputs_path=inputs_path,
                    template_path=TEMPLATE,
                    output_directory=directory,
                    folder_id="folder",
                    drive_service=FakeDrive(export_bytes),
                    sheets_service=FakeSheets(),
                    poll_attempts=1,
                    poll_delay_seconds=0,
                )

    def test_publish_rejects_nonbyte_export_and_missing_created_id(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary)
            inputs_path, populated, _ = self._build_files(directory)
            with self.assertRaisesRegex(GoogleDeliveryError, "empty spreadsheet export"):
                publish_and_verify(
                    company_name="Example Foods",
                    run_id="nonbytes",
                    excel_file_path=populated,
                    normalized_inputs_path=inputs_path,
                    template_path=TEMPLATE,
                    output_directory=directory,
                    folder_id="folder",
                    drive_service=FakeDrive("not bytes"),
                    sheets_service=FakeSheets(),
                    poll_attempts=1,
                    poll_delay_seconds=0,
                )

            drive = FakeDrive(b"unused")
            drive.file_resource.create = Mock(
                return_value=Executable(
                    {
                        "webViewLink": "https://docs.google.com/spreadsheets/d/no-id",
                        "name": "No ID",
                    }
                )
            )
            with self.assertRaisesRegex(GoogleDeliveryError, "did not return a file ID"):
                publish_and_verify(
                    company_name="Example Foods",
                    run_id="no-id",
                    excel_file_path=populated,
                    normalized_inputs_path=inputs_path,
                    template_path=TEMPLATE,
                    output_directory=directory,
                    folder_id="folder",
                    drive_service=drive,
                    sheets_service=FakeSheets(),
                    poll_attempts=1,
                    poll_delay_seconds=0,
                )


if __name__ == "__main__":
    unittest.main()
