from __future__ import annotations

import argparse
import json
import os
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from common import atomic_write_json, file_sha256
from verify_workbook import WorkbookVerificationError, verify_recalculated

SCOPES = ["https://www.googleapis.com/auth/drive.file"]
GOOGLE_SHEET_MIME_TYPE = "application/vnd.google-apps.spreadsheet"
EXCEL_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class GoogleDeliveryError(RuntimeError):
    pass


def credential_preflight(
    *,
    service_account_path: str | Path | None = None,
    token_path: str | Path | None = None,
) -> dict[str, Any]:
    service_path = Path(
        service_account_path or os.environ.get("GOOGLE_SERVICE_ACCOUNT_PATH", "")
    ).expanduser()
    if str(service_path) not in {"", "."}:
        if not service_path.is_file():
            raise GoogleDeliveryError("GOOGLE_SERVICE_ACCOUNT_PATH does not point to a readable file.")
        try:
            payload = json.loads(service_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as error:
            raise GoogleDeliveryError("The service-account file is not valid JSON.") from error
        required = {"type", "client_email", "private_key", "token_uri"}
        missing = sorted(required - set(payload))
        if missing or payload.get("type") != "service_account":
            raise GoogleDeliveryError(
                "The service-account file is missing required Google credential fields."
            )
        try:
            from google.oauth2 import service_account

            service_account.Credentials.from_service_account_info(payload, scopes=SCOPES)
        except Exception as error:
            raise GoogleDeliveryError(
                "The service-account key could not be parsed. The key may be invalid or revoked."
            ) from error
        return {
            "status": "configured",
            "credential_type": "service_account",
            "path_exists": True,
            "requires_writable_folder": True,
        }

    resolved_token = Path(token_path or os.environ.get("GOOGLE_TOKEN_PATH", "")).expanduser()
    if str(resolved_token) not in {"", "."}:
        if not resolved_token.is_file():
            raise GoogleDeliveryError("GOOGLE_TOKEN_PATH does not point to a readable file.")
        try:
            payload = json.loads(resolved_token.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as error:
            raise GoogleDeliveryError("The OAuth token file is not valid JSON.") from error
        if not payload.get("token") and not payload.get("refresh_token"):
            raise GoogleDeliveryError("The OAuth token file does not contain usable token fields.")
        return {
            "status": "configured",
            "credential_type": "oauth_token",
            "path_exists": True,
            "requires_writable_folder": True,
        }

    return {
        "status": "not_configured",
        "credential_type": None,
        "path_exists": False,
        "requires_writable_folder": True,
    }


def _load_credentials(
    *,
    service_account_path: str | Path | None = None,
    token_path: str | Path | None = None,
    allow_adc: bool = True,
    interactive_auth: bool = False,
    client_secret_path: str | Path | None = None,
) -> tuple[Any, str]:
    service_path_text = service_account_path or os.environ.get("GOOGLE_SERVICE_ACCOUNT_PATH")
    if service_path_text:
        from google.oauth2 import service_account

        service_path = Path(service_path_text).expanduser()
        credential_preflight(service_account_path=service_path)
        return (
            service_account.Credentials.from_service_account_file(service_path, scopes=SCOPES),
            "service_account",
        )

    token_path_text = token_path or os.environ.get("GOOGLE_TOKEN_PATH")
    if token_path_text:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials

        resolved_token = Path(token_path_text).expanduser()
        if resolved_token.is_file() or not interactive_auth:
            credential_preflight(token_path=resolved_token)
            credentials = Credentials.from_authorized_user_file(resolved_token, SCOPES)
            if credentials.expired and credentials.refresh_token:
                credentials.refresh(Request())
            if not credentials.valid:
                raise GoogleDeliveryError("OAuth credentials are not valid.")
            return credentials, "oauth_token"

    if allow_adc:
        try:
            import google.auth

            credentials, _ = google.auth.default(scopes=SCOPES)
            if credentials is not None:
                return credentials, "adc"
        except Exception:
            pass

    if not interactive_auth:
        raise GoogleDeliveryError(
            "No Google credentials are configured. Set GOOGLE_SERVICE_ACCOUNT_PATH or "
            "GOOGLE_TOKEN_PATH, configure ADC, or explicitly enable interactive authentication."
        )

    secret_path = Path(
        client_secret_path or os.environ.get("GOOGLE_CLIENT_SECRET_PATH", "")
    ).expanduser()
    if str(secret_path) in {"", "."} or not secret_path.is_file():
        raise GoogleDeliveryError("Interactive authentication requires GOOGLE_CLIENT_SECRET_PATH.")
    from google_auth_oauthlib.flow import InstalledAppFlow

    flow = InstalledAppFlow.from_client_secrets_file(secret_path, SCOPES)
    credentials = flow.run_local_server(port=0, open_browser=False)
    output_path = Path(token_path or os.environ.get("GOOGLE_TOKEN_PATH", "~/.config/valuation/token.json")).expanduser()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(credentials.to_json(), encoding="utf-8")
    output_path.chmod(0o600)
    return credentials, "interactive"


def _configure_calculation(sheets_service, file_id: str) -> None:
    body = {
        "requests": [
            {
                "updateSpreadsheetProperties": {
                    "properties": {
                        "autoRecalc": "ON_CHANGE",
                        "iterativeCalculationSettings": {
                            "maxIterations": 100,
                            "convergenceThreshold": 0.0001,
                        },
                    },
                    "fields": "autoRecalc,iterativeCalculationSettings",
                }
            }
        ]
    }
    sheets_service.spreadsheets().batchUpdate(spreadsheetId=file_id, body=body).execute()


def _set_run_status(sheets_service, file_id: str, status: str) -> None:
    sheets_service.spreadsheets().values().update(
        spreadsheetId=file_id,
        range="'Run Metadata'!B2",
        valueInputOption="RAW",
        body={"values": [[status]]},
    ).execute()


def _download_snapshot(drive_service, file_id: str, output_path: Path) -> None:
    request = drive_service.files().export_media(fileId=file_id, mimeType=EXCEL_MIME_TYPE)
    content = request.execute()
    if not isinstance(content, (bytes, bytearray)) or not content:
        raise GoogleDeliveryError("Google Drive returned an empty spreadsheet export.")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(bytes(content))


def publish_and_verify(
    *,
    company_name: str,
    run_id: str,
    excel_file_path: str | Path,
    normalized_inputs_path: str | Path,
    template_path: str | Path,
    output_directory: str | Path,
    folder_id: str,
    file_id: str | None = None,
    replace: bool = False,
    share_with: str | None = None,
    service_account_path: str | Path | None = None,
    token_path: str | Path | None = None,
    allow_adc: bool = True,
    interactive_auth: bool = False,
    client_secret_path: str | Path | None = None,
    drive_service=None,
    sheets_service=None,
    poll_attempts: int = 5,
    poll_delay_seconds: float = 2.0,
) -> dict[str, Any]:
    source = Path(excel_file_path)
    if not source.is_file():
        raise GoogleDeliveryError(f"Workbook does not exist: {source}.")
    if not folder_id:
        raise GoogleDeliveryError("A writable Google Drive folder or shared-drive folder ID is required.")
    if file_id and not replace:
        raise GoogleDeliveryError("Updating an existing Google Sheet requires replace=True.")
    if not file_id and replace:
        raise GoogleDeliveryError("replace=True requires an explicit Google Sheet file ID.")
    if poll_attempts < 1:
        raise GoogleDeliveryError("poll_attempts must be at least 1.")

    credential_type = "injected"
    if drive_service is None or sheets_service is None:
        credentials, credential_type = _load_credentials(
            service_account_path=service_account_path,
            token_path=token_path,
            allow_adc=allow_adc,
            interactive_auth=interactive_auth,
            client_secret_path=client_secret_path,
        )
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload

        drive_service = drive_service or build("drive", "v3", credentials=credentials, cache_discovery=False)
        sheets_service = sheets_service or build("sheets", "v4", credentials=credentials, cache_discovery=False)
        media = MediaFileUpload(source, mimetype=EXCEL_MIME_TYPE, resumable=True)
    else:
        media = str(source)

    if file_id:
        updated = (
            drive_service.files()
            .update(
                fileId=file_id,
                media_body=media,
                fields="id,webViewLink,name",
                supportsAllDrives=True,
            )
            .execute()
        )
        delivery_mode = "replace"
        delivered = updated
    else:
        timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        file_metadata = {
            "name": f"{company_name} Valuation {timestamp} {run_id[:8]}",
            "mimeType": GOOGLE_SHEET_MIME_TYPE,
            "parents": [folder_id],
        }
        delivered = (
            drive_service.files()
            .create(
                body=file_metadata,
                media_body=media,
                fields="id,webViewLink,name",
                supportsAllDrives=True,
            )
            .execute()
        )
        file_id = delivered.get("id")
        delivery_mode = "create"

    if not file_id:
        raise GoogleDeliveryError("Google Drive did not return a file ID.")
    if share_with:
        drive_service.permissions().create(
            fileId=file_id,
            body={"type": "user", "role": "writer", "emailAddress": share_with},
            sendNotificationEmail=True,
            supportsAllDrives=True,
        ).execute()

    _configure_calculation(sheets_service, file_id)
    _set_run_status(sheets_service, file_id, "Recalculated; verification pending")

    output_dir = Path(output_directory)
    snapshot_path = output_dir / f"{run_id}.recalculated.xlsx"
    verification: dict[str, Any] | None = None
    last_error: Exception | None = None
    for attempt in range(1, poll_attempts + 1):
        _download_snapshot(drive_service, file_id, snapshot_path)
        try:
            verification = verify_recalculated(
                snapshot_path,
                normalized_inputs_path,
                template_path,
            )
            last_error = None
            break
        except WorkbookVerificationError as error:
            last_error = error
            if attempt < poll_attempts and poll_delay_seconds:
                time.sleep(poll_delay_seconds)
    if last_error is not None or verification is None:
        raise GoogleDeliveryError(
            "The Google Sheet was created but its recalculated export did not pass verification."
        ) from last_error

    _set_run_status(sheets_service, file_id, "Complete")
    receipt = {
        "status": "complete",
        "delivery_mode": delivery_mode,
        "credential_type": credential_type,
        "run_id": run_id,
        "file_id": file_id,
        "file_name": delivered.get("name"),
        "web_view_link": delivered.get("webViewLink"),
        "folder_id": folder_id,
        "source_workbook_sha256": file_sha256(source),
        "snapshot_path": str(snapshot_path.resolve()),
        "snapshot_sha256": file_sha256(snapshot_path),
        "verified_at": datetime.now(timezone.utc).isoformat(),
        "verification": verification,
    }
    atomic_write_json(output_dir / f"{run_id}.google-delivery.json", receipt)
    return receipt


def main() -> None:
    parser = argparse.ArgumentParser(description="Publish and verify a v2 valuation in Google Sheets")
    subparsers = parser.add_subparsers(dest="command", required=True)

    auth_parser = subparsers.add_parser("auth-check", help="Validate local credential configuration without uploading")
    auth_parser.add_argument("--service-account-path")
    auth_parser.add_argument("--token-path")

    publish_parser = subparsers.add_parser("publish", help="Create or explicitly replace a Google Sheet")
    publish_parser.add_argument("--company", required=True)
    publish_parser.add_argument("--run-id", required=True)
    publish_parser.add_argument("--file", required=True)
    publish_parser.add_argument("--inputs", required=True)
    publish_parser.add_argument("--template", required=True)
    publish_parser.add_argument("--output-dir", required=True)
    publish_parser.add_argument("--folder-id", required=True)
    publish_parser.add_argument("--file-id")
    publish_parser.add_argument("--replace", action="store_true")
    publish_parser.add_argument("--share-with")
    publish_parser.add_argument("--service-account-path")
    publish_parser.add_argument("--token-path")
    publish_parser.add_argument("--no-adc", action="store_true")
    publish_parser.add_argument("--interactive-auth", action="store_true")
    publish_parser.add_argument("--client-secret-path")
    args = parser.parse_args()

    try:
        if args.command == "auth-check":
            receipt = credential_preflight(
                service_account_path=args.service_account_path,
                token_path=args.token_path,
            )
        else:
            receipt = publish_and_verify(
                company_name=args.company,
                run_id=args.run_id,
                excel_file_path=args.file,
                normalized_inputs_path=args.inputs,
                template_path=args.template,
                output_directory=args.output_dir,
                folder_id=args.folder_id,
                file_id=args.file_id,
                replace=args.replace,
                share_with=args.share_with,
                service_account_path=args.service_account_path,
                token_path=args.token_path,
                allow_adc=not args.no_adc,
                interactive_auth=args.interactive_auth,
                client_secret_path=args.client_secret_path,
            )
    except Exception as error:
        print(f"Google delivery failed: {error}", file=sys.stderr)
        raise SystemExit(1) from error
    print(json.dumps(receipt, indent=2))


if __name__ == "__main__":  # pragma: no cover - exercised through main()
    main()
