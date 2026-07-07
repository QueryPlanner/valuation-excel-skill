# Google delivery

## Service accounts

Set `GOOGLE_SERVICE_ACCOUNT_PATH` to a service-account JSON file. Do not copy the
key into the skill directory.

Grant the service account writer access to a target folder or shared-drive
folder, then pass that folder ID to the workflow. Key validity and folder access
are separate checks.

Use a shared drive for production automation when ownership and storage policy
require organizational control.

## OAuth or ADC

Alternatively set `GOOGLE_TOKEN_PATH`, or configure Application Default
Credentials. Interactive OAuth is disabled unless explicitly requested.

Token files created through interactive authentication are written with
owner-only permissions.

## Creation and replacement

New delivery creates a uniquely named Google Sheet in the specified folder.

Replacement requires:

- an explicit Google Sheet file ID,
- the `--replace` flag,
- user confirmation before the command is invoked.

Replacement uploads the complete workbook and replaces the existing Sheet
contents.

## Verification

After import, the uploader:

1. Enables recalculation and iterative calculation.
2. Exports a recalculated XLSX snapshot.
3. Checks WACC, operating assets, common equity, value per share, active formula
   paths, input identity, and protected formulas.
4. Writes a JSON delivery receipt.

An upload without a passing exported snapshot is not complete.
