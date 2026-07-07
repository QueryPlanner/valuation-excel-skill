# Valuation Excel Skill v3

Auditable SEC-based intrinsic valuation workflow for Codex. The skill downloads
SEC filings, validates evidence-backed assumptions, populates the valuation
workbook, and verifies recalculated outputs through LibreOffice or Google
Sheets.

## What changed in v3

- Versioned run directories with `run.json` status receipts.
- Strict input validation using `schema_version: "2.0"`.
- Required evidence for numeric assumptions, tied to filing accession, URL, and
  SHA-256 manifest entries.
- Workbook audit sheets for sources and run metadata.
- Local LibreOffice recalculation with isolated profile handling.
- Optional Google Sheets delivery with export and verification.
- Test coverage for CLI entrypoints, validation, workbook population,
  recalculation, SEC downloading, and Google delivery safety.

## Repository layout

```text
assets/template.xlsx                 Valuation workbook template
references/extraction_contract.md    Required input schema and evidence rules
references/valuation_rules.md        Assumption and valuation review guidance
references/google_delivery.md        Google Sheets credential and write safety
scripts/get_financial_reports.py     SEC filing and Company Facts downloader
scripts/validate_inputs.py           Input validator and normalizer
scripts/run_valuation.py             End-to-end workflow runner
scripts/fill_excel.py                Workbook population
scripts/verify_workbook.py           Workbook verification
scripts/recalculate_with_libreoffice.py
scripts/upload_to_sheets.py
tests/                              Unit and integration-style tests
```

## Requirements

- Python 3.11 or newer
- `uv`
- LibreOffice for local recalculation with `--backend libreoffice`
- Google service-account credentials only when using `--backend google`

Install and run commands through `uv`:

```bash
uv sync
```

## SEC downloader setup

Set a monitored contact in `SEC_USER_AGENT` before requesting filings:

```bash
export SEC_USER_AGENT="valuation-excel-skill your-email@example.com"
```

Download a filing inventory:

```bash
uv run --project . python scripts/get_financial_reports.py \
  --company NVDA \
  --download-dir downloads/NVDA \
  --output NVDA_reports.v2.json
```

The downloader requests up to ten 10-Ks, two 10-Qs, and Company Facts. Shortages
are recorded instead of rejected so young issuers can still be reviewed.

## Prepare inputs

Create an inputs JSON file that follows
`references/extraction_contract.md`. Use:

```json
{
  "schema_version": "2.0"
}
```

Every required numeric input needs evidence. Filing evidence must match the
accession, URL, and SHA-256 values in the filing manifest.

Validate and normalize inputs:

```bash
uv run --project . python scripts/validate_inputs.py \
  --inputs NVDA_inputs.json \
  --template assets/template.xlsx \
  --output NVDA_inputs.normalized.json \
  --report NVDA_validation.json
```

Fix validation errors before running the valuation. Review warnings against
`references/valuation_rules.md`.

## Run a valuation

For a fully verified local workbook:

```bash
uv run --project . python scripts/run_valuation.py \
  --inputs NVDA_inputs.json \
  --template assets/template.xlsx \
  --output-root valuation-runs \
  --backend libreoffice
```

The LibreOffice backend requires `soffice` or `libreoffice` on `PATH`. Use
`--libreoffice-executable PATH` to override discovery.

For workbook population without recalculation:

```bash
uv run --project . python scripts/run_valuation.py \
  --inputs NVDA_inputs.json \
  --template assets/template.xlsx \
  --output-root valuation-runs \
  --backend none
```

Exit code `2` means the workbook is `awaiting_recalculation`. Do not treat that
state as a completed valuation.

For Google Sheets recalculation and verified delivery:

```bash
uv run --project . python scripts/run_valuation.py \
  --inputs NVDA_inputs.json \
  --template assets/template.xlsx \
  --output-root valuation-runs \
  --backend google \
  --google-folder-id WRITABLE_FOLDER_ID
```

Read `references/google_delivery.md` before using Google mode. Check credential
structure without uploading:

```bash
uv run --project . python scripts/upload_to_sheets.py auth-check
```

## Run outputs

Each run directory contains:

- `run.json`
- `inputs.normalized.json`
- `valuation.awaiting-recalculation.xlsx`
- `precalculation-verification.json`
- `valuation.recalculated.xlsx` when recalculation succeeds
- `recalculation-verification.json` when recalculation succeeds
- backend-specific receipts for LibreOffice or Google Sheets

The completion statuses are:

- `complete`: a calculation backend recalculated the workbook and verification
  passed.
- `awaiting_recalculation`: formulas were written, but outputs are not yet
  trustworthy.
- `failed`: validation, population, recalculation, verification, or delivery
  failed.

## Test

Run the test suite:

```bash
uv run coverage run --branch --source=scripts -m unittest discover -s tests -t . -q
uv run coverage report
```

## Current limits

This version does not implement:

- Complete automatic Inline XBRL mapping
- A standalone DCF engine independent of the workbook
- Monte Carlo simulation
- Bank, insurer, or other specialized valuation models
