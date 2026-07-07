---
name: valuation-excel-skill-v3
description: Builds and verifies an auditable SEC-based intrinsic valuation with strict evidence validation, versioned run manifests, Excel audit sheets, local LibreOffice recalculation, and optional verified Google Sheets delivery. Use for reliable 10-K/10-Q valuation runs, workbook population, assumption review, recalculation verification, local Excel delivery, or Google delivery troubleshooting.
---

# Reliable Valuation Excel Workflow

Resolve this skill directory before running scripts:

```bash
VALUATION_V3_DIR="${CODEX_HOME:-$HOME/.codex}/skills/valuation-excel-skill-v3"
```

For a development checkout, set `VALUATION_V3_DIR` to the directory containing
this file.

## Completion contract

Use these statuses exactly:

- `complete`: a calculation backend recalculated the workbook and
  `verify_workbook.py --stage recalculated` passed.
- `awaiting_recalculation`: the workbook was populated and passed structural
  verification, but calculated outputs are not yet trustworthy.
- `failed`: validation, population, recalculation, verification, or delivery
  failed.

Never describe `awaiting_recalculation` as a completed valuation. OpenPyXL writes
formulas but does not calculate them.

## 1. Build the filing inventory

Set `SEC_USER_AGENT` to an application name and monitored contact email. Then run:

```bash
uv run --project "$VALUATION_V3_DIR" python \
  "$VALUATION_V3_DIR/scripts/get_financial_reports.py" \
  --company TICKER \
  --download-dir "downloads/TICKER" \
  --output "TICKER_reports.v2.json"
```

The downloader requests up to ten 10-Ks, two 10-Qs, and Company Facts. It records
shortages instead of rejecting a young issuer. Treat Company Facts as a
cross-check, not a complete substitute for filing-level notes and tables.

## 2. Create evidence-backed inputs

Read [references/extraction_contract.md](references/extraction_contract.md).
Use `schema_version: "2.0"`. Record one evidence item for every required numeric
input. Filing evidence must match an accession, URL, and SHA-256 in the filing
manifest.

Do not:

- default missing reported facts to zero,
- mix currencies or unit scales,
- use a current market price with an unrelated valuation date,
- use company segment names as workbook industry labels,
- enable operating-lease conversion when book debt already includes those
  liabilities.

If required disclosures are unavailable, populate
`missing_documents_required` and stop.

## 3. Validate and review assumptions

Run the validator:

```bash
uv run --project "$VALUATION_V3_DIR" python \
  "$VALUATION_V3_DIR/scripts/validate_inputs.py" \
  --inputs "TICKER_inputs.json" \
  --template "$VALUATION_V3_DIR/assets/template.xlsx" \
  --output "TICKER_inputs.normalized.json" \
  --report "TICKER_validation.json"
```

Fix every error. Review every warning. Industry-range warnings are review prompts,
not automatic reasons to force assumptions toward industry medians.

Read [references/valuation_rules.md](references/valuation_rules.md) before
approving the base case. Keep the story and numbers consistent.

## 4. Run the workflow

For a verified local workbook, use LibreOffice:

```bash
uv run --project "$VALUATION_V3_DIR" python \
  "$VALUATION_V3_DIR/scripts/run_valuation.py" \
  --inputs "TICKER_inputs.json" \
  --template "$VALUATION_V3_DIR/assets/template.xlsx" \
  --output-root "valuation-runs" \
  --backend libreoffice
```

The command requires `soffice` or `libreoffice` on `PATH`. Override discovery
with `--libreoffice-executable PATH`. The backend recalculates in an isolated
temporary profile, exports `valuation.recalculated.xlsx`, restores Excel
calculation metadata, and runs recalculated-stage verification. Do not trust
the workbook if the backend or verification fails.

For a workbook that intentionally remains incomplete:

```bash
uv run --project "$VALUATION_V3_DIR" python \
  "$VALUATION_V3_DIR/scripts/run_valuation.py" \
  --inputs "TICKER_inputs.json" \
  --template "$VALUATION_V3_DIR/assets/template.xlsx" \
  --output-root "valuation-runs" \
  --backend none
```

Exit code `2` means `awaiting_recalculation`, not failure.

For Google Sheets recalculation and verified delivery instead:

```bash
uv run --project "$VALUATION_V3_DIR" python \
  "$VALUATION_V3_DIR/scripts/run_valuation.py" \
  --inputs "TICKER_inputs.json" \
  --template "$VALUATION_V3_DIR/assets/template.xlsx" \
  --output-root "valuation-runs" \
  --backend google \
  --google-folder-id "WRITABLE_FOLDER_ID"
```

Google mode creates a uniquely named Sheet, enables iterative calculation,
exports a recalculated snapshot, verifies outputs, and persists a delivery
receipt.

## Google credentials and write safety

Read [references/google_delivery.md](references/google_delivery.md) when using
Google delivery.

Check local credential structure without uploading:

```bash
uv run --project "$VALUATION_V3_DIR" python \
  "$VALUATION_V3_DIR/scripts/upload_to_sheets.py" auth-check
```

A valid service-account key is insufficient by itself. The target folder or
shared-drive folder must grant that account write access.

Create a new Sheet by default. Update an existing Sheet only when the user
explicitly identifies the file ID and confirms full-content replacement. Do not
infer overwrite permission from a matching filename.

Never print credential contents, private keys, OAuth tokens, or client secrets.

## Inspect the run

Each run directory contains:

- `run.json`: authoritative state and step receipt,
- `inputs.normalized.json`: validated model inputs,
- `valuation.awaiting-recalculation.xlsx`: populated workbook,
- `precalculation-verification.json`: formula and input checks,
- `valuation.recalculated.xlsx`, `recalculation-verification.json`, and
  `libreoffice-recalculation.json` when local recalculation succeeds,
- a recalculated snapshot and Google delivery receipt when Google succeeds.

The workbook includes `Sources & Audit` and `Run Metadata` sheets. Verify the
recalculated snapshot before relying on any valuation output.

## Current limits

This version deliberately does not implement:

- a complete automatic Inline XBRL mapper,
- a standalone DCF engine independent of the workbook,
- Monte Carlo simulation,
- banks, insurers, or other specialized valuation models.

Do not represent those capabilities as present.
