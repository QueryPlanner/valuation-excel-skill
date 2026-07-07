# Extraction contract

## Required top-level objects

Use `schema_version: "2.0"` and provide:

- `company_context`
- `filing_manifest`
- `financial_data`
- `revenue_splits`
- `cost_of_capital_inputs`
- `employee_options`
- `operating_lease_details`
- `r_and_d_details`
- `single_value_metrics`
- `market_inputs`
- `base_case_assumptions`
- `company_story`
- `advanced_assumptions`
- `source_evidence`
- `missing_documents_required`

Use `tests/fixtures/valid_inputs.json` as the structural example. Replace every
fixture value and source with company-specific evidence.

## Company and period identity

`company_context` requires:

- `company_name`
- `valuation_date` in `YYYY-MM-DD`
- `ticker`
- ten-digit `cik`
- workbook-valid `country_of_incorporation`
- workbook-valid US and global industries
- three-letter uppercase currency
- units: `units`, `thousands`, `millions`, or `billions`
- fiscal year-end label

All financial statement values and share counts must use the stated currency and
scale.

## Filing manifest

Embed the output of `get_financial_reports.py` as `filing_manifest`. Filing
evidence must match its accession number, HTML URL, and SHA-256 hash.

## Evidence records

Every required numeric input needs exactly one evidence record with:

- `metric`: dotted input path
- `source_type`: `filing`, `market`, `industry`, `analyst_assumption`, or
  `model_default`
- `value`
- `units`
- `period`
- `calculation`

Filing evidence also requires:

- `accession_number`
- `source_url`
- `sha256`
- `section`
- `reported_label`

Market and industry evidence require `source_url`. Analyst assumptions and model
defaults require `rationale`.

The evidence value must equal the normalized input value. Do not cite one fact
while entering another.

## Revenue geography

Set `revenue_splits.geography_mode` explicitly:

- `incorporation`: leave country and region mappings empty.
- `country`: provide `by_country` only.
- `region`: provide `by_region` only.

Business and selected geography totals must reconcile to `period_revenue` within
2%. Every mapping value must be finite and greater than zero.

If a country split includes `Rest of the World`, provide an independently
supported `cost_of_capital_inputs.rest_of_world_erp`.

## Flow and snapshot facts

For flow metrics at an interim date:

`LTM = latest fiscal year + current YTD - prior comparable YTD`

Use the same duration for both YTD facts.

Use latest quarter-end values for equity, debt, cash, non-operating assets,
minority interests, shares, and option balances. Do not calculate LTM balances.

## Operating leases

Set `capitalize: false` when reported debt already includes operating lease
liabilities.

When conversion is required, set:

- `capitalize: true`
- `book_debt_excludes_operating_leases: true`
- current-year operating lease expense
- commitments for years 1 through 5
- commitments for year 6 and beyond

Do not enable the converter without the complete schedule.

## Market inputs

Provide stock price, risk-free rate, and a market `as_of_date`. Cite the source.
The valuation date and market date should normally match; a mismatch produces a
review warning.

## Missing disclosures

List an exact missing filing, footnote, schedule, or market source in
`missing_documents_required`. Any non-empty entry blocks workbook creation.
