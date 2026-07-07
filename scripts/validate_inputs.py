from __future__ import annotations

import argparse
import copy
import json
import re
import sys
from datetime import date
from pathlib import Path
from typing import Any, Callable

from common import atomic_write_json, is_finite_number, nested_get, numbers_match
from workbook_contract import (
    ContractValueError,
    WorkbookContract,
    canonicalize_country,
    canonicalize_industry,
    canonicalize_rating,
    canonicalize_region,
    load_workbook_contract,
)

SCHEMA_VERSION = "2.0"
FINANCIAL_METRICS = (
    "Revenues",
    "Operating_income_or_EBIT",
    "Interest_expense",
    "Book_value_of_equity",
    "Book_value_of_debt",
    "Cash_and_Marketable_Securities",
    "Cross_holdings_and_other_non_operating_assets",
    "Minority_interests",
)
NON_NEGATIVE_FINANCIAL_METRICS = (
    "Interest_expense",
    "Book_value_of_debt",
    "Cash_and_Marketable_Securities",
    "Cross_holdings_and_other_non_operating_assets",
    "Minority_interests",
)
FINANCIAL_PERIODS = ("Most_Recent_12_months", "Last_10K_before_LTM")
STORY_FIELDS = (
    "title",
    "summary",
    "growth",
    "profitability",
    "tax",
    "reinvestment",
    "competitive_advantage",
    "risk",
)
BASE_CASE_FIELDS = (
    "revenue_growth_next_year",
    "revenue_cagr_years_2_5",
    "operating_margin_next_year",
    "target_operating_margin",
    "margin_convergence_year",
    "sales_to_capital_years_1_5",
    "sales_to_capital_years_6_10",
)
ADVANCED_DEFAULTS = {
    "terminal_cost_of_capital_override": None,
    "stable_return_on_capital_override": None,
    "probability_of_failure": 0.0,
    "failure_proceeds_basis": "V",
    "failure_proceeds_percent": 0.5,
    "reinvestment_lag_years": 1,
    "hold_effective_tax_rate": False,
    "net_operating_loss": 0.0,
    "stable_riskfree_rate_override": None,
    "terminal_growth_override": None,
    "trapped_cash": 0.0,
    "trapped_cash_tax_rate": 0.0,
}
SOURCE_TYPES = {"filing", "market", "industry", "analyst_assumption", "model_default"}


class InputValidationError(ValueError):
    def __init__(self, errors: list[str], findings: list[dict[str, str]] | None = None):
        self.errors = errors
        self.findings = findings or []
        super().__init__("\n".join(f"- {error}" for error in errors))


def _require_object(errors: list[str], path: str, value: Any) -> dict[str, Any]:
    if not isinstance(value, dict):
        errors.append(f"{path} must be an object.")
        return {}
    return value


def _require_string(errors: list[str], path: str, value: Any) -> str:
    if not isinstance(value, str) or not value.strip():
        errors.append(f"{path} must be a non-empty string.")
        return ""
    return value.strip()


def _require_number(errors: list[str], path: str, value: Any) -> float | None:
    if not is_finite_number(value):
        errors.append(f"{path} must be a finite number.")
        return None
    return float(value)


def _parse_iso_date(errors: list[str], path: str, value: Any) -> str:
    text = _require_string(errors, path, value)
    if not text:
        return ""
    try:
        parsed = date.fromisoformat(text)
    except ValueError:
        errors.append(f"{path} must use YYYY-MM-DD format.")
        return text
    if parsed > date.today():
        errors.append(f"{path} cannot be in the future.")
    return text


def _canonicalize_mapping(
    values: dict[str, Any],
    canonicalizer: Callable[[str], str],
    path: str,
    errors: list[str],
) -> dict[str, float]:
    canonical: dict[str, float] = {}
    for name, raw_value in values.items():
        value = _require_number(errors, f"{path}.{name}", raw_value)
        if value is None:
            continue
        if value <= 0:
            errors.append(f"{path}.{name} must be greater than zero.")
            continue
        try:
            canonical_name = canonicalizer(str(name))
        except ContractValueError as error:
            errors.append(str(error))
            continue
        canonical[canonical_name] = canonical.get(canonical_name, 0.0) + value
    return canonical


def _canonicalize_country_or_rest_of_world(value: str, contract: WorkbookContract) -> str:
    normalized = " ".join(value.casefold().split())
    if normalized in {"rest of world", "rest of the world"}:
        return "Rest of the World"
    return canonicalize_country(value, contract)


def _validate_reconciliation(
    errors: list[str],
    label: str,
    values: dict[str, float],
    expected_total: float,
) -> None:
    actual_total = sum(values.values())
    tolerance = max(abs(expected_total) * 0.02, 1e-6)
    if abs(actual_total - expected_total) > tolerance:
        errors.append(
            f"{label} total {actual_total:,.6f} does not reconcile to "
            f"revenue_splits.period_revenue {expected_total:,.6f} within 2%."
        )


def _flatten_manifest_filings(manifest: Any) -> list[dict[str, Any]]:
    if not isinstance(manifest, dict):
        return []
    candidates: list[Any] = []
    candidates.extend(manifest.get("filings", []))
    candidates.extend(manifest.get("latest_four_10ks", []))
    candidates.extend(manifest.get("latest_two_10qs", []))
    output: list[dict[str, Any]] = []
    seen: set[str] = set()
    for candidate in candidates:
        if not isinstance(candidate, dict):
            continue
        accession = candidate.get("accession_number")
        if isinstance(accession, str) and accession not in seen:
            output.append(candidate)
            seen.add(accession)
    return output


def _upgrade_legacy_input(
    raw_inputs: dict[str, Any],
    findings: list[dict[str, str]],
) -> dict[str, Any]:
    upgraded = copy.deepcopy(raw_inputs)
    upgraded["schema_version"] = SCHEMA_VERSION
    revenue_splits = upgraded.setdefault("revenue_splits", {})
    if "geography_mode" not in revenue_splits:
        if revenue_splits.get("by_country"):
            revenue_splits["geography_mode"] = "country"
        elif revenue_splits.get("by_region"):
            revenue_splits["geography_mode"] = "region"
        else:
            revenue_splits["geography_mode"] = "incorporation"
    findings.append(
        {
            "severity": "warning",
            "code": "legacy_input_upgraded",
            "message": "A v1 input was upgraded in memory. Supply schema_version 2.0 and a filing_manifest.",
        }
    )
    return upgraded


def _required_evidence_metrics(inputs: dict[str, Any]) -> set[str]:
    metrics = {
        f"financial_data.{metric}.{period}"
        for metric in FINANCIAL_METRICS
        for period in FINANCIAL_PERIODS
    }
    metrics.update(
        {
            "single_value_metrics.Years_since_last_10K",
            "single_value_metrics.Number_of_shares_outstanding",
            "single_value_metrics.Effective_tax_rate",
            "single_value_metrics.Marginal_tax_rate",
            "cost_of_capital_inputs.average_maturity_of_debt_years",
            "market_inputs.stock_price",
            "market_inputs.riskfree_rate",
            "revenue_splits.period_revenue",
        }
    )
    metrics.update(f"base_case_assumptions.{field}" for field in BASE_CASE_FIELDS)
    cost_inputs = inputs.get("cost_of_capital_inputs", {})
    for field in ("direct_erp", "rest_of_world_erp"):
        if cost_inputs.get(field) is not None:
            metrics.add(f"cost_of_capital_inputs.{field}")

    splits = inputs.get("revenue_splits", {})
    for split_name in ("by_country", "by_region", "by_business"):
        for name in splits.get(split_name, {}):
            metrics.add(f"revenue_splits.{split_name}.{name}")

    r_and_d = inputs.get("r_and_d_details", {})
    if is_finite_number(r_and_d.get("current_year_expense")) and r_and_d["current_year_expense"] > 0:
        metrics.add("r_and_d_details.current_year_expense")
        amortization_years = r_and_d.get("amortization_period_years")
        if isinstance(amortization_years, int):
            for year_number in range(1, amortization_years + 1):
                metrics.add(f"r_and_d_details.historical_expenses.Year_Minus_{year_number}")

    options = inputs.get("employee_options", {})
    if is_finite_number(options.get("total_options_outstanding")) and options["total_options_outstanding"] > 0:
        metrics.update(
            {
                "employee_options.total_options_outstanding",
                "employee_options.weighted_average_exercise_price",
                "employee_options.average_maturity_years",
                "employee_options.stock_price_standard_deviation",
            }
        )

    lease = inputs.get("operating_lease_details", {})
    if lease.get("capitalize"):
        metrics.add("operating_lease_details.current_year_expense")
        for key in (
            "year_1",
            "year_2",
            "year_3",
            "year_4",
            "year_5",
            "years_6_and_beyond",
        ):
            metrics.add(f"operating_lease_details.commitments.{key}")
    return metrics


def _validate_source_evidence(
    inputs: dict[str, Any],
    errors: list[str],
    allow_legacy: bool,
) -> None:
    records = inputs.get("source_evidence")
    if not isinstance(records, list) or not records:
        errors.append("source_evidence must contain evidence records.")
        return

    manifest_filings = _flatten_manifest_filings(inputs.get("filing_manifest"))
    manifest_by_accession = {
        filing.get("accession_number"): filing
        for filing in manifest_filings
        if isinstance(filing.get("accession_number"), str)
    }
    evidence_by_metric: dict[str, dict[str, Any]] = {}
    for index, record in enumerate(records):
        path = f"source_evidence[{index}]"
        if not isinstance(record, dict):
            errors.append(f"{path} must be an object.")
            continue
        metric = _require_string(errors, f"{path}.metric", record.get("metric"))
        source_type = _require_string(errors, f"{path}.source_type", record.get("source_type"))
        _require_string(errors, f"{path}.calculation", record.get("calculation"))
        _require_string(errors, f"{path}.units", record.get("units"))
        if source_type and source_type not in SOURCE_TYPES:
            errors.append(f"{path}.source_type must be one of {sorted(SOURCE_TYPES)}.")
        if metric in evidence_by_metric:
            errors.append(f"source_evidence contains duplicate metric {metric}.")
        elif metric:
            evidence_by_metric[metric] = record

        if source_type == "filing":
            accession = _require_string(errors, f"{path}.accession_number", record.get("accession_number"))
            source_url = _require_string(errors, f"{path}.source_url", record.get("source_url"))
            sha256 = _require_string(errors, f"{path}.sha256", record.get("sha256"))
            _require_string(errors, f"{path}.period", record.get("period"))
            if not re.fullmatch(r"[0-9a-fA-F]{64}", sha256):
                errors.append(f"{path}.sha256 must be a 64-character SHA-256 digest.")
            manifest_record = manifest_by_accession.get(accession)
            if not manifest_record:
                if not allow_legacy:
                    errors.append(f"{path}.accession_number is not present in filing_manifest.")
            else:
                if manifest_record.get("sha256") != sha256:
                    errors.append(f"{path}.sha256 does not match filing_manifest for {accession}.")
                if manifest_record.get("html_url") and manifest_record.get("html_url") != source_url:
                    errors.append(f"{path}.source_url does not match filing_manifest for {accession}.")
        elif source_type in {"market", "industry"}:
            _require_string(errors, f"{path}.source_url", record.get("source_url"))
            _require_string(errors, f"{path}.period", record.get("period"))
        elif source_type in {"analyst_assumption", "model_default"}:
            _require_string(errors, f"{path}.rationale", record.get("rationale"))

        if metric:
            try:
                input_value = nested_get(inputs, metric)
            except KeyError:
                errors.append(f"{path}.metric {metric!r} does not resolve to an input value.")
                continue
            evidence_value = record.get("value")
            if is_finite_number(input_value):
                if not is_finite_number(evidence_value) or not numbers_match(input_value, evidence_value):
                    errors.append(f"{path}.value does not match {metric}.")
            elif evidence_value != input_value:
                errors.append(f"{path}.value does not match {metric}.")

    missing = sorted(_required_evidence_metrics(inputs) - set(evidence_by_metric))
    errors.extend(f"source_evidence is missing {metric}." for metric in missing)


def _append_model_default_evidence(inputs: dict[str, Any], defaults_applied: list[str]) -> None:
    records = inputs.setdefault("source_evidence", [])
    existing = {record.get("metric") for record in records if isinstance(record, dict)}
    for field in defaults_applied:
        metric = f"advanced_assumptions.{field}"
        if metric in existing:
            continue
        records.append(
            {
                "metric": metric,
                "source_type": "model_default",
                "period": inputs["company_context"]["valuation_date"],
                "calculation": "Deterministic v2 model default",
                "value": inputs["advanced_assumptions"][field],
                "units": "model setting",
                "rationale": "Explicit default from the v2 valuation contract.",
            }
        )


def _assumption_findings(
    inputs: dict[str, Any],
    contract: WorkbookContract,
    findings: list[dict[str, str]],
) -> None:
    industry = inputs["company_context"]["industry_global"]
    distribution = contract.industry_distributions.get(industry)
    if not distribution:
        findings.append(
            {
                "severity": "warning",
                "code": "industry_distribution_missing",
                "message": f"No assumption distribution is available for {industry}.",
            }
        )
        return

    base_case = inputs["base_case_assumptions"]
    checks = (
        ("revenue_growth_next_year", distribution.growth_q1, distribution.growth_q3),
        ("revenue_cagr_years_2_5", distribution.growth_q1, distribution.growth_q3),
        ("operating_margin_next_year", distribution.margin_q1, distribution.margin_q3),
        ("target_operating_margin", distribution.margin_q1, distribution.margin_q3),
    )
    for field, lower, upper in checks:
        value = float(base_case[field])
        if value < lower or value > upper:
            findings.append(
                {
                    "severity": "warning",
                    "code": "outside_industry_interquartile_range",
                    "message": (
                        f"base_case_assumptions.{field}={value:.2%} is outside the "
                        f"{industry} interquartile range {lower:.2%} to {upper:.2%}."
                    ),
                }
            )


def normalize_and_validate_inputs(
    raw_inputs: dict[str, Any],
    contract: WorkbookContract,
    *,
    allow_legacy: bool = False,
) -> dict[str, Any]:
    if not isinstance(raw_inputs, dict):
        raise InputValidationError(["The valuation inputs must be an object."])

    errors: list[str] = []
    findings: list[dict[str, str]] = []
    schema_version = raw_inputs.get("schema_version")
    if schema_version != SCHEMA_VERSION:
        if allow_legacy and schema_version is None:
            inputs = _upgrade_legacy_input(raw_inputs, findings)
        else:
            raise InputValidationError([f"schema_version must be {SCHEMA_VERSION!r}."])
    else:
        inputs = copy.deepcopy(raw_inputs)

    company_context = _require_object(errors, "company_context", inputs.get("company_context"))
    valuation_date = _parse_iso_date(errors, "company_context.valuation_date", company_context.get("valuation_date"))
    _require_string(errors, "company_context.company_name", company_context.get("company_name"))
    ticker = _require_string(errors, "company_context.ticker", company_context.get("ticker"))
    if ticker and len(ticker) > 20:
        errors.append("company_context.ticker must be at most 20 characters.")
    cik = company_context.get("cik")
    if not allow_legacy:
        cik_text = _require_string(errors, "company_context.cik", cik)
        if cik_text and not re.fullmatch(r"\d{10}", cik_text):
            errors.append("company_context.cik must contain exactly 10 digits.")
    country = company_context.get("country_of_incorporation")
    if country:
        try:
            company_context["country_of_incorporation"] = canonicalize_country(country, contract)
        except ContractValueError as error:
            errors.append(str(error))
    else:
        errors.append("company_context.country_of_incorporation is required.")
    currency = _require_string(errors, "company_context.currency", company_context.get("currency"))
    if currency and not re.fullmatch(r"[A-Z]{3}", currency):
        errors.append("company_context.currency must be a three-letter uppercase code.")
    units = _require_string(errors, "company_context.units", company_context.get("units"))
    if units and units not in {"units", "thousands", "millions", "billions"}:
        errors.append("company_context.units must be units, thousands, millions, or billions.")
    _require_string(errors, "company_context.fiscal_year_end", company_context.get("fiscal_year_end"))

    filing_manifest = inputs.get("filing_manifest")
    if not allow_legacy:
        _require_object(errors, "filing_manifest", filing_manifest)
        if not _flatten_manifest_filings(filing_manifest):
            errors.append("filing_manifest must contain at least one filing with an accession number.")

    financial_data = _require_object(errors, "financial_data", inputs.get("financial_data"))
    for metric_name in FINANCIAL_METRICS:
        metric = _require_object(errors, f"financial_data.{metric_name}", financial_data.get(metric_name))
        for period_name in FINANCIAL_PERIODS:
            value = _require_number(
                errors,
                f"financial_data.{metric_name}.{period_name}",
                metric.get(period_name),
            )
            if value is not None and metric_name in NON_NEGATIVE_FINANCIAL_METRICS and value < 0:
                errors.append(f"financial_data.{metric_name}.{period_name} cannot be negative.")
    for period_name in FINANCIAL_PERIODS:
        revenue = financial_data.get("Revenues", {}).get(period_name)
        if is_finite_number(revenue) and revenue <= 0:
            errors.append(f"financial_data.Revenues.{period_name} must be greater than zero.")

    single_metrics = _require_object(errors, "single_value_metrics", inputs.get("single_value_metrics"))
    years_since = _require_number(
        errors,
        "single_value_metrics.Years_since_last_10K",
        single_metrics.get("Years_since_last_10K"),
    )
    if years_since is not None and years_since not in {0.25, 0.5, 0.75, 1.0}:
        errors.append("single_value_metrics.Years_since_last_10K must be 0.25, 0.50, 0.75, or 1.00.")
    shares = _require_number(
        errors,
        "single_value_metrics.Number_of_shares_outstanding",
        single_metrics.get("Number_of_shares_outstanding"),
    )
    if shares is not None and shares <= 0:
        errors.append("single_value_metrics.Number_of_shares_outstanding must be greater than zero.")
    for field in ("Effective_tax_rate", "Marginal_tax_rate"):
        value = _require_number(errors, f"single_value_metrics.{field}", single_metrics.get(field))
        if value is not None and not 0 <= value <= 1:
            errors.append(f"single_value_metrics.{field} must be between 0 and 1.")

    revenue_splits = _require_object(errors, "revenue_splits", inputs.get("revenue_splits"))
    _require_string(errors, "revenue_splits.period_label", revenue_splits.get("period_label"))
    period_revenue = _require_number(errors, "revenue_splits.period_revenue", revenue_splits.get("period_revenue"))
    if period_revenue is not None and period_revenue <= 0:
        errors.append("revenue_splits.period_revenue must be greater than zero.")
    geography_mode = _require_string(errors, "revenue_splits.geography_mode", revenue_splits.get("geography_mode"))
    if geography_mode and geography_mode not in {"incorporation", "country", "region"}:
        errors.append("revenue_splits.geography_mode must be incorporation, country, or region.")

    by_country_raw = _require_object(errors, "revenue_splits.by_country", revenue_splits.get("by_country", {}))
    by_region_raw = _require_object(errors, "revenue_splits.by_region", revenue_splits.get("by_region", {}))
    by_business_raw = _require_object(errors, "revenue_splits.by_business", revenue_splits.get("by_business"))
    by_country = _canonicalize_mapping(
        by_country_raw,
        lambda value: _canonicalize_country_or_rest_of_world(value, contract),
        "revenue_splits.by_country",
        errors,
    )
    by_region = _canonicalize_mapping(
        by_region_raw,
        lambda value: canonicalize_region(value, contract),
        "revenue_splits.by_region",
        errors,
    )
    by_business = _canonicalize_mapping(
        by_business_raw,
        lambda value: canonicalize_industry(value, contract, "global"),
        "revenue_splits.by_business",
        errors,
    )
    revenue_splits["by_country"] = by_country
    revenue_splits["by_region"] = by_region
    revenue_splits["by_business"] = by_business

    if geography_mode == "incorporation" and (by_country or by_region):
        errors.append("Incorporation geography mode requires empty by_country and by_region mappings.")
    if geography_mode == "country" and (not by_country or by_region):
        errors.append("Country geography mode requires by_country only.")
    if geography_mode == "region" and (not by_region or by_country):
        errors.append("Region geography mode requires by_region only.")
    if len([name for name in by_country if name != "Rest of the World"]) > 11:
        errors.append("revenue_splits.by_country supports at most 11 named countries.")
    if not by_business or len(by_business) > 12:
        errors.append("revenue_splits.by_business must contain between 1 and 12 positive industries.")
    if period_revenue is not None:
        if by_country:
            _validate_reconciliation(errors, "Country revenue", by_country, period_revenue)
        if by_region:
            _validate_reconciliation(errors, "Regional revenue", by_region, period_revenue)
        if by_business:
            _validate_reconciliation(errors, "Business revenue", by_business, period_revenue)

    primary_industry = company_context.get("industry_global") or next(iter(by_business), None)
    if primary_industry:
        try:
            global_industry = canonicalize_industry(primary_industry, contract, "global")
            us_industry = canonicalize_industry(
                company_context.get("industry_us", global_industry),
                contract,
                "us",
            )
            company_context["industry_global"] = global_industry
            company_context["industry_us"] = us_industry
        except ContractValueError as error:
            errors.append(str(error))
    else:
        errors.append("company_context.industry_global is required.")

    cost_of_capital = _require_object(
        errors,
        "cost_of_capital_inputs",
        inputs.get("cost_of_capital_inputs"),
    )
    maturity = _require_number(
        errors,
        "cost_of_capital_inputs.average_maturity_of_debt_years",
        cost_of_capital.get("average_maturity_of_debt_years"),
    )
    debt_value = financial_data.get("Book_value_of_debt", {}).get("Most_Recent_12_months")
    if maturity is not None:
        if is_finite_number(debt_value) and debt_value > 0 and maturity <= 0:
            errors.append("Debt maturity must be greater than zero when debt is outstanding.")
        if maturity < 0:
            errors.append("Debt maturity cannot be negative.")
    debt_rating = cost_of_capital.get("debt_rating")
    if debt_rating and str(debt_rating).strip().upper() != "N/A":
        try:
            cost_of_capital["debt_rating"] = canonicalize_rating(debt_rating, contract)
        except ContractValueError as error:
            errors.append(str(error))
    else:
        cost_of_capital["debt_rating"] = "N/A"
        if cost_of_capital.get("synthetic_rating_company_type") not in {1, 2, 1.0, 2.0}:
            errors.append("synthetic_rating_company_type must be 1 or 2 when no actual rating is used.")
    direct_erp = cost_of_capital.get("direct_erp")
    if direct_erp is not None:
        direct_erp_number = _require_number(errors, "cost_of_capital_inputs.direct_erp", direct_erp)
        if direct_erp_number is not None and not 0 < direct_erp_number < 0.5:
            errors.append("cost_of_capital_inputs.direct_erp must be greater than 0 and less than 0.5.")
    rest_of_world_erp = cost_of_capital.get("rest_of_world_erp")
    if "Rest of the World" in by_country:
        rest_of_world_erp_number = _require_number(
            errors,
            "cost_of_capital_inputs.rest_of_world_erp",
            rest_of_world_erp,
        )
        if rest_of_world_erp_number is not None and not 0 < rest_of_world_erp_number < 0.5:
            errors.append("cost_of_capital_inputs.rest_of_world_erp must be greater than 0 and less than 0.5.")

    r_and_d = _require_object(errors, "r_and_d_details", inputs.get("r_and_d_details"))
    current_r_and_d = _require_number(errors, "r_and_d_details.current_year_expense", r_and_d.get("current_year_expense"))
    if current_r_and_d is not None and current_r_and_d < 0:
        errors.append("r_and_d_details.current_year_expense cannot be negative.")
    if current_r_and_d is not None and current_r_and_d > 0:
        years = r_and_d.get("amortization_period_years")
        if not isinstance(years, int) or not 1 <= years <= 10:
            errors.append("r_and_d_details.amortization_period_years must be an integer from 1 to 10.")
        history = _require_object(errors, "r_and_d_details.historical_expenses", r_and_d.get("historical_expenses"))
        if isinstance(years, int):
            for year_number in range(1, years + 1):
                key = f"Year_Minus_{year_number}"
                value = _require_number(errors, f"r_and_d_details.historical_expenses.{key}", history.get(key))
                if value is not None and value < 0:
                    errors.append(f"r_and_d_details.historical_expenses.{key} cannot be negative.")
        if company_context.get("industry_global"):
            r_and_d["industry_name"] = company_context["industry_global"]

    options = _require_object(errors, "employee_options", inputs.get("employee_options"))
    option_count = _require_number(errors, "employee_options.total_options_outstanding", options.get("total_options_outstanding"))
    if option_count is not None and option_count < 0:
        errors.append("employee_options.total_options_outstanding cannot be negative.")
    if option_count is not None and option_count > 0:
        for field in ("weighted_average_exercise_price", "average_maturity_years", "stock_price_standard_deviation"):
            value = _require_number(errors, f"employee_options.{field}", options.get(field))
            if value is not None and value <= 0:
                errors.append(f"employee_options.{field} must be greater than zero.")
        volatility = options.get("stock_price_standard_deviation")
        if is_finite_number(volatility) and volatility > 5:
            errors.append("employee_options.stock_price_standard_deviation must be expressed as a decimal.")

    lease = _require_object(errors, "operating_lease_details", inputs.get("operating_lease_details", {"capitalize": False}))
    if not isinstance(lease.get("capitalize"), bool):
        errors.append("operating_lease_details.capitalize must be boolean.")
    if lease.get("capitalize"):
        if lease.get("book_debt_excludes_operating_leases") is not True:
            errors.append("Capitalized leases require book_debt_excludes_operating_leases=true.")
        lease_expense = _require_number(
            errors,
            "operating_lease_details.current_year_expense",
            lease.get("current_year_expense"),
        )
        if lease_expense is not None and lease_expense < 0:
            errors.append("operating_lease_details.current_year_expense cannot be negative.")
        commitments = _require_object(
            errors,
            "operating_lease_details.commitments",
            lease.get("commitments"),
        )
        for key in ("year_1", "year_2", "year_3", "year_4", "year_5", "years_6_and_beyond"):
            value = _require_number(errors, f"operating_lease_details.commitments.{key}", commitments.get(key))
            if value is not None and value < 0:
                errors.append(f"operating_lease_details.commitments.{key} cannot be negative.")

    market = _require_object(errors, "market_inputs", inputs.get("market_inputs"))
    stock_price = _require_number(errors, "market_inputs.stock_price", market.get("stock_price"))
    if stock_price is not None and stock_price <= 0:
        errors.append("market_inputs.stock_price must be greater than zero.")
    riskfree_rate = _require_number(errors, "market_inputs.riskfree_rate", market.get("riskfree_rate"))
    if riskfree_rate is not None and not -0.05 < riskfree_rate < 0.25:
        errors.append("market_inputs.riskfree_rate must be between -0.05 and 0.25.")
    market_date = _parse_iso_date(errors, "market_inputs.as_of_date", market.get("as_of_date"))
    if valuation_date and market_date and market_date != valuation_date:
        findings.append(
            {
                "severity": "warning",
                "code": "market_date_differs",
                "message": "market_inputs.as_of_date differs from company_context.valuation_date.",
            }
        )

    base_case = _require_object(errors, "base_case_assumptions", inputs.get("base_case_assumptions"))
    for field in BASE_CASE_FIELDS:
        _require_number(errors, f"base_case_assumptions.{field}", base_case.get(field))
    convergence = base_case.get("margin_convergence_year")
    if is_finite_number(convergence) and (int(convergence) != convergence or not 1 <= convergence <= 10):
        errors.append("base_case_assumptions.margin_convergence_year must be an integer from 1 to 10.")
    for field in ("sales_to_capital_years_1_5", "sales_to_capital_years_6_10"):
        value = base_case.get(field)
        if is_finite_number(value) and value <= 0:
            errors.append(f"base_case_assumptions.{field} must be greater than zero.")
    for field in ("revenue_growth_next_year", "revenue_cagr_years_2_5", "operating_margin_next_year", "target_operating_margin"):
        value = base_case.get(field)
        if is_finite_number(value) and not -1 < value <= 1:
            errors.append(f"base_case_assumptions.{field} must be greater than -1 and at most 1.")

    story = _require_object(errors, "company_story", inputs.get("company_story"))
    for field in STORY_FIELDS:
        _require_string(errors, f"company_story.{field}", story.get(field))

    advanced = _require_object(errors, "advanced_assumptions", inputs.get("advanced_assumptions", {}))
    defaults_applied: list[str] = []
    for field, default in ADVANCED_DEFAULTS.items():
        if field not in advanced:
            advanced[field] = default
            defaults_applied.append(field)
    for field in ("terminal_cost_of_capital_override", "stable_return_on_capital_override", "stable_riskfree_rate_override", "terminal_growth_override"):
        value = advanced.get(field)
        if value is not None:
            number = _require_number(errors, f"advanced_assumptions.{field}", value)
            if number is not None and not -0.25 < number < 1:
                errors.append(f"advanced_assumptions.{field} is outside the supported range.")
    failure_probability = _require_number(
        errors,
        "advanced_assumptions.probability_of_failure",
        advanced.get("probability_of_failure"),
    )
    if failure_probability is not None and not 0 <= failure_probability <= 1:
        errors.append("advanced_assumptions.probability_of_failure must be between 0 and 1.")
    if advanced.get("failure_proceeds_basis") not in {"B", "V"}:
        errors.append("advanced_assumptions.failure_proceeds_basis must be B or V.")
    failure_proceeds = _require_number(
        errors,
        "advanced_assumptions.failure_proceeds_percent",
        advanced.get("failure_proceeds_percent"),
    )
    if failure_proceeds is not None and not 0 <= failure_proceeds <= 1:
        errors.append("advanced_assumptions.failure_proceeds_percent must be between 0 and 1.")
    lag = advanced.get("reinvestment_lag_years")
    if not isinstance(lag, int) or not 0 <= lag <= 3:
        errors.append("advanced_assumptions.reinvestment_lag_years must be an integer from 0 to 3.")
    if not isinstance(advanced.get("hold_effective_tax_rate"), bool):
        errors.append("advanced_assumptions.hold_effective_tax_rate must be boolean.")
    for field in ("net_operating_loss", "trapped_cash"):
        number = _require_number(errors, f"advanced_assumptions.{field}", advanced.get(field))
        if number is not None and number < 0:
            errors.append(f"advanced_assumptions.{field} cannot be negative.")
    trapped_tax = _require_number(
        errors,
        "advanced_assumptions.trapped_cash_tax_rate",
        advanced.get("trapped_cash_tax_rate"),
    )
    if trapped_tax is not None and not 0 <= trapped_tax <= 1:
        errors.append("advanced_assumptions.trapped_cash_tax_rate must be between 0 and 1.")
    inputs["advanced_assumptions"] = advanced
    inputs["defaults_applied"] = defaults_applied

    missing_documents = inputs.get("missing_documents_required", [])
    if not isinstance(missing_documents, list):
        errors.append("missing_documents_required must be an array.")
    elif any(str(document).strip() for document in missing_documents):
        errors.append("missing_documents_required is not empty.")

    supplied_evidence = inputs.get("source_evidence")
    if not isinstance(supplied_evidence, list) or not supplied_evidence:
        errors.append("source_evidence must contain evidence records.")
        if not isinstance(supplied_evidence, list):
            inputs["source_evidence"] = []

    _append_model_default_evidence(inputs, defaults_applied)
    _validate_source_evidence(inputs, errors, allow_legacy)
    if company_context.get("industry_global"):
        _assumption_findings(inputs, contract, findings)

    inputs["validation_findings"] = findings
    inputs["workbook_contract"] = {
        "template_sha256": contract.template_sha256,
        "formula_fingerprint": contract.formula_fingerprint,
    }
    if errors:
        raise InputValidationError(errors, findings)
    return inputs


def main() -> None:
    parser = argparse.ArgumentParser(description="Validate and normalize v2 valuation inputs")
    parser.add_argument("--inputs", required=True)
    parser.add_argument("--template", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--report")
    parser.add_argument("--allow-legacy", action="store_true")
    args = parser.parse_args()

    with Path(args.inputs).open(encoding="utf-8") as source:
        raw_inputs = json.load(source)
    contract = load_workbook_contract(args.template)
    try:
        normalized = normalize_and_validate_inputs(
            raw_inputs,
            contract,
            allow_legacy=args.allow_legacy,
        )
    except InputValidationError as error:
        report = {"status": "failed", "errors": error.errors, "findings": error.findings}
        if args.report:
            atomic_write_json(args.report, report)
        print(json.dumps(report, indent=2), file=sys.stderr)
        raise SystemExit(1) from error

    atomic_write_json(args.output, normalized)
    report = {"status": "valid", "errors": [], "findings": normalized["validation_findings"]}
    if args.report:
        atomic_write_json(args.report, report)
    print(json.dumps(report, indent=2))


if __name__ == "__main__":  # pragma: no cover - exercised through main()
    main()
