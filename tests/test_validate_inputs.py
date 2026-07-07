from __future__ import annotations

import copy
import unittest

from validate_inputs import (
    InputValidationError,
    _append_model_default_evidence,
    _assumption_findings,
    _flatten_manifest_filings,
    _validate_source_evidence,
    normalize_and_validate_inputs,
)
from workbook_contract import load_workbook_contract

from tests.helpers import (
    TEMPLATE,
    analyst_evidence,
    filing_evidence,
    load_valid_inputs,
    replace_evidence,
)


class ValidateInputsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.contract = load_workbook_contract(TEMPLATE)

    def assert_invalid(self, inputs: dict, text: str, *, allow_legacy: bool = False) -> None:
        with self.assertRaises(InputValidationError) as context:
            normalize_and_validate_inputs(inputs, self.contract, allow_legacy=allow_legacy)
        self.assertIn(text, "\n".join(context.exception.errors))

    def test_valid_input_normalizes_defaults_and_warnings(self) -> None:
        normalized = normalize_and_validate_inputs(load_valid_inputs(), self.contract)
        self.assertEqual(normalized["schema_version"], "2.0")
        self.assertEqual(normalized["advanced_assumptions"]["reinvestment_lag_years"], 1)
        self.assertIn("reinvestment_lag_years", normalized["defaults_applied"])
        self.assertTrue(
            any(
                record["source_type"] == "model_default"
                for record in normalized["source_evidence"]
            )
        )
        self.assertTrue(normalized["validation_findings"])

    def test_legacy_input_requires_explicit_flag(self) -> None:
        inputs = load_valid_inputs()
        del inputs["schema_version"]
        self.assert_invalid(inputs, "schema_version")
        normalized = normalize_and_validate_inputs(inputs, self.contract, allow_legacy=True)
        self.assertEqual(normalized["schema_version"], "2.0")
        self.assertEqual(normalized["validation_findings"][0]["code"], "legacy_input_upgraded")

    def test_rejects_bad_company_metadata_and_dates(self) -> None:
        cases = (
            (("company_context", "company_name"), "", "company_context.company_name"),
            (("company_context", "valuation_date"), "not-a-date", "YYYY-MM-DD"),
            (("company_context", "valuation_date"), "2999-01-01", "cannot be in the future"),
            (("company_context", "ticker"), "X" * 21, "at most 20"),
            (("company_context", "cik"), "12", "exactly 10 digits"),
            (("company_context", "currency"), "usd", "three-letter uppercase"),
            (("company_context", "units"), "crores", "must be units"),
        )
        for path, value, expected in cases:
            with self.subTest(path=path):
                inputs = load_valid_inputs()
                inputs[path[0]][path[1]] = value
                self.assert_invalid(inputs, expected)

    def test_rejects_non_finite_and_invalid_financial_values(self) -> None:
        inputs = load_valid_inputs()
        inputs["financial_data"]["Revenues"]["Most_Recent_12_months"] = float("nan")
        replace_evidence(
            inputs,
            "financial_data.Revenues.Most_Recent_12_months",
            filing_evidence("financial_data.Revenues.Most_Recent_12_months", float("nan")),
        )
        self.assert_invalid(inputs, "must be a finite number")

        inputs = load_valid_inputs()
        inputs["financial_data"]["Book_value_of_debt"]["Most_Recent_12_months"] = -1
        replace_evidence(
            inputs,
            "financial_data.Book_value_of_debt.Most_Recent_12_months",
            filing_evidence("financial_data.Book_value_of_debt.Most_Recent_12_months", -1),
        )
        self.assert_invalid(inputs, "cannot be negative")

    def test_rejects_silent_split_failures(self) -> None:
        for value, expected in (("bad", "finite number"), (0, "greater than zero"), (-1, "greater than zero")):
            with self.subTest(value=value):
                inputs = load_valid_inputs()
                inputs["revenue_splits"]["by_business"]["Food Processing"] = value
                self.assert_invalid(inputs, expected)

    def test_country_region_and_rest_of_world_modes(self) -> None:
        country = load_valid_inputs()
        country["revenue_splits"].update(
            {
                "geography_mode": "country",
                "by_country": {"United States": 1000.0},
                "by_region": {},
            }
        )
        country["source_evidence"].append(
            filing_evidence("revenue_splits.by_country.United States", 1000.0)
        )
        normalized = normalize_and_validate_inputs(country, self.contract)
        self.assertEqual(normalized["revenue_splits"]["by_country"], {"United States": 1000.0})

        region = load_valid_inputs()
        region["revenue_splits"].update(
            {
                "geography_mode": "region",
                "by_country": {},
                "by_region": {"North America": 1000.0},
            }
        )
        region["source_evidence"].append(
            filing_evidence("revenue_splits.by_region.North America", 1000.0)
        )
        self.assertEqual(
            normalize_and_validate_inputs(region, self.contract)["revenue_splits"]["by_region"],
            {"North America": 1000.0},
        )

        row = copy.deepcopy(country)
        row["revenue_splits"]["by_country"] = {"United States": 900.0, "Rest of the World": 100.0}
        replace_evidence(
            row,
            "revenue_splits.by_country.United States",
            filing_evidence("revenue_splits.by_country.United States", 900.0),
        )
        row["source_evidence"].append(
            filing_evidence("revenue_splits.by_country.Rest of the World", 100.0)
        )
        self.assert_invalid(row, "rest_of_world_erp")
        row["cost_of_capital_inputs"]["rest_of_world_erp"] = 0.07
        row["source_evidence"].append(
            analyst_evidence("cost_of_capital_inputs.rest_of_world_erp", 0.07)
        )
        normalize_and_validate_inputs(row, self.contract)

    def test_rejects_geography_mode_mismatch_and_reconciliation(self) -> None:
        inputs = load_valid_inputs()
        inputs["revenue_splits"]["geography_mode"] = "country"
        self.assert_invalid(inputs, "requires by_country only")
        inputs = load_valid_inputs()
        inputs["revenue_splits"]["by_business"]["Food Processing"] = 500
        replace_evidence(
            inputs,
            "revenue_splits.by_business.Food Processing",
            filing_evidence("revenue_splits.by_business.Food Processing", 500),
        )
        self.assert_invalid(inputs, "does not reconcile")

    def test_r_and_d_options_and_leases(self) -> None:
        inputs = load_valid_inputs()
        inputs["r_and_d_details"] = {
            "current_year_expense": 30.0,
            "amortization_period_years": 3,
            "historical_expenses": {
                "Year_Minus_1": 25.0,
                "Year_Minus_2": 20.0,
                "Year_Minus_3": 15.0,
            },
        }
        for metric, value in (
            ("r_and_d_details.current_year_expense", 30.0),
            ("r_and_d_details.historical_expenses.Year_Minus_1", 25.0),
            ("r_and_d_details.historical_expenses.Year_Minus_2", 20.0),
            ("r_and_d_details.historical_expenses.Year_Minus_3", 15.0),
        ):
            inputs["source_evidence"].append(filing_evidence(metric, value))
        normalized = normalize_and_validate_inputs(inputs, self.contract)
        self.assertEqual(normalized["r_and_d_details"]["industry_name"], "Food Processing")

        inputs = load_valid_inputs()
        inputs["employee_options"] = {
            "total_options_outstanding": 5.0,
            "weighted_average_exercise_price": 10.0,
            "average_maturity_years": 4.0,
            "stock_price_standard_deviation": 0.3,
        }
        for field, value in inputs["employee_options"].items():
            inputs["source_evidence"].append(
                filing_evidence(f"employee_options.{field}", value)
            )
        normalize_and_validate_inputs(inputs, self.contract)

        inputs = load_valid_inputs()
        inputs["operating_lease_details"] = {
            "capitalize": True,
            "book_debt_excludes_operating_leases": True,
            "current_year_expense": 12.0,
            "commitments": {
                "year_1": 10.0,
                "year_2": 9.0,
                "year_3": 8.0,
                "year_4": 7.0,
                "year_5": 6.0,
                "years_6_and_beyond": 20.0,
            },
        }
        inputs["source_evidence"].append(
            filing_evidence("operating_lease_details.current_year_expense", 12.0)
        )
        for key, value in inputs["operating_lease_details"]["commitments"].items():
            inputs["source_evidence"].append(
                filing_evidence(f"operating_lease_details.commitments.{key}", value)
            )
        normalize_and_validate_inputs(inputs, self.contract)

    def test_rejects_incomplete_lease_and_bad_options(self) -> None:
        inputs = load_valid_inputs()
        inputs["operating_lease_details"] = {"capitalize": True}
        self.assert_invalid(inputs, "book_debt_excludes_operating_leases")
        inputs = load_valid_inputs()
        inputs["employee_options"] = {
            "total_options_outstanding": 1.0,
            "weighted_average_exercise_price": 1.0,
            "average_maturity_years": 1.0,
            "stock_price_standard_deviation": 6.0,
        }
        for field, value in inputs["employee_options"].items():
            inputs["source_evidence"].append(
                filing_evidence(f"employee_options.{field}", value)
            )
        self.assert_invalid(inputs, "expressed as a decimal")

    def test_evidence_integrity(self) -> None:
        inputs = load_valid_inputs()
        inputs["source_evidence"][0]["value"] = 999
        self.assert_invalid(inputs, "does not match")
        inputs = load_valid_inputs()
        inputs["source_evidence"].append(copy.deepcopy(inputs["source_evidence"][0]))
        self.assert_invalid(inputs, "duplicate metric")
        inputs = load_valid_inputs()
        inputs["source_evidence"][0]["sha256"] = "b" * 64
        self.assert_invalid(inputs, "does not match filing_manifest")
        inputs = load_valid_inputs()
        del inputs["source_evidence"][0]["source_type"]
        self.assert_invalid(inputs, "source_type")

    def test_cost_market_advanced_and_missing_document_rules(self) -> None:
        cases = []
        inputs = load_valid_inputs()
        inputs["cost_of_capital_inputs"]["direct_erp"] = 0.8
        replace_evidence(
            inputs,
            "cost_of_capital_inputs.direct_erp",
            analyst_evidence("cost_of_capital_inputs.direct_erp", 0.8),
        )
        cases.append((inputs, "direct_erp"))

        inputs = load_valid_inputs()
        inputs["market_inputs"]["riskfree_rate"] = 1.0
        replace_evidence(
            inputs,
            "market_inputs.riskfree_rate",
            {
                "metric": "market_inputs.riskfree_rate",
                "source_type": "market",
                "source_url": "https://example.com",
                "period": "2026-01-02",
                "calculation": "Test",
                "value": 1.0,
                "units": "decimal",
            },
        )
        cases.append((inputs, "riskfree_rate"))

        inputs = load_valid_inputs()
        inputs["advanced_assumptions"] = {"reinvestment_lag_years": 4}
        cases.append((inputs, "reinvestment_lag_years"))

        inputs = load_valid_inputs()
        inputs["missing_documents_required"] = ["Debt footnote"]
        cases.append((inputs, "missing_documents_required"))

        for inputs, expected in cases:
            with self.subTest(expected=expected):
                self.assert_invalid(inputs, expected)

    def test_zero_debt_allows_zero_maturity_and_actual_rating(self) -> None:
        inputs = load_valid_inputs()
        inputs["financial_data"]["Book_value_of_debt"]["Most_Recent_12_months"] = 0.0
        inputs["cost_of_capital_inputs"]["average_maturity_of_debt_years"] = 0.0
        replace_evidence(
            inputs,
            "financial_data.Book_value_of_debt.Most_Recent_12_months",
            filing_evidence("financial_data.Book_value_of_debt.Most_Recent_12_months", 0.0),
        )
        replace_evidence(
            inputs,
            "cost_of_capital_inputs.average_maturity_of_debt_years",
            filing_evidence("cost_of_capital_inputs.average_maturity_of_debt_years", 0.0, "years"),
        )
        normalize_and_validate_inputs(inputs, self.contract)

        inputs = load_valid_inputs()
        inputs["cost_of_capital_inputs"]["debt_rating"] = "bbb"
        normalize_and_validate_inputs(inputs, self.contract)

    def test_defensive_validation_matrix(self) -> None:
        cases = []

        def add(path: tuple[str, ...], value, expected: str) -> None:
            inputs = load_valid_inputs()
            target = inputs
            for component in path[:-1]:
                target = target[component]
            target[path[-1]] = value
            cases.append((inputs, expected))

        add(("company_context", "valuation_date"), "", "valuation_date")
        add(("company_context", "country_of_incorporation"), "Not a country", "country")
        add(("company_context", "country_of_incorporation"), "", "country_of_incorporation")
        add(("filing_manifest", "filings"), [], "at least one filing")
        add(("financial_data", "Revenues", "Most_Recent_12_months"), 0, "greater than zero")
        add(("single_value_metrics", "Years_since_last_10K"), 0.3, "must be 0.25")
        add(("single_value_metrics", "Number_of_shares_outstanding"), 0, "greater than zero")
        add(("single_value_metrics", "Effective_tax_rate"), 1.5, "between 0 and 1")
        add(("revenue_splits", "period_revenue"), 0, "greater than zero")
        add(("revenue_splits", "geography_mode"), "wrong", "geography_mode")
        add(("company_context", "industry_global"), "Not an industry", "global industry")
        add(("cost_of_capital_inputs", "average_maturity_of_debt_years"), 0, "maturity")
        add(("cost_of_capital_inputs", "average_maturity_of_debt_years"), -1, "cannot be negative")
        add(("cost_of_capital_inputs", "debt_rating"), "not-rated", "debt rating")
        add(("cost_of_capital_inputs", "synthetic_rating_company_type"), 3, "must be 1 or 2")
        add(("r_and_d_details", "current_year_expense"), -1, "cannot be negative")
        add(("employee_options", "total_options_outstanding"), -1, "cannot be negative")
        add(("operating_lease_details", "capitalize"), "yes", "must be boolean")
        add(("market_inputs", "stock_price"), 0, "stock_price")
        add(("base_case_assumptions", "margin_convergence_year"), 2.5, "integer from 1 to 10")
        add(("base_case_assumptions", "sales_to_capital_years_1_5"), 0, "greater than zero")
        add(("base_case_assumptions", "revenue_growth_next_year"), 2, "at most 1")
        add(("advanced_assumptions",), {"terminal_growth_override": 2}, "outside the supported range")
        add(("advanced_assumptions",), {"probability_of_failure": 2}, "probability_of_failure")
        add(("advanced_assumptions",), {"failure_proceeds_basis": "X"}, "failure_proceeds_basis")
        add(("advanced_assumptions",), {"failure_proceeds_percent": 2}, "failure_proceeds_percent")
        add(("advanced_assumptions",), {"hold_effective_tax_rate": "no"}, "must be boolean")
        add(("advanced_assumptions",), {"net_operating_loss": -1}, "cannot be negative")
        add(("advanced_assumptions",), {"trapped_cash_tax_rate": 2}, "trapped_cash_tax_rate")
        add(("missing_documents_required",), "not-a-list", "must be an array")

        for inputs, expected in cases:
            with self.subTest(expected=expected):
                self.assert_invalid(inputs, expected)

    def test_complex_defensive_paths(self) -> None:
        inputs = load_valid_inputs()
        inputs["revenue_splits"]["by_business"] = {"Not an industry": 1000}
        self.assert_invalid(inputs, "global industry")

        inputs = load_valid_inputs()
        inputs["revenue_splits"].update(
            {
                "geography_mode": "incorporation",
                "by_country": {"United States": 1000},
            }
        )
        inputs["source_evidence"].append(
            filing_evidence("revenue_splits.by_country.United States", 1000)
        )
        self.assert_invalid(inputs, "Incorporation geography")

        inputs = load_valid_inputs()
        inputs["revenue_splits"].update(
            {
                "geography_mode": "region",
                "by_country": {"United States": 1000},
                "by_region": {},
            }
        )
        inputs["source_evidence"].append(
            filing_evidence("revenue_splits.by_country.United States", 1000)
        )
        self.assert_invalid(inputs, "Region geography")

        inputs = load_valid_inputs()
        countries = list(self.contract.countries[:12])
        inputs["revenue_splits"].update(
            {
                "geography_mode": "country",
                "by_country": {name: 1000 / 12 for name in countries},
                "by_region": {},
            }
        )
        for name, value in inputs["revenue_splits"]["by_country"].items():
            inputs["source_evidence"].append(
                filing_evidence(f"revenue_splits.by_country.{name}", value)
            )
        self.assert_invalid(inputs, "at most 11")

        inputs = load_valid_inputs()
        inputs["r_and_d_details"] = {
            "current_year_expense": 1,
            "amortization_period_years": "three",
            "historical_expenses": {},
        }
        self.assert_invalid(inputs, "integer from 1 to 10")

        inputs = load_valid_inputs()
        inputs["r_and_d_details"] = {
            "current_year_expense": 1,
            "amortization_period_years": 1,
            "historical_expenses": {"Year_Minus_1": -1},
        }
        inputs["source_evidence"].extend(
            (
                filing_evidence("r_and_d_details.current_year_expense", 1),
                filing_evidence("r_and_d_details.historical_expenses.Year_Minus_1", -1),
            )
        )
        self.assert_invalid(inputs, "cannot be negative")

        inputs = load_valid_inputs()
        inputs["employee_options"] = {
            "total_options_outstanding": 1,
            "weighted_average_exercise_price": 0,
            "average_maturity_years": 1,
            "stock_price_standard_deviation": 0.2,
        }
        for field, value in inputs["employee_options"].items():
            inputs["source_evidence"].append(
                filing_evidence(f"employee_options.{field}", value)
            )
        self.assert_invalid(inputs, "weighted_average_exercise_price")

        inputs = load_valid_inputs()
        inputs["operating_lease_details"] = {
            "capitalize": True,
            "book_debt_excludes_operating_leases": True,
            "current_year_expense": -1,
            "commitments": {
                "year_1": -1,
                "year_2": 1,
                "year_3": 1,
                "year_4": 1,
                "year_5": 1,
                "years_6_and_beyond": 1,
            },
        }
        self.assert_invalid(inputs, "cannot be negative")

    def test_evidence_defensive_paths(self) -> None:
        inputs = load_valid_inputs()
        inputs["source_evidence"] = []
        self.assert_invalid(inputs, "must contain evidence records")

        inputs = load_valid_inputs()
        inputs["source_evidence"] = "not-an-array"
        self.assert_invalid(inputs, "must contain evidence records")

        inputs = load_valid_inputs()
        inputs["source_evidence"].append("not-an-object")
        self.assert_invalid(inputs, "must be an object")

        inputs = load_valid_inputs()
        inputs["source_evidence"][0]["source_type"] = "unknown"
        self.assert_invalid(inputs, "must be one of")

        inputs = load_valid_inputs()
        inputs["source_evidence"][0]["sha256"] = "short"
        self.assert_invalid(inputs, "64-character")

        inputs = load_valid_inputs()
        inputs["source_evidence"][0]["accession_number"] = "missing"
        self.assert_invalid(inputs, "not present in filing_manifest")

        inputs = load_valid_inputs()
        inputs["source_evidence"][0]["source_url"] = "https://wrong.example"
        self.assert_invalid(inputs, "source_url does not match")

        inputs = load_valid_inputs()
        inputs["source_evidence"][0]["metric"] = "missing.path"
        self.assert_invalid(inputs, "does not resolve")

        inputs = load_valid_inputs()
        del inputs["source_evidence"][0]
        self.assert_invalid(inputs, "source_evidence is missing")

    def test_internal_helpers_cover_legacy_and_distribution_edges(self) -> None:
        manifest = {
            "filings": [None, {"accession_number": "a"}],
            "latest_four_10ks": [{"accession_number": "a"}],
            "latest_two_10qs": [{"accession_number": "b"}],
        }
        self.assertEqual(
            [item["accession_number"] for item in _flatten_manifest_filings(manifest)],
            ["a", "b"],
        )
        self.assertEqual(_flatten_manifest_filings(None), [])

        for split_name, expected_mode in (("by_country", "country"), ("by_region", "region")):
            inputs = load_valid_inputs()
            del inputs["schema_version"]
            del inputs["revenue_splits"]["geography_mode"]
            inputs["revenue_splits"][split_name] = {
                "United States" if split_name == "by_country" else "North America": 1000
            }
            other = "by_region" if split_name == "by_country" else "by_country"
            inputs["revenue_splits"][other] = {}
            metric = (
                "revenue_splits.by_country.United States"
                if split_name == "by_country"
                else "revenue_splits.by_region.North America"
            )
            inputs["source_evidence"].append(filing_evidence(metric, 1000))
            normalized = normalize_and_validate_inputs(
                inputs,
                self.contract,
                allow_legacy=True,
            )
            self.assertEqual(normalized["revenue_splits"]["geography_mode"], expected_mode)

        inputs = load_valid_inputs()
        inputs["advanced_assumptions"] = {}
        inputs["source_evidence"].append(
            {
                "metric": "advanced_assumptions.reinvestment_lag_years",
                "source_type": "model_default",
                "period": "2026-01-02",
                "calculation": "Existing default",
                "value": 1,
                "units": "model setting",
                "rationale": "Existing.",
            }
        )
        normalized = normalize_and_validate_inputs(inputs, self.contract)
        matching = [
            item
            for item in normalized["source_evidence"]
            if item["metric"] == "advanced_assumptions.reinvestment_lag_years"
        ]
        self.assertEqual(len(matching), 1)

        findings: list[dict[str, str]] = []
        contract = copy.copy(self.contract)
        object.__setattr__(contract, "industry_distributions", {})
        _assumption_findings(load_valid_inputs(), contract, findings)
        self.assertEqual(findings[0]["code"], "industry_distribution_missing")

    def test_remaining_validation_control_flow(self) -> None:
        with self.assertRaises(InputValidationError):
            normalize_and_validate_inputs([], self.contract)

        legacy = load_valid_inputs()
        del legacy["schema_version"]
        del legacy["revenue_splits"]["geography_mode"]
        normalized = normalize_and_validate_inputs(
            legacy,
            self.contract,
            allow_legacy=True,
        )
        self.assertEqual(normalized["revenue_splits"]["geography_mode"], "incorporation")

        errors: list[str] = []
        _validate_source_evidence({"source_evidence": []}, errors, False)
        self.assertIn("must contain evidence records", errors[0])

        missing_metric = load_valid_inputs()
        missing_metric["source_evidence"].append(
            {
                "source_type": "analyst_assumption",
                "calculation": "Test",
                "units": "decimal",
            }
        )
        self.assert_invalid(missing_metric, ".metric must be a non-empty string")

        legacy_manifest = load_valid_inputs()
        legacy_manifest["filing_manifest"] = {"filings": []}
        errors = []
        _validate_source_evidence(legacy_manifest, errors, True)
        self.assertFalse(
            any("accession_number is not present" in error for error in errors)
        )

        no_period_revenue = load_valid_inputs()
        no_period_revenue["revenue_splits"]["period_revenue"] = None
        self.assert_invalid(no_period_revenue, "period_revenue must be a finite number")

        no_industry = load_valid_inputs()
        no_industry["company_context"]["industry_global"] = ""
        no_industry["company_context"]["industry_us"] = ""
        no_industry["revenue_splits"]["by_business"] = {}
        no_industry["r_and_d_details"] = {
            "current_year_expense": 1,
            "amortization_period_years": 1,
            "historical_expenses": {"Year_Minus_1": 1},
        }
        self.assert_invalid(no_industry, "industry_global is required")

        no_maturity = load_valid_inputs()
        no_maturity["cost_of_capital_inputs"]["average_maturity_of_debt_years"] = None
        self.assert_invalid(no_maturity, "must be a finite number")

        no_direct_erp = load_valid_inputs()
        no_direct_erp["cost_of_capital_inputs"]["direct_erp"] = None
        replace_evidence(
            no_direct_erp,
            "cost_of_capital_inputs.direct_erp",
            analyst_evidence("cost_of_capital_inputs.direct_erp", None),
        )
        normalize_and_validate_inputs(no_direct_erp, self.contract)

        valid_override = load_valid_inputs()
        valid_override["advanced_assumptions"] = {
            "terminal_growth_override": 0.03,
        }
        normalize_and_validate_inputs(valid_override, self.contract)

        invalid_override_type = load_valid_inputs()
        invalid_override_type["advanced_assumptions"] = {
            "terminal_growth_override": "three percent",
        }
        self.assert_invalid(invalid_override_type, "must be a finite number")

        row = load_valid_inputs()
        row["revenue_splits"].update(
            {
                "geography_mode": "country",
                "by_country": {"Rest of the World": 1000},
                "by_region": {},
            }
        )
        row["cost_of_capital_inputs"]["rest_of_world_erp"] = 0.8
        row["source_evidence"].extend(
            (
                filing_evidence("revenue_splits.by_country.Rest of the World", 1000),
                analyst_evidence("cost_of_capital_inputs.rest_of_world_erp", 0.8),
            )
        )
        self.assert_invalid(row, "rest_of_world_erp must be greater than 0")


if __name__ == "__main__":
    unittest.main()
