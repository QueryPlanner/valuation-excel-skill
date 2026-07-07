from __future__ import annotations

import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import requests

from get_financial_reports import (
    SEC_COMPANYFACTS_URL,
    SEC_SUBMISSIONS_URL,
    SEC_TICKERS_URL,
    _build_filing_url,
    _collect_filing_rows,
    _default_output_path,
    _download_companyfacts,
    _download_filing,
    _filing_rows,
    _get_json,
    _load_environment,
    _request,
    _resolve_company,
    _safe_filename,
    _sec_headers,
    _select_filings,
    get_financial_reports,
)


class FakeResponse:
    def __init__(
        self,
        *,
        status_code: int = 200,
        json_value=None,
        content: bytes = b"",
        headers: dict | None = None,
    ):
        self.status_code = status_code
        self._json_value = json_value
        self.content = content
        self.headers = headers or {}

    def raise_for_status(self) -> None:
        if self.status_code >= 400:
            raise requests.HTTPError(f"HTTP {self.status_code}")

    def json(self):
        return self._json_value


class FakeSession:
    def __init__(self, responses: dict[str, list[FakeResponse] | FakeResponse]):
        self.responses = responses
        self.headers: dict[str, str] = {}
        self.calls: list[tuple[str, int]] = []

    def get(self, url: str, timeout: int):
        self.calls.append((url, timeout))
        response = self.responses[url]
        if isinstance(response, list):
            return response.pop(0)
        return response


class SecReportTests(unittest.TestCase):
    def test_headers_environment_and_company_resolution(self) -> None:
        self.assertIn("User-Agent", _sec_headers("App test@example.com"))
        with self.assertRaises(ValueError):
            _sec_headers("anonymous")
        with self.assertRaises(FileNotFoundError):
            _load_environment("/definitely/missing/.env")

        data = {
            "0": {"ticker": "ABC", "title": "ABC Corporation", "cik_str": 1},
            "1": {"ticker": "ABD", "title": "ABC Holdings", "cik_str": 2},
        }
        self.assertEqual(_resolve_company(data, "ABC")[0], "0000000001")
        self.assertEqual(_resolve_company(data, "ABC Corporation")[2], "ABC")
        with self.assertRaises(ValueError):
            _resolve_company(data, "AB")
        with self.assertRaises(ValueError):
            _resolve_company(data, "missing")

    def test_request_retries_and_raises(self) -> None:
        sleeps: list[float] = []
        session = FakeSession(
            {
                "url": [
                    FakeResponse(status_code=500),
                    FakeResponse(status_code=429, headers={"Retry-After": "1"}),
                    FakeResponse(status_code=200, content=b"ok"),
                ]
            }
        )
        response = _request(session, "url", timeout=1, sleep=sleeps.append)
        self.assertEqual(response.content, b"ok")
        self.assertEqual(sleeps, [0.5, 1.0, 0.12])

        failing = FakeSession({"url": [FakeResponse(status_code=500) for _ in range(4)]})
        with self.assertRaises(requests.HTTPError):
            _request(failing, "url", timeout=1, sleep=lambda _: None)

    def test_filing_rows_and_selection(self) -> None:
        columns = {
            "form": ["10-K", "8-K"],
            "accessionNumber": ["1-1", "1-2"],
            "primaryDocument": ["a.htm", "b.htm"],
            "filingDate": ["2026-01-01", "2026-01-02"],
            "reportDate": ["2025-12-31", "2026-01-01"],
        }
        rows = _filing_rows(columns, {"10-K"})
        self.assertEqual(len(rows), 1)
        selected, shortages = _select_filings("0000000001", rows, {"10-K": 2, "10-Q": 1})
        self.assertEqual(len(selected["10-K"]), 1)
        self.assertEqual(shortages, {"10-K": 1, "10-Q": 1})

        bad = dict(columns)
        bad["form"] = ["10-K"]
        with self.assertRaises(ValueError):
            _filing_rows(bad, {"10-K"})
        with self.assertRaises(ValueError):
            _filing_rows({"form": []}, {"10-K"})

    def test_get_reports_returns_shortages_without_rejecting_young_issuer(self) -> None:
        ticker_json = {
            "0": {"ticker": "ABC", "title": "ABC Corporation", "cik_str": 1}
        }
        submissions_url = f"{SEC_SUBMISSIONS_URL}/CIK0000000001.json"
        submissions = {
            "filings": {
                "recent": {
                    "form": ["10-K", "10-Q"],
                    "accessionNumber": ["0000000001-26-000001", "0000000001-26-000002"],
                    "primaryDocument": ["annual.htm", "quarter.htm"],
                    "filingDate": ["2026-02-01", "2026-05-01"],
                    "reportDate": ["2025-12-31", "2026-03-31"],
                },
                "files": [],
            }
        }
        session = FakeSession(
            {
                SEC_TICKERS_URL: FakeResponse(json_value=ticker_json),
                submissions_url: FakeResponse(json_value=submissions),
            }
        )
        manifest = get_financial_reports(
            "ABC",
            user_agent="App test@example.com",
            annual_count=2,
            quarterly_count=2,
            include_companyfacts=False,
            session=session,
            sleep=lambda _: None,
        )
        self.assertEqual(manifest["shortages"], {"10-K": 1, "10-Q": 1})
        self.assertEqual(len(manifest["filings"]), 2)
        self.assertIsNone(manifest["companyfacts"])
        with self.assertRaises(ValueError):
            get_financial_reports(
                "ABC",
                user_agent="App test@example.com",
                annual_count=0,
                session=session,
            )

    def test_downloads_hash_html_and_companyfacts(self) -> None:
        filing = {
            "form": "10-K",
            "accession_number": "0000000001-26-000001",
            "primary_document": "annual.htm",
            "filing_date": "2026-02-01",
            "report_date": "2025-12-31",
            "html_url": "https://example.com/annual.htm",
        }
        html = b"<!doctype html><html><body>Filing</body></html>"
        companyfacts_url = f"{SEC_COMPANYFACTS_URL}/CIK0000000001.json"
        session = FakeSession(
            {
                filing["html_url"]: FakeResponse(
                    content=html,
                    headers={"Content-Type": "text/html"},
                ),
                companyfacts_url: FakeResponse(
                    content=json.dumps({"facts": {}}).encode(),
                    headers={"Content-Type": "application/json"},
                ),
            }
        )
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary)
            downloaded = _download_filing(session, filing, directory, sleep=lambda _: None)
            self.assertTrue(Path(downloaded["local_path"]).is_file())
            self.assertEqual(len(downloaded["sha256"]), 64)
            facts = _download_companyfacts(
                session,
                "0000000001",
                directory / "facts.json",
                sleep=lambda _: None,
            )
            self.assertTrue(Path(facts["local_path"]).is_file())

        bad_session = FakeSession(
            {
                filing["html_url"]: FakeResponse(
                    content=b"binary",
                    headers={"Content-Type": "application/octet-stream"},
                )
            }
        )
        with tempfile.TemporaryDirectory() as temporary:
            with self.assertRaises(ValueError):
                _download_filing(
                    bad_session,
                    filing,
                    Path(temporary),
                    sleep=lambda _: None,
                )

    def test_environment_headers_json_and_request_defenses(self) -> None:
        _load_environment(None)
        with tempfile.TemporaryDirectory() as temporary:
            environment = Path(temporary) / ".env"
            environment.write_text("SEC_USER_AGENT=App env@example.com\n", encoding="utf-8")
            with patch("get_financial_reports.load_dotenv") as load:
                _load_environment(environment)
            load.assert_called_once_with(environment)

        with patch.dict(os.environ, {"SEC_USER_AGENT": "App env@example.com"}, clear=True):
            self.assertEqual(_sec_headers()["User-Agent"], "App env@example.com")

        list_session = FakeSession({"json": FakeResponse(json_value=[])})
        with self.assertRaisesRegex(ValueError, "JSON object"):
            _get_json(list_session, "json", sleep=lambda _: None)

        error_session = FakeSession({"missing": FakeResponse(status_code=404)})
        with self.assertRaises(requests.HTTPError):
            _request(error_session, "missing", timeout=1, sleep=lambda _: None)

        retry_response = FakeResponse(status_code=500)
        retry_response.raise_for_status = Mock()
        retry_session = FakeSession({"retry": retry_response})
        with self.assertRaisesRegex(RuntimeError, "Unreachable retry state"):
            _request(
                retry_session,
                "retry",
                timeout=1,
                max_attempts=1,
                sleep=lambda _: None,
            )

    def test_partial_resolution_collection_deduplication_and_break(self) -> None:
        companies = {
            "0": {"ticker": "ABC", "title": "ABC Corporation", "cik_str": 1},
            "ignored": "not a company",
        }
        self.assertEqual(_resolve_company(companies, "Corpor")[2], "ABC")

        recent = {
            "form": ["10-K"],
            "accessionNumber": ["0000000001-25-000001"],
            "primaryDocument": ["annual-old.htm"],
            "filingDate": ["2025-02-01"],
            "reportDate": ["2024-12-31"],
        }
        older_one = {
            "form": ["10-K", "10-Q", "8-K"],
            "accessionNumber": [
                "0000000001-25-000001",
                "0000000001-25-000002",
                "0000000001-25-000004",
            ],
            "primaryDocument": [
                "annual-duplicate.htm",
                "quarter.htm",
                "current-report.htm",
            ],
            "filingDate": ["2025-02-01", "2025-05-01", "2025-01-15"],
            "reportDate": ["2024-12-31", "2025-03-31", "2025-01-14"],
        }
        older_two = {
            "form": ["10-K"],
            "accessionNumber": ["0000000001-26-000003"],
            "primaryDocument": ["annual-new.htm"],
            "filingDate": ["2026-02-01"],
            "reportDate": ["2025-12-31"],
        }
        session = FakeSession(
            {
                f"{SEC_SUBMISSIONS_URL}/older-one.json": FakeResponse(json_value=older_one),
                f"{SEC_SUBMISSIONS_URL}/older-two.json": FakeResponse(json_value=older_two),
            }
        )
        submissions = {
            "filings": {
                "recent": recent,
                "files": [
                    "invalid",
                    {},
                    {"name": "older-one.json"},
                    {"name": "older-two.json"},
                    {"name": "should-not-be-fetched.json"},
                ],
            }
        }
        rows = _collect_filing_rows(
            session,
            submissions,
            accepted_forms={"10-K", "10-Q", "8-K"},
            targets={"10-K": 2, "10-Q": 1},
            sleep=lambda _: None,
        )
        self.assertEqual(len(rows), 4)
        self.assertEqual(rows[0]["accession_number"], "0000000001-26-000003")
        self.assertEqual(len(session.calls), 2)

        complete = {
            "filings": {
                "recent": older_two,
                "files": [{"name": "should-not-be-fetched.json"}],
            }
        }
        no_fetch = FakeSession({})
        rows = _collect_filing_rows(
            no_fetch,
            complete,
            accepted_forms={"10-K"},
            targets={"10-K": 1},
            sleep=lambda _: None,
        )
        self.assertEqual(len(rows), 1)
        self.assertEqual(no_fetch.calls, [])

        with self.assertRaisesRegex(ValueError, "filings.recent"):
            _collect_filing_rows(
                FakeSession({}),
                {"filings": []},
                accepted_forms={"10-K"},
                targets={"10-K": 1},
                sleep=lambda _: None,
            )

    def test_selection_skip_and_safe_filename_fallbacks(self) -> None:
        rows = [
            {
                "form": "8-K",
                "accession_number": "1-0",
                "primary_document": "ignored.htm",
                "filing_date": "2026-01-01",
                "report_date": "2025-12-31",
            },
            {
                "form": "10-K",
                "accession_number": "1-1",
                "primary_document": "first",
                "filing_date": "2026-02-01",
                "report_date": "",
            },
            {
                "form": "10-K",
                "accession_number": "1-2",
                "primary_document": "second.htm",
                "filing_date": "2025-02-01",
                "report_date": "2024-12-31",
            },
        ]
        selected, shortages = _select_filings("0000000001", rows, {"10-K": 1})
        self.assertEqual(len(selected["10-K"]), 1)
        self.assertEqual(shortages, {})
        self.assertEqual(_safe_filename(selected["10-K"][0]), "10k_2026-02-01_11.htm")
        self.assertEqual(
            _build_filing_url("0000000001", "1-1", "first"),
            "https://www.sec.gov/Archives/edgar/data/1/11/first",
        )
        self.assertEqual(_default_output_path("  A/B & C  ").name, "a_b_c_reports.v2.json")

    def test_get_reports_rejects_issuer_without_10k(self) -> None:
        ticker_json = {
            "0": {"ticker": "ABC", "title": "ABC Corporation", "cik_str": 1}
        }
        submissions_url = f"{SEC_SUBMISSIONS_URL}/CIK0000000001.json"
        submissions = {
            "filings": {
                "recent": {
                    "form": ["10-Q"],
                    "accessionNumber": ["0000000001-26-000002"],
                    "primaryDocument": ["quarter.htm"],
                    "filingDate": ["2026-05-01"],
                    "reportDate": ["2026-03-31"],
                },
                "files": [],
            }
        }
        session = FakeSession(
            {
                SEC_TICKERS_URL: FakeResponse(json_value=ticker_json),
                submissions_url: FakeResponse(json_value=submissions),
            }
        )
        with self.assertRaisesRegex(ValueError, "No 10-K"):
            get_financial_reports(
                "ABC",
                user_agent="App test@example.com",
                annual_count=1,
                quarterly_count=1,
                session=session,
                sleep=lambda _: None,
            )

    def test_get_reports_downloads_all_artifacts_and_uses_default_session(self) -> None:
        ticker_json = {
            "0": {"ticker": "ABC", "title": "ABC Corporation", "cik_str": 1}
        }
        submissions_url = f"{SEC_SUBMISSIONS_URL}/CIK0000000001.json"
        annual_accession = "0000000001-26-000001"
        quarterly_accession = "0000000001-26-000002"
        annual_url = _build_filing_url("0000000001", annual_accession, "annual.htm")
        quarterly_url = _build_filing_url("0000000001", quarterly_accession, "quarter.htm")
        companyfacts_url = f"{SEC_COMPANYFACTS_URL}/CIK0000000001.json"
        submissions = {
            "filings": {
                "recent": {
                    "form": ["10-K", "10-Q"],
                    "accessionNumber": [annual_accession, quarterly_accession],
                    "primaryDocument": ["annual.htm", "quarter.htm"],
                    "filingDate": ["2026-02-01", "2026-05-01"],
                    "reportDate": ["2025-12-31", "2026-03-31"],
                },
                "files": [],
            }
        }
        html = b"<html><body>inline filing</body></html>"
        responses = {
            SEC_TICKERS_URL: FakeResponse(json_value=ticker_json),
            submissions_url: FakeResponse(json_value=submissions),
            annual_url: FakeResponse(content=html, headers={"Content-Type": "text/html"}),
            quarterly_url: FakeResponse(content=b"<ix:header/>"),
            companyfacts_url: FakeResponse(content=b'{"facts":{}}'),
        }
        default_session = FakeSession(responses)
        with tempfile.TemporaryDirectory() as temporary, patch(
            "get_financial_reports.requests.Session",
            return_value=default_session,
        ):
            manifest = get_financial_reports(
                "ABC",
                user_agent="App test@example.com",
                download_directory=temporary,
                annual_count=1,
                quarterly_count=1,
                include_companyfacts=True,
                sleep=lambda _: None,
            )
            self.assertEqual(len(manifest["filings"]), 2)
            self.assertTrue(Path(manifest["companyfacts"]["local_path"]).is_file())
            self.assertTrue(all(Path(item["local_path"]).is_file() for item in manifest["filings"]))

        no_download_session = FakeSession(
            {
                SEC_TICKERS_URL: FakeResponse(json_value=ticker_json),
                submissions_url: FakeResponse(json_value=submissions),
            }
        )
        manifest = get_financial_reports(
            "ABC",
            user_agent="App test@example.com",
            annual_count=1,
            quarterly_count=1,
            include_companyfacts=True,
            session=no_download_session,
            sleep=lambda _: None,
        )
        self.assertEqual(manifest["companyfacts"], {"url": companyfacts_url})


if __name__ == "__main__":
    unittest.main()
