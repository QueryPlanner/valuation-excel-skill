from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import time
from pathlib import Path
from typing import Any, Callable

import requests
from dotenv import load_dotenv

from common import atomic_write_json

SEC_TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
SEC_SUBMISSIONS_URL = "https://data.sec.gov/submissions"
SEC_COMPANYFACTS_URL = "https://data.sec.gov/api/xbrl/companyfacts"
SEC_ARCHIVES_URL = "https://www.sec.gov/Archives/edgar/data"
REQUEST_DELAY_SECONDS = 0.12
RETRYABLE_STATUS_CODES = {429, 500, 502, 503, 504}


def _load_environment(environment_file: str | Path | None) -> None:
    if environment_file:
        path = Path(environment_file)
        if not path.is_file():
            raise FileNotFoundError(f"Environment file does not exist: {path}.")
        load_dotenv(path)


def _sec_headers(user_agent: str | None = None) -> dict[str, str]:
    resolved = (user_agent or os.environ.get("SEC_USER_AGENT", "")).strip()
    if not resolved or "@" not in resolved:
        raise ValueError(
            "SEC_USER_AGENT must identify the application and include a monitored contact email."
        )
    return {"User-Agent": resolved, "Accept-Encoding": "gzip, deflate"}


def _request(
    session: requests.Session,
    url: str,
    *,
    timeout: int,
    max_attempts: int = 4,
    sleep: Callable[[float], None] = time.sleep,
) -> requests.Response:
    last_response: requests.Response | None = None
    for attempt in range(max_attempts):
        response = session.get(url, timeout=timeout)
        last_response = response
        if response.status_code not in RETRYABLE_STATUS_CODES:
            response.raise_for_status()
            sleep(REQUEST_DELAY_SECONDS)
            return response
        if attempt + 1 < max_attempts:
            retry_after = response.headers.get("Retry-After")
            delay = float(retry_after) if retry_after and retry_after.isdigit() else 0.5 * (2**attempt)
            sleep(delay)
    assert last_response is not None
    last_response.raise_for_status()
    raise RuntimeError("Unreachable retry state.")


def _get_json(
    session: requests.Session,
    url: str,
    *,
    sleep: Callable[[float], None] = time.sleep,
) -> dict[str, Any]:
    value = _request(session, url, timeout=30, sleep=sleep).json()
    if not isinstance(value, dict):
        raise ValueError(f"SEC endpoint did not return a JSON object: {url}.")
    return value


def _resolve_company(ticker_data: dict[str, Any], company_query: str) -> tuple[str, str, str]:
    query = company_query.strip()
    query_upper = query.upper()
    companies = [company for company in ticker_data.values() if isinstance(company, dict)]
    ticker_matches = [company for company in companies if str(company.get("ticker", "")).upper() == query_upper]
    if len(ticker_matches) == 1:
        company = ticker_matches[0]
        return str(company["cik_str"]).zfill(10), str(company["title"]), str(company["ticker"])

    title_matches = [company for company in companies if str(company.get("title", "")).upper() == query_upper]
    if len(title_matches) == 1:
        company = title_matches[0]
        return str(company["cik_str"]).zfill(10), str(company["title"]), str(company["ticker"])

    partial_matches = [company for company in companies if query_upper in str(company.get("title", "")).upper()]
    if len(partial_matches) == 1:
        company = partial_matches[0]
        return str(company["cik_str"]).zfill(10), str(company["title"]), str(company["ticker"])
    if partial_matches:
        candidates = ", ".join(
            f"{company.get('ticker')} ({company.get('title')})" for company in partial_matches[:10]
        )
        raise ValueError(f"Company query {company_query!r} is ambiguous. Candidates: {candidates}.")
    raise ValueError(f"Could not map company or ticker {company_query!r} to an SEC CIK.")


def _filing_rows(filings: dict[str, list[Any]], accepted_forms: set[str]) -> list[dict[str, str]]:
    columns = ("form", "accessionNumber", "primaryDocument", "filingDate", "reportDate")
    missing = [column for column in columns if column not in filings]
    if missing:
        raise ValueError(f"SEC submissions data is missing columns: {', '.join(missing)}.")
    lengths = {column: len(filings[column]) for column in columns}
    if len(set(lengths.values())) != 1:
        raise ValueError(f"SEC submissions columns have inconsistent lengths: {lengths}.")

    rows: list[dict[str, str]] = []
    for values in zip(*(filings[column] for column in columns), strict=True):
        form, accession, primary_document, filing_date, report_date = values
        if form not in accepted_forms:
            continue
        rows.append(
            {
                "form": str(form),
                "accession_number": str(accession),
                "primary_document": str(primary_document),
                "filing_date": str(filing_date),
                "report_date": str(report_date),
            }
        )
    return rows


def _collect_filing_rows(
    session: requests.Session,
    submissions: dict[str, Any],
    *,
    accepted_forms: set[str],
    targets: dict[str, int],
    sleep: Callable[[float], None] = time.sleep,
) -> list[dict[str, str]]:
    filings = submissions.get("filings")
    if not isinstance(filings, dict) or not isinstance(filings.get("recent"), dict):
        raise ValueError("SEC submissions response does not contain filings.recent.")
    rows = _filing_rows(filings["recent"], accepted_forms)
    accessions_by_form = {
        form: {
            row["accession_number"]
            for row in rows
            if row["form"] == form
        }
        for form in targets
    }
    for older_file in filings.get("files", []):
        if all(
            len(accessions_by_form[form]) >= target
            for form, target in targets.items()
        ):
            break
        if not isinstance(older_file, dict) or not older_file.get("name"):
            continue
        older = _get_json(session, f"{SEC_SUBMISSIONS_URL}/{older_file['name']}", sleep=sleep)
        older_rows = _filing_rows(older, accepted_forms)
        rows.extend(older_rows)
        for row in older_rows:
            if row["form"] in accessions_by_form:
                accessions_by_form[row["form"]].add(row["accession_number"])
    unique = {row["accession_number"]: row for row in rows}
    return sorted(
        unique.values(),
        key=lambda row: (row["filing_date"], row["accession_number"]),
        reverse=True,
    )


def _build_filing_url(cik: str, accession_number: str, primary_document: str) -> str:
    cik_unpadded = str(int(cik))
    accession_without_dashes = accession_number.replace("-", "")
    return f"{SEC_ARCHIVES_URL}/{cik_unpadded}/{accession_without_dashes}/{primary_document}"


def _select_filings(
    cik: str,
    rows: list[dict[str, str]],
    targets: dict[str, int],
) -> tuple[dict[str, list[dict[str, str]]], dict[str, int]]:
    selected = {form: [] for form in targets}
    for row in rows:
        form = row["form"]
        if form not in selected or len(selected[form]) >= targets[form]:
            continue
        filing = dict(row)
        filing["html_url"] = _build_filing_url(cik, row["accession_number"], row["primary_document"])
        selected[form].append(filing)
    shortages = {
        form: targets[form] - len(selected[form])
        for form in targets
        if len(selected[form]) < targets[form]
    }
    return selected, shortages


def _safe_filename(filing: dict[str, str]) -> str:
    form = filing["form"].lower().replace("-", "")
    report_date = filing["report_date"] or filing["filing_date"]
    accession = filing["accession_number"].replace("-", "")
    extension = Path(filing["primary_document"]).suffix or ".htm"
    return f"{form}_{report_date}_{accession}{extension}"


def _download_filing(
    session: requests.Session,
    filing: dict[str, str],
    download_directory: Path,
    *,
    sleep: Callable[[float], None] = time.sleep,
) -> dict[str, str]:
    response = _request(session, filing["html_url"], timeout=60, sleep=sleep)
    content_type = response.headers.get("Content-Type", "")
    start = response.content[:1000].lower()
    appears_html = b"<html" in start or b"<!doctype html" in start or b"<ix:" in start
    accepted_type = "html" in content_type or "xml" in content_type
    if not appears_html and not accepted_type:
        raise ValueError(f"Filing {filing['accession_number']} did not look like Inline XBRL or HTML.")
    download_directory.mkdir(parents=True, exist_ok=True)
    output_path = download_directory / _safe_filename(filing)
    output_path.write_bytes(response.content)
    result = dict(filing)
    result.update(
        {
            "local_path": str(output_path.resolve()),
            "sha256": hashlib.sha256(response.content).hexdigest(),
            "content_type": content_type,
        }
    )
    return result


def _download_companyfacts(
    session: requests.Session,
    cik: str,
    destination: Path,
    *,
    sleep: Callable[[float], None] = time.sleep,
) -> dict[str, str]:
    url = f"{SEC_COMPANYFACTS_URL}/CIK{cik}.json"
    response = _request(session, url, timeout=60, sleep=sleep)
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_bytes(response.content)
    return {
        "url": url,
        "local_path": str(destination.resolve()),
        "sha256": hashlib.sha256(response.content).hexdigest(),
    }


def get_financial_reports(
    company_query: str,
    *,
    user_agent: str | None = None,
    download_directory: str | Path | None = None,
    annual_count: int = 10,
    quarterly_count: int = 2,
    include_companyfacts: bool = True,
    session: requests.Session | None = None,
    sleep: Callable[[float], None] = time.sleep,
) -> dict[str, Any]:
    if annual_count < 1 or quarterly_count < 0:
        raise ValueError("annual_count must be at least 1 and quarterly_count cannot be negative.")
    active_session = session or requests.Session()
    active_session.headers.update(_sec_headers(user_agent))
    ticker_data = _get_json(active_session, SEC_TICKERS_URL, sleep=sleep)
    cik, company_name, ticker = _resolve_company(ticker_data, company_query)
    submissions_url = f"{SEC_SUBMISSIONS_URL}/CIK{cik}.json"
    submissions = _get_json(active_session, submissions_url, sleep=sleep)
    targets = {"10-K": annual_count, "10-Q": quarterly_count}
    rows = _collect_filing_rows(
        active_session,
        submissions,
        accepted_forms=set(targets),
        targets=targets,
        sleep=sleep,
    )
    selected, shortages = _select_filings(cik, rows, targets)
    if not selected["10-K"]:
        raise ValueError("No 10-K filing was found. This v2 workflow currently requires a US 10-K issuer.")

    destination = Path(download_directory) if download_directory else None
    if destination:
        selected = {
            form: [
                _download_filing(active_session, filing, destination, sleep=sleep)
                for filing in filings
            ]
            for form, filings in selected.items()
        }
    flat_filings = [filing for form in ("10-K", "10-Q") for filing in selected[form]]
    companyfacts = None
    if include_companyfacts:
        companyfacts_url = f"{SEC_COMPANYFACTS_URL}/CIK{cik}.json"
        if destination:
            companyfacts = _download_companyfacts(
                active_session,
                cik,
                destination / "companyfacts.json",
                sleep=sleep,
            )
        else:
            companyfacts = {"url": companyfacts_url}

    return {
        "manifest_version": "2.0",
        "company_name": company_name,
        "ticker": ticker,
        "cik": cik,
        "source": "SEC EDGAR",
        "submissions_url": submissions_url,
        "requested_counts": targets,
        "shortages": shortages,
        "filings": flat_filings,
        "latest_10ks": selected["10-K"],
        "latest_10qs": selected["10-Q"],
        "companyfacts": companyfacts,
    }


def _default_output_path(company_query: str) -> Path:
    safe_name = re.sub(r"[^a-z0-9]+", "_", company_query.casefold()).strip("_")
    return Path(f"{safe_name}_reports.v2.json")


def main() -> None:
    parser = argparse.ArgumentParser(description="Download a versioned SEC filing inventory")
    parser.add_argument("--company", required=True)
    parser.add_argument("--download-dir")
    parser.add_argument("--output")
    parser.add_argument("--annual-count", type=int, default=10)
    parser.add_argument("--quarterly-count", type=int, default=2)
    parser.add_argument("--without-companyfacts", action="store_true")
    parser.add_argument("--sec-user-agent")
    parser.add_argument("--env-file")
    args = parser.parse_args()
    _load_environment(args.env_file)
    manifest = get_financial_reports(
        args.company,
        user_agent=args.sec_user_agent,
        download_directory=args.download_dir,
        annual_count=args.annual_count,
        quarterly_count=args.quarterly_count,
        include_companyfacts=not args.without_companyfacts,
    )
    output_path = Path(args.output) if args.output else _default_output_path(args.company)
    atomic_write_json(output_path, manifest)
    print(json.dumps(manifest, indent=2))


if __name__ == "__main__":  # pragma: no cover - exercised through main()
    main()
