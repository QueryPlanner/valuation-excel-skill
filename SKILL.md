---
name: valuation-excel-skill
description: Performs end-to-end valuation using an Excel template. Finds PDF links to financial reports using the webfetch tool, downloads them, manually parses the PDFs to extract and calculate valuation inputs (including precise LTM math), populates a predefined Excel valuation spreadsheet using openpyxl, and uploads the populated spreadsheet to Google Drive as a Google Sheet.
license: MIT
compatibility: opencode
metadata:
  audience: developers
  workflow: valuation-excel
---

# Valuation Excel End-to-End Workflow

<workflow>
## Step 1: Document Retrieval via Gemini
- Use the `get_financial_reports.py` script provided within this skill directory.
- This script accepts a `--company` flag and uses the Gemini API (e.g., `gemini-3-flash-preview`) with Google Search grounding to find downloadable PDF links for the target company's latest Annual Report and the last two Quarterly Reports (financial results).
- The script outputs a structured JSON containing the company name and URLs to the PDFs.
- Read the output JSON to get the URLs, then manually download the PDFs using `curl` or `wget`. If the Gemini URLs are broken (404), use `webfetch` to find the correct ones from the investor relations page.

## Step 2: Agent Extraction & Calculation
- Parse the downloaded PDFs (using `pdftotext` or python scripts) to extract the raw text of the Income Statements, Balance Sheets, and Cash Flow Statements.
- Strictly follow the extraction rules defined in `prompt.txt` within this skill directory.
- Manually perform the math to calculate Last Twelve Months (LTM) values (e.g., LTM Revenue = FY Revenue + Current YTD Revenue - Prior YTD Revenue).
- Be extremely careful with units (e.g., converting billions to millions if required) and currency (ensure you are using the correct currency or converting to USD if requested).
- Construct a structured JSON file (e.g., `<company_name>_valuation_inputs.json`) containing all the extracted metrics matching the format specified in `prompt.txt`.

## Step 3: Populate the Excel Spreadsheet
- Use the `fill_excel.py` script provided within this skill directory.
- This script accepts flags (`--company`, `--inputs`, `--output`, `--price`, `--rf_rate`, `--erp`, `--ticker`) to cleanly inject your generated JSON data into the `template.xlsx` file.
- We strongly recommend passing the `--ticker` flag (e.g. `--ticker NVDA`) so the script can automatically pull the current stock price and risk-free rate using `yfinance`.
- **Cost of Capital Nuances:** 
  - The script intelligently handles mapping both the `Input sheet` and the `Cost of capital worksheet`. 
  - It links the calculated Cost of Capital back to the Input sheet dynamically.
  - **Single vs Multibusiness:** The script detects whether the company operates in multiple industries based on the extracted `revenue_splits -> by_business` array. 
    - If single, it uses `Single Business(Global)` and references the primary industry.
    - If multibusiness, it changes the approach to `Multibusiness(US)`, clears out template dummy data, and populates the specific industries and their revenue breakdowns so the spreadsheet calculates a weighted composite beta and cost of capital.
- The output will be a newly saved Excel file (e.g., `<company_name>_valuation.xlsx`).

## Step 4: Upload to Google Sheets
- Use the `upload_to_sheets.py` script provided within this skill directory.
- This script accepts the flags (`--company`, `--file`).
- It will authenticate using OAuth 2.0 credentials (`client_secret_*.json` and the generated `token.json` session file).
- The script intelligently queries Google Drive first. If a sheet with the name `<company_name> Valuation (Auto-filled)` already exists, it will **update** that sheet in place. If not, it will create a new one.
- Output the `webViewLink` to the user so they can view the fully calculated intrinsic valuation in Google Sheets.

## Step 5: Subagent Story Building
- Do not guide the user to write the story manually.
- Instead, use the `task` tool to spawn a subagent to write a comprehensive narrative (the "story") about the company.
- The story should cover the company's core business model, competitors, growth story, profitability, capital efficiency, competitive advantage (moat), and risk profile.

## Step 6: Critique Valuation via Subagent
- Use the `task` tool to spawn a second subagent to critique the valuation.
- You must explicitly instruct this subagent to fully read the `.opencode/skills/valuation-excel-skill/valuation_theory.txt` file (it should read the entire file).
- The subagent should critique the quantitative inputs, expected growth rates, cost of capital, and the story generated in Step 5, ensuring everything aligns strictly with the intrinsic valuation framework outlined in `valuation_theory.txt`.
</workflow>
