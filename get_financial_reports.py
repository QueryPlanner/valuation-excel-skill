import os
import json
import argparse
from dotenv import load_dotenv
from google import genai
from google.genai import types

load_dotenv()

def get_financial_reports_links(company_name: str):
    """
    Uses Gemini API to find downloadable PDF links for a company's financial reports.
    """
    # Initialize the client. Assumes GEMINI_API_KEY environment variable is set.
    try:
        client = genai.Client()
    except Exception as e:
        print(f"Error initializing Google GenAI client. Make sure GEMINI_API_KEY is set: {e}")
        return

    # User specified "gemini 3 flash preview"
    model_name = "gemini-3.5-flash"  # Update this if the exact API model string differs

    prompt = f"""
Can you get me downloadable PDF links to the latest annual report and the last two quarterly reports (Financial results, not press releases or presentation reports) for the company: {company_name}.

Please provide the output as a valid JSON object with the following structure:
{{
  "company_name": "{company_name}",
  "annual_report": {{
    "year": "YYYY-YYYY",
    "pdf_link": "https://..."
  }},
  "quarterly_reports": [
    {{
      "quarter": "Qx YYYY-YYYY",
      "pdf_link": "https://..."
    }},
    {{
      "quarter": "Qx YYYY-YYYY",
      "pdf_link": "https://..."
    }}
  ]
}}
Only return the JSON object and no other text. Ensure the links point directly to PDFs if possible.
"""

    print(f"Querying Gemini API for {company_name} reports...")
    
    try:
        response = client.models.generate_content(
            model=model_name,
            contents=prompt,
            config=types.GenerateContentConfig(
                temperature=0.2,
                tools=[{"google_search": {}}],
            ),
        )
        
        # Parse the JSON response
        try:
            # strip markdown json blocks if present
            response_text = response.text.strip()
            if response_text.startswith("```json"):
                response_text = response_text[7:]
            if response_text.endswith("```"):
                response_text = response_text[:-3]
            response_text = response_text.strip()
            
            result = json.loads(response_text)
            print("\nSuccessfully retrieved links:")
            print(json.dumps(result, indent=2))
            
            # Save to a file
            output_file = f"{company_name.lower().replace(' ', '_')}_reports.json"
            with open(output_file, "w") as f:
                json.dump(result, f, indent=2)
            print(f"\nSaved results to {output_file}")
            
        except json.JSONDecodeError:
            print("Failed to parse the response as JSON. Raw response:")
            print(response.text)
            
    except Exception as e:
        print(f"Error calling Gemini API: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Get company financial report links using Gemini API")
    parser.add_argument("--company", type=str, required=True, help="Target company name")
    args = parser.parse_args()
    
    if not os.environ.get("GEMINI_API_KEY"):
        print("Warning: GEMINI_API_KEY environment variable is not set.")
        
    get_financial_reports_links(args.company)
