import os
import pandas as pd
from firecrawl import FirecrawlApp
from google import genai
from pydantic import BaseModel, Field

# Initialize client connections from secret keys
firecrawl_key = os.environ.get("FIRECRAWL_API_KEY")
gemini_key = os.environ.get("GEMINI_API_KEY")

firecrawl = FirecrawlApp(api_key=firecrawl_key)
ai_client = genai.Client(api_key=gemini_key)

# Define exact JSON output schema
class AuditResult(BaseModel):
    outcome: str = Field(description="Must be 'Accurate' or 'Needs review'")
    issue: str = Field(description="Explanation of discrepancies, or 'None' if accurate")
    verbatim_evidence: str = Field(description="Exact quote from webpage supporting the finding, or 'None'")

def get_field(row, possible_names, default=""):
    """Helper to extract column value regardless of variations in header names."""
    for name in possible_names:
        if name in row and pd.notna(row[name]):
            return str(row[name]).strip()
    return default

def run_audit():
    input_file = "input_services.csv"
    output_file = "service_audit_report.xlsx"

    if not os.path.exists(input_file):
        print(f"Error: Could not find {input_file}")
        return

    df = pd.read_csv(input_file)
    results = []

    for index, row in df.iterrows():
        # Read fields flexibly across common column naming conventions
        url = get_field(row, ["Source website URL", "URL", "url", "Website URL"])
        service_name = get_field(row, ["Service name", "service_name", "Service Name"], f"Row {index+1}")
        summary = get_field(row, ["Summary text", "Description", "summary", "description"])
        self_referral = get_field(row, ["Self-referral", "Self-referrals accepted", "self_referral"])
        age = get_field(row, ["Age", "A list of ages eligible", "age", "ages"])

        source_checked = "No"
        outcome = "Needs review"
        issue = "URL inaccessible or broken"
        evidence = "N/A"

        if url and url.startswith("http"):
            try:
                # Fetch page content via Firecrawl
                scrape_result = firecrawl.scrape_url(url, params={'formats': ['markdown']})
                page_text = scrape_result.get('markdown', '')

                if page_text:
                    source_checked = "Yes"

                    prompt = f"""
                    Perform quality assurance on this youth service directory record.

                    SERVICE DETAILS:
                    - Service Name: {service_name}
                    - Directory Summary: {summary}
                    - Self-Referral Flag: {self_referral} (1 = Yes accepted, Blank/0 = No)
                    - Eligible Ages: {age} (numbers separated by |)

                    WEBPAGE TEXT:
                    {page_text[:8000]}

                    VERIFICATION RULES:
                    1. Self-referral means a young person or parent/carer can contact directly without professional/GP referral.
                    2. Do NOT flag if age or self-referral data is missing on the webpage unless explicitly contradicted by site text.
                    3. Accept plain-English rewording and reasonable simplifications.
                    4. Mark 'Needs review' only if webpage text directly contradicts summary, self-referral, or age criteria.
                    """

                    response = ai_client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=prompt,
                        config={
                            'response_mime_type': 'application/json',
                            'response_schema': AuditResult,
                        }
                    )

                    audit = AuditResult.model_validate_json(response.text)
                    outcome = audit.outcome
                    issue = audit.issue
                    evidence = audit.verbatim_evidence

            except Exception as e:
                issue = f"Check failed: {str(e)}"

        results.append({
            "Source checked successfully": source_checked,
            "Outcome": outcome,
            "Issue": issue,
            "Verbatim Evidence": evidence
        })

    # Combine results and save Excel file
    audit_df = pd.DataFrame(results)
    
    # Drop existing audit columns if present in input CSV to avoid duplicates
    cols_to_drop = [c for c in audit_df.columns if c in df.columns]
    clean_input_df = df.drop(columns=cols_to_drop)
    
    final_df = pd.concat([clean_input_df, audit_df], axis=1)
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Audit Report')

    print(f"Audit completed successfully! Saved to {output_file}")

if __name__ == "__main__":
    run_audit()
