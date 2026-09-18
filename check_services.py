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

def run_audit():
    input_file = "input_services.csv"
    output_file = "service_audit_report.xlsx"

    if not os.path.exists(input_file):
        print(f"Error: Could not find {input_file}")
        return

    df = pd.read_csv(input_file)
    results = []

    for index, row in df.iterrows():
        # Read fields flexibly
        url = str(row.get("URL", row.get("url", ""))).strip()
        service_name = str(row.get("Service name", row.get("service_name", f"Row {index+1}"))).strip()
        summary = str(row.get("Description", "")).strip()
        self_referral = str(row.get("Self-referrals accepted", "")).strip()
        age = str(row.get("A list of ages eligible", "")).strip()

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
    final_df = pd.concat([df, audit_df], axis=1)
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Audit Report')

    print(f"Audit completed successfully! Saved to {output_file}")

if __name__ == "__main__":
    run_audit()
