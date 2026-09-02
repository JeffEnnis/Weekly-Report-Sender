"""
Weekly Market Intelligence Report — automation script (Outlook/Graph edition)
-------------------------------------------------------------------------------
Runs the search described in the Blueprint (Section 4), builds a formatted
.docx report, appends new bids/contacts to leads_tracker.xlsx, and emails
the report through your real Outlook mailbox via Microsoft Graph.

Designed to run inside the provided GitHub Actions workflow
(.github/workflows/weekly-report.yml) every Wednesday at 8:00 AM — see
Setup_Guide.docx for the full walkthrough. It also runs fine locally/manually.

SETUP (one-time):
  1. pip install anthropic python-docx openpyxl msal requests
  2. Set the environment variables listed in the CONFIG section below.
  3. Test manually:  python3 weekly_report.py
"""

import os
import json
import datetime
import base64

import anthropic
import requests
import msal
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import openpyxl

# ----------------------------- CONFIG ---------------------------------
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")          # required
MODEL = "claude-sonnet-4-6"

SEND_EMAIL = True

# Microsoft Graph / Outlook — from your Azure App Registration (see Setup Guide)
AZURE_CLIENT_ID = os.environ.get("AZURE_CLIENT_ID")
AZURE_TENANT_ID = os.environ.get("AZURE_TENANT_ID")
AZURE_CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET")
SENDER_EMAIL = os.environ.get("SENDER_EMAIL")          # the Outlook mailbox sending the report
REPORT_EMAIL_TO = os.environ.get("REPORT_EMAIL_TO")    # where the report should land

TRACKER_PATH = os.environ.get("LEADS_TRACKER_PATH", "leads_tracker.xlsx")
OUTPUT_DIR = os.environ.get("REPORT_OUTPUT_DIR", ".")

# Edit this as your focus shifts — regions, trades, target companies.
SEARCH_BRIEF = """
Search BC Bid, MERX, BC Hydro, TransLink, Metro Vancouver, the BC
Environmental Assessment Office project registry, Journal of Commerce,
Northern Miner, Mining.com, Business in Vancouver, and BC government news
releases for the past 7 days plus any newly posted or upcoming
opportunities. Focus on: heavy civil, general construction, mining/mineral
exploration, and civil/structural engineering, in British Columbia.

Return exactly:
- 3 Future Opportunities (projects likely to tender in 1-6 months, not yet
  formally posted — the best signals are EAO filings, municipal capital
  plans, and mining company investor updates)
- 3 Current Opportunities (open bids/RFPs closing in the next 30-60 days)
- 3 Current News items relevant to bidding or client strategy this week
- 3 Key Connections (named individuals worth contacting this week — role,
  company, why they matter, and a suggested outreach angle)

For each item give: title, one-line summary, source name, source URL,
date, and a suggested next action.

Respond with ONLY valid JSON, no markdown fences, no preamble, in this
exact shape:
{
  "future_opportunities": [{"title": "", "summary": "", "source": "", "url": "", "date": "", "next_action": ""}],
  "current_opportunities": [{"title": "", "summary": "", "source": "", "url": "", "date": "", "next_action": ""}],
  "current_news": [{"title": "", "summary": "", "source": "", "url": "", "date": "", "next_action": ""}],
  "key_connections": [{"title": "", "summary": "", "source": "", "url": "", "date": "", "next_action": "", "name": "", "company": ""}]
}
"""


def fetch_weekly_data() -> dict:
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    messages = [{"role": "user", "content": SEARCH_BRIEF}]
    response = None

    # Web search can take several search rounds to finish. If the model
    # pauses mid-search (stop_reason "pause_turn") or is cut off by the
    # token limit, feed its partial turn back in and let it continue,
    # rather than treating an incomplete response as a final answer.
    for _ in range(6):
        response = client.messages.create(
            model=MODEL,
            max_tokens=8000,
            tools=[{"type": "web_search_20250305", "name": "web_search"}],
            messages=messages,
        )
        if response.stop_reason == "pause_turn":
            messages.append({"role": "assistant", "content": response.content})
            continue
        break

    text_blocks = [b.text for b in response.content if b.type == "text"]
    raw = "\n".join(text_blocks).strip()
    raw = raw.replace("```json", "").replace("```", "").strip()

    if not raw:
        print("No text content in the model's response. Full response for debugging:")
        print(response.model_dump_json(indent=2))
        raise RuntimeError(
            f"Empty response from Claude (stop_reason={response.stop_reason}). "
            "See the debug output above."
        )

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        print("Response was not valid JSON. Raw text received:")
        print(raw)
        raise


def build_docx(data: dict, week_of: str) -> str:
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    title = doc.add_heading("Weekly Market Intelligence Report", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph(f"Week of {week_of}  |  BC Construction / Engineering / Mining / Heavy Civil")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sections = [
        ("future_opportunities", "3 Future Opportunities"),
        ("current_opportunities", "3 Current Opportunities"),
        ("current_news", "3 Current News"),
        ("key_connections", "3 Key Connections"),
    ]

    for key, label in sections:
        doc.add_heading(label, level=1)
        for item in data.get(key, []):
            p = doc.add_paragraph()
            run = p.add_run(item.get("title", "Untitled"))
            run.bold = True
            run.font.size = Pt(12)
            doc.add_paragraph(item.get("summary", ""))
            meta = doc.add_paragraph()
            meta_run = meta.add_run(
                f"Source: {item.get('source','')}  |  Date: {item.get('date','')}  |  "
                f"Next action: {item.get('next_action','')}"
            )
            meta_run.italic = True
            meta_run.font.size = Pt(9)
            meta_run.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
            if item.get("url"):
                doc.add_paragraph(item["url"])
            doc.add_paragraph()

    filename = os.path.join(OUTPUT_DIR, f"Weekly_Report_{week_of}.docx")
    doc.save(filename)
    return filename


def append_to_tracker(data: dict, found_date: str):
    if not os.path.exists(TRACKER_PATH):
        print(f"Tracker not found at {TRACKER_PATH} — skipping append.")
        return
    wb = openpyxl.load_workbook(TRACKER_PATH)

    opp_ws = wb["Opportunities"]
    for stage_key, stage_label in [("future_opportunities", "Future"), ("current_opportunities", "Current")]:
        for item in data.get(stage_key, []):
            opp_ws.append([
                found_date, stage_label, item.get("title", ""), "", "", "", "",
                "", item.get("source", ""), item.get("url", ""), "Not started",
                item.get("next_action", ""), item.get("summary", ""),
            ])

    contacts_ws = wb["Contacts"]
    for item in data.get("key_connections", []):
        contacts_ws.append([
            found_date, item.get("name", item.get("title", "")), "", item.get("company", ""),
            "", item.get("source", ""), item.get("url", ""), "", "", "Not contacted",
            item.get("next_action", ""), item.get("summary", ""),
        ])

    wb.save(TRACKER_PATH)


def get_graph_token() -> str:
    """Client-credentials auth against Azure AD, scoped to Microsoft Graph."""
    authority = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        AZURE_CLIENT_ID, authority=authority, client_credential=AZURE_CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Failed to get Graph token: {result.get('error_description')}")
    return result["access_token"]


def send_email_via_graph(filepath: str, week_of: str):
    if not (AZURE_CLIENT_ID and AZURE_TENANT_ID and AZURE_CLIENT_SECRET and SENDER_EMAIL and REPORT_EMAIL_TO):
        print("Microsoft Graph credentials not fully set — skipping send. File saved at:", filepath)
        return

    token = get_graph_token()
    with open(filepath, "rb") as f:
        content_b64 = base64.b64encode(f.read()).decode("utf-8")

    message = {
        "message": {
            "subject": f"Weekly Market Intelligence Report — {week_of}",
            "body": {"contentType": "Text", "content": "Attached: this week's bid, news, and connections report."},
            "toRecipients": [{"emailAddress": {"address": REPORT_EMAIL_TO}}],
            "attachments": [{
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": os.path.basename(filepath),
                "contentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "contentBytes": content_b64,
            }],
        },
        "saveToSentItems": "true",
    }

    url = f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail"
    resp = requests.post(url, headers={"Authorization": f"Bearer {token}"}, json=message)
    if resp.status_code >= 300:
        raise RuntimeError(f"Graph sendMail failed ({resp.status_code}): {resp.text}")
    print("Email sent via Outlook.")


def main():
    if not ANTHROPIC_API_KEY:
        raise SystemExit("Set the ANTHROPIC_API_KEY environment variable before running.")

    week_of = datetime.date.today().isoformat()
    print("Fetching weekly data...")
    data = fetch_weekly_data()

    print("Building report...")
    filepath = build_docx(data, week_of)

    print("Updating leads tracker...")
    append_to_tracker(data, week_of)

    if SEND_EMAIL:
        print("Emailing report via Outlook...")
        send_email_via_graph(filepath, week_of)

    print(f"Done. Report saved to {filepath}")


if __name__ == "__main__":
    main()
