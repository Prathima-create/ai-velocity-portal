"""
Cloud Sync: Fetch SharePoint list data via REST API and save as CSV.
Runs in GitHub Actions (no browser needed).

Setup: Add these GitHub Secrets:
  SHAREPOINT_SITE_URL    = https://amazon.sharepoint.com/sites/AI-Velocity-site
  SHAREPOINT_LIST_NAME   = AI Velocity Submission Portal
  SHAREPOINT_ACCESS_TOKEN = Bearer token from Azure AD app registration
"""

import os
import csv
import json
import sys
import requests
from datetime import datetime


def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def get_sharepoint_items(site_url, list_name, access_token):
    """Fetch all items from a SharePoint list via REST API"""
    # SharePoint REST API endpoint for list items
    api_url = f"{site_url}/_api/web/lists/getbytitle('{list_name}')/items"
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json;odata=nometadata",
        "Content-Type": "application/json"
    }
    
    all_items = []
    next_url = api_url + "?$top=500"  # Get 500 items per page
    
    while next_url:
        log(f"Fetching: {next_url[:80]}...")
        resp = requests.get(next_url, headers=headers, timeout=30)
        
        if resp.status_code == 401:
            log("ERROR: Authentication failed (401). Check your SHAREPOINT_ACCESS_TOKEN secret.")
            log("You need to register an Azure AD app and get a Bearer token.")
            log("See docs/SHAREPOINT_API_GUIDE.md for instructions.")
            sys.exit(1)
        
        if resp.status_code == 403:
            log("ERROR: Access denied (403). Your app may not have permissions to this list.")
            sys.exit(1)
        
        if resp.status_code != 200:
            log(f"ERROR: HTTP {resp.status_code} - {resp.text[:200]}")
            sys.exit(1)
        
        data = resp.json()
        items = data.get("value", [])
        all_items.extend(items)
        log(f"  Got {len(items)} items (total: {len(all_items)})")
        
        # Check for pagination
        next_url = data.get("odata.nextLink") or data.get("@odata.nextLink")
    
    return all_items


# SharePoint internal field names → CSV column names mapping
FIELD_MAP = {
    "Title": "Name",
    "field_1": "What would you like to do",
    "field_2": "Process",
    "field_3": "Sub Process",
    "field_4": "Problem Statement",
    "field_5": "Current Manual Effort",
    "field_6": "Proposed AI Solution",
    "field_7": "Estimated Volume",
    "field_8": "Expected Impact if Implemented?",
    "field_9": "How do you plan to execute this idea?",
    "field_10": "Data available",
    "field_11": "Support Required (if any)",
    "field_12": "Target Timeline",
    "field_13": "Can this be replicated across teams?",
    "field_14": "Project Name",
    "field_15": "Project Owner/Lead",
    "field_16": "Project Team",
    "field_17": "Tech team POC",
    "field_18": "Challenge addressed ",
    "field_19": "AI solution ",
    "field_20": "Impact",
    "field_21": "Can this solution be replicated by others?",
    "field_22": "Which AI Win are you interested in replicating?",
    "field_23": "Briefly describe your current process",
    "field_24": "Your Manager",
    "field_25": "Your Team ",
    "field_26": "Approval Status",
    "field_27": "Manager approval",
    "field_28": "L6/L7 approval",
    "field_29": "Tech team approval",
    "field_30": "If your idea is being implemented by you , what stage it is in?",
    "Created": "Created",
    "Modified": "Modified",
    "Author": "Created By",
    "Editor": "Modified By",
}


def items_to_csv(items, output_path):
    """Convert SharePoint list items to CSV"""
    if not items:
        log("WARNING: No items to write")
        return 0
    
    # Get all unique field names from items
    all_fields = set()
    for item in items:
        all_fields.update(item.keys())
    
    log(f"Available fields: {sorted(all_fields)[:20]}...")
    
    # Determine CSV columns — use FIELD_MAP if fields match, otherwise use raw field names
    # First try to map known fields
    csv_columns = []
    field_to_csv = {}
    
    for sp_field, csv_name in FIELD_MAP.items():
        if sp_field in all_fields:
            csv_columns.append(csv_name)
            field_to_csv[sp_field] = csv_name
    
    # If no mapped fields found, use all fields directly (common with different SP configurations)
    if len(csv_columns) < 5:
        log("INFO: Using raw SharePoint field names as CSV columns")
        csv_columns = sorted(all_fields - {"__metadata", "odata.type", "odata.id", "odata.etag", "odata.editLink"})
        field_to_csv = {f: f for f in csv_columns}
    
    # Write CSV
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=csv_columns)
        writer.writeheader()
        
        for item in items:
            row = {}
            for sp_field, csv_name in field_to_csv.items():
                value = item.get(sp_field, "")
                # Handle complex fields (lookups, people, etc.)
                if isinstance(value, dict):
                    value = value.get("Title", value.get("LookupValue", str(value)))
                elif isinstance(value, list):
                    value = "; ".join(str(v) for v in value)
                row[csv_name] = str(value) if value is not None else ""
            writer.writerow(row)
    
    return len(items)


def main():
    log("=" * 60)
    log("SharePoint Cloud Sync - Starting")
    log("=" * 60)
    
    site_url = os.environ.get("SHAREPOINT_SITE_URL", "").strip().rstrip("/")
    list_name = os.environ.get("SHAREPOINT_LIST_NAME", "AI Velocity Submission Portal").strip()
    access_token = os.environ.get("SHAREPOINT_ACCESS_TOKEN", "").strip()
    
    if not site_url:
        log("ERROR: SHAREPOINT_SITE_URL not set. Add it as a GitHub Secret.")
        log("Example: https://amazon.sharepoint.com/sites/AI-Velocity-site")
        sys.exit(1)
    
    if not access_token:
        log("ERROR: SHAREPOINT_ACCESS_TOKEN not set. Add it as a GitHub Secret.")
        log("")
        log("To get a token, you need to:")
        log("1. Register an app in Azure AD (ask IT/admin)")
        log("2. Grant it Sites.Read.All permission")
        log("3. Get a client_id + client_secret")
        log("4. Use them to get a Bearer token")
        log("")
        log("OR use the manual CSV upload approach instead:")
        log("  - Download CSV from SharePoint manually")
        log("  - Upload via the Upload CSV button on the dashboard")
        sys.exit(1)
    
    log(f"Site URL: {site_url}")
    log(f"List: {list_name}")
    log(f"Token: {access_token[:20]}...")
    
    output_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "submissions.csv")
    
    # Fetch items
    items = get_sharepoint_items(site_url, list_name, access_token)
    
    if not items:
        log("WARNING: No items returned from SharePoint. CSV not updated.")
        sys.exit(0)
    
    # Convert to CSV
    count = items_to_csv(items, output_path)
    
    log(f"SUCCESS: Wrote {count} rows to {output_path}")
    log(f"File size: {os.path.getsize(output_path):,} bytes")
    log("=" * 60)


if __name__ == "__main__":
    main()
