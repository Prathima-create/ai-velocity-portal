"""
Resolve SharePoint Person IDs to names using Edge browser cookies.
Creates config/person_id_mapping.json that the backend uses.
"""
import json, os, csv, sys, time

# Get unique Person IDs from CSV
DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data")
CSV_PATH = os.path.join(DATA_DIR, "submissions.csv")
MAPPING_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "config", "person_id_mapping.json")

SITE_URL = "https://amazon.sharepoint.com/sites/AI-Velocity-site"

def get_all_person_ids():
    """Extract all unique numeric Person IDs from the CSV."""
    ids = set()
    if not os.path.exists(CSV_PATH):
        print("No CSV found")
        return ids
    with open(CSV_PATH, encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            for field in ['NameStringId', 'NameId', 'Project Owner/LeadId', 'Project TeamId', 'AuthorId']:
                val = row.get(field, '').strip()
                if val and val.isdigit() and int(val) > 0:
                    ids.add(val)
    return ids

def resolve_ids_via_sharepoint(person_ids):
    """Resolve Person IDs using SharePoint REST API via Edge cookies."""
    try:
        import browser_cookie3
        import requests
    except ImportError:
        print("pip install browser_cookie3 requests")
        return {}
    
    print("Getting Edge cookies for SharePoint...")
    try:
        cj = browser_cookie3.edge(domain_name='.sharepoint.com')
    except Exception as e:
        print(f"Could not get Edge cookies: {e}")
        print("Make sure Edge is CLOSED before running this script!")
        return {}
    
    session = requests.Session()
    session.cookies = cj
    session.headers.update({
        'Accept': 'application/json;odata=verbose',
        'User-Agent': 'Mozilla/5.0'
    })
    
    mapping = {}
    total = len(person_ids)
    
    for i, pid in enumerate(sorted(person_ids, key=int)):
        try:
            url = f"{SITE_URL}/_api/web/siteusers/getbyid({pid})"
            resp = session.get(url, timeout=10)
            if resp.status_code == 200:
                data = resp.json()
                user_data = data.get('d', {})
                title = user_data.get('Title', '')
                email = user_data.get('Email', '')
                login = user_data.get('LoginName', '')
                
                # Extract name from Title or Email
                name = title
                if not name and email:
                    name = email.split('@')[0].replace('.', ' ').title()
                if not name and login:
                    parts = login.split('|')
                    name = parts[-1].split('@')[0].replace('.', ' ').title() if parts else ''
                
                if name:
                    mapping[pid] = name
                    print(f"  [{i+1}/{total}] ID {pid} -> {name}")
                else:
                    print(f"  [{i+1}/{total}] ID {pid} -> (no name found)")
            else:
                print(f"  [{i+1}/{total}] ID {pid} -> HTTP {resp.status_code}")
            
            time.sleep(0.1)  # Be polite
        except Exception as e:
            print(f"  [{i+1}/{total}] ID {pid} -> ERROR: {e}")
    
    return mapping

def main():
    # Load existing mapping if any
    existing = {}
    if os.path.exists(MAPPING_PATH):
        with open(MAPPING_PATH) as f:
            existing = json.load(f)
        print(f"Loaded {len(existing)} existing mappings")
    
    # Get all Person IDs from CSV
    person_ids = get_all_person_ids()
    print(f"Found {len(person_ids)} unique Person IDs in CSV")
    
    # Filter out already-resolved ones
    unresolved = [pid for pid in person_ids if pid not in existing]
    print(f"Need to resolve {len(unresolved)} new IDs")
    
    if unresolved:
        new_mappings = resolve_ids_via_sharepoint(unresolved)
        existing.update(new_mappings)
        print(f"\nResolved {len(new_mappings)} new names")
    
    # Save
    os.makedirs(os.path.dirname(MAPPING_PATH), exist_ok=True)
    with open(MAPPING_PATH, 'w') as f:
        json.dump(existing, f, indent=2, ensure_ascii=False)
    print(f"Saved {len(existing)} total mappings to {MAPPING_PATH}")

if __name__ == '__main__':
    main()
