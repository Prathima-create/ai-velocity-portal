"""
SharePoint CSV Auto-Sync via Selenium + REST API (Mozilla Firefox)
Uses Firefox SSO session to authenticate, then fetches data via SharePoint REST API.

Usage:
  python scripts/sync_sharepoint.py              # One-time sync
  python scripts/sync_sharepoint.py --loop 30    # Auto-sync every 30 minutes
"""

import os
import sys
import time
import csv
import shutil
import argparse
import requests
from datetime import datetime

# ─── Config ───────────────────────────────────────────────────────────────────
SHAREPOINT_SITE_URL = "https://amazon.sharepoint.com/sites/AI-Velocity-site"
SHAREPOINT_LIST_URL = f"{SHAREPOINT_SITE_URL}/Lists/AI%20Velocity%20Submission%20Portal/AllItems.aspx"
# REST API endpoint to get all list items
SP_API_URL = f"{SHAREPOINT_SITE_URL}/_api/web/lists/getbytitle('AI Velocity Submission Portal')/items"

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(PROJECT_ROOT, "data")
TARGET_CSV = os.path.join(DATA_DIR, "submissions.csv")

# Firefox profile — reuses your existing SSO session cookies
FIREFOX_PROFILES_DIR = os.path.join(os.environ.get("APPDATA", ""), "Mozilla", "Firefox", "Profiles")


def log(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


def find_firefox_profile():
    """Find the most recently used Firefox profile"""
    if os.path.isdir(FIREFOX_PROFILES_DIR):
        profiles = [os.path.join(FIREFOX_PROFILES_DIR, d) 
                    for d in os.listdir(FIREFOX_PROFILES_DIR)
                    if os.path.isdir(os.path.join(FIREFOX_PROFILES_DIR, d))]
        if profiles:
            return max(profiles, key=os.path.getmtime)
    return None


def get_cookies_from_firefox():
    """Open Firefox with SSO session, navigate to SharePoint, and grab auth cookies"""
    from selenium import webdriver
    from selenium.webdriver.firefox.service import Service
    from selenium.webdriver.firefox.options import Options

    log("Starting Firefox browser...")
    options = Options()
    
    profile_path = find_firefox_profile()
    if profile_path:
        log(f"Using Firefox profile: {os.path.basename(profile_path)}")
        options.profile = profile_path

    # Set download dir (not really needed but avoids errors)
    options.set_preference("browser.download.folderList", 2)
    options.set_preference("browser.download.dir", os.path.join(os.path.expanduser("~"), "Downloads"))
    options.set_preference("browser.download.useDownloadDir", True)
    options.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/csv,application/csv")

    try:
        from webdriver_manager.firefox import GeckoDriverManager
        service = Service(GeckoDriverManager().install())
        driver = webdriver.Firefox(service=service, options=options)
    except Exception as e:
        log(f"webdriver-manager failed ({e}), trying default...")
        driver = webdriver.Firefox(options=options)

    driver.implicitly_wait(15)

    try:
        # Navigate to SharePoint to trigger SSO
        log("Navigating to SharePoint...")
        driver.get(SHAREPOINT_LIST_URL)
        
        log("Waiting for page to load (SSO auth)...")
        time.sleep(5)
        
        # Check if login required
        current_url = driver.current_url.lower()
        if "login" in current_url or "adfs" in current_url or "microsoftonline" in current_url:
            log("SSO login page detected. Please log in manually...")
            wait_start = time.time()
            while time.time() - wait_start < 120:
                if "sharepoint.com" in driver.current_url and "login" not in driver.current_url.lower():
                    log("Login successful!")
                    break
                time.sleep(3)
            else:
                log("ERROR: Login timed out.")
                return None, None
        
        # Wait for page to fully load
        time.sleep(5)
        log("Page loaded. Extracting cookies...")
        
        # Get all cookies from the browser
        cookies = {}
        for cookie in driver.get_cookies():
            cookies[cookie['name']] = cookie['value']
        
        # Get the request digest for API calls
        digest = None
        try:
            digest_script = """
            return document.getElementById('__REQUESTDIGEST') ? 
                   document.getElementById('__REQUESTDIGEST').value : null;
            """
            digest = driver.execute_script(digest_script)
        except:
            pass
        
        if not digest:
            # Try getting it via the contextinfo API
            try:
                digest_js = """
                var xhr = new XMLHttpRequest();
                xhr.open('POST', '%s/_api/contextinfo', false);
                xhr.setRequestHeader('Accept', 'application/json');
                xhr.send();
                return JSON.parse(xhr.responseText).FormDigestValue;
                """ % SHAREPOINT_SITE_URL
                digest = driver.execute_script(digest_js)
            except:
                pass
        
        log(f"Got {len(cookies)} cookies" + (f" and request digest" if digest else ""))
        return cookies, digest
        
    except Exception as e:
        log(f"ERROR: {e}")
        return None, None
    finally:
        driver.quit()
        log("Browser closed.")


def fetch_list_items_via_api(cookies, digest=None):
    """Fetch all SharePoint list items using REST API with browser cookies"""
    log("Fetching list items via SharePoint REST API...")
    
    headers = {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json",
    }
    if digest:
        headers["X-RequestDigest"] = digest
    
    all_items = []
    url = f"{SP_API_URL}?$top=5000"
    
    session = requests.Session()
    # Set all cookies from the browser
    for name, value in cookies.items():
        session.cookies.set(name, value)
    
    while url:
        log(f"Fetching batch... (total so far: {len(all_items)})")
        try:
            resp = session.get(url, headers=headers, timeout=30)
            if resp.status_code == 200:
                data = resp.json()
                results = data.get("d", {}).get("results", [])
                all_items.extend(results)
                
                # Check for pagination
                next_url = data.get("d", {}).get("__next", None)
                url = next_url if next_url else None
                log(f"Got {len(results)} items in this batch")
            elif resp.status_code == 403:
                log("ERROR: Access denied (403). Cookies may have expired.")
                return None
            else:
                log(f"ERROR: API returned status {resp.status_code}")
                log(f"Response: {resp.text[:500]}")
                return None
        except Exception as e:
            log(f"ERROR fetching from API: {e}")
            return None
    
    log(f"Total items fetched: {len(all_items)}")
    return all_items


def fetch_list_via_browser_js(cookies):
    """Alternative: Fetch list data directly using JavaScript in the browser"""
    from selenium import webdriver
    from selenium.webdriver.firefox.service import Service
    from selenium.webdriver.firefox.options import Options

    log("Fetching list data via browser JavaScript...")
    options = Options()
    
    profile_path = find_firefox_profile()
    if profile_path:
        options.profile = profile_path

    try:
        from webdriver_manager.firefox import GeckoDriverManager
        service = Service(GeckoDriverManager().install())
        driver = webdriver.Firefox(service=service, options=options)
    except Exception:
        driver = webdriver.Firefox(options=options)

    driver.implicitly_wait(15)

    try:
        driver.get(SHAREPOINT_LIST_URL)
        time.sleep(8)
        
        # Use JavaScript to fetch all list items via the REST API  
        fetch_js = """
        async function fetchAll() {
            let items = [];
            let url = '%s?$top=5000';
            while (url) {
                let resp = await fetch(url, {
                    headers: {'Accept': 'application/json;odata=verbose'},
                    credentials: 'same-origin'
                });
                let data = await resp.json();
                items = items.concat(data.d.results);
                url = data.d.__next || null;
            }
            return JSON.stringify(items);
        }
        return await fetchAll();
        """ % SP_API_URL
        
        log("Executing REST API fetch via browser...")
        result = driver.execute_script(fetch_js)
        
        import json
        items = json.loads(result)
        log(f"Got {len(items)} items via browser JS")
        return items
        
    except Exception as e:
        log(f"Browser JS fetch failed: {e}")
        return None
    finally:
        driver.quit()
        log("Browser closed.")


def decode_sp_name(name):
    """Decode SharePoint _x00XX_ encoded field names to readable text"""
    import re
    def replace_hex(match):
        hex_code = match.group(1)
        try:
            return chr(int(hex_code, 16))
        except:
            return match.group(0)
    return re.sub(r'_x([0-9a-fA-F]{4})_', replace_hex, name)


def items_to_csv(items, output_path):
    """Convert SharePoint list items (JSON) to CSV matching the backend's expected format"""
    if not items:
        log("No items to write")
        return False
    
    # Skip metadata / internal fields
    skip_keys = {
        "__metadata", "__deferred", "odata.type", "odata.id", "odata.editLink",
        "FileSystemObjectType", "ServerRedirectedEmbedUri", "ServerRedirectedEmbedUrl",
        "ContentTypeId", "ComplianceAssetId", "OData__UIVersionString",
        "GUID", "Id", "ID", "AuthorId", "EditorId", "Attachments",
        "FieldValuesAsHtml", "FieldValuesAsText", "FieldValuesForEdit",
        "File", "Folder", "ParentList", "Properties", "RoleAssignments",
        "FirstUniqueAncestorSecurableObject", "GetDlpPolicyTip",
        "ContentType", "AttachmentFiles", "LikedByInformation", "Versions",
    }
    
    # Collect all unique keys from items
    all_keys = set()
    for item in items:
        all_keys.update(item.keys())
    
    # Filter out skip keys and deferred/metadata objects
    useful_keys = []
    for key in all_keys:
        if key in skip_keys:
            continue
        # Check if any item has a non-metadata value for this key
        has_real_value = False
        for item in items[:5]:  # Sample first 5
            val = item.get(key)
            if val is not None and not isinstance(val, dict):
                has_real_value = True
                break
        if has_real_value:
            useful_keys.append(key)
    
    # Build column map: SP internal name → clean CSV name
    # Decode _x0020_ etc. to spaces, _x002f_ to /, etc.
    col_map = {}
    for key in useful_keys:
        col_map[key] = decode_sp_name(key)
    
    # Sort columns: put important ones first
    priority_cols = [
        "Title", "Name", "What would you like to do",
        "Process", "Sub Process", "Select your process", "Select your Sub process",
        "Problem Statement", "Current Manual Effort", "Proposed AI Solution",
        "Estimated Volume", "Target Timeline",
        "Project Name", "Project Owner/Lead", "Project Team", "Tech team POC",
        "Challenge addressed ", "AI solution ", "Impact",
        "Expected Impact if Implemented?",
        "Can this solution be replicated by others?",
        "Data available", "Support Required (if any)",
        "How do you plan to execute this idea?",
        "Which AI Win are you interested in replicating?",
        "Briefly describe your current process",
        "Can this be replicated across teams?",
        "If your idea is being implemented by you , what stage it is in?",
        "Approval Status", "Manager approval", "L6/L7 approval", "Tech team approval",
        "Your Manager", "Your Team ", "Your name",
        "Created", "Modified",
    ]
    
    # Sort: priority columns first (in order), then remaining alphabetically
    def sort_key(sp_key):
        clean = col_map[sp_key]
        for i, p in enumerate(priority_cols):
            if clean.strip() == p.strip() or clean.startswith(p.strip()[:20]):
                return (0, i)
        return (1, clean)
    
    sorted_keys = sorted(useful_keys, key=sort_key)
    
    # Write CSV
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    # Backup existing
    if os.path.exists(output_path):
        backup = output_path + f".backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        shutil.copy2(output_path, backup)
        log(f"Backed up existing CSV")
    
    with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        # Header row with decoded names
        writer.writerow([col_map[k] for k in sorted_keys])
        # Data rows
        for item in items:
            row = []
            for key in sorted_keys:
                val = item.get(key, "")
                if val is None:
                    val = ""
                elif isinstance(val, dict):
                    val = str(val)
                row.append(str(val))
            writer.writerow(row)
    
    log(f"CSV written: {len(items)} rows, {len(sorted_keys)} columns")
    return True


def sync_via_browser(headless=False):
    """
    Use a single Firefox session: navigate to SharePoint, authenticate via SSO,
    then fetch all list data via JavaScript REST API call directly in the browser.
    """
    from selenium import webdriver
    from selenium.webdriver.firefox.service import Service
    from selenium.webdriver.firefox.options import Options
    import json

    log("=" * 50)
    log("Starting SharePoint sync (Browser JS method)...")
    log("=" * 50)
    
    options = Options()
    profile_path = find_firefox_profile()
    if profile_path:
        log(f"Using Firefox profile: {os.path.basename(profile_path)}")
        options.profile = profile_path
    if headless:
        options.add_argument("--headless")

    try:
        from webdriver_manager.firefox import GeckoDriverManager
        service = Service(GeckoDriverManager().install())
        driver = webdriver.Firefox(service=service, options=options)
    except Exception as e:
        log(f"webdriver-manager failed ({e})")
        driver = webdriver.Firefox(options=options)

    driver.implicitly_wait(15)

    try:
        # Step 1: Navigate to SharePoint (triggers SSO)
        log("Navigating to SharePoint...")
        driver.get(SHAREPOINT_LIST_URL)
        
        log("Waiting for SSO auth...")
        time.sleep(5)
        
        # Handle login if needed
        current_url = driver.current_url.lower()
        if "login" in current_url or "adfs" in current_url or "microsoftonline" in current_url:
            log("SSO login page detected. Please log in manually...")
            wait_start = time.time()
            while time.time() - wait_start < 120:
                if "sharepoint.com" in driver.current_url and "login" not in driver.current_url.lower():
                    log("Login successful!")
                    break
                time.sleep(3)
            else:
                log("ERROR: Login timed out.")
                return False
        
        # Wait for page to fully load
        time.sleep(8)
        log("SharePoint page loaded.")
        
        # Step 2: First, discover the actual list name by querying lists
        log("Discovering SharePoint list name...")
        discover_js = """
        var callback = arguments[arguments.length - 1];
        (async function() {
            try {
                let resp = await fetch('%s/_api/web/lists?$select=Title,ItemCount&$filter=ItemCount gt 50', {
                    headers: {'Accept': 'application/json;odata=verbose'},
                    credentials: 'same-origin'
                });
                let data = await resp.json();
                callback(JSON.stringify(data.d.results.map(l => ({title: l.Title, count: l.ItemCount}))));
            } catch(e) {
                callback(JSON.stringify({error: e.message}));
            }
        })();
        """ % SHAREPOINT_SITE_URL
        
        driver.set_script_timeout(30)
        lists_result = driver.execute_async_script(discover_js)
        lists_data = json.loads(lists_result)
        log(f"Found lists: {lists_data}")
        
        # Find the right list (the one with >100 items that contains "Velocity" or "Contribution")
        list_title = None
        if isinstance(lists_data, list):
            for lst in lists_data:
                title = lst.get("title", "")
                if "velocity" in title.lower() or "contribution" in title.lower() or "submission" in title.lower():
                    list_title = title
                    log(f"Found matching list: '{title}' ({lst.get('count', '?')} items)")
                    break
            if not list_title and lists_data:
                # Just use the first list with many items
                list_title = lists_data[0].get("title", "")
                log(f"Using first large list: '{list_title}'")
        
        if not list_title:
            log("Could not discover list name. Trying known variations...")
            # Try several possible names
            for name in ["AI Velocity Contribution Form", "AI Velocity Submission Portal", 
                         "AI%20Velocity%20Contribution%20Form", "AI_Velocity_Submission_Portal"]:
                test_js = """
                try {
                    let resp = await fetch("%s/_api/web/lists/getbytitle('%s')/ItemCount", {
                        headers: {'Accept': 'application/json;odata=verbose'},
                        credentials: 'same-origin'
                    });
                    let data = await resp.json();
                    return JSON.stringify(data);
                } catch(e) {
                    return JSON.stringify({error: e.message});
                }
                """ % (SHAREPOINT_SITE_URL, name)
                try:
                    result = driver.execute_script(test_js)
                    parsed = json.loads(result)
                    if "error" not in parsed:
                        list_title = name
                        log(f"Found list: '{name}'")
                        break
                except:
                    continue
        
        if not list_title:
            log("ERROR: Could not find the SharePoint list!")
            return False
        
        # Step 3: Fetch all items via REST API in browser
        log(f"Fetching all items from '{list_title}'...")
        fetch_js = """
        var callback = arguments[arguments.length - 1];
        (async function() {
            try {
                let items = [];
                let url = "%s/_api/web/lists/getbytitle('%s')/items?$top=5000";
                while (url) {
                    let resp = await fetch(url, {
                        headers: {'Accept': 'application/json;odata=verbose'},
                        credentials: 'same-origin'
                    });
                    let data = await resp.json();
                    if (data.d && data.d.results) {
                        items = items.concat(data.d.results);
                        url = data.d.__next || null;
                    } else {
                        break;
                    }
                }
                callback(JSON.stringify(items));
            } catch(e) {
                callback(JSON.stringify({error: e.message}));
            }
        })();
        """ % (SHAREPOINT_SITE_URL, list_title)
        
        driver.set_script_timeout(60)
        result = driver.execute_async_script(fetch_js)
        items = json.loads(result)
        log(f"Fetched {len(items)} items!")
        
        if not items:
            log("No items returned!")
            return False
        
        # Step 4: Save as CSV
        success = items_to_csv(items, TARGET_CSV)
        if success:
            log(f"SUCCESS! Updated {TARGET_CSV} with {len(items)} rows")
            return True
        return False
        
    except Exception as e:
        log(f"ERROR: {e}")
        return False
    finally:
        driver.quit()
        log("Browser closed.")


def sync_once(headless=False):
    """Perform a single sync"""
    return sync_via_browser(headless=headless)


def sync_loop(interval_minutes=30, headless=False):
    """Continuously sync at specified interval"""
    log(f"Starting auto-sync loop (every {interval_minutes} minutes)")
    while True:
        try:
            sync_once(headless=headless)
        except KeyboardInterrupt:
            log("Stopped.")
            break
        except Exception as e:
            log(f"ERROR: {e}")
        log(f"\nNext sync in {interval_minutes} minutes...")
        try:
            time.sleep(interval_minutes * 60)
        except KeyboardInterrupt:
            break


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="SharePoint CSV Auto-Sync")
    parser.add_argument("--loop", type=int, default=0)
    parser.add_argument("--headless", action="store_true")
    args = parser.parse_args()
    
    os.makedirs(DATA_DIR, exist_ok=True)
    
    if args.loop > 0:
        sync_loop(interval_minutes=args.loop, headless=args.headless)
    else:
        success = sync_once(headless=args.headless)
        sys.exit(0 if success else 1)
