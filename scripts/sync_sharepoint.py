"""
SharePoint Auto-Sync via Edge Browser (SSO)
Uses a COPY of your Edge profile to avoid locking issues.
Fetches list data via SharePoint REST API executed inside the browser.

Usage:
  python scripts/sync_sharepoint.py              # One-time sync
  python scripts/sync_sharepoint.py --loop 60    # Auto-sync every 60 minutes
"""

import os
import sys
import time
import csv
import json
import shutil
import tempfile
import argparse
import re
from datetime import datetime

# ─── Config ───────────────────────────────────────────────────────────────────
SHAREPOINT_SITE_URL = "https://amazon.sharepoint.com/sites/AI-Velocity-site"
SHAREPOINT_LIST_PAGE = f"{SHAREPOINT_SITE_URL}/Lists/AI%20Velocity%20Submission%20Portal/AllItems.aspx"

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(PROJECT_ROOT, "data")
TARGET_CSV = os.path.join(DATA_DIR, "submissions.csv")
LOG_FILE = os.path.join(DATA_DIR, "sync_log.txt")


def log(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line, flush=True)
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def find_edge_profile():
    """Find the default Edge user data directory"""
    local_app = os.environ.get("LOCALAPPDATA", "")
    edge_dir = os.path.join(local_app, "Microsoft", "Edge", "User Data")
    if os.path.isdir(edge_dir):
        return edge_dir
    return None


def copy_edge_profile(edge_user_data):
    """
    Copy essential Edge profile files to a temp directory.
    This avoids the 'profile is locked' error when Edge is already open.
    Only copies cookies/session data, not the entire profile.
    """
    temp_dir = os.path.join(tempfile.gettempdir(), "edge_sync_profile")

    # Clean up previous temp profile
    if os.path.exists(temp_dir):
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception:
            pass

    os.makedirs(temp_dir, exist_ok=True)

    # Copy Local State (needed for cookie decryption)
    local_state = os.path.join(edge_user_data, "Local State")
    if os.path.exists(local_state):
        shutil.copy2(local_state, os.path.join(temp_dir, "Local State"))

    # Copy the Default profile's essential files
    src_profile = os.path.join(edge_user_data, "Default")
    dst_profile = os.path.join(temp_dir, "Default")
    os.makedirs(dst_profile, exist_ok=True)

    essential_files = [
        "Cookies", "Cookies-journal",
        "Login Data", "Login Data-journal",
        "Web Data", "Web Data-journal",
        "Preferences", "Secure Preferences",
    ]

    for fname in essential_files:
        src = os.path.join(src_profile, fname)
        if os.path.exists(src):
            try:
                shutil.copy2(src, os.path.join(dst_profile, fname))
            except Exception as e:
                log(f"  Could not copy {fname}: {e}")

    log(f"Edge profile copied to temp dir")
    return temp_dir


def decode_sp_field_name(name):
    """Decode SharePoint _x00XX_ encoded field names"""
    def replace_hex(match):
        try:
            return chr(int(match.group(1), 16))
        except Exception:
            return match.group(0)
    return re.sub(r'_x([0-9a-fA-F]{4})_', replace_hex, name)


def sync_via_edge():
    """
    Open Edge with copied profile, navigate to SharePoint (SSO auto-login),
    then fetch all list items via REST API in the browser context.
    """
    from selenium import webdriver
    from selenium.webdriver.edge.service import Service as EdgeService
    from selenium.webdriver.edge.options import Options as EdgeOptions

    log("=" * 55)
    log("SharePoint Sync — Starting (Edge + SSO)")
    log("=" * 55)

    # ── Prepare Edge profile copy ──
    edge_user_data = find_edge_profile()
    if not edge_user_data:
        log("ERROR: Edge user data directory not found")
        return False

    log("Copying Edge profile (to avoid lock conflicts)...")
    temp_profile = copy_edge_profile(edge_user_data)

    # ── Launch Edge ──
    options = EdgeOptions()
    options.add_argument(f"--user-data-dir={temp_profile}")
    options.add_argument("--profile-directory=Default")
    options.add_argument("--no-first-run")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-popup-blocking")
    # Run headless so no window pops up
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")

    try:
        driver = webdriver.Edge(options=options)
    except Exception as e:
        log(f"ERROR launching Edge: {e}")
        log("Make sure Microsoft Edge and Edge WebDriver are installed.")
        log("Download WebDriver from: https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/")
        return False

    driver.implicitly_wait(15)
    driver.set_script_timeout(60)

    try:
        # ── Step 1: Navigate to SharePoint ──
        log("Navigating to SharePoint...")
        driver.get(SHAREPOINT_LIST_PAGE)
        time.sleep(5)

        # Check if we landed on a login page
        url = driver.current_url.lower()
        if "login" in url or "adfs" in url or "microsoftonline" in url:
            log("SSO login page detected — waiting for auto-redirect...")
            # Give SSO up to 30s to auto-complete with cached creds
            for _ in range(10):
                time.sleep(3)
                url = driver.current_url.lower()
                if "sharepoint.com" in url and "login" not in url:
                    break
            else:
                log("ERROR: SSO did not auto-complete. You may need to log in manually once.")
                log("Open Edge, go to SharePoint, log in, then try again.")
                return False

        log("SharePoint loaded successfully")
        time.sleep(3)

        # ── Step 2: Discover the list ──
        log("Discovering SharePoint list...")
        discover_js = """
        var cb = arguments[arguments.length - 1];
        (async function() {
            try {
                var r = await fetch('%s/_api/web/lists?$select=Title,ItemCount&$filter=ItemCount gt 10', {
                    headers: {'Accept': 'application/json;odata=verbose'},
                    credentials: 'same-origin'
                });
                var d = await r.json();
                cb(JSON.stringify(d.d.results.map(function(l){return {title:l.Title, count:l.ItemCount};})));
            } catch(e) { cb(JSON.stringify({error: e.message})); }
        })();
        """ % SHAREPOINT_SITE_URL

        lists_raw = driver.execute_async_script(discover_js)
        lists_data = json.loads(lists_raw)

        list_title = None
        if isinstance(lists_data, list):
            for lst in lists_data:
                t = lst.get("title", "")
                if any(kw in t.lower() for kw in ["velocity", "submission", "contribution"]):
                    list_title = t
                    log(f"Found list: '{t}' ({lst.get('count', '?')} items)")
                    break
            if not list_title and lists_data:
                # Pick the largest list as fallback
                lists_data.sort(key=lambda x: x.get("count", 0), reverse=True)
                list_title = lists_data[0]["title"]
                log(f"Using largest list: '{list_title}' ({lists_data[0].get('count')} items)")

        if not list_title:
            # Try known names directly
            for name in ["AI Velocity Submission Portal", "AI Velocity Contribution Form"]:
                test_js = """
                var cb = arguments[arguments.length - 1];
                (async function() {
                    try {
                        var r = await fetch("%s/_api/web/lists/getbytitle('%s')/ItemCount", {
                            headers: {'Accept': 'application/json;odata=verbose'},
                            credentials: 'same-origin'
                        });
                        var d = await r.json();
                        cb(JSON.stringify(d));
                    } catch(e) { cb(JSON.stringify({error: e.message})); }
                })();
                """ % (SHAREPOINT_SITE_URL, name)
                try:
                    result = json.loads(driver.execute_async_script(test_js))
                    if "error" not in result:
                        list_title = name
                        log(f"Found list by name: '{name}'")
                        break
                except Exception:
                    continue

        if not list_title:
            log("ERROR: Could not find the SharePoint list")
            return False

        # ── Step 3: Fetch all items ──
        log(f"Fetching items from '{list_title}'...")
        # Escape single quotes in list title for JS
        safe_title = list_title.replace("'", "\\'")
        fetch_js = """
        var cb = arguments[arguments.length - 1];
        (async function() {
            try {
                var items = [];
                var url = "%s/_api/web/lists/getbytitle('%s')/items?$top=5000";
                while (url) {
                    var r = await fetch(url, {
                        headers: {'Accept': 'application/json;odata=verbose'},
                        credentials: 'same-origin'
                    });
                    var d = await r.json();
                    if (d.d && d.d.results) {
                        items = items.concat(d.d.results);
                        url = d.d.__next || null;
                    } else { break; }
                }
                cb(JSON.stringify(items));
            } catch(e) { cb(JSON.stringify({error: e.message})); }
        })();
        """ % (SHAREPOINT_SITE_URL, safe_title)

        raw = driver.execute_async_script(fetch_js)
        items = json.loads(raw)

        if isinstance(items, dict) and "error" in items:
            log(f"ERROR fetching items: {items['error']}")
            return False

        log(f"Fetched {len(items)} items")

        if not items:
            log("WARNING: No items returned")
            return False

        # ── Step 4: Convert to CSV ──
        return save_items_as_csv(items)

    except Exception as e:
        log(f"ERROR: {e}")
        return False
    finally:
        try:
            driver.quit()
        except Exception:
            pass
        log("Browser closed")
        # Clean up temp profile
        try:
            shutil.rmtree(temp_profile, ignore_errors=True)
        except Exception:
            pass


def save_items_as_csv(items):
    """Convert SharePoint JSON items to CSV matching the backend format"""
    # Collect all field names, skip metadata
    skip = {"__metadata", "odata.type", "odata.id", "odata.editLink", "odata.etag",
            "FileSystemObjectType", "ServerRedirectedEmbedUri", "ServerRedirectedEmbedUrl",
            "ContentTypeId", "ComplianceAssetId", "OData__UIVersionString",
            "GUID", "Attachments", "AuthorId", "EditorId"}

    all_keys = set()
    for item in items:
        all_keys.update(k for k in item.keys() if k not in skip and not isinstance(item.get(k), dict))

    # Decode field names
    col_map = {k: decode_sp_field_name(k) for k in all_keys}

    # Sort columns sensibly
    priority = [
        "Created By", "What would you like to do", "Name", "Process", "Sub Process",
        "Problem Statement", "Current Manual Effort", "Proposed AI Solution",
        "Project Name", "Project Owner", "Impact", "Created", "Modified",
    ]

    def sort_key(k):
        clean = col_map[k]
        for i, p in enumerate(priority):
            if p.lower() in clean.lower():
                return (0, i)
        return (1, clean)

    sorted_keys = sorted(all_keys, key=sort_key)

    # Backup existing CSV
    os.makedirs(DATA_DIR, exist_ok=True)
    if os.path.exists(TARGET_CSV):
        backup = TARGET_CSV + f".backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        shutil.copy2(TARGET_CSV, backup)
        log("Backed up existing CSV")

    # Write
    try:
        with open(TARGET_CSV, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow([col_map[k] for k in sorted_keys])
            for item in items:
                row = []
                for k in sorted_keys:
                    val = item.get(k, "")
                    if val is None:
                        val = ""
                    elif isinstance(val, dict):
                        val = ""
                    elif isinstance(val, list):
                        val = ", ".join(str(v) for v in val)
                    # Sanitize Unicode for safe CSV writing
                    val = str(val).replace('\u2192', '->').replace('\u2013', '-').replace('\u2014', '-').replace('\u2018', "'").replace('\u2019', "'").replace('\u201c', '"').replace('\u201d', '"')
                    row.append(val)
                writer.writerow(row)
    except Exception as e:
        log(f"ERROR writing CSV: {e}")
        return False

    log(f"CSV saved: {len(items)} rows, {len(sorted_keys)} columns -> {TARGET_CSV}")

    # Validate
    size = os.path.getsize(TARGET_CSV)
    if size < 500:
        log(f"WARNING: CSV is only {size} bytes — may be incomplete")
        return False

    log(f"File size: {size:,} bytes")
    return True


def sync_once():
    """Run a single sync"""
    return sync_via_edge()


def sync_loop(interval_minutes):
    """Run sync on a loop"""
    log(f"Starting sync loop (every {interval_minutes} min)")
    while True:
        try:
            sync_once()
        except KeyboardInterrupt:
            log("Stopped by user")
            break
        except Exception as e:
            log(f"Sync error: {e}")
        log(f"Next sync in {interval_minutes} minutes...")
        try:
            time.sleep(interval_minutes * 60)
        except KeyboardInterrupt:
            break


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="SharePoint Auto-Sync (Edge)")
    parser.add_argument("--loop", type=int, default=0, help="Sync interval in minutes (0 = once)")
    args = parser.parse_args()

    if args.loop > 0:
        sync_loop(args.loop)
    else:
        ok = sync_once()
        sys.exit(0 if ok else 1)
