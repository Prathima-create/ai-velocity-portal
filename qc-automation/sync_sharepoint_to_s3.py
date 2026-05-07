"""
SharePoint → S3 Sync Script for QC Automation Dashboard
========================================================
Pulls audit data from Amazon SharePoint (via Edge SSO) and uploads to S3.
Runs on a schedule (cron/systemd timer) on EC2 or locally.

Architecture:
  SharePoint (someone's account) → This script (Edge SSO) → S3 bucket → Streamlit app

Usage:
  python sync_sharepoint_to_s3.py                    # One-time sync
  python sync_sharepoint_to_s3.py --loop 30          # Sync every 30 minutes
  python sync_sharepoint_to_s3.py --local-only       # Download to local data/ only (no S3)

Environment Variables:
  QC_S3_BUCKET          - S3 bucket name (default: fincom-qc-data)
  QC_S3_PREFIX          - S3 prefix/folder (default: current/)
  QC_SHAREPOINT_SITE    - SharePoint site URL
  QC_SHAREPOINT_FOLDER  - SharePoint folder path containing the CSVs
"""

import os
import sys
import time
import json
import shutil
import tempfile
import argparse
from pathlib import Path
from datetime import datetime

# ─── Config ───────────────────────────────────────────────────────────────────
SHAREPOINT_SITE_URL = os.environ.get(
    'QC_SHAREPOINT_SITE',
    'https://amazon.sharepoint.com/sites/FinCom-QC'
)
SHAREPOINT_FOLDER = os.environ.get(
    'QC_SHAREPOINT_FOLDER',
    '/Shared Documents/QC_Data'
)

S3_BUCKET = os.environ.get('QC_S3_BUCKET', 'fincom-qc-data')
S3_PREFIX = os.environ.get('QC_S3_PREFIX', 'current/')

PROJECT_ROOT = Path(__file__).parent
DATA_DIR = PROJECT_ROOT / "data"
DATA_DIR.mkdir(exist_ok=True)
LOG_FILE = DATA_DIR / "sync_log.txt"

# Files we expect to find on SharePoint
EXPECTED_FILES = [
    "Fincom_Process.csv",
    "Fincom_Analyst.csv",
    "Disputes.csv",
    "Rectification.csv",
    "IVOC.csv",
    "Defect Reduction.csv",
]


def log(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


# ============================================================
# SHAREPOINT DOWNLOAD (via Edge SSO)
# ============================================================
def find_edge_profile():
    """Find the default Edge user data directory."""
    local_app = os.environ.get("LOCALAPPDATA", "")
    edge_dir = os.path.join(local_app, "Microsoft", "Edge", "User Data")
    if os.path.isdir(edge_dir):
        return edge_dir
    # Linux/Mac fallback
    home = Path.home()
    for path in [
        home / ".config/microsoft-edge/Default",
        home / "Library/Application Support/Microsoft Edge/Default",
    ]:
        if path.exists():
            return str(path.parent)
    return None


def copy_edge_profile(edge_user_data):
    """Copy essential Edge profile files to temp dir to avoid lock conflicts."""
    temp_dir = os.path.join(tempfile.gettempdir(), "edge_qc_sync_profile")
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir, ignore_errors=True)

    os.makedirs(temp_dir, exist_ok=True)

    local_state = os.path.join(edge_user_data, "Local State")
    if os.path.exists(local_state):
        shutil.copy2(local_state, os.path.join(temp_dir, "Local State"))

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
            except Exception:
                pass

    return temp_dir


def download_from_sharepoint():
    """
    Use Edge + Selenium to download files from SharePoint document library.
    Leverages corporate SSO (no credentials needed — uses cached session).
    """
    try:
        from selenium import webdriver
        from selenium.webdriver.edge.options import Options as EdgeOptions
        from selenium.webdriver.common.by import By
    except ImportError:
        log("ERROR: selenium not installed. Run: pip install selenium")
        return []

    log("=" * 55)
    log("SharePoint → S3 Sync — Starting")
    log("=" * 55)

    edge_user_data = find_edge_profile()
    if not edge_user_data:
        log("ERROR: Edge user data directory not found")
        return []

    log("Copying Edge profile...")
    temp_profile = copy_edge_profile(edge_user_data)

    options = EdgeOptions()
    options.add_argument(f"--user-data-dir={temp_profile}")
    options.add_argument("--profile-directory=Default")
    options.add_argument("--no-first-run")
    options.add_argument("--disable-extensions")
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")

    try:
        driver = webdriver.Edge(options=options)
    except Exception as e:
        log(f"ERROR launching Edge: {e}")
        return []

    driver.implicitly_wait(15)
    driver.set_script_timeout(120)
    downloaded_files = []

    try:
        # Navigate to SharePoint
        log(f"Navigating to SharePoint: {SHAREPOINT_SITE_URL}")
        driver.get(SHAREPOINT_SITE_URL)
        time.sleep(5)

        # Check SSO
        url = driver.current_url.lower()
        if "login" in url or "microsoftonline" in url:
            log("SSO login detected — waiting for auto-redirect...")
            for _ in range(10):
                time.sleep(3)
                if "sharepoint.com" in driver.current_url.lower() and "login" not in driver.current_url.lower():
                    break
            else:
                log("ERROR: SSO did not auto-complete")
                return []

        log("SharePoint loaded successfully")
        time.sleep(3)

        # Download each expected file via REST API
        for filename in EXPECTED_FILES:
            try:
                file_path = f"{SHAREPOINT_FOLDER}/{filename}"
                download_js = """
                var cb = arguments[arguments.length - 1];
                (async function() {
                    try {
                        var url = '%s/_api/web/GetFileByServerRelativeUrl(\'%s\')/$value';
                        var r = await fetch(url, {
                            credentials: 'same-origin',
                            headers: {'Accept': 'application/octet-stream'}
                        });
                        if (!r.ok) { cb(JSON.stringify({error: r.status + ' ' + r.statusText})); return; }
                        var text = await r.text();
                        cb(JSON.stringify({data: text, size: text.length}));
                    } catch(e) { cb(JSON.stringify({error: e.message})); }
                })();
                """ % (SHAREPOINT_SITE_URL, file_path.replace("'", "\\'"))

                result = json.loads(driver.execute_async_script(download_js))

                if 'error' in result:
                    log(f"  ⚠️ {filename}: Not found or error ({result['error']})")
                    continue

                # Save locally
                local_path = DATA_DIR / filename
                with open(local_path, 'w', encoding='utf-8-sig') as f:
                    f.write(result['data'])

                log(f"  ✅ {filename}: {result['size']} bytes downloaded")
                downloaded_files.append(local_path)

            except Exception as e:
                log(f"  ❌ {filename}: Error — {e}")

    except Exception as e:
        log(f"ERROR: {e}")
    finally:
        try:
            driver.quit()
        except Exception:
            pass
        shutil.rmtree(temp_profile, ignore_errors=True)

    log(f"Downloaded {len(downloaded_files)}/{len(EXPECTED_FILES)} files")
    return downloaded_files


# ============================================================
# ALTERNATIVE: Download from SharePoint Excel file (direct URL)
# ============================================================
def download_from_sharepoint_url(file_url):
    """
    Download a single file from a SharePoint sharing URL.
    Works when someone shares a direct link to an Excel/CSV file.
    
    Usage:
      Set QC_SHAREPOINT_FILE_URL environment variable to the sharing URL.
    """
    try:
        from selenium import webdriver
        from selenium.webdriver.edge.options import Options as EdgeOptions
    except ImportError:
        log("ERROR: selenium not installed")
        return None

    edge_user_data = find_edge_profile()
    if not edge_user_data:
        return None

    temp_profile = copy_edge_profile(edge_user_data)

    options = EdgeOptions()
    options.add_argument(f"--user-data-dir={temp_profile}")
    options.add_argument("--profile-directory=Default")
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")

    # Set download directory
    prefs = {"download.default_directory": str(DATA_DIR)}
    options.add_experimental_option("prefs", prefs)

    try:
        driver = webdriver.Edge(options=options)
        driver.get(file_url)
        time.sleep(10)  # Wait for download/redirect

        # If it opens in browser, try to get download link
        download_js = """
        var cb = arguments[arguments.length - 1];
        (async function() {
            try {
                var r = await fetch(window.location.href.replace('?web=1','') + '?download=1', {
                    credentials: 'same-origin'
                });
                var text = await r.text();
                cb(JSON.stringify({data: text, size: text.length}));
            } catch(e) { cb(JSON.stringify({error: e.message})); }
        })();
        """
        result = json.loads(driver.execute_async_script(download_js))
        if 'data' in result:
            local_path = DATA_DIR / "downloaded_file.csv"
            with open(local_path, 'w', encoding='utf-8-sig') as f:
                f.write(result['data'])
            log(f"Downloaded file: {result['size']} bytes")
            return local_path
    except Exception as e:
        log(f"ERROR downloading URL: {e}")
    finally:
        try:
            driver.quit()
        except Exception:
            pass
        shutil.rmtree(temp_profile, ignore_errors=True)

    return None


# ============================================================
# S3 UPLOAD
# ============================================================
def upload_to_s3(files):
    """Upload downloaded files to S3 bucket."""
    try:
        import boto3
    except ImportError:
        log("ERROR: boto3 not installed. Run: pip install boto3")
        return False

    log(f"Uploading {len(files)} files to s3://{S3_BUCKET}/{S3_PREFIX}")

    try:
        s3 = boto3.client('s3')

        for filepath in files:
            key = f"{S3_PREFIX}{filepath.name}"
            s3.upload_file(
                str(filepath),
                S3_BUCKET,
                key,
                ExtraArgs={'ContentType': 'text/csv'}
            )
            log(f"  ✅ Uploaded: s3://{S3_BUCKET}/{key}")

        # Upload sync metadata
        metadata = {
            'last_sync': datetime.now().isoformat(),
            'files_synced': [f.name for f in files],
            'source': 'sharepoint_edge_sso',
        }
        s3.put_object(
            Bucket=S3_BUCKET,
            Key=f"{S3_PREFIX}_sync_metadata.json",
            Body=json.dumps(metadata, indent=2),
            ContentType='application/json'
        )

        log("S3 upload complete!")
        return True

    except Exception as e:
        log(f"ERROR uploading to S3: {e}")
        return False


# ============================================================
# MAIN
# ============================================================
def sync_once(local_only=False):
    """Run a single sync cycle."""
    # Step 1: Download from SharePoint
    files = download_from_sharepoint()

    if not files:
        log("No files downloaded. Checking for existing local files...")
        files = list(DATA_DIR.glob("*.csv"))
        if not files:
            log("ERROR: No data available at all!")
            return False

    # Step 2: Upload to S3 (unless local-only mode)
    if not local_only:
        upload_to_s3(files)
    else:
        log("Local-only mode — skipping S3 upload")

    log("Sync complete!")
    return True


def sync_loop(interval_minutes, local_only=False):
    """Run sync on a loop."""
    log(f"Starting sync loop (every {interval_minutes} min)")
    while True:
        try:
            sync_once(local_only=local_only)
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
    parser = argparse.ArgumentParser(description="SharePoint → S3 Sync for QC Dashboard")
    parser.add_argument("--loop", type=int, default=0,
                        help="Sync interval in minutes (0 = run once)")
    parser.add_argument("--local-only", action="store_true",
                        help="Only download to local data/ folder, don't upload to S3")
    args = parser.parse_args()

    if args.loop > 0:
        sync_loop(args.loop, local_only=args.local_only)
    else:
        ok = sync_once(local_only=args.local_only)
        sys.exit(0 if ok else 1)
