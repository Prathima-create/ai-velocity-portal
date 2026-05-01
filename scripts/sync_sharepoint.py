"""
SharePoint CSV Auto-Sync via Selenium (Microsoft Edge)
Downloads the SharePoint list as CSV using your existing browser SSO session.

Usage:
  python scripts/sync_sharepoint.py              # One-time sync
  python scripts/sync_sharepoint.py --loop 30    # Auto-sync every 30 minutes
  python scripts/sync_sharepoint.py --headless   # Run without browser window
"""

import os
import sys
import time
import shutil
import glob
import argparse
from datetime import datetime

# ─── Config ───────────────────────────────────────────────────────────────────
SHAREPOINT_LIST_URL = (
    "https://amazon.sharepoint.com/sites/AI-Velocity-site"
    "/Lists/AI%20Velocity%20Submission%20Portal/AllItems.aspx"
)
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(PROJECT_ROOT, "data")
TARGET_CSV = os.path.join(DATA_DIR, "submissions.csv")
DOWNLOAD_DIR = os.path.join(PROJECT_ROOT, "data", "_downloads")

# Edge user data dir — reuses your existing SSO session cookies
EDGE_USER_DATA = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Edge", "User Data")


def log(msg):
    """Print with timestamp"""
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


def setup_edge_driver(headless=False):
    """Create Edge WebDriver with SSO session reuse"""
    from selenium import webdriver
    from selenium.webdriver.edge.service import Service
    from selenium.webdriver.edge.options import Options

    # Ensure download dir exists
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    options = Options()

    # Use a copy of user profile to avoid "profile in use" errors
    # We copy cookies from default profile
    options.add_argument(f"--user-data-dir={os.path.join(DOWNLOAD_DIR, '_edge_profile')}")

    if headless:
        options.add_argument("--headless=new")

    # Set download directory
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--no-first-run")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--start-maximized")

    try:
        # Try using webdriver-manager to get the correct EdgeDriver
        from webdriver_manager.microsoft import EdgeChromiumDriverManager
        service = Service(EdgeChromiumDriverManager().install())
        driver = webdriver.Edge(service=service, options=options)
    except Exception:
        # Fallback: let Selenium find it automatically
        driver = webdriver.Edge(options=options)

    driver.implicitly_wait(15)
    return driver


def wait_for_download(timeout=120):
    """Wait for CSV file to appear in download directory"""
    log("Waiting for CSV download to complete...")
    start = time.time()
    while time.time() - start < timeout:
        # Look for CSV files (SharePoint exports as .csv)
        csv_files = glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv"))
        # Filter out partial downloads (.crdownload, .tmp)
        partial = glob.glob(os.path.join(DOWNLOAD_DIR, "*.crdownload")) + \
                  glob.glob(os.path.join(DOWNLOAD_DIR, "*.tmp"))
        
        if csv_files and not partial:
            # Get the most recent CSV
            latest = max(csv_files, key=os.path.getmtime)
            log(f"Download complete: {os.path.basename(latest)}")
            return latest
        
        time.sleep(2)
    
    log("ERROR: Download timed out!")
    return None


def export_sharepoint_list(headless=False):
    """Navigate to SharePoint list and trigger CSV export"""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.action_chains import ActionChains
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    log("Starting Edge browser...")
    driver = setup_edge_driver(headless=headless)

    try:
        # Step 1: Navigate to SharePoint list
        log(f"Navigating to SharePoint list...")
        driver.get(SHAREPOINT_LIST_URL)

        # Step 2: Wait for page to load (may need SSO login)
        log("Waiting for page to load (SSO may be required)...")
        time.sleep(5)

        # Check if we're on a login page
        current_url = driver.current_url.lower()
        if "login" in current_url or "adfs" in current_url or "microsoftonline" in current_url:
            log("SSO login page detected. Please log in manually in the browser window.")
            log("The script will wait up to 120 seconds for you to complete login...")
            
            # Wait for redirect back to SharePoint
            wait_start = time.time()
            while time.time() - wait_start < 120:
                if "sharepoint.com" in driver.current_url and "login" not in driver.current_url.lower():
                    log("Login successful! Continuing...")
                    break
                time.sleep(3)
            else:
                log("ERROR: Login timed out. Please try again.")
                return None

        # Step 3: Wait for the list view to fully load
        log("Waiting for list view to load...")
        time.sleep(8)

        # Step 4: Click "Export" button in the toolbar
        # SharePoint list has an "Export to CSV" option in the toolbar
        log("Looking for Export button...")
        
        try:
            # Method 1: Click the "Export" command bar button
            export_btn = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 
                    "button[name='Export'], button[data-automationid='exportCommand'], "
                    "span[data-automationid='exportCommand']"))
            )
            export_btn.click()
            log("Clicked Export button!")
        except Exception:
            try:
                # Method 2: Try finding by text content
                buttons = driver.find_elements(By.TAG_NAME, "button")
                for btn in buttons:
                    if "export" in btn.text.lower():
                        btn.click()
                        log(f"Clicked button: {btn.text}")
                        break
                else:
                    # Method 3: Use keyboard shortcut or command bar
                    log("Trying command bar approach...")
                    # Click the "..." more commands if needed
                    more_btns = driver.find_elements(By.CSS_SELECTOR, 
                        "button[data-automationid='moreCommand'], button[aria-label='More']")
                    if more_btns:
                        more_btns[0].click()
                        time.sleep(2)
                        # Look for Export in dropdown
                        menu_items = driver.find_elements(By.CSS_SELECTOR, 
                            "button[role='menuitem'], div[role='menuitem']")
                        for item in menu_items:
                            if "export" in item.text.lower():
                                item.click()
                                log(f"Clicked menu item: {item.text}")
                                break
            except Exception as e:
                log(f"Could not find Export button: {e}")
                log("Taking screenshot for debugging...")
                driver.save_screenshot(os.path.join(DOWNLOAD_DIR, "debug_screenshot.png"))
                return None

        # Step 5: Handle any export dialog (CSV option)
        time.sleep(3)
        try:
            # If there's a sub-menu for export type, click CSV
            csv_options = driver.find_elements(By.XPATH, 
                "//*[contains(text(), 'CSV') or contains(text(), 'csv')]")
            if csv_options:
                csv_options[0].click()
                log("Selected CSV export format")
        except Exception:
            pass  # Direct CSV export without submenu

        # Step 6: Wait for download
        csv_file = wait_for_download(timeout=120)
        return csv_file

    except Exception as e:
        log(f"ERROR: {e}")
        try:
            driver.save_screenshot(os.path.join(DOWNLOAD_DIR, "error_screenshot.png"))
            log("Screenshot saved for debugging")
        except Exception:
            pass
        return None

    finally:
        driver.quit()
        log("Browser closed.")


def update_data(csv_path):
    """Copy downloaded CSV to the data directory"""
    if not csv_path or not os.path.exists(csv_path):
        log("ERROR: No CSV file to process")
        return False

    # Backup existing file
    if os.path.exists(TARGET_CSV):
        backup = TARGET_CSV + f".backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        shutil.copy2(TARGET_CSV, backup)
        log(f"Backed up existing CSV to {os.path.basename(backup)}")

    # Copy new file
    shutil.copy2(csv_path, TARGET_CSV)
    
    # Count rows
    with open(TARGET_CSV, 'r', encoding='utf-8-sig', errors='replace') as f:
        row_count = sum(1 for _ in f) - 1  # minus header

    log(f"SUCCESS! Updated submissions.csv with {row_count} rows")
    
    # Clean up downloads
    for f in glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv")):
        try:
            os.remove(f)
        except Exception:
            pass

    return True


def sync_once(headless=False):
    """Perform a single sync"""
    log("=" * 50)
    log("Starting SharePoint CSV sync...")
    log("=" * 50)
    
    csv_path = export_sharepoint_list(headless=headless)
    
    if csv_path:
        success = update_data(csv_path)
        if success:
            log("Sync completed successfully!")
            return True
    
    log("Sync failed. Data file unchanged.")
    return False


def sync_loop(interval_minutes=30, headless=False):
    """Continuously sync at specified interval"""
    log(f"Starting auto-sync loop (every {interval_minutes} minutes)")
    log("Press Ctrl+C to stop\n")
    
    while True:
        try:
            sync_once(headless=headless)
        except KeyboardInterrupt:
            log("Stopped by user.")
            break
        except Exception as e:
            log(f"ERROR in sync loop: {e}")
        
        log(f"\nNext sync in {interval_minutes} minutes...")
        try:
            time.sleep(interval_minutes * 60)
        except KeyboardInterrupt:
            log("Stopped by user.")
            break


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="SharePoint CSV Auto-Sync")
    parser.add_argument("--loop", type=int, default=0,
                       help="Auto-sync interval in minutes (0 = one-time sync)")
    parser.add_argument("--headless", action="store_true",
                       help="Run browser in headless mode (no visible window)")
    args = parser.parse_args()
    
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    
    if args.loop > 0:
        sync_loop(interval_minutes=args.loop, headless=args.headless)
    else:
        success = sync_once(headless=args.headless)
        sys.exit(0 if success else 1)
