"""
Auto Sync & Push: Downloads CSV from SharePoint via Selenium, then pushes to GitHub.
Run this via Windows Task Scheduler every hour for automatic dashboard updates.

Setup: 
  1. Run: python scripts/auto_sync_and_push.py  (to test once)
  2. Then set up Windows Task Scheduler (see setup_scheduler.bat)
"""

import os
import sys
import subprocess
from datetime import datetime

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOG_FILE = os.path.join(PROJECT_ROOT, "data", "sync_log.txt")


def log(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line, flush=True)
    try:
        os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except:
        pass


def run_selenium_sync():
    """Run Selenium sync, then check Downloads folder for latest CSV"""
    import glob
    import shutil
    
    csv_target = os.path.join(PROJECT_ROOT, "data", "submissions.csv")
    downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    
    # Step 1: Try Selenium sync
    sync_script = os.path.join(PROJECT_ROOT, "scripts", "sync_sharepoint.py")
    if os.path.exists(sync_script):
        log("Starting Selenium SharePoint sync...")
        try:
            result = subprocess.run(
                [sys.executable, sync_script],
                cwd=PROJECT_ROOT,
                capture_output=True,
                text=True,
                timeout=300
            )
            if result.returncode == 0:
                log("Selenium sync completed successfully")
                return True
            else:
                log(f"Selenium sync exited with code {result.returncode}")
        except subprocess.TimeoutExpired:
            log("Selenium sync timed out")
        except Exception as e:
            log(f"Selenium error: {e}")
    
    # Step 2: Fallback — find latest "AI Velocity" CSV in Downloads folder
    log("Checking Downloads folder for latest SharePoint CSV...")
    csv_pattern = os.path.join(downloads_dir, "AI Velocity*.csv")
    csv_files = glob.glob(csv_pattern)
    
    if not csv_files:
        log("No AI Velocity CSV files found in Downloads folder")
        return False
    
    # Get the newest one
    latest_csv = max(csv_files, key=os.path.getmtime)
    latest_mtime = os.path.getmtime(latest_csv)
    current_mtime = os.path.getmtime(csv_target) if os.path.exists(csv_target) else 0
    
    if latest_mtime > current_mtime:
        log(f"Found newer CSV: {os.path.basename(latest_csv)}")
        shutil.copy2(latest_csv, csv_target)
        log(f"Copied to {csv_target}")
        return True
    else:
        log("No newer CSV found — data is already up to date")
        return True


# Render Deploy Hook — triggers a new deploy without waiting for auto-deploy
RENDER_DEPLOY_HOOK = "https://api.render.com/deploy/srv-d7q2n3lckfvc739fk6h0?key=Whwg99rFWN8"

def trigger_render_deploy():
    """Trigger Render deploy via deploy hook URL"""
    try:
        import urllib.request
        resp = urllib.request.urlopen(RENDER_DEPLOY_HOOK, timeout=15)
        data = resp.read().decode()
        log(f"Render deploy triggered: {data}")
    except Exception as e:
        log(f"Render deploy hook failed (non-critical): {e}")


def validate_csv(csv_path):
    """Validate that the CSV has real data (not just headers or truncated columns)"""
    import csv as csv_mod
    try:
        with open(csv_path, 'r', encoding='utf-8-sig', errors='replace') as f:
            reader = csv_mod.DictReader(f)
            rows = list(reader)
        
        if not rows:
            log("CSV validation FAILED: no data rows")
            return False
        
        num_cols = len(rows[0])
        if num_cols < 30:
            log(f"CSV validation FAILED: only {num_cols} columns (expected 80+). Sync produced truncated data.")
            return False
        
        # Check that values actually have data (not all empty)
        non_empty = sum(1 for v in rows[0].values() if v and str(v).strip())
        if non_empty < 5:
            log(f"CSV validation FAILED: first row has only {non_empty} non-empty values. Data appears empty.")
            return False
        
        log(f"CSV validation passed: {len(rows)} rows, {num_cols} cols, {non_empty} values in row 1")
        return True
    except Exception as e:
        log(f"CSV validation error: {e}")
        return False


def git_push():
    """Commit and push the updated CSV to GitHub"""
    csv_path = os.path.join(PROJECT_ROOT, "data", "submissions.csv")
    
    if not os.path.exists(csv_path):
        log("ERROR: submissions.csv not found after sync!")
        return False
    
    # Check file size to verify it's valid
    size = os.path.getsize(csv_path)
    if size < 100:
        log(f"WARNING: CSV file is suspiciously small ({size} bytes). Skipping push.")
        return False
    
    # Validate CSV has real data (not truncated/empty)
    if not validate_csv(csv_path):
        log("SKIPPING push: CSV failed validation. Restoring from git...")
        subprocess.run(["git", "checkout", "--", "data/submissions.csv"], cwd=PROJECT_ROOT)
        return False
    
    log(f"CSV file: {size:,} bytes")
    
    try:
        # Check if there are changes
        result = subprocess.run(
            ["git", "diff", "--stat", "data/submissions.csv"],
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True
        )
        
        if not result.stdout.strip():
            log("No changes in CSV — data is already up to date")
            return True
        
        log(f"Changes detected: {result.stdout.strip()}")
        
        # Stage, commit, push
        subprocess.run(["git", "add", "data/submissions.csv"], cwd=PROJECT_ROOT, check=True)
        
        ts = datetime.now().strftime("%Y-%m-%d %H:%M")
        subprocess.run(
            ["git", "commit", "-m", f"Auto-sync: Update SharePoint data {ts}"],
            cwd=PROJECT_ROOT,
            check=True
        )
        
        result = subprocess.run(
            ["git", "push", "origin", "main"],
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True,
            timeout=60
        )
        
        if result.returncode == 0:
            log("Git push successful! Triggering Render deploy...")
            trigger_render_deploy()
            return True
        else:
            log(f"Git push failed: {result.stderr}")
            return False
            
    except subprocess.CalledProcessError as e:
        log(f"Git command failed: {e}")
        return False
    except Exception as e:
        log(f"ERROR during git push: {e}")
        return False


def main():
    log("=" * 60)
    log("AUTO SYNC & PUSH — Starting")
    log("=" * 60)
    
    # Step 1: Selenium sync
    sync_ok = run_selenium_sync()
    
    if not sync_ok:
        log("Sync failed — will not push. Retrying next cycle.")
        log("=" * 60)
        return
    
    # Step 2: Git push
    push_ok = git_push()
    
    if push_ok:
        log("✅ SUCCESS: Data synced and pushed. Dashboard will update shortly.")
    else:
        log("⚠️ Push failed but sync was ok. Data is saved locally.")
    
    log("=" * 60)


if __name__ == "__main__":
    main()
