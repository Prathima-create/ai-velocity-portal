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
    """Run the existing Selenium sync script"""
    sync_script = os.path.join(PROJECT_ROOT, "scripts", "sync_sharepoint.py")
    if not os.path.exists(sync_script):
        log("ERROR: sync_sharepoint.py not found!")
        return False
    
    log("Starting Selenium SharePoint sync...")
    try:
        result = subprocess.run(
            [sys.executable, sync_script],
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True,
            timeout=300  # 5 min timeout
        )
        if result.returncode == 0:
            log("Selenium sync completed successfully")
            log(f"Output: {result.stdout[-200:]}" if result.stdout else "No output")
            return True
        else:
            log(f"Selenium sync failed (exit code {result.returncode})")
            log(f"Error: {result.stderr[-200:]}" if result.stderr else "No error output")
            return False
    except subprocess.TimeoutExpired:
        log("ERROR: Selenium sync timed out after 5 minutes")
        return False
    except Exception as e:
        log(f"ERROR: {e}")
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
            log("Git push successful! Render will auto-deploy in ~2 minutes.")
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
