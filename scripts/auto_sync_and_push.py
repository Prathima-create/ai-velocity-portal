"""
Auto Sync & Push
  1. Runs Edge-based SharePoint sync (copies profile to avoid lock issues)
  2. Validates the CSV
  3. Pushes to GitHub
  4. Render auto-deploys from the push

Run manually:   python scripts/auto_sync_and_push.py
Schedule:       Run setup_scheduler.bat (creates hourly Windows Task)
"""

import os
import sys
import csv as csv_mod
import subprocess
from datetime import datetime

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CSV_PATH = os.path.join(PROJECT_ROOT, "data", "submissions.csv")
LOG_FILE = os.path.join(PROJECT_ROOT, "data", "sync_log.txt")


def log(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line, flush=True)
    try:
        os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def run_sync():
    """Run the SharePoint sync script"""
    sync_script = os.path.join(PROJECT_ROOT, "scripts", "sync_sharepoint.py")
    log("Running SharePoint sync...")
    try:
        result = subprocess.run(
            [sys.executable, sync_script],
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True,
            timeout=180,
        )
        # Print sync output
        if result.stdout:
            for line in result.stdout.strip().split("\n"):
                log(f"  {line}")
        if result.returncode == 0:
            log("Sync completed successfully")
            return True
        else:
            log(f"Sync failed (exit code {result.returncode})")
            if result.stderr:
                log(f"  stderr: {result.stderr[:300]}")
            return False
    except subprocess.TimeoutExpired:
        log("Sync timed out after 3 minutes")
        return False
    except Exception as e:
        log(f"Sync error: {e}")
        return False


def fallback_check_downloads():
    """Fallback: check Downloads folder for a newer CSV export"""
    import glob
    import shutil

    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    patterns = [
        os.path.join(downloads, "AI Velocity*.csv"),
        os.path.join(downloads, "AI_Velocity*.csv"),
    ]

    csv_files = []
    for p in patterns:
        csv_files.extend(glob.glob(p))

    if not csv_files:
        return False

    latest = max(csv_files, key=os.path.getmtime)
    latest_time = os.path.getmtime(latest)
    current_time = os.path.getmtime(CSV_PATH) if os.path.exists(CSV_PATH) else 0

    if latest_time > current_time:
        log(f"Found newer CSV in Downloads: {os.path.basename(latest)}")
        shutil.copy2(latest, CSV_PATH)
        log("Copied to data/submissions.csv")
        return True

    log("No newer CSV in Downloads")
    return False


def validate_csv():
    """Check that the CSV has real data"""
    if not os.path.exists(CSV_PATH):
        log("CSV file not found")
        return False

    size = os.path.getsize(CSV_PATH)
    if size < 500:
        log(f"CSV too small ({size} bytes)")
        return False

    try:
        with open(CSV_PATH, "r", encoding="utf-8-sig", errors="replace") as f:
            reader = csv_mod.DictReader(f)
            rows = list(reader)

        if len(rows) < 10:
            log(f"CSV has only {len(rows)} rows — seems incomplete")
            return False

        cols = len(rows[0]) if rows else 0
        if cols < 20:
            log(f"CSV has only {cols} columns — seems truncated")
            return False

        log(f"CSV valid: {len(rows)} rows, {cols} columns, {size:,} bytes")
        return True
    except Exception as e:
        log(f"CSV validation error: {e}")
        return False


def git_push():
    """Commit and push the updated CSV"""
    try:
        # Check for changes
        diff = subprocess.run(
            ["git", "diff", "--stat", "data/submissions.csv"],
            cwd=PROJECT_ROOT, capture_output=True, text=True,
        )
        if not diff.stdout.strip():
            log("No changes in CSV — already up to date")
            return True

        log(f"Changes detected: {diff.stdout.strip()}")

        # Stage and commit
        subprocess.run(["git", "add", "data/submissions.csv"], cwd=PROJECT_ROOT, check=True)

        ts = datetime.now().strftime("%Y-%m-%d %H:%M")
        subprocess.run(
            ["git", "commit", "-m", f"chore: Auto-sync SharePoint data {ts}"],
            cwd=PROJECT_ROOT, check=True,
        )

        # Push
        result = subprocess.run(
            ["git", "push"],
            cwd=PROJECT_ROOT, capture_output=True, text=True, timeout=60,
        )

        if result.returncode == 0:
            log("Git push successful — Render will auto-deploy")
            return True
        else:
            log(f"Git push failed: {result.stderr}")
            return False

    except subprocess.CalledProcessError as e:
        log(f"Git error: {e}")
        return False
    except Exception as e:
        log(f"Push error: {e}")
        return False


def main():
    log("=" * 55)
    log("AUTO SYNC & PUSH — Starting")
    log("=" * 55)

    # Step 1: Try Edge-based sync
    synced = run_sync()

    # Step 2: Fallback — check Downloads for manual export
    if not synced:
        log("Edge sync failed. Checking Downloads folder as fallback...")
        synced = fallback_check_downloads()

    if not synced:
        log("No new data available. Exiting.")
        log("=" * 55)
        return

    # Step 3: Validate
    if not validate_csv():
        log("CSV validation failed — restoring previous version")
        subprocess.run(["git", "checkout", "--", "data/submissions.csv"], cwd=PROJECT_ROOT)
        log("=" * 55)
        return

    # Step 4: Push to GitHub
    pushed = git_push()

    if pushed:
        log("SUCCESS — Data synced and pushed. Dashboard will update in ~2 min.")
    else:
        log("Push failed but CSV is updated locally.")

    log("=" * 55)


if __name__ == "__main__":
    main()
