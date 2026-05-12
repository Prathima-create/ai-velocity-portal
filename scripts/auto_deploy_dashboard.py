#!/usr/bin/env python3
"""
auto_deploy_dashboard.py — Auto-Deploy Watcher
================================================
Watches for CSV changes in your SharePoint Fincom_QC folder and
automatically re-deploys the dashboard when changes are detected.

Runs every 2 minutes. Only re-deploys if CSV files have been modified.

Usage:
    python scripts/auto_deploy_dashboard.py          # Run watcher (every 2 min)
    python scripts/auto_deploy_dashboard.py --once   # Run once and exit
    python scripts/auto_deploy_dashboard.py --interval 60  # Check every 60 seconds

To run in background:
    start /min python scripts/auto_deploy_dashboard.py
    
Or use the batch file:
    scripts\\start_auto_deploy.bat
"""

import os, sys, time, hashlib, json
from pathlib import Path
from datetime import datetime

# Add scripts dir to path so we can import deploy_dashboard
sys.path.insert(0, str(Path(__file__).parent))

# ─── Config ───────────────────────────────────────────────────────────────────
CHECK_INTERVAL = 120  # seconds (2 minutes)
STATE_FILE = Path(__file__).parent / ".deploy_state.json"


def find_sharepoint_base():
    """Find the SharePoint Fincom_QC folder."""
    home = Path.home()
    search_roots = [home, home / "Documents" / "DRIVE", home / "Documents", home / "OneDrive"]
    
    for root in search_roots:
        amazon_dir = root / "amazon.com"
        if not amazon_dir.exists():
            continue
        try:
            for d in amazon_dir.iterdir():
                if d.is_dir() and "automation" in d.name.lower() and "hosting" in d.name.lower():
                    fq = d / "Fincom_QC"
                    if fq.exists():
                        return fq
        except PermissionError:
            continue
    
    return home / "amazon.com" / "Automation hosting - Documents" / "Fincom_QC"


def get_csv_fingerprint(base_path: Path):
    """Generate a fingerprint of all CSV files (based on modification times + sizes).
    This is fast — doesn't read file contents, just metadata."""
    fingerprint_parts = []
    
    for csv_file in sorted(base_path.rglob("*.csv")):
        try:
            stat = csv_file.stat()
            fingerprint_parts.append(f"{csv_file.name}:{stat.st_size}:{stat.st_mtime:.0f}")
        except (OSError, PermissionError):
            continue
    
    if not fingerprint_parts:
        return None
    
    combined = "|".join(fingerprint_parts)
    return hashlib.md5(combined.encode()).hexdigest()


def load_state():
    """Load previous fingerprint from state file."""
    if STATE_FILE.exists():
        try:
            return json.loads(STATE_FILE.read_text())
        except:
            pass
    return {"fingerprint": None, "last_deploy": None}


def save_state(fingerprint):
    """Save current fingerprint to state file."""
    state = {
        "fingerprint": fingerprint,
        "last_deploy": datetime.now().isoformat()
    }
    STATE_FILE.write_text(json.dumps(state, indent=2))


def run_deploy():
    """Run the deploy_dashboard.py pipeline."""
    print(f"\n{'='*60}")
    print(f"[{datetime.now().strftime('%H:%M:%S')}] CSV CHANGES DETECTED — Deploying...")
    print(f"{'='*60}\n")
    
    try:
        import deploy_dashboard
        deploy_dashboard.main()
        return True
    except SystemExit:
        # deploy_dashboard calls sys.exit() on success
        return True
    except Exception as e:
        print(f"\nERROR during deploy: {e}")
        import traceback
        traceback.print_exc()
        return False


def log(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Auto-deploy dashboard on CSV changes")
    parser.add_argument("--once", action="store_true", help="Run once and exit")
    parser.add_argument("--interval", type=int, default=CHECK_INTERVAL,
                        help=f"Check interval in seconds (default: {CHECK_INTERVAL})")
    parser.add_argument("--force", action="store_true", help="Force deploy even if no changes")
    args = parser.parse_args()

    base_path = find_sharepoint_base()
    
    print("=" * 60)
    print("FinCom QC Dashboard — Auto-Deploy Watcher")
    print("=" * 60)
    print(f"  Watching:   {base_path}")
    print(f"  Interval:   {args.interval} seconds")
    print(f"  Mode:       {'Once' if args.once else 'Continuous'}")
    print(f"  Press Ctrl+C to stop")
    print("=" * 60)
    print()

    if not base_path.exists():
        print(f"ERROR: SharePoint folder not found: {base_path}")
        print("Make sure SharePoint/OneDrive is synced.")
        sys.exit(1)

    state = load_state()
    
    if args.force:
        log("Force deploy requested...")
        run_deploy()
        fingerprint = get_csv_fingerprint(base_path)
        if fingerprint:
            save_state(fingerprint)
        if args.once:
            return

    while True:
        try:
            fingerprint = get_csv_fingerprint(base_path)
            
            if fingerprint is None:
                log("No CSV files found yet. Waiting...")
            elif fingerprint != state.get("fingerprint"):
                log(f"Changes detected! (fingerprint: {fingerprint[:8]}...)")
                success = run_deploy()
                if success:
                    save_state(fingerprint)
                    state["fingerprint"] = fingerprint
                    log("✓ Deploy complete! Dashboard updated.")
                else:
                    log("✗ Deploy failed. Will retry next cycle.")
            else:
                log(f"No changes. (last deploy: {state.get('last_deploy', 'never')})")
            
            if args.once:
                break
            
            time.sleep(args.interval)
            
        except KeyboardInterrupt:
            print("\n\nStopped by user. Goodbye!")
            break
        except Exception as e:
            log(f"ERROR: {e}")
            if args.once:
                break
            time.sleep(args.interval)


if __name__ == "__main__":
    main()
