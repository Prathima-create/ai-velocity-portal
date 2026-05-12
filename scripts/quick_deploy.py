#!/usr/bin/env python3
"""
quick_deploy.py — One-Click Dashboard Deploy
==============================================
Any admin can run this to regenerate + upload the QC dashboard.
Works on any machine that has:
  1. SharePoint Fincom_QC folder synced via OneDrive
  2. AWS CLI with 'qc-dashboard' profile configured
  3. Python with pandas, openpyxl

Usage:
    python scripts/quick_deploy.py          # Auto-detect and deploy
    python scripts/quick_deploy.py --help   # Show options

This is the SAME as deploy_dashboard.py but simplified for other admins.
Just double-click or run from command line!
"""

import subprocess
import sys
from pathlib import Path

def main():
    # Find deploy_dashboard.py relative to this script
    script_dir = Path(__file__).parent
    deploy_script = script_dir / "deploy_dashboard.py"
    
    if not deploy_script.exists():
        print(f"ERROR: deploy_dashboard.py not found at {deploy_script}")
        sys.exit(1)
    
    print("=" * 60)
    print("  FinCom QC Dashboard — Quick Deploy")
    print("  Running full pipeline...")
    print("=" * 60)
    print()
    
    # Pass through any command-line arguments
    args = [sys.executable, str(deploy_script)] + sys.argv[1:]
    result = subprocess.run(args)
    sys.exit(result.returncode)

if __name__ == "__main__":
    main()
