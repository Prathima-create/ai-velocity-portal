#!/usr/bin/env python3
"""
deploy_dashboard.py — Full Automated QC Dashboard Pipeline
===========================================================
Auto-discovers process folders in the SharePoint sync folder,
generates HTML dashboards, wraps them in a tabbed page,
and deploys to S3 → CloudFront.

Usage:
    python scripts/deploy_dashboard.py              # auto-discover latest month
    python scripts/deploy_dashboard.py --month Apr  # specific month
    python scripts/deploy_dashboard.py --year 2026  # specific year
    python scripts/deploy_dashboard.py --dry-run    # preview only, no upload

Pipeline:
    1. Scan SharePoint folder → discover Process_Month_Year folders
    2. Find the latest month (or use --month/--year)
    3. For each process folder: run qc_automation to generate HTML dashboards
    4. Combine all process dashboards into a tabbed index.html
    5. Upload index.html to S3
    6. Invalidate CloudFront cache

Requirements:
    - pandas, openpyxl, matplotlib (for qc_automation.py)
    - aws CLI configured with 'qc-dashboard' profile
"""

import os, sys, json, re, io, argparse, subprocess
from pathlib import Path
from datetime import datetime
from collections import defaultdict

# Ensure UTF-8 output
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# ============================================================
# CONFIGURATION — Change these if your setup differs
# ============================================================
# SharePoint sync folder (where process folders live)
SHAREPOINT_BASE = Path(r"C:\Users\pratpk\amazon.com\Automation hosting - Documents\Fincom_QC")

# S3 bucket and CloudFront
S3_BUCKET = "fincom-qc-data"
S3_PATHS = ["index.html", "dashboard/index.html"]  # upload to both paths
CF_DISTRIBUTION_ID = "E1ZWY804XF47SW"
AWS_PROFILE = "qc-dashboard"

# Dashboard URL
DASHBOARD_URL = "https://d1ggg5slalzwhl.cloudfront.net"

# Month ordering
MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

# Tab display names (process_name -> tab label)
# Add new processes here as they appear
TAB_LABELS = {
    "General": "General",
    "SHT": "Shipment Hold",
    # Future: "Returns": "Returns", "Payments": "Payments", etc.
}


# ============================================================
# STEP 1: Discover process folders
# ============================================================
def discover_folders(base: Path):
    """Scan base folder for Process_Month_Year directories.
    Returns dict: {(month, year): {process_name: folder_path}}"""
    if not base.exists():
        print(f"ERROR: SharePoint folder not found: {base}")
        sys.exit(1)

    months_data = defaultdict(dict)
    for d in sorted(base.iterdir()):
        if not d.is_dir():
            continue
        parts = d.name.split('_')
        if len(parts) < 3:
            continue
        process = parts[0]
        month = parts[1]
        year = parts[2]
        if month not in MONTHS:
            continue
        if not year.isdigit():
            continue
        months_data[(month, year)][process] = d

    return months_data


def find_latest_month(months_data, target_month=None, target_year=None):
    """Find the latest month key, or match target_month/target_year."""
    if not months_data:
        return None

    def sort_key(k):
        m, y = k
        return (int(y), MONTHS.index(m) if m in MONTHS else 0)

    sorted_keys = sorted(months_data.keys(), key=sort_key, reverse=True)

    if target_month and target_year:
        target = (target_month, target_year)
        return target if target in months_data else None
    elif target_month:
        # Find latest year with this month
        for k in sorted_keys:
            if k[0] == target_month:
                return k
    elif target_year:
        # Find latest month in this year
        for k in sorted_keys:
            if k[1] == target_year:
                return k

    return sorted_keys[0] if sorted_keys else None


# ============================================================
# STEP 2: Generate dashboards (uses qc_automation.py)
# ============================================================
def generate_dashboards(process_folders: dict, script_dir: Path):
    """Run qc_automation.py for each process folder.
    Returns dict: {process_name: html_path}"""
    qc_script = script_dir / "qc_automation.py"
    if not qc_script.exists():
        print(f"ERROR: qc_automation.py not found at {qc_script}")
        sys.exit(1)

    # Import qc_automation dynamically
    import importlib.util
    spec = importlib.util.spec_from_file_location("qc_automation", str(qc_script))
    qc = importlib.util.module_from_spec(spec)

    # Override BASE_FOLDER to point to SharePoint
    # We need to temporarily modify the module's BASE_FOLDER
    original_base = None

    try:
        spec.loader.exec_module(qc)
        original_base = qc.BASE_FOLDER
        qc.BASE_FOLDER = process_folders[list(process_folders.keys())[0]].parent
    except Exception as e:
        print(f"WARNING: Could not load qc_automation module: {e}")
        print("Falling back to subprocess execution...")
        return _generate_via_subprocess(process_folders, qc_script)

    html_files = {}
    for process_name, folder in process_folders.items():
        print(f"\n{'='*60}")
        print(f"Generating dashboard: {process_name} ({folder.name})")
        print(f"{'='*60}")

        # Check if Fincom_Process.csv exists
        if not (folder / "Fincom_Process.csv").exists():
            print(f"  SKIP: No Fincom_Process.csv in {folder.name}")
            continue

        try:
            qc.run_dataset(folder)
            # Find the generated HTML
            html_path = folder / f"QC_Dashboard_{folder.name}.html"
            if html_path.exists():
                html_files[process_name] = html_path
                print(f"  ✓ Dashboard generated: {html_path.name}")
            else:
                print(f"  WARNING: Expected HTML not found: {html_path.name}")
        except Exception as e:
            print(f"  ERROR generating {process_name}: {e}")
            import traceback
            traceback.print_exc()

    return html_files


def _generate_via_subprocess(process_folders, qc_script):
    """Fallback: run qc_automation.py via subprocess for each folder."""
    html_files = {}
    for process_name, folder in process_folders.items():
        if not (folder / "Fincom_Process.csv").exists():
            print(f"  SKIP: No Fincom_Process.csv in {folder.name}")
            continue

        print(f"\nGenerating dashboard for {process_name} via subprocess...")
        # We need to create a small wrapper that calls run_dataset non-interactively
        wrapper = f'''
import sys
sys.path.insert(0, r"{qc_script.parent}")
from qc_automation import run_dataset, BASE_FOLDER
from pathlib import Path
# Override BASE_FOLDER
import qc_automation
qc_automation.BASE_FOLDER = Path(r"{folder.parent}")
run_dataset(Path(r"{folder}"))
'''
        try:
            result = subprocess.run(
                [sys.executable, "-c", wrapper],
                capture_output=True, text=True, timeout=300
            )
            print(result.stdout[-500:] if len(result.stdout) > 500 else result.stdout)
            if result.returncode != 0:
                print(f"  STDERR: {result.stderr[-300:]}")

            html_path = folder / f"QC_Dashboard_{folder.name}.html"
            if html_path.exists():
                html_files[process_name] = html_path
                print(f"  ✓ Dashboard generated: {html_path.name}")
        except Exception as e:
            print(f"  ERROR: {e}")

    return html_files


# ============================================================
# STEP 3: Create tabbed wrapper (dynamic number of tabs)
# ============================================================
def create_tabbed_dashboard(html_files: dict, month: str, year: str, output_path: Path):
    """Combine multiple process dashboards into a single tabbed HTML page.
    Supports any number of process tabs dynamically."""

    if not html_files:
        print("ERROR: No HTML files to combine!")
        return None

    print(f"\nCreating tabbed dashboard with {len(html_files)} tabs...")

    # Read each HTML file
    dashboards = {}
    for process_name, html_path in html_files.items():
        print(f"  Reading: {html_path.name} ({html_path.stat().st_size // 1024}KB)")
        dashboards[process_name] = html_path.read_text(encoding='utf-8')

    # Sort processes: General first, then alphabetical
    def sort_key(name):
        if name.lower() == 'general':
            return (0, name)
        return (1, name)

    sorted_processes = sorted(dashboards.keys(), key=sort_key)

    import html as html_module

    # Build tab buttons
    tab_buttons = []
    for i, proc in enumerate(sorted_processes):
        label = TAB_LABELS.get(proc, proc)
        active = ' active' if i == 0 else ''
        tab_buttons.append(
            f'  <button class="tab-btn{active}" onclick="switchTab(\'{proc.lower()}\', this)">{label}</button>'
        )

    # Build iframes with srcdoc ATTRIBUTE (entity-encoded) - this is the approach that works
    iframe_elements = []
    for i, proc in enumerate(sorted_processes):
        active = ' active' if i == 0 else ''
        tab_id = proc.lower()
        # Entity-encode the HTML for srcdoc attribute
        encoded_html = html_module.escape(dashboards[proc], quote=True)
        iframe_elements.append(
            f'<iframe id="frame-{tab_id}" class="tab-frame{active}"\n'
            f'  srcdoc="{encoded_html}"></iframe>'
        )

    wrapper = f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>FinCom QC Dashboard - {month} {year}</title>
<style>
body {{ margin:0; font-family: 'Segoe UI', Arial, sans-serif; background: #f0f2f5; }}
nav.tab-bar {{ display:flex; background:#1a1a2e; padding:0 20px; position:sticky; top:0; z-index:9999; }}
.tab-btn {{ padding:14px 28px; color:#aaa; cursor:pointer; border:none; background:none; font-size:15px; font-weight:600; border-bottom:3px solid transparent; transition:all 0.2s; }}
.tab-btn:hover {{ color:#fff; }}
.tab-btn.active {{ color:#fff; border-bottom-color:#4285f4; background:rgba(66,133,244,0.1); }}
.tab-frame {{ display:none; width:100%; border:none; height:calc(100vh - 52px); }}
.tab-frame.active {{ display:block; }}
</style>
</head>
<body>
<nav class="tab-bar">
{chr(10).join(tab_buttons)}
</nav>

{chr(10).join(iframe_elements)}

<script>
function switchTab(tab, btn) {{
  document.querySelectorAll('.tab-frame').forEach(function(el) {{ el.classList.remove('active'); }});
  document.querySelectorAll('.tab-btn').forEach(function(el) {{ el.classList.remove('active'); }});
  document.getElementById('frame-' + tab).classList.add('active');
  btn.classList.add('active');
}}
</script>
</body>
</html>'''

    output_path.write_text(wrapper, encoding='utf-8')
    size_kb = output_path.stat().st_size // 1024
    print(f"  ✓ Tabbed dashboard saved: {output_path} ({size_kb} KB)")
    return output_path


# ============================================================
# STEP 4: Upload to S3 + Invalidate CloudFront
# ============================================================
def upload_to_s3(html_path: Path, dry_run=False):
    """Upload HTML to S3 and invalidate CloudFront."""
    if dry_run:
        print(f"\n[DRY RUN] Would upload {html_path.name} to S3 and invalidate CloudFront")
        return True

    print(f"\nUploading to S3...")
    success = True

    for s3_key in S3_PATHS:
        cmd = [
            "aws", "s3", "cp", str(html_path),
            f"s3://{S3_BUCKET}/{s3_key}",
            "--content-type", "text/html",
            "--profile", AWS_PROFILE
        ]
        print(f"  → s3://{S3_BUCKET}/{s3_key}")
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            print(f"  ERROR: {result.stderr}")
            success = False
        else:
            print(f"  ✓ Uploaded")

    # Invalidate CloudFront
    print(f"\nInvalidating CloudFront ({CF_DISTRIBUTION_ID})...")
    cmd = [
        "aws", "cloudfront", "create-invalidation",
        "--distribution-id", CF_DISTRIBUTION_ID,
        "--paths", "/*",
        "--profile", AWS_PROFILE,
        "--query", "Invalidation.Status",
        "--output", "text"
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode == 0:
        print(f"  ✓ Invalidation: {result.stdout.strip()}")
    else:
        print(f"  ERROR: {result.stderr}")
        success = False

    if success:
        print(f"\n✓ Dashboard live at: {DASHBOARD_URL}")

    return success


# ============================================================
# MAIN
# ============================================================
def main():
    parser = argparse.ArgumentParser(description="Automated QC Dashboard Pipeline")
    parser.add_argument("--month", help="Target month (e.g., Apr, May)")
    parser.add_argument("--year", help="Target year (e.g., 2026)")
    parser.add_argument("--dry-run", action="store_true", help="Preview only, don't upload")
    parser.add_argument("--skip-generate", action="store_true",
                        help="Skip dashboard generation, use existing HTMLs")
    parser.add_argument("--base", help="Override SharePoint base folder path")
    args = parser.parse_args()

    global SHAREPOINT_BASE
    if args.base:
        SHAREPOINT_BASE = Path(args.base)

    print("=" * 60)
    print("FinCom QC Dashboard — Automated Pipeline")
    print("=" * 60)
    print(f"SharePoint folder: {SHAREPOINT_BASE}")
    print(f"S3 bucket:         {S3_BUCKET}")
    print(f"CloudFront:        {CF_DISTRIBUTION_ID}")
    print()

    # Step 1: Discover folders
    print("Step 1: Discovering process folders...")
    months_data = discover_folders(SHAREPOINT_BASE)

    if not months_data:
        print("ERROR: No process folders found!")
        print(f"Expected folders like 'General_Apr_2026' in {SHAREPOINT_BASE}")
        sys.exit(1)

    print(f"  Found {sum(len(v) for v in months_data.values())} folders across {len(months_data)} months:")
    for (m, y), procs in sorted(months_data.items(), key=lambda x: (int(x[0][1]), MONTHS.index(x[0][0]) if x[0][0] in MONTHS else 0)):
        proc_list = ", ".join(sorted(procs.keys()))
        print(f"    {m} {y}: {proc_list}")

    # Find target month
    latest = find_latest_month(months_data, args.month, args.year)
    if not latest:
        print(f"ERROR: Could not find target month ({args.month or 'latest'} {args.year or ''})")
        sys.exit(1)

    month, year = latest
    process_folders = months_data[latest]
    print(f"\n  → Target: {month} {year} ({len(process_folders)} processes: {', '.join(sorted(process_folders.keys()))})")

    # Step 2: Generate dashboards
    script_dir = Path(__file__).parent
    if args.skip_generate:
        print("\nStep 2: SKIPPED (--skip-generate)")
        html_files = {}
        for proc_name, folder in process_folders.items():
            html_path = folder / f"QC_Dashboard_{folder.name}.html"
            if html_path.exists():
                html_files[proc_name] = html_path
                print(f"  Using existing: {html_path.name}")
            else:
                print(f"  WARNING: {html_path.name} not found, skipping {proc_name}")
    else:
        print("\nStep 2: Generating dashboards...")
        html_files = generate_dashboards(process_folders, script_dir)

    if not html_files:
        print("ERROR: No dashboards generated!")
        sys.exit(1)

    print(f"\n  ✓ {len(html_files)} dashboards ready: {', '.join(sorted(html_files.keys()))}")

    # Step 3: Create tabbed wrapper
    print("\nStep 3: Creating tabbed dashboard...")
    output_path = SHAREPOINT_BASE / "index.html"
    result = create_tabbed_dashboard(html_files, month, year, output_path)

    if not result:
        print("ERROR: Failed to create tabbed dashboard!")
        sys.exit(1)

    # Step 4: Upload to S3
    print("\nStep 4: Deploying to CloudFront...")
    upload_to_s3(output_path, dry_run=args.dry_run)

    print("\n" + "=" * 60)
    print("PIPELINE COMPLETE")
    print("=" * 60)
    print(f"  Processes: {', '.join(sorted(html_files.keys()))}")
    print(f"  Month:     {month} {year}")
    print(f"  Dashboard: {DASHBOARD_URL}")
    if args.dry_run:
        print("  (DRY RUN — no upload performed)")


if __name__ == "__main__":
    main()
