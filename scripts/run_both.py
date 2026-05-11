"""Run qc_automation.py for both General and SHT datasets."""
import sys, os, io
from pathlib import Path

# Fix Windows console encoding
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# Override BASE_FOLDER before importing
BASE = Path(r"C:\Users\pratpk\amazon.com\Automation hosting - Documents\Fincom_QC")

# Read and exec the script with modified BASE_FOLDER
script_path = Path(__file__).parent / "qc_automation.py"
code = script_path.read_text(encoding='utf-8')
code = code.replace(
    "BASE_FOLDER = DESKTOP / \"FinCom_QC\"",
    f'BASE_FOLDER = Path(r"{BASE}")'
)
# Also handle single-quote variant
code = code.replace(
    "BASE_FOLDER = DESKTOP / 'FinCom_QC'",
    f'BASE_FOLDER = Path(r"{BASE}")'
)

# Execute in namespace
ns = {}
exec(compile(code, str(script_path), 'exec'), ns)

# Run both datasets non-interactively
folders = ns['list_datasets']()
print(f"\nFound {len(folders)} datasets:")
for f in folders:
    print(f"  - {f.name}")

for folder in folders:
    if 'Apr_2026' in folder.name:
        print(f"\n>>> Processing: {folder.name}")
        ns['run_dataset'](folder)

print("\n=== ALL DONE ===")
print("Generated HTML files in each folder.")
