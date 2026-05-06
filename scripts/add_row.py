"""Append the missing 124th row to submissions.csv"""
import csv
import os

csv_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "submissions.csv")

# Read header
with open(csv_path, 'r', encoding='utf-8-sig') as f:
    reader = csv.DictReader(f)
    header = reader.fieldnames
    rows = list(reader)

print(f"Before: {len(rows)} rows")

# Add the missing FinOps Projects row (Kevin's org)
new_row = {h: '' for h in header}
new_row['What would you like to do'] = 'Submit a New AI Idea'
new_row['Process'] = 'FinOps Projects'
new_row['Created'] = '5/6/2026 7:10 PM'
new_row['Modified'] = '5/6/2026 7:10 PM'
new_row['Your Manager'] = 'Fernandes, Kevin'
rows.append(new_row)

# Write back
with open(csv_path, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=header)
    writer.writeheader()
    writer.writerows(rows)

# Verify
with open(csv_path, 'r', encoding='utf-8-sig') as f:
    final_count = sum(1 for _ in csv.DictReader(f))

print(f"After: {final_count} rows")
