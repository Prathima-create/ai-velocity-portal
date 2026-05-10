#!/bin/bash
# ============================================================
# refresh_dashboard.sh
# Runs on EC2 every 30 minutes via cron
# Downloads CSVs from S3 → runs qc_automation.py → serves HTML
# ============================================================

set -e

# Config
S3_BUCKET="fincom-qc-data"
CURRENT_PREFIX="General_Apr_2026/"
PREV_PREFIX="General_Mar_2026/"
WORK_DIR="/home/ec2-user/qc-dashboard"
DATA_DIR="$WORK_DIR/data"
WEB_DIR="/home/ec2-user/web"
LOG_FILE="$WORK_DIR/refresh.log"

echo "$(date) — Starting refresh..." >> "$LOG_FILE"

# Create dirs
mkdir -p "$DATA_DIR/current" "$DATA_DIR/previous" "$WEB_DIR"

# Download current month CSVs from S3
aws s3 sync "s3://$S3_BUCKET/$CURRENT_PREFIX" "$DATA_DIR/current/" --delete 2>> "$LOG_FILE"

# Download previous month CSVs from S3 (for MoM comparison)
aws s3 sync "s3://$S3_BUCKET/$PREV_PREFIX" "$DATA_DIR/previous/" --delete 2>> "$LOG_FILE"

# Run qc_automation.py to generate HTML
cd "$WORK_DIR"
python3 qc_automation.py --data-dir "$DATA_DIR/current" --prev-dir "$DATA_DIR/previous" --output-dir "$WEB_DIR" 2>> "$LOG_FILE"

# Rename output to index.html for web serving
HTML_FILE=$(ls "$WEB_DIR"/QC_Dashboard_*.html 2>/dev/null | head -1)
if [ -n "$HTML_FILE" ]; then
    cp "$HTML_FILE" "$WEB_DIR/index.html"
    echo "$(date) — Dashboard regenerated: $HTML_FILE" >> "$LOG_FILE"
else
    echo "$(date) — ERROR: No HTML file generated!" >> "$LOG_FILE"
fi

# Ensure web server is running on port 8503
if ! pgrep -f "http.server 8503" > /dev/null; then
    cd "$WEB_DIR"
    nohup python3 -m http.server 8503 --bind 0.0.0.0 > /dev/null 2>&1 &
    echo "$(date) — Web server started on port 8503" >> "$LOG_FILE"
fi

echo "$(date) — Refresh complete." >> "$LOG_FILE"
