@echo off
REM ============================================================
REM SharePoint → S3 Auto-Sync (Windows Task Scheduler)
REM Watches the SharePoint-synced folder and uploads CSVs to S3
REM Schedule: Every 30 minutes via Task Scheduler
REM ============================================================

set S3_BUCKET=fincom-qc-data
set SHAREPOINT_FOLDER=C:\Users\pratpk\amazon.com\Automation hosting - Documents\General_Apr_2026
set S3_PREFIX=General_Apr_2026/
set AWS_PROFILE=auditiq

echo [%date% %time%] Syncing SharePoint to S3...

REM Upload all CSVs from SharePoint folder to S3
"C:\Program Files\Amazon\AWSCLIV2\aws.exe" s3 sync "%SHAREPOINT_FOLDER%" "s3://%S3_BUCKET%/%S3_PREFIX%" --exclude "*" --include "*.csv" --profile %AWS_PROFILE%

echo [%date% %time%] Sync complete.
