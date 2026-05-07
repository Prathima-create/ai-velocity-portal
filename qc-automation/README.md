# 📊 FinCom QC Automation Dashboard

**Live, shareable QC dashboard** hosted on EC2 behind corporate VPN.  
Syncs data automatically from SharePoint → S3 → Streamlit app.

---

## 🏗️ Architecture

```
┌──────────────────────┐     ┌─────────────────┐     ┌──────────────────────┐
│  Amazon SharePoint   │────▶│   S3 Bucket     │────▶│  EC2 (Streamlit)     │
│  (QC Audit CSVs)     │     │  fincom-qc-data │     │  http://10.x.x.x:8501│
│  Someone else's acct │     │  (encrypted)    │     │  VPN-accessible only │
└──────────────────────┘     └─────────────────┘     └──────────────────────┘
         │                           ▲                         │
         │    Sync every 30 min      │                         │
         └───────────────────────────┘                         │
              (Edge SSO + Selenium)                             │
                                                               ▼
                                                    ┌──────────────────┐
                                                    │  Team members    │
                                                    │  (VPN link)      │
                                                    └──────────────────┘
```

---

## 📁 File Structure

```
qc-automation/
├── app.py                      # Streamlit dashboard (main app)
├── sync_sharepoint_to_s3.py    # SharePoint → S3 sync script
├── requirements.txt            # Python dependencies
├── .env.template               # Environment config template
├── .streamlit/
│   └── config.toml             # Streamlit settings
├── data/                       # Local data cache (auto-created)
├── deploy/
│   ├── setup_ec2.sh            # One-click EC2 setup script
│   └── create_s3_bucket.sh     # Create & configure S3 bucket
└── README.md                   # This file
```

---

## 🚀 Deployment Guide (Step by Step)

### Prerequisites
- AWS Account (via Isengard)
- EC2 instance (t3.micro, Amazon Linux 2023)
- IAM role with S3 access
- Security Group: Allow port **8501** from **10.0.0.0/8** (VPN)

---

### Step 1: Create S3 Bucket

```bash
# On your local machine (with AWS CLI configured)
cd qc-automation/deploy
chmod +x create_s3_bucket.sh
./create_s3_bucket.sh
```

This creates a **private, encrypted** S3 bucket with:
- ✅ Versioning enabled
- ✅ Public access blocked
- ✅ AES-256 encryption
- ✅ 90-day old version cleanup

---

### Step 2: Launch EC2 Instance

1. Go to **AWS Console → EC2 → Launch Instance**
2. Settings:
   - **AMI**: Amazon Linux 2023
   - **Instance type**: t3.micro (free tier eligible)
   - **IAM role**: Attach role with S3 access policy (see `create_s3_bucket.sh` output)
   - **Security Group**:
     - Inbound: TCP 8501 from 10.0.0.0/8 (Amazon VPN)
     - Inbound: TCP 22 from your IP (SSH)
   - **Storage**: 10 GB gp3

3. Note down the **private IP** (e.g., `10.0.5.42`)

---

### Step 3: Deploy on EC2

```bash
# SSH into your EC2 instance
ssh ec2-user@10.x.x.x

# Download and run setup script
curl -O https://raw.githubusercontent.com/Prathima-create/ai-velocity-portal/main/qc-automation/deploy/setup_ec2.sh
chmod +x setup_ec2.sh
./setup_ec2.sh
```

Or manually:
```bash
# Clone repo
git clone https://github.com/Prathima-create/ai-velocity-portal.git /opt/qc-dashboard
cd /opt/qc-dashboard/qc-automation

# Setup Python
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# Configure
cp .env.template .env
nano .env  # Edit with your S3 bucket, SharePoint URL
```

---

### Step 4: Configure Environment

Edit `/opt/qc-dashboard/qc-automation/.env`:

```env
QC_S3_BUCKET=fincom-qc-data
QC_S3_PREFIX=current/
QC_SHAREPOINT_SITE=https://amazon.sharepoint.com/sites/YOUR-SITE
QC_SHAREPOINT_FOLDER=/sites/YOUR-SITE/Shared Documents/QC_Data
```

---

### Step 5: Start the Dashboard

```bash
# Start Streamlit
sudo systemctl start qc-dashboard

# Start auto-sync (every 30 min)
sudo systemctl start qc-sync.timer

# Check status
sudo systemctl status qc-dashboard
```

---

### Step 6: Share the Link! 🎉

Your dashboard is now live at:
```
http://10.x.x.x:8501
```

Share this link with anyone on the Amazon VPN. They can view the dashboard immediately!

---

## 🔄 Data Sync Options

### Option A: Automatic (Recommended)
The `qc-sync.timer` runs every 30 minutes:
1. Opens SharePoint via Edge SSO
2. Downloads latest CSV files
3. Uploads to S3
4. Streamlit auto-refreshes (5-min cache)

### Option B: Manual Upload
Upload CSVs directly to S3:
```bash
aws s3 cp Fincom_Process.csv s3://fincom-qc-data/current/
aws s3 cp Fincom_Analyst.csv s3://fincom-qc-data/current/
```

### Option C: Local Testing
Place CSVs in `data/` folder and run locally:
```bash
streamlit run app.py
```

---

## 🔐 Security Features

| Feature | Implementation |
|---------|---------------|
| Network isolation | EC2 only accessible from VPN (10.0.0.0/8) |
| Data encryption | S3 AES-256, in-transit via VPN |
| No public access | S3 bucket public access blocked |
| No credentials in code | IAM role on EC2, env vars for config |
| PII protection | Data never leaves Amazon network |
| Access control | VPN = authenticated Amazon employee |

---

## 🛠️ Management Commands

```bash
# Dashboard
sudo systemctl start qc-dashboard      # Start
sudo systemctl stop qc-dashboard       # Stop
sudo systemctl restart qc-dashboard    # Restart
sudo systemctl status qc-dashboard     # Status
sudo journalctl -u qc-dashboard -f     # Live logs

# Sync
sudo systemctl start qc-sync           # Run sync now
sudo systemctl status qc-sync.timer    # Check timer
sudo journalctl -u qc-sync -f          # Sync logs

# Manual sync
cd /opt/qc-dashboard/qc-automation
source venv/bin/activate
python sync_sharepoint_to_s3.py --local-only  # Download only
python sync_sharepoint_to_s3.py               # Download + S3 upload
```

---

## 🧪 Local Development

```bash
# From this folder (qc-automation/)
pip install -r requirements.txt

# Place test CSVs in data/ folder
# Then run:
streamlit run app.py

# Opens at http://localhost:8501
```

---

## 📋 Expected CSV Files

The dashboard expects these files (in S3 or `data/` folder):

| File | Purpose | Required? |
|------|---------|-----------|
| `Fincom_Process.csv` | Process-level audit data | ✅ Yes |
| `Fincom_Analyst.csv` | Analyst-level audit data | Recommended |
| `Disputes.csv` | Dispute tracking | Optional |
| `Rectification.csv` | Rectification tracking | Optional |
| `IVOC.csv` | IVOC tracking | Optional |
| `Defect Reduction.csv` | Defect reduction tracking | Optional |

---

## ❓ Troubleshooting

| Issue | Solution |
|-------|----------|
| "No data files found" | Check S3 bucket has CSVs, or place them in `data/` |
| "S3 not available" | Check IAM role is attached to EC2 |
| Dashboard not loading | Check security group allows port 8501 |
| Sync failing | Ensure Edge SSO is working (may need initial manual login) |
| "Permission denied" on S3 | Verify IAM policy includes your bucket ARN |

---

## 📞 Support

- **Repository**: https://github.com/Prathima-create/ai-velocity-portal
- **Folder**: `qc-automation/`
- **Original script**: `scripts/qc_automation.py` (standalone version)
