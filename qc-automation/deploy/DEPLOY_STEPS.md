# 🚀 QC Dashboard Deployment — Step-by-Step Console Guide

## You have: Account `035268397124` with `AdministratorAccess` ✅

Follow these steps in order. Total time: ~20 minutes.

---

## Step 1: Create IAM User for QC Dashboard (5 min)

1. In AWS Console → **IAM** → **Users** → **Create user**
2. User name: `qc-dashboard-user`
3. Click **Next**
4. Select **Attach policies directly**
5. Search and attach these policies:
   - `AmazonS3FullAccess`
   - `AmazonEC2FullAccess`
6. Click **Next** → **Create user**
7. Click on the new user → **Security credentials** tab
8. Click **Create access key**
9. Select **Command Line Interface (CLI)**
10. Check the acknowledgment box → **Next** → **Create access key**
11. **COPY** the Access Key ID and Secret Access Key (save them!)

---

## Step 2: Add AWS Profile Locally

Open a terminal and run:
```cmd
"C:\Program Files\Amazon\AWSCLIV2\aws.exe" configure --profile qc-dashboard
```

Enter:
- Access Key ID: (paste from step 1)
- Secret Access Key: (paste from step 1)
- Region: `us-east-1`
- Output: `json`

---

## Step 3: Create S3 Bucket (2 min)

In AWS Console → **S3** → **Create bucket**:
- Bucket name: `fincom-qc-dashboard-035268397124` (must be globally unique)
- Region: US East (N. Virginia)
- ✅ Block all public access (keep checked!)
- ✅ Enable bucket versioning
- Encryption: SSE-S3 (AES-256)
- Click **Create bucket**

Then create a folder inside:
- Click on the bucket → **Create folder** → Name: `current` → Create

---

## Step 4: Launch EC2 Instance (10 min)

In AWS Console → **EC2** → **Launch Instance**:

### 4a. Basic Settings:
- **Name**: `qc-dashboard-server`
- **AMI**: Amazon Linux 2023 (free tier eligible)
- **Instance type**: `t2.micro` (free tier) or `t3.micro`
- **Key pair**: Create new → Name: `qc-dashboard-key` → RSA → .pem → Download it!

### 4b. Network Settings:
- Click **Edit** on Network settings
- **Auto-assign public IP**: Enable
- **Security group**: Create new:
  - Name: `qc-dashboard-sg`
  - Rule 1: SSH (port 22) → My IP
  - Rule 2: Custom TCP → Port `8501` → Source: `0.0.0.0/0` (or restrict to VPN: `10.0.0.0/8`)

### 4c. Advanced Details (scroll down):
- **IAM instance profile**: Create one if needed (see below) OR skip and use access keys

### 4d. Storage:
- 10 GiB gp3 (default is fine)

Click **Launch Instance**!

---

## Step 5: Connect to EC2 & Deploy (5 min)

### Option A: Session Manager (since you have it)
In AWS Console → EC2 → select your instance → **Connect** → **Session Manager** → Connect

### Option B: SSH
```bash
ssh -i "qc-dashboard-key.pem" ec2-user@<public-ip>
```

### Once connected, run these commands:

```bash
# Install dependencies
sudo yum update -y
sudo yum install -y python3 python3-pip git

# Clone your repo
git clone https://github.com/Prathima-create/ai-velocity-portal.git /home/ec2-user/qc-dashboard
cd /home/ec2-user/qc-dashboard/qc-automation

# Install Python packages
pip3 install --user -r requirements.txt

# Configure AWS (for S3 access)
aws configure
# Enter your qc-dashboard-user credentials
# Region: us-east-1
# Output: json

# Create data directory
mkdir -p data

# Test the app
streamlit run app.py --server.port=8501 --server.address=0.0.0.0 --server.headless=true &

# Check it's running
curl http://localhost:8501
```

---

## Step 6: Make it Permanent (auto-start)

```bash
# Create systemd service
sudo tee /etc/systemd/system/qc-dashboard.service << 'EOF'
[Unit]
Description=FinCom QC Dashboard (Streamlit)
After=network.target

[Service]
Type=simple
User=ec2-user
WorkingDirectory=/home/ec2-user/qc-dashboard/qc-automation
ExecStart=/home/ec2-user/.local/bin/streamlit run app.py --server.port=8501 --server.address=0.0.0.0 --server.headless=true
Restart=always
RestartSec=5
Environment=HOME=/home/ec2-user

[Install]
WantedBy=multi-user.target
EOF

sudo systemctl daemon-reload
sudo systemctl enable qc-dashboard
sudo systemctl start qc-dashboard

# Verify
sudo systemctl status qc-dashboard
```

---

## Step 7: Get Your Shareable Link! 🎉

Go back to EC2 Console → select your instance → copy the **Public IPv4 address**

Your dashboard is at:
```
http://<public-ip>:8501
```

Share this link with anyone! They can view the QC dashboard.

---

## Step 8: (Optional) Set Up Auto-Sync from SharePoint

For now, manually upload CSVs to S3:
```bash
# From your local machine
"C:\Program Files\Amazon\AWSCLIV2\aws.exe" s3 cp Fincom_Process.csv s3://fincom-qc-dashboard-035268397124/current/ --profile qc-dashboard
"C:\Program Files\Amazon\AWSCLIV2\aws.exe" s3 cp Fincom_Analyst.csv s3://fincom-qc-dashboard-035268397124/current/ --profile qc-dashboard
```

Later, set up the auto-sync timer on EC2 (see README.md).

---

## 🎯 Summary

| What | Value |
|------|-------|
| Account | 035268397124 |
| IAM User | qc-dashboard-user |
| S3 Bucket | fincom-qc-dashboard-035268397124 |
| EC2 Instance | qc-dashboard-server |
| Dashboard URL | http://<public-ip>:8501 |
| Security | Port 8501 open, SSH from your IP only |

---

## ⚠️ Won't Affect Project Athena

- Project Athena uses: `bedrock-callaudit-user-athena` (separate user, separate credentials)
- QC Dashboard uses: `qc-dashboard-user` (new user, isolated permissions)
- They share the same ACCOUNT but are completely independent
- Nothing in this setup touches Bedrock or Athena's resources
