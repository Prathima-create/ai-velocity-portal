# 🚀 AI Velocity Portal - Hosting Guide

## Goal: Get a URL like `https://aivelocity.a2z.com` accessible to all Amazon employees

---

## 🏆 Option 1: AWS Amplify Hosting (RECOMMENDED - Easiest)

**URL format**: `https://aivelocity.amplifyapp.com` or custom `https://aivelocity.a2z.com`  
**Time to deploy**: ~15 minutes  
**Cost**: Free tier covers this easily  

### Steps:

1. **Get an AWS Account** (via Isengard)
   - Go to https://isengard.amazon.com
   - Create a new account or use your team's existing one
   - Select your org and create an account (e.g., "AI-Velocity-Portal")

2. **Push code to AWS CodeCommit** (or GitHub Enterprise)
   ```bash
   # Initialize git repo
   cd c:\Users\pratpk\ai-velocity-portal
   git init
   git add .
   git commit -m "AI Velocity Portal v3.0"
   
   # Push to CodeCommit (create repo first in AWS Console)
   git remote add origin https://git-codecommit.us-east-1.amazonaws.com/v1/repos/ai-velocity-portal
   git push -u origin main
   ```

3. **Deploy via Amplify Console**
   - Go to AWS Console → Amplify
   - Click "New app" → "Host web app"
   - Connect your CodeCommit repo
   - It will detect the Dockerfile and deploy automatically
   - Get URL: `https://main.xxxxx.amplifyapp.com`

4. **Custom Domain** (to get `aivelocity.a2z.com`)
   - In Amplify Console → Domain Management → Add domain
   - Add `aivelocity.a2z.com`
   - Follow DNS setup instructions

---

## 🏆 Option 2: EC2 + ALB (Most Control)

**URL format**: `https://aivelocity.a2z.com`  
**Time to deploy**: ~30 minutes  
**Best for**: Full control, persistent data, auto-sync  

### Steps:

1. **Launch EC2 Instance** (via Isengard account)
   ```
   - AMI: Amazon Linux 2023
   - Instance type: t3.micro (free tier)
   - Security Group: Allow ports 80, 443 from Amazon VPN (10.0.0.0/8)
   ```

2. **SSH in and deploy**
   ```bash
   # Install dependencies
   sudo yum install -y python3 python3-pip git
   
   # Clone your code
   git clone <your-codecommit-repo-url>
   cd ai-velocity-portal
   
   # Install Python deps
   pip3 install -r backend/requirements.txt
   
   # Run with auto-restart
   nohup python3 backend/main.py &
   ```

3. **Set up ALB + HTTPS**
   - Create an Application Load Balancer
   - Request ACM certificate for `aivelocity.a2z.com`
   - Point ALB to EC2 target group (port 3000)
   - Update DNS to point `aivelocity.a2z.com` → ALB

4. **Auto-start on reboot** (create systemd service)
   ```bash
   sudo tee /etc/systemd/system/aivelocity.service << 'EOF'
   [Unit]
   Description=AI Velocity Portal
   After=network.target
   
   [Service]
   Type=simple
   User=ec2-user
   WorkingDirectory=/home/ec2-user/ai-velocity-portal
   ExecStart=/usr/bin/python3 backend/main.py
   Restart=always
   
   [Install]
   WantedBy=multi-user.target
   EOF
   
   sudo systemctl enable aivelocity
   sudo systemctl start aivelocity
   ```

---

## 🏆 Option 3: ECS Fargate (Serverless Container - Auto-scaling)

**URL format**: `https://aivelocity.a2z.com`  
**Time to deploy**: ~45 minutes  
**Best for**: Zero maintenance, auto-scaling  

### Steps:

1. **Build & Push Docker Image**
   ```bash
   # Login to ECR
   aws ecr get-login-password | docker login --username AWS --password-stdin <account-id>.dkr.ecr.us-east-1.amazonaws.com
   
   # Build and push
   docker build -t ai-velocity-portal .
   docker tag ai-velocity-portal:latest <account-id>.dkr.ecr.us-east-1.amazonaws.com/ai-velocity-portal:latest
   docker push <account-id>.dkr.ecr.us-east-1.amazonaws.com/ai-velocity-portal:latest
   ```

2. **Create ECS Service**
   - Create ECS Cluster (Fargate)
   - Create Task Definition pointing to your ECR image
   - Create Service with ALB
   - Map port 3000

3. **Custom Domain**
   - Same as Option 2: ACM cert + Route53 → ALB

---

## 🏆 Option 4: S3 + CloudFront (Static Frontend Only)

**URL format**: `https://aivelocity.cloudfront.net` or custom domain  
**Limitation**: Need to convert to static site (pre-render data into JSON files)  
**Not recommended** for this app since it needs a backend to read CSV dynamically.

---

## 🔧 Getting a Custom Domain (`aivelocity.a2z.com`)

For any option above, to get a custom Amazon domain:

1. **File a DNS ticket** via your team's IT contact or:
   - Go to https://dns.amazon.com (internal)
   - Request a CNAME record: `aivelocity.a2z.com` → your ALB/Amplify/CloudFront URL

2. **SSL Certificate**:
   - Use AWS Certificate Manager (ACM) to request a cert for `aivelocity.a2z.com`
   - Validate via DNS (add the CNAME record ACM gives you)

---

## 📊 Keeping Data Fresh

### Option A: Manual CSV Upload
- SSH into EC2 and replace `data/submissions.csv`
- Or set up a simple upload endpoint

### Option B: Scheduled Sync (Recommended)
Add a cron job on EC2 to auto-sync from SharePoint:
```bash
# Every 30 minutes, sync from SharePoint
*/30 * * * * cd /home/ec2-user/ai-velocity-portal && python3 scripts/sync_sharepoint.py >> /var/log/aivelocity-sync.log 2>&1
```

### Option C: API Integration
If you get SharePoint API access (see `docs/SHAREPOINT_API_GUIDE.md`), the backend can pull data directly from SharePoint — no CSV needed.

---

## 🎯 My Recommendation

**For quick launch**: **Option 1 (Amplify)** — 15 minutes, zero maintenance  
**For long-term with auto-sync**: **Option 2 (EC2)** — more control, can run Selenium sync  
**For enterprise-grade**: **Option 3 (ECS Fargate)** — auto-scaling, zero patching  

All options work with Midway authentication if you add the Midway ALB integration.

---

## 🔐 Adding Amazon Midway Authentication

To require Amazon SSO login (so only Amazon employees can access):

1. Use an ALB with **OIDC authentication** action
2. Configure Midway as the identity provider
3. ALB will handle login — your app doesn't need any auth code
4. Every request gets `x-amzn-oidc-identity` header with the user's alias

This ensures only authenticated Amazon employees can access `https://aivelocity.a2z.com`.
