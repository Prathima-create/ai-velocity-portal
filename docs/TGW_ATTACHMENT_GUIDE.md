# Transit Gateway (TGW) Attachment Guide for Conduit Account

## Your Details
- **Account**: 035268397124 (Conduit)
- **VPC**: vpc-0e6eaf3dac4d50a42
- **Region**: us-east-1
- **Goal**: VPN users → Transit Gateway → Internal ALB (Athena AuditIQ)

---

## Overview: Why You Need TGW

Your internal ALB is in a private VPC in a conduit account. For VPN-connected users (on Amazon corporate network) to reach it, the VPC needs to be attached to the **Amazon Corporate Transit Gateway**. This creates a route from the corporate network → TGW → your VPC → internal ALB.

---

## Step-by-Step Process

### Step 1: Verify Your VPC Subnet Configuration

Before requesting TGW attachment, ensure:
```bash
# Check your VPC CIDR (must not overlap with Amazon corporate ranges)
aws ec2 describe-vpcs --vpc-ids vpc-0e6eaf3dac4d50a42 --query "Vpcs[0].CidrBlockAssociationSet[*].CidrBlock" --region us-east-1 --profile auditiq

# Check subnets - you need at least 1 subnet in at least 2 AZs for TGW attachment
aws ec2 describe-subnets --filters "Name=vpc-id,Values=vpc-0e6eaf3dac4d50a42" --query "Subnets[*].[SubnetId,AvailabilityZone,CidrBlock,Tags[?Key=='Name'].Value|[0]]" --output table --region us-east-1 --profile auditiq
```

**Important**: Create dedicated TGW subnets (small /28 or /27) in at least 2 AZs. Don't use your application subnets for TGW attachment.

### Step 2: Request Transit Gateway Attachment

There are multiple ways to request this:

#### Option A: Self-Service via Conduit Console (Preferred)
1. Go to **Conduit Console** → Your account (035268397124)
2. Navigate to **Networking** → **Transit Gateway Attachments**
3. If available, click **Request TGW Attachment**
4. Select your VPC and the subnets for attachment
5. This auto-creates the SIM ticket

#### Option B: SIM Ticket (Manual)
1. Go to: **https://sim.amazon.com/issues/create**
2. Use template: **"Transit Gateway VPC Attachment Request"**
   - Search for: `AWS/Networking/TransitGateway` category
   - Or search: `EC2 VPC Transit Gateway Attachment`
3. Fill in:
   - **Account ID**: 035268397124
   - **Account Type**: Conduit
   - **VPC ID**: vpc-0e6eaf3dac4d50a42
   - **Region**: us-east-1
   - **Subnets for TGW attachment**: (list your TGW subnet IDs)
   - **Business justification**: "Internal tool (Project Athena - Call Audit Automation) hosted on EC2 behind internal ALB. Need corporate VPN users to access the app via internal ALB DNS."

#### Option C: Via Network Configuration Service (NCS)
1. Go to: **https://ncs.amazon.com** (Network Configuration Service)
2. Search for your account/VPC
3. Request TGW attachment through the portal

### Step 3: Create TGW Subnets (if not existing)
```bash
# Create dedicated TGW subnets in 2 AZs (use small /28 CIDRs)
# Adjust CIDR based on your VPC's available space
aws ec2 create-subnet \
  --vpc-id vpc-0e6eaf3dac4d50a42 \
  --cidr-block 10.x.x.0/28 \
  --availability-zone us-east-1a \
  --tag-specifications 'ResourceType=subnet,Tags=[{Key=Name,Value=athena-tgw-subnet-1a}]' \
  --region us-east-1 --profile auditiq

aws ec2 create-subnet \
  --vpc-id vpc-0e6eaf3dac4d50a42 \
  --cidr-block 10.x.x.16/28 \
  --availability-zone us-east-1b \
  --tag-specifications 'ResourceType=subnet,Tags=[{Key=Name,Value=athena-tgw-subnet-1b}]' \
  --region us-east-1 --profile auditiq
```

### Step 4: After TGW Attachment is Approved

Once the attachment is created:

1. **Update Route Tables** — Add route to corporate network via TGW:
```bash
# Get your route table ID
aws ec2 describe-route-tables --filters "Name=vpc-id,Values=vpc-0e6eaf3dac4d50a42" --query "RouteTables[*].[RouteTableId,Tags[?Key=='Name'].Value|[0]]" --output table --region us-east-1 --profile auditiq

# Add route for Amazon corporate network via TGW
# (TGW ID will be provided after attachment approval)
aws ec2 create-route \
  --route-table-id rtb-XXXXX \
  --destination-cidr-block 10.0.0.0/8 \
  --transit-gateway-id tgw-XXXXX \
  --region us-east-1 --profile auditiq
```

2. **Security Group** — Ensure your ALB SG allows traffic from corporate VPN CIDR:
```bash
# Check current ALB security group rules
aws ec2 describe-security-groups --filters "Name=group-name,Values=*athena*" --query "SecurityGroups[*].[GroupId,GroupName,IpPermissions]" --region us-east-1 --profile auditiq

# If needed, add inbound rule for corporate VPN range
aws ec2 authorize-security-group-ingress \
  --group-id sg-XXXXX \
  --protocol tcp \
  --port 443 \
  --cidr 10.0.0.0/8 \
  --region us-east-1 --profile auditiq
```

3. **DNS** — Register a friendly CNAME:
   - Request via SIM: DNS CNAME pointing to your internal ALB DNS
   - e.g., `athena-auditiq.corp.amazon.com` → ALB DNS name

---

## AppSec / LOAF / GenAI Security Controls

Per the links you provided:

### AppSTAR LOAF (Line of Assurance Framework)
- URL: `https://w.amazon.com/bin/view/InfoSec/Application_Security/AppSTAR/LOAF/Tools/AppsecMCP`
- LOAF provides security tooling for application security review
- For your ASR, you'll need to address LOAF findings

### GenAI Security Controls
- URL: `https://w.amazon.com/bin/view/InfoSec/Application_Security/AppSTAR/Appsec_AI/GenAI_Security_Controls/`
- Since Project Athena uses **Bedrock (Claude)** for AI-powered audit analysis, you MUST comply with GenAI security controls:
  1. **Data Classification**: Ensure call audit data is classified correctly
  2. **PII Handling**: We've already added REDACTED masking ✅
  3. **Model Access Controls**: Using IAM role with least privilege ✅
  4. **Prompt Injection Protection**: Review audit_engine.py prompts
  5. **Output Validation**: Ensure AI outputs are validated before storage
  6. **Logging**: CloudWatch logging enabled ✅ with PII scrubbing ✅

### For ASR Review Readiness:
1. Complete LOAF assessment via the AppsecMCP tool
2. Document GenAI security controls in your ASR
3. Ensure TGW attachment resolves the "no network access" BPA finding
4. Update DI diagram to show: VPN User → TGW → VPC → Internal ALB → EC2

---

## Alternative: If TGW Takes Too Long

If TGW attachment is slow (can take days), consider these alternatives:

### Option 1: SSM Port Forwarding (Already Working!)
```bash
aws ssm start-session --target i-0dfb47b248fc6f8e8 \
  --document-name AWS-StartPortForwardingSessionToRemoteHost \
  --parameters host="internal-athena-alb-XXXXX.us-east-1.elb.amazonaws.com",portNumber="80",localPortNumber="8501" \
  --region us-east-1 --profile auditiq
```
Then access via: http://localhost:8501

### Option 2: Client VPN Endpoint
Set up AWS Client VPN endpoint in your VPC (requires more setup but gives direct VPC access).

---

## Quick Action Items

| # | Action | How |
|---|--------|-----|
| 1 | Create TGW subnets | AWS CLI (see Step 3) |
| 2 | Request TGW attachment | SIM ticket or Conduit Console |
| 3 | Update route tables | After TGW approved (Step 4) |
| 4 | Update ALB security group | Allow 10.0.0.0/8 on port 443 |
| 5 | Complete LOAF assessment | AppsecMCP tool |
| 6 | Document GenAI controls | ASR document |
| 7 | Request DNS CNAME | SIM ticket |
