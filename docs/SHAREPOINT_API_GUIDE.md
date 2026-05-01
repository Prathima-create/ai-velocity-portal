# SharePoint API Access Guide for AI Velocity Portal

## What You Need

To pull data directly from your SharePoint list via API, you need **3 things**:

### 1. Azure AD App Registration (Admin Required)
You need an **Azure AD App** registered in your Amazon/corporate tenant. This gives you:
- **Tenant ID** — Your org's Azure AD tenant ID
- **Client ID** — The app's unique identifier
- **Client Secret** — A secret key for authentication

### 2. SharePoint Site & List Details
From your SharePoint URL:
```
https://amazon.sharepoint.com/sites/AI-Velocity-site/Lists/AI%20Velocity%20Submission%20Portal/AllItems.aspx
```
- **Site URL**: `https://amazon.sharepoint.com/sites/AI-Velocity-site`
- **List Name**: `AI Velocity Submission Portal`

### 3. API Permissions Required
The Azure AD app needs these Microsoft Graph permissions:
- `Sites.Read.All` — Read SharePoint sites
- `Lists.Read.All` — Read SharePoint lists (alternative: `Sites.ReadWrite.All`)

---

## Step-by-Step: Check If You Already Have Access

### Step 1: Check if you have Azure AD access
Go to: https://portal.azure.com → Azure Active Directory → App registrations

If you can see this page, you have access. If not, you'll need to request access from your IT admin.

### Step 2: Register a new App (if you have access)
1. Click **"New registration"**
2. Name: `AI Velocity Portal - SharePoint Sync`
3. Supported account types: **Single tenant**
4. Click **Register**
5. Copy the **Application (client) ID** and **Directory (tenant) ID**

### Step 3: Create a Client Secret
1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Description: `AI Velocity sync`
4. Expiry: 12 months
5. Copy the **Value** immediately (it won't be shown again)

### Step 4: Grant API Permissions
1. Go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Choose **Application permissions**
5. Search and add:
   - `Sites.Read.All`
6. Click **Grant admin consent** (requires admin)

### Step 5: Test with the script below
Run the test script I've created: `scripts/test_sharepoint_access.py`

---

## Alternative: SharePoint REST API (No Azure AD needed)
If you have a **SharePoint access token** from your browser session, you can use it directly:

1. Open SharePoint in your browser
2. Open Developer Tools (F12) → Network tab
3. Look for any request to `amazon.sharepoint.com`
4. Copy the `Authorization: Bearer <token>` header value
5. Use it in the test script

---

## Quick Test: Can You Access SharePoint API?

Run this in PowerShell to test basic access:
```powershell
# Replace with your actual token from browser DevTools
$token = "YOUR_BEARER_TOKEN"
$siteUrl = "https://amazon.sharepoint.com/sites/AI-Velocity-site"

$headers = @{
    "Authorization" = "Bearer $token"
    "Accept" = "application/json"
}

# Test: Get site info
Invoke-RestMethod -Uri "$siteUrl/_api/web" -Headers $headers -Method Get
```

If this returns JSON with site info, you have API access!

---

## Fallback Options (If No API Access)

### Option A: File Watcher (Easiest)
Drop the CSV export into `data/submissions.csv` and the server auto-reloads.

### Option B: Selenium Automation
Uses a headless browser to:
1. Open SharePoint
2. Login via SSO
3. Navigate to list → Export to CSV
4. Save to `data/submissions.csv`

### Option C: Power Automate Flow
If you have Power Automate access:
1. Create a flow triggered on schedule (every 30 min)
2. Action: "Get items" from SharePoint list
3. Action: "Create CSV" and save to OneDrive/local folder
4. The portal picks up the new CSV

---

## Environment Variables Needed (for API approach)

Create a `.env` file in the project root:
```env
SHAREPOINT_TENANT_ID=your-tenant-id
SHAREPOINT_CLIENT_ID=your-client-id
SHAREPOINT_CLIENT_SECRET=your-client-secret
SHAREPOINT_SITE_URL=https://amazon.sharepoint.com/sites/AI-Velocity-site
SHAREPOINT_LIST_NAME=AI Velocity Submission Portal
SYNC_INTERVAL_MINUTES=30
```
