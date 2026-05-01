"""
SharePoint API Access Test Script
Run this to check if you can connect to your SharePoint list.

Usage:
  Option 1 (Azure AD App): python scripts/test_sharepoint_access.py --mode app
  Option 2 (Browser Token): python scripts/test_sharepoint_access.py --mode token --token YOUR_BEARER_TOKEN
  Option 3 (Check REST API): python scripts/test_sharepoint_access.py --mode rest
"""

import argparse
import json
import sys
import os

try:
    import requests
except ImportError:
    print("❌ 'requests' package not installed. Run: pip install requests")
    sys.exit(1)

SITE_URL = "https://amazon.sharepoint.com/sites/AI-Velocity-site"
LIST_NAME = "AI Velocity Submission Portal"


def test_with_app_credentials():
    """Test using Azure AD App Registration (client credentials flow)"""
    print("\n🔑 Testing with Azure AD App Credentials...")
    print("=" * 60)
    
    tenant_id = input("Enter Tenant ID: ").strip()
    client_id = input("Enter Client ID: ").strip()
    client_secret = input("Enter Client Secret: ").strip()
    
    if not all([tenant_id, client_id, client_secret]):
        print("❌ All three values are required!")
        return False
    
    # Step 1: Get access token
    print("\n📡 Step 1: Getting access token from Azure AD...")
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    token_data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default"
    }
    
    try:
        resp = requests.post(token_url, data=token_data, timeout=15)
        if resp.status_code == 200:
            token = resp.json().get("access_token")
            print("✅ Got access token!")
        else:
            print(f"❌ Failed to get token: {resp.status_code}")
            print(f"   Error: {resp.json().get('error_description', resp.text[:200])}")
            return False
    except Exception as e:
        print(f"❌ Connection error: {e}")
        return False
    
    # Step 2: Get SharePoint site info via Graph API
    print("\n📡 Step 2: Getting SharePoint site info via Microsoft Graph...")
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    
    # Extract site path from URL
    site_path = "/sites/AI-Velocity-site"
    graph_url = f"https://graph.microsoft.com/v1.0/sites/amazon.sharepoint.com:{site_path}"
    
    try:
        resp = requests.get(graph_url, headers=headers, timeout=15)
        if resp.status_code == 200:
            site = resp.json()
            site_id = site.get("id")
            print(f"✅ Site found: {site.get('displayName')}")
            print(f"   Site ID: {site_id}")
        else:
            print(f"❌ Cannot access site: {resp.status_code}")
            print(f"   Error: {resp.text[:300]}")
            if resp.status_code == 403:
                print("\n⚠️  You need 'Sites.Read.All' permission. Ask your admin to grant it.")
            return False
    except Exception as e:
        print(f"❌ Error: {e}")
        return False
    
    # Step 3: Get list items
    print(f"\n📡 Step 3: Getting list '{LIST_NAME}'...")
    lists_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists"
    
    try:
        resp = requests.get(lists_url, headers=headers, timeout=15)
        if resp.status_code == 200:
            lists = resp.json().get("value", [])
            target_list = None
            for lst in lists:
                if lst.get("displayName") == LIST_NAME:
                    target_list = lst
                    break
            
            if target_list:
                list_id = target_list.get("id")
                print(f"✅ List found: {target_list.get('displayName')}")
                print(f"   List ID: {list_id}")
                
                # Step 4: Get items count
                items_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items?$top=5&$expand=fields"
                resp = requests.get(items_url, headers=headers, timeout=15)
                if resp.status_code == 200:
                    items = resp.json().get("value", [])
                    print(f"\n✅ SUCCESS! Can read list items. Got {len(items)} sample items.")
                    if items:
                        fields = items[0].get("fields", {})
                        print(f"\n📋 Available columns ({len(fields)} fields):")
                        for key in sorted(fields.keys())[:20]:
                            print(f"   - {key}: {str(fields[key])[:60]}")
                        if len(fields) > 20:
                            print(f"   ... and {len(fields) - 20} more")
                    
                    # Save credentials for later use
                    save_env(tenant_id, client_id, client_secret, site_id, list_id)
                    return True
                else:
                    print(f"❌ Cannot read items: {resp.status_code}")
                    print(f"   {resp.text[:200]}")
            else:
                print(f"❌ List '{LIST_NAME}' not found. Available lists:")
                for lst in lists[:10]:
                    print(f"   - {lst.get('displayName')}")
        else:
            print(f"❌ Cannot list site lists: {resp.status_code}")
            print(f"   {resp.text[:200]}")
    except Exception as e:
        print(f"❌ Error: {e}")
    
    return False


def test_with_browser_token(token):
    """Test using a bearer token copied from browser DevTools"""
    print("\n🔑 Testing with Browser Bearer Token...")
    print("=" * 60)
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json;odata=verbose"
    }
    
    # Test SharePoint REST API
    print(f"\n📡 Testing SharePoint REST API at {SITE_URL}...")
    api_url = f"{SITE_URL}/_api/web"
    
    try:
        resp = requests.get(api_url, headers=headers, timeout=15)
        if resp.status_code == 200:
            data = resp.json()
            title = data.get("d", {}).get("Title", "Unknown")
            print(f"✅ Connected to site: {title}")
            
            # Try to get list data
            list_url = f"{SITE_URL}/_api/web/lists/getbytitle('{LIST_NAME}')/items?$top=5"
            resp2 = requests.get(list_url, headers=headers, timeout=15)
            if resp2.status_code == 200:
                items = resp2.json().get("d", {}).get("results", [])
                print(f"✅ SUCCESS! Can read list items. Got {len(items)} sample items.")
                if items:
                    print(f"\n📋 Sample fields:")
                    for key in sorted(items[0].keys())[:15]:
                        print(f"   - {key}: {str(items[0][key])[:60]}")
                return True
            else:
                print(f"❌ Cannot read list: {resp2.status_code}")
        else:
            print(f"❌ Cannot access site: {resp.status_code}")
            if resp.status_code == 401:
                print("   Token may have expired. Get a fresh one from browser DevTools.")
            elif resp.status_code == 403:
                print("   You don't have permission to access this site via API.")
    except Exception as e:
        print(f"❌ Error: {e}")
    
    return False


def test_rest_api():
    """Test if SharePoint REST API is reachable (no auth)"""
    print("\n🌐 Testing SharePoint REST API Reachability...")
    print("=" * 60)
    
    try:
        resp = requests.get(SITE_URL, timeout=10, allow_redirects=False)
        print(f"📡 Status: {resp.status_code}")
        
        if resp.status_code in (200, 302, 301):
            print("✅ SharePoint site is reachable!")
            if resp.status_code in (302, 301):
                redirect = resp.headers.get("Location", "")
                print(f"   Redirects to: {redirect[:100]}")
                if "login" in redirect.lower() or "adfs" in redirect.lower():
                    print("   → Uses SSO/ADFS authentication")
                    print("\n💡 Recommendation: Use Azure AD App credentials approach")
                    print("   OR use Selenium for browser-based SSO login")
        else:
            print(f"⚠️  Unexpected status: {resp.status_code}")
    except requests.exceptions.SSLError:
        print("⚠️  SSL error — you may be behind a corporate proxy/firewall")
        print("   Try running from your corporate VPN")
    except requests.exceptions.ConnectionError:
        print("❌ Cannot reach SharePoint. Check your network/VPN connection.")
    except Exception as e:
        print(f"❌ Error: {e}")
    
    print("\n" + "=" * 60)
    print("📋 NEXT STEPS:")
    print("=" * 60)
    print("""
1. If you have Azure AD access (portal.azure.com):
   → Run: python scripts/test_sharepoint_access.py --mode app
   → You'll need: tenant_id, client_id, client_secret

2. If you can get a token from browser DevTools:
   → Open SharePoint in browser → F12 → Network → copy Bearer token
   → Run: python scripts/test_sharepoint_access.py --mode token --token YOUR_TOKEN

3. If neither works (common in corporate environments):
   → Use Selenium approach: I'll build a browser automation script
   → Or use Power Automate to export CSV on schedule
   → Or manually drop CSV → server auto-reloads

4. Check if you have Power Automate:
   → Go to: https://make.powerautomate.com
   → If accessible, this is the easiest no-code approach
""")


def save_env(tenant_id, client_id, client_secret, site_id, list_id):
    """Save credentials to .env file"""
    env_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), ".env")
    with open(env_path, "w") as f:
        f.write(f"SHAREPOINT_TENANT_ID={tenant_id}\n")
        f.write(f"SHAREPOINT_CLIENT_ID={client_id}\n")
        f.write(f"SHAREPOINT_CLIENT_SECRET={client_secret}\n")
        f.write(f"SHAREPOINT_SITE_ID={site_id}\n")
        f.write(f"SHAREPOINT_LIST_ID={list_id}\n")
        f.write(f"SHAREPOINT_SITE_URL={SITE_URL}\n")
        f.write(f"SHAREPOINT_LIST_NAME={LIST_NAME}\n")
        f.write(f"SYNC_INTERVAL_MINUTES=30\n")
    print(f"\n💾 Credentials saved to {env_path}")
    print("   ⚠️  Add .env to .gitignore!")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Test SharePoint API Access")
    parser.add_argument("--mode", choices=["app", "token", "rest"], default="rest",
                       help="Test mode: 'app' for Azure AD, 'token' for browser token, 'rest' for basic check")
    parser.add_argument("--token", help="Bearer token from browser DevTools", default="")
    args = parser.parse_args()
    
    print("🚀 SharePoint API Access Tester for AI Velocity Portal")
    print("=" * 60)
    
    if args.mode == "app":
        success = test_with_app_credentials()
    elif args.mode == "token":
        if not args.token:
            args.token = input("Enter Bearer token: ").strip()
        success = test_with_browser_token(args.token)
    else:
        test_rest_api()
        success = False
    
    if success:
        print("\n" + "=" * 60)
        print("🎉 SUCCESS! SharePoint API access is working!")
        print("=" * 60)
        print("Next: I'll set up automatic sync in the backend.")
    else:
        print("\n💡 Run with --mode rest first to check connectivity:")
        print("   python scripts/test_sharepoint_access.py --mode rest")
