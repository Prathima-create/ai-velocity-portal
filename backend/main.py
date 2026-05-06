"""
AI Velocity Portal - Backend API
Serves AI Wins Dashboard data from SharePoint CSV export
"""

from fastapi import FastAPI, Query, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import List, Optional, Dict, Any
import csv
import os
import sys
from datetime import datetime
from collections import Counter

app = FastAPI(
    title="AI Velocity Portal API",
    description="Backend API for AI Velocity - Accounts Payable",
    version="2.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─── Leader / Org Mapping ─────────────────────────────────────────────────────
# Process-name based leader mapping (default fallback)
LEADER_MAPPING = {
    "Corp Invoice Processing": {"leader": "Hari", "poc": "Bighnaraja"},
    "Critical Vendors - Invoice Processing": {"leader": "Hari", "poc": "Bighnaraja"},
    "Corp AP - FinCoM -SVOT": {"leader": "Hari", "poc": "Bighnaraja"},
    "Corp AP - FinCoM- Inbound": {"leader": "Hari", "poc": "Bighnaraja"},
    "FinCoM_Expense": {"leader": "Hari", "poc": "Bighnaraja"},
    "Expense": {"leader": "Hari", "poc": "Bighnaraja"},
    "Corp Cards": {"leader": "Hari", "poc": "Bighnaraja"},
    "TTT": {"leader": "Sanjeev", "poc": "Yashika"},
    "Accts Payable Retail- VAR": {"leader": "Ritesh", "poc": "Vomsee"},
    "Accts Payable - VAR": {"leader": "Ritesh", "poc": "Vomsee"},
    "Accts Payable NonInventory-VAR": {"leader": "Kevin", "poc": "Arunima"},
    "FinOps Projects": {"leader": "Kevin", "poc": "Arunima"},
    "Invoice on Hold": {"leader": "Kevin", "poc": "Arunima"},
    "FinOps - AR CDO - VAR": {"leader": "Leela", "poc": "Renuka"},
}

# Manager-name → L7 Leader mapping (takes PRIORITY over process-based mapping)
# This ensures correct org attribution regardless of which process the person selects
MANAGER_TO_LEADER = {
    # L7: Hari (hathamma) — POC: Bighnaraja Dash (bighnard)
    "THAMMALA, HARI":                       {"leader": "Hari", "poc": "Bighnaraja"},
    "V, UPENDRA":                           {"leader": "Hari", "poc": "Bighnaraja"},
    "Medakkar, Mayur":                      {"leader": "Hari", "poc": "Bighnaraja"},
    "Ch, Raghunadh":                        {"leader": "Hari", "poc": "Bighnaraja"},
    "SINGH, ROHIT":                         {"leader": "Hari", "poc": "Bighnaraja"},
    "Dash, Bighnaraja":                     {"leader": "Hari", "poc": "Bighnaraja"},
    "Manchili, Sree Somasekhar":            {"leader": "Hari", "poc": "Bighnaraja"},
    "Ali, Md Mujahed":                      {"leader": "Hari", "poc": "Bighnaraja"},
    "Janardhanan, Madhusudanan(MJ)":        {"leader": "Hari", "poc": "Bighnaraja"},
    "Srikanth, Kanumarlapudi":              {"leader": "Hari", "poc": "Bighnaraja"},
    "Mishra, Renuka":                       {"leader": "Leela", "poc": "Renuka"},
    "Tandon, Bhawana":                      {"leader": "Hari", "poc": "Bighnaraja"},
    "Sharma, Yogendra":                     {"leader": "Hari", "poc": "Bighnaraja"},
    "Khan, Akbar":                          {"leader": "Hari", "poc": "Bighnaraja"},
    "Adhikari, Neha":                       {"leader": "Hari", "poc": "Bighnaraja"},
    "Mahindrakar, Vishnu":                  {"leader": "Hari", "poc": "Bighnaraja"},
    "Devi, Renuka":                         {"leader": "Hari", "poc": "Bighnaraja"},
    "Khan, Shahroof":                       {"leader": "Hari", "poc": "Bighnaraja"},
    "Tangellamudi, Prasad":                 {"leader": "Hari", "poc": "Bighnaraja"},
    "Tangellamudi, Jayadev":                {"leader": "Hari", "poc": "Bighnaraja"},
    "Walia, Shivalika":                     {"leader": "Hari", "poc": "Bighnaraja"},
    "Baloch, Zeenat Hanif":                 {"leader": "Hari", "poc": "Bighnaraja"},
    "Chandrasekharan, Prashanth":           {"leader": "Hari", "poc": "Bighnaraja"},
    "Harivanam, Naga Sravan":               {"leader": "Hari", "poc": "Bighnaraja"},
    "Venkata Amrutha Sai, Pandeswara":      {"leader": "Hari", "poc": "Bighnaraja"},
    "Dachepalli, Venkatesh":                {"leader": "Hari", "poc": "Bighnaraja"},
    "Bera, Kumari Payal":                   {"leader": "Hari", "poc": "Bighnaraja"},
    "Sharma, Shil Nidhi":                   {"leader": "Hari", "poc": "Bighnaraja"},
    "Sharma, Niraj":                        {"leader": "Hari", "poc": "Bighnaraja"},
    "Palle, Suresh":                        {"leader": "Hari", "poc": "Bighnaraja"},
    "Nathari, Srilatha":                    {"leader": "Hari", "poc": "Bighnaraja"},
    "Bohra, Hamza":                         {"leader": "Hari", "poc": "Bighnaraja"},
    "Mohammed, Abdul Rahaman":              {"leader": "Hari", "poc": "Bighnaraja"},
    "Koduru, Harish":                       {"leader": "Hari", "poc": "Bighnaraja"},
    "Srigari, Sruthika":                    {"leader": "Hari", "poc": "Bighnaraja"},
    "Jashwanth, Sappidi":                   {"leader": "Hari", "poc": "Bighnaraja"},
    "Vishnubhatla, Satyanaga Vidya Sagar":  {"leader": "Hari", "poc": "Bighnaraja"},
    # L7: Kevin (fekevn) — POC: Arunima (HYD) / Komal (PNQ)
    "Fernandes, Kevin":                     {"leader": "Kevin", "poc": "Arunima"},
    "Gaddam, Shirish":                      {"leader": "Kevin", "poc": "Arunima"},
    "Paravastu, Samudrika":                 {"leader": "Kevin", "poc": "Arunima"},
    "Bob, Terence":                         {"leader": "Kevin", "poc": "Komal"},
    "Unnisa, Habeeb":                       {"leader": "Kevin", "poc": "Arunima"},
    "Khureja, Swati":                       {"leader": "Kevin", "poc": "Komal"},
    "Kunapareddy, Harika":                  {"leader": "Kevin", "poc": "Arunima"},
    "Mynampati, Sindhuri":                  {"leader": "Kevin", "poc": "Arunima"},
    "Sulthana, Faheem":                     {"leader": "Kevin", "poc": "Arunima"},
    "Kotikalapudi, Ravi Kiran":             {"leader": "Kevin", "poc": "Arunima"},
    "Mohammad, Rafiz":                      {"leader": "Kevin", "poc": "Arunima"},
    "Mukherjee, Arunima":                   {"leader": "Kevin", "poc": "Arunima"},
    "Jadhav, Komal":                        {"leader": "Kevin", "poc": "Komal"},
    "Pawar, Amit Subhash":                  {"leader": "Kevin", "poc": "Komal"},
    "Rodrigues, Clifford":                  {"leader": "Kevin", "poc": "Komal"},
    "D Sri, Lokesh":                        {"leader": "Kevin", "poc": "Arunima"},
    "Hemanth Kumar, G.":                    {"leader": "Kevin", "poc": "Arunima"},
    # L7: Parul (guppar) — POC: Naren Maheshwari (mahenare)
    "Dixit, Girish":                        {"leader": "Parul", "poc": "Naren"},
    "Madadi, Keshava Reddy":                {"leader": "Parul", "poc": "Naren"},
    "Vijayakumar, Priyadarsini":            {"leader": "Parul", "poc": "Naren"},
    "Maheshwari, Naren":                    {"leader": "Parul", "poc": "Naren"},
    "Prashanthi, S":                        {"leader": "Parul", "poc": "Naren"},
    "Sadhu, Sumanth":                       {"leader": "Parul", "poc": "Naren"},
    "Manchala, Sumalatha":                  {"leader": "Parul", "poc": "Naren"},
    "Singarapu, Nishanth":                  {"leader": "Parul", "poc": "Naren"},
    "Uttekar, Nilesh":                      {"leader": "Parul", "poc": "Naren"},
    "Ranjan, Saurabh":                      {"leader": "Parul", "poc": "Naren"},
    "Madhunantuni, Madhavi":                {"leader": "Parul", "poc": "Naren"},
    "Singh, Satish Bahadur":                {"leader": "Parul", "poc": "Naren"},
    "Sadhukhan, Tapas":                     {"leader": "Parul", "poc": "Naren"},
    "Singh, Anjesh":                        {"leader": "Parul", "poc": "Naren"},
    "Iyer, Pratik":                         {"leader": "Parul", "poc": "Naren"},
    "Dass, Christopher":                    {"leader": "Parul", "poc": "Naren"},
    "Matana, Priya":                        {"leader": "Parul", "poc": "Naren"},
    "Barretto, Rachel":                     {"leader": "Parul", "poc": "Naren"},
    # L7: Ritchie (rajivkmr) — POC: Harshita Arora (haaroraz)
    "., Ritchie":                           {"leader": "Ritchie", "poc": "Harshita"},
    "Thilak Kumar, Paavai":                 {"leader": "Ritchie", "poc": "Harshita"},
    "Kumar B, Harish":                      {"leader": "Ritchie", "poc": "Harshita"},
    "Raghavendran, Hareesh":                {"leader": "Ritchie", "poc": "Harshita"},
    "BL, Padmavathi":                       {"leader": "Ritchie", "poc": "Harshita"},
    "Govindarajan, Devaraj":                {"leader": "Ritchie", "poc": "Harshita"},
    "Suddoju, Ranjith":                     {"leader": "Ritchie", "poc": "Harshita"},
    "Bandukwala, Zainab Abbas":             {"leader": "Ritchie", "poc": "Harshita"},
    "Hussain, Shaik Fayaz":                 {"leader": "Ritchie", "poc": "Harshita"},
    "Pudugopuram, Mahesh Babu babu":        {"leader": "Ritchie", "poc": "Harshita"},
    "Kumar Koduri, Ajay":                   {"leader": "Ritchie", "poc": "Harshita"},
    "G, Satish":                            {"leader": "Ritchie", "poc": "Harshita"},
    "Srinivasan, Ramya":                    {"leader": "Ritchie", "poc": "Harshita"},
    "Arora, Harshita":                      {"leader": "Ritchie", "poc": "Harshita"},
    # L7: Ritesh (rjoshi) — POC: Vomsee Gowtham (vomseey)
    "Joshi, Ritesh":                        {"leader": "Ritesh", "poc": "Vomsee"},
    "Yerubandi, Vomsee Gowtham":            {"leader": "Ritesh", "poc": "Vomsee"},
    "Arockiamary, Cindrella":               {"leader": "Ritesh", "poc": "Vomsee"},
    # L7: Leela — POC: Renuka Mishra (renmish)
    "Gera, Leela":                          {"leader": "Leela", "poc": "Renuka"},
    "Natin, Arshad":                        {"leader": "Leela", "poc": "Renuka"},
    "Koripalli, Anvesh":                    {"leader": "Leela", "poc": "Renuka"},
    "Shrivastav, Rajesh":                   {"leader": "Leela", "poc": "Renuka"},
    "Yadav, Sarika":                        {"leader": "Leela", "poc": "Renuka"},
    "Kumar, Ashok":                         {"leader": "Leela", "poc": "Renuka"},
    # L7: Sanjeev (skmittal) — POC: Yashika Verma (yasver)
    "Mittal, Sanjeev":                      {"leader": "Sanjeev", "poc": "Yashika"},
    "Begari, Praveen Kumar":                {"leader": "Sanjeev", "poc": "Yashika"},
    "Macherla, Archana":                    {"leader": "Sanjeev", "poc": "Yashika"},
    "Korumilli, Raj":                       {"leader": "Sanjeev", "poc": "Yashika"},
    "Verma, Yashika":                       {"leader": "Sanjeev", "poc": "Yashika"},
    "Devulapalli, Deepthi":                 {"leader": "Sanjeev", "poc": "Yashika"},
    "Ravikanti, Divya":                     {"leader": "Sanjeev", "poc": "Yashika"},
    "Duttagupta, Sudip":                    {"leader": "Sanjeev", "poc": "Yashika"},
    "Dhiman, Navnoor":                      {"leader": "Sanjeev", "poc": "Yashika"},
    # Program Lead (under Hari's org)
    "K, Prathima":                          {"leader": "Hari", "poc": "Bighnaraja"},
}

# ─── SDE / Tech Team Contact Mapping ─────────────────────────────────────────
SDE_CONTACTS = {
    "Corp Invoice Processing": {"alias": "kartheek", "name": "Yarram Kartheek Reddy", "team": "ACES BD"},
    "Critical Vendors - Invoice Processing": {"alias": "mjanarad", "name": "Madhusudanan MJ", "team": "ACES BD"},
    "Corp AP - FinCoM -SVOT": {"alias": "raghunadh", "name": "Raghunadh Ch", "team": "ACES BD"},
    "Corp AP - FinCoM- Inbound": {"alias": "raghunadh", "name": "Raghunadh Ch", "team": "ACES BD"},
    "FinCoM_Expense": {"alias": "somasekh", "name": "Sree Somasekhar Manchili", "team": "ACES BD"},
    "Expense": {"alias": "somasekh", "name": "Sree Somasekhar Manchili", "team": "ACES BD"},
    "Corp Cards": {"alias": "bighnar", "name": "Bighnaraja Dash", "team": "ACES BD"},
    "TTT": {"alias": "kanumark", "name": "Kanumarlapudi Srikanth", "team": "ACES BD"},
    "Accts Payable Retail- VAR": {"alias": "bighnar", "name": "Bighnaraja Dash", "team": "ACES BD"},
    "Accts Payable - VAR": {"alias": "bighnar", "name": "Bighnaraja Dash", "team": "ACES BD"},
    "Accts Payable NonInventory-VAR": {"alias": "bighnar", "name": "Bighnaraja Dash", "team": "ACES BD"},
    "FinOps Projects": {"alias": "adarshsr", "name": "Adarsh Srivastav", "team": "SDE"},
    "Invoice on Hold": {"alias": "habeebu", "name": "Habeeb Unnisa", "team": "ACES BD"},
    "FinOps - AR CDO - VAR": {"alias": "bighnar", "name": "Bighnaraja Dash", "team": "ACES BD"},
}

# ─── AI Tool Suggestions based on problem type ───────────────────────────────
TOOL_SUGGESTIONS = {
    "document": ["Amazon Textract", "Amazon Comprehend", "Orcha AI", "Party Rock"],
    "extraction": ["Amazon Textract", "Party Rock", "Python + PDF Libraries", "Orcha AI"],
    "chatbot": ["Amazon Q Business", "Amazon Quick Suite", "Amazon Lex"],
    "knowledge": ["Amazon Q Business", "Amazon Quick Suite Chat Agent", "RAG Framework"],
    "automation": ["Amazon Quick Suite Flow", "Python Automation", "AWS Step Functions"],
    "analytics": ["Amazon QuickSight", "Amazon Q Business", "Python + Pandas"],
    "sentiment": ["Amazon Comprehend", "Custom LLM Agent", "Amazon Quick Suite"],
    "translation": ["Amazon Translate", "Amazon Comprehend", "Custom AI Agent"],
    "audit": ["Amazon Quick Suite", "Python Automation", "Custom AI Agent"],
    "email": ["Amazon Quick Suite Flow", "Python + Outlook Integration", "AWS SES"],
    "ocr": ["Amazon Textract", "Orcha AI", "Python + Tesseract"],
    "classification": ["Amazon Comprehend", "Custom ML Model", "Amazon SageMaker"],
    "dashboard": ["Amazon QuickSight", "GDA Dashboard", "Custom Web Dashboard"],
    "sop": ["Amazon Q Business", "Quick Suite Chat Agent", "RAG + Knowledge Base"],
    "invoice": ["Amazon Textract", "Orcha AI", "Python + PDF Extraction", "CREATURE Template Automation"],
    "reconciliation": ["Python Automation", "Amazon QuickSight", "Data Central + AI"],
    "validation": ["Amazon Textract + Rules Engine", "Python Automation", "Custom AI Agent"],
    "payment": ["Python Automation", "Amazon Quick Suite Flow", "RPA Integration"],
}

# ─── CSV Parsing ──────────────────────────────────────────────────────────────
def get_field(row, *names):
    """Get a field value by trying multiple possible column names (handles both manual CSV and REST API CSV)"""
    for name in names:
        val = row.get(name, "")
        if val and str(val).strip():
            return str(val).strip()
    # Also try partial matches for truncated REST API names (SharePoint truncates to ~20 chars)
    for name in names:
        for prefix_len in [20, 15, 12]:
            prefix = name[:prefix_len]
            if len(prefix) < 5:
                continue
            for key in row:
                if key.startswith(prefix) and row[key] and str(row[key]).strip():
                    return str(row[key]).strip()
    # Also try with _x0020_ encoded spaces
    for name in names:
        encoded = name.replace(" ", "_x0020_")[:20]
        for key in row:
            if key.startswith(encoded[:12]) and row[key] and str(row[key]).strip():
                return str(row[key]).strip()
    return ""


# ─── GitHub CSV Auto-Fetch (for Render / cloud) ──────────────────────────────
GITHUB_CSV_URL = "https://raw.githubusercontent.com/Prathima-create/ai-velocity-portal/main/data/submissions.csv"
_last_github_fetch = 0
GITHUB_FETCH_INTERVAL = 120  # seconds — fetch from GitHub every 2 minutes on cloud

def maybe_fetch_csv_from_github(csv_path):
    """On Render (cloud), fetch the latest CSV from GitHub every 2 minutes.
    This avoids needing Docker rebuilds for data updates."""
    global _last_github_fetch
    import time
    
    # Only auto-fetch on cloud (Render sets RENDER env var, or check if PORT is set by cloud)
    is_cloud = os.environ.get("RENDER") or os.environ.get("RENDER_SERVICE_ID") or (os.environ.get("PORT") and os.name != 'nt')
    if not is_cloud:
        return  # local dev — use local file
    
    now = time.time()
    if now - _last_github_fetch < GITHUB_FETCH_INTERVAL:
        return  # too soon, skip
    
    try:
        import urllib.request
        req = urllib.request.Request(GITHUB_CSV_URL, headers={"Cache-Control": "no-cache"})
        resp = urllib.request.urlopen(req, timeout=15)
        content = resp.read()
        
        if len(content) > 100:  # sanity check
            # Validate the fetched CSV has "Your Manager" column (Form 12 format)
            text = content.decode('utf-8-sig', errors='replace')
            if 'Your Manager' not in text.split('\n')[0]:
                print(f"[GitHub Fetch] Skipped: CSV missing 'Your Manager' column (CDN cache stale)", flush=True)
                _last_github_fetch = now  # don't retry for 2 min
                return
            os.makedirs(os.path.dirname(csv_path), exist_ok=True)
            with open(csv_path, 'wb') as f:
                f.write(content)
            _last_github_fetch = now
            print(f"[GitHub Fetch] Updated CSV from GitHub: {len(content):,} bytes", flush=True)
    except Exception as e:
        print(f"[GitHub Fetch] Failed: {e}", flush=True)


def load_submissions():
    """Load and parse the SharePoint CSV export"""
    csv_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "submissions.csv")
    
    # On cloud, auto-fetch latest CSV from GitHub
    maybe_fetch_csv_from_github(csv_path)
    
    if not os.path.exists(csv_path):
        return []
    
    submissions = []
    with open(csv_path, 'r', encoding='utf-8-sig', errors='replace') as f:
        reader = csv.DictReader(f)
        for idx, row in enumerate(reader):
            submission_type = get_field(row, "What would you like to do", "What_x0020_would_x0020_you_x0020")
            
            # Determine category
            if "Completed AI Win" in submission_type or "completed" in submission_type.lower():
                category = "ai_win"
            elif "Replicate" in submission_type or "replicate" in submission_type.lower():
                category = "replicate"
            else:
                category = "new_idea"
            
            # Parse expected impact
            impact_raw = get_field(row, "Expected Impact if Implemented?", "Expected_x0020_Impact_x0020_if_x", "Rating (0-5)")
            
            # Parse approval status — REST API uses "ApprovalStatus" or "OData__ApprovalStatus"
            approval_status_raw = get_field(row, "Approval Status", "Approval status", "ApprovalStatus", "OData__ApprovalStatus")
            manager_approval = get_field(row, "Manager approval", "Managerapproval")
            l6_approval = get_field(row, "L6/L7 approval")
            tech_approval = get_field(row, "Tech team approval", "Techteamapproval")
            
            # Determine overall status
            if approval_status_raw == "1":
                status = "Approved"
            elif approval_status_raw == "2":
                status = "In Review"
            elif approval_status_raw == "3":
                status = "Pending"
            elif approval_status_raw == "0":
                status = "New"
            else:
                status = "Pending"
            
            # For AI Wins, status is always "Completed"
            if category == "ai_win":
                status = "Completed"
            
            # Parse execution plan
            execution_plan = get_field(row, "How do you plan to execute this idea?", "How_x0020_do_x0020_you_x0020_pla")
            
            # Build submission object
            name = get_field(row, "Name", "Your name", "Title")
            process = get_field(row, "Process", "Process you are in", "Select your process", "Select_x0020_your_x0020_process", "Process_x0020_you_x0020_are_x002")
            sub_process = get_field(row, "Sub Process", "Sub Process you are in", "Select your Sub process", "Select_x0020_your_x0020_Sub_x002", "Sub_x0020_Process_x0020_you_x002")
            
            # For AI wins
            project_name = get_field(row, "Project Name")
            project_owner = get_field(row, "Project Owner/Lead")
            project_team = get_field(row, "Project Team")
            tech_poc = get_field(row, "Tech team POC")
            challenge = get_field(row, "Challenge addressed ", "Challenge addressed")
            ai_solution_win = get_field(row, "AI solution ", "AI solution")
            impact_win = get_field(row, "Impact")
            replicable = get_field(row, "Can this solution be replicated by others?", "Can_x0020_this_x0020_solution_x0")
            
            # For ideas
            problem_statement = get_field(row, "Problem Statement")
            current_effort = get_field(row, "Current Manual Effort", "Current_x0020_Manual_x0020_Effor")
            proposed_solution = get_field(row, "Proposed AI Solution")
            estimated_volume = get_field(row, "Estimated Volume", "Estimated_x0020_volume0")
            impact_types = get_field(row, "Expected Impact if Implemented?", "Expected_x0020_Impact_x0020_if_x", "Other Impact", "Other_x0020_Impact")
            data_available = get_field(row, "Data available")
            support_required = get_field(row, "Support Required (if any)", "Support_x0020_Required_x0020__x0")
            target_timeline = get_field(row, "Target Timeline")
            can_replicate = get_field(row, "Can this be replicated across teams?", "Can_x0020_this_x0020_be_x0020_re")
            
            # For replicate requests
            which_win = get_field(row, "Which AI Win are you interested in replicating?", "Which_x0020_AI_x0020_Win_x0020_a")
            current_process_desc = get_field(row, "Briefly describe your current process", "Briefly_x0020_describe_x0020_you", "Define your Current Process", "Define_x0020_your_x0020_Current_")
            
            # Implementation stage (new column)
            impl_stage_raw = get_field(row, "If your idea is being implemented by you , what stage it is in?", "If_x0020_your_x0020_idea_x0020_i")
            
            # Dates
            created = get_field(row, "Created")
            modified = get_field(row, "Modified")
            created_by = get_field(row, "Created By")
            modified_by = get_field(row, "Modified By")
            manager = get_field(row, "Your Manager")
            team = get_field(row, "Your Team ", "Your Team")
            
            # Suggest tools based on problem keywords
            suggested_tools = suggest_tools(
                problem_statement + " " + proposed_solution + " " + challenge + " " + ai_solution_win
            )
            
            # Get SDE contact
            sde_contact = get_sde_contact(process)
            
            # Normalize implementation stage
            if "not required" in impl_stage_raw.lower() or "completed win" in impl_stage_raw.lower() or "ready for production" in impl_stage_raw.lower():
                impl_stage = "Completed (Production)"
            elif "awaiting approval" in impl_stage_raw.lower():
                impl_stage = "Completed (Awaiting Approvals)"
            elif "uat" in impl_stage_raw.lower():
                impl_stage = "In Progress (UAT Stage)"
            elif "development" in impl_stage_raw.lower():
                impl_stage = "In Progress (Development Stage)"
            elif impl_stage_raw:
                impl_stage = impl_stage_raw
            else:
                impl_stage = ""
            
            # Get leader info — manager-based lookup takes PRIORITY
            leader_info = MANAGER_TO_LEADER.get(manager) or get_leader(process)
            
            # Get Tech team POC / Support Team from CSV (if available)
            support_team = get_field(row, "Support Team/ Partnership Team", "Support Team/ StringId")
            tech_team_poc = get_field(row, "Tech team POC", "Tech team POCStringId")
            
            # Build exploration tip — helps submitter know which tools to start with
            if suggested_tools and category == "new_idea":
                top_tools = suggested_tools[:3]
                exploration_tip = f"💡 You can start exploring your idea with: {', '.join(top_tools)}. Reach out to your Tech POC ({sde_contact.get('name', 'AI Velocity Team')}) for guidance on implementation."
            elif category == "replicate":
                exploration_tip = f"💡 Connect with your Tech POC ({sde_contact.get('name', 'AI Velocity Team')}) to understand the existing solution and how to adapt it for your process."
            else:
                exploration_tip = ""
            
            submission = {
                "id": idx + 1,
                "category": category,
                "submission_type": submission_type,
                "name": name,
                "created_by": created_by,
                "process": process,
                "sub_process": sub_process,
                "project_name": project_name,
                "project_owner": project_owner,
                "project_team": project_team,
                "tech_poc": tech_poc or tech_team_poc,
                "support_team": support_team,
                "challenge": challenge,
                "ai_solution": ai_solution_win or proposed_solution,
                "impact": impact_win or impact_types,
                "replicable": replicable or can_replicate,
                "problem_statement": problem_statement,
                "current_effort": current_effort,
                "proposed_solution": proposed_solution,
                "estimated_volume": estimated_volume,
                "execution_plan": execution_plan,
                "data_available": data_available,
                "support_required": support_required,
                "target_timeline": target_timeline,
                "which_win_to_replicate": which_win,
                "current_process_desc": current_process_desc,
                "status": status,
                "manager_approval": manager_approval,
                "l6_approval": l6_approval,
                "tech_approval": tech_approval,
                "manager": manager,
                "team": team,
                "created": created,
                "modified": modified,
                "modified_by": modified_by,
                "suggested_tools": suggested_tools,
                "exploration_tip": exploration_tip,
                "sde_contact": sde_contact,
                "leader": leader_info.get("leader", "Unknown"),
                "leader_poc": leader_info.get("poc", "TBD"),
                "implementation_stage": impl_stage,
            }
            
            submissions.append(submission)
    
    return submissions


def suggest_tools(text: str) -> List[str]:
    """Suggest AI tools based on problem/solution keywords"""
    text_lower = text.lower()
    tools = set()
    
    keyword_map = {
        "document": ["pdf", "document", "extract", "bol", "credit note", "invoice copy"],
        "extraction": ["extract", "parse", "read", "ocr", "scan"],
        "chatbot": ["chatbot", "chat bot", "chat agent", "conversational"],
        "knowledge": ["sop", "knowledge", "wiki", "knowledge base", "information"],
        "automation": ["automate", "automation", "workflow", "flow", "trigger", "schedule"],
        "analytics": ["analytics", "dashboard", "report", "metrics", "kpi", "trend"],
        "sentiment": ["sentiment", "nrr", "negative response", "frustration"],
        "translation": ["translat", "language", "german", "local language"],
        "audit": ["audit", "quality", "qc", "validation check", "compliance check"],
        "email": ["email", "notification", "reminder", "follow-up", "correspondence"],
        "ocr": ["ocr", "scan", "image", "handwritten"],
        "classification": ["classif", "categoriz", "routing", "cti"],
        "dashboard": ["dashboard", "visualization", "quicksight", "report"],
        "sop": ["sop", "standard operating", "procedure", "blurb"],
        "invoice": ["invoice", "payment", "billing", "creature"],
        "reconciliation": ["reconcil", "matching", "comparison", "vlookup"],
        "validation": ["validat", "verify", "check", "compliance"],
        "payment": ["payment", "batch", "scheduling", "template"],
    }
    
    for category, keywords in keyword_map.items():
        for kw in keywords:
            if kw in text_lower:
                tools.update(TOOL_SUGGESTIONS.get(category, []))
                break
    
    if not tools:
        tools = {"Amazon Q Business", "Amazon Quick Suite", "Python Automation"}
    
    return list(tools)[:5]


def get_sde_contact(process: str) -> Dict[str, str]:
    """Get the SDE/Tech contact for a given process"""
    if process in SDE_CONTACTS:
        return SDE_CONTACTS[process]
    
    # Fuzzy match
    for key, contact in SDE_CONTACTS.items():
        if key.lower() in process.lower() or process.lower() in key.lower():
            return contact
    
    return {"alias": "pratpk", "name": "Prathima K", "team": "AI Velocity Program Lead"}


def get_leader(process: str) -> Dict[str, str]:
    """Get the org leader for a given process"""
    if process in LEADER_MAPPING:
        return LEADER_MAPPING[process]
    for key, info in LEADER_MAPPING.items():
        if key.lower() in process.lower() or process.lower() in key.lower():
            return info
    return {"leader": "Unknown", "poc": "TBD"}


# ─── API Routes ───────────────────────────────────────────────────────────────

@app.get("/api/submissions")
async def get_submissions(
    category: Optional[str] = None,
    process: Optional[str] = None,
    status: Optional[str] = None,
    search: Optional[str] = None
):
    """Get all submissions with optional filters"""
    data = load_submissions()
    
    if category:
        data = [s for s in data if s["category"] == category]
    if process:
        data = [s for s in data if process.lower() in s["process"].lower()]
    if status:
        data = [s for s in data if s["status"].lower() == status.lower()]
    if search:
        search_lower = search.lower()
        data = [s for s in data if 
            search_lower in s.get("name", "").lower() or
            search_lower in s.get("problem_statement", "").lower() or
            search_lower in s.get("proposed_solution", "").lower() or
            search_lower in s.get("project_name", "").lower() or
            search_lower in s.get("process", "").lower() or
            search_lower in s.get("challenge", "").lower() or
            search_lower in s.get("ai_solution", "").lower()
        ]
    
    return data


@app.get("/api/submissions/{submission_id}")
async def get_submission(submission_id: int):
    """Get a specific submission by ID"""
    data = load_submissions()
    for s in data:
        if s["id"] == submission_id:
            return s
    return {"error": "Not found"}


@app.get("/api/stats")
async def get_stats():
    """Get dashboard statistics"""
    data = load_submissions()
    
    total = len(data)
    ideas = [s for s in data if s["category"] == "new_idea"]
    wins = [s for s in data if s["category"] == "ai_win"]
    replicates = [s for s in data if s["category"] == "replicate"]
    
    # Process breakdown
    processes = Counter(s["process"] for s in data if s["process"])
    
    # Status breakdown
    statuses = Counter(s["status"] for s in data)
    
    # Unique submitters
    submitters = set(s["name"] for s in data if s["name"])
    
    # Timeline breakdown
    timeline_counts = Counter(s["target_timeline"] for s in data if s["target_timeline"])
    
    # Implementation stage counts — only from completed wins
    live_in_prod = len([s for s in wins if s.get("implementation_stage") == "Completed (Production)"])
    uat_in_progress = len([s for s in wins if s.get("implementation_stage") != "Completed (Production)"])
    
    return {
        "total_submissions": total,
        "live_in_production": live_in_prod,
        "uat_in_progress": uat_in_progress,
        "new_ideas": len(ideas),
        "completed_wins": len(wins),
        "replicate_requests": len(replicates),
        "unique_submitters": len(submitters),
        "processes": dict(processes.most_common(15)),
        "statuses": dict(statuses),
        "timelines": dict(timeline_counts),
        "server_time": datetime.now().isoformat()
    }


@app.get("/api/ai-wins")
async def get_ai_wins():
    """Get only completed AI wins"""
    data = load_submissions()
    return [s for s in data if s["category"] == "ai_win"]


@app.get("/api/ideas")
async def get_ideas():
    """Get only new AI ideas"""
    data = load_submissions()
    return [s for s in data if s["category"] == "new_idea"]


@app.get("/api/replicates")
async def get_replicates():
    """Get only replicate requests"""
    data = load_submissions()
    return [s for s in data if s["category"] == "replicate"]


@app.get("/api/suggest-tools")
async def suggest_tools_api(problem: str = Query(..., description="Problem description")):
    """Suggest AI tools based on problem description"""
    tools = suggest_tools(problem)
    return {"suggested_tools": tools}


@app.get("/api/sde-contact")
async def get_sde_contact_api(process: str = Query(..., description="Process name")):
    """Get SDE/Tech contact for a process"""
    contact = get_sde_contact(process)
    return contact


@app.get("/api/leaderboard")
async def get_leaderboard():
    """Get leader-wise breakdown of submissions"""
    data = load_submissions()
    leaders = {}
    for s in data:
        leader = s.get("leader", "Unknown")
        if leader not in leaders:
            leaders[leader] = {"leader": leader, "poc": s.get("leader_poc", "TBD"), "total": 0, "wins": 0, "ideas": 0, "replicates": 0, "in_progress": 0, "processes": set(), "contributors": set()}
        leaders[leader]["total"] += 1
        if s["category"] == "ai_win":
            leaders[leader]["wins"] += 1
        elif s["category"] == "new_idea":
            if s["status"] in ("Approved", "In Review"):
                leaders[leader]["in_progress"] += 1
            else:
                leaders[leader]["ideas"] += 1
        elif s["category"] == "replicate":
            leaders[leader]["replicates"] += 1
        if s["process"]:
            leaders[leader]["processes"].add(s["process"])
        if s["name"]:
            leaders[leader]["contributors"].add(s["name"])
    
    result = []
    for l in sorted(leaders.values(), key=lambda x: x["total"], reverse=True):
        result.append({
            "leader": l["leader"],
            "poc": l["poc"],
            "total": l["total"],
            "wins": l["wins"],
            "ideas": l["ideas"],
            "in_progress": l["in_progress"],
            "replicates": l["replicates"],
            "contributors": len(l["contributors"]),
        })
    return result


@app.get("/api/stages")
async def get_stages():
    """Get implementation stage breakdown for ideas/wins being implemented"""
    data = load_submissions()
    stages = Counter(s["implementation_stage"] for s in data if s["implementation_stage"])
    # Order them in a logical pipeline order
    stage_order = [
        "In Progress (Development Stage)",
        "In Progress (UAT Stage)",
        "Completed (Awaiting Approvals)",
        "Completed (Production)",
    ]
    result = {}
    for stage in stage_order:
        if stage in stages:
            result[stage] = stages[stage]
    # Add any others not in the order
    for stage, count in stages.items():
        if stage not in result:
            result[stage] = count
    return result


@app.get("/api/processes")
async def get_processes():
    """Get list of all unique processes"""
    data = load_submissions()
    processes = sorted(set(s["process"] for s in data if s["process"]))
    return processes


@app.post("/api/sync")
async def trigger_sync():
    """Trigger SharePoint data sync via GitHub Actions"""
    import requests as req
    
    github_token = os.environ.get("GITHUB_PAT", "")
    repo = os.environ.get("GITHUB_REPO", "Prathima-create/ai-velocity-portal")
    
    if not github_token:
        # Fallback: try local Selenium sync (only works on Windows with Edge browser)
        if os.name == 'nt':
            import subprocess
            script_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "scripts", "sync_sharepoint.py")
            if os.path.exists(script_path):
                try:
                    kwargs = {"stdout": subprocess.PIPE, "stderr": subprocess.PIPE, "creationflags": subprocess.CREATE_NO_WINDOW}
                    proc = subprocess.Popen([sys.executable, script_path], **kwargs)
                    return {"status": "started", "message": "🚀 Local sync started! Edge browser will open. Dashboard refreshes in 60s.", "pid": proc.pid}
                except Exception as e:
                    return {"status": "error", "message": str(e)}
        # Cloud server without GITHUB_PAT
        return {"status": "started", "message": "📡 Data auto-syncs every hour from your laptop. Click 🔄 Refresh to see the latest data. Or use 📤 Upload to update manually."}
    
    # Trigger GitHub Actions workflow via repository_dispatch
    try:
        resp = req.post(
            f"https://api.github.com/repos/{repo}/dispatches",
            json={"event_type": "sync-sharepoint"},
            headers={
                "Authorization": f"token {github_token}",
                "Accept": "application/vnd.github.v3+json"
            },
            timeout=10
        )
        if resp.status_code == 204:
            return {"status": "started", "message": "🚀 SharePoint sync triggered! Data will update in ~2 minutes. Click Refresh after that."}
        else:
            return {"status": "error", "message": f"GitHub API returned {resp.status_code}: {resp.text[:200]}"}
    except Exception as e:
        return {"status": "error", "message": f"Could not trigger sync: {str(e)}"}


@app.get("/api/sync-status")
async def sync_status():
    """Check last sync time by checking CSV file modification time"""
    csv_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "submissions.csv")
    if os.path.exists(csv_path):
        mtime = os.path.getmtime(csv_path)
        from datetime import datetime as dt
        last_sync = dt.fromtimestamp(mtime).isoformat()
        size = os.path.getsize(csv_path)
        return {"last_sync": last_sync, "file_size": size, "exists": True}
    return {"last_sync": None, "exists": False}


@app.post("/api/upload-csv")
async def upload_csv(file: UploadFile = File(...)):
    """Upload a new SharePoint CSV to update dashboard data instantly"""
    if not file.filename.endswith('.csv'):
        return {"status": "error", "message": "Only CSV files are allowed"}
    csv_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "submissions.csv")
    os.makedirs(os.path.dirname(csv_path), exist_ok=True)
    content = await file.read()
    with open(csv_path, 'wb') as f:
        f.write(content)
    # Count rows to verify
    with open(csv_path, 'r', encoding='utf-8-sig', errors='replace') as f:
        reader = csv.reader(f)
        rows = sum(1 for _ in reader) - 1  # minus header
    return {"status": "success", "message": f"CSV uploaded! {rows} rows loaded.", "rows": rows, "size": len(content)}


# ─── Serve Frontend ──────────────────────────────────────────────────────────

frontend_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "frontend")

@app.get("/")
async def serve_frontend():
    return FileResponse(os.path.join(frontend_path, "index.html"))

@app.get("/program.html")
async def serve_program():
    return FileResponse(os.path.join(frontend_path, "program.html"))

@app.get("/training.html")
async def serve_training():
    return FileResponse(os.path.join(frontend_path, "training.html"))

if os.path.exists(frontend_path):
    app.mount("/static", StaticFiles(directory=frontend_path), name="static")
    app.mount("/", StaticFiles(directory=frontend_path, html=True), name="frontend")

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 3000))
    uvicorn.run(app, host="0.0.0.0", port=port)
