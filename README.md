# 🚀 AI Velocity Portal - AI Wins Dashboard

> **Accelerating AI adoption across Amazon's Accounts Payable organization**

A real-time dashboard powered by SharePoint data that tracks all AI submissions, completed wins, new ideas, and replication requests across the AP organization.

## 📊 Dashboard Features

| Section | Description |
|---------|-------------|
| **KPI Dashboard** | Real-time counts: Total Submissions, Ideas, Wins, Replicates, Contributors |
| **Process Charts** | Visual breakdown of submissions by process and approval status |
| **🏆 AI Wins** | 16 completed AI projects with impact details, team info, and replicability |
| **💡 Ideas Pipeline** | 90 submitted ideas with search, filters, approval status, tool suggestions |
| **🔁 Replicate** | 9 requests to replicate successful AI wins across teams |
| **⚙️ Workflow** | Visual pipeline: Submit → Approve → Tool Suggest → SDE Assign → Build |

## 🔗 Data Source

Data is automatically loaded from the **SharePoint AI Velocity Submission Portal** CSV export:
- `data/submissions.csv` — Export from [SharePoint List](https://amazon.sharepoint.com/sites/AI-Velocity-site/Lists/AI%20Velocity%20Submission%20Portal/AllItems.aspx)

### To refresh data:
1. Go to SharePoint → Export to Excel/CSV
2. Replace `data/submissions.csv` with the new export
3. Restart the server — dashboard updates automatically

## 🤖 Smart Features

- **AI Tool Suggestions**: Automatically suggests appropriate tools (Textract, Q Business, Quick Suite, etc.) based on problem description keywords
- **SDE Contact Mapping**: Shows the right tech contact for each process area
- **Approval Pipeline**: Visual tracker showing submission → manager → L6/L7 → tech review → implementation status
- **Click-to-Detail**: Click any card to see full details including problem statement, solution, impact, suggested tools, and SDE contact

## 🏗️ Tech Stack

- **Frontend**: HTML5, CSS3 (Dark Glassmorphism), Vanilla JavaScript
- **Backend**: Python FastAPI with CSV parsing
- **Data**: SharePoint CSV export (115 submissions from 60 contributors)

## 🚀 Quick Start

```bash
# Option 1: Run the start script
start.bat

# Option 2: Manual start
pip install fastapi uvicorn pydantic python-dotenv
python -m uvicorn backend.main:app --host 0.0.0.0 --port 3000
```

Then open **http://localhost:3000**

## 📁 Project Structure

```
ai-velocity-portal/
├── frontend/
│   ├── index.html     # AI Wins Dashboard (main page)
│   ├── styles.css     # Dark theme with glassmorphism
│   └── app.js         # Dashboard logic, API integration, modals
├── backend/
│   ├── main.py        # FastAPI server, CSV parser, tool suggestions, SDE mapping
│   └── requirements.txt
├── data/
│   └── submissions.csv # SharePoint export (auto-parsed by backend)
├── start.bat          # One-click launcher
└── README.md
```

## 📊 API Endpoints

| Endpoint | Description |
|----------|-------------|
| `GET /api/stats` | Dashboard KPIs and breakdowns |
| `GET /api/submissions` | All submissions (filterable by category, process, status, search) |
| `GET /api/ai-wins` | Only completed AI wins |
| `GET /api/ideas` | Only new AI ideas |
| `GET /api/replicates` | Only replicate requests |
| `GET /api/processes` | List of all unique processes |
| `GET /api/suggest-tools?problem=...` | Get AI tool suggestions |
| `GET /api/sde-contact?process=...` | Get SDE contact for a process |

## 🏆 Current AI Wins (16 Projects)

| Project | Owner | Key Impact |
|---------|-------|------------|
| ATLAS | Subhra Majumder | 90% efficiency, 1,519 hrs saved annually |
| Project Athena | Prathima K | 3% → 100% call audit coverage |
| Sentinel | Venkatesh D. | 100% transaction audit coverage |
| GPP Invoicing | Arunima M. | 3.55 FTE savings |
| FinGenie | Arunima M. | 2 FTE savings |
| OptiGuide | Nirmal Kumar S. | SOP review minutes → seconds |
| NRR Sentiment | Somasekhar M. | NRR 11.9% → 2.9% |
| TicketIQ Analytics | Niraj Sharma | Structured KPI tracking |
| BOL PDF Extractor | Prafful Kakkar | 100 hrs annual savings |
| Project RACE | Madhavi M. | 240 hrs annual savings |

**Program Lead**: Prathima K | **Sponsor**: Amit's Org

---
*Built for AI Velocity - Accounts Payable | Amazon Internal*
