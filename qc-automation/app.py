"""
FinCom QC Automation Dashboard — Streamlit Web App v3
=====================================================
Matches the interactive HTML dashboard look & feel.
RAG colors, KPI cards, status cards, bar charts, org scorecard,
analyst table, filter bar — all matching the original qc_automation.py HTML output.

Data flow: SharePoint → S3 → EC2 Streamlit (CloudFront)
"""

import streamlit as st
import pandas as pd
import os, re, json, warnings
from pathlib import Path
from datetime import datetime, timedelta
from collections import Counter

warnings.filterwarnings('ignore')

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="FinCom QC Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ============================================================
# CONSTANTS
# ============================================================
DATA_DIR = Path(__file__).parent / "data"
DATA_DIR.mkdir(exist_ok=True)

PARAMETERS = [
    "1. Greeting", "2. Empathy", "3. Language",
    "4. Appropriate Resolution", "5. Delay/Escalation",
    "6. Phone Calls", "7. Proactive/Self-Service",
    "8. Transfer/SIM CTI", "9. Confidentiality",
    "10. Annotations", "11. Resolution Code",
]
FATAL_NAMES = {"appropriateresolution", "confidentiality"}


# ============================================================
# CUSTOM CSS (matching original HTML dashboard)
# ============================================================
def inject_css():
    st.markdown("""
    <style>
    /* Hide Streamlit chrome */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .block-container { padding: 16px 24px; max-width: 1400px; }
    
    /* Typography */
    .dashboard-title { font-size: 18px; font-weight: 700; color: #1f2937; margin: 0 0 4px 0; }
    .dashboard-sub { font-size: 12px; color: #6b7280; margin-bottom: 16px; }
    .section-title { font-size: 14px; font-weight: 600; color: #374151; margin: 0 0 10px 0; }
    
    /* KPI Cards */
    .kpi-row { display: flex; gap: 12px; margin-bottom: 16px; flex-wrap: wrap; }
    .kpi-card { flex: 1 1 0; min-width: 180px; background: #fff; padding: 16px;
                border-radius: 10px; box-shadow: 0 1px 3px rgba(0,0,0,.06);
                transition: transform .1s; cursor: default; }
    .kpi-card:hover { transform: translateY(-2px); box-shadow: 0 4px 10px rgba(0,0,0,.1); }
    .kpi-label { font-size: 11px; color: #6b7280; text-transform: uppercase; letter-spacing: 0.5px; }
    .kpi-value { font-size: 28px; font-weight: 700; margin: 6px 0; color: #1f2937; }
    .kpi-sub { font-size: 11px; color: #6b7280; }
    .kpi-fatal .kpi-value { color: #dc2626; }
    
    /* RAG Colors */
    .rag-green { background: #d1fae5 !important; }
    .rag-amber { background: #fef3c7 !important; }
    .rag-red { background: #fee2e2 !important; }
    .rag-green-text { color: #047857; font-weight: 600; }
    .rag-amber-text { color: #b45309; font-weight: 600; }
    .rag-red-text { color: #b91c1c; font-weight: 600; }
    
    /* Status Cards */
    .status-row-block { display: flex; gap: 12px; margin-bottom: 16px; flex-wrap: wrap; }
    .status-card { flex: 1 1 0; min-width: 200px; background: #fff; border-radius: 10px;
                   padding: 14px; box-shadow: 0 1px 3px rgba(0,0,0,.06); }
    .status-header { font-weight: 600; font-size: 13px; color: #374151;
                     border-bottom: 1px solid #e5e7eb; padding-bottom: 6px; margin-bottom: 8px; }
    .status-total { font-size: 13px; color: #6b7280; margin-bottom: 8px; }
    .status-item { display: flex; justify-content: space-between; font-size: 13px; padding: 3px 0; }
    .status-sla { margin-top: 10px; padding: 6px 8px; border-radius: 6px;
                  text-align: center; font-size: 12px; font-weight: 600; }
    .sla-green { background: #d1fae5; color: #047857; }
    .sla-red { background: #fee2e2; color: #b91c1c; }
    
    /* Bar Chart */
    .bar-chart { margin-top: 8px; }
    .bar-row { display: flex; align-items: center; margin-bottom: 6px; gap: 8px; }
    .bar-label { width: 180px; font-size: 12px; color: #374151; white-space: nowrap;
                 overflow: hidden; text-overflow: ellipsis; flex-shrink: 0; }
    .bar-track { flex: 1; height: 20px; background: #f3f4f6; border-radius: 4px; overflow: hidden; }
    .bar-fill { height: 100%; background: #3b82f6; border-radius: 4px; transition: width .3s; }
    .bar-fill-fatal { background: #dc2626; }
    .bar-fill-amber { background: #f59e0b; }
    .bar-value { width: 36px; text-align: right; font-size: 12px; font-weight: 600; flex-shrink: 0; }
    
    /* WoW bars */
    .wow-block { margin-bottom: 16px; }
    .wow-title { font-size: 12px; font-weight: 600; color: #374151; margin-bottom: 6px; }
    .wow-bars { display: flex; gap: 6px; align-items: flex-end; height: 100px; }
    .wow-bar-col { display: flex; flex-direction: column; align-items: center; flex: 1; }
    .wow-bar { width: 100%; max-width: 40px; height: 80px; background: #f3f4f6;
               border-radius: 4px 4px 0 0; position: relative; overflow: hidden;
               display: flex; align-items: flex-end; }
    .wow-fill { width: 100%; border-radius: 4px 4px 0 0; transition: height .3s; }
    .wow-num { font-size: 11px; font-weight: 600; margin-bottom: 2px; }
    .wow-lbl { font-size: 10px; color: #6b7280; margin-top: 4px; }
    .arr-up { color: #dc2626; font-size: 11px; }
    .arr-dn { color: #047857; font-size: 11px; }
    
    /* MoM Table */
    .mom-table { width: 100%; border-collapse: collapse; font-size: 13px; }
    .mom-table th { text-align: center; color: #6b7280; font-size: 11px; padding: 6px; }
    .mom-table td { padding: 8px 6px; border-bottom: 1px solid #f3f4f6; }
    .mom-lbl { font-weight: 500; color: #374151; }
    .mom-val { text-align: center; font-weight: 600; }
    .mom-delta { font-size: 11px; margin-top: 2px; }
    .vs { text-align: center; color: #9ca3af; font-size: 11px; }
    
    /* Data Tables */
    .data-table { width: 100%; border-collapse: collapse; font-size: 12px; }
    .data-table th { background: #f9fafb; padding: 8px; text-align: left;
                     font-size: 11px; color: #6b7280; text-transform: uppercase;
                     border-bottom: 2px solid #e5e7eb; cursor: pointer; }
    .data-table th:hover { background: #f3f4f6; }
    .data-table td { padding: 8px; border-bottom: 1px solid #f3f4f6; }
    .data-table tr:hover { background: #f9fafb; }
    
    /* Section containers */
    .section-box { background: #fff; border-radius: 10px; padding: 14px;
                   box-shadow: 0 1px 3px rgba(0,0,0,.06); margin-bottom: 16px; }
    .flex-row { display: flex; gap: 16px; margin-bottom: 16px; }
    .flex-half { flex: 1 1 calc(50% - 8px); min-width: 300px; }
    
    /* Search */
    .search-box { width: 100%; padding: 8px 12px; border: 1px solid #d1d5db;
                  border-radius: 6px; font-size: 13px; margin-bottom: 8px; }
    
    /* Filter bar */
    .filter-bar { background: #fff; padding: 10px 14px; border-radius: 10px;
                  box-shadow: 0 1px 3px rgba(0,0,0,.06); margin-bottom: 16px;
                  font-size: 12px; color: #6b7280; }
    </style>
    """, unsafe_allow_html=True)


# ============================================================
# HELPERS
# ============================================================
def norm(s):
    return (str(s).strip().lower()
            .replace('/', '').replace('-', '')
            .replace('_', '').replace(' ', '')
            .replace('.', '').replace(',', ''))


def is_defect(cell):
    return str(cell).strip() == "0"


def is_fatal_param(param_name):
    n = norm(param_name)
    return any(fk in n for fk in FATAL_NAMES)


def find_col_strict(df, *candidates):
    cols = list(df.columns)
    for cand in candidates:
        key = norm(cand)
        if not key:
            continue
        for col in cols:
            if col == key:
                return col
        for col in cols:
            if col.startswith(key):
                rest = col[len(key):]
                if not rest or not rest[0].isalnum():
                    return col
    return None


def get_series(df, *candidates):
    col = find_col_strict(df, *candidates)
    if col is not None:
        return df[col].fillna('').astype(str).str.strip()
    return pd.Series([''] * len(df), index=df.index)


def parse_accuracy_series(series):
    def _p(v):
        try:
            v = float(str(v).replace('%', '').strip())
            return round(v * 100, 2) if v <= 1.0 else round(v, 2)
        except Exception:
            return None
    return series.apply(_p)


def parse_date_series(series):
    def _p(v):
        v = str(v).strip().split(' ')[0].split(',')[0]
        if not v or v.lower() in ('nan', 'none', '', 'na', 'n/a'):
            return None
        parts = v.replace('-', '/').split('/')
        if len(parts) != 3:
            return None
        try:
            a, b, c = int(parts[0]), int(parts[1]), int(parts[2])
            year = c if c > 1000 else 2000 + c
            if a > 12: return datetime(year, b, a)
            if b > 12: return datetime(year, a, b)
            return datetime(year, a, b)
        except Exception:
            return None
    return series.apply(_p)


def iso_week(d):
    if d is None or pd.isna(d):
        return None
    try:
        return f"W{d.isocalendar()[1]}"
    except Exception:
        return None


def workdays_between(start, end):
    if start is None or end is None:
        return None
    if end < start:
        return 0
    days = 0
    cur = start
    while cur <= end:
        if cur.weekday() < 5:
            days += 1
        cur += timedelta(days=1)
    return max(0, days - 1)


def rag_class(pct, inverted=False):
    if pct is None:
        return ''
    if inverted:
        if pct <= 5: return 'rag-green'
        if pct <= 10: return 'rag-amber'
        return 'rag-red'
    if pct >= 95: return 'rag-green'
    if pct >= 90: return 'rag-amber'
    return 'rag-red'


def rag_text_class(pct, inverted=False):
    if pct is None:
        return ''
    if inverted:
        if pct <= 5: return 'rag-green-text'
        if pct <= 10: return 'rag-amber-text'
        return 'rag-red-text'
    if pct >= 95: return 'rag-green-text'
    if pct >= 90: return 'rag-amber-text'
    return 'rag-red-text'


def fmt_pct(p):
    return '—' if p is None else f'{p:.1f}%'


# ============================================================
# DATA LOADING — S3 bucket or local SharePoint-synced folder
# ============================================================
SHAREPOINT_BASE = Path(r"C:\Users\pratpk\amazon.com\Automation hosting - Documents")
CURRENT_MONTH_FOLDER = SHAREPOINT_BASE / "General_Apr_2026"
PREV_MONTH_FOLDER = SHAREPOINT_BASE / "General_Mar_2026"


def _get_s3_client():
    import boto3
    profile = os.environ.get('AWS_PROFILE', None)
    if profile:
        session = boto3.Session(profile_name=profile)
    else:
        session = boto3.Session()
    return session.client('s3')


@st.cache_data(ttl=300)
def load_data_from_s3():
    bucket = os.environ.get('QC_S3_BUCKET', 'fincom-qc-data')
    prefix = os.environ.get('QC_S3_PREFIX', 'General_Apr_2026/')

    try:
        s3 = _get_s3_client()
        response = s3.list_objects_v2(Bucket=bucket, Prefix=prefix)
        contents = response.get('Contents', [])
        if contents:
            files = {}
            for obj in contents:
                key = obj['Key']
                filename = key.split('/')[-1]
                if filename.endswith('.csv'):
                    local_path = DATA_DIR / filename
                    s3.download_file(bucket, key, str(local_path))
                    files[filename] = local_path
            if files:
                return files, "S3", prefix
    except Exception:
        pass

    local_data = os.environ.get('QC_LOCAL_DATA', '')
    folder = Path(local_data) if local_data else CURRENT_MONTH_FOLDER
    if folder.exists():
        files = {}
        for f in folder.glob("*.csv"):
            files[f.name] = f
        if files:
            return files, "Local (SharePoint)", str(folder)

    files = {}
    for f in DATA_DIR.glob("*.csv"):
        files[f.name] = f
    if files:
        return files, "Local (data/)", str(DATA_DIR)
    return {}, "No data found", ""


@st.cache_data(ttl=300)
def load_prev_month_from_s3():
    bucket = os.environ.get('QC_S3_BUCKET', 'fincom-qc-data')
    prefix = os.environ.get('QC_S3_PREV_PREFIX', 'General_Mar_2026/')

    try:
        s3 = _get_s3_client()
        response = s3.list_objects_v2(Bucket=bucket, Prefix=prefix)
        contents = response.get('Contents', [])
        files = {}
        for obj in contents:
            key = obj['Key']
            filename = key.split('/')[-1]
            if filename.endswith('.csv'):
                local_path = DATA_DIR / f"prev_{filename}"
                s3.download_file(bucket, key, str(local_path))
                files[filename] = local_path
        if files:
            return files
    except Exception:
        pass

    local_prev = os.environ.get('QC_LOCAL_PREV_DATA', '')
    folder = Path(local_prev) if local_prev else PREV_MONTH_FOLDER
    if folder.exists():
        files = {}
        for f in folder.glob("*.csv"):
            files[f.name] = f
        return files
    return {}


# ============================================================
# AUDIT FILE PARSER
# ============================================================
def parse_audit_file(filepath):
    if not filepath.exists():
        return pd.DataFrame()
    try:
        raw = pd.read_csv(filepath, header=None, encoding='utf-8-sig',
                          on_bad_lines='skip', dtype=str)
    except Exception:
        return pd.DataFrame()

    header_row_idx = None
    for i in range(min(200, len(raw))):
        row = raw.iloc[i]
        cells = [str(v).strip() for v in row.values if pd.notna(v) and str(v).strip()]
        if len(cells) < 10:
            continue
        normalized = [norm(c) for c in cells if len(c) <= 60]
        has_analyst = any(c == 'analyst' or c == 'primaryanalyst' for c in normalized)
        has_accuracy = any('accuracy' in c and len(c) <= 15 for c in normalized)
        has_date = any(c == 'auditdate' or c == 'date' for c in normalized)
        has_case = any(c == 'casenumber' or c == 'caseno' or c == 'caseid' for c in normalized)
        if has_analyst and has_accuracy and has_date and has_case:
            header_row_idx = i
            break

    if header_row_idx is None:
        return pd.DataFrame()

    df = pd.read_csv(filepath, header=header_row_idx, encoding='utf-8-sig',
                     on_bad_lines='skip', dtype=str)
    df = df.dropna(how='all')
    df.columns = [norm(c) for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated(keep='first')]

    missed_idx = None
    for i, c in enumerate(df.columns):
        if 'missedparameters' in c or 'ofmissed' in c or c == 'missed':
            missed_idx = i
            break
    if missed_idx is None:
        missed_idx = len(df.columns)

    param_col_positions = []
    for i in range(11):
        pos = missed_idx + 1 + (i * 2)
        if pos < len(df.columns):
            param_col_positions.append(pos)

    param_cols = {}
    for i, p_name in enumerate(PARAMETERS):
        if i < len(param_col_positions):
            param_cols[p_name] = df.columns[param_col_positions[i]]
        else:
            param_cols[p_name] = None

    case_col = find_col_strict(df, "Case Number", "Case No", "Case ID")
    if case_col:
        df = df[df[case_col].astype(str).str.strip().str.match(r'^\d+(-\d+)?$', na=False)].copy()

    if len(df) == 0:
        return pd.DataFrame()

    df['_analyst'] = get_series(df, "Analyst")
    df['_org'] = get_series(df, "ORG", "Org").replace('', 'Unknown')
    df['_case'] = get_series(df, "Case Number", "Case No", "Case ID")
    df['_topic'] = get_series(df, "Issue", "Topic")
    df['_category'] = get_series(df, "Case Category", "Category")
    df['_manager'] = get_series(df, "Manager Login", "Manager")
    df['_missed'] = pd.to_numeric(
        get_series(df, "# of Missed Parameters", "Missed Parameters"),
        errors='coerce').fillna(0).astype(int)
    df['_accuracy'] = parse_accuracy_series(get_series(df, "Accuracy %", "Accuracy"))
    df['_date'] = parse_date_series(get_series(df, "Audit Date", "Date"))
    df['_week'] = df['_date'].apply(iso_week)

    param_comment_cols = {}
    for i, p_name in enumerate(PARAMETERS):
        if i < len(param_col_positions):
            comment_pos = param_col_positions[i] + 1
            if comment_pos < len(df.columns):
                param_comment_cols[p_name] = df.columns[comment_pos]

    fatal_inap, fatal_conf, nonfatal, hits_list, fatal_comments = [], [], [], [], []
    for _, row in df.iterrows():
        hits = []
        is_inap = is_conf = False
        comments_for_row = []
        for p in PARAMETERS:
            c = param_cols.get(p)
            if c is None:
                continue
            if is_defect(row.get(c, '')):
                hits.append(p)
                if "appropriateresolution" in norm(p):
                    is_inap = True
                if "confidentiality" in norm(p):
                    is_conf = True
                ccol = param_comment_cols.get(p)
                if ccol:
                    cval = str(row.get(ccol, '')).strip()
                    if cval and cval.lower() != 'nan':
                        comments_for_row.append(cval)
        hits_list.append(hits)
        fatal_inap.append(is_inap)
        fatal_conf.append(is_conf)
        nonfatal.append(len(hits) > 0 and not (is_inap or is_conf))
        fatal_comments.append(" | ".join(comments_for_row) if comments_for_row else "")

    df['_param_hits'] = hits_list
    df['_fatal_inap'] = fatal_inap
    df['_fatal_conf'] = fatal_conf
    df['_is_fatal'] = pd.Series(fatal_inap, index=df.index) | pd.Series(fatal_conf, index=df.index)
    df['_is_nonfatal'] = nonfatal
    df['_has_any_def'] = df['_param_hits'].apply(lambda x: len(x) > 0)
    df['_comment'] = pd.Series(fatal_comments, index=df.index)

    return df


# ============================================================
# DISPUTES / RECTIFICATION / IVOC / DEFECT REDUCTION PARSERS
# ============================================================
def parse_tracker_file(filepath, sla_days=7, tracker_type='disputes'):
    if not filepath.exists():
        return pd.DataFrame()
    try:
        raw = pd.read_csv(filepath, header=None, encoding='utf-8-sig',
                          on_bad_lines='skip', dtype=str)
    except Exception:
        return pd.DataFrame()

    header_idx = 0
    for i in range(min(20, len(raw))):
        row = raw.iloc[i]
        normalized = [norm(v) for v in row.values if pd.notna(v) and str(v).strip()]
        if len(normalized) >= 3 and any('case' in c or 'analyst' in c for c in normalized):
            header_idx = i
            break

    df = pd.read_csv(filepath, header=header_idx, encoding='utf-8-sig',
                     on_bad_lines='skip', dtype=str)
    df.columns = [norm(c) for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated(keep='first')]
    df = df.dropna(how='all')

    case_col = find_col_strict(df, "Case Number", "Case No", "Case ID")
    if case_col:
        df = df[df[case_col].astype(str).str.strip().str.match(r'^\d+(-\d+)?$', na=False)].copy()

    if len(df) == 0:
        return pd.DataFrame()

    df['_case'] = get_series(df, "Case Number", "Case No", "Case ID")
    df['_org'] = get_series(df, "ORG", "Org").replace('', 'Unknown')
    df['_analyst'] = get_series(df, "Analyst", "Primary Analyst")
    df['_audit_date'] = parse_date_series(get_series(df, "Audit Date", "Paste Date"))

    if tracker_type == 'disputes':
        df['_owner'] = get_series(df, "Defect Owner", "Owner").str.lower()
        df['_dispute_date'] = parse_date_series(get_series(df, "Day of Dispute", "Dispute Date"))
        def _cat(row):
            owner = str(row.get('_owner', '')).lower()
            if 'auditor' in owner: return "QC Error (Auditor)"
            if 'quest lead' in owner or 'questlead' in owner: return "To Quest Lead"
            if 'backup' in owner or 'back up' in owner: return "Moved to Backup"
            if 'primary' in owner: return "Stayed with Primary"
            return "Other"
        df['_status'] = df.apply(_cat, axis=1)
        def _sla(row):
            ad, dd = row.get('_audit_date'), row.get('_dispute_date')
            try:
                if dd is None or ad is None or pd.isna(dd) or pd.isna(ad):
                    return "Pending"
            except: return "Pending"
            wd = workdays_between(ad, dd)
            if wd is None: return "Pending"
            return "SLA Met" if wd <= sla_days else "SLA Breached"
        df['_sla'] = df.apply(_sla, axis=1)
    else:
        df['_action'] = get_series(df, "Action Taken", "Action", "Feedback Provided")
        df['_end_date'] = parse_date_series(get_series(df, "Rectification Date", "Closed Date", "Date", "IVOC Rectification Date"))
        df['_status_raw'] = get_series(df, "Status", "Rectification Status", "IVOC Status")
        def _status(row):
            end_date = row.get('_end_date')
            action = str(row.get('_action', '')).lower().strip()
            status_raw = str(row.get('_status_raw', '')).lower().strip()
            try:
                if end_date is not None and pd.notna(end_date): return "Rectified"
            except: pass
            if 'rectif' in status_raw and 'not' not in status_raw: return "Rectified"
            if action and action not in ('na', 'n/a', 'none', 'pending', 'no action', ''): return "Rectified"
            audit_date = row.get('_audit_date')
            try:
                if audit_date is not None and pd.notna(audit_date):
                    wd = workdays_between(audit_date, datetime.now())
                    if wd is not None and wd > sla_days: return "Pending — Missed SLA"
            except: pass
            return "Pending — Within SLA"
        df['_status'] = df.apply(_status, axis=1)
        def _sla(row):
            s = row.get('_status', '')
            if s == "Rectified": return "SLA Met"
            if "Missed SLA" in s: return "SLA Breached"
            return "Pending"
        df['_sla'] = df.apply(_sla, axis=1)

    return df


def find_file(files, *keywords):
    for name, path in files.items():
        name_lower = name.lower()
        if all(kw.lower() in name_lower for kw in keywords):
            return Path(path)
    return Path("nonexistent")


# ============================================================
# HTML COMPONENTS (matching original dashboard)
# ============================================================
def render_kpi_row(total_audited, accuracy_avg, fatal_count, inap_count, conf_count, defects_total):
    defect_rate = (defects_total / total_audited * 100) if total_audited else 0
    acc_rag = rag_class(accuracy_avg)
    def_rag = rag_class(defect_rate, inverted=True)

    html = f'''
    <div class="kpi-row">
      <div class="kpi-card">
        <div class="kpi-label">Cases Audited</div>
        <div class="kpi-value">{total_audited}</div>
        <div class="kpi-sub">From Process file</div>
      </div>
      <div class="kpi-card {acc_rag}">
        <div class="kpi-label">Accuracy</div>
        <div class="kpi-value">{fmt_pct(accuracy_avg)}</div>
        <div class="kpi-sub">Average across all audits</div>
      </div>
      <div class="kpi-card kpi-fatal">
        <div class="kpi-label">Fatal Errors</div>
        <div class="kpi-value">{fatal_count}</div>
        <div class="kpi-sub">Inappropriate Resolution: {inap_count} &nbsp;|&nbsp; Confidentiality: {conf_count}</div>
      </div>
      <div class="kpi-card {def_rag}">
        <div class="kpi-label">Defects</div>
        <div class="kpi-value">{defects_total}</div>
        <div class="kpi-sub">{fmt_pct(defect_rate)} defect rate</div>
      </div>
    </div>'''
    st.markdown(html, unsafe_allow_html=True)


def render_status_cards(disputes_df, rect_df, ivoc_df, defred_df):
    def _card(title, df):
        if df is None or len(df) == 0:
            return f'''<div class="status-card">
              <div class="status-header">{title}</div>
              <div class="status-total" style="color:#9ca3af;">No data</div>
            </div>'''
        total = len(df)
        breakdown = df['_status'].value_counts()
        sla_met = (df['_sla'] == 'SLA Met').sum()
        sla_total = total
        sla_pct = round(sla_met / sla_total * 100, 1) if sla_total else 0
        sla_cls = 'sla-green' if sla_pct >= 100 else 'sla-red'

        items = ''
        for label, count in breakdown.items():
            items += f'<div class="status-item"><span>{label}</span><strong>{count}</strong></div>'

        return f'''<div class="status-card">
          <div class="status-header">{title}</div>
          <div class="status-total">Total: <strong>{total}</strong></div>
          {items}
          <div class="status-sla {sla_cls}">SLA: {sla_met}/{sla_total} ({fmt_pct(sla_pct)})</div>
        </div>'''

    html = '<div class="status-row-block">'
    html += _card('Disputes', disputes_df if len(disputes_df) > 0 else None)
    html += _card('Rectification', rect_df if len(rect_df) > 0 else None)
    html += _card('IVOC', ivoc_df if len(ivoc_df) > 0 else None)
    html += _card('Defect Reduction', defred_df if len(defred_df) > 0 else None)
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)


def render_wow(filtered_df):
    if len(filtered_df) == 0:
        return
    
    weeks = sorted([w for w in filtered_df['_week'].unique() if w],
                   key=lambda w: int(w.replace('W', '')) if w.startswith('W') else 0)
    if not weeks:
        st.markdown('<p style="color:#6b7280;font-size:12px;">No weekly data.</p>', unsafe_allow_html=True)
        return

    wow_data = []
    prev_fatal = prev_nonfatal = 0
    for w in weeks:
        g = filtered_df[filtered_df['_week'] == w]
        f_count = int(g['_is_fatal'].sum())
        n_count = int(g['_is_nonfatal'].sum())
        f_arrow = 'up' if f_count > prev_fatal else ('down' if f_count < prev_fatal else '')
        n_arrow = 'up' if n_count > prev_nonfatal else ('down' if n_count < prev_nonfatal else '')
        wow_data.append({'week': w, 'audited': len(g), 'fatal': f_count, 'nonfatal': n_count,
                         'f_arrow': f_arrow, 'n_arrow': n_arrow})
        prev_fatal, prev_nonfatal = f_count, n_count

    for metric_key, metric_label, color in [
        ('audited', 'Cases Audited', '#3b82f6'),
        ('fatal', 'Fatal Defects', '#dc2626'),
        ('nonfatal', 'Non-Fatal Defects', '#f59e0b'),
    ]:
        max_val = max(w[metric_key] for w in wow_data) or 1
        bars = ''
        for w in wow_data:
            pct = w[metric_key] / max_val * 100
            arrow_html = ''
            if metric_key == 'fatal' and w['f_arrow']:
                arrow_html = f'<span class="arr-{"up" if w["f_arrow"]=="up" else "dn"}">{"↑" if w["f_arrow"]=="up" else "↓"}</span>'
            elif metric_key == 'nonfatal' and w['n_arrow']:
                arrow_html = f'<span class="arr-{"up" if w["n_arrow"]=="up" else "dn"}">{"↑" if w["n_arrow"]=="up" else "↓"}</span>'
            bars += f'''<div class="wow-bar-col">
              <div class="wow-num">{w[metric_key]} {arrow_html}</div>
              <div class="wow-bar"><div class="wow-fill" style="height:{pct}%;background:{color}"></div></div>
              <div class="wow-lbl">{w['week']}</div>
            </div>'''
        st.markdown(f'''<div class="wow-block">
          <div class="wow-title">{metric_label}</div>
          <div class="wow-bars">{bars}</div>
        </div>''', unsafe_allow_html=True)


def render_mom(filtered_df, prev_df):
    if prev_df is None or len(prev_df) == 0:
        st.markdown('<p style="color:#6b7280;font-size:12px;">Previous month data not available.</p>', unsafe_allow_html=True)
        return

    cur_aud = len(filtered_df)
    cur_fat = int(filtered_df['_is_fatal'].sum())
    cur_acc = round(filtered_df['_accuracy'].mean(), 1) if cur_aud else 0
    prev_aud = len(prev_df)
    prev_fat = int(prev_df['_is_fatal'].sum())
    prev_acc = round(prev_df['_accuracy'].mean(), 1) if prev_aud else 0

    def _delta(cur, prev, better_lower=False):
        d = cur - prev
        if d == 0: return '<span style="color:#6b7280;font-size:11px;">no change</span>'
        if better_lower:
            cls = 'arr-dn' if d < 0 else 'arr-up'
            arrow = '↓' if d < 0 else '↑'
        else:
            cls = 'arr-up' if d > 0 else 'arr-dn'
            arrow = '↑' if d > 0 else '↓'
        return f'<span class="{cls}">{arrow} {abs(d)}</span>'

    html = f'''<table class="mom-table">
      <tr><th></th><th>Previous</th><th></th><th>Current</th></tr>
      <tr><td class="mom-lbl">Total Audited</td><td class="mom-val">{prev_aud}</td><td class="vs">vs</td>
          <td class="mom-val">{cur_aud}<div class="mom-delta">{_delta(cur_aud, prev_aud)}</div></td></tr>
      <tr><td class="mom-lbl">Fatal Errors</td><td class="mom-val">{prev_fat}</td><td class="vs">vs</td>
          <td class="mom-val">{cur_fat}<div class="mom-delta">{_delta(cur_fat, prev_fat, better_lower=True)}</div></td></tr>
      <tr><td class="mom-lbl">Accuracy</td><td class="mom-val">{fmt_pct(prev_acc)}</td><td class="vs">vs</td>
          <td class="mom-val">{fmt_pct(cur_acc)}<div class="mom-delta">{_delta(round(cur_acc-prev_acc,1), 0)}%</div></td></tr>
    </table>'''
    st.markdown(html, unsafe_allow_html=True)


def render_top_defects(filtered_df):
    counter = Counter()
    for hits in filtered_df['_param_hits']:
        for p in hits:
            counter[p] += 1
    if not counter:
        st.markdown('<p style="color:#6b7280;font-size:12px;">No defects recorded.</p>', unsafe_allow_html=True)
        return

    top = counter.most_common()
    max_c = top[0][1] if top else 1
    html = '<div class="bar-chart">'
    for name, count in top:
        pct = (count / max_c * 100) if max_c else 0
        html += f'''<div class="bar-row">
          <div class="bar-label">{name}</div>
          <div class="bar-track"><div class="bar-fill" style="width:{pct}%"></div></div>
          <div class="bar-value">{count}</div>
        </div>'''
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)


def render_org_scorecard(filtered_df):
    org_data = []
    for org, g in filtered_df.groupby('_org'):
        if not org or org == 'Unknown':
            continue
        oa = len(g)
        of = int(g['_is_fatal'].sum())
        on = int(g['_is_nonfatal'].sum())
        acc = g['_accuracy'].mean()
        dr = (of + on) / oa * 100 if oa else 0
        org_data.append({'org': org, 'audited': oa, 'accuracy': acc, 'fatal': of, 'nonfatal': on, 'defect_rate': dr})

    if not org_data:
        return

    html = '''<table class="data-table">
      <thead><tr><th>Org</th><th>Cases Audited</th><th>Accuracy</th><th>Fatal</th><th>Non-Fatal</th><th>Defect Rate</th></tr></thead><tbody>'''
    for o in sorted(org_data, key=lambda x: x['accuracy']):
        acc_cls = rag_text_class(o['accuracy'])
        dr_cls = rag_text_class(o['defect_rate'], inverted=True)
        html += f'''<tr>
          <td><strong>{o['org']}</strong></td>
          <td>{o['audited']}</td>
          <td class="{acc_cls}">{fmt_pct(o['accuracy'])}</td>
          <td>{o['fatal']}</td>
          <td>{o['nonfatal']}</td>
          <td class="{dr_cls}">{fmt_pct(o['defect_rate'])}</td>
        </tr>'''
    html += '</tbody></table>'
    st.markdown(html, unsafe_allow_html=True)


def render_fatal_by_analyst(analyst_df):
    if len(analyst_df) == 0:
        st.markdown('<p style="color:#6b7280;font-size:12px;">No fatal errors in Analyst file.</p>', unsafe_allow_html=True)
        return

    fatal_counts = {}
    for _, row in analyst_df.iterrows():
        if row['_is_fatal']:
            a = row['_analyst']
            if a:
                fatal_counts[a] = fatal_counts.get(a, 0) + 1

    if not fatal_counts:
        st.markdown('<p style="color:#6b7280;font-size:12px;">No fatal errors found.</p>', unsafe_allow_html=True)
        return

    sorted_fc = sorted(fatal_counts.items(), key=lambda x: x[1], reverse=True)[:10]
    max_c = sorted_fc[0][1] if sorted_fc else 1
    html = '<div class="bar-chart">'
    for analyst, count in sorted_fc:
        pct = (count / max_c * 100) if max_c else 0
        html += f'''<div class="bar-row">
          <div class="bar-label">{analyst}</div>
          <div class="bar-track"><div class="bar-fill bar-fill-fatal" style="width:{pct}%"></div></div>
          <div class="bar-value">{count}</div>
        </div>'''
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)


def render_analyst_table(analyst_df):
    if len(analyst_df) == 0:
        return

    perf = []
    for analyst, g in analyst_df.groupby('_analyst'):
        if not analyst:
            continue
        aa = len(g)
        af = int(g['_is_fatal'].sum())
        an = int(g['_is_nonfatal'].sum())
        acc = g['_accuracy'].mean()
        perf.append({'analyst': analyst, 'org': g['_org'].mode().iloc[0] if not g['_org'].mode().empty else '',
                     'audited': aa, 'accuracy': acc, 'fatal': af, 'nonfatal': an})

    if not perf:
        return

    html = '''<table class="data-table">
      <thead><tr><th>Analyst</th><th>Org</th><th>Audited</th><th>Accuracy</th><th>Fatal</th><th>Non-Fatal</th></tr></thead><tbody>'''
    for a in sorted(perf, key=lambda x: x['accuracy']):
        acc_cls = rag_text_class(a['accuracy'])
        html += f'''<tr>
          <td>{a['analyst']}</td>
          <td>{a['org']}</td>
          <td>{a['audited']}</td>
          <td class="{acc_cls}">{fmt_pct(a['accuracy'])}</td>
          <td>{a['fatal']}</td>
          <td>{a['nonfatal']}</td>
        </tr>'''
    html += '</tbody></table>'
    st.markdown(html, unsafe_allow_html=True)


# ============================================================
# MAIN DASHBOARD
# ============================================================
def main():
    inject_css()

    # Load data
    files, source, prefix = load_data_from_s3()
    prev_files = load_prev_month_from_s3()

    if not files:
        st.error("❌ No data files found!")
        st.info("Upload CSV files to S3 bucket `fincom-qc-data` or place them in the `data/` folder.")
        return

    # Parse files
    process_df = parse_audit_file(find_file(files, "process"))
    analyst_df = parse_audit_file(find_file(files, "analyst"))
    disputes_df = parse_tracker_file(find_file(files, "dispute"), sla_days=7, tracker_type='disputes')
    rect_df = parse_tracker_file(find_file(files, "rectification"), sla_days=5, tracker_type='rect')
    ivoc_df = parse_tracker_file(find_file(files, "ivoc"), sla_days=5, tracker_type='ivoc')
    defred_df = parse_tracker_file(find_file(files, "defect", "reduction"), sla_days=5, tracker_type='defred')
    prev_process_df = parse_audit_file(find_file(prev_files, "process")) if prev_files else pd.DataFrame()

    if len(process_df) == 0:
        st.error("❌ Could not parse Fincom_Process.csv!")
        st.info(f"Data source: {source} | Files: {list(files.keys())}")
        return

    # ── TITLE + DATA SOURCE ──
    st.markdown(f'''
    <div class="dashboard-title">📊 FinCom QC Dashboard — General Apr 2026</div>
    <div class="dashboard-sub">Data: {source} | {len(files)} files loaded | Last refreshed: {datetime.now().strftime("%Y-%m-%d %H:%M")}</div>
    ''', unsafe_allow_html=True)

    # ── FILTER BAR ──
    orgs = sorted([o for o in process_df['_org'].unique() if o and o != 'Unknown'])
    analysts = sorted([a for a in process_df['_analyst'].unique() if a])
    categories = sorted([c for c in process_df['_category'].unique() if c])
    weeks = sorted([w for w in process_df['_week'].unique() if w],
                   key=lambda w: int(w.replace('W', '')) if w.startswith('W') else 0)
    topics = sorted([t for t in process_df['_topic'].unique() if t])

    with st.container():
        fc1, fc2, fc3, fc4, fc5 = st.columns(5)
        with fc1:
            sel_orgs = st.multiselect("Org", orgs, default=orgs, key="f_org")
        with fc2:
            sel_analysts = st.multiselect("Analyst", analysts, default=analysts, key="f_analyst")
        with fc3:
            sel_categories = st.multiselect("Category", categories, default=categories, key="f_cat")
        with fc4:
            sel_weeks = st.multiselect("Week", weeks, default=weeks, key="f_week")
        with fc5:
            sel_topics = st.multiselect("Topic", topics, default=topics, key="f_topic")

    # Apply filters
    mask = pd.Series(True, index=process_df.index)
    if sel_orgs:
        mask &= process_df['_org'].isin(sel_orgs)
    if sel_analysts:
        mask &= process_df['_analyst'].isin(sel_analysts)
    if sel_categories:
        mask &= process_df['_category'].isin(sel_categories)
    if sel_weeks:
        mask &= process_df['_week'].isin(sel_weeks)
    if sel_topics:
        mask &= process_df['_topic'].isin(sel_topics)
    filtered_df = process_df[mask]

    # Filter analyst df too
    if len(analyst_df) > 0:
        a_mask = pd.Series(True, index=analyst_df.index)
        if sel_orgs:
            a_mask &= analyst_df['_org'].isin(sel_orgs)
        if sel_analysts:
            a_mask &= analyst_df['_analyst'].isin(sel_analysts)
        filtered_analyst = analyst_df[a_mask]
    else:
        filtered_analyst = pd.DataFrame()

    # ══════════════════════════════════════════════════════════
    # KPI ROW
    # ══════════════════════════════════════════════════════════
    total_audited = len(filtered_df)
    accuracy_avg = round(filtered_df['_accuracy'].mean(), 1) if total_audited else 0
    inap_count = int(filtered_df['_fatal_inap'].sum())
    conf_count = int(filtered_df['_fatal_conf'].sum())
    fatal_count = inap_count + conf_count
    defects_total = int(filtered_df['_missed'].sum())

    render_kpi_row(total_audited, accuracy_avg, fatal_count, inap_count, conf_count, defects_total)

    # ══════════════════════════════════════════════════════════
    # STATUS CARDS
    # ══════════════════════════════════════════════════════════
    st.markdown('<div class="section-title" style="margin-top:16px;">📋 Tracker Status</div>', unsafe_allow_html=True)
    render_status_cards(disputes_df, rect_df, ivoc_df, defred_df)

    # ══════════════════════════════════════════════════════════
    # WoW + MoM (side by side)
    # ══════════════════════════════════════════════════════════
    col_wow, col_mom = st.columns(2)
    with col_wow:
        st.markdown('<div class="section-box"><div class="section-title">📈 Week over Week</div>', unsafe_allow_html=True)
        render_wow(filtered_df)
        st.markdown('</div>', unsafe_allow_html=True)
    with col_mom:
        st.markdown('<div class="section-box"><div class="section-title">📊 Month over Month</div>', unsafe_allow_html=True)
        render_mom(filtered_df, prev_process_df)
        st.markdown('</div>', unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════
    # TOP DEFECT PARAMETERS
    # ══════════════════════════════════════════════════════════
    st.markdown('<div class="section-box"><div class="section-title">📊 Top Defect Parameters</div>', unsafe_allow_html=True)
    render_top_defects(filtered_df)
    st.markdown('</div>', unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════
    # ORG SCORECARD
    # ══════════════════════════════════════════════════════════
    st.markdown('<div class="section-box"><div class="section-title">🏢 Org Scorecard</div>', unsafe_allow_html=True)
    render_org_scorecard(filtered_df)
    st.markdown('</div>', unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════
    # FATAL BY ANALYST + ANALYST PERFORMANCE (side by side)
    # ══════════════════════════════════════════════════════════
    col_fba, col_apt = st.columns(2)
    with col_fba:
        st.markdown('<div class="section-box"><div class="section-title">⚠️ Fatal Errors by Analyst</div>', unsafe_allow_html=True)
        render_fatal_by_analyst(filtered_analyst)
        st.markdown('</div>', unsafe_allow_html=True)
    with col_apt:
        st.markdown('<div class="section-box"><div class="section-title">👤 Analyst Performance</div>', unsafe_allow_html=True)
        render_analyst_table(filtered_analyst)
        st.markdown('</div>', unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════
    # CASE-LEVEL DRILL-DOWN
    # ══════════════════════════════════════════════════════════
    with st.expander("🔎 Case-Level Detail (click to expand)"):
        if total_audited > 0:
            detail = filtered_df[['_case', '_analyst', '_org', '_category', '_topic', '_week',
                                  '_accuracy', '_missed', '_is_fatal', '_is_nonfatal', '_comment']].copy()
            detail.columns = ['Case', 'Analyst', 'Org', 'Category', 'Topic', 'Week',
                              'Accuracy', 'Missed', 'Fatal', 'Non-Fatal', 'Comment']
            st.dataframe(detail, use_container_width=True, hide_index=True, height=400)

    # ── SIDEBAR: Refresh button ──
    with st.sidebar:
        st.markdown("### ⚙️ Controls")
        if st.button("🔄 Refresh Data"):
            st.cache_data.clear()
            st.rerun()
        st.markdown("---")
        st.markdown(f"**Source:** {source}")
        st.markdown(f"**Files:** {len(files)}")
        for name in sorted(files.keys()):
            st.caption(f"📄 {name}")


if __name__ == "__main__":
    main()
