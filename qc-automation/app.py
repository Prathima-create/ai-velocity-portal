"""
FinCom QC Automation Dashboard — Streamlit Web App v2
=====================================================
Converted from qc_automation.py.py (HTML generator) to live Streamlit dashboard.
Reads data from S3 bucket (synced from SharePoint).
Data flow: SharePoint → S3 → EC2 Streamlit app

Key rules:
  - Process file = source for ALL org-level numbers
  - Analyst file = source for analyst-level rows ONLY
  - Defect cell rule: string == "0" => defect (letter "O" is NOT a defect)
  - Defects KPI = sum of "# of Missed Parameters" column
  - Fatal: "Appropriate Resolution" or "Confidentiality" == "0"
  - Non-Fatal Defect: any param == "0" but neither fatal param == "0"
"""

import streamlit as st
import pandas as pd
import os, re, warnings
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
    initial_sidebar_state="expanded"
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


# ============================================================
# DATA LOADING FROM S3
# ============================================================
@st.cache_data(ttl=300)
def load_data_from_s3():
    """Load CSV data from S3 bucket. Falls back to local data/ folder."""
    try:
        import boto3
        s3 = boto3.client('s3')
        bucket = os.environ.get('QC_S3_BUCKET', 'fincom-qc-data')

        # Try to load current month folder
        prefixes_to_try = [
            os.environ.get('QC_S3_PREFIX', 'General_Apr_2026/'),
            'current/',
        ]

        for prefix in prefixes_to_try:
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
        return {}, "S3 (empty)", ""
    except Exception as e:
        # Fall back to local files
        files = {}
        for f in DATA_DIR.glob("*.csv"):
            files[f.name] = f
        return files, f"Local ({e})" if not files else "Local", ""


@st.cache_data(ttl=300)
def load_prev_month_from_s3():
    """Load previous month data for MoM comparison."""
    try:
        import boto3
        s3 = boto3.client('s3')
        bucket = os.environ.get('QC_S3_BUCKET', 'fincom-qc-data')
        prefix = os.environ.get('QC_S3_PREV_PREFIX', 'General_Mar_2026/')

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
        return files
    except Exception:
        return {}


# ============================================================
# AUDIT FILE PARSER
# ============================================================
def parse_audit_file(filepath):
    """Parse a FinCom audit CSV file (Process or Analyst)."""
    if not filepath.exists():
        return pd.DataFrame()

    try:
        raw = pd.read_csv(filepath, header=None, encoding='utf-8-sig',
                          on_bad_lines='skip', dtype=str)
    except Exception:
        return pd.DataFrame()

    # Find header row
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

    # Find missed parameters column
    missed_idx = None
    for i, c in enumerate(df.columns):
        if 'missedparameters' in c or 'ofmissed' in c or c == 'missed':
            missed_idx = i
            break

    if missed_idx is None:
        missed_idx = len(df.columns)

    # Map parameter columns by position
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

    # Filter to numeric case numbers
    case_col = find_col_strict(df, "Case Number", "Case No", "Case ID")
    if case_col:
        df = df[df[case_col].astype(str).str.strip().str.match(r'^\d+(-\d+)?$', na=False)].copy()

    if len(df) == 0:
        return pd.DataFrame()

    # Extract fields
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

    # Detect defects per row
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
# DISPUTES PARSER
# ============================================================
def parse_disputes(filepath, sla_days=7):
    if not filepath.exists():
        return pd.DataFrame()

    try:
        raw = pd.read_csv(filepath, header=None, encoding='utf-8-sig',
                          on_bad_lines='skip', dtype=str)
    except Exception:
        return pd.DataFrame()

    # Find header
    header_idx = 0
    for i in range(min(20, len(raw))):
        row = raw.iloc[i]
        normalized = [norm(v) for v in row.values if pd.notna(v) and str(v).strip()]
        if len(normalized) >= 3 and any('case' in c for c in normalized):
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
    df['_owner'] = get_series(df, "Defect Owner", "Owner").str.lower()
    df['_org'] = get_series(df, "ORG", "Org").replace('', 'Unknown')
    df['_analyst'] = get_series(df, "Analyst", "Primary Analyst")

    df['_audit_date'] = parse_date_series(get_series(df, "Audit Date"))
    df['_dispute_date'] = parse_date_series(get_series(df, "Day of Dispute", "Dispute Date", "Date"))

    # Category
    def _category(row):
        owner = str(row.get('_owner', '')).lower()
        if 'auditor' in owner: return "QC Error (Auditor)"
        if 'quest lead' in owner or 'questlead' in owner: return "To Quest Lead"
        if 'backup' in owner or 'back up' in owner: return "Moved to Backup"
        if 'primary' in owner: return "Stayed with Primary"
        return "Other"
    df['_category'] = df.apply(_category, axis=1)

    # SLA
    def _sla(row):
        ad = row.get('_audit_date')
        dd = row.get('_dispute_date')
        try:
            if dd is None or ad is None or pd.isna(dd) or pd.isna(ad):
                return "Pending"
        except Exception:
            return "Pending"
        wd = workdays_between(ad, dd)
        if wd is None: return "Pending"
        return "SLA Met" if wd <= sla_days else "SLA Breached"
    df['_sla'] = df.apply(_sla, axis=1)

    return df


# ============================================================
# RECTIFICATION / IVOC PARSER
# ============================================================
def parse_rectification(filepath, sla_days=5):
    return _two_state_parser(filepath, sla_days,
        status_keys=["Rectification Status", "Status"],
        end_keys=["Rectification Date", "Closed Date", "Date"])


def parse_ivoc(filepath, sla_days=5):
    return _two_state_parser(filepath, sla_days,
        status_keys=["IVOC Status", "Rectification Status", "Status"],
        end_keys=["IVOC Rectification Date", "Rectification Date", "Closed Date", "Date"])


def _two_state_parser(filepath, sla_days, status_keys, end_keys):
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
        if len(normalized) >= 3 and any('case' in c for c in normalized):
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
    df['_action'] = get_series(df, "Action Taken", "Action", "Feedback Provided")
    df['_audit_date'] = parse_date_series(get_series(df, "Audit Date", "Paste Date"))
    df['_end_date'] = parse_date_series(get_series(df, *end_keys))
    df['_status_raw'] = get_series(df, *status_keys)

    def _status(row):
        end_date = row.get('_end_date')
        audit_date = row.get('_audit_date')
        action = str(row.get('_action', '')).lower().strip()
        status_raw = str(row.get('_status_raw', '')).lower().strip()

        try:
            if end_date is not None and pd.notna(end_date):
                return "Rectified"
        except Exception:
            pass
        if 'rectif' in status_raw and 'not' not in status_raw:
            return "Rectified"
        if action and action not in ('na', 'n/a', 'none', 'pending', 'no action', ''):
            return "Rectified"
        try:
            if audit_date is not None and pd.notna(audit_date):
                wd_today = workdays_between(audit_date, datetime.now())
                if wd_today is not None and wd_today > sla_days:
                    return "Pending — Missed SLA"
        except Exception:
            pass
        return "Pending — Within SLA"

    df['_status'] = df.apply(_status, axis=1)

    def _sla(row):
        s = row.get('_status', '')
        if s == "Rectified": return "SLA Met"
        if s == "Pending — Missed SLA": return "SLA Breached"
        return "Pending"
    df['_sla'] = df.apply(_sla, axis=1)

    return df


# ============================================================
# DEFECT REDUCTION PARSER
# ============================================================
def parse_defect_reduction(filepath, sla_days=5):
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
        if len(normalized) >= 3 and any('analyst' in c for c in normalized):
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
    df['_analyst'] = get_series(df, "Analyst")
    df['_action'] = get_series(df, "Action Taken", "Action", "Feedback Provided")
    df['_status_raw'] = get_series(df, "Status")
    df['_case_update_date'] = parse_date_series(get_series(df, "Case Update Date", "Update Date"))

    def _status(row):
        status = str(row.get('_status_raw', '')).lower().strip()
        if 'complet' in status or 'done' in status or 'closed' in status:
            return "Action Taken"
        cud = row.get('_case_update_date')
        try:
            if cud is not None and pd.notna(cud):
                wd_today = workdays_between(cud, datetime.now())
                if wd_today is not None and wd_today > sla_days:
                    return "Pending — Missed SLA"
        except Exception:
            pass
        return "Pending — Within SLA"

    df['_status'] = df.apply(_status, axis=1)

    def _sla(row):
        s = row.get('_status', '')
        if s == "Action Taken": return "SLA Met"
        if s == "Pending — Missed SLA": return "SLA Breached"
        return "Pending"
    df['_sla'] = df.apply(_sla, axis=1)

    return df


# ============================================================
# FIND FILE HELPER
# ============================================================
def find_file(files, *keywords):
    """Find a file by keyword in filename."""
    for name, path in files.items():
        name_lower = name.lower()
        if all(kw.lower() in name_lower for kw in keywords):
            return Path(path)
    return Path("nonexistent")


# ============================================================
# MAIN DASHBOARD
# ============================================================
def main():
    # Custom CSS
    st.markdown("""
    <style>
        .block-container { padding-top: 2rem; max-width: 1200px; }
        div[data-testid="stMetricValue"] { font-size: 1.8rem !important; font-weight: 700 !important; }
        div[data-testid="stMetricLabel"] { font-size: 0.75rem !important; text-transform: uppercase; }
        .stDataFrame { font-size: 0.85rem !important; }
        h1 { font-size: 1.5rem !important; }
        h2 { font-size: 1.2rem !important; }
        h3 { font-size: 1rem !important; }
    </style>
    """, unsafe_allow_html=True)

    st.title("📊 FinCom QC Dashboard")
    st.caption(f"Live dashboard • Last refreshed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # Load data
    files, source, prefix = load_data_from_s3()
    prev_files = load_prev_month_from_s3()

    if not files:
        st.error("❌ No data files found!")
        st.info("""
        **Setup required:** Upload CSV files to S3 bucket or place them in the `data/` folder.
        
        Expected files: `Fincom_Process.csv`, `Fincom_Analyst.csv`, `Disputes.csv`, 
        `Rectification.csv`, `IVOC.csv`, `Defect Reduction.csv`
        """)
        return

    st.sidebar.success(f"📡 Data: **{source}**")
    st.sidebar.info(f"Files: {len(files)} | Prefix: {prefix or 'local'}")

    # Parse files
    process_df = parse_audit_file(find_file(files, "process"))
    analyst_df = parse_audit_file(find_file(files, "analyst"))
    disputes_df = parse_disputes(find_file(files, "dispute"))
    rect_df = parse_rectification(find_file(files, "rectification"))
    ivoc_df = parse_ivoc(find_file(files, "ivoc"))
    defred_df = parse_defect_reduction(find_file(files, "defect", "reduction"))

    # Previous month
    prev_process_df = parse_audit_file(find_file(prev_files, "process")) if prev_files else pd.DataFrame()

    if len(process_df) == 0:
        st.error("❌ Could not parse Fincom_Process.csv!")
        st.info("Make sure the file has columns: Analyst, Accuracy %, Audit Date, Case Number")
        return

    # ── SIDEBAR FILTERS ──
    st.sidebar.header("🔍 Filters")

    orgs = sorted([o for o in process_df['_org'].unique() if o and o != 'Unknown'])
    selected_orgs = st.sidebar.multiselect("Org", orgs, default=orgs)

    analysts = sorted([a for a in process_df['_analyst'].unique() if a])
    selected_analysts = st.sidebar.multiselect("Analyst", analysts, default=analysts)

    weeks = sorted([w for w in process_df['_week'].unique() if w],
                   key=lambda w: int(w.replace('W', '')) if w and w.startswith('W') else 0)
    selected_weeks = st.sidebar.multiselect("Week", weeks, default=weeks)

    # Apply filters
    mask = (process_df['_org'].isin(selected_orgs) &
            process_df['_analyst'].isin(selected_analysts) &
            process_df['_week'].isin(selected_weeks))
    filtered_df = process_df[mask]

    # Also filter analyst df
    if len(analyst_df) > 0:
        analyst_mask = (analyst_df['_org'].isin(selected_orgs) &
                       analyst_df['_analyst'].isin(selected_analysts))
        filtered_analyst = analyst_df[analyst_mask]
    else:
        filtered_analyst = pd.DataFrame()

    # ══════════════════════════════════════════════════════════
    # KPI ROW
    # ══════════════════════════════════════════════════════════
    st.markdown("---")
    col1, col2, col3, col4 = st.columns(4)

    total_audited = len(filtered_df)
    accuracy_avg = round(filtered_df['_accuracy'].mean(), 2) if total_audited else 0
    inap_count = int(filtered_df['_fatal_inap'].sum())
    conf_count = int(filtered_df['_fatal_conf'].sum())
    fatal_count = inap_count + conf_count
    defects_total = int(filtered_df['_missed'].sum())

    with col1:
        st.metric("📋 Cases Audited", total_audited)
    with col2:
        color = "🟢" if accuracy_avg >= 95 else ("🟡" if accuracy_avg >= 90 else "🔴")
        st.metric("🎯 Accuracy", f"{accuracy_avg:.1f}%", delta=color)
    with col3:
        st.metric("⚠️ Fatal Errors", fatal_count,
                  help=f"Inap Resolution: {inap_count} | Confidentiality: {conf_count}")
    with col4:
        rate = f"{defects_total/total_audited*100:.1f}%" if total_audited else "0%"
        st.metric("🔍 Total Defects", defects_total, delta=f"Rate: {rate}", delta_color="inverse")

    # ══════════════════════════════════════════════════════════
    # STATUS CARDS ROW
    # ══════════════════════════════════════════════════════════
    st.markdown("---")
    st.subheader("📋 Tracker Status")
    sc1, sc2, sc3, sc4 = st.columns(4)

    def status_card(container, title, df, status_col='_status'):
        with container:
            st.markdown(f"**{title}**")
            if df is not None and len(df) > 0:
                st.caption(f"Total: {len(df)}")
                counts = df[status_col].value_counts()
                for status, count in counts.items():
                    st.write(f"• {status}: **{count}**")
                # SLA
                if '_sla' in df.columns:
                    sla_met = (df['_sla'] == 'SLA Met').sum()
                    sla_total = len(df)
                    sla_pct = round(sla_met / sla_total * 100, 1) if sla_total else 0
                    color = "🟢" if sla_pct >= 100 else "🔴"
                    st.caption(f"{color} SLA: {sla_met}/{sla_total} ({sla_pct}%)")
            else:
                st.caption("No data")

    status_card(sc1, "Disputes", disputes_df, '_category')
    status_card(sc2, "Rectification", rect_df, '_status')
    status_card(sc3, "IVOC", ivoc_df, '_status')
    status_card(sc4, "Defect Reduction", defred_df, '_status')

    # ══════════════════════════════════════════════════════════
    # WEEK OVER WEEK + MONTH OVER MONTH
    # ══════════════════════════════════════════════════════════
    st.markdown("---")
    wow_col, mom_col = st.columns(2)

    with wow_col:
        st.subheader("📈 Week over Week")
        if total_audited > 0:
            wow_data = []
            wk_groups = filtered_df.groupby('_week')
            for w in sorted([w for w in wk_groups.groups.keys() if w],
                           key=lambda w: int(w.replace('W', '')) if w.startswith('W') else 0):
                g = wk_groups.get_group(w)
                wow_data.append({
                    'Week': w,
                    'Audited': len(g),
                    'Fatal': int(g['_fatal_inap'].sum()) + int(g['_fatal_conf'].sum()),
                    'Non-Fatal': int(g['_is_nonfatal'].sum()),
                })
            if wow_data:
                wow_df = pd.DataFrame(wow_data)
                st.dataframe(wow_df, use_container_width=True, hide_index=True)
                st.bar_chart(wow_df.set_index('Week')[['Audited', 'Fatal', 'Non-Fatal']])
            else:
                st.info("No weekly data available.")
        else:
            st.info("No data for WoW.")

    with mom_col:
        st.subheader("📊 Month over Month")
        if len(prev_process_df) > 0:
            prev_aud = len(prev_process_df)
            prev_fat = int(prev_process_df['_fatal_inap'].sum()) + int(prev_process_df['_fatal_conf'].sum())
            prev_acc = round(prev_process_df['_accuracy'].mean(), 2)

            mom_data = {
                'Metric': ['Cases Audited', 'Fatal Errors', 'Accuracy'],
                'Previous': [prev_aud, prev_fat, f"{prev_acc:.1f}%"],
                'Current': [total_audited, fatal_count, f"{accuracy_avg:.1f}%"],
                'Delta': [total_audited - prev_aud, fatal_count - prev_fat,
                          f"{accuracy_avg - prev_acc:+.1f}%"],
            }
            st.dataframe(pd.DataFrame(mom_data), use_container_width=True, hide_index=True)
        else:
            st.info("Previous month data not available for MoM comparison.")

    # ══════════════════════════════════════════════════════════
    # TOP DEFECT PARAMETERS
    # ══════════════════════════════════════════════════════════
    st.markdown("---")
    st.subheader("📊 Top Defect Parameters")

    counter = Counter()
    for hits in filtered_df['_param_hits']:
        for p in hits:
            counter[p] += 1

    if counter:
        defect_data = pd.DataFrame(counter.most_common(), columns=['Parameter', 'Count'])
        st.bar_chart(defect_data.set_index('Parameter'))
    else:
        st.info("No defects found with current filters.")

    # ══════════════════════════════════════════════════════════
    # ORG SCORECARD + FATAL BY ANALYST
    # ══════════════════════════════════════════════════════════
    st.markdown("---")
    org_col, fba_col = st.columns(2)

    with org_col:
        st.subheader("🏢 Org Scorecard")
        org_data = []
        for org, g in filtered_df.groupby('_org'):
            if not org:
                continue
            oa = len(g)
            of = int(g['_fatal_inap'].sum()) + int(g['_fatal_conf'].sum())
            on = int(g['_is_nonfatal'].sum())
            org_data.append({
                'Org': org,
                'Audited': oa,
                'Accuracy': f"{g['_accuracy'].mean():.1f}%",
                'Fatal': of,
                'Non-Fatal': on,
                'Defect Rate': f"{(of + on) / oa * 100:.1f}%" if oa else "0%",
            })
        if org_data:
            st.dataframe(pd.DataFrame(org_data), use_container_width=True, hide_index=True)

    with fba_col:
        st.subheader("⚠️ Fatal Errors by Analyst")
        if len(filtered_analyst) > 0:
            fa_copy = filtered_analyst.copy()
            fa_copy['_fatal_marks'] = fa_copy['_fatal_inap'].astype(int) + fa_copy['_fatal_conf'].astype(int)
            fba = (fa_copy[fa_copy['_fatal_marks'] > 0]
                   .groupby('_analyst')['_fatal_marks'].sum()
                   .sort_values(ascending=False).head(10))
            if not fba.empty:
                st.bar_chart(fba)
            else:
                st.info("No fatal errors in analyst data.")
        else:
            st.info("Analyst file not loaded.")

    # ══════════════════════════════════════════════════════════
    # ANALYST PERFORMANCE TABLE
    # ══════════════════════════════════════════════════════════
    st.markdown("---")
    st.subheader("👤 Analyst Performance")

    if len(filtered_analyst) > 0:
        analyst_perf = []
        for analyst, g in filtered_analyst.groupby('_analyst'):
            if not analyst:
                continue
            aa = len(g)
            af = int(g['_fatal_inap'].sum()) + int(g['_fatal_conf'].sum())
            an = int(g['_is_nonfatal'].sum())
            analyst_perf.append({
                'Analyst': analyst,
                'Org': g['_org'].mode().iloc[0] if not g['_org'].mode().empty else '',
                'Audited': aa,
                'Accuracy': f"{g['_accuracy'].mean():.1f}%",
                'Fatal': af,
                'Non-Fatal': an,
            })
        if analyst_perf:
            st.dataframe(pd.DataFrame(analyst_perf).sort_values('Accuracy'),
                         use_container_width=True, hide_index=True)
    else:
        st.info("Analyst file not loaded.")

    # ══════════════════════════════════════════════════════════
    # CASE-LEVEL DETAIL
    # ══════════════════════════════════════════════════════════
    st.markdown("---")
    st.subheader("🔎 Case-Level Detail")

    with st.expander("View all cases (click to expand)"):
        detail_cols = ['_case', '_analyst', '_org', '_accuracy', '_missed', '_is_fatal', '_is_nonfatal']
        available_cols = [c for c in detail_cols if c in filtered_df.columns]
        display_df = filtered_df[available_cols].copy()
        col_map = {'_case': 'Case', '_analyst': 'Analyst', '_org': 'Org',
                   '_accuracy': 'Accuracy', '_missed': 'Missed Params',
                   '_is_fatal': 'Fatal', '_is_nonfatal': 'Non-Fatal'}
        display_df.columns = [col_map.get(c, c) for c in display_df.columns]
        st.dataframe(display_df, use_container_width=True, hide_index=True)

    # ══════════════════════════════════════════════════════════
    # SIDEBAR CONTROLS
    # ══════════════════════════════════════════════════════════
    st.sidebar.markdown("---")
    if st.sidebar.button("🔄 Refresh Data from S3"):
        st.cache_data.clear()
        st.rerun()

    st.sidebar.markdown("---")
    st.sidebar.subheader("📡 Data Files")
    for name in sorted(files.keys()):
        st.sidebar.caption(f"📄 {name}")


if __name__ == "__main__":
    main()
