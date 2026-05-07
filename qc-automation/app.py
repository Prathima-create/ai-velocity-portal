"""
FinCom QC Automation Dashboard — Streamlit Web App
Hosted on EC2 behind corporate VPN.
Reads data from S3 bucket (synced from SharePoint).
"""

import streamlit as st
import pandas as pd
import json
import os
from pathlib import Path
from datetime import datetime
from collections import Counter

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
# DATA LOADING (S3 or Local fallback)
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


@st.cache_data(ttl=300)  # Cache for 5 minutes, then refresh from S3
def load_data_from_s3():
    """Try to load data from S3, fall back to local CSV."""
    try:
        import boto3
        s3 = boto3.client('s3')
        bucket = os.environ.get('QC_S3_BUCKET', 'fincom-qc-data')
        prefix = os.environ.get('QC_S3_PREFIX', 'current/')

        # List objects in the bucket
        response = s3.list_objects_v2(Bucket=bucket, Prefix=prefix)
        files = {}
        for obj in response.get('Contents', []):
            key = obj['Key']
            filename = key.split('/')[-1]
            if filename.endswith('.csv'):
                local_path = DATA_DIR / filename
                s3.download_file(bucket, key, str(local_path))
                files[filename] = local_path

        return files, "S3"
    except Exception as e:
        st.sidebar.warning(f"S3 not available: {e}\nUsing local data.")
        # Fall back to local files
        files = {}
        for f in DATA_DIR.glob("*.csv"):
            files[f.name] = f
        return files, "Local"


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


def parse_audit_file(filepath):
    """Parse a FinCom audit CSV file."""
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
        row_vals = [str(v) for v in raw.iloc[i].values if pd.notna(v)]
        normalized = [norm(v) for v in row_vals]
        has_analyst = any('analyst' in v for v in normalized)
        has_accuracy = any('accuracy' in v for v in normalized)
        has_date = any('auditdate' in v for v in normalized)
        if has_analyst and has_accuracy and has_date:
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
        if 'missedparameters' in c or 'ofmissed' in c:
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
        df = df[df[case_col].astype(str).str.strip().str.match(r'^\d+$', na=False)].copy()

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

    # Detect defects per row
    fatal_inap, fatal_conf, nonfatal, hits_list = [], [], [], []
    for _, row in df.iterrows():
        hits = []
        is_inap = is_conf = False
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
        hits_list.append(hits)
        fatal_inap.append(is_inap)
        fatal_conf.append(is_conf)
        nonfatal.append(len(hits) > 0 and not (is_inap or is_conf))

    df['_param_hits'] = hits_list
    df['_fatal_inap'] = fatal_inap
    df['_fatal_conf'] = fatal_conf
    df['_is_fatal'] = pd.Series(fatal_inap, index=df.index) | pd.Series(fatal_conf, index=df.index)
    df['_is_nonfatal'] = nonfatal
    df['_has_any_def'] = df['_param_hits'].apply(lambda x: len(x) > 0)

    return df


# ============================================================
# DASHBOARD UI
# ============================================================
def main():
    st.title("📊 FinCom QC Automation Dashboard")
    st.caption(f"Last refreshed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # Load data
    files, source = load_data_from_s3()

    if not files:
        st.error("❌ No data files found! Please ensure data is synced from SharePoint to S3.")
        st.info("""
        **Setup required:**
        1. Place CSV files in the `data/` folder, OR
        2. Configure S3 bucket sync (see deployment guide)
        
        Expected files:
        - `Fincom_Process.csv` — Process audit data
        - `Fincom_Analyst.csv` — Analyst audit data  
        - `Disputes.csv` — Dispute tracking
        - `Rectification.csv` — Rectification tracking
        - `IVOC.csv` — IVOC tracking
        - `Defect Reduction.csv` — Defect reduction tracking
        """)
        return

    st.sidebar.success(f"📡 Data source: **{source}**")
    st.sidebar.info(f"Files found: {len(files)}")

    # Parse process file
    process_file = None
    for name, path in files.items():
        if 'process' in name.lower():
            process_file = path
            break

    if process_file is None:
        st.error("❌ Fincom_Process.csv not found in data!")
        return

    process_df = parse_audit_file(Path(process_file))

    if len(process_df) == 0:
        st.error("❌ No valid data in Process file!")
        return

    # Parse analyst file
    analyst_file = None
    for name, path in files.items():
        if 'analyst' in name.lower():
            analyst_file = path
            break

    analyst_df = parse_audit_file(Path(analyst_file)) if analyst_file else pd.DataFrame()

    # ── SIDEBAR FILTERS ──
    st.sidebar.header("🔍 Filters")

    orgs = sorted(process_df['_org'].unique())
    selected_orgs = st.sidebar.multiselect("Org", orgs, default=orgs)

    analysts = sorted(process_df['_analyst'].unique())
    selected_analysts = st.sidebar.multiselect("Analyst", analysts, default=analysts)

    # Apply filters
    mask = process_df['_org'].isin(selected_orgs) & process_df['_analyst'].isin(selected_analysts)
    filtered_df = process_df[mask]

    # ── KPI ROW ──
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
        st.metric("🎯 Accuracy", f"{accuracy_avg:.1f}%",
                  delta=f"{'🟢' if accuracy_avg >= 95 else '🟡' if accuracy_avg >= 90 else '🔴'}")
    with col3:
        st.metric("⚠️ Fatal Errors", fatal_count,
                  help=f"Inappropriate Resolution: {inap_count} | Confidentiality: {conf_count}")
    with col4:
        st.metric("🔍 Total Defects", defects_total,
                  delta=f"{defects_total/total_audited*100:.1f}% rate" if total_audited else "0%",
                  delta_color="inverse")

    # ── TOP DEFECT PARAMETERS ──
    st.markdown("---")
    st.subheader("📊 Top Defect Parameters")

    counter = Counter()
    for hits in filtered_df['_param_hits']:
        for p in hits:
            counter[p] += 1

    if counter:
        defect_df = pd.DataFrame(counter.most_common(), columns=['Parameter', 'Count'])
        st.bar_chart(defect_df.set_index('Parameter'))
    else:
        st.info("No defects found with current filters.")

    # ── ORG SCORECARD ──
    st.markdown("---")
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
            'Cases Audited': oa,
            'Accuracy': f"{g['_accuracy'].mean():.1f}%",
            'Fatal': of,
            'Non-Fatal': on,
            'Defect Rate': f"{(of + on) / oa * 100:.1f}%" if oa else "0%",
        })

    if org_data:
        st.dataframe(pd.DataFrame(org_data), use_container_width=True, hide_index=True)

    # ── ANALYST PERFORMANCE ──
    st.markdown("---")
    col_left, col_right = st.columns(2)

    with col_left:
        st.subheader("👤 Fatal Errors by Analyst")
        if analyst_df is not None and len(analyst_df) > 0:
            analyst_mask = analyst_df['_analyst'].isin(selected_analysts) & analyst_df['_org'].isin(selected_orgs)
            filt_analyst = analyst_df[analyst_mask]
            filt_analyst_copy = filt_analyst.copy()
            filt_analyst_copy['_fatal_marks'] = filt_analyst_copy['_fatal_inap'].astype(int) + filt_analyst_copy['_fatal_conf'].astype(int)
            fatal_by_analyst = (filt_analyst_copy[filt_analyst_copy['_fatal_marks'] > 0]
                                .groupby('_analyst')['_fatal_marks'].sum()
                                .sort_values(ascending=False).head(10))
            if not fatal_by_analyst.empty:
                st.bar_chart(fatal_by_analyst)
            else:
                st.info("No fatal errors in analyst data.")
        else:
            st.info("Analyst file not loaded.")

    with col_right:
        st.subheader("📈 Analyst Performance Table")
        if analyst_df is not None and len(analyst_df) > 0:
            analyst_mask = analyst_df['_analyst'].isin(selected_analysts) & analyst_df['_org'].isin(selected_orgs)
            filt_analyst = analyst_df[analyst_mask]
            analyst_perf = []
            for analyst, g in filt_analyst.groupby('_analyst'):
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

    # ── DETAILED DRILL-DOWN ──
    st.markdown("---")
    st.subheader("🔎 Case-Level Detail")

    with st.expander("View all cases (click to expand)"):
        detail_cols = ['_case', '_analyst', '_org', '_accuracy', '_missed', '_is_fatal', '_is_nonfatal']
        display_df = filtered_df[detail_cols].copy()
        display_df.columns = ['Case', 'Analyst', 'Org', 'Accuracy', 'Missed Params', 'Fatal', 'Non-Fatal']
        st.dataframe(display_df, use_container_width=True, hide_index=True)

    # ── REFRESH BUTTON ──
    st.sidebar.markdown("---")
    if st.sidebar.button("🔄 Refresh Data from S3"):
        st.cache_data.clear()
        st.rerun()

    # ── DATA SYNC STATUS ──
    st.sidebar.markdown("---")
    st.sidebar.subheader("📡 Sync Status")
    sync_log = DATA_DIR / "sync_log.txt"
    if sync_log.exists():
        last_line = sync_log.read_text().strip().split('\n')[-1]
        st.sidebar.code(last_line, language=None)
    else:
        st.sidebar.warning("No sync log found. Run sync_sharepoint_to_s3.py")


if __name__ == "__main__":
    main()
