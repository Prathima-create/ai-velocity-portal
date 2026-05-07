#!/usr/bin/env python3
"""
FinCom QC Automation Dashboard - v4
Outputs (per process+month folder):
  - QC_Dashboard_<folder>.html   (interactive dashboard, share via SharePoint)
  - QC_Report_<folder>.xlsx       (formal Excel report for leadership)
  - QC_Snapshot_<folder>.png      (screenshot to paste in email body)

Key rules:
  - Process file = source for ALL org-level numbers
  - Analyst file = source for analyst-level rows ONLY
  - Defect cell rule: string == "0" => defect (letter "O" is NOT a defect)
  - Defects KPI = sum of "# of Missed Parameters" column
  - Fatal: "Appropriate Resolution" or "Confidentiality" == "0"
  - Non-Fatal Defect: any param == "0" but neither fatal param == "0"
"""

import os, sys, json, base64, warnings
warnings.filterwarnings('ignore')
from pathlib import Path
from datetime import datetime, timedelta
from collections import Counter, defaultdict

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas not installed. Run: python -m pip install pandas --user")
    sys.exit(1)


# ============================================================
# CONSTANTS
# ============================================================
DESKTOP     = Path.home() / "Desktop"
BASE_FOLDER = DESKTOP / "FinCom_QC"

PARAMETERS = [
    "1. Greeting", "2. Empathy", "3. Language",
    "4. Appropriate Resolution", "5. Delay/Escalation",
    "6. Phone Calls", "7. Proactive/Self-Service",
    "8. Transfer/SIM CTI", "9. Confidentiality",
    "10. Annotations", "11. Resolution Code",
]
FATAL_NAMES = {"appropriateresolution", "confidentiality"}
MONTHS_ORDER = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]


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
    """Strict column lookup. Exact match, then prefix match (with non-alphanumeric boundary)."""
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
        for col in cols:
            if len(col) >= 3 and key.startswith(col):
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
            v = float(str(v).replace('%','').strip())
            return round(v * 100, 2) if v <= 1.0 else round(v, 2)
        except Exception:
            return None
    return series.apply(_p)


def parse_date_series(series):
    def _p(v):
        v = str(v).strip().split(' ')[0].split(',')[0]
        if not v or v.lower() in ('nan','none','','na','n/a'):
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
# CSV LOADER + CONFIG
# ============================================================
def load_csv_smart(filepath, required_keywords=None):
    if not filepath.exists():
        return None, f"File not found: {filepath.name}"
    try:
        raw = pd.read_csv(filepath, header=None, encoding='utf-8-sig',
                          on_bad_lines='skip', dtype=str)
    except Exception as e:
        return None, str(e)

    header_idx = 0
    if required_keywords:
        for i, row in raw.iterrows():
            normalized = [norm(v) for v in row.values if pd.notna(v)]
            if all(any(kw in cell for cell in normalized) for kw in required_keywords):
                header_idx = i
                break

    df = pd.read_csv(filepath, header=header_idx, encoding='utf-8-sig',
                     on_bad_lines='skip', dtype=str)
    df.columns = [norm(c) for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]
    df = df.dropna(how='all')
    return df, None


def load_config(folder):
    config_path = folder / "config.csv"
    defaults = {
        "disputes_sla_days":      5,
        "rectification_sla_days": 5,
        "ivoc_sla_days":          5,
        "reduction_sla_days":     5,
        "accuracy_target":        95.0,
    }
    if not config_path.exists():
        print("  WARNING: config.csv not found — using defaults")
        return defaults

    df, err = load_csv_smart(config_path)
    if err or df is None or len(df) == 0:
        return defaults

    row = df.iloc[0]
    def _get(*keys):
        for k in keys:
            col = find_col_strict(df, k)
            if col:
                try:
                    return float(str(row[col]).replace(',','').strip())
                except Exception:
                    pass
        return None

    cfg = dict(defaults)
    for k in cfg.keys():
        v = _get(k)
        if v is not None and v > 0:
            cfg[k] = v
    return cfg


# ============================================================
# AUDIT FILE PARSER (Process or Analyst)
#
# Real Fincom audit files have a complex structure:
#   - 60+ rows of preamble (definitions, instructions)
#   - True header row contains: Analyst, ORG, Audit Date, Accuracy %, # of Missed Parameters
#   - Parameter columns come AFTER "# of Missed Parameters" in pairs:
#     [point_value, Comment, point_value, Comment, ...]
#   - Header for parameter columns is the POINT VALUE (3, 8, 15, "O", 22, etc.)
#   - Cell value: point value = passed; "0" = defect; "O" = not applicable
#
# Strategy: find the real header row, then read 11 parameters by POSITION
# (alternating param-col, comment-col) starting after "# of Missed Parameters".
# ============================================================
def parse_audit_file(filepath, source_label):
    if not filepath.exists():
        print(f"  WARNING: {filepath.name} not found")
        return pd.DataFrame()

    print(f"\n  --- {filepath.name} ---")

    # Read raw file to find the header row
    try:
        raw = pd.read_csv(filepath, header=None, encoding='utf-8-sig',
                          on_bad_lines='skip', dtype=str)
    except Exception as e:
        print(f"  ERROR: Could not load: {e}")
        return pd.DataFrame()

    # Find the header row: must contain Analyst + Accuracy + Audit Date
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
        print("  ERROR: Could not find header row containing Analyst + Accuracy + Audit Date")
        return pd.DataFrame()

    print(f"  Header row found at line {header_row_idx + 1}")

    # Re-read with the correct header row
    df = pd.read_csv(filepath, header=header_row_idx, encoding='utf-8-sig',
                     on_bad_lines='skip', dtype=str)
    df = df.dropna(how='all')

    # Get raw column names BEFORE normalization (needed for positional logic)
    raw_cols = list(df.columns)
    df.columns = [norm(c) for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated(keep='first')]

    # Find the column index of "# of Missed Parameters" — params start right after
    missed_idx = None
    for i, c in enumerate(df.columns):
        if 'missedparameters' in c or 'ofmissed' in c or c == 'missed':
            missed_idx = i
            break

    if missed_idx is None:
        print("  WARNING: '# of Missed Parameters' column not found — parameter detection may fail")
        missed_idx = len(df.columns)  # fallback

    # The 11 parameters live at columns: missed_idx+1, +3, +5, +7, +9, +11, +13, +15, +17, +19, +21
    # (every other column after Missed; the columns in between are 'Comment' columns)
    param_col_positions = []
    for i in range(11):
        pos = missed_idx + 1 + (i * 2)
        if pos < len(df.columns):
            param_col_positions.append(pos)

    print(f"  Found Missed Params col at index {missed_idx}; reading 11 parameters at positions {param_col_positions}")

    # Map each parameter NAME to a real column NAME
    param_cols = {}
    for i, p_name in enumerate(PARAMETERS):
        if i < len(param_col_positions):
            param_cols[p_name] = df.columns[param_col_positions[i]]
        else:
            param_cols[p_name] = None

    found = sum(1 for c in param_cols.values() if c is not None)
    print(f"  Parameter columns matched: {found}/{len(PARAMETERS)}")
    for p, c in param_cols.items():
        marker = "OK " if c else "MISS"
        print(f"     {marker}  {p:32s} -> col '{c}'")

    # Filter rows: keep only those with a real Case Number (numeric)
    case_col = find_col_strict(df, "Case Number", "Case No", "Case ID")
    if case_col:
        df = df[df[case_col].astype(str).str.strip().str.match(r'^\d+$', na=False)].copy()
        print(f"  After filtering to numeric Case Numbers: {len(df)} rows")

    if len(df) == 0:
        print("  ERROR: No data rows found after filtering")
        return pd.DataFrame()

    # Basic fields
    df['_source']  = source_label
    df['_analyst'] = get_series(df, "Analyst")
    df['_org']     = get_series(df, "ORG", "Org").replace('', 'Unknown')
    df['_case']    = get_series(df, "Case Number", "Case No", "Case ID")
    df['_topic']   = get_series(df, "Issue", "Topic")
    df['_category']= get_series(df, "Case Category", "Category")
    df['_manager'] = get_series(df, "Manager Login", "Manager")
    # Comment column for fatal drill-downs — use the right-most "Comment" column near the end,
    # or fall back to any "Comment" column. Real file has many "Comment" columns (one per param).
    # Best match for a single audit comment: a column named "Audit Comment" or the last
    # standalone "Comment" column. We'll just use the get_series helper which finds the first.
    df['_comment'] = get_series(df, "Audit Comment", "Comment", "Auditor Remark", "Notes")

    missed = pd.to_numeric(get_series(df, "# of Missed Parameters", "Missed Parameters", "# of Missed", "Missed"),
                           errors='coerce').fillna(0).astype(int)
    df['_missed'] = missed

    df['_accuracy'] = parse_accuracy_series(get_series(df, "Accuracy %", "Accuracy"))

    df['_date'] = parse_date_series(get_series(df, "Audit Date", "Date"))
    df['_week'] = df['_date'].apply(iso_week)

    # Per-row defect detection (string equality)
    fatal_inap, fatal_conf, nonfatal, hits_list = [], [], [], []
    for _, row in df.iterrows():
        hits = []
        is_inap = is_conf = False
        for p in PARAMETERS:
            c = param_cols[p]
            if c is None:
                continue
            if is_defect(row[c]):
                hits.append(p)
                if "appropriateresolution" in norm(p):
                    is_inap = True
                if "confidentiality" in norm(p):
                    is_conf = True
        hits_list.append(hits)
        fatal_inap.append(is_inap)
        fatal_conf.append(is_conf)
        nonfatal.append(len(hits) > 0 and not (is_inap or is_conf))

    df['_param_hits']  = hits_list
    df['_fatal_inap']  = fatal_inap
    df['_fatal_conf']  = fatal_conf
    df['_is_fatal']    = pd.Series(fatal_inap, index=df.index) | pd.Series(fatal_conf, index=df.index)
    df['_is_nonfatal'] = nonfatal
    df['_has_any_def'] = df['_param_hits'].apply(lambda x: len(x) > 0)

    print(f"  {filepath.name}: {len(df)} records loaded")
    keep = ['_source','_analyst','_org','_case','_topic','_category','_manager',
            '_comment','_missed','_accuracy','_date','_week',
            '_param_hits','_fatal_inap','_fatal_conf',
            '_is_fatal','_is_nonfatal','_has_any_def']
    return df[keep].copy()


# ============================================================
# DISPUTES
# ============================================================
def parse_disputes(filepath, sla_days=5):
    if not filepath.exists():
        return pd.DataFrame()
    df, err = load_csv_smart(filepath, ["case"])
    if err or df is None:
        return pd.DataFrame()

    case_col = find_col_strict(df, "Case Number", "Case No", "Case ID")
    if case_col:
        df = df[df[case_col].astype(str).str.strip().str.match(r'^\d+$', na=False)].copy()

    df['_case']    = get_series(df, "Case Number", "Case No", "Case ID")
    df['_owner']   = get_series(df, "Defect Owner", "Owner").str.lower()
    df['_report']  = get_series(df, "Report Information", "Dispute Outcome", "Outcome").str.lower()
    df['_org']     = get_series(df, "ORG", "Org").replace('', 'Unknown')
    df['_analyst'] = get_series(df, "Analyst", "Primary Analyst")
    df['_manager'] = get_series(df, "Manager Login", "Manager")

    df['_audit_date']   = parse_date_series(get_series(df, "Audit Date"))
    df['_dispute_date'] = parse_date_series(get_series(df, "Day of Dispute", "Dispute Date", "Date"))
    df['_workdays'] = df.apply(lambda r: workdays_between(r['_audit_date'], r['_dispute_date']), axis=1)

    def _sla(row):
        if row['_dispute_date'] is None or row['_audit_date'] is None:
            return "Pending"
        wd = row['_workdays']
        if wd is None: return "Pending"
        return "SLA Met" if wd <= sla_days else "SLA Breached"
    df['_sla'] = df.apply(_sla, axis=1)

    def _category(row):
        owner = row['_owner']; report = row['_report']
        if 'auditor' in owner: return "QC Error (Auditor)"
        if 'quest lead' in owner or 'questlead' in owner: return "To Quest Lead"
        if 'backup' in owner or 'back up' in owner: return "Moved to Backup"
        if 'primary' in owner: return "Stayed with Primary"
        if 'reverse' in report: return "Reversed"
        if 'reject' in report: return "Rejected"
        return "Other"
    df['_category'] = df.apply(_category, axis=1)

    print(f"  Disputes: {len(df)} records loaded")
    return df


# ============================================================
# RECTIFICATION + IVOC (two-state)
# ============================================================
def _two_state_parser(filepath, label, sla_days, status_keys, end_keys):
    if not filepath.exists():
        return pd.DataFrame()
    df, err = load_csv_smart(filepath, ["case"])
    if err or df is None:
        return pd.DataFrame()

    case_col = find_col_strict(df, "Case Number", "Case No", "Case ID")
    if case_col:
        df = df[df[case_col].astype(str).str.strip().str.match(r'^\d+$', na=False)].copy()

    df['_case']    = get_series(df, "Case Number", "Case No", "Case ID")
    df['_org']     = get_series(df, "ORG", "Org").replace('', 'Unknown')
    df['_analyst'] = get_series(df, "Analyst", "Primary Analyst")
    df['_manager'] = get_series(df, "Manager Login", "Manager")
    df['_status_raw'] = get_series(df, *status_keys)
    df['_action']     = get_series(df, "Action Taken", "Action", "Feedback Provided")

    df['_audit_date'] = parse_date_series(get_series(df, "Audit Date", "Paste Date"))
    df['_end_date']   = parse_date_series(get_series(df, *end_keys))
    df['_workdays'] = df.apply(lambda r: workdays_between(r['_audit_date'], r['_end_date']), axis=1)

    def _status(row):
        action = row['_action'].lower(); status = row['_status_raw'].lower()
        if row['_end_date'] is not None: return "Rectified"
        if 'rectif' in status and 'not' not in status: return "Rectified"
        if action and action not in ('na','n/a','none','pending','no action',''): return "Rectified"
        return "Pending"
    df['_status'] = df.apply(_status, axis=1)

    def _sla(row):
        if row['_end_date'] is None:
            if row['_audit_date'] is not None:
                wd = workdays_between(row['_audit_date'], datetime.now())
                if wd is not None and wd > sla_days:
                    return "SLA Breached"
            return "Pending"
        if row['_workdays'] is None: return "Pending"
        return "SLA Met" if row['_workdays'] <= sla_days else "SLA Breached"
    df['_sla'] = df.apply(_sla, axis=1)

    print(f"  {label}: {len(df)} records loaded")
    return df


def parse_rectification(filepath, sla_days=5):
    return _two_state_parser(filepath, "Rectification", sla_days,
        status_keys=["Rectification Status", "Status"],
        end_keys=["Rectification Date", "Closed Date", "Date"])


def parse_ivoc(filepath, sla_days=5):
    return _two_state_parser(filepath, "IVOC", sla_days,
        status_keys=["IVOC Status", "Rectification Status", "Status"],
        end_keys=["IVOC Rectification Date", "Rectification Date", "Closed Date", "Date"])


# ============================================================
# DEFECT REDUCTION (filtered to current month)
# ============================================================
def parse_defect_reduction(filepath, current_month_short, sla_days=5):
    if not filepath.exists():
        return pd.DataFrame()
    df, err = load_csv_smart(filepath, ["analyst"])
    if err or df is None:
        return pd.DataFrame()

    df['_case']    = get_series(df, "Case Number", "Case No", "Case ID")
    df['_org']     = get_series(df, "ORG", "Org").replace('', 'Unknown')
    df['_analyst'] = get_series(df, "Analyst")
    df['_manager'] = get_series(df, "Manager Login", "Manager")
    df['_action']  = get_series(df, "Action Taken", "Action", "Feedback Provided")
    df['_topic']   = get_series(df, "Topic", "Defect Parameter", "Parameter")
    df['_month']   = get_series(df, "Month")

    target = current_month_short.strip().lower()[:3]
    if target:
        before = len(df)
        df = df[df['_month'].str.lower().str[:3] == target].copy()
        print(f"  Defect Reduction: filtered to {current_month_short} ({before} -> {len(df)})")

    df['_review_date'] = parse_date_series(get_series(df, "Date of Review", "Review Date"))
    df['_update_date'] = parse_date_series(get_series(df, "Case Update Date", "Update Date", "Closed Date"))
    df['_workdays']    = df.apply(lambda r: workdays_between(r['_review_date'], r['_update_date']), axis=1)

    def _status(row):
        action = row['_action'].lower()
        if action and action not in ('na','n/a','none','pending','no action',''):
            return "Action Taken"
        return "Pending"
    df['_status'] = df.apply(_status, axis=1)

    def _sla(row):
        if row['_update_date'] is None:
            if row['_review_date'] is not None:
                wd = workdays_between(row['_review_date'], datetime.now())
                if wd is not None and wd > sla_days:
                    return "SLA Breached"
            return "Pending"
        if row['_workdays'] is None: return "Pending"
        return "SLA Met" if row['_workdays'] <= sla_days else "SLA Breached"
    df['_sla'] = df.apply(_sla, axis=1)

    print(f"  Defect Reduction: {len(df)} records loaded")
    return df

# ============================================================
# METRICS COMPUTATION
# Process file -> all org-level numbers
# Analyst file -> analyst rows only
# ============================================================
def compute_metrics(process_df, analyst_df,
                    disputes_df, rect_df, ivoc_df, defred_df,
                    prev_process_df, prev_analyst_df, config,
                    process_label, month_label):
    """Build all dashboard numbers."""
    if process_df is None or len(process_df) == 0:
        return None

    m = {}

    # ---- HEADLINE KPIs (from Process file only) ----
    total_audited = len(process_df)
    accuracy_avg  = round(process_df['_accuracy'].mean(), 2) if total_audited else 0.0

    inap_count = int(process_df['_fatal_inap'].sum())
    conf_count = int(process_df['_fatal_conf'].sum())
    # Fatal Defects = simple sum (each fatal mark is its own defect, not deduped by case)
    fatal_count = inap_count + conf_count

    # Defects = SUM of "# of Missed Parameters" column
    defects_total = int(process_df['_missed'].sum())
    nonfatal_defects = int(process_df['_is_nonfatal'].sum())

    m['headline'] = {
        'audited':         total_audited,
        'accuracy':        accuracy_avg,
        'fatal':           fatal_count,
        'inap':            inap_count,
        'conf':            conf_count,
        'defects':         defects_total,
        'nonfatal_defects': nonfatal_defects,
        'defect_rate':     round(defects_total / total_audited * 100, 2) if total_audited else 0,
    }

    # ---- STATUS CARDS ----
    def _status_card(df, status_field):
        if df is None or len(df) == 0:
            return None
        breakdown = df[status_field].value_counts().to_dict() if status_field in df.columns else {}
        sla_counts = df['_sla'].value_counts().to_dict() if '_sla' in df.columns else {}
        sla_met = sla_counts.get('SLA Met', 0)
        sla_total = len(df)
        return {
            'total':     sla_total,
            'breakdown': breakdown,
            'sla_met':   sla_met,
            'sla_total': sla_total,
            'sla_pct':   round(sla_met / sla_total * 100, 2) if sla_total else 0,
        }

    m['disputes']         = _status_card(disputes_df, '_category')
    m['rectification']    = _status_card(rect_df, '_status')
    m['ivoc']             = _status_card(ivoc_df, '_status')
    m['defect_reduction'] = _status_card(defred_df, '_status')

    # ---- WoW (Cases Audited / Fatal Defects / Non-Fatal Defects per ISO week) ----
    wow_data = []
    if total_audited:
        wk_groups = process_df.groupby('_week')
        weeks_sorted = sorted(
            [w for w in wk_groups.groups.keys() if w and w != 'Unknown'],
            key=lambda w: int(w.replace('W', '')) if w.startswith('W') else 0
        )
        prev_fatal = None
        prev_nonfatal = None
        for w in weeks_sorted:
            g = wk_groups.get_group(w)
            audited = len(g)
            # Fatal = sum of fatal marks (Inap + Conf) for this week
            fatal_w = int(g['_fatal_inap'].sum()) + int(g['_fatal_conf'].sum())
            nonfatal_w = int(g['_is_nonfatal'].sum())
            wow_data.append({
                'week':      w,
                'audited':   audited,
                'fatal':     fatal_w,
                'nonfatal':  nonfatal_w,
                'fatal_delta_arrow':    None if prev_fatal is None else ('up' if fatal_w > prev_fatal else ('down' if fatal_w < prev_fatal else 'flat')),
                'nonfatal_delta_arrow': None if prev_nonfatal is None else ('up' if nonfatal_w > prev_nonfatal else ('down' if nonfatal_w < prev_nonfatal else 'flat')),
            })
            prev_fatal = fatal_w
            prev_nonfatal = nonfatal_w
    m['wow'] = wow_data

    # ---- MoM (current month vs previous month) ----
    mom = None
    if prev_process_df is not None and len(prev_process_df) > 0:
        cur_aud = total_audited
        cur_fat = fatal_count
        cur_proc_acc = accuracy_avg  # Process accuracy (current)
        cur_anl_acc = round(analyst_df['_accuracy'].mean(), 2) if analyst_df is not None and len(analyst_df) > 0 else None

        prv_aud = len(prev_process_df)
        prv_fat = int(prev_process_df['_fatal_inap'].sum()) + int(prev_process_df['_fatal_conf'].sum())
        prv_proc_acc = round(prev_process_df['_accuracy'].mean(), 2) if prv_aud else 0.0
        prv_anl_acc = round(prev_analyst_df['_accuracy'].mean(), 2) if prev_analyst_df is not None and len(prev_analyst_df) > 0 else None

        mom = {
            'prev_label': month_label.split()[0] if month_label else 'Prev',
            'cur_label':  month_label,
            'prev_audited': prv_aud, 'cur_audited': cur_aud,
            'prev_fatal':   prv_fat, 'cur_fatal':   cur_fat,
            'prev_proc_acc': prv_proc_acc, 'cur_proc_acc': cur_proc_acc,
            'prev_anl_acc':  prv_anl_acc,  'cur_anl_acc':  cur_anl_acc,
            'prev_fatal_pct': round(prv_fat / prv_aud * 100, 2) if prv_aud else 0,
            'cur_fatal_pct':  round(cur_fat / cur_aud * 100, 2) if cur_aud else 0,
            'fatal_delta':    cur_fat - prv_fat,
            'audited_delta':  cur_aud - prv_aud,
            'proc_acc_delta': round(cur_proc_acc - prv_proc_acc, 2),
            'anl_acc_delta':  round((cur_anl_acc or 0) - (prv_anl_acc or 0), 2) if (cur_anl_acc is not None and prv_anl_acc is not None) else None,
        }
    m['mom'] = mom

    # ---- TOP DEFECT PARAMETERS (all 11, ranked) ----
    counter = Counter()
    for hits in process_df['_param_hits']:
        for p in hits:
            counter[p] += 1
    m['top_defects'] = counter.most_common(len(PARAMETERS))

    # ---- ORG TABLE (sortable) ----
    org_rows = []
    for org, g in process_df.groupby('_org'):
        if not org:
            continue
        oa = len(g)
        # Fatal = simple sum of fatal marks (Inap + Conf), matching headline KPI
        of = int(g['_fatal_inap'].sum()) + int(g['_fatal_conf'].sum())
        # Non-fatal defects = cases with at least one non-fatal defect (no fatal mark)
        on = int(g['_is_nonfatal'].sum())
        org_rows.append({
            'org':         org,
            'audited':     oa,
            'accuracy':    round(g['_accuracy'].mean(), 2) if oa else 0,
            'fatal':       of,
            'nonfatal':    on,
            'defect_rate': round((of + on) / oa * 100, 2) if oa else 0,
        })
    m['orgs'] = sorted(org_rows, key=lambda x: -x['audited'])

    # ---- FATAL BY ANALYST (from ANALYST file only) ----
    if analyst_df is not None and len(analyst_df) > 0:
        # Sum of fatal marks (Inap + Conf) per analyst — matches KPI counting
        analyst_df_copy = analyst_df.copy()
        analyst_df_copy['_fatal_marks'] = analyst_df_copy['_fatal_inap'].astype(int) + analyst_df_copy['_fatal_conf'].astype(int)
        fba = (analyst_df_copy[analyst_df_copy['_fatal_marks'] > 0]
               .groupby('_analyst')['_fatal_marks'].sum()
               .sort_values(ascending=False).head(10))
        m['fatal_by_analyst'] = [(a, int(c)) for a, c in fba.items() if a]
    else:
        m['fatal_by_analyst'] = []

    # ---- ANALYST TABLE (from ANALYST file only) ----
    analysts = []
    if analyst_df is not None and len(analyst_df) > 0:
        for analyst, g in analyst_df.groupby('_analyst'):
            if not analyst:
                continue
            aa = len(g)
            # Fatal = sum of fatal marks (Inap + Conf)
            af = int(g['_fatal_inap'].sum()) + int(g['_fatal_conf'].sum())
            an = int(g['_is_nonfatal'].sum())
            analysts.append({
                'analyst':  analyst,
                'org':      g['_org'].mode().iloc[0] if not g['_org'].mode().empty else '',
                'manager':  g['_manager'].mode().iloc[0] if not g['_manager'].mode().empty else '',
                'audited':  aa,
                'accuracy': round(g['_accuracy'].mean(), 2),
                'fatal':    af,
                'nonfatal': an,
            })
    m['analysts'] = sorted(analysts, key=lambda x: x['accuracy'])

    # ---- FILTER VALUES ----
    m['filters'] = {
        'orgs':       sorted({o for o in process_df['_org'].unique() if o}),
        'analysts':   sorted({a for a in process_df['_analyst'].unique() if a}),
        'categories': sorted({c for c in process_df['_category'].unique() if c}),
        'weeks':      sorted({w for w in process_df['_week'].unique() if w},
                             key=lambda w: int(w.replace('W','')) if w and w.startswith('W') else 0),
        'topics':     sorted({t for t in process_df['_topic'].unique() if t}),
    }

    # ---- ROW-LEVEL DATA for client-side drill-downs ----
    rows = []
    for _, r in process_df.iterrows():
        rows.append({
            'case':        r['_case'],
            'analyst':     r['_analyst'],
            'org':         r['_org'],
            'manager':     r['_manager'],
            'category':    r['_category'],
            'topic':       r['_topic'],
            'week':        r['_week'] or '',
            'date':        r['_date'].strftime('%Y-%m-%d') if r['_date'] is not None else '',
            'accuracy':    r['_accuracy'],
            'is_fatal':    bool(r['_is_fatal']),
            'is_nonfatal': bool(r['_is_nonfatal']),
            'has_def':     bool(r['_has_any_def']),
            'missed':      int(r['_missed']),
            'params':      r['_param_hits'],
            'comment':     r['_comment'],
        })
    m['rows'] = rows
    return m


# ============================================================
# HTML DASHBOARD BUILDER
# ============================================================
def rag_class(pct, *, inverted=False, sla_card=False):
    if pct is None:
        return 'rag-na'
    if sla_card:
        return 'rag-green' if pct >= 100 else 'rag-red'
    if inverted:
        if pct <= 5:  return 'rag-green'
        if pct <= 10: return 'rag-amber'
        return 'rag-red'
    if pct >= 95: return 'rag-green'
    if pct >= 90: return 'rag-amber'
    return 'rag-red'


def fmt_pct(p):
    return '—' if p is None else f'{p:.1f}%'


def build_html(metrics, process_label, month_label, prev_month_label, missing_files):
    h = metrics['headline']

    # KPI ROW
    kpi_html = f'''
    <div class="row kpi-row">
      <div class="kpi-card" onclick="openDrill('audited')">
        <div class="kpi-label">Cases Audited</div>
        <div class="kpi-value">{h['audited']}</div>
        <div class="kpi-sub">From Process file</div>
      </div>
      <div class="kpi-card {rag_class(h['accuracy'])}" onclick="openDrill('accuracy')">
        <div class="kpi-label">Accuracy</div>
        <div class="kpi-value">{fmt_pct(h['accuracy'])}</div>
        <div class="kpi-sub">Average across all audits</div>
      </div>
      <div class="kpi-card kpi-fatal" onclick="openDrill('fatal')">
        <div class="kpi-label">Fatal Errors</div>
        <div class="kpi-value">{h['fatal']}</div>
        <div class="kpi-sub">Inappropriate Resolution: {h['inap']} &nbsp;|&nbsp; Confidentiality: {h['conf']}</div>
      </div>
      <div class="kpi-card {rag_class(h['defect_rate'], inverted=True)}" onclick="openDrill('defects')">
        <div class="kpi-label">Defects</div>
        <div class="kpi-value">{h['defects']}</div>
        <div class="kpi-sub">{fmt_pct(h['defect_rate'])} defect rate</div>
      </div>
    </div>'''

    # STATUS ROW
    def _status_html(title, data, key):
        if data is None:
            return ''
        rows_inner = ''
        for label, count in data['breakdown'].items():
            label_safe = str(label).replace("'", "&#39;")
            rows_inner += f'<div class="status-row" onclick="event.stopPropagation();openStatusDrill(\'{key}\',\'{label_safe}\')"><span>{label_safe}</span><strong>{count}</strong></div>'
        sla_cls = rag_class(data['sla_pct'], sla_card=True)
        return f'''
        <div class="status-card" onclick="openDrill('{key}')">
          <div class="status-header">{title}</div>
          <div class="status-total">Total: <strong>{data['total']}</strong></div>
          <div class="status-body">{rows_inner}</div>
          <div class="status-sla {sla_cls}">SLA: {data['sla_met']}/{data['sla_total']} ({fmt_pct(data['sla_pct'])})</div>
        </div>'''

    status_html = '<div class="row status-row-block">'
    status_html += _status_html('Disputes', metrics['disputes'], 'disputes')
    status_html += _status_html('Rectification', metrics['rectification'], 'rectification')
    status_html += _status_html('IVOC', metrics['ivoc'], 'ivoc')
    status_html += _status_html('Defect Reduction', metrics['defect_reduction'], 'defect_reduction')
    status_html += '</div>'

    # WoW
    wow_html = '<div class="section half"><h2>Week over Week</h2>'
    if metrics['wow']:
        for metric_key, metric_label, color in [
            ('audited', 'Cases Audited', '#3b82f6'),
            ('fatal', 'Fatal Defects', '#dc2626'),
            ('nonfatal', 'Non-Fatal Defects', '#f59e0b'),
        ]:
            max_val = max(w[metric_key] for w in metrics['wow']) or 1
            wow_html += f'<div class="wow-block"><div class="wow-title">{metric_label}</div><div class="wow-bars">'
            for w in metrics['wow']:
                pct = w[metric_key] / max_val * 100
                arrow = ''
                if metric_key in ('fatal','nonfatal'):
                    a = w.get(f'{metric_key}_delta_arrow')
                    if a == 'up':   arrow = '<span class="arr-up">↑</span>'
                    elif a == 'down': arrow = '<span class="arr-dn">↓</span>'
                wow_html += f'''<div class="wow-bar-col">
                  <div class="wow-num">{w[metric_key]} {arrow}</div>
                  <div class="wow-bar"><div class="wow-fill" style="height:{pct}%;background:{color}"></div></div>
                  <div class="wow-lbl">{w['week']}</div>
                </div>'''
            wow_html += '</div></div>'
    else:
        wow_html += '<p class="muted">No weekly data.</p>'
    wow_html += '</div>'

    # MoM
    mom_html = '<div class="section half"><h2>Month over Month</h2>'
    mom = metrics['mom']
    if mom:
        prev = prev_month_label or 'Previous'
        cur  = month_label
        def _delta(d, suffix='', better_is_lower=False):
            if d == 0:
                return '<span class="muted">no change</span>'
            if better_is_lower:
                cls = 'arr-dn' if d < 0 else 'arr-up'
                arrow = '↓' if d < 0 else '↑'
            else:
                cls = 'arr-up' if d > 0 else 'arr-dn'
                arrow = '↑' if d > 0 else '↓'
            return f'<span class="{cls}">{arrow} {abs(d)}{suffix} vs {prev}</span>'
        # Helper: format an accuracy MoM cell (handle None for missing files)
        def _acc_cell(prev_v, cur_v, delta_v):
            prev_str = fmt_pct(prev_v) if prev_v is not None else '—'
            cur_str  = fmt_pct(cur_v)  if cur_v  is not None else '—'
            delta_str = _delta(delta_v, '%') if delta_v is not None else ''
            return prev_str, cur_str, delta_str

        proc_prev, proc_cur, proc_delta = _acc_cell(mom['prev_proc_acc'], mom['cur_proc_acc'], mom['proc_acc_delta'])
        anl_prev,  anl_cur,  anl_delta  = _acc_cell(mom['prev_anl_acc'],  mom['cur_anl_acc'],  mom['anl_acc_delta'])

        mom_html += f'''
        <table class="mom-table">
          <tr><th></th><th>{prev}</th><th></th><th>{cur}</th></tr>
          <tr>
            <td class="mom-lbl">Total Audited</td>
            <td class="mom-val">{mom['prev_audited']}</td>
            <td class="vs">vs</td>
            <td class="mom-val">{mom['cur_audited']}<div class="mom-delta">{_delta(mom['audited_delta'])}</div></td>
          </tr>
          <tr>
            <td class="mom-lbl">Fatal</td>
            <td class="mom-val">{mom['prev_fatal']} ({mom['prev_fatal_pct']:.1f}%)</td>
            <td class="vs">vs</td>
            <td class="mom-val mom-fatal" onclick="openMoMDrill('fatal')">{mom['cur_fatal']} ({mom['cur_fatal_pct']:.1f}%)<div class="mom-delta">{_delta(mom['fatal_delta'], '', better_is_lower=True)}</div></td>
          </tr>
          <tr>
            <td class="mom-lbl">Process Accuracy</td>
            <td class="mom-val">{proc_prev}</td>
            <td class="vs">vs</td>
            <td class="mom-val" onclick="openMoMDrill('proc_accuracy')">{proc_cur}<div class="mom-delta">{proc_delta}</div></td>
          </tr>
          <tr>
            <td class="mom-lbl">Analyst Accuracy</td>
            <td class="mom-val">{anl_prev}</td>
            <td class="vs">vs</td>
            <td class="mom-val" onclick="openMoMDrill('anl_accuracy')">{anl_cur}<div class="mom-delta">{anl_delta}</div></td>
          </tr>
        </table>
        <p class="muted small">Click <strong>Fatal</strong> or <strong>Accuracy</strong> on the right to drill down.</p>
        '''
    else:
        mom_html += f'<p class="muted">Previous month folder not found. To enable MoM, ensure the previous month\'s data exists at <code>{BASE_FOLDER}\\{process_label}_&lt;PrevMonth&gt;_&lt;Year&gt;</code>.</p>'
    mom_html += '</div>'

    # TOP DEFECTS (all 11, ranked)
    top_html = '<div class="section"><h2>Top Defect Parameters</h2>'
    if metrics['top_defects']:
        max_c = metrics['top_defects'][0][1] if metrics['top_defects'] else 1
        top_html += '<div class="bar-chart">'
        for name, count in metrics['top_defects']:
            pct = (count / max_c * 100) if max_c else 0
            name_safe = str(name).replace("'", "&#39;")
            top_html += f'''
            <div class="bar-row" onclick="openParamDrill('{name_safe}')">
              <div class="bar-label">{name_safe}</div>
              <div class="bar-track"><div class="bar-fill" style="width:{pct}%"></div></div>
              <div class="bar-value">{count}</div>
            </div>'''
        top_html += '</div>'
    else:
        top_html += '<p class="muted">No defects recorded.</p>'
    top_html += '</div>'

    # ORG TABLE (sortable)
    org_html = '<div class="section"><h2>Org Scorecard</h2>'
    org_html += '''<table class="data-table" id="orgTable">
      <thead><tr>
        <th onclick="sortDataTable('orgTable',0,'str')">Org</th>
        <th onclick="sortDataTable('orgTable',1,'num')">Cases Audited</th>
        <th onclick="sortDataTable('orgTable',2,'num')">Accuracy</th>
        <th onclick="sortDataTable('orgTable',3,'num')">Fatal</th>
        <th onclick="sortDataTable('orgTable',4,'num')">Non-Fatal Defects</th>
        <th onclick="sortDataTable('orgTable',5,'num')">Defect Rate</th>
      </tr></thead><tbody>'''
    for o in metrics['orgs']:
        acc_cls = rag_class(o['accuracy'])
        dr_cls  = rag_class(o['defect_rate'], inverted=True)
        org_html += f'''<tr onclick="openOrgDrill('{o['org']}')">
          <td><strong>{o['org']}</strong></td>
          <td>{o['audited']}</td>
          <td class="{acc_cls}">{fmt_pct(o['accuracy'])}</td>
          <td>{o['fatal']}</td>
          <td>{o['nonfatal']}</td>
          <td class="{dr_cls}">{fmt_pct(o['defect_rate'])}</td>
        </tr>'''
    org_html += '</tbody></table></div>'

    # FATAL BY ANALYST
    fba_html = '<div class="section half"><h2>Fatal Errors by Analyst</h2>'
    if metrics['fatal_by_analyst']:
        max_c = metrics['fatal_by_analyst'][0][1] if metrics['fatal_by_analyst'] else 1
        fba_html += '<div class="bar-chart">'
        for analyst, count in metrics['fatal_by_analyst']:
            pct = (count / max_c * 100) if max_c else 0
            a_safe = str(analyst).replace("'", "&#39;")
            fba_html += f'''<div class="bar-row" onclick="openAnalystFatalDrill('{a_safe}')">
              <div class="bar-label">{a_safe}</div>
              <div class="bar-track"><div class="bar-fill bar-fatal" style="width:{pct}%"></div></div>
              <div class="bar-value">{count}</div>
            </div>'''
        fba_html += '</div>'
    else:
        fba_html += '<p class="muted">No fatal errors found in Analyst file.</p>'
    fba_html += '</div>'

    # ANALYST TABLE
    at_html = '''<div class="section half">
      <h2>Analyst Performance</h2>
      <input type="text" id="analystSearch" placeholder="Search analyst..." oninput="filterAnalystTable()" class="search-box"/>
      <div class="table-wrap"><table class="data-table" id="analystTable">
        <thead><tr>
          <th onclick="sortDataTable('analystTable',0,'str')">Analyst</th>
          <th onclick="sortDataTable('analystTable',1,'str')">Org</th>
          <th onclick="sortDataTable('analystTable',2,'num')">Audited</th>
          <th onclick="sortDataTable('analystTable',3,'num')">Accuracy</th>
          <th onclick="sortDataTable('analystTable',4,'num')">Fatal</th>
          <th onclick="sortDataTable('analystTable',5,'num')">Non-Fatal Defects</th>
        </tr></thead><tbody>'''
    for a in metrics['analysts']:
        acc_cls = rag_class(a['accuracy'])
        at_html += f'''<tr onclick="openAnalystDrill('{a['analyst']}')">
          <td>{a['analyst']}</td>
          <td>{a['org']}</td>
          <td>{a['audited']}</td>
          <td class="{acc_cls}">{fmt_pct(a['accuracy'])}</td>
          <td>{a['fatal']}</td>
          <td>{a['nonfatal']}</td>
        </tr>'''
    at_html += '</tbody></table></div></div>'

    # FILTER BAR
    def _multi_dd(label, key, values):
        opts = ''
        for v in values:
            v_safe = str(v).replace("'", "&#39;")
            opts += f'<label><input type="checkbox" value="{v_safe}" onchange="onFilterChange()"/> {v_safe}</label>'
        return f'''<div class="filter-dd" data-key="{key}">
          <button class="filter-btn" onclick="toggleDD(this)">{label} <span class="badge" id="badge-{key}">All</span> ▾</button>
          <div class="filter-panel">
            <label class="all-toggle"><input type="checkbox" checked onchange="toggleAll(this,'{key}')"/> All</label>
            <div class="filter-options">{opts}</div>
          </div>
        </div>'''
    f = metrics['filters']
    filter_html = '<div class="filter-bar">'
    filter_html += _multi_dd('Org', 'orgs', f['orgs'])
    filter_html += _multi_dd('Analyst', 'analysts', f['analysts'])
    filter_html += _multi_dd('Category', 'categories', f['categories'])
    filter_html += _multi_dd('Week', 'weeks', f['weeks'])
    filter_html += _multi_dd('Topic', 'topics', f['topics'])
    filter_html += '<button class="reset-btn" onclick="resetFilters()">Reset</button></div>'

    # FOOTER
    footer = ''
    if missing_files:
        footer = f'''<div class="footer-note"><strong>Note:</strong> This process does not yet track {", ".join(missing_files)}. Cards are auto-hidden until those files are added.</div>'''

    rows_json = json.dumps(metrics['rows'], default=str)
    mom_json  = json.dumps(metrics.get('mom') or {}, default=str)
    orgs_json = json.dumps(metrics['orgs'], default=str)
    analysts_json = json.dumps(metrics['analysts'], default=str)

    html = HTML_TEMPLATE.format(
        process=process_label, month=month_label,
        filter_bar=filter_html, kpi_row=kpi_html, status_row=status_html,
        wow=wow_html, mom=mom_html, top_defects=top_html, orgs=org_html,
        fatal_by_analyst=fba_html, analyst_table=at_html, footer=footer,
        rows_json=rows_json, mom_json=mom_json, orgs_json=orgs_json,
        analysts_json=analysts_json,
    )
    return html


# ============================================================
# HTML TEMPLATE
# ============================================================
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"/>
<title>QC Dashboard — {process} {month}</title>
<style>
* {{ box-sizing: border-box; }}
body {{ font-family: -apple-system, "Segoe UI", Arial, sans-serif; margin: 0; padding: 16px;
       background: #f5f6fa; color: #1f2937; }}
h2 {{ font-size: 14px; margin: 0 0 10px 0; color: #374151; }}
.row {{ display: flex; gap: 12px; margin-bottom: 16px; flex-wrap: wrap; }}
.row > * {{ flex: 1 1 0; min-width: 0; }}
.section {{ background: #fff; border-radius: 10px; padding: 14px;
            box-shadow: 0 1px 3px rgba(0,0,0,.06); margin-bottom: 16px; }}
.section.half {{ flex: 1 1 calc(50% - 8px); }}
.muted {{ color: #6b7280; font-size: 12px; }}
.muted.small {{ font-size: 11px; }}

.ph {{ font-size: 13px; color: #6b7280; margin-bottom: 8px; }}

/* Filter bar */
.filter-bar {{ background: #fff; padding: 10px; border-radius: 10px;
               display: flex; gap: 8px; flex-wrap: wrap; align-items: center;
               box-shadow: 0 1px 3px rgba(0,0,0,.06); margin-bottom: 16px; }}
.filter-dd {{ position: relative; }}
.filter-btn {{ background: #f3f4f6; border: 1px solid #d1d5db; padding: 6px 12px;
               border-radius: 6px; cursor: pointer; font-size: 13px; }}
.filter-btn:hover {{ background: #e5e7eb; }}
.badge {{ background: #2563eb; color: #fff; border-radius: 10px;
          padding: 1px 6px; font-size: 11px; margin-left: 4px; }}
.filter-panel {{ display: none; position: absolute; top: 100%; left: 0;
                 background: #fff; border: 1px solid #d1d5db; border-radius: 6px;
                 padding: 8px; min-width: 200px; max-height: 300px; overflow-y: auto;
                 z-index: 100; box-shadow: 0 4px 10px rgba(0,0,0,.1); }}
.filter-panel.open {{ display: block; }}
.filter-panel label {{ display: block; padding: 4px 0; font-size: 13px; cursor: pointer; }}
.all-toggle {{ font-weight: 600; border-bottom: 1px solid #e5e7eb; margin-bottom: 4px;
               padding-bottom: 6px !important; }}
.reset-btn {{ background: #ef4444; color: #fff; border: none; padding: 6px 14px;
              border-radius: 6px; cursor: pointer; font-size: 13px; }}

/* KPI cards */
.kpi-card {{ background: #fff; padding: 16px; border-radius: 10px;
             box-shadow: 0 1px 3px rgba(0,0,0,.06); cursor: pointer;
             transition: transform .1s; }}
.kpi-card:hover {{ transform: translateY(-2px); box-shadow: 0 4px 10px rgba(0,0,0,.1); }}
.kpi-label {{ font-size: 12px; color: #6b7280; text-transform: uppercase; }}
.kpi-value {{ font-size: 28px; font-weight: 700; margin: 6px 0; }}
.kpi-sub {{ font-size: 11px; color: #6b7280; }}
.kpi-fatal .kpi-value {{ color: #dc2626; }}

/* RAG */
.rag-green {{ background: #d1fae5 !important; }}
.rag-amber {{ background: #fef3c7 !important; }}
.rag-red {{ background: #fee2e2 !important; }}
.rag-na    {{ background: #f3f4f6 !important; }}
td.rag-green {{ color: #047857; font-weight: 600; }}
td.rag-amber {{ color: #b45309; font-weight: 600; }}
td.rag-red   {{ color: #b91c1c; font-weight: 600; }}

/* Status cards */
.status-card {{ background: #fff; border-radius: 10px; padding: 14px;
                box-shadow: 0 1px 3px rgba(0,0,0,.06); cursor: pointer; }}
.status-card:hover {{ box-shadow: 0 4px 10px rgba(0,0,0,.1); }}
.status-header {{ font-weight: 600; font-size: 13px; color: #374151;
                  border-bottom: 1px solid #e5e7eb; padding-bottom: 6px; margin-bottom: 8px; }}
.status-total {{ font-size: 13px; color: #6b7280; margin-bottom: 8px; }}
.status-row {{ display: flex; justify-content: space-between; font-size: 13px;
               padding: 3px 0; cursor: pointer; }}
.status-row:hover {{ background: #f9fafb; }}
.status-sla {{ margin-top: 10px; padding: 6px 8px; border-radius: 6px;
               text-align: center; font-size: 12px; font-weight: 600; }}

/* Bar chart */
.bar-chart {{ }}
.bar-row {{ display: grid; grid-template-columns: 1.6fr 3fr 50px;
            gap: 10px; align-items: center; padding: 4px 0;
            cursor: pointer; font-size: 13px; }}
.bar-row:hover {{ background: #f9fafb; }}
.bar-label {{ overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
.bar-track {{ background: #f3f4f6; border-radius: 4px; height: 18px; overflow: hidden; }}
.bar-fill  {{ background: #3b82f6; height: 100%; }}
.bar-fatal {{ background: #dc2626; }}
.bar-value {{ text-align: right; font-weight: 600; }}

/* Org/Analyst tables */
.data-table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
.data-table th {{ background: #f3f4f6; padding: 8px; text-align: left;
                  cursor: pointer; user-select: none; position: sticky; top: 0; }}
.data-table td {{ padding: 6px 8px; border-bottom: 1px solid #f3f4f6; }}
.data-table tr:hover {{ background: #f9fafb; cursor: pointer; }}
.search-box {{ width: 100%; padding: 6px 10px; border: 1px solid #d1d5db;
               border-radius: 6px; margin-bottom: 8px; font-size: 13px; }}
.table-wrap {{ max-height: 320px; overflow-y: auto; }}

/* WoW */
.wow-block {{ margin-bottom: 12px; }}
.wow-title {{ font-size: 12px; font-weight: 600; color: #6b7280; margin-bottom: 4px; }}
.wow-bars {{ display: flex; gap: 8px; align-items: flex-end; height: 80px; }}
.wow-bar-col {{ flex: 1; display: flex; flex-direction: column; align-items: center;
                justify-content: flex-end; height: 100%; }}
.wow-num {{ font-size: 12px; font-weight: 700; }}
.wow-bar {{ width: 100%; height: 60%; background: #f3f4f6; border-radius: 3px 3px 0 0;
            display: flex; flex-direction: column; justify-content: flex-end; }}
.wow-fill {{ width: 100%; border-radius: 3px 3px 0 0; transition: height .3s; }}
.wow-lbl {{ font-size: 11px; color: #6b7280; margin-top: 4px; }}

/* MoM */
.mom-table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
.mom-table th {{ background: #f3f4f6; padding: 8px; text-align: center;
                 font-weight: 600; font-size: 12px; }}
.mom-table td {{ padding: 8px; text-align: center; }}
.mom-table .mom-lbl {{ text-align: left; font-weight: 600; color: #374151; }}
.mom-table .mom-val {{ font-size: 16px; font-weight: 700; }}
.mom-table .vs {{ color: #9ca3af; font-size: 11px; }}
.mom-table .mom-delta {{ font-size: 11px; font-weight: 500; margin-top: 2px; }}
.mom-table .mom-fatal {{ cursor: pointer; color: #dc2626; }}
.arr-up {{ color: #b91c1c; font-weight: 700; }}
.arr-dn {{ color: #047857; font-weight: 700; }}

/* Modal */
.modal-overlay {{ position: fixed; inset: 0; background: rgba(0,0,0,.5);
                  display: none; align-items: center; justify-content: center; z-index: 200; }}
.modal-overlay.open {{ display: flex; }}
.modal {{ background: #fff; border-radius: 10px; padding: 20px;
          max-width: 90%; max-height: 85vh; overflow-y: auto; min-width: 600px; }}
.modal-header {{ display: flex; justify-content: space-between; align-items: center;
                 border-bottom: 1px solid #e5e7eb; padding-bottom: 10px; margin-bottom: 12px; }}
.modal-header h3 {{ margin: 0; font-size: 16px; }}
.modal-close {{ background: #ef4444; color: #fff; border: none; padding: 4px 10px;
                border-radius: 6px; cursor: pointer; }}
.modal-table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
.modal-table th {{ background: #f3f4f6; padding: 6px 8px; text-align: left; }}
.modal-table td {{ padding: 6px 8px; border-bottom: 1px solid #f3f4f6; vertical-align: top; }}

.footer-note {{ background: #fef3c7; border-left: 4px solid #f59e0b;
                padding: 10px 14px; border-radius: 6px; font-size: 12px;
                color: #92400e; margin-top: 16px; }}
</style></head>
<body>
<div class="ph">{process} • {month}</div>

{filter_bar}

{kpi_row}

{status_row}

<div class="row">
  {wow}
  {mom}
</div>

{top_defects}

{orgs}

<div class="row">
  {fatal_by_analyst}
  {analyst_table}
</div>

{footer}

<div id="modalOverlay" class="modal-overlay" onclick="if(event.target===this)closeModal()">
  <div class="modal">
    <div class="modal-header"><h3 id="modalTitle">Drill-down</h3>
      <button class="modal-close" onclick="closeModal()">Close</button></div>
    <div id="modalBody"></div>
  </div>
</div>

<script>
const ROWS = {rows_json};
const MOM = {mom_json};
const ORGS = {orgs_json};
const ANALYSTS = {analysts_json};
let activeFilters = {{ orgs: [], analysts: [], categories: [], weeks: [], topics: [] }};

// Filter dropdown UI
function toggleDD(btn) {{
  document.querySelectorAll('.filter-panel.open').forEach(p => {{
    if (p !== btn.nextElementSibling) p.classList.remove('open');
  }});
  btn.nextElementSibling.classList.toggle('open');
}}
document.addEventListener('click', e => {{
  if (!e.target.closest('.filter-dd')) {{
    document.querySelectorAll('.filter-panel.open').forEach(p => p.classList.remove('open'));
  }}
}});
function toggleAll(checkbox, key) {{
  const panel = checkbox.closest('.filter-panel');
  panel.querySelectorAll('.filter-options input[type=checkbox]').forEach(b => b.checked = false);
  onFilterChange();
}}
function onFilterChange() {{
  ['orgs','analysts','categories','weeks','topics'].forEach(k => {{
    const dd = document.querySelector(`[data-key='${{k}}']`);
    const checked = Array.from(dd.querySelectorAll('.filter-options input:checked')).map(b => b.value);
    activeFilters[k] = checked;
    document.getElementById('badge-' + k).textContent = checked.length === 0 ? 'All' : `${{checked.length}}`;
  }});
  filterAnalystTable();
}}
function resetFilters() {{
  document.querySelectorAll('.filter-options input:checked').forEach(b => b.checked = false);
  document.querySelectorAll('.all-toggle input').forEach(b => b.checked = true);
  onFilterChange();
}}

function applyFilters(rows) {{
  return rows.filter(r => {{
    if (activeFilters.orgs.length && !activeFilters.orgs.includes(r.org)) return false;
    if (activeFilters.analysts.length && !activeFilters.analysts.includes(r.analyst)) return false;
    if (activeFilters.categories.length && !activeFilters.categories.includes(r.category)) return false;
    if (activeFilters.weeks.length && !activeFilters.weeks.includes(r.week)) return false;
    if (activeFilters.topics.length && !activeFilters.topics.includes(r.topic)) return false;
    return true;
  }});
}}

function buildTable(headers, rows) {{
  if (rows.length === 0) return '<p style="color:#6b7280">No matching cases.</p>';
  let html = '<table class="modal-table"><thead><tr>';
  headers.forEach(h => html += `<th>${{h}}</th>`);
  html += '</tr></thead><tbody>';
  rows.forEach(r => {{
    html += '<tr>';
    r.forEach(c => html += `<td>${{c}}</td>`);
    html += '</tr>';
  }});
  return html + '</tbody></table>';
}}

function openModal(title, html) {{
  document.getElementById('modalTitle').textContent = title;
  document.getElementById('modalBody').innerHTML = html;
  document.getElementById('modalOverlay').classList.add('open');
}}
function closeModal() {{ document.getElementById('modalOverlay').classList.remove('open'); }}

// Table sorting
function sortDataTable(tableId, idx, type) {{
  const tbl = document.getElementById(tableId);
  const tbody = tbl.querySelector('tbody');
  const dir = tbl.dataset.sortDir === 'desc' && tbl.dataset.sortIdx == idx ? 'asc' : 'desc';
  tbl.dataset.sortDir = dir;
  tbl.dataset.sortIdx = idx;
  const rows = Array.from(tbody.querySelectorAll('tr'));
  rows.sort((a, b) => {{
    let av = a.children[idx].textContent.trim();
    let bv = b.children[idx].textContent.trim();
    if (type === 'num') {{
      av = parseFloat(av.replace('%','').replace(',','')) || 0;
      bv = parseFloat(bv.replace('%','').replace(',','')) || 0;
      return dir === 'desc' ? bv - av : av - bv;
    }}
    return dir === 'desc' ? bv.localeCompare(av) : av.localeCompare(bv);
  }});
  rows.forEach(r => tbody.appendChild(r));
}}

function filterAnalystTable() {{
  const q = (document.getElementById('analystSearch').value || '').toLowerCase();
  const tbl = document.getElementById('analystTable');
  for (const tr of tbl.querySelectorAll('tbody tr')) {{
    const cells = tr.children;
    const name = cells[0].textContent.toLowerCase();
    const org = cells[1].textContent;
    let show = true;
    if (q && !name.includes(q)) show = false;
    if (activeFilters.orgs.length && !activeFilters.orgs.includes(org)) show = false;
    tr.style.display = show ? '' : 'none';
  }}
}}

// ----- Drill-downs -----
function openDrill(kind) {{
  const filtered = applyFilters(ROWS);
  let title='', headers=[], rows=[];
  if (kind === 'audited') {{
    title = `Cases Audited (${{filtered.length}})`;
    headers = ['Case','Analyst','Org','Date','Accuracy'];
    rows = filtered.map(r => [r.case, r.analyst, r.org, r.date, r.accuracy + '%']);
  }} else if (kind === 'accuracy') {{
    title = 'Accuracy by Case (lowest first)';
    headers = ['Case','Analyst','Org','Accuracy','Defects'];
    const sorted = [...filtered].sort((a,b) => a.accuracy - b.accuracy);
    rows = sorted.map(r => [r.case, r.analyst, r.org, r.accuracy + '%', r.params.join(', ')]);
  }} else if (kind === 'fatal') {{
    // One row per fatal mark — case with both fatal params hits shows TWICE
    // so total rows match KPI's simple sum (Inap + Conf).
    const fatalRows = [];
    filtered.forEach(r => {{
      const hasInap = r.params.some(p => p.toLowerCase().replace(/[^a-z]/g,'').includes('appropriateresolution'));
      const hasConf = r.params.some(p => p.toLowerCase().replace(/[^a-z]/g,'').includes('confidentiality'));
      if (hasInap) fatalRows.push([r.case, r.analyst, r.comment || '(no comment)']);
      if (hasConf) fatalRows.push([r.case, r.analyst, r.comment || '(no comment)']);
    }});
    title = `Fatal Errors (${{fatalRows.length}})`;
    headers = ['Case Number','Analyst','Audit Comment'];
    rows = fatalRows;
  }} else if (kind === 'defects') {{
    const total = filtered.reduce((s,r) => s + r.missed, 0);
    title = `Defects — Total Missed Parameters: ${{total}}`;
    headers = ['Case','Analyst','Org','Missed','Parameters','Comment'];
    rows = filtered.filter(r => r.missed > 0).map(r => [r.case, r.analyst, r.org, r.missed, r.params.join(', '), r.comment]);
  }}
  openModal(title, buildTable(headers, rows));
}}

function openParamDrill(param) {{
  const filtered = applyFilters(ROWS).filter(r => r.params.includes(param));
  openModal(`Cases hitting: ${{param}} (${{filtered.length}})`,
            buildTable(['Case','Analyst','Org','Comment'],
                       filtered.map(r => [r.case, r.analyst, r.org, r.comment])));
}}

function openOrgDrill(org) {{
  const filtered = applyFilters(ROWS).filter(r => r.org === org);
  openModal(`${{org}} — Audits (${{filtered.length}})`,
            buildTable(['Case','Analyst','Manager','Date','Accuracy','Fatal','Defects','Comment'],
                       filtered.map(r => [r.case, r.analyst, r.manager || '(none)', r.date, r.accuracy + '%',
                                          r.is_fatal ? 'Yes' : '', r.missed, r.comment])));
}}

function openAnalystFatalDrill(analyst) {{
  // One row per fatal mark (matches sum-based fatal count)
  const filtered = applyFilters(ROWS).filter(r => r.analyst === analyst);
  const fatalRows = [];
  filtered.forEach(r => {{
    const hasInap = r.params.some(p => p.toLowerCase().replace(/[^a-z]/g,'').includes('appropriateresolution'));
    const hasConf = r.params.some(p => p.toLowerCase().replace(/[^a-z]/g,'').includes('confidentiality'));
    if (hasInap) fatalRows.push([r.case, r.comment || '(no comment)']);
    if (hasConf) fatalRows.push([r.case, r.comment || '(no comment)']);
  }});
  openModal(`${{analyst}} — Fatal Errors (${{fatalRows.length}})`,
            buildTable(['Case Number','Audit Comment'], fatalRows));
}}

function openAnalystDrill(analyst) {{
  const filtered = applyFilters(ROWS).filter(r => r.analyst === analyst);
  openModal(`${{analyst}} — Audits (${{filtered.length}})`,
            buildTable(['Case','Org','Date','Accuracy','Defects','Comment'],
                       filtered.map(r => [r.case, r.org, r.date, r.accuracy + '%', r.params.join(', '), r.comment])));
}}

function openStatusDrill(card, label) {{
  openModal(`${{card}} — ${{label}}`,
            '<p style="color:#6b7280">Detailed table for this status segment will be available in a future iteration.</p>');
}}

function openMoMDrill(kind) {{
  if (!MOM || Object.keys(MOM).length === 0) return;
  // Manager-Org-Count breakdown for current month
  let html = '';
  if (kind === 'fatal') {{
    const fatalRows = ROWS.filter(r => r.is_fatal);
    const buckets = {{}};
    fatalRows.forEach(r => {{
      const key = (r.manager || 'Unknown Manager') + ' • ' + (r.org || 'Unknown Org');
      buckets[key] = (buckets[key] || 0) + 1;
    }});
    const sorted = Object.entries(buckets).sort((a,b) => b[1] - a[1]);
    html = buildTable(['Manager • Org', 'Fatal Count'], sorted.map(([k,v]) => [k, v]));
    openModal(`Fatal Errors — ${{MOM.cur_label}} (Manager × Org breakdown)`, html);
  }} else if (kind === 'proc_accuracy' || kind === 'anl_accuracy') {{
    // Bottom 10 by accuracy (works for both Process and Analyst drill, since
    // ANALYSTS table aggregates per analyst from the Analyst file source)
    const sorted = [...ANALYSTS].sort((a,b) => a.accuracy - b.accuracy).slice(0, 10);
    html = buildTable(['Analyst','Org','Audited','Accuracy','Fatal','Non-Fatal Defects'],
                      sorted.map(a => [a.analyst, a.org, a.audited, a.accuracy + '%', a.fatal, a.nonfatal]));
    const label = kind === 'proc_accuracy' ? 'Process' : 'Analyst';
    openModal(`${{label}} Accuracy — Bottom 10 (${{MOM.cur_label}})`, html);
  }}
}}
</script>
</body></html>
"""


# ============================================================
# EXCEL REPORT BUILDER
# ============================================================
def build_excel(metrics, process_label, month_label, prev_month_label, out_path):
    """Build a multi-sheet Excel report for leadership."""
    try:
        import openpyxl
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("  WARNING: openpyxl not installed; skipping Excel report.")
        print("           Install with: python -m pip install openpyxl --user")
        return None

    wb = openpyxl.Workbook()
    thin = Side(border_style='thin', color='D1D5DB')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill('solid', fgColor='F3F4F6')
    hdr_font = Font(bold=True, color='1F2937')
    green = PatternFill('solid', fgColor='D1FAE5')
    amber = PatternFill('solid', fgColor='FEF3C7')
    red   = PatternFill('solid', fgColor='FEE2E2')

    def color_for(pct, inverted=False, sla=False):
        if pct is None: return None
        if sla: return green if pct >= 100 else red
        if inverted:
            if pct <= 5: return green
            if pct <= 10: return amber
            return red
        if pct >= 95: return green
        if pct >= 90: return amber
        return red

    def write_table(ws, headers, rows, start_row=1, color_cols=None):
        """Write a table starting at start_row. color_cols = {col_idx: ('rag', inverted, sla)}"""
        for c, h in enumerate(headers, 1):
            cell = ws.cell(row=start_row, column=c, value=h)
            cell.font = hdr_font
            cell.fill = hdr_fill
            cell.alignment = Alignment(horizontal='left', vertical='center')
            cell.border = border
        for r_idx, row in enumerate(rows, start_row + 1):
            for c, val in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c, value=val)
                cell.border = border
                if color_cols and c in color_cols:
                    rule = color_cols[c]
                    if rule[0] == 'rag':
                        try:
                            num = float(str(val).replace('%','').replace(',',''))
                            f = color_for(num, inverted=rule[1], sla=rule[2])
                            if f: cell.fill = f
                        except Exception:
                            pass
        # Auto-size columns
        for c in range(1, len(headers) + 1):
            max_len = max(len(str(headers[c-1])),
                          max((len(str(r[c-1])) for r in rows), default=10))
            ws.column_dimensions[get_column_letter(c)].width = min(max_len + 4, 50)
        ws.freeze_panes = ws.cell(row=start_row + 1, column=1)

    # --- Sheet 1: Summary ---
    ws = wb.active
    ws.title = "Summary"
    ws.cell(1, 1, f"QC Report — {process_label} {month_label}").font = Font(bold=True, size=14)
    ws.cell(2, 1, f"Source: Process file. Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}").font = Font(italic=True, color='6B7280')

    h = metrics['headline']
    kpi_rows = [
        ['Cases Audited',     h['audited']],
        ['Accuracy',          f"{h['accuracy']:.2f}%"],
        ['Fatal Errors',      h['fatal']],
        ['  Inappropriate Resolution', h['inap']],
        ['  Confidentiality', h['conf']],
        ['Defects (sum of Missed Parameters)', h['defects']],
        ['Non-Fatal Defects', h['nonfatal_defects']],
        ['Defect Rate',       f"{h['defect_rate']:.2f}%"],
    ]
    write_table(ws, ['KPI', 'Value'], kpi_rows, start_row=4)

    # MoM section
    if metrics.get('mom'):
        m = metrics['mom']
        prev = prev_month_label or 'Prev'
        cur = month_label
        ws.cell(15, 1, "Month over Month").font = Font(bold=True, size=12)
        def _xl_acc(prev_v, cur_v, delta_v):
            p = f"{prev_v:.2f}%" if prev_v is not None else '—'
            c = f"{cur_v:.2f}%"  if cur_v  is not None else '—'
            d = f"{delta_v:+.2f}%" if delta_v is not None else '—'
            return p, c, d
        proc_p, proc_c, proc_d = _xl_acc(m['prev_proc_acc'], m['cur_proc_acc'], m['proc_acc_delta'])
        anl_p,  anl_c,  anl_d  = _xl_acc(m['prev_anl_acc'],  m['cur_anl_acc'],  m['anl_acc_delta'])
        mom_rows = [
            ['Total Audited', m['prev_audited'], m['cur_audited'], m['audited_delta']],
            ['Fatal',         f"{m['prev_fatal']} ({m['prev_fatal_pct']:.1f}%)",
                              f"{m['cur_fatal']} ({m['cur_fatal_pct']:.1f}%)", m['fatal_delta']],
            ['Process Accuracy', proc_p, proc_c, proc_d],
            ['Analyst Accuracy', anl_p,  anl_c,  anl_d],
        ]
        write_table(ws, ['Metric', prev, cur, 'Δ'], mom_rows, start_row=16)

    # Top defects
    ws.cell(23, 1, "Top Defect Parameters").font = Font(bold=True, size=12)
    write_table(ws, ['Parameter', 'Count'], metrics['top_defects'], start_row=24)

    # --- Sheet 2: Org Scorecard ---
    ws2 = wb.create_sheet("Org Scorecard")
    ws2.cell(1, 1, "Org Scorecard").font = Font(bold=True, size=14)
    org_headers = ['Org','Cases Audited','Accuracy','Fatal','Non-Fatal Defects','Defect Rate']
    org_rows = [[o['org'], o['audited'], f"{o['accuracy']:.2f}%", o['fatal'], o['nonfatal'], f"{o['defect_rate']:.2f}%"]
                for o in metrics['orgs']]
    write_table(ws2, org_headers, org_rows, start_row=3,
                color_cols={3:('rag',False,False), 6:('rag',True,False)})

    # --- Sheet 3: Analyst Performance ---
    ws3 = wb.create_sheet("Analysts")
    ws3.cell(1, 1, "Analyst Performance").font = Font(bold=True, size=14)
    ws3.cell(2, 1, "Source: Analyst file").font = Font(italic=True, color='6B7280')
    a_headers = ['Analyst','Org','Manager','Audited','Accuracy','Fatal','Non-Fatal Defects']
    a_rows = [[a['analyst'], a['org'], a.get('manager') or '(none)', a['audited'],
               f"{a['accuracy']:.2f}%", a['fatal'], a['nonfatal']] for a in metrics['analysts']]
    write_table(ws3, a_headers, a_rows, start_row=4,
                color_cols={5:('rag',False,False)})

    # --- Sheet 4: Fatal Cases ---
    ws4 = wb.create_sheet("Fatal Cases")
    ws4.cell(1, 1, "Fatal Cases").font = Font(bold=True, size=14)
    fatal_rows = [[r['case'], r['analyst'], r['org'], r.get('manager') or '(none)',
                   r['date'], ', '.join(r['params']), r['comment'] or '(no comment)']
                  for r in metrics['rows'] if r['is_fatal']]
    write_table(ws4, ['Case','Analyst','Org','Manager','Date','Fatal Parameter(s)','Audit Comment'],
                fatal_rows, start_row=3)

    # --- Sheet 5: Top Defects detail ---
    ws5 = wb.create_sheet("Defect Detail")
    ws5.cell(1, 1, "All Defects (one row per case)").font = Font(bold=True, size=14)
    def_rows = [[r['case'], r['analyst'], r['org'], r['missed'],
                 ', '.join(r['params']), r['comment']]
                for r in metrics['rows'] if r['has_def']]
    write_table(ws5, ['Case','Analyst','Org','# Missed','Parameters','Comment'],
                def_rows, start_row=3)

    wb.save(out_path)
    return out_path


# ============================================================
# SCREENSHOT (using imgkit/wkhtmltopdf if available, else skip with note)
# ============================================================
def build_screenshot(html_path, png_path):
    """Try to render the HTML to PNG. Best-effort — silently skip if tools missing."""
    # Try imgkit (wraps wkhtmltoimage)
    try:
        import imgkit
        imgkit.from_file(str(html_path), str(png_path),
                         options={'width': 1400, 'quiet': ''})
        return png_path
    except Exception:
        pass
    # Try playwright
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page(viewport={'width': 1400, 'height': 1800})
            page.goto(f'file://{html_path}')
            page.wait_for_load_state('networkidle')
            page.screenshot(path=str(png_path), full_page=True)
            browser.close()
        return png_path
    except Exception:
        pass
    print("  NOTE: Screenshot skipped (wkhtmltoimage / playwright not available).")
    print("        Install with one of:")
    print("          python -m pip install imgkit --user   (also needs wkhtmltopdf)")
    print("          python -m pip install playwright --user && playwright install chromium")
    return None


# ============================================================
# MAIN ORCHESTRATION
# ============================================================
def list_datasets():
    if not BASE_FOLDER.exists():
        print(f"\nERROR: Folder not found: {BASE_FOLDER}")
        sys.exit(1)
    return sorted([f for f in BASE_FOLDER.iterdir() if f.is_dir()], key=lambda p: p.name)


def parse_folder_name(name):
    """Parse 'Process_Month_Year' (e.g. 'General_Mar_2026') into (process, month, year)."""
    parts = name.split('_')
    if len(parts) >= 3:
        return parts[0], parts[1], parts[2]
    return name, '', ''


def find_prev_month_folder(current_folder):
    """Given e.g. 'General_Mar_2026', find 'General_Feb_2026' if it exists."""
    process, month, year = parse_folder_name(current_folder.name)
    if month not in MONTHS_ORDER:
        return None, None
    idx = MONTHS_ORDER.index(month)
    if idx == 0:
        prev_month = "Dec"
        prev_year = str(int(year) - 1) if year.isdigit() else year
    else:
        prev_month = MONTHS_ORDER[idx - 1]
        prev_year = year
    candidate = BASE_FOLDER / f"{process}_{prev_month}_{prev_year}"
    if candidate.exists() and candidate.is_dir():
        return candidate, f"{prev_month} {prev_year}"
    return None, None


def run_dataset(folder):
    process, month, year = parse_folder_name(folder.name)
    process_label = process
    month_label = f"{month} {year}".strip()

    print(f"\n{'=' * 60}")
    print(f"Running: {process_label} — {month_label}")
    print('=' * 60)

    print("\nLoading config...")
    config = load_config(folder)

    # Parse current month's process & analyst files
    print("Reading current month files...")
    process_df = parse_audit_file(folder / "Fincom_Process.csv", "Process Audit")
    analyst_df = parse_audit_file(folder / "Fincom_Analyst.csv", "Analyst Audit")

    if len(process_df) == 0:
        print("\nERROR: No data in Process file. Cannot continue.")
        return

    # Other files
    disputes_df = parse_disputes(folder / "Disputes.csv", int(config['disputes_sla_days']))
    rect_df     = parse_rectification(folder / "Rectification.csv", int(config['rectification_sla_days']))
    ivoc_df     = parse_ivoc(folder / "IVOC.csv", int(config['ivoc_sla_days']))
    defred_df   = parse_defect_reduction(folder / "Defect Reduction.csv", month, int(config['reduction_sla_days']))

    # Track missing
    missing = []
    if not (folder / "Rectification.csv").exists(): missing.append("Rectification")
    if not (folder / "IVOC.csv").exists():         missing.append("IVOC")
    if not (folder / "Defect Reduction.csv").exists(): missing.append("Defect Reduction")

    # Previous month for MoM
    prev_folder, prev_month_label = find_prev_month_folder(folder)
    prev_process_df = pd.DataFrame()
    prev_analyst_df = pd.DataFrame()
    if prev_folder:
        print(f"\nPrevious month folder found: {prev_folder.name}")
        prev_process_df = parse_audit_file(prev_folder / "Fincom_Process.csv", "Process Audit (prev)")
        prev_analyst_df = parse_audit_file(prev_folder / "Fincom_Analyst.csv", "Analyst Audit (prev)")
    else:
        print("\nNo previous month folder found — MoM section will be hidden.")

    # Compute metrics
    print("\nComputing metrics...")
    metrics = compute_metrics(process_df, analyst_df, disputes_df, rect_df, ivoc_df, defred_df,
                              prev_process_df, prev_analyst_df, config, process_label, month_label)
    if metrics is None:
        print("ERROR: Could not compute metrics.")
        return

    # Print sanity check
    h = metrics['headline']
    print(f"\n--- HEADLINE NUMBERS (for verification) ---")
    print(f"  Cases Audited:        {h['audited']}")
    print(f"  Accuracy:             {h['accuracy']:.2f}%")
    print(f"  Fatal:                {h['fatal']}  (= Inap {h['inap']} + Conf {h['conf']})")
    print(f"  Defects (Σ Missed):   {h['defects']}")
    print(f"  Non-Fatal Defects:    {h['nonfatal_defects']}")

    # Build outputs
    print("\nBuilding HTML dashboard...")
    html = build_html(metrics, process_label, month_label, prev_month_label, missing)
    html_path = folder / f"QC_Dashboard_{folder.name}.html"
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"  Saved: {html_path}")

    print("Building Excel report...")
    xlsx_path = folder / f"QC_Report_{folder.name}.xlsx"
    if build_excel(metrics, process_label, month_label, prev_month_label, xlsx_path):
        print(f"  Saved: {xlsx_path}")

    print("Building screenshot...")
    png_path = folder / f"QC_Snapshot_{folder.name}.png"
    if build_screenshot(html_path, png_path):
        print(f"  Saved: {png_path}")

    print("\n" + "=" * 60)
    print("ALL DONE")
    print("=" * 60)


def main():
    print("FinCom QC Automation — v4")
    folders = list_datasets()
    if not folders:
        print("No process folders found in", BASE_FOLDER)
        sys.exit(1)
    print("\nAvailable datasets:")
    for i, f in enumerate(folders, 1):
        print(f"  {i}. {f.name}")
    try:
        choice = int(input("\nEnter number to select: ").strip())
        if choice < 1 or choice > len(folders):
            print("Invalid choice.")
            sys.exit(1)
    except (ValueError, KeyboardInterrupt):
        print("\nCancelled.")
        sys.exit(0)
    run_dataset(folders[choice - 1])


if __name__ == "__main__":
    main()
