#!/usr/bin/env python3
"""
generate_dashboard.py
=====================
Reads live Workload and Attendance Excel files, computes all dashboard
datasets, and regenerates weekly_output_dashboard.html with fresh embedded
data.

Usage:
    python generate_dashboard.py
    python generate_dashboard.py --week-end 2026-08-02
    python generate_dashboard.py --workload "path/to/workload.xlsx"

See README_generate_dashboard.txt for full documentation.
"""

import os, sys, re, json, glob, argparse
import pandas as pd
import numpy as np
import unicodedata
from datetime import datetime, timedelta
from dateutil import parser as dateutil_parser

sys.stdout.reconfigure(encoding='utf-8')

# ── FILE PATHS ─────────────────────────────────────────────────────────────────
_SKILL_DIR      = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKLOAD_DIR    = r"G:\My Drive\Dữ liệu nhân sự\Workload\Construction Team"
ATTENDANCE_PATH = r"G:\My Drive\Dữ liệu nhân sự\Data\Timesheet\HR_Fact_Attendance.xlsx"
KPI_PATH        = os.path.join(_SKILL_DIR, "references", "Customer KPI Mapping.json")
TEMPLATE_PATH   = os.path.join(_SKILL_DIR, "template", "weekly_output_dashboard.html")
OUTPUT_PATH     = os.path.join(_SKILL_DIR, "index.html")

# ── CONSTANTS ──────────────────────────────────────────────────────────────────
DATA_CUTOFF   = pd.Timestamp("2026-04-01")
EXCLUDE_USERS = {
    'huong huynh', 'unallocated', 'adrian', 'thuy ho',
    'thuy  ho', 'long nguyen', 'tam doan'
}
CLEVELAND_IDS = {'MTVN0001', 'MTVN0021', 'MTVN0049', 'MTVN0073', 'MTVN0082', 'MTVN0086'}
ACTIVE_TEAMS  = [
    'Cleveland', 'Drafting', 'Estimating',
    'Engineered Wood Products (EWP)', 'Frame & Truss'
]
TRAILING_WEEKS = 4  # weeks used for backlog forecast trailing average

# Custom ID overrides where employee name in Workload differs from DIM_Employee
CUSTOM_ID_MAP = {
    'nam nguyen': 'MTVN0090',
    'phuoc vo':   'MTVN0029',
    'son ho':     'MTVN0085',
    'thuy  ho':   'MTVN0058',
    'adrian r':   'MTVN0059',
}
# Hard-coded team overrides by employee ID
SPECIAL_TEAM_MAP = {
    'MTVN0042':     'Estimating',
    'MTVN-HCM0020': 'Frame & Truss',
    'MTVN0035':     'Drafting',
}

# ── UTILITIES ──────────────────────────────────────────────────────────────────

def normalize_name(s: str) -> str:
    if not isinstance(s, str): return ""
    s = unicodedata.normalize('NFC', s.strip())
    return " ".join(s.split()).lower()

def safe_parse_dt(val):
    if pd.isna(val) or str(val).strip() in ('', 'nan', 'NaT'): return pd.NaT
    try:   return dateutil_parser.parse(str(val))
    except: return pd.NaT

def parse_date_col(series: pd.Series) -> pd.Series:
    return series.apply(safe_parse_dt)

def get_flow_type(task_type_name: str) -> str:
    s = str(task_type_name).lower()
    if 'design'                in s: return 'Design'
    if 'check'                 in s: return 'Check'
    if 'review' in s and 'fix' in s: return 'Review & Fix'
    if 'train'                 in s: return 'Training'
    return 'Other'

def get_job_prefix(job_type: str) -> str:
    s = str(job_type).strip()
    if s.startswith('Amd.'): return 'Amd.'
    if s.startswith('Est.'): return 'Est.'
    if s.startswith('Det.'): return 'Det.'
    return 'Other'

def sla_bucket(due_date, today) -> str:
    if pd.isna(due_date): return 'No due date'
    try:
        d = due_date.date() if hasattr(due_date, 'date') else due_date
        days_left = (d - today).days
        if days_left >= 5: return '5d'
        if days_left == 4: return '4d'
        if days_left == 3: return '3d'
        if days_left == 2: return '2d'
        if days_left == 1: return '1d'
        return '0d'
    except:
        return 'No due date'

def to_iso_str(val) -> str:
    """Convert datetime/Timestamp to ISO-8601 string, safe for JSON."""
    if val is None or (hasattr(val, '__bool__') and pd.isna(val)): return ''
    return str(val)[:19]

# ── FILE DISCOVERY ─────────────────────────────────────────────────────────────

def find_workload_file(workload_dir: str, explicit: str | None = None) -> str:
    if explicit and os.path.exists(explicit):
        return explicit
    files = sorted(glob.glob(os.path.join(workload_dir, "Report_Past Month.xlsx")))
    if not files:
        raise FileNotFoundError(f"No workload file found in: {workload_dir}")
    return files[-1]   # alphabetically latest = most recent date in filename

# ── DATA LOADING ───────────────────────────────────────────────────────────────

def load_all(workload_path: str, attendance_path: str):
    print(f"  Workload  : {os.path.basename(workload_path)}")
    jobs_df  = pd.read_excel(workload_path, sheet_name="Jobs")
    tasks_df = pd.read_excel(workload_path, sheet_name="Tasks")

    print(f"  Attendance: {os.path.basename(attendance_path)}")
    emp_df   = pd.read_excel(attendance_path, sheet_name="DIM_Employee")
    att_df   = pd.read_excel(attendance_path, sheet_name="FACT_Attendance_Daily")
    leave_df = pd.read_excel(attendance_path, sheet_name="Req_Leave")
    ot_df    = pd.read_excel(attendance_path, sheet_name="Req_OT")

    return jobs_df, tasks_df, emp_df, att_df, leave_df, ot_df

# ── DATA PREPARATION ───────────────────────────────────────────────────────────

def prep_jobs(jobs_df: pd.DataFrame) -> pd.DataFrame:
    for col in ['Time Created', 'Time Started', 'Time Complete']:
        jobs_df[col] = parse_date_col(jobs_df[col])
    jobs_df = jobs_df[jobs_df['Company'] != 'Test'].copy()
    jobs_df = jobs_df[jobs_df['Time Created'] > DATA_CUTOFF].copy()
    return jobs_df.reset_index(drop=True)

def prep_tasks(tasks_df: pd.DataFrame, valid_job_ids: set) -> pd.DataFrame:
    for col in ['timeCreated', 'timeCouldBeginAt', 'timeAccepted', 'timeOutcome']:
        if col in tasks_df.columns:
            tasks_df[col] = parse_date_col(tasks_df[col])
    tasks_df = tasks_df[tasks_df['jobId'].isin(valid_job_ids)].copy()
    tasks_df = tasks_df[~tasks_df['taskTypeName'].str.contains(
        r'Admin\s*-\s*QS|Admin\s*QS', case=False, na=False)].copy()
    tasks_df = tasks_df[tasks_df['timeCreated'] > DATA_CUTOFF].copy()
    for col in ['Designed Square Meter', 'Designed Layout', 'Dwelling Units']:
        if col in tasks_df.columns:
            tasks_df[col] = pd.to_numeric(tasks_df[col], errors='coerce').fillna(0)
    if 'investedSeconds' in tasks_df.columns:
        tasks_df['investedSeconds'] = pd.to_numeric(
            tasks_df['investedSeconds'], errors='coerce').fillna(0)

    # Fallback for Complete/Rejected tasks missing timeOutcome
    if 'timeAccepted' in tasks_df.columns and 'investedSeconds' in tasks_df.columns:
        mask = (
            tasks_df['Flow Task_status'].isin(['Complete', 'Completed', 'Rejected']) &
            tasks_df['timeOutcome'].isna() &
            tasks_df['timeAccepted'].notna()
        )
        tasks_df.loc[mask, 'timeOutcome'] = tasks_df.loc[mask, 'timeAccepted'] + pd.to_timedelta(tasks_df.loc[mask, 'investedSeconds'], unit='s')

    return tasks_df.reset_index(drop=True)

def apply_completion_override(
    jobs_df: pd.DataFrame, tasks_df: pd.DataFrame
) -> pd.DataFrame:
    """Override Time Complete / Status for RES Engineering & Prime Design
    using the date of the latest completed Check task."""
    target = jobs_df[
        jobs_df['Company'].str.contains(
            'RES Engineering|Prime Design', case=False, na=False)
    ]['Job #'].tolist()

    checks = tasks_df[
        (tasks_df['jobId'].isin(target)) &
        (tasks_df['taskTypeName'].str.contains('Check', case=False, na=False)) &
        (tasks_df['Flow Task_status'].isin(['Complete', 'Completed'])) &
        (tasks_df['timeOutcome'].notna())
    ]
    if len(checks):
        for job_id, t_out in checks.groupby('jobId')['timeOutcome'].max().items():
            mask = jobs_df['Job #'] == job_id
            jobs_df.loc[mask, 'Time Complete'] = t_out
            jobs_df.loc[mask, 'Status']        = 'Complete'
    return jobs_df

# ── USER → TEAM MAPPING ────────────────────────────────────────────────────────

def build_user_maps(tasks_df: pd.DataFrame, emp_df: pd.DataFrame):
    """Return (user_to_team, workload_to_emp_id) dicts."""
    emp_df = emp_df.copy()
    emp_df['EmployeeID'] = emp_df['EmployeeID'].astype(str).str.strip()

    user_to_team   = {}
    workload_to_id = {}

    for name in sorted(tasks_df['userName'].dropna().unique()):
        norm = normalize_name(name)
        if norm in EXCLUDE_USERS:
            user_to_team[name] = 'Excluded'
            continue

        # 1. Custom hard-coded ID override
        emp_id = CUSTOM_ID_MAP.get(norm)
        if emp_id is None:
            # 2. Match by EN name
            m = emp_df[emp_df['FullNameEN'].apply(normalize_name) == norm]
            if len(m) == 0:
                # 3. Match by VN name
                m = emp_df[emp_df['FullNameVN'].apply(normalize_name) == norm]
            emp_id = m.iloc[0]['EmployeeID'] if len(m) > 0 else None

        if emp_id:
            workload_to_id[name] = emp_id
            if emp_id in CLEVELAND_IDS:
                user_to_team[name] = 'Cleveland'
            elif emp_id in SPECIAL_TEAM_MAP:
                user_to_team[name] = SPECIAL_TEAM_MAP[emp_id]
            else:
                row = emp_df[emp_df['EmployeeID'] == emp_id]
                user_to_team[name] = (
                    row.iloc[0]['Team'] if len(row) > 0 else 'Unknown')
        else:
            user_to_team[name] = 'Unknown'

    return user_to_team, workload_to_id

# ── JOB → TEAM PRE-COMPUTATION ────────────────────────────────────────────────

def build_job_team_cache(
    tasks_df: pd.DataFrame, user_to_team: dict
) -> dict:
    """Pre-compute {job_id: team} for every job that has tasks.
    Priority: Design task → Check task → any task.
    """
    df = tasks_df.copy()
    df['_team'] = df['userName'].map(user_to_team).fillna('Unknown')
    df = df[~df['_team'].isin(['Unknown', 'Excluded'])]

    def priority(name):
        n = str(name).lower()
        if 'design' in n: return 0
        if 'check'  in n: return 1
        return 2

    df['_pri'] = df['taskTypeName'].apply(priority)
    df = df.sort_values(['jobId', '_pri'])
    return df.groupby('jobId')['_team'].first().to_dict()

def job_team(job_id, job_team_cache: dict) -> str:
    return job_team_cache.get(job_id, 'Unmapped')

# ── WEEK RANGES ────────────────────────────────────────────────────────────────

def compute_weeks(w0_end: datetime) -> list[tuple]:
    """All Mon–Sun weeks from DATA_CUTOFF to w0_end.
    Returns list of (week_start, week_end, 'YYYY-MM-DD' Monday label)."""
    # Find the Sunday of the w0_end week
    days_to_sun = (6 - w0_end.weekday()) % 7
    sunday = (w0_end + timedelta(days=days_to_sun)).replace(
        hour=23, minute=59, second=59, microsecond=0)

    weeks = []
    cur_sun = sunday
    while True:
        ws = (cur_sun - timedelta(days=6)).replace(
            hour=0, minute=0, second=0, microsecond=0)
        if ws.date() < DATA_CUTOFF.date():
            break
        weeks.append((ws, cur_sun, ws.strftime('%Y-%m-%d')))
        cur_sun -= timedelta(weeks=1)

    return list(reversed(weeks))

# ── DATASET BUILDERS ───────────────────────────────────────────────────────────

def _comp_jobs(jobs_df, ws, we):
    return jobs_df[
        (jobs_df['Time Complete'] >= ws) &
        (jobs_df['Time Complete'] <= we) &
        (jobs_df['Status'] == 'Complete')
    ]

def _comp_tasks(tasks_df, ws, we):
    return tasks_df[
        (tasks_df['timeOutcome'] >= ws) &
        (tasks_df['timeOutcome'] <= we) &
        (tasks_df['Flow Task_status'].isin(['Complete', 'Completed']))
    ]

def build_customer_weekly(
    jobs_df, tasks_df, weeks, job_team_cache, job_map
) -> list[dict]:
    rows = []
    for ws, we, wlabel in weeks:
        cj = _comp_jobs(jobs_df, ws, we).copy()
        if cj.empty: continue
        cj['_team'] = cj['Job #'].map(job_team_cache).fillna('Unmapped')
        cj = cj[~cj['_team'].isin(['Excluded', 'Unknown', 'Unmapped'])]

        # Design-task volumes for completed jobs
        dt = tasks_df[
            (tasks_df['jobId'].isin(cj['Job #'])) &
            (tasks_df['Flow Task_status'].isin(['Complete', 'Completed'])) &
            (tasks_df['taskTypeName'].str.contains('Design', case=False, na=False))
        ].copy()
        dt['_company'] = dt['jobId'].map(lambda x: job_map.get(x, {}).get('Company', 'Unknown'))
        dt['_team']    = dt['jobId'].map(job_team_cache).fillna('Unmapped')
        dt = dt[~dt['_team'].isin(['Excluded', 'Unknown', 'Unmapped'])]

        jg = cj.groupby(['_team', 'Company']).size().reset_index(name='jobs')
        vg = (dt.groupby(['_team', '_company'])
                [['Dwelling Units', 'Designed Square Meter', 'Designed Layout']]
                .sum().reset_index()
                .rename(columns={'_company': 'Company'}))
        hg = (dt.groupby(['_team', '_company'])['userName']
                .nunique().reset_index(name='headcount')
                .rename(columns={'_company': 'Company'}))

        mg = (jg.merge(vg, on=['_team', 'Company'], how='outer')
                .merge(hg, on=['_team', 'Company'], how='outer')
                .fillna(0))

        for _, r in mg.iterrows():
            j = float(r.get('jobs', 0))
            s = float(r.get('Designed Square Meter', 0))
            rows.append({
                'Week':    wlabel,
                'Team':    r['_team'],
                'Customer': r['Company'],
                'jobs':    j,
                'layouts': float(r.get('Designed Layout', 0)),
                'units':   float(r.get('Dwelling Units', 0)),
                'sqm':     s,
                'avg_sqm_per_job': round(s / j, 1) if j > 0 and s > 0 else 0.0,
                'headcount': float(r.get('headcount', 0)),
            })
    return rows


def build_employee_weekly(
    tasks_df, weeks, user_to_team, job_map
) -> list[dict]:
    rows = []
    for ws, we, wlabel in weeks:
        wt = _comp_tasks(tasks_df, ws, we).copy()
        wt = wt[wt['userName'].apply(normalize_name).isin(
            set(tasks_df['userName'].apply(normalize_name).unique()) - EXCLUDE_USERS)]
        wt['_team'] = wt['userName'].map(user_to_team).fillna('Unknown')
        wt = wt[~wt['_team'].isin(['Excluded', 'Unknown'])]
        if wt.empty: continue

        wt['_company']  = wt['jobId'].map(lambda x: job_map.get(x, {}).get('Company', 'Unknown'))
        wt['_flowtype'] = wt['taskTypeName'].apply(get_flow_type)

        counts = (wt.groupby(['userName', '_team', '_company', '_flowtype'])
                    .size().unstack(fill_value=0).reset_index())
        for c in ['Design', 'Check', 'Review & Fix']:
            if c not in counts.columns: counts[c] = 0

        design_wt = wt[wt['_flowtype'] == 'Design']
        vols = (design_wt.groupby(['userName', '_team', '_company'])
                         [['Dwelling Units', 'Designed Square Meter', 'Designed Layout']]
                         .sum().reset_index())

        mg = counts.merge(vols, on=['userName', '_team', '_company'], how='outer').fillna(0)
        for _, r in mg.iterrows():
            rows.append({
                'Week':         wlabel,
                'Employee':     r['userName'],
                'Team':         r['_team'],
                'Customer':     r['_company'],
                'design_tasks': float(r.get('Design', 0)),
                'check_tasks':  float(r.get('Check', 0)),
                'layouts':      float(r.get('Designed Layout', 0)),
                'units':        float(r.get('Dwelling Units', 0)),
                'sqm':          float(r.get('Designed Square Meter', 0)),
            })
    return rows


def build_rework_datasets(
    tasks_df, weeks, user_to_team, job_map
) -> tuple[list[dict], list[dict]]:
    team_rows, emp_rows = [], []
    for ws, we, wlabel in weeks:
        wt = _comp_tasks(tasks_df, ws, we).copy()
        wt['_team']     = wt['userName'].map(user_to_team).fillna('Unknown')
        wt = wt[~wt['_team'].isin(['Excluded', 'Unknown'])]
        if wt.empty: continue
        wt['_company']  = wt['jobId'].map(lambda x: job_map.get(x, {}).get('Company', 'Unknown'))
        wt['_flowtype'] = wt['taskTypeName'].apply(get_flow_type)

        for (grp_cols, target) in [
            (['_team', '_company'],           team_rows),
            (['userName', '_team', '_company'], emp_rows),
        ]:
            g = (wt.groupby(grp_cols + ['_flowtype'])
                   .size().unstack(fill_value=0).reset_index())
            for c in ['Design', 'Review & Fix']:
                if c not in g.columns: g[c] = 0
            for _, r in g.iterrows():
                rec = {'Week': wlabel}
                for col in grp_cols:
                    key = {'_team': 'Team', '_company': 'Customer',
                           'userName': 'Employee'}.get(col, col)
                    rec[key] = r[col]
                rec['design_tasks']    = float(r.get('Design', 0))
                rec['reviewfix_tasks'] = float(r.get('Review & Fix', 0))
                target.append(rec)
    return team_rows, emp_rows


def build_passthrough_weekly(
    jobs_df, weeks, job_team_cache, tasks_df, job_map
) -> list[dict]:
    latest_label = weeks[-1][2] if weeks else ''
    # Current backlog snapshot per (team, company)
    active = jobs_df[jobs_df['Status'] != 'Complete'].copy()
    active['_team'] = active['Job #'].map(job_team_cache).fillna('Unmapped')
    active = active[~active['_team'].isin(['Excluded', 'Unknown', 'Unmapped'])]
    backlog_snap = active.groupby(['_team', 'Company']).size().to_dict()

    rows = []
    for ws, we, wlabel in weeks:
        sent = jobs_df[
            (jobs_df['Time Created'] >= ws) &
            (jobs_df['Time Created'] <= we)
        ].copy()
        sent['_team'] = sent['Job #'].map(job_team_cache).fillna('Unmapped')
        sent = sent[~sent['_team'].isin(['Excluded', 'Unknown', 'Unmapped'])]

        comp = _comp_jobs(jobs_df, ws, we).copy()
        comp['_team'] = comp['Job #'].map(job_team_cache).fillna('Unmapped')
        comp = comp[~comp['_team'].isin(['Excluded', 'Unknown', 'Unmapped'])]

        intra_ids = set(comp[
            (comp['Time Created'] >= ws) &
            (comp['Time Created'] <= we)
        ]['Job #'])

        combos = (set(zip(sent['_team'], sent['Company'])) |
                  set(zip(comp['_team'],  comp['Company'])))

        for team, company in combos:
            s_n = len(sent[(sent['_team'] == team) & (sent['Company'] == company)])
            c_sub = comp[(comp['_team'] == team) & (comp['Company'] == company)]
            i_n = len(c_sub[c_sub['Job #'].isin(intra_ids)])
            p_n = len(c_sub[~c_sub['Job #'].isin(intra_ids)])
            bl  = int(backlog_snap.get((team, company), 0)) if wlabel == latest_label else 0
            rows.append({
                'Week': wlabel, 'Team': team, 'Customer': company,
                'jobs_sent': s_n, 'intra_week': i_n,
                'pass_week': p_n, 'backlog': bl,
            })
    return rows


def build_jobtype_weekly(
    jobs_df, weeks, job_team_cache
) -> list[dict]:
    rows = []
    for ws, we, wlabel in weeks:
        cj = _comp_jobs(jobs_df, ws, we).copy()
        if cj.empty: continue
        cj['_team'] = cj['Job #'].map(job_team_cache).fillna('Unmapped')
        cj = cj[~cj['_team'].isin(['Excluded', 'Unknown', 'Unmapped'])]
        for _, r in cj.iterrows():
            jt = str(r.get('Job Type', '')).strip() or 'Other'
            rows.append({
                'Week': wlabel, 'Team': r['_team'],
                'Customer': r.get('Company', 'Unknown'),
                'ProjectType': get_job_prefix(jt),
                'JobType': jt, 'jobs': 1,
            })
    if not rows: return []
    df = pd.DataFrame(rows)
    agg = (df.groupby(['Week', 'Team', 'Customer', 'ProjectType', 'JobType'])
             ['jobs'].sum().reset_index())
    return agg.to_dict('records')


def build_job_scatter(
    jobs_df, tasks_df, weeks, job_team_cache, job_map
) -> list[dict]:
    rows = []
    for ws, we, wlabel in weeks:
        cj = _comp_jobs(jobs_df, ws, we).copy()
        if cj.empty: continue
        cj['_team'] = cj['Job #'].map(job_team_cache).fillna('Unmapped')
        cj = cj[~cj['_team'].isin(['Excluded', 'Unknown', 'Unmapped'])]

        for _, job in cj.iterrows():
            jid = job['Job #']
            dt = tasks_df[
                (tasks_df['jobId'] == jid) &
                (tasks_df['taskTypeName'].str.contains('Design', case=False, na=False)) &
                (tasks_df['Flow Task_status'].isin(['Complete', 'Completed']))
            ]
            sqm   = float(dt['Designed Square Meter'].sum()) if len(dt) else 0.0
            units = float(dt['Dwelling Units'].sum())        if len(dt) else 0.0
            hrs   = float(dt['investedSeconds'].sum()) / 3600.0 if len(dt) else 0.0
            designers = ', '.join(sorted(dt['userName'].dropna().unique())) if len(dt) else ''
            if sqm == 0 and units == 0 and hrs == 0: continue

            created   = job.get('Time Created')
            completed = job.get('Time Complete')
            age_days  = (completed - created).days if pd.notna(created) and pd.notna(completed) else 0
            jt = str(job.get('Job Type', '')).strip() or 'Other'
            rows.append({
                'Week':     wlabel,
                'Team':     job['_team'],
                'Customer': job.get('Company', 'Unknown'),
                'jobId':    str(jid),
                'JobType':  jt,
                'sqm':      round(sqm, 1),
                'units':    int(units),
                'invested_hours': round(hrs, 2),
                'Designers':  designers,
                'TimeCreated':  to_iso_str(created)[:10],
                'TimeComplete': to_iso_str(completed)[:10],
                'AgeDays':  age_days,
            })
    return rows


def build_leadtime_weekly(
    jobs_df, weeks, job_team_cache
) -> list[dict]:
    rows = []
    for ws, we, wlabel in weeks:
        cj = _comp_jobs(jobs_df, ws, we).copy()
        if cj.empty: continue
        cj['_team'] = cj['Job #'].map(job_team_cache).fillna('Unmapped')
        cj = cj[~cj['_team'].isin(['Excluded', 'Unknown', 'Unmapped'])]
        cj['_lt'] = (cj['Time Complete'] - cj['Time Started']).dt.total_seconds() / 86400.0
        cj = cj[cj['_lt'] >= 0]
        grp = cj.groupby(['_team', 'Company'])['_lt'].agg(['mean', 'count']).reset_index()
        for _, r in grp.iterrows():
            rows.append({
                'Week': wlabel, 'Team': r['_team'],
                'Customer': r['Company'],
                'avg_leadtime': round(float(r['mean']), 2),
                'n': int(r['count']),
            })
    return rows


def build_sla_datasets(
    jobs_df, job_team_cache
) -> tuple[list, list, list, str]:
    today     = datetime.now().date()
    today_str = today.strftime('%Y-%m-%d')

    active = jobs_df[jobs_df['Status'] != 'Complete'].copy()
    active['_team'] = active['Job #'].map(job_team_cache).fillna('Unmapped')
    active = active[~active['_team'].isin(['Excluded', 'Unknown', 'Unmapped'])]

    if 'Due Date' in active.columns:
        active['Due Date'] = pd.to_datetime(active['Due Date'], errors='coerce')
        active['_bucket'] = active['Due Date'].apply(lambda d: sla_bucket(d, today))
    else:
        active['_bucket'] = 'No due date'

    act_agg = active.groupby(['_team', 'Company']).size().reset_index(name='active_jobs')
    sla_active = [
        {'Team': r['_team'], 'Customer': r['Company'], 'active_jobs': int(r['active_jobs'])}
        for _, r in act_agg.iterrows()
    ]

    dist_agg = active.groupby(['_team', 'Company', '_bucket']).size().reset_index(name='n')
    sla_dist  = [
        {'Team': r['_team'], 'Customer': r['Company'],
         'SLABucket': r['_bucket'], 'n': int(r['n'])}
        for _, r in dist_agg.iterrows()
    ]

    sla_jobs = [
        {
            'Team': r['_team'], 'Customer': r.get('Company', ''),
            'SLABucket': r['_bucket'], 'jobId': str(r['Job #']),
            'jobRef': str(r.get('Address', r.get('Job Ref', r['Job #']))).strip(),
        }
        for _, r in active.iterrows()
    ]
    return sla_active, sla_dist, sla_jobs, today_str


def build_headcount_weekly(
    tasks_df, weeks, user_to_team, job_map
) -> list[dict]:
    rows = []
    for ws, we, wlabel in weeks:
        wt = _comp_tasks(tasks_df, ws, we).copy()
        wt['_team'] = wt['userName'].map(user_to_team).fillna('Unknown')
        wt = wt[~wt['_team'].isin(['Excluded', 'Unknown'])]
        if wt.empty: continue
        wt['_company'] = wt['jobId'].map(lambda x: job_map.get(x, {}).get('Company', 'Unknown'))
        grp = wt.groupby(['_team', '_company'])['userName'].nunique().reset_index(name='headcount')
        for _, r in grp.iterrows():
            rows.append({
                'Week': wlabel, 'Team': r['_team'],
                'Customer': r['_company'], 'headcount': int(r['headcount']),
            })
    return rows


def build_backlog_forecast(
    jobs_df, tasks_df, weeks, job_team_cache, user_to_team, job_map,
    n_trailing: int = TRAILING_WEEKS
) -> tuple[list[dict], list[str]]:
    if len(weeks) < 2: return [], []
    trailing = weeks[-n_trailing:] if len(weeks) >= n_trailing else weeks
    trailing_labels = [w[2] for w in trailing]

    # Current backlog
    active = jobs_df[jobs_df['Status'] != 'Complete'].copy()
    active['_team'] = active['Job #'].map(job_team_cache).fillna('Unmapped')
    active = active[~active['_team'].isin(['Excluded', 'Unknown', 'Unmapped'])]
    backlog_grp = active.groupby(['_team', 'Company']).size().to_dict()

    wk_jobs  = {}  # (team, company) -> [count per trailing week]
    wk_hrs   = {}
    wk_hc    = {}

    for ws, we, _ in trailing:
        cj = _comp_jobs(jobs_df, ws, we).copy()
        cj['_team'] = cj['Job #'].map(job_team_cache).fillna('Unmapped')
        cj = cj[~cj['_team'].isin(['Excluded', 'Unknown', 'Unmapped'])]
        for (team, company), cnt in cj.groupby(['_team', 'Company']).size().items():
            wk_jobs.setdefault((team, company), []).append(cnt)

        wt = _comp_tasks(tasks_df, ws, we).copy()
        wt['_team']    = wt['userName'].map(user_to_team).fillna('Unknown')
        wt = wt[~wt['_team'].isin(['Excluded', 'Unknown'])]
        wt['_company'] = wt['jobId'].map(lambda x: job_map.get(x, {}).get('Company', 'Unknown'))
        for (team, company), g in wt.groupby(['_team', '_company']):
            total_hrs = float(g['investedSeconds'].sum()) / 3600.0
            wk_hrs.setdefault((team, company), []).append(total_hrs)
            wk_hc.setdefault((team, company), []).append(int(g['userName'].nunique()))

    forecast_rows = []
    for key in set(backlog_grp.keys()) | set(wk_jobs.keys()):
        team, company = key
        bl = backlog_grp.get(key, 0)
        if bl == 0: continue

        avg_wk_jobs = float(np.mean(wk_jobs.get(key, [0])))
        avg_wk_hrs  = float(np.mean(wk_hrs.get(key, [0])))
        avg_wk_hc   = float(np.mean(wk_hc.get(key, [0])))

        total_c_hrs  = sum(wk_hrs.get(key, [0]))
        total_c_jobs = sum(wk_jobs.get(key, [0]))
        avg_hrs_job  = total_c_hrs / total_c_jobs if total_c_jobs > 0 else 0.0
        hours_needed = bl * avg_hrs_job

        forecast_rows.append({
            'Customer': company, 'Team': team,
            'backlog':         int(bl),
            'avg_weekly_jobs': round(avg_wk_jobs, 2),
            'total_hours_needed': round(hours_needed, 1),
            'avg_weekly_hours':   round(avg_wk_hrs,  1),
            'avg_headcount':      round(avg_wk_hc,   1),
        })
    return forecast_rows, trailing_labels

# ── GANTT RAW DATA ─────────────────────────────────────────────────────────────

def build_gantt_raw(
    tasks_df, att_df, leave_df, weeks,
    user_to_team, workload_to_id,
    gantt_weeks: int = 2
) -> dict:
    if not weeks: return {}
    recent = weeks[-gantt_weeks:]
    g_start, g_end = recent[0][0], recent[-1][1]

    # Tasks visible in the Gantt window (created/started/ended within or overlapping)
    gt = tasks_df[
        (tasks_df['timeOutcome'].isna() |
         (tasks_df['timeOutcome'] >= g_start)) &
        (tasks_df['timeCreated'].isna() |
         (tasks_df['timeCreated'] <= g_end))
    ].copy()
    gt = gt[gt['userName'].apply(normalize_name).map(
        lambda n: n not in EXCLUDE_USERS)]
    gt['_team'] = gt['userName'].map(user_to_team).fillna('Unknown')
    gt = gt[~gt['_team'].isin(['Excluded', 'Unknown'])]

    # Attendance hours in Gantt window per employee
    att_df = att_df.copy()
    att_df['Date_Text'] = pd.to_datetime(att_df['Date_Text'], errors='coerce')
    g_att = att_df[
        (att_df['Date_Text'] >= g_start) &
        (att_df['Date_Text'] <= g_end)
    ].copy()
    g_att['_hrs'] = (
        (g_att['Type of Date'].astype(str).str.strip() == 'FullWorkDay').astype(float) * 8.0 +
        (g_att['Type of Date'].astype(str).str.strip() == 'HalfWorkDay').astype(float) * 4.0
    )
    att_per_emp = g_att.groupby('Employee_ID')['_hrs'].sum().to_dict()

    # Approved leaves in Gantt window per employee
    leave_df = leave_df.copy()
    leave_df['LeaveFromwithTime'] = pd.to_datetime(leave_df['LeaveFromwithTime'], errors='coerce')
    leave_df['LeaveTowithTime']   = pd.to_datetime(leave_df['LeaveTowithTime'],   errors='coerce')
    g_leaves = leave_df[
        (leave_df['Status'].astype(str).str.strip().str.lower() == 'đã duyệt') &
        (leave_df['LeaveFromwithTime'] <= g_end) &
        (leave_df['LeaveTowithTime']   >= g_start)
    ].copy()

    now_str = datetime.now().strftime('%Y-%m-%dT%H:%M:%S')
    raw = {}

    for team in ACTIVE_TEAMS:
        team_tasks = gt[gt['_team'] == team]
        if team_tasks.empty: continue
        raw.setdefault(team, {})

        for user in sorted(team_tasks['userName'].unique()):
            u_tasks = team_tasks[team_tasks['userName'] == user]
            emp_id  = workload_to_id.get(user, '')

            att_hrs   = float(att_per_emp.get(emp_id, 0))
            intra_hrs = float(u_tasks['investedSeconds'].sum()) / 3600.0
            util_str  = (f"{intra_hrs / att_hrs * 100:.1f}%"
                         if att_hrs > 0 else "N/A")

            task_list = []
            for _, t in u_tasks.iterrows():
                status = str(t.get('Flow Task_status', '')).strip()
                in_prog = status not in ('Complete', 'Completed')
                invested_t = float(t.get('investedSeconds', 0)) / 3600.0

                # jobRef: try jobRef field first, fall back to jobId
                job_ref_val = t.get('jobRef', t.get('jobId', ''))

                task_list.append({
                    'taskId':       str(t.get('taskId', t.name)),
                    'jobRef':       str(job_ref_val),
                    'parentJobRef': str(t.get('parentJobRef', '')),
                    'cat':          get_flow_type(str(t.get('taskTypeName', ''))),
                    'status':       status if status else ('In Progress' if in_prog else 'Complete'),
                    'created':      to_iso_str(t.get('timeCreated')),
                    'start':        to_iso_str(t.get('timeAccepted')) if pd.notna(t.get('timeAccepted')) else ('' if status.lower() == 'pending' else to_iso_str(t.get('timeCouldBeginAt'))),
                    'end':          now_str if in_prog else to_iso_str(t.get('timeOutcome')),
                    'investedTime': f"{invested_t:.2f}h",
                    'investedHrs':  invested_t,
                    'inProgress':   in_prog,
                })

            leave_list = []
            emp_att_list = []
            if emp_id:
                for _, lv in g_leaves[g_leaves['Employee_ID'] == emp_id].iterrows():
                    lt = str(lv.get('Leave_Type', 'Other'))
                    l_mapped = 'Annual' if any(
                        kw in lt.lower() for kw in ['annual', 'phép', 'năm']
                    ) else 'Other'
                    leave_list.append({
                        'start': to_iso_str(lv['LeaveFromwithTime']),
                        'end':   to_iso_str(lv['LeaveTowithTime']),
                        'type':  l_mapped,
                    })
                
                emp_att = g_att[g_att['Employee_ID'] == emp_id]
                for _, r in emp_att.iterrows():
                    actual = float(r.get('Số giờ làm việc thực tế', 0)) if pd.notna(r.get('Số giờ làm việc thực tế')) else 0.0
                    emp_att_list.append({
                        'date': to_iso_str(r['Date_Text'])[:10],
                        'hrs': float(r['_hrs']),
                        'actual_hrs': actual
                    })

            raw[team][user] = {
                'attHrs':      f"{att_hrs:.1f}",
                'intraHrs':    f"{intra_hrs:.1f}",
                'utilization': util_str,
                'name':        user,
                'tasks':       task_list,
                'leaves':      leave_list,
                'attendance':  emp_att_list,
            }
    return raw

# ── KPI / CUSTOMER TARGETS ────────────────────────────────────────────────────

def load_kpi_config(kpi_path: str) -> dict:
    if not os.path.exists(kpi_path): return {}
    try:
        with open(kpi_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return {}

def build_customer_targets(kpi_config: dict) -> dict:
    targets = {}
    for customer, cfg in kpi_config.items():
        if customer.startswith('_') or not isinstance(cfg, dict):
            continue   # skip metadata keys like _comment, _format
        metric = str(cfg.get('Expected Metric', 'Jobs')).lower()
        daily  = float(cfg.get('Daily Expected Output', 0))
        weekly = daily * 5
        if weekly > 0:
            targets[customer] = {'target': int(weekly), 'metric': metric}
    return targets

# ── HTML INJECTION ─────────────────────────────────────────────────────────────

def export_data_and_copy_template(
    template_path: str,
    output_path: str,
    dashdata: dict,
    gantt_raw: dict,
    gantt_week_start: str,
    gantt_week_end: str,
    customer_targets: dict,
) -> None:
    # 1. Write the JS data file
    js_path = os.path.join(os.path.dirname(template_path), "..", "references", "dashdata.js")
    os.makedirs(os.path.dirname(js_path), exist_ok=True)
    
    with open(js_path, 'w', encoding='utf-8') as f:
        f.write(f"window.DASH_DATA = {json.dumps(dashdata, ensure_ascii=False, default=str)};\n")
        f.write(f"window.GANTT_RAW = {json.dumps(gantt_raw, ensure_ascii=False, default=str)};\n")
        f.write(f"window.CUSTOMER_TARGETS = {json.dumps(customer_targets if customer_targets else {}, ensure_ascii=False, default=str)};\n")
        f.write(f"window.GANTT_DATES = {{'start': '{gantt_week_start}', 'end': '{gantt_week_end}'}};\n")
        
    print(f"  ✓  Exported data to dashdata.js")

    # 2. Copy the template to output_path (reports folder)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(template_path, 'r', encoding='utf-8') as src:
        content = src.read()
    with open(output_path, 'w', encoding='utf-8') as dst:
        dst.write(content)
        
    print(f"  ✓  Dashboard copied to → {output_path}")


# ── MAIN ───────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(
        description='Generate Weekly Output Dashboard from live Excel data.')
    ap.add_argument('--workload',    type=str, default=None,
                    help='Path to Workload Excel file (auto-detected if omitted)')
    ap.add_argument('--attendance',  type=str, default=None,
                    help='Path to Attendance Excel file')
    ap.add_argument('--week-end',    type=str, default=None,
                    help='Last day of reporting week YYYY-MM-DD (defaults to last Sunday)')
    ap.add_argument('--output',      type=str, default=None,
                    help='Override output HTML path')
    args = ap.parse_args()

    output_path = args.output or OUTPUT_PATH

    # ── Resolve source files ───────────────────────────────────────────────────
    workload_path = find_workload_file(WORKLOAD_DIR, args.workload)
    att_path      = args.attendance or ATTENDANCE_PATH

    for label, path in [
        ('Workload',    workload_path),
        ('Attendance',  att_path),
        ('Template',    TEMPLATE_PATH),
    ]:
        if not os.path.exists(path):
            print(f"❌  {label} file not found:\n   {path}")
            sys.exit(1)

    # ── Determine W0 end date ──────────────────────────────────────────────────
    if args.week_end:
        w0_end = datetime.strptime(args.week_end, '%Y-%m-%d')
        w0_end = w0_end.replace(hour=23, minute=59, second=59)
    else:
        today         = datetime.now()
        current_sun   = today + timedelta(days=6 - today.weekday())  # Current Sunday
        w0_end        = current_sun.replace(hour=23, minute=59, second=59)

    print(f"\n{'─'*60}")
    print(f"  Weekly Output Dashboard Generator")
    print(f"  Reporting week: W/E {w0_end.strftime('%d %b %Y (%A)')}")
    print(f"{'─'*60}\n")

    # ── Load data ──────────────────────────────────────────────────────────────
    print("📂  Loading source data...")
    jobs_df, tasks_df, emp_df, att_df, leave_df, ot_df = \
        load_all(workload_path, att_path)

    # ── Prepare ────────────────────────────────────────────────────────────────
    print("🔧  Preparing data...")
    jobs_df  = prep_jobs(jobs_df)
    tasks_df = prep_tasks(tasks_df, set(jobs_df['Job #']))
    jobs_df  = apply_completion_override(jobs_df, tasks_df)

    user_to_team, workload_to_id = build_user_maps(tasks_df, emp_df)
    job_team_cache = build_job_team_cache(tasks_df, user_to_team)
    job_map = jobs_df.set_index('Job #').to_dict('index')

    # ── Compute week ranges ────────────────────────────────────────────────────
    weeks = compute_weeks(w0_end)
    print(f"📊  {len(weeks)} weeks to compute "
          f"({weeks[0][2]} → {weeks[-1][2]})\n")

    # ── Build all datasets ─────────────────────────────────────────────────────
    steps = [
        ("customer_weekly",       lambda: build_customer_weekly(
            jobs_df, tasks_df, weeks, job_team_cache, job_map)),
        ("employee_weekly",       lambda: build_employee_weekly(
            tasks_df, weeks, user_to_team, job_map)),
        ("rework datasets",       lambda: build_rework_datasets(
            tasks_df, weeks, user_to_team, job_map)),
        ("passthrough_weekly",    lambda: build_passthrough_weekly(
            jobs_df, weeks, job_team_cache, tasks_df, job_map)),
        ("jobtype_weekly",        lambda: build_jobtype_weekly(
            jobs_df, weeks, job_team_cache)),
        ("job_scatter",           lambda: build_job_scatter(
            jobs_df, tasks_df, weeks, job_team_cache, job_map)),
        ("leadtime_weekly",       lambda: build_leadtime_weekly(
            jobs_df, weeks, job_team_cache)),
        ("SLA datasets",          lambda: build_sla_datasets(
            jobs_df, job_team_cache)),
        ("headcount_weekly",      lambda: build_headcount_weekly(
            tasks_df, weeks, user_to_team, job_map)),
        ("backlog_forecast",      lambda: build_backlog_forecast(
            jobs_df, tasks_df, weeks, job_team_cache, user_to_team, job_map)),
        ("Gantt RAW",             lambda: build_gantt_raw(
            tasks_df, att_df, leave_df, weeks,
            user_to_team, workload_to_id, gantt_weeks=2)),
    ]

    results = {}
    for name, fn in steps:
        print(f"  → {name}...")
        results[name] = fn()

    rework_team, rework_emp = results["rework datasets"]
    sla_active, sla_dist, sla_jobs, sla_today = results["SLA datasets"]
    forecast_rows, trailing_labels = results["backlog_forecast"]
    gantt_raw = results["Gantt RAW"]

    # ── KPI targets ────────────────────────────────────────────────────────────
    print("  → CUSTOMER_TARGETS from KPI config...")
    kpi_config = load_kpi_config(KPI_PATH)
    customer_targets = build_customer_targets(kpi_config)
    if not customer_targets:
        print("    (KPI config not found — CUSTOMER_TARGETS left as-is in HTML)")

    # ── Assemble dashdata ──────────────────────────────────────────────────────
    dashdata = {
        'weeks':                 [w[2] for w in weeks],
        'teams':                 ACTIVE_TEAMS,
        'customer_weekly':       results["customer_weekly"],
        'employee_weekly':       results["employee_weekly"],
        'rework_team_weekly':    rework_team,
        'rework_employee_weekly': rework_emp,
        'passthrough_weekly':    results["passthrough_weekly"],
        'jobtype_weekly':        results["jobtype_weekly"],
        'job_scatter':           results["job_scatter"],
        'leadtime_weekly':       results["leadtime_weekly"],
        'sla_active_jobs':       sla_active,
        'sla_distribution':      sla_dist,
        'sla_job_list':          sla_jobs,
        'backlog_forecast':      forecast_rows,
        'headcount_weekly':      results["headcount_weekly"],
        'sla_today_ref':         sla_today,
        'forecast_trailing_weeks': trailing_labels,
        'customer_targets':      customer_targets,   # consumed by CUSTOMER_TARGETS JS var
    }

    # ── Inject & write dashboard ───────────────────────────────────────────────
    w0_start_str = (w0_end - timedelta(days=6)).strftime('%Y-%m-%d')
    w0_end_str   = w0_end.strftime('%Y-%m-%d')

    print(f"\n📝  Exporting data and copying template...")
    export_data_and_copy_template(
        TEMPLATE_PATH, output_path,
        dashdata, gantt_raw,
        gantt_week_start=w0_start_str,
        gantt_week_end=w0_end_str,
        customer_targets=customer_targets,
    )

    # ── Summary ────────────────────────────────────────────────────────────────
    total_active = sum(r['active_jobs'] for r in sla_active)
    print(f"\n{'─'*60}")
    print(f"  ✅  Dashboard generated successfully!")
    print(f"  Weeks computed     : {len(weeks)}")
    print(f"  customer_weekly    : {len(results['customer_weekly'])} rows")
    print(f"  employee_weekly    : {len(results['employee_weekly'])} rows")
    print(f"  job_scatter        : {len(results['job_scatter'])} points")
    print(f"  Active backlog     : {total_active} jobs")
    print(f"  Gantt teams        : {len(gantt_raw)}")
    print(f"\n  Open in browser    :")
    print(f"  {output_path}")
    print(f"{'─'*60}\n")


if __name__ == '__main__':
    main()
