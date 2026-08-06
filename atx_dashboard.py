"""
ATX Peace — Lead Dashboard + Report Builder
Tab 1: Visual dashboard (lead-driven, playbook standards)
Tab 2: Excel report generator (existing flow)
"""
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date, timedelta
import plotly.graph_objects as go
import plotly.express as px
import io, re

try:
    import report_core
    REPORT_CORE_AVAILABLE = True
except ImportError:
    REPORT_CORE_AVAILABLE = False

# ── Page config ─────────────────────────────────────────────────────
st.set_page_config(
    page_title="ATX Peace · Lead Dashboard",
    page_icon="🕊️",
    layout="wide",
)

# ── CSS ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .card {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 16px 18px;
        margin-bottom: 12px;
    }
    .metric-big { font-size: 2rem; font-weight: 700; color: #1F3864; line-height: 1.1; }
    .metric-label { font-size: 0.78rem; color: #64748B; text-transform: uppercase; letter-spacing: .06em; margin-bottom: 4px; }
    .section-title {
        font-size: 1.05rem; font-weight: 700; color: #1F3864;
        border-bottom: 2px solid #E2E8F0; padding-bottom: 6px; margin-bottom: 14px; margin-top: 6px;
    }
    .badge-green  { background:#D1FAE5; color:#065F46; padding:2px 8px; border-radius:12px; font-size:0.78rem; font-weight:600; }
    .badge-amber  { background:#FEF3C7; color:#92400E; padding:2px 8px; border-radius:12px; font-size:0.78rem; font-weight:600; }
    .badge-red    { background:#FEE2E2; color:#991B1B; padding:2px 8px; border-radius:12px; font-size:0.78rem; font-weight:600; }
    .badge-gray   { background:#F1F5F9; color:#475569; padding:2px 8px; border-radius:12px; font-size:0.78rem; font-weight:600; }
    .place-bar-wrap { margin-bottom: 8px; }
    .place-label { font-size:0.8rem; font-weight:600; color:#374151; margin-bottom:2px; }
    .place-bar-bg { background:#E5E7EB; border-radius:5px; height:16px; overflow:hidden; }
    .stDataFrame { font-size: 0.83rem !important; }
</style>
""", unsafe_allow_html=True)

# ── Constants ────────────────────────────────────────────────────────
DARK_BLUE = "#1F3864"
MED_BLUE  = "#2E5FA3"
LITE_BLUE = "#BDD7EE"
GREEN     = "#1A7340"
AMBER     = "#B8600A"
RED_COL   = "#B91C1C"

ATX_PRIMARY  = ["Dove Springs", "Oltorf", "Douglas Landing", "South Lamar"]
ATX_OUTER    = ["Pleasant Valley", "Riverside", "Windsor Hills", "East Cesar Chavez", "Parker Lane", "Montopolis"]
HACA_PROPS   = ["Meadowbrook", "Bouldin Oaks", "Booker T", "Chalmers",
                "Manchaca Village", "Rosewood", "Shadow Bend", "Cardinal Hills", "Gaston Place"]
TRAVIS_KEYS  = ["travis", "wildflower"]

QUARTER_PRESETS = {
    "Q1 · Oct–Dec 2025": ("2025-10-01", "2025-12-31"),
    "Q2 · Jan–Mar 2026": ("2026-01-01", "2026-03-31"),
    "Q3 · Apr–Jun 2026": ("2026-04-01", "2026-06-30"),
    "Q4 · Jul–Sep 2026": ("2026-07-01", "2026-09-30"),
}

# Playbook targets (ATX Peace)
Q_TARGET_CM       = 19
Q_TARGET_IMPROVED = 15
DAILY_CO_HRS      = 2      # canvass & outreach per TM per day
TRAVIS_HRS_DAILY  = 1      # Travis HS / Wildflower per TM per day
FU_PER_CASE_WK    = 2      # follow-ups per case per week (3 for Critical)

# ── Helpers ──────────────────────────────────────────────────────────
def clean_pid(v):
    try:
        return str(int(float(str(v).strip())))
    except:
        return str(v).strip()

def email_name(e):
    if pd.isna(e): return ""
    e = str(e).strip().lower()
    base = e.split("@")[0].replace(".", " ").title().split()
    if len(base) >= 2:
        return f"{base[0][0]}. {base[-1]}"
    return base[0] if base else e

def find_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

def days_ago_from(dt):
    if pd.isna(dt): return None
    try:
        return max(0, (pd.Timestamp.now().normalize() - pd.Timestamp(dt).normalize()).days)
    except:
        return None

def progress_bar(value, target, label, unit=""):
    pct = min(value / target * 100, 100) if target > 0 else 0
    color = GREEN if pct >= 90 else (AMBER if pct >= 60 else RED_COL)
    val_str = f"{value:.1f}{unit}" if isinstance(value, float) else f"{value}{unit}"
    tgt_str = f"{target:.1f}{unit}" if isinstance(target, float) else f"{target}{unit}"
    return f"""
    <div style="margin-bottom:10px">
      <div style="display:flex;justify-content:space-between;margin-bottom:3px">
        <span style="font-size:0.8rem;font-weight:600;color:#374151">{label}</span>
        <span style="font-size:0.8rem;color:#6B7280">{val_str} / {tgt_str}</span>
      </div>
      <div style="background:#E5E7EB;border-radius:5px;height:12px;overflow:hidden">
        <div style="background:{color};width:{pct:.1f}%;height:100%;border-radius:5px"></div>
      </div>
      <div style="font-size:0.7rem;color:#9CA3AF;margin-top:1px">{pct:.0f}%</div>
    </div>"""

def load_csv(f):
    if f is None: return pd.DataFrame()
    f.seek(0)
    try:
        return pd.read_csv(f, low_memory=False, encoding='utf-8')
    except UnicodeDecodeError:
        f.seek(0)
        return pd.read_csv(f, low_memory=False, encoding='latin-1')

def parse_dates(df, candidates):
    col = find_col(df, candidates)
    if col:
        df = df.copy()
        df['_date'] = pd.to_datetime(df[col], errors='coerce')
    else:
        df = df.copy()
        df['_date'] = pd.NaT
    return df

def in_range(df, fs, fe):
    if '_date' not in df.columns or df.empty: return df
    return df[(df['_date'] >= fs) & (df['_date'] <= fe)].copy()

def working_days_in(d_start, d_end):
    return sum(1 for d in pd.date_range(d_start, d_end) if d.weekday() < 5)

# ── Sidebar ──────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div style="background:{DARK_BLUE};padding:14px 16px;border-radius:8px;margin-bottom:14px">
      <div style="color:white;font-size:1.05rem;font-weight:700">🕊️ ATX Peace</div>
      <div style="color:{LITE_BLUE};font-size:0.78rem">Lead Dashboard · Life Anew</div>
    </div>""", unsafe_allow_html=True)

    # ── Date filter
    st.subheader("📅 Date Range")
    view_mode = st.selectbox("View By", ["Quarter", "Month", "Week", "Day", "Custom Range"])

    if view_mode == "Quarter":
        q_choice = st.selectbox("Quarter", list(QUARTER_PRESETS.keys()), index=2)
        fs_str, fe_str = QUARTER_PRESETS[q_choice]
        filter_start = date.fromisoformat(fs_str)
        filter_end   = date.fromisoformat(fe_str)

    elif view_mode == "Month":
        anchor = st.date_input("Month", value=date(2026, 6, 1))
        filter_start = anchor.replace(day=1)
        if anchor.month == 12:
            filter_end = anchor.replace(day=31)
        else:
            filter_end = (anchor.replace(month=anchor.month + 1, day=1) - timedelta(days=1))

    elif view_mode == "Week":
        anchor = st.date_input("Any day in week", value=date.today())
        filter_start = anchor - timedelta(days=anchor.weekday())
        filter_end   = filter_start + timedelta(days=6)

    elif view_mode == "Day":
        filter_start = st.date_input("Date", value=date.today())
        filter_end   = filter_start

    else:
        filter_start = st.date_input("From", value=date(2026, 4, 1))
        filter_end   = st.date_input("To",   value=date(2026, 6, 30))

    wd = working_days_in(filter_start, filter_end)
    st.caption(f"📆 {filter_start.strftime('%b %d')} – {filter_end.strftime('%b %d, %Y')} · {wd} working days")

    st.markdown("---")

    # ── File uploads
    st.subheader("📁 Upload CSV Files")
    parts_file    = st.file_uploader("👤 Participant Folders", type="csv", key="pf")
    goals_file    = st.file_uploader("🎯 Goals",              type="csv", key="gl")
    followup_file = st.file_uploader("📞 Follow-ups",         type="csv", key="fu")
    outreach_file = st.file_uploader("🚶 Outreach & Canvass", type="csv", key="oc")
    circles_file  = st.file_uploader("⭕ Circles & Classes",  type="csv", key="ci")
    incidents_file= st.file_uploader("🚨 Incidents",          type="csv", key="inc")
    pre_file      = st.file_uploader("📋 Pre-Assessment",     type="csv", key="pre")
    post_file     = st.file_uploader("📋 Post-Assessment",    type="csv", key="post")
    attest_file   = st.file_uploader("✍️ Self-Attestation",   type="csv", key="att")

    all_files = [parts_file, goals_file, followup_file, outreach_file,
                 circles_file, incidents_file, pre_file, post_file, attest_file]
    n_uploaded = sum(1 for f in all_files if f is not None)
    if n_uploaded == 9:
        st.success(f"✅ All 9 files loaded")
    elif n_uploaded > 0:
        st.info(f"{n_uploaded}/9 files · {9 - n_uploaded} more for full view")
    else:
        st.caption("Upload files to load dashboard")

# ── Load data ────────────────────────────────────────────────────────
pf_raw  = load_csv(parts_file)
gl_raw  = load_csv(goals_file)
fu_raw  = load_csv(followup_file)
oc_raw  = load_csv(outreach_file)
ci_raw  = load_csv(circles_file)
inc_raw = load_csv(incidents_file)
pre_raw = load_csv(pre_file)
post_raw= load_csv(post_file)
att_raw = load_csv(attest_file)

FS = pd.Timestamp(filter_start)
FE = pd.Timestamp(filter_end) + pd.Timedelta(hours=23, minutes=59)

# Parse dates once
fu_df  = parse_dates(fu_raw,  ['Date of Follow Up','Date','Follow Up Date'])
oc_df  = parse_dates(oc_raw,  ['Date','Activity Date','Event Date'])
inc_df = parse_dates(inc_raw, ['Date','Incident Date','Date of Incident'])
gl_df  = parse_dates(gl_raw,  ['Date Created','Created Date','Date'])

# In-range slices
fu_p   = in_range(fu_df,  FS, FE)
oc_p   = in_range(oc_df,  FS, FE)
inc_p  = in_range(inc_df, FS, FE)

# ── Coord name parsing ───────────────────────────────────────────────
COORD_COL = ['Created by','Coordinator','coordinator','Staff']
def add_coord(df):
    c = find_col(df, COORD_COL)
    if c:
        df = df.copy()
        df['_coord'] = df[c].apply(email_name)
    return df

oc_p  = add_coord(oc_p)
fu_p  = add_coord(fu_p)
inc_p = add_coord(inc_p)

# ── Outreach hours ───────────────────────────────────────────────────
def compute_hours(df):
    dur_col   = find_col(df, ['Duration (Hours)','Duration','Hours','Hours Logged','Time Spent (Hours)'])
    start_col = find_col(df, ['Start Time','Arrival Time','Start'])
    end_col   = find_col(df, ['End Time','Departure Time','End'])
    df = df.copy()
    if dur_col:
        df['_hrs'] = pd.to_numeric(df[dur_col], errors='coerce').fillna(0)
    elif start_col and end_col:
        ts = pd.to_datetime(df[start_col], errors='coerce')
        te = pd.to_datetime(df[end_col],   errors='coerce')
        df['_hrs'] = ((te - ts).dt.total_seconds() / 3600).clip(lower=0).fillna(0)
    else:
        df['_hrs'] = 1.0  # 1 hr assumed per session when no time data
    return df

if not oc_p.empty:
    oc_p = compute_hours(oc_p)
    loc_col_oc = find_col(oc_p, ['Location','Neighborhood','Location/Neighborhood','Place','Area'])
    if loc_col_oc:
        oc_p['_loc'] = oc_p[loc_col_oc].fillna('Unknown').str.strip()
        oc_p['_loc_lower'] = oc_p['_loc'].str.lower()
        oc_p['_is_travis'] = oc_p['_loc_lower'].apply(lambda x: any(t in x for t in TRAVIS_KEYS))
    else:
        oc_p['_loc'] = 'Unknown'
        oc_p['_is_travis'] = False

# ── PID helpers ──────────────────────────────────────────────────────
PID_COLS = ['Participant iD','Participant ID','PID','Participant_ID']
def add_pid(df):
    c = find_col(df, PID_COLS)
    if c:
        df = df.copy()
        df['_pid'] = df[c].apply(clean_pid)
    return df

pf_df  = add_pid(pf_raw.copy())
gl_df  = add_pid(gl_df)
fu_p   = add_pid(fu_p)

# ══════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════
tab1, tab2 = st.tabs(["📊 Dashboard", "📄 Generate Report"])

# ╔══════════════════════════════════════════════════════════════════╗
# ║  TAB 1 · DASHBOARD                                               ║
# ╚══════════════════════════════════════════════════════════════════╝
with tab1:
    if n_uploaded == 0:
        st.markdown("""
        <div style="text-align:center;padding:60px 0;color:#94A3B8">
            <div style="font-size:3rem">🕊️</div>
            <div style="font-size:1.2rem;font-weight:600;margin-top:8px">ATX Peace Lead Dashboard</div>
            <div style="margin-top:6px">Upload CSV files in the sidebar to load the dashboard.</div>
        </div>""", unsafe_allow_html=True)
        st.stop()

    # Header bar
    st.markdown(f"""
    <div style="background:{DARK_BLUE};padding:14px 20px;border-radius:8px;margin-bottom:18px;
                display:flex;justify-content:space-between;align-items:center">
      <div>
        <span style="color:white;font-size:1.3rem;font-weight:700">ATX Peace — Lead Dashboard</span>
        <span style="color:{LITE_BLUE};font-size:0.82rem;margin-left:12px">Life Anew · CVI Program</span>
      </div>
      <div style="color:{LITE_BLUE};font-size:0.85rem">
        {filter_start.strftime('%b %d')} – {filter_end.strftime('%b %d, %Y')} &nbsp;·&nbsp; {wd} working days
      </div>
    </div>""", unsafe_allow_html=True)

    # ─────────────────────────────────────────────────────────────────
    # SECTION 1 · PROGRAM SCORECARD
    # ─────────────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">📈 Program Scorecard</div>', unsafe_allow_html=True)

    # Compute top-level metrics
    n_cm = 0
    if not pf_df.empty and '_pid' in pf_df.columns:
        n_cm = pf_df['_pid'].nunique()

    n_improved = 0
    if not post_raw.empty:
        n_improved = len(post_raw)

    n_outreach = len(oc_p) if not oc_p.empty else 0
    total_co_hrs = oc_p['_hrs'].sum() if ('_hrs' in oc_p.columns and not oc_p.empty) else 0

    n_fu = len(fu_p) if not fu_p.empty else 0
    fu_connected = 0
    if not fu_p.empty:
        conn_col = find_col(fu_p, ['Contact Made','Connected','Reached','Result','Was contact made'])
        if conn_col:
            fu_connected = fu_p[conn_col].astype(str).str.lower().isin(['yes','true','1','connected','reached','made']).sum()

    sc1, sc2, sc3, sc4 = st.columns(4)
    with sc1:
        st.markdown(f"""<div class="card">
            <div class="metric-label">Case-Managed</div>
            <div class="metric-big">{n_cm}</div>
            {progress_bar(n_cm, Q_TARGET_CM, "vs Q target")}
        </div>""", unsafe_allow_html=True)
    with sc2:
        st.markdown(f"""<div class="card">
            <div class="metric-label">Improved</div>
            <div class="metric-big">{n_improved}</div>
            {progress_bar(n_improved, Q_TARGET_IMPROVED, "vs Q target")}
        </div>""", unsafe_allow_html=True)
    with sc3:
        st.markdown(f"""<div class="card">
            <div class="metric-label">Outreach Sessions</div>
            <div class="metric-big">{n_outreach}</div>
            <div style="color:#6B7280;font-size:0.8rem;margin-top:4px">{total_co_hrs:.1f} hrs logged</div>
        </div>""", unsafe_allow_html=True)
    with sc4:
        st.markdown(f"""<div class="card">
            <div class="metric-label">Follow-ups</div>
            <div class="metric-big">{n_fu}</div>
            <div style="color:#6B7280;font-size:0.8rem;margin-top:4px">{fu_connected} connected</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("---")

    # ─────────────────────────────────────────────────────────────────
    # SECTION 2 · PARTICIPANT STATUS TABLE
    # ─────────────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">👥 Participant Status</div>', unsafe_allow_html=True)

    if pf_df.empty or '_pid' not in pf_df.columns:
        st.info("Upload Participant Folders CSV to see participant status.")
    else:
        fn_col     = find_col(pf_df, ['Participant First Name','First Name','first_name','FirstName'])
        ln_col     = find_col(pf_df, ['Participant Last Name','Last Name','last_name','LastName'])
        status_col = find_col(pf_df, ['Caseload Status','Status','Case Status','status'])
        coord_col  = find_col(pf_df, COORD_COL)
        start_col  = find_col(pf_df, ['Case Start Date','Case Start','Date of Case Start','case_start'])

        # Build one row per unique participant
        rows = []
        for _, row in pf_df.drop_duplicates('_pid').iterrows():
            pid  = row['_pid']
            fn   = str(row.get(fn_col, '') or '') if fn_col else ''
            ln   = str(row.get(ln_col, '') or '') if ln_col else ''
            name = f"{fn} {ln}".strip() or f"PID {pid}"
            status = str(row.get(status_col, 'Unknown') or 'Unknown') if status_col else 'Unknown'
            coord  = email_name(row.get(coord_col, '')) if coord_col else ''

            # Last follow-up date
            last_fu = None
            days_fu = None
            if not fu_df.empty and '_pid' in fu_df.columns and '_date' in fu_df.columns:
                p_fu = fu_df[fu_df['_pid'] == pid].sort_values('_date', ascending=False)
                if not p_fu.empty:
                    last_fu = p_fu.iloc[0]['_date']
                    days_fu = days_ago_from(last_fu)

            # Goals
            n_goals = 0
            mid_due = False
            if not gl_df.empty and '_pid' in gl_df.columns:
                p_goals = gl_df[gl_df['_pid'] == pid].copy()
                n_goals = len(p_goals)
                # Check open goal > 21 days
                gst_col = find_col(p_goals, ['Status','Goal Status','status'])
                if '_date' in p_goals.columns and gst_col:
                    open_g = p_goals[~p_goals[gst_col].astype(str).str.lower().isin(['complete','completed','closed','done'])]
                    if not open_g.empty:
                        oldest = (pd.Timestamp.now() - open_g['_date'].min()).days
                        if pd.notna(oldest) and oldest >= 21:
                            mid_due = True

            # Intake docs — look for any doc/consent/release column
            doc_col = next((c for c in pf_df.columns
                            if any(k in c.lower() for k in ['intake','document','consent','release','signed'])), None)
            has_docs = True
            if doc_col:
                val = str(row.get(doc_col, '') or '')
                has_docs = val.strip() not in ['', 'nan', 'None', 'No', '0', 'false']

            rows.append({
                'PID': pid,
                'Name': name,
                'Status': status,
                'Coordinator': coord,
                'Last Follow-up': last_fu.strftime('%b %d, %Y') if last_fu and not pd.isna(last_fu) else '—',
                'Days Since F/U': days_fu if days_fu is not None else 999,
                'Goals': n_goals,
                'Mid-Assess': '⚠️ Due' if mid_due else '✓',
                'Intake Docs': '✓' if has_docs else '—',
                '_mid': mid_due,
                '_days_fu': days_fu if days_fu is not None else 999,
                '_status': status,
            })

        pt = pd.DataFrame(rows)

        if not pt.empty:
            # Filters row
            f1, f2 = st.columns(2)
            with f1:
                status_opts = sorted(pt['Status'].unique().tolist())
                status_sel  = st.multiselect("Status", status_opts, default=status_opts, key="s_filter")
            with f2:
                coord_opts = sorted([c for c in pt['Coordinator'].unique() if c])
                coord_sel  = st.multiselect("Coordinator", coord_opts, default=coord_opts, key="c_filter")

            mask = pt['Status'].isin(status_sel)
            if coord_sel:
                mask &= pt['Coordinator'].isin(coord_sel)
            pt_show = pt[mask]

            # Style dataframe
            def style_row(row):
                styles = []
                for col in row.index:
                    if col == 'Days Since F/U':
                        if row[col] > 14:
                            styles.append('background-color:#FEE2E2')
                        elif row[col] > 7:
                            styles.append('background-color:#FEF3C7')
                        else:
                            styles.append('')
                    elif col == 'Mid-Assess' and '⚠️' in str(row[col]):
                        styles.append('background-color:#FEF3C7;font-weight:600')
                    elif col == 'Intake Docs' and row[col] == '—':
                        styles.append('background-color:#FEE2E2')
                    else:
                        styles.append('')
                return styles

            display = ['PID','Name','Status','Coordinator','Last Follow-up','Days Since F/U','Goals','Mid-Assess','Intake Docs']
            st.dataframe(
                pt_show[display].style.apply(style_row, axis=1),
                hide_index=True,
                use_container_width=True,
            )

            # Alert summaries
            overdue_fu = pt_show[pt_show['_days_fu'] > 14]
            needs_mid  = pt_show[pt_show['_mid'] == True]
            no_docs    = pt_show[pt_show['Intake Docs'] == '—']

            alert_cols = st.columns(3)
            with alert_cols[0]:
                if not overdue_fu.empty:
                    st.error(f"**{len(overdue_fu)}** participant(s) — no follow-up in 14+ days")
            with alert_cols[1]:
                if not needs_mid.empty:
                    st.warning(f"**{len(needs_mid)}** participant(s) — mid-assessment due (goal 3+ wks open)")
            with alert_cols[2]:
                if not no_docs.empty:
                    st.warning(f"**{len(no_docs)}** participant(s) — intake docs missing")

    st.markdown("---")

    # ─────────────────────────────────────────────────────────────────
    # SECTION 3 · TRUSTED MESSENGER DAILY TRACKER
    # ─────────────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">📋 Trusted Messenger Activity Tracker</div>', unsafe_allow_html=True)
    st.caption(f"Standard: {DAILY_CO_HRS} hrs canvass & outreach/day · {TRAVIS_HRS_DAILY} hr Travis HS/Wildflower/day · follow-ups logged same day")

    if oc_p.empty and fu_p.empty:
        st.info("Upload Outreach and Follow-up CSVs to see TM activity.")
    else:
        # Build per-TM stats
        tm_data = {}

        if not oc_p.empty and '_coord' in oc_p.columns and '_hrs' in oc_p.columns:
            for coord, grp in oc_p.groupby('_coord'):
                if not coord: continue
                days_logged = grp['_date'].dt.date.nunique()
                total_hrs   = grp['_hrs'].sum()
                travis_hrs  = grp[grp['_is_travis']]['_hrs'].sum() if '_is_travis' in grp.columns else 0
                daily_hrs   = grp.groupby(grp['_date'].dt.date)['_hrs'].sum()
                days_2hr    = int((daily_hrs >= DAILY_CO_HRS).sum())

                tm_data[coord] = {
                    'Days C&O Logged': days_logged,
                    'Total C&O Hrs': round(total_hrs, 1),
                    'Travis/Wildflower Hrs': round(travis_hrs, 1),
                    'Days Meeting 2hr': days_2hr,
                    'Follow-ups': 0,
                    'FU Connected': 0,
                }

        if not fu_p.empty and '_coord' in fu_p.columns:
            for coord, grp in fu_p.groupby('_coord'):
                if not coord: continue
                conn_col = find_col(grp, ['Contact Made','Connected','Reached','Result'])
                n_conn = 0
                if conn_col:
                    n_conn = grp[conn_col].astype(str).str.lower().isin(
                        ['yes','true','1','connected','reached','made']).sum()
                if coord not in tm_data:
                    tm_data[coord] = {'Days C&O Logged':0,'Total C&O Hrs':0,
                                      'Travis/Wildflower Hrs':0,'Days Meeting 2hr':0,
                                      'Follow-ups':0,'FU Connected':0}
                tm_data[coord]['Follow-ups']   = len(grp)
                tm_data[coord]['FU Connected'] = int(n_conn)

        if tm_data:
            # Two-col grid of TM cards
            tm_names = sorted(tm_data.keys())
            for i in range(0, len(tm_names), 2):
                cols = st.columns(2)
                for j, coord in enumerate(tm_names[i:i+2]):
                    s = tm_data[coord]
                    target_co_hrs    = wd * DAILY_CO_HRS
                    target_travis    = wd * TRAVIS_HRS_DAILY
                    with cols[j]:
                        st.markdown(f"""<div class="card">
                            <div style="font-size:0.95rem;font-weight:700;color:{DARK_BLUE};margin-bottom:10px">{coord}</div>
                            {progress_bar(s['Days C&O Logged'], wd, "Days Logged")}
                            {progress_bar(s['Total C&O Hrs'], target_co_hrs, "Canvass & Outreach Hrs", "h")}
                            {progress_bar(s['Travis/Wildflower Hrs'], target_travis, "Travis HS / Wildflower Hrs", "h")}
                            <div style="margin-top:8px;font-size:0.78rem;color:#6B7280;display:flex;gap:16px">
                                <span>Follow-ups: <strong>{s['Follow-ups']}</strong></span>
                                <span>Connected: <strong>{s['FU Connected']}</strong></span>
                                <span>Days w/ 2hr C&O: <strong>{s['Days Meeting 2hr']}/{wd}</strong></span>
                            </div>
                        </div>""", unsafe_allow_html=True)
        else:
            st.info("No coordinator activity found in the selected date range.")

    st.markdown("---")

    # ─────────────────────────────────────────────────────────────────
    # SECTION 4 · PLACES COVERAGE
    # ─────────────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">📍 Places Coverage</div>', unsafe_allow_html=True)

    if oc_p.empty or '_loc' not in oc_p.columns:
        st.info("Upload Outreach & Canvass CSV with a Location column to see places coverage.")
    else:
        def categorize(loc):
            l = loc.lower()
            if any(t in l for t in TRAVIS_KEYS): return 'Travis HS / Wildflower'
            for p in ATX_PRIMARY:
                if p.lower() in l: return 'Primary'
            for p in ATX_OUTER:
                if p.lower() in l: return 'Outer Ring'
            for p in HACA_PROPS:
                if p.lower() in l: return 'HACA'
            return 'Other'

        loc_stats = oc_p.groupby('_loc').agg(
            visits=('_date','count'),
            hours=('_hrs','sum') if '_hrs' in oc_p.columns else ('_date','count')
        ).reset_index()
        loc_stats['Category'] = loc_stats['_loc'].apply(categorize)
        loc_stats['hours']    = loc_stats['hours'].round(1)
        loc_stats = loc_stats.sort_values('hours', ascending=True)

        cat_colors = {
            'Primary':               DARK_BLUE,
            'Outer Ring':            MED_BLUE,
            'HACA':                  '#6B7280',
            'Travis HS / Wildflower': GREEN,
            'Other':                 '#9CA3AF',
        }

        ch_col, sum_col = st.columns([3, 1])

        with ch_col:
            fig = go.Figure()
            for cat in ['Primary','Outer Ring','HACA','Travis HS / Wildflower','Other']:
                sub = loc_stats[loc_stats['Category'] == cat]
                if sub.empty: continue
                fig.add_trace(go.Bar(
                    y=sub['_loc'],
                    x=sub['hours'],
                    name=cat,
                    orientation='h',
                    marker_color=cat_colors.get(cat, '#9CA3AF'),
                    text=[f"{h:.1f}h" for h in sub['hours']],
                    textposition='outside',
                    hovertemplate='<b>%{y}</b><br>Hours: %{x:.1f}h<br>Sessions: %{customdata}<extra></extra>',
                    customdata=sub['visits'],
                ))
            height = max(300, len(loc_stats) * 30 + 100)
            fig.update_layout(
                title=dict(text="Hours Logged by Location", font=dict(size=12, color=DARK_BLUE)),
                xaxis_title="Hours",
                barmode='stack',
                height=height,
                margin=dict(l=0, r=60, t=40, b=30),
                legend=dict(orientation='h', yanchor='bottom', y=1.02, font=dict(size=10)),
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font=dict(size=11),
            )
            st.plotly_chart(fig, use_container_width=True)

        with sum_col:
            st.markdown("**Coverage vs Expected**")
            expected_map = {
                'Primary':               (ATX_PRIMARY,  'Required'),
                'Outer Ring':            (ATX_OUTER,    'Target'),
                'HACA':                  (HACA_PROPS,   'Target'),
                'Travis HS / Wildflower': (['Travis HS / Wildflower'], 'Required'),
            }
            covered_lower = loc_stats['_loc'].str.lower().tolist()

            for cat, (place_list, flag) in expected_map.items():
                cat_data  = loc_stats[loc_stats['Category'] == cat]
                reached   = len(cat_data)
                expected  = len(place_list)
                total_hrs = cat_data['hours'].sum()
                pct       = min(reached / expected * 100, 100) if expected > 0 else 0
                bar_c     = GREEN if pct >= 80 else (AMBER if pct >= 50 else RED_COL)
                badge_cls = 'badge-green' if pct >= 80 else ('badge-amber' if pct >= 50 else 'badge-red')

                st.markdown(f"""
                <div style="margin-bottom:14px">
                    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:3px">
                        <span style="font-size:0.78rem;font-weight:700;color:#374151">{cat}</span>
                        <span class="{badge_cls}">{flag}</span>
                    </div>
                    <div style="background:#E5E7EB;border-radius:4px;height:10px;overflow:hidden;margin-bottom:3px">
                        <div style="background:{bar_c};width:{pct:.0f}%;height:100%;border-radius:4px"></div>
                    </div>
                    <div style="font-size:0.7rem;color:#6B7280">{reached}/{expected} places · {total_hrs:.1f}h total</div>
                </div>""", unsafe_allow_html=True)

            # Not-yet-reached primary locations
            missing_primary = [p for p in ATX_PRIMARY
                               if not any(p.lower() in c for c in covered_lower)]
            if missing_primary:
                st.markdown("**Not yet reached (Primary):**")
                for p in missing_primary:
                    st.markdown(f"<span style='color:#B91C1C;font-size:0.78rem'>✕ {p}</span>", unsafe_allow_html=True)

    st.markdown("---")

    # ─────────────────────────────────────────────────────────────────
    # SECTION 5 · INCIDENTS
    # ─────────────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">🚨 Incidents</div>', unsafe_allow_html=True)

    if inc_df.empty:
        st.info("Upload Incidents CSV to see incident tracking.")
    else:
        n_inc = len(inc_p)
        st.markdown(f"**{n_inc} incident(s)** in selected period &nbsp; (all-time: {len(inc_df)})")

        if not inc_p.empty:
            i1, i2 = st.columns([2, 1])

            with i1:
                num_col  = find_col(inc_p, ['Record number','Record Number','#','ID','Incident ID'])
                type_col = find_col(inc_p, ['Type','Incident Type','Type of Incident','Category'])
                loc_col2 = find_col(inc_p, ['Location','Neighborhood','Location of Incident','Area'])

                show = {}
                if num_col:  show['Record']   = inc_p[num_col]
                if '_date' in inc_p.columns:
                    show['Date'] = inc_p['_date'].dt.strftime('%b %d, %Y')
                if type_col: show['Type']     = inc_p[type_col]
                if loc_col2: show['Location'] = inc_p[loc_col2]
                if '_coord' in inc_p.columns: show['Coordinator'] = inc_p['_coord']

                if show:
                    st.dataframe(pd.DataFrame(show), hide_index=True, use_container_width=True)

            with i2:
                if loc_col2:
                    loc_counts = inc_p[loc_col2].value_counts().head(10)
                    if not loc_counts.empty:
                        fig_i = px.bar(
                            x=loc_counts.values, y=loc_counts.index,
                            orientation='h', title="By Location",
                            color_discrete_sequence=[RED_COL],
                        )
                        fig_i.update_layout(
                            height=250, margin=dict(l=0, r=0, t=30, b=0),
                            showlegend=False, xaxis_title="", yaxis_title="",
                            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                        )
                        st.plotly_chart(fig_i, use_container_width=True)

        elif not inc_df.empty:
            st.success("No incidents in the selected date range.")

# ╔══════════════════════════════════════════════════════════════════╗
# ║  TAB 2 · REPORT GENERATOR                                        ║
# ╚══════════════════════════════════════════════════════════════════╝
with tab2:
    st.subheader("Generate Excel PMRQ Report")
    st.caption("Uses the same CSV files uploaded in the sidebar. Configure settings below, then click Generate.")

    if not REPORT_CORE_AVAILABLE:
        st.error("report_core.py not found — place it in the same folder as this app.")
        st.stop()

    # ── Report config ─────────────────────────────────────────────
    rc1, rc2 = st.columns(2)
    with rc1:
        r_quarter = st.selectbox("Quarter", ["Q1","Q2","Q3","Q4"], index=2, key="r_q")
        r_fy      = st.text_input("Fiscal Year", value="FY 2025-2026", key="r_fy")
    with rc2:
        r_q_start = st.date_input("Quarter Start", value=date(2026, 4, 1),  key="r_qs")
        r_q_end   = st.date_input("Quarter End",   value=date(2026, 6, 30), key="r_qe")

    r_period = st.text_input("Period Label", value="April 1 – June 30, 2026", key="r_period")

    with st.expander("Week Range & Targets"):
        wc1, wc2 = st.columns(2)
        with wc1:
            r_wk_start = st.date_input("Week Start", value=date(2026, 6, 1), key="r_ws")
            r_wk_end   = st.date_input("Week End",   value=date(2026, 6, 7), key="r_we")
        with wc2:
            r_target_5b    = st.number_input("5B Target (%)", 0.0, 1.0, 0.50, 0.01, key="r_5b")
            r_q_la_target  = st.number_input("Q LA Participant Target", 0, value=19, key="r_la")
            r_q_quota_cm   = st.number_input("Q Case-Managed Quota",    0, value=19, key="r_cm")
            r_q_quota_out  = st.number_input("Q Outreach Sessions Quota",0, value=38, key="r_out")
            r_ann_cm       = st.number_input("Annual CM Goal",           0, value=75, key="r_acm")
            r_ann_out      = st.number_input("Annual Outreach Goal",     0, value=150,key="r_aout")

    months_map = {"Q1":["Oct","Nov","Dec"], "Q2":["Jan","Feb","Mar"],
                  "Q3":["Apr","May","Jun"], "Q4":["Jul","Aug","Sep"]}
    months = months_map[r_quarter]

    with st.expander("Monthly Coordinator Quotas"):
        default_quotas = {
            'J. Cooper': {months[0]: 3, months[1]: 3, months[2]: 3},
            'R. Herd':   {months[0]: 3, months[1]: 3, months[2]: 3},
            'N. Dunn':   {months[0]: 0, months[1]: 2, months[2]: 3},
            'K. Young':  {months[0]: 0, months[1]: 0, months[2]: 0},
        }
        quota = {}
        for coord, mq in default_quotas.items():
            st.markdown(f"**{coord}**")
            qc = st.columns(3)
            quota[coord] = {}
            for i, mo in enumerate(months):
                quota[coord][mo] = qc[i].number_input(mo, 0, 20, mq[mo], key=f"q_{coord}_{mo}")

    # ── Generate ──────────────────────────────────────────────────
    missing_files = []
    file_map = {
        'participants': parts_file, 'goals': goals_file, 'outreach': outreach_file,
        'circles': circles_file, 'incidents': incidents_file, 'followup': followup_file,
        'pre': pre_file, 'post': post_file, 'attestation': attest_file,
    }
    for k, f in file_map.items():
        if f is None: missing_files.append(k)

    if missing_files:
        st.info(f"Still need: **{', '.join(missing_files)}** (upload in sidebar)")

    gen_disabled = bool(missing_files)
    if st.button("🚀 Generate Report", type="primary", disabled=gen_disabled, key="gen_btn"):
        with st.spinner("Building report… 15–30 seconds…"):
            try:
                dfs = {}
                for key, f in file_map.items():
                    f.seek(0)
                    try:
                        dfs[key] = pd.read_csv(f, low_memory=False, encoding='utf-8')
                    except UnicodeDecodeError:
                        f.seek(0)
                        dfs[key] = pd.read_csv(f, low_memory=False, encoding='latin-1')

                month_map_vals = {
                    r_q_start.month: months[0],
                    (r_q_start.month % 12) + 1: months[1],
                    ((r_q_start.month + 1) % 12) + 1: months[2],
                }
                cfg = {
                    'QUARTER': r_quarter, 'FISCAL_YEAR': r_fy, 'PERIOD': r_period,
                    'Q_START': pd.Timestamp(r_q_start), 'Q_END': pd.Timestamp(r_q_end),
                    'WEEK_START': pd.Timestamp(r_wk_start), 'WEEK_END': pd.Timestamp(r_wk_end),
                    'TARGET_5B': r_target_5b, 'Q_LA_TARGET': r_q_la_target,
                    'Q_QUOTA_CM': r_q_quota_cm, 'Q_QUOTA_OUTREACH': r_q_quota_out,
                    'Q_QUOTA_TOTAL': r_q_la_target, 'ANNUAL_GOAL_CM': r_ann_cm,
                    'ANNUAL_GOAL_OUT': r_ann_out, 'MONTHS': months,
                    'MONTH_MAP': month_map_vals, 'QUOTA': quota,
                }
                result = report_core.build_report(dfs, cfg)

            except Exception as e:
                st.error(f"Report generation failed: {e}")
                import traceback
                with st.expander("Error details"):
                    st.code(traceback.format_exc())
                st.stop()

        st.success("✅ Report ready!")
        s = result['stats']
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Participants",       s['participants'])
        m2.metric("5B Improvement %",   f"{s['5b_pct']:.1%}")
        m3.metric("Outreach Sessions",  s['outreach'])
        m4.metric("Corrections Needed", s['corrections'])

        st.download_button(
            label="⬇️ Download Excel Report",
            data=result['bytes'],
            file_name=result['fname'],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

st.markdown("---")
st.caption("ATX Peace · Lead Dashboard · Life Anew · Built with Streamlit")
