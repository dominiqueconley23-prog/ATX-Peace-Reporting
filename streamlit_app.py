"""
ATX Peace Report Builder — Streamlit Web App
Upload CSVs → click Generate → download Excel report.
Deploy to Streamlit Cloud: https://streamlit.io/cloud
"""
import streamlit as st
import pandas as pd
import io
from datetime import datetime, date
import report_core

# ── Page config ───────────────────────────────────────────────────
st.set_page_config(
    page_title="ATX Peace Report Builder",
    page_icon="📊",
    layout="wide",
)

# ── Header ────────────────────────────────────────────────────────
st.markdown(
    """
    <div style="background:#1F3864;padding:20px 28px;border-radius:8px;margin-bottom:16px">
        <h1 style="color:#fff;margin:0;font-size:1.8rem">📊 ATX Peace Report Builder</h1>
        <p style="color:#BDD7EE;margin:4px 0 0;font-size:0.9rem">
            Life Anew · CVI Program · PMRQ Reporting
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Sidebar config ────────────────────────────────────────────────
st.sidebar.header("⚙️ Quarter Configuration")

quarter = st.sidebar.selectbox("Quarter", ["Q1", "Q2", "Q3", "Q4"], index=2)
fiscal_year = st.sidebar.text_input("Fiscal Year", value="FY 2025-2026")

# Quarter date presets
preset_ranges = {
    "Q1": ("2025-10-01", "2025-12-31", "Oct 1 – Dec 31, 2025"),
    "Q2": ("2026-01-01", "2026-03-31", "Jan 1 – Mar 31, 2026"),
    "Q3": ("2026-04-01", "2026-06-30", "April 1 – June 30, 2026"),
    "Q4": ("2026-07-01", "2026-09-30", "Jul 1 – Sep 30, 2026"),
}
default_start, default_end, period_label = preset_ranges[quarter]

q_start = st.sidebar.date_input("Quarter Start", value=date.fromisoformat(default_start))
q_end   = st.sidebar.date_input("Quarter End",   value=date.fromisoformat(default_end))
period  = st.sidebar.text_input("Period Label (for report headers)", value=period_label)

st.sidebar.subheader("This-Week Range (for coordinator sheets)")
week_start = st.sidebar.date_input("Week Start", value=date(2026, 6, 1))
week_end   = st.sidebar.date_input("Week End",   value=date(2026, 6, 7))

st.sidebar.subheader("Targets")
target_5b     = st.sidebar.number_input("5B Target (%)", min_value=0.0, max_value=1.0, value=0.50, step=0.01, format="%.2f")
q_la_target   = st.sidebar.number_input(f"{quarter} LA Participant Target", min_value=0, value=19, step=1)
q_quota_cm    = st.sidebar.number_input(f"{quarter} Case-Managed Quota", min_value=0, value=19, step=1)
q_quota_out   = st.sidebar.number_input(f"{quarter} Outreach Sessions Quota", min_value=0, value=38, step=1)
annual_goal_cm  = st.sidebar.number_input("Annual CM Goal (FY)", min_value=0, value=75, step=1)
annual_goal_out = st.sidebar.number_input("Annual Outreach Goal (FY)", min_value=0, value=150, step=1)

st.sidebar.subheader("Monthly Case-Management Quotas")
st.sidebar.caption("Fill in coordinator quota targets per month")

months_labels = {
    "Q1": ["Oct", "Nov", "Dec"],
    "Q2": ["Jan", "Feb", "Mar"],
    "Q3": ["Apr", "May", "Jun"],
    "Q4": ["Jul", "Aug", "Sep"],
}
months = months_labels[quarter]

default_quotas = {
    'J. Cooper': {months[0]: 3, months[1]: 3, months[2]: 3},
    'R. Herd':   {months[0]: 3, months[1]: 3, months[2]: 3},
    'N. Dunn':   {months[0]: 0, months[1]: 2, months[2]: 3},
    'K. Young':  {months[0]: 0, months[1]: 0, months[2]: 0},
}
quota = {}
for coord, mq in default_quotas.items():
    st.sidebar.markdown(f"**{coord}**")
    row_cols = st.sidebar.columns(3)
    quota[coord] = {}
    for i, mo in enumerate(months):
        quota[coord][mo] = row_cols[i].number_input(
            f"{mo}", min_value=0, max_value=20,
            value=mq[mo], key=f"quota_{coord}_{mo}", step=1)

st.sidebar.markdown("---")
st.sidebar.caption("💡 Changes here only affect this run. To change locked YTD data, contact your report admin.")

# ── File upload section ───────────────────────────────────────────
st.subheader("1. Upload CSV Exports from Kintone")
st.caption("Export each form as CSV and upload below. All 9 files required.")

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("**Core Data**")
    parts_file   = st.file_uploader("👤 Participant Folders",      type="csv", key="participants")
    goals_file   = st.file_uploader("🎯 Goals",                    type="csv", key="goals")
    outreach_file = st.file_uploader("🚶 Outreach & Canvass",      type="csv", key="outreach")

with col2:
    st.markdown("**Activity**")
    circles_file  = st.file_uploader("⭕ Circles & Classes",       type="csv", key="circles")
    incidents_file= st.file_uploader("🚨 Incident Responses",      type="csv", key="incidents")
    followup_file = st.file_uploader("📞 Follow-ups",              type="csv", key="followup")

with col3:
    st.markdown("**Assessments**")
    pre_file     = st.file_uploader("📋 Pre-Assessment",           type="csv", key="pre")
    post_file    = st.file_uploader("📋 Post-Assessment",          type="csv", key="post")
    attest_file  = st.file_uploader("✍️ Self-Attestation Form",    type="csv", key="attestation")

# ── Status check ──────────────────────────────────────────────────
all_files = {
    'participants': parts_file,
    'goals':        goals_file,
    'outreach':     outreach_file,
    'circles':      circles_file,
    'incidents':    incidents_file,
    'followup':     followup_file,
    'pre':          pre_file,
    'post':         post_file,
    'attestation':  attest_file,
}
labels = {
    'participants': 'Participant Folders',
    'goals':        'Goals',
    'outreach':     'Outreach & Canvass',
    'circles':      'Circles & Classes',
    'incidents':    'Incident Responses',
    'followup':     'Follow-ups',
    'pre':          'Pre-Assessment',
    'post':         'Post-Assessment',
    'attestation':  'Self-Attestation Form',
}
uploaded = {k for k, v in all_files.items() if v is not None}
missing  = [labels[k] for k in all_files if k not in uploaded]

if missing:
    st.info(f"Still needed: **{', '.join(missing)}**")
else:
    st.success("✅ All 9 files uploaded — ready to generate!")

st.markdown("---")

# ── Generate button ───────────────────────────────────────────────
st.subheader("2. Generate Report")

generate_disabled = bool(missing)
if st.button("🚀 Generate Report", type="primary", disabled=generate_disabled,
             help="Upload all 9 files first" if generate_disabled else "Click to build Excel report"):

    with st.spinner("Building report… this usually takes 15–30 seconds."):
        try:
            # Load all DataFrames
            dfs = {}
            for key, f in all_files.items():
                f.seek(0)
                try:
                    dfs[key] = pd.read_csv(f, low_memory=False)
                except Exception as e:
                    st.error(f"Error reading **{labels[key]}**: {e}")
                    st.stop()

            # Build config
            month_map_vals = {
                q_start.month: months[0],
                (q_start.month % 12) + 1: months[1],
                ((q_start.month + 1) % 12) + 1: months[2],
            }

            cfg = {
                'QUARTER':           quarter,
                'FISCAL_YEAR':       fiscal_year,
                'PERIOD':            period,
                'Q_START':           pd.Timestamp(q_start),
                'Q_END':             pd.Timestamp(q_end),
                'WEEK_START':        pd.Timestamp(week_start),
                'WEEK_END':          pd.Timestamp(week_end),
                'TARGET_5B':         target_5b,
                'Q_LA_TARGET':       q_la_target,
                'Q_QUOTA_CM':        q_quota_cm,
                'Q_QUOTA_OUTREACH':  q_quota_out,
                'Q_QUOTA_TOTAL':     q_la_target,
                'ANNUAL_GOAL_CM':    annual_goal_cm,
                'ANNUAL_GOAL_OUT':   annual_goal_out,
                'MONTHS':            months,
                'MONTH_MAP':         month_map_vals,
                'QUOTA':             quota,
            }

            # Run the report
            result = report_core.build_report(dfs, cfg)

        except Exception as e:
            st.error(f"Report generation failed: {e}")
            import traceback
            with st.expander("Error details"):
                st.code(traceback.format_exc())
            st.stop()

    # ── Stats row ─────────────────────────────────────────────────
    st.markdown("### ✅ Report Ready")
    s = result['stats']
    mc1, mc2, mc3, mc4 = st.columns(4)
    mc1.metric("Participants", s['participants'])
    mc2.metric("5B Improvement %", f"{s['5b_pct']:.1%}")
    mc3.metric("Outreach Sessions", s['outreach'])
    mc4.metric("Corrections Needed", s['corrections'],
               delta=None if s['corrections'] == 0 else f"{s['corrections']} issues",
               delta_color="inverse")

    st.columns(4)[0]  # spacer
    sc1, sc2, sc3, sc4 = st.columns(4)
    sc1.metric("Rollover Cases", s['rollover'])
    sc2.metric("Circles / Classes", s['circles'])
    sc3.metric("Incidents", s['incidents'])
    sc4.metric("Cross-Quarter Dups", s['duplicates'],
               delta=None if s['duplicates'] == 0 else f"⚠ {s['duplicates']} found",
               delta_color="inverse")

    # ── Download button ───────────────────────────────────────────
    st.markdown("---")
    st.download_button(
        label="⬇️ Download Excel Report",
        data=result['bytes'],
        file_name=result['fname'],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    st.caption(f"File: `{result['fname']}`")

    if s['corrections'] > 0:
        st.warning(
            f"⚠️ **{s['corrections']} correction(s) needed.** "
            "Open the report and check the **Corrections — Master** sheet and each coordinator's sheet."
        )
    if s['duplicates'] > 0:
        st.error(
            f"🔴 **{s['duplicates']} cross-quarter duplicate PID(s) found.** "
            "Check the **Duplicate PIDs** sheet before submitting to the City."
        )

st.markdown("---")
st.caption("ATX Peace CVI Report Builder · Life Anew · Built with ❤️ using Streamlit")
