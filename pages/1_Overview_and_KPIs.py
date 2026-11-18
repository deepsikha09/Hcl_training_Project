"""
UI-01: Executive Overview & Baseline KPIs Page
Multi-branch outpatient attendance benchmarking, rolling trends, and branch tables.
"""

import streamlit as st
from src.data import load_data, clean_dataset
from src.analytics import (
    compute_overview_kpis,
    compute_branch_breakdown,
    compute_temporal_trend,
    filter_records
)
from src.ui import (
    render_tailwind_header,
    render_kpi_card_row,
    render_insight_card,
    plot_rolling_trend,
    plot_branch_comparison,
    records_to_markdown_table
)

st.set_page_config(page_title="Overview & KPIs | ApexCare", page_icon="📊", layout="wide")

raw_records = load_data()
dataset, _ = clean_dataset(raw_records)

# Sidebar Filters
st.sidebar.markdown("### 🔍 **Filter Overview**")
clinics = ["All Clinics"] + sorted(list({r.get("clinic_name") for r in dataset if r.get("clinic_name")}))
specialties = ["All Specialties"] + sorted(list({r.get("specialty_name") for r in dataset if r.get("specialty_name")}))

sel_clinic = st.sidebar.selectbox("Clinic Branch", clinics, index=0, key="ov_clinic")
sel_spec = st.sidebar.selectbox("Specialty", specialties, index=0, key="ov_spec")

filtered = filter_records(dataset, clinic_name=sel_clinic, specialty_name=sel_spec)

render_tailwind_header(
    title="Executive Outpatient Performance & Branch Benchmarking",
    subtitle=f"Aggregated operational view across {sel_clinic} • {sel_spec}",
    badge="UI-01 • Core KPIs"
)

# Metric Cards
kpis = compute_overview_kpis(filtered)
kpi_cards = [
    {"title": "Total Volume", "value": f"{kpis['total_appointments']:,}", "subtitle": f"{kpis['total_shows']} Attended", "trend": "Active", "color": "blue"},
    {"title": "Attendance Rate", "value": f"{kpis['attendance_rate_pct']}%", "subtitle": "Target: 85.0%", "trend": "Baseline", "color": "green" if kpis['attendance_rate_pct'] >= 85.0 else "blue"},
    {"title": "No-Show Rate", "value": f"{kpis['noshow_rate_pct']}%", "subtitle": f"{kpis['total_noshows']} Missed Consults", "trend": "High" if kpis['noshow_rate_pct'] > 18.0 else "Normal", "color": "red" if kpis['noshow_rate_pct'] > 18.0 else "amber"},
    {"title": "Uncollected Revenue", "value": f"₹{int(kpis['total_financial_loss_inr']):,}", "subtitle": f"{kpis['total_capacity_loss_hours']} Idle Hours", "trend": "Leakage", "color": "amber"}
]
render_kpi_card_row(kpi_cards)

st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)

# Charts Row
c1, c2 = st.columns([6, 5])
with c1:
    trend = compute_temporal_trend(filtered)
    if trend:
        st.plotly_chart(plot_rolling_trend(trend), use_container_width=True)
with c2:
    branches = compute_branch_breakdown(filtered)
    if branches:
        st.plotly_chart(plot_branch_comparison(branches), use_container_width=True)

# Detailed Branch Performance Table
st.markdown("### 📋 **Clinical Facility Operational Scorecard**")
if branches:
    table_cols = ["clinic_name", "total_appointments", "shows", "no_shows", "noshow_rate_pct", "financial_loss_inr", "capacity_loss_hours"]
    st.markdown(records_to_markdown_table(branches, columns=table_cols))

# Operational Notes
st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
render_insight_card(
    title="Branch Variance Finding",
    text="Sector 62 (IT Corridor Hub) demonstrates highest outpatient volume with balanced attendance, whereas Greater Noida branch experiences extended transit distances contributing to higher non-attendance.",
    insight_type="info"
)
