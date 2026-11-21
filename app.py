"""
ApexCare Healthcare Analytics & Optimization Platform
Main Application Entrypoint & Executive Overview Dashboard
"""

import streamlit as st
import datetime
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
    plot_branch_comparison
)

# 1. Page Configuration
st.set_page_config(
    page_title="ApexCare Outpatient Analytics",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Data Loading & Ingestion
try:
    raw_records = load_data()
    dataset, clean_report = clean_dataset(raw_records)
except Exception as e:
    st.error(f"Error loading database records: {e}")
    st.stop()

# 3. Sidebar Global Filters & Session State
st.sidebar.image("https://images.unsplash.com/photo-1519494026892-80bbd2d6fd0d?w=400&auto=format&fit=crop&q=60", use_container_width=True)
st.sidebar.markdown("### 🏥 **ApexCare Operations**")
st.sidebar.markdown("<p style='font-size:12px; color:#64748B;'>Noida Outpatient Network (Sectors 62, 18, 137, Greater Noida)</p>", unsafe_allow_html=True)
st.sidebar.markdown("---")

# Extract unique filter values
clinics = ["All Clinics"] + sorted(list({r.get("clinic_name") for r in dataset if r.get("clinic_name")}))
specialties = ["All Specialties"] + sorted(list({r.get("specialty_name") for r in dataset if r.get("specialty_name")}))
shifts = ["All Shifts"] + sorted(list({r.get("corporate_shift_type") for r in dataset if r.get("corporate_shift_type")}))

# Filter widgets
selected_clinic = st.sidebar.selectbox("Filter by Clinic Branch", clinics, index=0)
selected_specialty = st.sidebar.selectbox("Filter by Medical Specialty", specialties, index=0)
selected_shift = st.sidebar.selectbox("Filter by Corporate Shift", shifts, index=0)

# Filter dataset
filtered_data = filter_records(
    dataset,
    clinic_name=selected_clinic,
    specialty_name=selected_specialty,
    shift_type=selected_shift
)

st.sidebar.markdown("---")
st.sidebar.info(
    f"**Active Records:** {len(filtered_data)} of {len(dataset)}\n\n"
    f"**Database:** SQLite (`data/app.db`)\n\n"
    f"**Project ID:** P101\n\n"
    f"**Batch:** ABES_OND26 MCA 2ND YEAR_DATA ANALYTICS_BATCH 01\n\n"
    f"**Student:** Deepshikha (`2500320140047`)\n\n"
    f"**Mentor:** Dinesh Sereno"
)

# 4. Header Banner
render_tailwind_header(
    title="Executive Outpatient Analytics & No-Show Diagnostic Platform",
    subtitle=f"Real-time operational monitoring across {selected_clinic} • {selected_specialty}",
    badge="Enterprise Healthcare"
)

# 5. Core Operational KPIs Row
kpis = compute_overview_kpis(filtered_data)

kpi_cards = [
    {
        "title": "Total Appointments",
        "value": f"{kpis['total_appointments']:,}",
        "subtitle": f"{kpis['total_shows']} Attended | {kpis['total_noshows']} No-Shows",
        "trend": "100% Monitored",
        "color": "blue"
    },
    {
        "title": "Attendance Rate",
        "value": f"{kpis['attendance_rate_pct']}%",
        "subtitle": "Target: 85.0% Adherence",
        "trend": "Acceptable" if kpis['attendance_rate_pct'] >= 80.0 else "Action Required",
        "color": "green" if kpis['attendance_rate_pct'] >= 85.0 else "blue"
    },
    {
        "title": "No-Show Rate",
        "value": f"{kpis['noshow_rate_pct']}%",
        "subtitle": f"Target: < 15.0% Non-Attendance",
        "trend": "Vulnerable" if kpis['noshow_rate_pct'] > 18.0 else "Controlled",
        "color": "red" if kpis['noshow_rate_pct'] > 18.0 else "amber"
    },
    {
        "title": "Revenue Leakage",
        "value": f"₹{int(kpis['total_financial_loss_inr']):,}",
        "subtitle": f"{kpis['total_capacity_loss_hours']} Doctor Hours Idle",
        "trend": "Uncollected",
        "color": "amber"
    }
]

render_kpi_card_row(kpi_cards)

st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

# 6. Primary Visualizations Grid
col_trend, col_branch = st.columns([6, 5])

with col_trend:
    trend_records = compute_temporal_trend(filtered_data)
    if trend_records:
        fig_trend = plot_rolling_trend(trend_records)
        st.plotly_chart(fig_trend, use_container_width=True)
    else:
        st.warning("No time-series data available for active filters.")

with col_branch:
    branch_records = compute_branch_breakdown(filtered_data)
    if branch_records:
        fig_branch = plot_branch_comparison(branch_records)
        st.plotly_chart(fig_branch, use_container_width=True)
    else:
        st.warning("No branch records available for active filters.")

# 7. Executive Action Insights
st.markdown("### 💡 **Executive Strategic Insights & Clinical Directives**")

col_ins1, col_ins2, col_ins3 = st.columns(3)

with col_ins1:
    render_insight_card(
        title="Digital Reminder Lift",
        text="Automated SMS notifications reduce outpatient no-shows by 48.8% (p = 0.0051). Mandatory reminder integration recommended for all long lead-time bookings.",
        insight_type="success"
    )

with col_ins2:
    render_insight_card(
        title="Specialty Vulnerability",
        text="Mental Wellness and Dermatology experience elevated no-show rates (22%-25%). Recommend implementing 24-hour pre-appointment phone confirmations.",
        insight_type="warning"
    )

with col_ins3:
    render_insight_card(
        title="Capacity Recovery",
        text=f"Outpatient non-attendance wasted {kpis['total_capacity_loss_hours']} specialist hours. Applying calibrated overbooking (+15%) can recover ~₹28,500 monthly.",
        insight_type="info"
    )

st.markdown("---")
st.markdown("<p style='text-align: center; color: #94A3B8; font-size: 12px;'>ApexCare Health Network • Outpatient Operations Decision Support System • MCA Capstone 2026</p>", unsafe_allow_html=True)
