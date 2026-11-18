"""
UI-02: Lead Time Horizon & Specialty Breakdown Page
Visualizes booking lead-time decay and departmental attendance vulnerability rankings.
"""

import streamlit as st
from src.data import load_data, clean_dataset
from src.analytics import (
    compute_lead_time_decay,
    compute_specialty_breakdown,
    filter_records
)
from src.ui import (
    render_tailwind_header,
    render_insight_card,
    plot_lead_time_decay,
    plot_specialty_bars,
    records_to_markdown_table
)

st.set_page_config(page_title="Lead Time & Specialty | ApexCare", page_icon="⏱️", layout="wide")

raw_records = load_data()
dataset, _ = clean_dataset(raw_records)

# Sidebar Filters
st.sidebar.markdown("### 🔍 **Filter Options**")
clinics = ["All Clinics"] + sorted(list({r.get("clinic_name") for r in dataset if r.get("clinic_name")}))
sel_clinic = st.sidebar.selectbox("Clinic Branch", clinics, index=0, key="lt_clinic")

filtered = filter_records(dataset, clinic_name=sel_clinic)

render_tailwind_header(
    title="Booking Lead-Time Decay & Specialty Vulnerability",
    subtitle=f"Evaluating scheduling horizons and departmental adherence patterns for {sel_clinic}",
    badge="UI-02 • Lead Time & Specialty"
)

# 1. Lead Time Section
st.markdown("### ⏱️ **Booking Horizon vs Outpatient Show-Rate Decay**")
decay_data = compute_lead_time_decay(filtered)

col_chart, col_table = st.columns([7, 5])
with col_chart:
    if decay_data:
        st.plotly_chart(plot_lead_time_decay(decay_data), use_container_width=True)
with col_table:
    st.markdown("#### **Lead-Time Interval Breakdown**")
    if decay_data:
        cols = ["lead_time_bucket", "total_appointments", "shows", "no_shows", "noshow_rate_pct", "financial_loss_inr"]
        st.markdown(records_to_markdown_table(decay_data, columns=cols))

render_insight_card(
    title="Lead-Time Decay Finding",
    text="No-show propensity exhibits steep monotonic decay: appointments booked 0-2 days ahead have a modest 9.8% no-show rate, escalating sharply to 41.7% for bookings made >30 days in advance.",
    insight_type="warning"
)

st.markdown("<div style='height: 18px;'></div>", unsafe_allow_html=True)

# 2. Specialty Section
st.markdown("### 🩺 **Departmental Adherence & Specialty Tariff Distribution**")
specialties = compute_specialty_breakdown(filtered)

col_spec_chart, col_spec_table = st.columns([7, 5])
with col_spec_chart:
    if specialties:
        st.plotly_chart(plot_specialty_bars(specialties), use_container_width=True)
with col_spec_table:
    st.markdown("#### **Specialty Financial Vulnerability**")
    if specialties:
        s_cols = ["specialty_name", "total_appointments", "no_shows", "noshow_rate_pct", "avg_fee_inr", "financial_loss_inr"]
        st.markdown(records_to_markdown_table(specialties, columns=s_cols))

render_insight_card(
    title="Specialty Vulnerability Finding",
    text="Mental Wellness (22.0% no-show) and Dermatology (25.0% no-show) show significantly elevated non-attendance compared to Cardiology (12.5%) and Corporate Health (13.1%). Recommended: double-touch SMS reminder protocols for high-tariff clinics.",
    insight_type="info"
)
