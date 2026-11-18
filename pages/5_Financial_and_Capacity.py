"""
UI-05: Financial Leakage & Capacity Loss Page
Quantifies uncollected consultation revenues and specialist physician idle hours.
"""

import streamlit as st
from src.data import load_data, clean_dataset
from src.analytics import (
    compute_financial_leakage,
    compute_capacity_waste,
    compute_monthly_financial_trend,
    compute_doctor_idle_hours,
    filter_records
)
from src.ui import (
    render_tailwind_header,
    render_kpi_card_row,
    render_insight_card,
    plot_monthly_leakage,
    records_to_markdown_table
)

st.set_page_config(page_title="Financial & Capacity | ApexCare", page_icon="💰", layout="wide")

raw_records = load_data()
dataset, _ = clean_dataset(raw_records)

# Sidebar Filters
st.sidebar.markdown("### 🔍 **Filter Options**")
clinics = ["All Clinics"] + sorted(list({r.get("clinic_name") for r in dataset if r.get("clinic_name")}))
specialties = ["All Specialties"] + sorted(list({r.get("specialty_name") for r in dataset if r.get("specialty_name")}))

sel_clinic = st.sidebar.selectbox("Clinic Branch", clinics, index=0, key="fin_clinic")
sel_spec = st.sidebar.selectbox("Specialty", specialties, index=0, key="fin_spec")

filtered = filter_records(dataset, clinic_name=sel_clinic, specialty_name=sel_spec)

render_tailwind_header(
    title="Financial Leakage Accounting & Clinical Capacity Utilization",
    subtitle=f"Tracking uncollected consultation revenue and physician idle hours across {sel_clinic}",
    badge="UI-05 • Financial & Capacity"
)

# Financial & Capacity KPIs
total_loss = compute_financial_leakage(filtered)
total_hours = compute_capacity_waste(filtered)
total_noshows = sum(1 for r in filtered if r.get("is_no_show") == 1 or r.get("status") == "No-Show")
avg_loss = (total_loss / total_noshows) if total_noshows > 0 else 0.0
annualized_projected = total_loss * 4.0 # Extrapolating sample volume

fin_kpis = [
    {
        "title": "Total Revenue Leakage",
        "value": f"₹{int(total_loss):,}",
        "subtitle": f"{total_noshows} Unattended Consultations",
        "trend": "Direct Loss",
        "color": "red"
    },
    {
        "title": "Physician Hours Wasted",
        "value": f"{round(total_hours, 1)} hrs",
        "subtitle": "Unutilized Clinical Capacity",
        "trend": "Doctor Idleness",
        "color": "amber"
    },
    {
        "title": "Avg Loss per No-Show",
        "value": f"₹{int(avg_loss):,}",
        "subtitle": "Tariff-Weighted Loss",
        "trend": "Unit Cost",
        "color": "blue"
    },
    {
        "title": "Annualized Network Risk",
        "value": f"₹{int(annualized_projected):,}",
        "subtitle": "Projected Across 4 Branches",
        "trend": "Opportunity",
        "color": "amber"
    }
]
render_kpi_card_row(fin_kpis)

st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)

# Monthly Loss Chart
monthly_data = compute_monthly_financial_trend(filtered)
if monthly_data:
    st.plotly_chart(plot_monthly_leakage(monthly_data), use_container_width=True)

# Doctor Idle Hours Table
st.markdown("### 👨‍⚕️ **Physician Consultation Waste & Idle Capacity Roster**")
doctor_data = compute_doctor_idle_hours(filtered)
if doctor_data:
    doc_cols = ["doctor_name", "specialty_name", "clinic_name", "total_appointments", "no_shows", "noshow_rate_pct", "idle_hours", "financial_loss_inr"]
    st.markdown(records_to_markdown_table(doctor_data, columns=doc_cols, limit=15))

# Insights Callout
st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
render_insight_card(
    title="Capacity Reclamation Strategy",
    text=f"Senior specialists in Cardiology and Orthopedics carry the highest hourly tariff (₹1,200 - ₹1,500/slot). Implementing a standby walk-in or emergency triage backfill protocol can reclaim up to 60% of the {round(total_hours, 1)} wasted clinical hours.",
    insight_type="info"
)
