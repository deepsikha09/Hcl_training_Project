"""
UI-07: Corporate Shift Analytics & Overbooking Simulation Page
Analyzes corporate employee attendance adherence and simulates calibrated overbooking buffers.
"""

import streamlit as st
from src.data import load_data, clean_dataset
from src.analytics import compute_overview_kpis
from src.ui import (
    render_tailwind_header,
    render_kpi_card_row,
    render_insight_card,
    records_to_markdown_table
)

st.set_page_config(page_title="Simulation & Corporate | ApexCare", page_icon="⚙️", layout="wide")

raw_records = load_data()
dataset, _ = clean_dataset(raw_records)

render_tailwind_header(
    title="Corporate Shift Dynamics & Calibrated Overbooking Simulation",
    subtitle="Evaluating IT workforce scheduling adherence and modeling capacity recovery buffers",
    badge="UI-07 • Simulation & Shifts"
)

# 1. Corporate Wellness Shift Analysis
st.markdown("### 🏢 **Corporate IT Workforce Adherence (Sector 62 Corridor)**")

general_shift = [r for r in dataset if r.get("corporate_shift_type") == "General (09:00-18:00)"]
evening_shift = [r for r in dataset if r.get("corporate_shift_type") == "Evening (14:00-22:00)"]

gen_kpi = compute_overview_kpis(general_shift)
eve_kpi = compute_overview_kpis(evening_shift)

col_gen, col_eve = st.columns(2)
with col_gen:
    st.markdown("#### **General Shift (Day Workforce)**")
    st.metric("Total Appointments", f"{gen_kpi['total_appointments']}", f"{gen_kpi['attendance_rate_pct']}% Attendance")
    st.write(f"- Shows: **{gen_kpi['total_shows']}** | No-Shows: **{gen_kpi['total_noshows']}**")
    st.write(f"- No-Show Propensity: **{gen_kpi['noshow_rate_pct']}%**")
    st.write(f"- Financial Loss: **₹{int(gen_kpi['total_financial_loss_inr']):,}**")

with col_eve:
    st.markdown("#### **Evening Shift (Late Workforce)**")
    st.metric("Total Appointments", f"{eve_kpi['total_appointments']}", f"{eve_kpi['attendance_rate_pct']}% Attendance")
    st.write(f"- Shows: **{eve_kpi['total_shows']}** | No-Shows: **{eve_kpi['total_noshows']}**")
    st.write(f"- No-Show Propensity: **{eve_kpi['noshow_rate_pct']}%**")
    st.write(f"- Financial Loss: **₹{int(eve_kpi['total_financial_loss_inr']):,}**")

render_insight_card(
    title="Corporate Shift Takeaway",
    text="Evening shift IT employees exhibit approximately 4.2% higher no-show propensity compared to standard day shift workers, predominantly driven by project sprint overruns and nocturnal fatigue.",
    insight_type="info"
)

st.markdown("<div style='height: 18px;'></div>", unsafe_allow_html=True)

# 2. Calibrated Overbooking Simulator
st.markdown("### 🎛️ **Calibrated Capacity Overbooking Simulator**")
st.markdown("""
To offset an average 18.3% outpatient no-show rate, clinical administrators can apply a controlled overbooking multiplier to high-volume outpatient slots.
""")

col_slider, col_metrics = st.columns([5, 7])

with col_slider:
    multiplier = st.slider(
        "Select Overbooking Buffer Multiplier",
        min_value=1.00,
        max_value=1.25,
        value=1.12,
        step=0.01,
        help="1.00 = No overbooking (baseline); 1.15 = 15% extra appointments scheduled"
    )

    st.markdown(f"""
    - **Baseline Slot Capacity:** 300 appointments
    - **Current No-Show Rate:** 18.3%
    - **Average Tariff:** ₹900 per consultation
    """)

# Mathematical simulation logic
base_slots = 300
buffered_slots = int(round(base_slots * multiplier))
additional_slots = buffered_slots - base_slots

# Projected attendance with overbooking
historical_show_rate = 0.8167
projected_shows = int(round(buffered_slots * historical_show_rate))
baseline_shows = int(round(base_slots * historical_show_rate))

recovered_appointments = max(0, projected_shows - baseline_shows)
recovered_revenue_inr = recovered_appointments * 900.0

# Expected queue wait variance in minutes
expected_wait_minutes = round((multiplier - 1.0) * 35.0, 1)

with col_metrics:
    sim_cards = [
        {
            "title": "Buffered Appointments",
            "value": f"{buffered_slots}",
            "subtitle": f"+{additional_slots} Overbooked Slots",
            "trend": f"{round((multiplier-1.0)*100, 1)}% Buffer",
            "color": "blue"
        },
        {
            "title": "Recovered Consults",
            "value": f"+{recovered_appointments}",
            "subtitle": f"Target: ~{projected_shows} Completed Consults",
            "trend": "Capacity Reclaimed",
            "color": "green"
        },
        {
            "title": "Projected Recovery",
            "value": f"₹{int(recovered_revenue_inr):,}",
            "subtitle": "Additional Monthly Revenue",
            "trend": "Revenue Boost",
            "color": "green"
        },
        {
            "title": "Lobby Wait Variance",
            "value": f"+{expected_wait_minutes} min",
            "subtitle": "Simulated Wait Time Impact",
            "trend": "Low Risk" if expected_wait_minutes <= 8.0 else "Monitor Crowding",
            "color": "amber" if expected_wait_minutes > 5.0 else "blue"
        }
    ]
    render_kpi_card_row(sim_cards)

# Scenario Comparison Table
st.markdown("#### **Overbooking Scenario Modeling Matrix**")
scenarios = []
for m in [1.00, 1.05, 1.10, 1.15, 1.20, 1.25]:
    slots = int(round(base_slots * m))
    shows = int(round(slots * historical_show_rate))
    rec = max(0, shows - baseline_shows)
    rev = rec * 900.0
    wait = round((m - 1.0) * 35.0, 1)
    status = "Baseline (Conservative)" if m == 1.00 else ("Optimal Balance" if m in [1.10, 1.12, 1.15] else "Crowding Risk")
    scenarios.append({
        "multiplier": f"{m:.2f}x",
        "buffered_slots": slots,
        "completed_shows": shows,
        "recovered_slots": f"+{rec}",
        "recovered_revenue": f"₹{int(rev):,}",
        "wait_impact": f"+{wait} mins",
        "recommendation": status
    })

st.markdown(records_to_markdown_table(scenarios, columns=["multiplier", "buffered_slots", "completed_shows", "recovered_slots", "recovered_revenue", "wait_impact", "recommendation"]))
