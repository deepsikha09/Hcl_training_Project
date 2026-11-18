"""
UI-06: Predictive Risk Triage Queue Page
Provides front-desk receptionists with advance calling rosters and repeat-offender Pareto analytics.
"""

import streamlit as st
import io
import csv
from src.data import load_data, clean_dataset
from src.models import enrich_records_with_risk, generate_triage_queue
from src.ui import (
    render_tailwind_header,
    render_insight_card,
    plot_risk_donut,
    plot_pareto_curve,
    records_to_markdown_table
)

st.set_page_config(page_title="Risk Triage | ApexCare", page_icon="📋", layout="wide")

raw_records = load_data()
dataset, _ = clean_dataset(raw_records)
enriched_dataset = enrich_records_with_risk(dataset)

render_tailwind_header(
    title="Predictive Risk Triage & Front-Desk Calling Queue",
    subtitle="Advance patient risk stratification and prioritized phone-confirmation roster",
    badge="UI-06 • Risk Triage"
)

# Visual Stratification Row
col_donut, col_pareto = st.columns([5, 6])
with col_donut:
    st.plotly_chart(plot_risk_donut(enriched_dataset), use_container_width=True)
with col_pareto:
    st.plotly_chart(plot_pareto_curve(enriched_dataset), use_container_width=True)

render_insight_card(
    title="Pareto Loss Concentration Finding",
    text="The top 20% of chronic non-attending patients account for nearly 65% of cumulative consultation financial losses. Targeting this habitual cohort with personal phone calls generates maximum return on front-desk staff time.",
    insight_type="warning"
)

st.markdown("<div style='height: 18px;'></div>", unsafe_allow_html=True)

# Front-Desk Action Queue
st.markdown("### 📞 **Daily Front-Desk Proactive Calling Roster**")

c_filter1, c_filter2 = st.columns([4, 4])
with c_filter1:
    tier_filter = st.selectbox("Filter Risk Tier", ["High Only (Score >= 60)", "High & Medium (Score >= 30)", "All Appointments"], index=0)
with c_filter2:
    search_query = st.text_input("Search Patient MRN, Doctor, or Specialty", "")

# Apply filters
min_tier_code = "High"
if "Medium" in tier_filter:
    min_tier_code = "Medium"
elif "All" in tier_filter:
    min_tier_code = "Low"

queue_records = generate_triage_queue(enriched_dataset, min_tier=min_tier_code, limit=100)

if search_query:
    q = search_query.lower()
    queue_records = [
        r for r in queue_records
        if q in str(r.get("patient_mrn", "")).lower()
        or q in str(r.get("doctor_name", "")).lower()
        or q in str(r.get("specialty_name", "")).lower()
    ]

# Add action column
for r in queue_records:
    tier = r.get("risk_tier", "Low")
    if tier == "High":
        r["action_protocol"] = "Mandatory 48h Phone Confirmation"
    elif tier == "Medium":
        r["action_protocol"] = "Automated 24h SMS Reminder"
    else:
        r["action_protocol"] = "Standard 12h App Notification"

st.markdown(f"**Showing {len(queue_records)} prioritized appointments requiring advance staff action:**")

table_cols = ["booking_reference", "patient_mrn", "doctor_name", "specialty_name", "lead_time_days", "risk_score", "risk_tier", "action_protocol"]
st.markdown(records_to_markdown_table(queue_records, columns=table_cols, limit=20))

# Export to CSV
if queue_records:
    csv_buf = io.StringIO()
    writer = csv.DictWriter(csv_buf, fieldnames=table_cols)
    writer.writeheader()
    for r in queue_records:
        writer.writerow({c: r.get(c, "") for c in table_cols})
    csv_bytes = csv_buf.getvalue()

    st.download_button(
        label="📥 Download Front-Desk Call Roster (CSV)",
        data=csv_bytes,
        file_name="apexcare_calling_roster.csv",
        mime="text/csv"
    )
