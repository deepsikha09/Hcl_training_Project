"""
Comprehensive Analytics Subpackage for Appointment No-Show Optimization.
"""

from .descriptive import (
    compute_overview_kpis,
    compute_branch_breakdown,
    compute_temporal_trend,
    filter_records
)
from .lead_time import bin_lead_time, compute_lead_time_decay
from .specialty import compute_specialty_breakdown
from .demographics import (
    bin_age_group,
    bin_distance_band,
    compute_age_group_breakdown,
    compute_gender_breakdown,
    compute_distance_breakdown,
    compute_comorbidity_breakdown
)
from .temporal import compute_arrival_matrix
from .financial import (
    compute_financial_leakage,
    compute_capacity_waste,
    compute_monthly_financial_trend,
    compute_doctor_idle_hours
)

__all__ = [
    "compute_overview_kpis",
    "compute_branch_breakdown",
    "compute_temporal_trend",
    "filter_records",
    "bin_lead_time",
    "compute_lead_time_decay",
    "compute_specialty_breakdown",
    "bin_age_group",
    "bin_distance_band",
    "compute_age_group_breakdown",
    "compute_gender_breakdown",
    "compute_distance_breakdown",
    "compute_comorbidity_breakdown",
    "compute_arrival_matrix",
    "compute_financial_leakage",
    "compute_capacity_waste",
    "compute_monthly_financial_trend",
    "compute_doctor_idle_hours"
]
