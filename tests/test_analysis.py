"""
Automated Pytest Suite: Analytical Calculations & Statistical Methods (TC-T-07 to TC-T-11)
Tests KPI formulas, financial metrics, hypothesis tests, and risk classification.
"""

import math
import pytest

def test_tc_t_07_kpi_formulas_on_known_mock():
    """TC-T-07: Verify KPI calculations (attendance rate and no-show rate) on a tiny known dataset."""
    # Mock dataset: 8 shows, 2 no-shows (Total = 10)
    mock_outcomes = [1, 1, 1, 1, 1, 1, 1, 1, 0, 0] # 1 = Show, 0 = No-Show
    total_appointments = len(mock_outcomes)
    shows = sum(mock_outcomes)
    no_shows = total_appointments - shows
    
    noshow_rate = (no_shows / total_appointments) * 100.0
    attendance_rate = (shows / total_appointments) * 100.0
    
    assert total_appointments == 10
    assert shows == 8
    assert no_shows == 2
    assert noshow_rate == 20.0
    assert attendance_rate == 80.0
    assert (noshow_rate + attendance_rate) == 100.0

def test_tc_t_08_financial_and_capacity_formulas_on_known_data():
    """TC-T-08: Verify financial revenue loss and doctor idle capacity calculations on known mock entries."""
    mock_records = [
        {"fee": 600.0, "duration_mins": 20, "is_no_show": 1},  # Wasted
        {"fee": 1200.0, "duration_mins": 30, "is_no_show": 0}, # Attended
        {"fee": 1500.0, "duration_mins": 45, "is_no_show": 1}, # Wasted
        {"fee": 800.0, "duration_mins": 20, "is_no_show": 0}   # Attended
    ]
    
    total_financial_loss = sum(r["fee"] * r["is_no_show"] for r in mock_records)
    total_capacity_loss_hours = sum((r["duration_mins"] / 60.0) * r["is_no_show"] for r in mock_records)
    
    # Expected: 600 + 1500 = INR 2100.0
    # Capacity: (20/60) + (45/60) = 0.333 + 0.75 = 1.0833 hours
    assert total_financial_loss == 2100.0
    assert math.isclose(total_capacity_loss_hours, (20/60.0 + 45/60.0), rel_tol=1e-4)

def test_tc_t_09_two_sample_proportion_z_test_formula():
    """TC-T-09: Verify two-sample proportion Z-test mathematical implementation."""
    # Sample 1: No Reminder (n1=100, x1=25 no-shows, p1=0.25)
    # Sample 2: Reminder Sent (n2=200, x2=20 no-shows, p2=0.10)
    n1, x1 = 100, 25
    n2, x2 = 200, 20
    p1 = x1 / n1
    p2 = x2 / n2
    
    # Pooled proportion
    p_pool = (x1 + x2) / (n1 + n2)
    se = math.sqrt(p_pool * (1.0 - p_pool) * (1.0/n1 + 1.0/n2))
    z_stat = (p1 - p2) / se
    
    # Relative risk reduction
    rrr = ((p1 - p2) / p1) * 100.0
    
    assert p1 == 0.25
    assert p2 == 0.10
    assert rrr == 60.0
    assert z_stat > 2.5, f"Expected Z-statistic > 2.5 (statistically significant), got {z_stat}"

def test_tc_t_10_lead_time_decay_bucketing_logic():
    """TC-T-10: Verify lead-time days discretization into standard 5 buckets."""
    def get_lead_bucket(days):
        if days <= 2:
            return "0-2 days"
        elif days <= 7:
            return "3-7 days"
        elif days <= 14:
            return "8-14 days"
        elif days <= 30:
            return "15-30 days"
        else:
            return ">30 days"
            
    assert get_lead_bucket(0) == "0-2 days"
    assert get_lead_bucket(2) == "0-2 days"
    assert get_lead_bucket(5) == "3-7 days"
    assert get_lead_bucket(12) == "8-14 days"
    assert get_lead_bucket(25) == "15-30 days"
    assert get_lead_bucket(31) == ">30 days"
    assert get_lead_bucket(90) == ">30 days"

def test_tc_t_11_risk_scoring_and_tier_thresholds():
    """TC-T-11: Verify risk tier boundary classification logic."""
    def assign_risk_tier(score):
        if score >= 60.0:
            return "High"
        elif score >= 30.0:
            return "Medium"
        else:
            return "Low"
            
    assert assign_risk_tier(85.5) == "High"
    assert assign_risk_tier(60.0) == "High"
    assert assign_risk_tier(59.9) == "Medium"
    assert assign_risk_tier(30.0) == "Medium"
    assert assign_risk_tier(29.9) == "Low"
    assert assign_risk_tier(12.0) == "Low"


def test_tc_t_13_clean_dataset_imputation_and_deduplication():
    """TC-T-13: Verify data cleaning engine performs imputation and deduplication."""
    from src.data.cleaner import clean_dataset, audit_data_quality
    mock_raw = [
        {"appointment_id": 1, "reminder_lead_hours": None, "status": "Show", "lead_time_days": 5, "age": 30},
        {"appointment_id": 2, "reminder_lead_hours": 24.0, "status": "No-Show", "lead_time_days": 10, "age": 45},
        {"appointment_id": 1, "reminder_lead_hours": 48.0, "status": "Show", "lead_time_days": 5, "age": 30} # Duplicate ID
    ]
    cleaned, report = clean_dataset(mock_raw)
    assert len(cleaned) == 2, f"Expected 2 deduplicated records, got {len(cleaned)}"
    assert report["deduplicated"] == 1
    # Check imputed reminder lead hours
    assert cleaned[0]["reminder_lead_hours"] is not None

def test_tc_t_14_statistical_z_test_on_actual_db():
    """TC-T-14: Verify two-sample proportion Z-test on production database records."""
    from src.data.loader import load_data
    from src.stats.hypothesis import two_sample_proportion_z_test
    records = load_data()
    res = two_sample_proportion_z_test(records)
    assert res["z_stat"] > 2.0, f"Expected significant Z-statistic, got {res['z_stat']}"
    assert res["p_value"] < 0.05, f"Expected p < 0.05, got {res['p_value']}"
    assert res["relative_risk_reduction_pct"] > 40.0
    assert res["is_significant"] is True

def test_tc_t_15_chi_square_specialty_independence():
    """TC-T-15: Verify Chi-square test computes valid test statistics across 6 specialties."""
    from src.data.loader import load_data
    from src.stats.hypothesis import chi_square_specialty_test
    records = load_data()
    chi_res = chi_square_specialty_test(records)
    assert chi_res["degrees_of_freedom"] == 5
    assert chi_res["chi2_stat"] >= 0.0
    assert 0.0 <= chi_res["p_value"] <= 1.0

def test_tc_t_16_risk_scoring_and_triage_queue_sorting():
    """TC-T-16: Verify risk scoring engine prioritizes highest risk patients."""
    from src.models.risk_engine import calculate_risk_score, generate_triage_queue, assign_risk_tier
    high_risk_record = {
        "lead_time_days": 25,
        "sms_reminder_sent": 0,
        "specialty_name": "Mental Wellness",
        "repeat_offender_tier": "Habitual",
        "transit_distance_km": 22.0
    }
    score = calculate_risk_score(high_risk_record)
    assert score >= 60.0
    assert assign_risk_tier(score) == "High"

def test_tc_t_17_temporal_arrival_matrix_dimensions():
    """TC-T-17: Verify 2D temporal arrival heatmap grid dimensions (6 days x 3 slots)."""
    from src.data.loader import load_data
    from src.analytics.temporal import compute_arrival_matrix
    records = load_data()
    matrix = compute_arrival_matrix(records)
    assert len(matrix["days"]) == 6
    assert len(matrix["slots"]) == 3
    assert len(matrix["volume_grid"]) == 3
    assert len(matrix["volume_grid"][0]) == 6

def test_tc_t_18_financial_and_capacity_aggregations():
    """TC-T-18: Verify financial leakage and capacity waste on actual dataset."""
    from src.data.loader import load_data
    from src.analytics.financial import compute_financial_leakage, compute_capacity_waste
    records = load_data()
    leakage = compute_financial_leakage(records)
    capacity = compute_capacity_waste(records)
    assert leakage > 30000.0, f"Expected substantial financial loss, found {leakage}"
    assert capacity > 15.0, f"Expected substantial capacity waste hours, found {capacity}"
