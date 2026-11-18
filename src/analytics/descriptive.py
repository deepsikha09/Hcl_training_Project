"""
Descriptive Analytics & High-Level Operational KPIs Module
Calculates aggregate metrics, branch performance, and time-series rolling averages.
"""

from typing import List, Dict, Any, Optional
from collections import defaultdict
import datetime

def filter_records(
    records: List[Dict[str, Any]],
    clinic_name: Optional[str] = None,
    specialty_name: Optional[str] = None,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    shift_type: Optional[str] = None
) -> List[Dict[str, Any]]:
    """Applies global interactive dashboard filters across analytical records."""
    filtered = []
    for r in records:
        if clinic_name and clinic_name != "All Clinics" and r.get("clinic_name") != clinic_name:
            continue
        if specialty_name and specialty_name != "All Specialties" and r.get("specialty_name") != specialty_name:
            continue
        if shift_type and shift_type != "All Shifts" and r.get("corporate_shift_type") != shift_type:
            continue
        app_date = r.get("appointment_date")
        if start_date and app_date and app_date < str(start_date):
            continue
        if end_date and app_date and app_date > str(end_date):
            continue
        filtered.append(r)
    return filtered

def compute_overview_kpis(records: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Computes foundational operational KPIs matching BR-01 and TC-T-07."""
    total = len(records)
    if total == 0:
        return {
            "total_appointments": 0,
            "total_shows": 0,
            "total_noshows": 0,
            "attendance_rate_pct": 0.0,
            "noshow_rate_pct": 0.0,
            "total_financial_loss_inr": 0.0,
            "total_capacity_loss_hours": 0.0,
            "avg_lead_time_days": 0.0
        }

    no_shows = sum(1 for r in records if r.get("is_no_show") == 1 or r.get("status") == "No-Show")
    shows = total - no_shows
    noshow_rate = round((no_shows / total) * 100.0, 2)
    attendance_rate = round((shows / total) * 100.0, 2)

    financial_loss = sum(float(r.get("actual_consultation_fee") or 0.0) for r in records if r.get("is_no_show") == 1)
    capacity_loss_hours = sum((float(r.get("scheduled_duration_mins") or 0.0) / 60.0) for r in records if r.get("is_no_show") == 1)
    
    lead_times = [float(r.get("lead_time_days") or 0.0) for r in records]
    avg_lead_time = round(sum(lead_times) / total, 1) if lead_times else 0.0

    return {
        "total_appointments": total,
        "total_shows": shows,
        "total_noshows": no_shows,
        "attendance_rate_pct": attendance_rate,
        "noshow_rate_pct": noshow_rate,
        "total_financial_loss_inr": round(financial_loss, 2),
        "total_capacity_loss_hours": round(capacity_loss_hours, 1),
        "avg_lead_time_days": avg_lead_time
    }

def compute_branch_breakdown(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Groups appointments by clinic branch to compute comparative attendance metrics."""
    groups = defaultdict(lambda: {"total": 0, "shows": 0, "noshows": 0, "loss_inr": 0.0, "loss_hours": 0.0})

    for r in records:
        branch = r.get("clinic_name", "Unknown Branch")
        groups[branch]["total"] += 1
        is_ns = (r.get("is_no_show") == 1 or r.get("status") == "No-Show")
        if is_ns:
            groups[branch]["noshows"] += 1
            groups[branch]["loss_inr"] += float(r.get("actual_consultation_fee") or 0.0)
            groups[branch]["loss_hours"] += float(r.get("scheduled_duration_mins") or 0.0) / 60.0
        else:
            groups[branch]["shows"] += 1

    results = []
    for branch, stats in groups.items():
        tot = stats["total"]
        ns = stats["noshows"]
        sh = stats["shows"]
        results.append({
            "clinic_name": branch,
            "total_appointments": tot,
            "shows": sh,
            "no_shows": ns,
            "attendance_rate_pct": round((sh / tot) * 100.0, 2) if tot > 0 else 0.0,
            "noshow_rate_pct": round((ns / tot) * 100.0, 2) if tot > 0 else 0.0,
            "financial_loss_inr": round(stats["loss_inr"], 2),
            "capacity_loss_hours": round(stats["loss_hours"], 1)
        })

    # Sort by total volume descending
    results.sort(key=lambda x: x["total_appointments"], reverse=True)
    return results

def compute_temporal_trend(records: List[Dict[str, Any]], window: int = 7) -> List[Dict[str, Any]]:
    """
    Computes daily volume, shows, no-shows, and rolling moving averages.
    """
    day_counts = defaultdict(lambda: {"total": 0, "shows": 0, "noshows": 0})
    for r in records:
        d = r.get("appointment_date")
        if d:
            day_counts[d]["total"] += 1
            if r.get("is_no_show") == 1 or r.get("status") == "No-Show":
                day_counts[d]["noshows"] += 1
            else:
                day_counts[d]["shows"] += 1

    sorted_dates = sorted(day_counts.keys())
    trend = []
    rolling_buffer = []

    for d in sorted_dates:
        tot = day_counts[d]["total"]
        ns = day_counts[d]["noshows"]
        sh = day_counts[d]["shows"]
        ns_rate = (ns / tot * 100.0) if tot > 0 else 0.0
        rolling_buffer.append(ns_rate)
        if len(rolling_buffer) > window:
            rolling_buffer.pop(0)
        
        m_avg = sum(rolling_buffer) / len(rolling_buffer)
        trend.append({
            "date": d,
            "total": tot,
            "shows": sh,
            "no_shows": ns,
            "noshow_rate_pct": round(ns_rate, 2),
            "moving_avg_pct": round(m_avg, 2)
        })

    return trend
