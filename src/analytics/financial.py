"""
Financial Leakage & Clinical Capacity Accounting Module
Calculates uncollected consultation revenue and physician idle consultation hours.
"""

from typing import List, Dict, Any
from collections import defaultdict

def compute_financial_leakage(records: List[Dict[str, Any]]) -> float:
    """Calculates total uncollected consultation fee revenue in INR matching TC-T-08."""
    return sum(float(r.get("actual_consultation_fee") or 0.0) for r in records if r.get("is_no_show") == 1 or r.get("status") == "No-Show")

def compute_capacity_waste(records: List[Dict[str, Any]]) -> float:
    """Calculates total doctor consultation hours wasted matching TC-T-08."""
    return sum((float(r.get("scheduled_duration_mins") or 0.0) / 60.0) for r in records if r.get("is_no_show") == 1 or r.get("status") == "No-Show")

def compute_monthly_financial_trend(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Groups financial loss and wasted capacity by calendar month (YYYY-MM)."""
    months = defaultdict(lambda: {"total_loss_inr": 0.0, "total_hours": 0.0, "no_shows": 0, "total_appts": 0})

    for r in records:
        date_str = r.get("appointment_date", "")
        month_key = date_str[:7] if len(date_str) >= 7 else "Unknown"
        months[month_key]["total_appts"] += 1

        if r.get("is_no_show") == 1 or r.get("status") == "No-Show":
            months[month_key]["no_shows"] += 1
            months[month_key]["total_loss_inr"] += float(r.get("actual_consultation_fee") or 0.0)
            months[month_key]["total_hours"] += float(r.get("scheduled_duration_mins") or 0.0) / 60.0

    sorted_months = sorted([m for m in months.keys() if m != "Unknown"])
    results = []
    for m in sorted_months:
        st = months[m]
        results.append({
            "month": m,
            "total_appointments": st["total_appts"],
            "no_shows": st["no_shows"],
            "financial_loss_inr": round(st["total_loss_inr"], 2),
            "capacity_loss_hours": round(st["total_hours"], 1)
        })

    return results

def compute_doctor_idle_hours(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Aggregates unutilized consultation slots and revenue leakage per physician."""
    doctors = defaultdict(lambda: {
        "doctor_name": "",
        "specialty": "",
        "clinic": "",
        "total": 0,
        "noshows": 0,
        "loss_inr": 0.0,
        "loss_hours": 0.0
    })

    for r in records:
        doc = r.get("doctor_name", "Unknown Doctor")
        doctors[doc]["doctor_name"] = doc
        doctors[doc]["specialty"] = r.get("specialty_name", "")
        doctors[doc]["clinic"] = r.get("clinic_name", "")
        doctors[doc]["total"] += 1

        if r.get("is_no_show") == 1 or r.get("status") == "No-Show":
            doctors[doc]["noshows"] += 1
            doctors[doc]["loss_inr"] += float(r.get("actual_consultation_fee") or 0.0)
            doctors[doc]["loss_hours"] += float(r.get("scheduled_duration_mins") or 0.0) / 60.0

    results = []
    for doc, st in doctors.items():
        tot = st["total"]
        ns = st["noshows"]
        results.append({
            "doctor_name": doc,
            "specialty_name": st["specialty"],
            "clinic_name": st["clinic"],
            "total_appointments": tot,
            "no_shows": ns,
            "noshow_rate_pct": round((ns / tot) * 100.0, 2) if tot > 0 else 0.0,
            "idle_hours": round(st["loss_hours"], 1),
            "financial_loss_inr": round(st["loss_inr"], 2)
        })

    results.sort(key=lambda x: x["financial_loss_inr"], reverse=True)
    return results
