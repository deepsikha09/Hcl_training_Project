"""
Specialty & Departmental Analytics Module
Evaluates attendance, no-show vulnerabilities, and tariffs across medical specialties.
"""

from typing import List, Dict, Any
from collections import defaultdict

def compute_specialty_breakdown(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Calculates operational metrics across clinical departments.
    """
    groups = defaultdict(lambda: {"total": 0, "shows": 0, "noshows": 0, "fees": [], "loss_inr": 0.0, "loss_hours": 0.0})

    for r in records:
        spec = r.get("specialty_name", "General Medicine")
        fee = float(r.get("actual_consultation_fee") or 0.0)
        dur = float(r.get("scheduled_duration_mins") or 0.0) / 60.0
        groups[spec]["total"] += 1
        groups[spec]["fees"].append(fee)

        is_ns = (r.get("is_no_show") == 1 or r.get("status") == "No-Show")
        if is_ns:
            groups[spec]["noshows"] += 1
            groups[spec]["loss_inr"] += fee
            groups[spec]["loss_hours"] += dur
        else:
            groups[spec]["shows"] += 1

    results = []
    for spec, st in groups.items():
        tot = st["total"]
        sh = st["shows"]
        ns = st["noshows"]
        avg_fee = sum(st["fees"]) / len(st["fees"]) if st["fees"] else 0.0
        results.append({
            "specialty_name": spec,
            "total_appointments": tot,
            "shows": sh,
            "no_shows": ns,
            "attendance_rate_pct": round((sh / tot) * 100.0, 2) if tot > 0 else 0.0,
            "noshow_rate_pct": round((ns / tot) * 100.0, 2) if tot > 0 else 0.0,
            "avg_fee_inr": round(avg_fee, 2),
            "financial_loss_inr": round(st["loss_inr"], 2),
            "capacity_loss_hours": round(st["loss_hours"], 1)
        })

    # Sort by no-show rate descending
    results.sort(key=lambda x: x["noshow_rate_pct"], reverse=True)
    return results
