"""
Lead Time Analytics & Booking Decay Engine
Analyzes the decay of attendance adherence as booking horizon expands.
"""

from typing import List, Dict, Any
from collections import defaultdict

BUCKET_ORDER = ["0-2 days", "3-7 days", "8-14 days", "15-30 days", ">30 days"]

def bin_lead_time(days: float) -> str:
    """Discretizes booking lead time in days into 5 standardized intervals matching TC-T-10."""
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

def compute_lead_time_decay(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Computes volume, show rate, no-show rate, and financial loss per lead-time interval.
    """
    groups = {b: {"total": 0, "shows": 0, "noshows": 0, "loss_inr": 0.0} for b in BUCKET_ORDER}

    for r in records:
        lead = float(r.get("lead_time_days") or 0.0)
        bucket = r.get("lead_time_bucket") or bin_lead_time(lead)
        if bucket not in groups:
            bucket = bin_lead_time(lead)
            
        groups[bucket]["total"] += 1
        is_ns = (r.get("is_no_show") == 1 or r.get("status") == "No-Show")
        if is_ns:
            groups[bucket]["noshows"] += 1
            groups[bucket]["loss_inr"] += float(r.get("actual_consultation_fee") or 0.0)
        else:
            groups[bucket]["shows"] += 1

    results = []
    for b in BUCKET_ORDER:
        st = groups[b]
        tot = st["total"]
        sh = st["shows"]
        ns = st["noshows"]
        results.append({
            "lead_time_bucket": b,
            "total_appointments": tot,
            "shows": sh,
            "no_shows": ns,
            "attendance_rate_pct": round((sh / tot) * 100.0, 2) if tot > 0 else 0.0,
            "noshow_rate_pct": round((ns / tot) * 100.0, 2) if tot > 0 else 0.0,
            "financial_loss_inr": round(st["loss_inr"], 2)
        })

    return results
