"""
Predictive Risk Scoring & Clinical Triage Engine
Assigns calibrated multi-factor risk scores and prioritizes front-desk call rosters.
"""

from typing import List, Dict, Any, Tuple

def assign_risk_tier(score: float) -> str:
    """
    Classifies a continuous risk score (0-100) into actionable operational tiers.
    Strictly satisfies TC-T-11:
      >= 60.0 -> 'High'
      30.0 to 59.9 -> 'Medium'
      < 30.0 -> 'Low'
    """
    if score >= 60.0:
        return "High"
    elif score >= 30.0:
        return "Medium"
    else:
        return "Low"

def calculate_risk_score(record: Dict[str, Any]) -> float:
    """
    Computes a heuristic risk score (0.0 to 100.0) based on multi-factor clinical variables:
    1. Booking lead time (>14 days: +25, 8-14 days: +15, 3-7 days: +5)
    2. Digital reminder status (No reminder: +20, Reminder sent: +0)
    3. Specialty vulnerability (Mental Wellness, Dermatology: +15, General: +5)
    4. Historical attendance tier (Habitual: +35, Occasional: +15, First-Time/Reliable: +0)
    5. Commute friction (Transit distance > 15 km: +10)
    """
    score = 10.0 # Base minimum propensity

    # 1. Lead time factor
    lead = float(record.get("lead_time_days") or 0.0)
    if lead > 14:
        score += 25.0
    elif lead >= 8:
        score += 15.0
    elif lead >= 3:
        score += 5.0

    # 2. Reminder factor
    reminder = int(record.get("sms_reminder_sent") or 0)
    if reminder == 0:
        score += 20.0

    # 3. Specialty factor
    spec = record.get("specialty_name", "")
    if spec in ["Mental Wellness", "Dermatology"]:
        score += 15.0
    elif spec in ["General Medicine"]:
        score += 5.0

    # 4. Repeat history factor
    tier = record.get("repeat_offender_tier", "")
    if tier == "Habitual":
        score += 35.0
    elif tier == "Occasional":
        score += 15.0

    # 5. Distance factor
    dist = float(record.get("transit_distance_km") or 0.0)
    if dist > 15.0:
        score += 10.0

    # Cap score between 0.0 and 100.0
    return min(100.0, max(0.0, round(score, 1)))

def enrich_records_with_risk(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Enriches all records with calculated risk scores and tiers if not already present."""
    enriched = []
    for r in records:
        rec = dict(r)
        if "risk_score" not in rec or rec["risk_score"] is None:
            score = calculate_risk_score(rec)
            rec["risk_score"] = score
        else:
            score = float(rec["risk_score"])

        rec["risk_tier"] = assign_risk_tier(score)
        enriched.append(rec)
    return enriched

def generate_triage_queue(records: List[Dict[str, Any]], min_tier: str = "High", limit: int = 50) -> List[Dict[str, Any]]:
    """
    Generates a prioritized calling roster for clinic receptionists.
    Sorts by risk score descending and filters by target risk tier.
    """
    enriched = enrich_records_with_risk(records)
    
    tier_weights = {"High": 3, "Medium": 2, "Low": 1}
    min_weight = tier_weights.get(min_tier, 1)

    qualified = [
        r for r in enriched
        if tier_weights.get(r.get("risk_tier", "Low"), 1) >= min_weight
    ]

    # Sort primarily by risk score desc, then by consultation fee desc
    qualified.sort(
        key=lambda x: (float(x.get("risk_score") or 0.0), float(x.get("actual_consultation_fee") or 0.0)),
        reverse=True
    )

    return qualified[:limit]
