"""
Demographic & Comorbidity Analytics Module
Analyzes attendance adherence across age cohorts, genders, transit distances, and comorbidities.
"""

from typing import List, Dict, Any
from collections import defaultdict

AGE_COHORTS = ["18-35", "36-50", "51-65", "65+"]
DISTANCE_BANDS = ["<5 km", "5-10 km", "10-20 km", ">20 km"]

def bin_age_group(age: int) -> str:
    """Classifies patient age into standard epidemiological cohorts."""
    if age <= 35:
        return "18-35"
    elif age <= 50:
        return "36-50"
    elif age <= 65:
        return "51-65"
    else:
        return "65+"

def bin_distance_band(km: float) -> str:
    """Classifies patient transit distance into commuting zones."""
    if km < 5.0:
        return "<5 km"
    elif km <= 10.0:
        return "5-10 km"
    elif km <= 20.0:
        return "10-20 km"
    else:
        return ">20 km"

def compute_age_group_breakdown(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Evaluates attendance metrics across patient age cohorts."""
    groups = {cohort: {"total": 0, "shows": 0, "noshows": 0} for cohort in AGE_COHORTS}

    for r in records:
        age = int(r.get("age") or 35)
        cohort = r.get("age_group") or bin_age_group(age)
        if cohort not in groups:
            cohort = bin_age_group(age)
            
        groups[cohort]["total"] += 1
        is_ns = (r.get("is_no_show") == 1 or r.get("status") == "No-Show")
        if is_ns:
            groups[cohort]["noshows"] += 1
        else:
            groups[cohort]["shows"] += 1

    results = []
    for c in AGE_COHORTS:
        tot = groups[c]["total"]
        sh = groups[c]["shows"]
        ns = groups[c]["noshows"]
        results.append({
            "age_group": c,
            "total_appointments": tot,
            "shows": sh,
            "no_shows": ns,
            "attendance_rate_pct": round((sh / tot) * 100.0, 2) if tot > 0 else 0.0,
            "noshow_rate_pct": round((ns / tot) * 100.0, 2) if tot > 0 else 0.0
        })
    return results

def compute_gender_breakdown(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Compares show vs no-show counts by gender."""
    groups = defaultdict(lambda: {"total": 0, "shows": 0, "noshows": 0})
    for r in records:
        g = r.get("gender", "Other")
        groups[g]["total"] += 1
        if r.get("is_no_show") == 1 or r.get("status") == "No-Show":
            groups[g]["noshows"] += 1
        else:
            groups[g]["shows"] += 1

    results = []
    for g, st in groups.items():
        tot = st["total"]
        sh = st["shows"]
        ns = st["noshows"]
        results.append({
            "gender": g,
            "total": tot,
            "shows": sh,
            "no_shows": ns,
            "noshow_rate_pct": round((ns / tot) * 100.0, 2) if tot > 0 else 0.0
        })
    return results

def compute_distance_breakdown(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Analyzes transit friction and distance to clinic."""
    groups = {band: {"total": 0, "shows": 0, "noshows": 0} for band in DISTANCE_BANDS}

    for r in records:
        dist = float(r.get("transit_distance_km") or 5.0)
        band = bin_distance_band(dist)
        groups[band]["total"] += 1
        if r.get("is_no_show") == 1 or r.get("status") == "No-Show":
            groups[band]["noshows"] += 1
        else:
            groups[band]["shows"] += 1

    results = []
    for band in DISTANCE_BANDS:
        tot = groups[band]["total"]
        sh = groups[band]["shows"]
        ns = groups[band]["noshows"]
        results.append({
            "distance_band": band,
            "total": tot,
            "shows": sh,
            "no_shows": ns,
            "noshow_rate_pct": round((ns / tot) * 100.0, 2) if tot > 0 else 0.0
        })
    return results

def compute_comorbidity_breakdown(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Cross-tabulates chronic clinical conditions (Hypertension, Diabetes)."""
    categories = {
        "Hypertension & Diabetes": {"total": 0, "noshows": 0},
        "Hypertension Only": {"total": 0, "noshows": 0},
        "Diabetes Only": {"total": 0, "noshows": 0},
        "Neither (Healthy Baseline)": {"total": 0, "noshows": 0}
    }

    for r in records:
        ht = int(r.get("hypertension_flag") or 0)
        db = int(r.get("diabetes_flag") or 0)
        is_ns = (r.get("is_no_show") == 1 or r.get("status") == "No-Show")

        if ht == 1 and db == 1:
            cat = "Hypertension & Diabetes"
        elif ht == 1 and db == 0:
            cat = "Hypertension Only"
        elif ht == 0 and db == 1:
            cat = "Diabetes Only"
        else:
            cat = "Neither (Healthy Baseline)"

        categories[cat]["total"] += 1
        if is_ns:
            categories[cat]["noshows"] += 1

    results = []
    for cat, st in categories.items():
        tot = st["total"]
        ns = st["noshows"]
        sh = tot - ns
        results.append({
            "category": cat,
            "total": tot,
            "shows": sh,
            "no_shows": ns,
            "noshow_rate_pct": round((ns / tot) * 100.0, 2) if tot > 0 else 0.0
        })
    return results
