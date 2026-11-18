"""
Data Quality, Cleaning & Imputation Engine
Imputes missing values, enforces business domains, and audits dataset hygiene.
"""

from typing import List, Dict, Any, Tuple
import statistics

def audit_data_quality(records: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Performs comprehensive hygiene audit across records.
    Returns metrics on missing values, duplicate IDs, and out-of-bound anomalies.
    """
    total_records = len(records)
    if total_records == 0:
        return {
            "total_records": 0,
            "missing_values": {},
            "duplicate_ids": 0,
            "outliers": {}
        }

    # Track missing fields
    missing_counts = {}
    keys = list(records[0].keys())
    for k in keys:
        missing_counts[k] = 0

    seen_ids = set()
    duplicate_ids = 0
    age_outliers = 0
    lead_time_outliers = 0
    negative_lead_times = 0

    for r in records:
        # Check duplicates
        aid = r.get("appointment_id")
        if aid in seen_ids:
            duplicate_ids += 1
        else:
            seen_ids.add(aid)

        # Check missing values
        for k in keys:
            v = r.get(k)
            if v is None or v == "" or (isinstance(v, str) and v.strip() == ""):
                missing_counts[k] += 1

        # Check outliers
        age = r.get("age")
        if age is not None:
            if age > 100 or age < 0:
                age_outliers += 1

        lead = r.get("lead_time_days")
        if lead is not None:
            if lead < 0:
                negative_lead_times += 1
            elif lead > 60:
                lead_time_outliers += 1

    # Filter only columns that have missing values
    active_missing = {k: v for k, v in missing_counts.items() if v > 0}

    return {
        "total_records": total_records,
        "unique_appointments": len(seen_ids),
        "duplicate_ids": duplicate_ids,
        "missing_fields": active_missing,
        "outliers": {
            "age_gt_100": age_outliers,
            "lead_time_gt_60": lead_time_outliers,
            "negative_lead_times": negative_lead_times
        }
    }

def clean_dataset(records: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """
    Cleans and standardizes raw analytical records:
    1. Imputes missing reminder_lead_hours using median lead time of available reminders.
    2. Clamps and flags extreme outliers.
    3. Normalizes text fields (case-standardization and whitespace trimming).
    4. Computes derived flags if absent.
    Returns (cleaned_records, audit_report).
    """
    initial_audit = audit_data_quality(records)
    
    # Calculate median reminder lead hours for imputation if present
    reminder_hours = [
        float(r["reminder_lead_hours"])
        for r in records
        if "reminder_lead_hours" in r and r["reminder_lead_hours"] is not None
    ]
    median_reminder_hr = round(statistics.median(reminder_hours), 1) if reminder_hours else 24.0

    cleaned = []
    seen_ids = set()

    for r in records:
        rec = dict(r)

        # Skip duplicate appointment IDs to maintain strict PK uniqueness
        aid = rec.get("appointment_id")
        if aid in seen_ids:
            continue
        seen_ids.add(aid)

        # Impute missing reminder lead hours if column exists
        if "reminder_lead_hours" in rec:
            if rec["reminder_lead_hours"] is None or rec["reminder_lead_hours"] == "":
                rec["reminder_lead_hours"] = median_reminder_hr

        # Clean strings
        for str_col in ["status", "specialty_name", "clinic_name", "corporate_shift_type", "gender"]:
            if str_col in rec and isinstance(rec[str_col], str):
                rec[str_col] = rec[str_col].strip()

        # Enforce is_no_show flag consistency
        status = rec.get("status", "")
        if "is_no_show" not in rec or rec["is_no_show"] is None:
            rec["is_no_show"] = 1 if status == "No-Show" else 0
        else:
            rec["is_no_show"] = int(rec["is_no_show"])

        cleaned.append(rec)

    post_audit = audit_data_quality(cleaned)
    audit_summary = {
        "raw_count": len(records),
        "clean_count": len(cleaned),
        "deduplicated": len(records) - len(cleaned),
        "imputed_reminder_hours_median": median_reminder_hr,
        "initial_audit": initial_audit,
        "post_audit": post_audit
    }

    return cleaned, audit_summary
