"""
Statistical Inference & Hypothesis Testing Engine
Calculates Two-Sample Proportion Z-tests, Chi-square tests of independence,
and epidemiological odds ratios using pure Python mathematical routines.
"""

import math
from typing import List, Dict, Any, Tuple, Optional

def _std_norm_cdf(z: float) -> float:
    """Cumulative distribution function for the standard normal distribution."""
    return 0.5 * (1.0 + math.erf(z / math.sqrt(2.0)))

def _two_tailed_p_value_z(z: float) -> float:
    """Calculates two-tailed p-value from a Z-score using complementary error function."""
    return math.erfc(abs(z) / math.sqrt(2.0))

def _chi2_upper_incomplete_gamma(a: float, x: float, eps: float = 1e-11) -> float:
    """
    Computes regularized upper incomplete gamma function Q(a, x) = 1 - P(a, x).
    Used to calculate Chi-Square distribution survival p-values: Q(df/2, chi2/2).
    """
    if x <= 0:
        return 1.0
    if a <= 0:
        return 0.0

    # Incomplete gamma series for P(a, x)
    term = 1.0 / a
    total_sum = term
    for n in range(1, 250):
        term *= x / (a + n)
        total_sum += term
        if term < eps * total_sum:
            break

    try:
        log_p = -x + a * math.log(x) - math.lgamma(a)
        if log_p < -700:
            P = 0.0
        else:
            P = total_sum * math.exp(log_p)
    except (ValueError, OverflowError):
        P = 0.0

    return max(0.0, min(1.0, 1.0 - P))

def two_sample_proportion_z_test(
    records: List[Dict[str, Any]],
    treatment_field: str = "sms_reminder_sent",
    outcome_field: str = "is_no_show"
) -> Dict[str, Any]:
    """
    Two-sample proportion Z-test evaluating the efficacy of SMS reminders.
    Sample 1: Control (No Reminder, treatment = 0)
    Sample 2: Intervention (Reminder Sent, treatment = 1)
    Tests H0: p1 <= p2 vs H1: p1 > p2 (Reminder reduces no-shows).
    """
    n1, x1 = 0, 0 # No reminder
    n2, x2 = 0, 0 # Reminder sent

    for r in records:
        is_treated = int(r.get(treatment_field) or 0)
        is_ns = 1 if (r.get(outcome_field) == 1 or r.get("status") == "No-Show") else 0

        if is_treated == 0:
            n1 += 1
            if is_ns == 1:
                x1 += 1
        else:
            n2 += 1
            if is_ns == 1:
                x2 += 1

    if n1 == 0 or n2 == 0:
        return {
            "n_control": n1, "events_control": x1, "rate_control_pct": 0.0,
            "n_treated": n2, "events_treated": x2, "rate_treated_pct": 0.0,
            "z_stat": 0.0, "p_value": 1.0, "relative_risk_reduction_pct": 0.0,
            "odds_ratio": 1.0, "ci_lower": 1.0, "ci_upper": 1.0, "is_significant": False
        }

    p1 = x1 / n1
    p2 = x2 / n2

    # Pooled proportion
    p_pool = (x1 + x2) / (n1 + n2)
    se = math.sqrt(p_pool * (1.0 - p_pool) * (1.0 / n1 + 1.0 / n2)) if p_pool > 0 and p_pool < 1 else 0.0001
    z_stat = (p1 - p2) / se if se > 0 else 0.0

    # Two-tailed p-value
    p_val = _two_tailed_p_value_z(z_stat)

    # Relative Risk Reduction (RRR)
    rrr = ((p1 - p2) / p1 * 100.0) if p1 > 0 else 0.0

    # Odds Ratio (OR) with 95% Wald CI
    # Contingency:
    #                 No-Show   Show
    # No Reminder       x1     n1 - x1
    # Reminder          x2     n2 - x2
    a = max(1, x1)
    b = max(1, n1 - x1)
    c = max(1, x2)
    d = max(1, n2 - x2)

    odds_ratio = (a * d) / (b * c)
    se_log_or = math.sqrt(1.0 / a + 1.0 / b + 1.0 / c + 1.0 / d)
    ci_lower = math.exp(math.log(odds_ratio) - 1.96 * se_log_or)
    ci_upper = math.exp(math.log(odds_ratio) + 1.96 * se_log_or)

    return {
        "n_control": n1,
        "events_control": x1,
        "rate_control_pct": round(p1 * 100.0, 2),
        "n_treated": n2,
        "events_treated": x2,
        "rate_treated_pct": round(p2 * 100.0, 2),
        "z_stat": round(z_stat, 4),
        "p_value": round(p_val, 6),
        "relative_risk_reduction_pct": round(rrr, 2),
        "odds_ratio": round(odds_ratio, 3),
        "ci_lower": round(ci_lower, 3),
        "ci_upper": round(ci_upper, 3),
        "is_significant": bool(p_val < 0.05)
    }

def chi_square_specialty_test(records: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Chi-Square Test of Independence evaluating whether no-show propensity
    varies significantly across medical specialties (contingency table r x 2).
    """
    specialties = sorted(list({r.get("specialty_name", "General") for r in records}))
    table = {s: {"shows": 0, "noshows": 0} for s in specialties}

    for r in records:
        s = r.get("specialty_name", "General")
        is_ns = 1 if (r.get("is_no_show") == 1 or r.get("status") == "No-Show") else 0
        if is_ns == 1:
            table[s]["noshows"] += 1
        else:
            table[s]["shows"] += 1

    total_shows = sum(table[s]["shows"] for s in specialties)
    total_noshows = sum(table[s]["noshows"] for s in specialties)
    grand_total = total_shows + total_noshows

    if grand_total == 0 or len(specialties) < 2:
        return {
            "chi2_stat": 0.0, "p_value": 1.0, "degrees_of_freedom": 1,
            "is_significant": False, "table": table
        }

    p_show_overall = total_shows / grand_total
    p_noshow_overall = total_noshows / grand_total

    chi2_stat = 0.0
    for s in specialties:
        row_total = table[s]["shows"] + table[s]["noshows"]
        if row_total == 0:
            continue
        exp_show = row_total * p_show_overall
        exp_noshow = row_total * p_noshow_overall

        if exp_show > 0:
            chi2_stat += ((table[s]["shows"] - exp_show) ** 2) / exp_show
        if exp_noshow > 0:
            chi2_stat += ((table[s]["noshows"] - exp_noshow) ** 2) / exp_noshow

    df = len(specialties) - 1
    p_val = _chi2_upper_incomplete_gamma(df / 2.0, chi2_stat / 2.0)

    return {
        "chi2_stat": round(chi2_stat, 4),
        "degrees_of_freedom": df,
        "p_value": round(p_val, 6),
        "is_significant": bool(p_val < 0.05),
        "contingency_table": table
    }

def comorbidity_relative_risk(records: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Calculates relative risk of no-show for chronic comorbidity cohorts."""
    # Cohort with hypertension vs healthy
    ht_pos, ht_pos_ns = 0, 0
    ht_neg, ht_neg_ns = 0, 0

    for r in records:
        ht = int(r.get("hypertension_flag") or 0)
        is_ns = 1 if (r.get("is_no_show") == 1 or r.get("status") == "No-Show") else 0
        if ht == 1:
            ht_pos += 1
            if is_ns:
                ht_pos_ns += 1
        else:
            ht_neg += 1
            if is_ns:
                ht_neg_ns += 1

    rate_ht = (ht_pos_ns / ht_pos) if ht_pos > 0 else 0.0
    rate_no_ht = (ht_neg_ns / ht_neg) if ht_neg > 0 else 0.0
    rr_ht = (rate_ht / rate_no_ht) if rate_no_ht > 0 else 1.0

    return {
        "hypertension_patients": ht_pos,
        "hypertension_noshow_rate_pct": round(rate_ht * 100.0, 2),
        "baseline_patients": ht_neg,
        "baseline_noshow_rate_pct": round(rate_no_ht * 100.0, 2),
        "relative_risk": round(rr_ht, 2)
    }
