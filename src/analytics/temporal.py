"""
Temporal Analytics & Bivariate Arrival Heatmap Engine
Computes 2D scheduling matrices across weekdays and consultation intervals.
"""

from typing import List, Dict, Any

WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
TIME_SLOTS = ["Morning", "Afternoon", "Evening"]

def compute_arrival_matrix(records: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Constructs a 2D bivariate matrix mapping weekdays to appointment time slots.
    Returns:
      - days: List[str]
      - slots: List[str]
      - volume_matrix: List[List[int]] (slots x days)
      - noshow_rate_matrix: List[List[float]] (slots x days)
      - cell_records: Dict[(day, slot), stats]
    """
    matrix = {(day, slot): {"total": 0, "noshows": 0} for day in WEEKDAYS for slot in TIME_SLOTS}

    for r in records:
        day = r.get("appointment_day_of_week")
        slot = r.get("appointment_time_slot")
        if day in WEEKDAYS and slot in TIME_SLOTS:
            matrix[(day, slot)]["total"] += 1
            if r.get("is_no_show") == 1 or r.get("status") == "No-Show":
                matrix[(day, slot)]["noshows"] += 1

    volume_grid = []
    rate_grid = []

    for slot in TIME_SLOTS:
        vol_row = []
        rate_row = []
        for day in WEEKDAYS:
            st = matrix[(day, slot)]
            tot = st["total"]
            ns = st["noshows"]
            rate = round((ns / tot) * 100.0, 1) if tot > 0 else 0.0
            vol_row.append(tot)
            rate_row.append(rate)
        volume_grid.append(vol_row)
        rate_grid.append(rate_row)

    return {
        "days": WEEKDAYS,
        "slots": TIME_SLOTS,
        "volume_grid": volume_grid,
        "rate_grid": rate_grid
    }
