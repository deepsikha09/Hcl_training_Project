"""
UI and Charting Subpackage.
Exposes both Tailwind CSS components and Plotly chart generators.
"""

from .components import (
    render_tailwind_header,
    render_kpi_card_row,
    render_insight_card,
    records_to_markdown_table
)

from .charts import (
    plot_rolling_trend,
    plot_branch_comparison,
    plot_lead_time_decay,
    plot_specialty_bars,
    plot_arrival_heatmap,
    plot_pareto_curve,
    plot_risk_donut,
    plot_monthly_leakage
)

__all__ = [
    "render_tailwind_header",
    "render_kpi_card_row",
    "render_insight_card",
    "plot_rolling_trend",
    "plot_branch_comparison",
    "plot_lead_time_decay",
    "plot_specialty_bars",
    "plot_arrival_heatmap",
    "plot_pareto_curve",
    "plot_risk_donut",
    "plot_monthly_leakage"
]
