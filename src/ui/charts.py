"""
Plotly Chart Factory & Healthcare Data Visualization Module
Generates responsive, publication-quality Plotly charts using standard corporate theme.
"""

from typing import List, Dict, Any
import plotly.graph_objects as go

# Enterprise Theme Palette
COLOR_PRIMARY = "#1F497D"      # Apex Navy
COLOR_SECONDARY = "#2E75B6"    # Accent Blue
COLOR_DANGER = "#EF4444"       # Rose / No-Show
COLOR_SUCCESS = "#10B981"      # Emerald / Show
COLOR_WARNING = "#F59E0B"      # Amber
COLOR_BG_CARD = "#FFFFFF"

LAYOUT_DEFAULTS = dict(
    template="plotly_white",
    font=dict(family="Inter, -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, sans-serif", size=12, color="#1E293B"),
    margin=dict(l=40, r=30, t=50, b=40),
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    hovermode="closest",
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
)

def plot_rolling_trend(trend_records: List[Dict[str, Any]]) -> go.Figure:
    """Renders temporal time series with daily volume bars and 7-day rolling moving average."""
    dates = [r["date"] for r in trend_records]
    total_vol = [r["total"] for r in trend_records]
    noshow_rates = [r["noshow_rate_pct"] for r in trend_records]
    moving_avg = [r["moving_avg_pct"] for r in trend_records]

    fig = go.Figure()

    # Daily volume bars
    fig.add_trace(go.Bar(
        x=dates,
        y=total_vol,
        name="Daily Appointments",
        marker_color="rgba(203, 213, 225, 0.6)",
        yaxis="y1"
    ))

    # Daily no-show rate scatter
    fig.add_trace(go.Scatter(
        x=dates,
        y=noshow_rates,
        mode="markers",
        name="Daily No-Show %",
        marker=dict(color=COLOR_DANGER, size=5, opacity=0.6),
        yaxis="y2"
    ))

    # 7-day moving average trendline
    fig.add_trace(go.Scatter(
        x=dates,
        y=moving_avg,
        mode="lines",
        name="Moving Average %",
        line=dict(color=COLOR_PRIMARY, width=3),
        yaxis="y2"
    ))

    layout = dict(LAYOUT_DEFAULTS)
    layout.update(
        title=dict(text="<b>30-Day Outpatient Attendance & Rolling Trend</b>", font=dict(size=14)),
        yaxis=dict(title="Appointment Volume", showgrid=True, gridcolor="#F1F5F9", side="left"),
        yaxis2=dict(title="No-Show Rate (%)", overlaying="y", side="right", range=[0, 60], showgrid=False),
        xaxis=dict(showgrid=False)
    )
    fig.update_layout(**layout)
    return fig

def plot_branch_comparison(branches: List[Dict[str, Any]]) -> go.Figure:
    """Renders stacked horizontal bar chart comparing Shows vs No-Shows across clinic branches."""
    names = [b["clinic_name"] for b in branches]
    shows = [b["shows"] for b in branches]
    noshows = [b["no_shows"] for b in branches]
    rates = [b["noshow_rate_pct"] for b in branches]

    fig = go.Figure()

    fig.add_trace(go.Bar(
        y=names,
        x=shows,
        name="Attended (Show)",
        orientation="h",
        marker_color=COLOR_SUCCESS,
        text=shows,
        textposition="inside"
    ))

    fig.add_trace(go.Bar(
        y=names,
        x=noshows,
        name="No-Show",
        orientation="h",
        marker_color=COLOR_DANGER,
        text=[f"{ns} ({r}%)" for ns, r in zip(noshows, rates)],
        textposition="inside"
    ))

    layout = dict(LAYOUT_DEFAULTS)
    layout.update(
        title=dict(text="<b>Clinical Branch Attendance Comparison</b>", font=dict(size=14)),
        barmode="stack",
        xaxis=dict(title="Total Scheduled Appointments", showgrid=True, gridcolor="#F1F5F9"),
        yaxis=dict(autorange="reversed")
    )
    fig.update_layout(**layout)
    return fig

def plot_lead_time_decay(decay_data: List[Dict[str, Any]]) -> go.Figure:
    """Renders dual-axis bar and line chart illustrating booking decay."""
    buckets = [d["lead_time_bucket"] for d in decay_data]
    totals = [d["total_appointments"] for d in decay_data]
    noshow_rates = [d["noshow_rate_pct"] for d in decay_data]

    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=buckets,
        y=totals,
        name="Booking Volume",
        marker_color=COLOR_SECONDARY,
        yaxis="y1"
    ))

    fig.add_trace(go.Scatter(
        x=buckets,
        y=noshow_rates,
        mode="lines+markers+text",
        name="No-Show Rate (%)",
        line=dict(color=COLOR_DANGER, width=3),
        marker=dict(size=8, color=COLOR_DANGER),
        text=[f"{r}%" for r in noshow_rates],
        textposition="top center",
        yaxis="y2"
    ))

    layout = dict(LAYOUT_DEFAULTS)
    layout.update(
        title=dict(text="<b>Booking Lead Time vs No-Show Decay Curve</b>", font=dict(size=14)),
        yaxis=dict(title="Scheduled Volume", showgrid=True, gridcolor="#F1F5F9"),
        yaxis2=dict(title="No-Show Propensity (%)", overlaying="y", side="right", range=[0, 60], showgrid=False),
        xaxis=dict(title="Booking Lead-Time Interval")
    )
    fig.update_layout(**layout)
    return fig

def plot_specialty_bars(specialties: List[Dict[str, Any]]) -> go.Figure:
    """Renders horizontal bar chart ranking specialties by no-show vulnerability."""
    specs = [s["specialty_name"] for s in reversed(specialties)]
    rates = [s["noshow_rate_pct"] for s in reversed(specialties)]
    colors = [COLOR_DANGER if r >= 20.0 else COLOR_PRIMARY for r in rates]

    fig = go.Figure(go.Bar(
        y=specs,
        x=rates,
        orientation="h",
        marker_color=colors,
        text=[f"{r}%" for r in rates],
        textposition="outside"
    ))

    layout = dict(LAYOUT_DEFAULTS)
    layout.update(
        title=dict(text="<b>Specialty No-Show Vulnerability Ranking</b>", font=dict(size=14)),
        xaxis=dict(title="No-Show Rate (%)", range=[0, max(rates) * 1.25 if rates else 30], showgrid=True, gridcolor="#F1F5F9"),
        yaxis=dict(title="")
    )
    fig.update_layout(**layout)
    return fig

def plot_arrival_heatmap(matrix_data: Dict[str, Any]) -> go.Figure:
    """Renders bivariate heatmap mapping weekdays against consultation time slots."""
    days = matrix_data["days"]
    slots = matrix_data["slots"]
    rate_grid = matrix_data["rate_grid"]
    vol_grid = matrix_data["volume_grid"]

    # Build annotations
    annotations = []
    for i, slot in enumerate(slots):
        for j, day in enumerate(days):
            rate = rate_grid[i][j]
            vol = vol_grid[i][j]
            annotations.append(dict(
                x=day,
                y=slot,
                text=f"<b>{rate}%</b><br><span style='font-size:10px; color:#64748B;'>({vol} appts)</span>",
                showarrow=False,
                font=dict(color="#1E293B" if rate < 25 else "#991B1B")
            ))

    fig = go.Figure(data=go.Heatmap(
        z=rate_grid,
        x=days,
        y=slots,
        colorscale=[
            [0.0, "#F0FDF4"],    # Low no-show (greenish)
            [0.4, "#FEF3C7"],    # Medium (amber)
            [1.0, "#FEE2E2"]     # High no-show (red)
        ],
        colorbar=dict(title="No-Show %"),
        hoverongaps=False
    ))

    layout = dict(LAYOUT_DEFAULTS)
    layout.update(
        title=dict(text="<b>Bivariate Arrival Heatmap: Weekday × Consultation Interval</b>", font=dict(size=14)),
        annotations=annotations,
        xaxis=dict(title="Day of Week"),
        yaxis=dict(title="Time Slot", autorange="reversed")
    )
    fig.update_layout(**layout)
    return fig

def plot_pareto_curve(records: List[Dict[str, Any]]) -> go.Figure:
    """Renders cumulative financial loss Pareto curve across patient volume."""
    # Sum financial loss per patient
    from collections import defaultdict
    patient_losses = defaultdict(float)
    for r in records:
        if r.get("is_no_show") == 1 or r.get("status") == "No-Show":
            mrn = r.get("patient_mrn", "Unknown")
            fee = float(r.get("actual_consultation_fee") or 0.0)
            patient_losses[mrn] += fee

    sorted_losses = sorted(patient_losses.values(), reverse=True)
    total_loss = sum(sorted_losses) if sorted_losses else 1.0

    cumulative = []
    curr = 0.0
    for loss in sorted_losses:
        curr += loss
        cumulative.append((curr / total_loss) * 100.0)

    x_pct = [((i + 1) / len(sorted_losses)) * 100.0 for i in range(len(sorted_losses))] if sorted_losses else [0]

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=x_pct,
        y=cumulative,
        mode="lines",
        name="Cumulative Loss %",
        line=dict(color=COLOR_DANGER, width=3),
        fill="tozeroy",
        fillcolor="rgba(239, 68, 68, 0.1)"
    ))

    # Reference 80/20 line
    fig.add_shape(
        type="line", x0=20, y0=0, x1=20, y1=100,
        line=dict(color="#64748B", dash="dash", width=1.5)
    )
    fig.add_shape(
        type="line", x0=0, y0=80, x1=100, y1=80,
        line=dict(color="#64748B", dash="dash", width=1.5)
    )

    layout = dict(LAYOUT_DEFAULTS)
    layout.update(
        title=dict(text="<b>Repeat Offender Pareto Loss Concentration</b>", font=dict(size=14)),
        xaxis=dict(title="% of Non-Attending Patients", range=[0, 100], showgrid=True, gridcolor="#F1F5F9"),
        yaxis=dict(title="Cumulative % of Financial Loss", range=[0, 105], showgrid=True, gridcolor="#F1F5F9")
    )
    fig.update_layout(**layout)
    return fig

def plot_risk_donut(records: List[Dict[str, Any]]) -> go.Figure:
    """Renders donut chart illustrating proportion of appointments in High, Medium, Low risk tiers."""
    tier_counts = {"High": 0, "Medium": 0, "Low": 0}
    for r in records:
        tier = r.get("risk_tier", "Low")
        if tier in tier_counts:
            tier_counts[tier] += 1
        else:
            tier_counts["Low"] += 1

    labels = ["High Risk (>=60)", "Medium Risk (30-59)", "Low Risk (<30)"]
    values = [tier_counts["High"], tier_counts["Medium"], tier_counts["Low"]]
    colors = [COLOR_DANGER, COLOR_WARNING, COLOR_SUCCESS]

    fig = go.Figure(data=go.Pie(
        labels=labels,
        values=values,
        hole=0.6,
        marker=dict(colors=colors),
        textinfo="label+percent"
    ))

    layout = dict(LAYOUT_DEFAULTS)
    layout.update(
        title=dict(text="<b>Predictive Risk Tier Distribution</b>", font=dict(size=14)),
        showlegend=False
    )
    fig.update_layout(**layout)
    return fig

def plot_monthly_leakage(monthly_data: List[Dict[str, Any]]) -> go.Figure:
    """Renders monthly revenue loss waterfall in INR."""
    months = [m["month"] for m in monthly_data]
    losses = [m["financial_loss_inr"] for m in monthly_data]

    fig = go.Figure(go.Bar(
        x=months,
        y=losses,
        marker_color=COLOR_DANGER,
        text=[f"₹{int(l):,}" for l in losses],
        textposition="outside"
    ))

    layout = dict(LAYOUT_DEFAULTS)
    layout.update(
        title=dict(text="<b>Monthly Outpatient Revenue Leakage (INR)</b>", font=dict(size=14)),
        xaxis=dict(title="Calendar Month"),
        yaxis=dict(title="Uncollected Consultation Fees (₹)", showgrid=True, gridcolor="#F1F5F9")
    )
    fig.update_layout(**layout)
    return fig
