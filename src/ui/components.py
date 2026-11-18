"""
Tailwind CSS-Powered Components for Streamlit Multi-Page Application
"""

from typing import List, Dict, Any, Optional
import streamlit as st
import streamlit.components.v1 as components

def render_tailwind_header(title: str, subtitle: str, badge: str = "Live Production"):
    """
    Renders an enterprise healthcare header banner styled with Tailwind CSS via CDN.
    """
    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <script src="https://cdn.tailwindcss.com"></script>
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
        <style>
            body {{ font-family: 'Inter', sans-serif; }}
        </style>
    </head>
    <body class="bg-transparent m-0 p-1">
        <div class="bg-gradient-to-r from-slate-900 via-blue-950 to-indigo-950 rounded-xl p-5 shadow-sm border border-slate-800 text-white flex flex-col md:flex-row justify-between items-start md:items-center gap-3">
            <div>
                <div class="flex items-center gap-2 mb-1">
                    <span class="inline-flex items-center px-2 py-0.5 rounded text-xs font-semibold bg-blue-500/20 text-blue-300 border border-blue-400/30">
                        {badge}
                    </span>
                    <span class="text-xs text-slate-400 font-medium">ApexCare Health Network • Noida Operations</span>
                </div>
                <h1 class="text-xl md:text-2xl font-bold tracking-tight text-white mb-0.5">{title}</h1>
                <p class="text-xs md:text-sm text-slate-300 font-normal m-0">{subtitle}</p>
            </div>
            <div class="flex items-center gap-3 self-end md:self-center">
                <div class="text-right">
                    <p class="text-xs text-slate-400 m-0">Analytics Engine</p>
                    <p class="text-xs font-semibold text-emerald-400 m-0">● Connected (SQLite)</p>
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    components.html(html_code, height=115, scrolling=False)


def render_kpi_card_row(kpis: list):
    """
    Renders a responsive 4-column row of Tailwind CSS metric KPI cards.
    kpis is a list of dicts: [
        {"title": str, "value": str, "subtitle": str, "trend": str, "color": "blue"|"green"|"amber"|"red"}
    ]
    """
    cards_html = ""
    color_schemes = {
        "blue": {"border": "border-blue-200", "bg": "bg-blue-50/50", "text": "text-blue-900", "badge": "bg-blue-100 text-blue-700"},
        "green": {"border": "border-emerald-200", "bg": "bg-emerald-50/50", "text": "text-emerald-950", "badge": "bg-emerald-100 text-emerald-700"},
        "amber": {"border": "border-amber-200", "bg": "bg-amber-50/50", "text": "text-amber-950", "badge": "bg-amber-100 text-amber-700"},
        "red": {"border": "border-rose-200", "bg": "bg-rose-50/50", "text": "text-rose-950", "badge": "bg-rose-100 text-rose-700"}
    }

    for kpi in kpis:
        c = color_schemes.get(kpi.get("color", "blue"), color_schemes["blue"])
        trend_badge = f'<span class="text-xs font-medium px-2 py-0.5 rounded-full {c["badge"]}">{kpi.get("trend", "")}</span>' if kpi.get("trend") else ""
        
        cards_html += f"""
        <div class="bg-white rounded-xl border {c['border']} p-4 shadow-sm hover:shadow transition-shadow duration-200 flex flex-col justify-between">
            <div class="flex justify-between items-start mb-2">
                <span class="text-xs font-semibold uppercase tracking-wider text-slate-500">{kpi['title']}</span>
                {trend_badge}
            </div>
            <div class="mb-1">
                <span class="text-2xl md:text-3xl font-bold {c['text']} tracking-tight">{kpi['value']}</span>
            </div>
            <div class="text-xs text-slate-500 font-medium">
                {kpi.get('subtitle', '')}
            </div>
        </div>
        """

    full_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <script src="https://cdn.tailwindcss.com"></script>
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
        <style>
            body {{ font-family: 'Inter', sans-serif; }}
        </style>
    </head>
    <body class="bg-transparent m-0 p-1">
        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-{len(kpis)} gap-3">
            {cards_html}
        </div>
    </body>
    </html>
    """
    height = 125 if len(kpis) <= 4 else 145
    components.html(full_html, height=height, scrolling=False)


def render_insight_card(title: str, text: str, insight_type: str = "info"):
    """
    Renders an executive insight/recommendation callout card.
    insight_type: "info" | "warning" | "success"
    """
    types = {
        "info": {"bg": "bg-blue-50", "border": "border-blue-200", "title": "text-blue-900", "text": "text-blue-800", "icon": "ℹ️"},
        "warning": {"bg": "bg-amber-50", "border": "border-amber-200", "title": "text-amber-900", "text": "text-amber-800", "icon": "⚠️"},
        "success": {"bg": "bg-emerald-50", "border": "border-emerald-200", "title": "text-emerald-900", "text": "text-emerald-800", "icon": "✅"}
    }
    t = types.get(insight_type, types["info"])

    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <script src="https://cdn.tailwindcss.com"></script>
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
        <style>body {{ font-family: 'Inter', sans-serif; }}</style>
    </head>
    <body class="bg-transparent m-0 p-1">
        <div class="rounded-lg {t['bg']} border {t['border']} p-3.5 flex items-start gap-3 shadow-xs">
            <span class="text-lg leading-none">{t['icon']}</span>
            <div>
                <h4 class="text-xs font-bold uppercase tracking-wider {t['title']} m-0 mb-0.5">{title}</h4>
                <p class="text-xs leading-relaxed {t['text']} m-0">{text}</p>
            </div>
        </div>
    </body>
    </html>
    """
    components.html(html_code, height=75, scrolling=False)


def records_to_markdown_table(records: List[Dict[str, Any]], columns: Optional[List[str]] = None, limit: int = 50) -> str:
    """
    Converts a list of dicts to a GitHub-flavored Markdown table for clean Streamlit rendering.
    """
    if not records:
        return "_No records available to display._"
    cols = columns or list(records[0].keys())
    headers = [str(c).replace("_", " ").title() for c in cols]
    header_line = "| " + " | ".join(headers) + " |"
    sep_line = "| " + " | ".join("---" for _ in cols) + " |"
    data_lines = []
    for r in records[:limit]:
        row_cells = [str(r.get(c, "")).replace("|", "/") for c in cols]
        data_lines.append("| " + " | ".join(row_cells) + " |")
    return "\n".join([header_line, sep_line] + data_lines)
