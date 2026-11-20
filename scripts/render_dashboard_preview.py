"""
Render high-fidelity dashboard screenshot mockup using Pillow for Slide 7.
"""

from PIL import Image, ImageDraw, ImageFont
import os

def render_dashboard():
    W, H = 1280, 720
    img = Image.new("RGB", (W, H), color="#F8FAFC")
    draw = ImageDraw.Draw(img)

    # Fonts (using default bitmap or standard TTF if available)
    try:
        font_large = ImageFont.truetype("arial.ttf", 22)
        font_mid = ImageFont.truetype("arial.ttf", 15)
        font_small = ImageFont.truetype("arial.ttf", 12)
        font_bold = ImageFont.truetype("arialbd.ttf", 15)
        font_title = ImageFont.truetype("arialbd.ttf", 20)
        font_kpi_val = ImageFont.truetype("arialbd.ttf", 28)
    except Exception:
        font_large = ImageFont.load_default()
        font_mid = ImageFont.load_default()
        font_small = ImageFont.load_default()
        font_bold = ImageFont.load_default()
        font_title = ImageFont.load_default()
        font_kpi_val = ImageFont.load_default()

    # Sidebar (Left)
    sidebar_w = 260
    draw.rectangle([(0, 0), (sidebar_w, H)], fill="#FFFFFF", outline="#E2E8F0")
    draw.rectangle([(15, 20), (sidebar_w - 15, 65)], fill="#1E293B", outline="#0F172A")
    draw.text((25, 32), "🏥 ApexCare Operations", fill="#FFFFFF", font=font_bold)
    draw.text((25, 80), "Noida Outpatient Network", fill="#64748B", font=font_small)

    # Sidebar Nav Items
    pages = [
        ("• Executive Overview", True),
        ("1. Overview & KPIs", False),
        ("2. Lead Time & Specialty", False),
        ("3. Demographics & Arrivals", False),
        ("4. Statistical Inference", False),
        ("5. Financial & Capacity", False),
        ("6. Predictive Risk Triage", False),
        ("7. Simulation & Corporate", False),
        ("8. Data Quality & Export", False)
    ]
    y_nav = 115
    for p_name, is_active in pages:
        if is_active:
            draw.rectangle([(15, y_nav - 4), (sidebar_w - 15, y_nav + 22)], fill="#EFF6FF", outline="#BFDBFE")
            draw.text((25, y_nav), p_name, fill="#1D4ED8", font=font_bold)
        else:
            draw.text((25, y_nav), p_name, fill="#475569", font=font_small)
        y_nav += 32

    # Main Area (Right)
    main_x = sidebar_w + 20
    main_w = W - main_x - 20

    # Header Banner
    banner_h = 75
    draw.rectangle([(main_x, 15), (main_x + main_w, 15 + banner_h)], fill="#1F497D", outline="#1E3A8A")
    draw.text((main_x + 20, 26), "Executive Outpatient Analytics & No-Show Diagnostic Platform", fill="#FFFFFF", font=font_title)
    draw.text((main_x + 20, 56), "Live Monitoring Across Noida Facilities (Sector 62, 18, 137, Greater Noida) • SQLite Connected", fill="#93C5FD", font=font_small)

    # 4 KPI Cards
    kpi_y = 105
    kpi_w = (main_w - 30) // 4
    kpis = [
        ("TOTAL APPOINTMENTS", "300", "245 Attended | 55 Missed", "#1F497D", "#DBEAFE"),
        ("ATTENDANCE RATE", "81.7%", "Target: 85.0% Adherence", "#15803D", "#DCFCE7"),
        ("NO-SHOW RATE", "18.3%", "Target: < 15.0% Non-Attendance", "#B91C1C", "#FEE2E2"),
        ("REVENUE LEAKAGE", "₹49,500", "23.8 Doctor Idle Hours", "#B45309", "#FEF3C7")
    ]
    for i, (title, val, sub, text_c, bg_c) in enumerate(kpis):
        x0 = main_x + i * (kpi_w + 10)
        draw.rectangle([(x0, kpi_y), (x0 + kpi_w, kpi_y + 85)], fill="#FFFFFF", outline="#E2E8F0")
        draw.rectangle([(x0, kpi_y), (x0 + kpi_w, kpi_y + 4)], fill=text_c)
        draw.text((x0 + 12, kpi_y + 12), title, fill="#64748B", font=font_small)
        draw.text((x0 + 12, kpi_y + 30), val, fill=text_c, font=font_kpi_val)
        draw.text((x0 + 12, kpi_y + 65), sub, fill="#64748B", font=font_small)

    # 2 Main Charts
    chart_y = 205
    chart_w = (main_w - 15) // 2
    chart_h = 360

    # Chart 1: 30-Day Rolling Trend
    c1_x = main_x
    draw.rectangle([(c1_x, chart_y), (c1_x + chart_w, chart_y + chart_h)], fill="#FFFFFF", outline="#E2E8F0")
    draw.text((c1_x + 15, chart_y + 15), "30-Day Outpatient Attendance & Rolling Trend", fill="#1E293B", font=font_bold)
    
    # Mock bars & line
    for b_idx in range(18):
        bx = c1_x + 40 + b_idx * 24
        bh = 60 + (b_idx * 7) % 110
        draw.rectangle([(bx, chart_y + chart_h - 40 - bh), (bx + 14, chart_y + chart_h - 40)], fill="#CBD5E1")
    # Trendline
    points = [(c1_x + 40 + b_idx * 24 + 7, chart_y + chart_h - 70 - ((b_idx * 5) % 60)) for b_idx in range(18)]
    for p_i in range(len(points) - 1):
        draw.line([points[p_i], points[p_i + 1]], fill="#1F497D", width=3)
    draw.text((c1_x + 40, chart_y + chart_h - 25), "Nov 01 ------------------------------------------- Nov 30", fill="#94A3B8", font=font_small)

    # Chart 2: Branch Comparison
    c2_x = main_x + chart_w + 15
    draw.rectangle([(c2_x, chart_y), (c2_x + chart_w, chart_y + chart_h)], fill="#FFFFFF", outline="#E2E8F0")
    draw.text((c2_x + 15, chart_y + 15), "Clinical Branch Attendance Comparison", fill="#1E293B", font=font_bold)

    branches = [
        ("Sector 62 (IT)", 85, 15, 230),
        ("Sector 18 (Hub)", 68, 14, 185),
        ("Sector 137 (Exp)", 52, 12, 145),
        ("Greater Noida", 40, 14, 120)
    ]
    by = chart_y + 70
    for name, shows, noshows, bar_len in branches:
        draw.text((c2_x + 20, by), name, fill="#1E293B", font=font_bold)
        show_len = int(bar_len * 0.82)
        ns_len = bar_len - show_len
        # Show bar (Green)
        draw.rectangle([(c2_x + 160, by), (c2_x + 160 + show_len, by + 26)], fill="#10B981")
        # No-Show bar (Red)
        draw.rectangle([(c2_x + 160 + show_len, by), (c2_x + 160 + bar_len, by + 26)], fill="#EF4444")
        draw.text((c2_x + 170 + bar_len, by + 5), f"{shows} Sh / {noshows} NS", fill="#64748B", font=font_small)
        by += 65

    # Bottom Insight Card
    ins_y = chart_y + chart_h + 15
    draw.rectangle([(main_x, ins_y), (main_x + main_w, ins_y + 60)], fill="#F0FDF4", outline="#BBF7D0")
    draw.text((main_x + 20, ins_y + 12), "💡 EXECUTIVE DIRECTIVE: SMS REMINDER INTERVENTION", fill="#166534", font=font_bold)
    draw.text((main_x + 20, ins_y + 34), "Automated SMS reminders reduce outpatient no-shows by 48.84% (p = 0.0052). Mandatory 48-hr call confirmations for bookings >14 days.", fill="#14532D", font=font_small)

    os.makedirs("docs", exist_ok=True)
    out_path = "docs/18_Dashboard_Preview.png"
    img.save(out_path)
    print(f"Dashboard mockup saved to: {out_path}")

if __name__ == "__main__":
    render_dashboard()
