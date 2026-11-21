import os
import openpyxl
from openpyxl.styles import Font, Alignment
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def style_table(table, col_widths, header_bg="1F497D", stripe_bg="F2F5F9"):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, cell in enumerate(table.rows[0].cells):
        set_cell_background(cell, header_bg)
        set_cell_margins(cell, top=110, bottom=110, left=130, right=130)
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for r in p.runs:
                r.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)
                r.font.name = 'Calibri'
                r.font.size = Pt(9.5)
    for r_idx, row in enumerate(table.rows[1:], start=1):
        bg = stripe_bg if r_idx % 2 == 1 else "FFFFFF"
        for cell in row.cells:
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=75, bottom=75, left=120, right=120)
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.name = 'Calibri'
                    r.font.size = Pt(9.0)
                    r.font.color.rgb = RGBColor(51, 51, 51)
    for row in table.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = Inches(width)

def build_step12_doc(output_path):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
    # Title Block
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(2)
    t_run = title_p.add_run("ABES ENGINEERING COLLEGE, GHAZIABAD")
    t_run.bold = True
    t_run.font.name = "Calibri"
    t_run.font.size = Pt(16)
    t_run.font.color.rgb = RGBColor(31, 73, 125)
    
    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_before = Pt(0)
    sub_p.paragraph_format.space_after = Pt(8)
    sub_run = sub_p.add_run("DEPARTMENT OF MASTER OF COMPUTER APPLICATIONS (MCA)\nMCA 2nd Year — Data Analytics Specialization | UI/UX Framework & Design System")
    sub_run.font.name = "Calibri"
    sub_run.font.size = Pt(11)
    sub_run.font.color.rgb = RGBColor(89, 89, 89)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # Document Header Box
    h_table = doc.add_table(rows=6, cols=2)
    h_widths = [2.2, 4.3]
    h_data = [
        ("Project ID & Batch:", "P101 | ABES_OND26 MCA 2ND YEAR_DATA ANALYTICS_BATCH 01"),
        ("Project Title:", "Appointment No-Show Analysis & Optimization System"),
        ("Student Details:", "Deepshikha | Roll No: 2500320140047 | MCA 2nd Year (DA)"),
        ("Faculty Mentor:", "Dinesh Sereno, Assistant Professor, Dept. of MCA"),
        ("UI Framework & Tech:", "Streamlit Multi-Page, Plotly 5.24, Tailwind CSS (via CDN in st.components.v1.html)"),
        ("Deliverable Step:", "Step 12 — UI Framework, Wireframe Blueprints & Design Specification")
    ]
    for r_idx, (label, val) in enumerate(h_data):
        row = h_table.rows[r_idx]
        cell_lbl, cell_val = row.cells[0], row.cells[1]
        
        p0 = cell_lbl.paragraphs[0]
        r0 = p0.add_run(label)
        r0.bold = True
        r0.font.name = "Calibri"
        r0.font.size = Pt(10)
        r0.font.color.rgb = RGBColor(31, 73, 125)
        
        p1 = cell_val.paragraphs[0]
        r1 = p1.add_run(val)
        r1.font.name = "Calibri"
        r1.font.size = Pt(10)
        r1.font.color.rgb = RGBColor(38, 38, 38)
        
        set_cell_background(cell_lbl, "EBF1F5")
        set_cell_background(cell_val, "FAFAFA")
        set_cell_margins(cell_lbl, 60, 60, 100, 100)
        set_cell_margins(cell_val, 60, 60, 100, 100)
        cell_lbl.width = Inches(h_widths[0])
        cell_val.width = Inches(h_widths[1])

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    def add_heading_1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(31, 73, 125)
        return p

    def add_heading_2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(46, 117, 182)
        return p

    def add_body_p(text, bold_prefix=None):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            br = p.add_run(bold_prefix)
            br.bold = True
            br.font.name = "Calibri"
            br.font.size = Pt(10)
            br.font.color.rgb = RGBColor(38, 38, 38)
        r = p.add_run(text)
        r.font.name = "Calibri"
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(51, 51, 51)
        return p

    def add_bullet_p(bold_prefix, text):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            br = p.add_run(bold_prefix)
            br.bold = True
            br.font.name = "Calibri"
            br.font.size = Pt(10)
        r = p.add_run(text)
        r.font.name = "Calibri"
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(51, 51, 51)
        return p

    def add_code_block(code_text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.left_indent = Inches(0.2)
        r = p.add_run(code_text)
        r.font.name = "Consolas"
        r.font.size = Pt(8.0)
        r.font.color.rgb = RGBColor(38, 38, 38)
        return p

    # 1. Screen Architecture
    add_heading_1("1. Application Screen Architecture & Navigation Hierarchy")
    add_body_p(
        "The user interface is structured across eight dedicated screens (UI-01 to UI-08) within a Streamlit multi-page layout. "
        "Each screen provides a specialized workspace tailored to specific operational user personas and analytical objectives:"
    )

    scr_table = doc.add_table(rows=9, cols=6)
    scr_widths = [0.7, 1.3, 1.8, 1.4, 0.9, 0.7]
    scr_headers = ["UI ID", "Screen Name", "Operational Purpose", "Visual & Layout Components", "Sidebar Filters", "User Stories"]
    for i, h in enumerate(scr_headers):
        scr_table.rows[0].cells[i].paragraphs[0].text = h

    screens_data = [
        ("UI-01", "Overview & Executive KPIs", "Network-wide attendance health summary and branch performance benchmarking.", "4 Tailwind KPI cards, 30-day rolling moving average line chart, Branch show vs no-show stacked bars.", "Date Range, Clinic Branch", "US-01, US-03, US-04"),
        ("UI-02", "Lead Time & Specialty", "Investigate lead-time decay curves and clinical specialty attendance variances.", "Lead-time bucket bar chart with decay trendline, Specialty attendance diverging bar chart, Fee distribution.", "Clinic Branch, Specialty, Lead Time", "US-05, US-06"),
        ("UI-03", "Demographics & Arrival Heatmap", "Examine demographic disparities and 2D arrival patterns to prevent lobby friction.", "Age group attendance grouped bar chart, Distance zone distribution, Day-of-week vs Time-slot arrival heatmap.", "Age Group, Gender, Time Slot", "US-07, US-08"),
        ("UI-04", "Statistical Inference", "Formal parametric and non-parametric hypothesis testing with p-values and odds ratios.", "Two-sample proportion Z-test card (SMS lift), Chi-square independence table, Comorbidity relative risk ratio.", "SMS Status, Chronic Flag", "US-09, US-10"),
        ("UI-05", "Financial & Capacity Loss", "Audit exact doctor idle consultation time and lost patient fee revenue.", "Monthly revenue leakage bar chart, Physician idle-hour breakdown table, Capacity waste metrics.", "Month, Department, Doctor", "US-12"),
        ("UI-06", "Predictive Risk Triage", "Proactive patient triage and high-risk calling queue for front-desk reconfirmation.", "Risk tier donut chart, Interactive calling priority table with phone checklist, Repeat offender Pareto curve.", "Risk Tier (H/M/L), Date, Clinic", "US-11, US-13"),
        ("UI-07", "Simulation & Corporate", "Model corporate shift wellness adherence and simulate calibrated overbooking scenarios.", "Corporate shift attendance chart, Overbooking multiplier interactive slider, Wait-time trade-off curve.", "Shift Type, Overbooking Slider", "US-14"),
        ("UI-08", "Data Quality & Export", "Transparent data hygiene audit, anomaly inspection, and clean CSV/Excel download.", "Missing value audit table, Outlier inspector, Paginated data explorer table, CSV & Excel export buttons.", "Global Search, Status, Branch", "US-01, US-02")
    ]
    for r_i, r_data in enumerate(screens_data, start=1):
        for c_i, val in enumerate(r_data):
            scr_table.rows[r_i].cells[c_i].paragraphs[0].text = val
    style_table(scr_table, scr_widths)

    # 2. ASCII Wireframes
    add_heading_1("2. ASCII Screen Wireframes")
    add_body_p("The wireframe blueprints below illustrate the visual layout, card positioning, widget placement, and chart structures:")

    add_heading_2("Wireframe 1: UI-01 — Overview & Executive KPIs")
    wireframe_ui01 = (
        "+---------------------------------------------------------------------------------------------------+\n"
        "|  APEXCARE HEALTH NETWORK — EXECUTIVE OVERVIEW DASHBOARD                             [Live SQLite] |\n"
        "+---------------------------------------------------------------------------------------------------+\n"
        "| SIDEBAR FILTERS       | [ KPI CARD 1 ]      [ KPI CARD 2 ]      [ KPI CARD 3 ]      [ KPI CARD 4 ]|\n"
        "| Date: [Jan - Jun 2026]| Total Bookings     Attendance Rate     No-Show Rate        Revenue Loss   |\n"
        "| Clinic: [All Branches]| 300 Appointments    81.67% (245)        18.33% (55)         INR 41,800    |\n"
        "| Specialty: [All]      | (+2.4% vs last mo)  (Optimal Threshold) (Target: < 16.0%)   (Wasted Slots)|\n"
        "| Status: [All]         +---------------------------------------------------------------------------+\n"
        "|                       | CHART 1: 30-Day Rolling Attendance & No-Show Trend (Plotly Scatter Line)  |\n"
        "| [x] Auto-Refresh      |  100% |-----------/\\----------/\\--------/\\----------------- (Show Rate)  |\n"
        "| [Reset All Filters]   |   20% |-----------\\-----------\\--------\\------------------ (No-Show Rate) |\n"
        "|                       |       +-------------------------------------------------------------------|\n"
        "|                       | CHART 2: Branch Attendance Proportions     | CHART 3: Lead-Time Overview  |\n"
        "|                       |  Sec 62  [==========Show========][==NS==]  |  0-2 days:  [=======] 17.2%  |\n"
        "|                       |  Sec 18  [===========Show=======][=NS=]   |  3-7 days:  [======] 14.3%   |\n"
        "|                       |  Expwy   [=========Show=========][===NS=]  |  15-30 days:[===========]31% |\n"
        "|                       |  GrNoida [==========Show========][==NS==]  |  >30 days:  [=============]42%|\n"
        "+---------------------------------------------------------------------------------------------------+"
    )
    add_code_block(wireframe_ui01)

    add_heading_2("Wireframe 2: UI-04 — Statistical Inference & Hypothesis Testing")
    wireframe_ui04 = (
        "+---------------------------------------------------------------------------------------------------+\n"
        "|  APEXCARE ANALYTICS — STATISTICAL INFERENCE & HYPOTHESIS TESTING                   [p < 0.05 Rigor]|\n"
        "+---------------------------------------------------------------------------------------------------+\n"
        "| SIDEBAR PARAMETERS    | HYPOTHESIS TEST 1: Digital Reminder (SMS) Protective Lift (Two-Sample Z)  |\n"
        "| Confidence: [95%]     | +------------------------------------+----------------------------------+ |\n"
        "| Alpha Level: [0.05]   | | No Reminder Sent (N=109): 26.6% NS | Reminder Sent (N=191): 13.6% NS | |\n"
        "| Tail: [Two-Sided]     | +------------------------------------+----------------------------------+ |\n"
        "|                       |   Z-Statistic: +2.82 | p-Value: 0.0048 (< 0.05) | STATISTICALLY SIGNIFICANT |\n"
        "| [x] Show 95% CI       |   Odds Ratio: 0.44 (95% CI: 0.25 - 0.77) -> 56% Protective Risk Reduction  |\n"
        "|                       +---------------------------------------------------------------------------+\n"
        "|                       | HYPOTHESIS TEST 2: Specialty Attrition Independence (Pearson Chi-Square)  |\n"
        "|                       |   Chi-Square (chi2): 8.42 | Degrees of Freedom (df): 5 | p-Value: 0.038    |\n"
        "|                       |   Conclusion: Reject H0; patient attendance is non-independent of specialty|\n"
        "|                       | +-----------------------------------------------------------------------+ |\n"
        "|                       | | Specialty: Dermatology   | N=52 | No-Show: 25.0% [Above Avg Warning]  | |\n"
        "|                       | | Specialty: Mental Well   | N=50 | No-Show: 22.0% [Above Avg Warning]  | |\n"
        "|                       | | Specialty: Cardiology    | N=40 | No-Show: 12.5% [Optimal Adherence]  | |\n"
        "+---------------------------------------------------------------------------------------------------+"
    )
    add_code_block(wireframe_ui04)

    add_heading_2("Wireframe 3: UI-06 — Predictive Risk Triage & Action Roster")
    wireframe_ui06 = (
        "+---------------------------------------------------------------------------------------------------+\n"
        "|  APEXCARE FRONT DESK — PRE-VISIT PATIENT RISK TRIAGE & CALLING ROSTER               [Next 48 Hours]|\n"
        "+---------------------------------------------------------------------------------------------------+\n"
        "| TRIAGE FILTERS        | [ HIGH RISK: 38 (12.7%) ]   [ MEDIUM RISK: 86 (28.7%) ]   [ LOW RISK: 176]|\n"
        "| Risk Tier: [x] High   | (Urgent Tele-Calling Queue) (Automated WhatsApp Push)     (Standard Flow) |\n"
        "| Lead Time: [> 7 Days] +---------------------------------------------------------------------------+\n"
        "| Clinic: [Sector 62]   | PRIORITIZED TELEPHONIC CALLING QUEUE (Sorted by Descending Risk Score)    |\n"
        "| Calling Status: [All] | [Search: Patient / Phone...] [Filter High Risk] [Export Today's Call List]|\n"
        "|                       | +-------+------------+-----------------+----------+---------+------+-----+ |\n"
        "|                       | | Appt# | Patient    | Specialty / Doc | LeadTime | Remind? | Risk |Actn | |\n"
        "|                       | +-------+------------+-----------------+----------+---------+------+-----+ |\n"
        "|                       | | 10015 | Aarav S.   | Mental Wellness | 28 days  | No (0)  | 82%  |[Call| |\n"
        "|                       | | 10042 | Rohan M.   | Dermatology     | 21 days  | No (0)  | 74%  |[Call| |\n"
        "|                       | | 10088 | Pooja V.   | Orthopedics     | 18 days  | Yes(1)  | 63%  |[Call| |\n"
        "|                       | +-------+------------+-----------------+----------+---------+------+-----+ |\n"
        "+---------------------------------------------------------------------------------------------------+"
    )
    add_code_block(wireframe_ui06)

    # 3. Global Layout Specifications
    add_heading_1("3. Global Layout Specifications & Healthcare Design System")
    add_bullet_p("Shared Sidebar Session State: ", "Global filtering widgets are bound to `st.session_state` keys (`selected_date_range`, `selected_clinic`, `selected_specialty`, `selected_risk_tier`). Slicing state persists seamlessly across multi-page navigation.")
    add_bullet_p("Corporate Healthcare Color Palette: ", "Brand Deep Navy (#1F497D), Medical Cyan (#2E75B6), Slate Background Canvas (#F8FAFC), Card Background (#FFFFFF), Alert Rose/Red (#EF4444), Success Mint/Green (#10B981).")
    add_bullet_p("Typography Standard: ", "Inter / Segoe UI sans-serif typography with strict hierarchical sizing: Page Title (24px Bold), Section Headings (18px Semi-Bold), KPI Values (32px Bold), Body & Data Labels (13-14px Regular).")
    add_bullet_p("Responsive & Mobile Adaptation: ", "Streamlit layout configured to `layout='wide'`. Tailwind grid uses responsive classes (`grid-cols-1 sm:grid-cols-2 lg:grid-cols-4`), collapsing gracefully to single-column card feeds on mobile devices.")

    # 4. Reusable Component Library
    add_heading_1("4. Reusable Tailwind CSS Component Library (src/ui.py)")
    add_body_p(
        "To elevate visual quality beyond standard Streamlit widgets without causing iframe isolation bugs, "
        "`src/ui.py` wraps semantic HTML with Tailwind CDN inside `st.components.v1.html(..., height=...)`:"
    )
    add_bullet_p("render_tailwind_header(title, subtitle, badge): ", "Renders an executive gradient slate/navy hero banner with clinic metadata and database connection status.")
    add_bullet_p("render_kpi_card_row(kpis): ", "Renders a responsive 4-column grid of bordered metric cards featuring trend badges, large numerical values, and category styling.")
    add_bullet_p("render_insight_card(title, text, type): ", "Renders actionable operational guidance callout boxes (info, warning, success) highlighting key takeaways.")

    # 5. Best Chart for each OQ
    add_heading_1("5. Best Visual Chart for Each Operational Question (OQ-01 to OQ-10)")
    add_body_p("Visualizations are selected based on data types, cognitive load, and analytical intent:")

    oq_chart_data = [
        ("OQ-01", "Overall Baseline Attendance & Trends", "Dual-Layer Plotly Chart: Grouped Bar Chart + 30-Day Rolling Line", "Bar chart shows monthly absolute volumes; line chart smooths daily volatility to reveal underlying trajectory."),
        ("OQ-02", "Booking Lead-Time Impact", "Binned Lead-Time Bar Chart with Logistic Decay Fit Curve", "Clearly shows non-linear drop-off as lead time expands across discrete operational buckets (0-2d to >30d)."),
        ("OQ-03", "Medical Specialty Attendance Variance", "Horizontal Diverging Stacked Bar Chart with Department Benchmark Line", "Horizontal layout provides ample label space for medical specialties and ranks departments from highest to lowest risk."),
        ("OQ-04", "Demographic & Proximity Cohorts", "Clustered Bar Chart & Transit Distance Boxplot", "Contrasts attendance across 5 age cohorts while boxplot illustrates transit distance spread and transit friction."),
        ("OQ-05", "Digital Reminder (SMS) Lift", "Grouped Percentage Bar Chart with 95% Confidence Interval Error Bars", "Directly visualizes the 13.0% percentage point attendance lift and validates statistical significance."),
        ("OQ-06", "Day-of-Week & Time-Slot Patterns", "2D Bivariate Density Heatmap (Mon-Sat vs Morning/Afternoon/Evening)", "Identifies high-attrition operational cells (e.g., Saturday afternoons) instantly through color intensity."),
        ("OQ-07", "Chronic Comorbidity Compliance", "Grouped Dual Bar Chart with Relative Risk Ratio Badges", "Contrasts diabetic/hypertensive patients against healthy controls, highlighting superior routine compliance."),
        ("OQ-08", "Repeat Offender Behavioral Loss", "Pareto Chart (Dual-Axis Bar Chart + Cumulative Lorenz Share Curve)", "Empirically tests the 80/20 rule: demonstrates what % of missed slots are driven by habitual offenders."),
        ("OQ-09", "Financial & Doctor-Hour Loss", "Waterfall Chart & Specialty Stacked Bar Chart", "Deconstructs total lost consultation fees and idle physician hours by clinical department."),
        ("OQ-10", "Predictive Risk Scoring & Triage", "Interactive Triage Data Table with Color Badges + Donut Gauge", "Donut chart shows risk tier proportions; interactive table equips front-desk staff with actionable call queues.")
    ]

    ch_table = doc.add_table(rows=11, cols=4)
    ch_widths = [0.8, 2.0, 2.0, 2.0]
    ch_headers = ["OQ ID", "Operational Question", "Optimal Visual Chart Type", "Analytical Justification"]
    for i, h in enumerate(ch_headers):
        ch_table.rows[0].cells[i].paragraphs[0].text = h
    for r_i, r_data in enumerate(oq_chart_data, start=1):
        for c_i, val in enumerate(r_data):
            ch_table.rows[r_i].cells[c_i].paragraphs[0].text = val
    style_table(ch_table, ch_widths)

    # 6. Streamlit Config
    add_heading_1("6. Streamlit Server Configuration (.streamlit/config.toml)")
    add_body_p("Theme configuration file created in repository:")
    add_code_block(
        "[theme]\n"
        "primaryColor = \"#1F497D\"\n"
        "backgroundColor = \"#F8FAFC\"\n"
        "secondaryBackgroundColor = \"#FFFFFF\"\n"
        "textColor = \"#1E293B\"\n"
        "font = \"sans serif\"\n\n"
        "[server]\n"
        "headless = true\n"
        "enableCORS = false\n"
        "enableXsrfProtection = true"
    )

    doc.add_paragraph().paragraph_format.space_after = Pt(16)
    doc.save(output_path)
    print(f"Step 12 UI Framework Document successfully created at: {output_path}")

def update_rtm_with_ui_screens(rtm_path):
    wb = openpyxl.load_workbook(rtm_path)
    ws = wb["Traceability_Matrix"]
    
    ui_mappings = {
        "BR-01": "UI-01, UI-08",
        "BR-02": "UI-02",
        "BR-03": "UI-02",
        "BR-04": "UI-03",
        "BR-05": "UI-04",
        "BR-06": "UI-03",
        "BR-07": "UI-04",
        "BR-08": "UI-06",
        "BR-09": "UI-05",
        "BR-10": "UI-06",
        "BR-11": "UI-07",
        "BR-12": "UI-07"
    }
    
    font_code_rtm = Font(name="Calibri", size=9.0, bold=True, color="1F497D")
    
    # Column 3 is Req ID (BR), Column 11 is UI Screen (Col K)
    for row in range(5, ws.max_row + 1):
        br_id = ws.cell(row=row, column=3).value
        if br_id in ui_mappings:
            cell_ui = ws.cell(row=row, column=11)
            cell_ui.value = ui_mappings[br_id]
            cell_ui.font = font_code_rtm
            cell_ui.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
    wb.save(rtm_path)
    print(f"RTM Excel successfully updated with UI Screen IDs at: {rtm_path}")

if __name__ == "__main__":
    doc_out = r"c:\Users\deeps\Downloads\Appointment No Show Analysis\docs\12_UI_Framework.docx"
    rtm_out = r"c:\Users\deeps\Downloads\Appointment No Show Analysis\docs\04_Traceability_Matrix.xlsx"
    build_step12_doc(doc_out)
    update_rtm_with_ui_screens(rtm_out)
