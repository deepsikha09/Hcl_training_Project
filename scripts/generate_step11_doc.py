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

def build_step11_doc(output_path):
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
    sub_run = sub_p.add_run("DEPARTMENT OF MASTER OF COMPUTER APPLICATIONS (MCA)\nMCA 2nd Year — Data Analytics Specialization | Agile User Story Specification")
    sub_run.font.name = "Calibri"
    sub_run.font.size = Pt(11)
    sub_run.font.color.rgb = RGBColor(89, 89, 89)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # Document Header Box
    h_table = doc.add_table(rows=5, cols=2)
    h_widths = [2.2, 4.3]
    h_data = [
        ("Project Title:", "Appointment No-Show Analysis & Optimization System"),
        ("Student Details:", "Deepshikha | Roll No: 2500320140047 | MCA 2nd Year (DA)"),
        ("Faculty Mentor:", "Alpna Lodhi mam, Assistant Professor, Dept. of MCA"),
        ("Agile Methodology:", "Scrum Framework | 5 Epics | 14 User Stories (44 Story Points) | 10-Day Sprint"),
        ("Deliverable Step:", "Step 11 — Agile User Stories & Acceptance Criteria Specification")
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

    # 1. Master User Stories Table
    add_heading_1("1. Master Agile User Stories Specification Table")
    add_body_p(
        "Fourteen agile user stories (US-01 through US-14) translate business requirements (BR) and technical functional features (FR) "
        "into discrete, testable development backlog increments. Sized using standard Fibonacci story points (1 to 8):"
    )

    ust_table = doc.add_table(rows=15, cols=7)
    ust_widths = [0.7, 1.2, 2.5, 0.7, 0.6, 0.5, 0.5]
    ust_headers = ["US ID", "Target Persona", "User Story: \"As a <persona>, I want <goal> so that <benefit>\"", "Feature", "BR ID", "Priority", "Pts"]
    for i, h in enumerate(ust_headers):
        ust_table.rows[0].cells[i].paragraphs[0].text = h

    stories_summary = [
        ("US-01", "Data Analyst (SH-05)", "As a Data Analyst, I want to load cached appointment data from SQLite with automated schema validation so that the application operates on clean, verified records without runtime crashes.", "FR-01", "BR-01", "Must", "3"),
        ("US-02", "Data Analyst (SH-05)", "As a Data Analyst, I want the system to automatically handle missing values, duplicates, and outliers so that downstream analytics reflect accurate, sanitized data.", "FR-01", "BR-01", "Must", "3"),
        ("US-03", "Ops Manager (SH-01)", "As a Clinic Operations Manager, I want to view top-level KPI cards (Total Bookings, Attendance Rate, No-Show Rate, Financial Loss) with modern Tailwind styling so that I can instantly evaluate high-level clinic performance.", "FR-01", "BR-01", "Must", "2"),
        ("US-04", "Ops Manager (SH-01)", "As a Clinic Operations Manager, I want to filter attendance trends across different clinic branches and date intervals so that I can identify locations and periods with rising non-attendance.", "FR-01", "BR-01", "Must", "3"),
        ("US-05", "Front Desk (SH-03)", "As a Front Desk Supervisor, I want to analyze patient attendance rates grouped by booking lead-time buckets so that I know at what booking intervals patients become high-risk for forgetfulness.", "FR-02", "BR-02", "Must", "3"),
        ("US-06", "CMO (SH-02)", "As a Chief Medical Officer, I want to examine attendance rates across all 6 clinical departments so that I can pinpoint specialties with severe patient drop-off and adjust staffing rosters.", "FR-03", "BR-03", "Must", "3"),
        ("US-07", "Ops Manager (SH-01)", "As a Clinic Operations Manager, I want to segment attendance by age groups, gender, and transit distance bands so that we can understand physical and transit barriers to clinical care.", "FR-04", "BR-04", "Should", "2"),
        ("US-08", "Front Desk (SH-03)", "As a Front Desk Supervisor, I want to inspect a 2D arrival heatmap of weekdays versus consultation time slots so that I can anticipate waiting lobby congestion and avoid overbooking peak times.", "FR-06", "BR-06", "Should", "3"),
        ("US-09", "Front Desk (SH-03)", "As a Front Desk Supervisor, I want to run a formal two-sample proportion Z-test and odds ratio on reminder dispatches so that our team can empirically prove whether SMS reminders produce statistically significant attendance lift.", "FR-05", "BR-05", "Must", "5"),
        ("US-10", "CMO (SH-02)", "As a Chief Medical Officer, I want to monitor follow-up attendance among diabetic and hypertensive patients so that clinical teams can ensure treatment continuity and prevent disease complications.", "FR-07", "BR-07", "Must", "3"),
        ("US-11", "Front Desk (SH-03)", "As a Front Desk Staff, I want to identify habitual repeat no-show patients and quantify their share of total missed slots so that our calling outreach focuses on high-frequency offenders rather than first-timers.", "FR-08", "BR-08", "Should", "3"),
        ("US-12", "Ops Manager (SH-01)", "As a Clinic Operations Manager, I want an automated accounting report of monthly consultation fee leakage and physician idle hours so that hospital leadership can quantify operational waste in rupees and hours.", "FR-09", "BR-09", "Must", "3"),
        ("US-13", "Front Desk (SH-03)", "As a Front Desk Coordinator, I want upcoming appointments classified into High, Medium, and Low risk tiers with an interactive calling list so that staff can prioritize high-risk patients for daily phone reconfirmation.", "FR-10", "BR-10", "Must", "5"),
        ("US-14", "Corporate HR (SH-04)", "As a Corporate Partner HR Lead & Ops Lead, I want to correlate corporate shift timings with checkup completion and simulate calibrated overbooking scenarios so that clinic capacity can be recovered safely.", "FR-11, FR-12", "BR-11, BR-12", "Could", "5")
    ]
    for r_i, r_data in enumerate(stories_summary, start=1):
        for c_i, val in enumerate(r_data):
            ust_table.rows[r_i].cells[c_i].paragraphs[0].text = val
    style_table(ust_table, ust_widths)

    # 2. Epic Groupings
    add_heading_1("2. Epic Architecture & Backlog Distribution")
    epics = [
        ("Epic 1: Data Foundation & Ingestion Engine", "US-01, US-02", "6 Points", "Handles SQLite persistence, cached data loading (@st.cache_data), and the automated cleaning pipeline."),
        ("Epic 2: Core Overview & Facility Monitoring", "US-03, US-04", "5 Points", "Delivers executive Tailwind KPI summary cards, branch filters, and temporal moving average visual trends."),
        ("Epic 3: Exploratory & Behavioral Analysis", "US-05, US-06, US-07, US-08", "11 Points", "Implements multi-dimensional cohort slicing: lead-time decay, medical specialties, demographics, and 2D arrival heatmaps."),
        ("Epic 4: Statistical Inference & Financial Accounting", "US-09, US-10, US-11, US-12", "14 Points", "Executes formal hypothesis testing (Z-test, p-value), comorbidity risk ratios, repeat offender Pareto curves, and doctor revenue loss formulas."),
        ("Epic 5: Predictive Risk Triage & Capacity Simulation", "US-13, US-14", "10 Points", "Calculates the composite no-show risk score, delivers the daily front-desk calling priority list, and models overbooking capacity.")
    ]
    e_table = doc.add_table(rows=6, cols=4)
    e_widths = [2.2, 1.2, 1.0, 2.3]
    e_headers = ["Epic Name", "Encompassed Stories", "Story Points", "Operational Scope"]
    for i, h in enumerate(e_headers):
        e_table.rows[0].cells[i].paragraphs[0].text = h
    for r_i, r_data in enumerate(epics, start=1):
        for c_i, val in enumerate(r_data):
            e_table.rows[r_i].cells[c_i].paragraphs[0].text = val
    style_table(e_table, e_widths)

    # 3. Acceptance Criteria in Given/When/Then Format
    add_heading_1("3. Detailed Acceptance Criteria (Given / When / Then)")
    add_body_p("Each user story is governed by strict, binary pass/fail acceptance criteria written in Gherkin behavioral syntax:")

    gherkin_criteria = [
        ("US-01", "Data Ingestion & Integrity Loader",
         "Given the SQLite database file exists at data/app.db,\n"
         "When the application boots or data/loader.py executes,\n"
         "Then exactly 300 fact records and 36 view columns are loaded into memory with zero schema validation errors."),

        ("US-02", "Data Quality & Cleaning Pipeline",
         "Given raw appointment records containing missing reminder hours and outlier entries,\n"
         "When src/data/cleaner.py processes the dataset,\n"
         "Then missing reminder hours are imputed with valid sentinel values, outliers are flagged, and clean DataFrames are returned."),

        ("US-03", "Executive KPI Summary Cards",
         "Given the Overview page is loaded with an active dataset slice,\n"
         "When the dashboard renders,\n"
         "Then 4 Tailwind-styled KPI cards display Total Appointments, Attendance Rate %, No-Show Rate %, and Total Revenue Loss (₹) with exact decimal accuracy."),

        ("US-04", "Multi-Branch Attendance & Trend Tracking",
         "Given the user is on the Overview dashboard,\n"
         "When the user selects 'ApexCare Sector 62 Clinic' in the sidebar clinic filter,\n"
         "Then all rendered charts and metrics update dynamically to display data exclusively for Sector 62 within 500 milliseconds."),

        ("US-05", "Lead-Time Attendance Decay Analysis",
         "Given the Lead Time Analysis tab is active,\n"
         "When the user reviews the lead-time distribution chart,\n"
         "Then no-show rates are displayed across 5 standard intervals (0-2d, 3-7d, 8-14d, 15-30d, >30d), showing higher no-show percentages in longer lead intervals."),

        ("US-06", "Specialty-Level Attendance Breakdown",
         "Given the Specialty Analysis tab is active,\n"
         "When consultation show rates are grouped by medical department,\n"
         "Then a comparative horizontal bar chart renders show vs. no-show proportions across all 6 clinical specialties."),

        ("US-07", "Demographic & Proximity Cohort Profiling",
         "Given the Demographics Analysis view is opened,\n"
         "When the user selects age group and transit distance breakdown,\n"
         "Then attendance percentages are segmented across 5 standard age brackets and 3 transit distance zones."),

        ("US-08", "Day & Time Slot Heatmap Visualization",
         "Given the Temporal Analysis tab is selected,\n"
         "When the 2D arrival heatmap renders,\n"
         "Then a color-coded matrix displays attendance rates cross-tabulated across Monday–Saturday and Morning, Afternoon, Evening slots."),

        ("US-09", "Digital Reminder Efficacy Hypothesis Testing",
         "Given the Statistical Inference dashboard is open,\n"
         "When the user reviews the Digital Reminder evaluation section,\n"
         "Then the two-sample proportion Z-test statistic, relative risk reduction %, and p-value (p < 0.05) are explicitly displayed with statistical interpretation banners."),

        ("US-10", "Chronic Disease Comorbidity Compliance Monitoring",
         "Given the Clinical Compliance tab is loaded,\n"
         "When filtering by patient chronic condition flags,\n"
         "Then compliance rates for diabetic and hypertensive patients are contrasted against non-chronic baselines with relative risk ratios."),

        ("US-11", "Repeat Offender Pareto Behavioral Analysis",
         "Given the Behavioral Analysis tab is active,\n"
         "When the repeat offender distribution is displayed,\n"
         "Then patients are categorized into 'First-Timer', 'Occasional Miss', and 'Habitual', showing the cumulative share of missed visits caused by habitual offenders."),

        ("US-12", "Financial Revenue Leakage & Capacity Loss Accounting",
         "Given the Financial & Capacity Loss dashboard tab is open,\n"
         "When the user selects a specific month or department,\n"
         "Then the exact consultation fee leakage in INR and wasted specialist consultation hours are computed and displayed in summary cards."),

        ("US-13", "Predictive Risk Scoring & Advance Triage Roster",
         "Given upcoming scheduled appointments exist in the system,\n"
         "When the Front Desk coordinator opens the Risk Triage screen,\n"
         "Then appointments are color-coded into High (>=60), Medium (30-59), and Low (<30) risk bands with an exportable telephonic calling priority list."),

        ("US-14", "Corporate Shift & Overbooking Simulation Engine",
         "Given the Advanced Simulation & Corporate page is open,\n"
         "When the user adjusts the overbooking buffer multiplier slider (1.05 to 1.25),\n"
         "Then projected recovered consultation slots and simulated waiting lobby queue times are calculated and displayed in real time.")
    ]

    for us_id, us_name, us_ghk in gherkin_criteria:
        add_heading_2(f"{us_id}: {us_name}")
        add_body_p(us_ghk)

    # 4. Sprint Schedule
    add_heading_1("4. Agile Sprint Implementation Roadmap (10-Day Plan)")
    add_body_p("The 44 story points are allocated across a 2-week, 10-day agile sprint cycle matching the upcoming Build steps:")

    sp_data = [
        ("Day 1 - Day 2", "Epic 1: Data Engineering & Foundation", "US-01, US-02", "6 Points", "Build SQLite connection, @st.cache_data loader, schema assertions, and missing value cleaning pipeline."),
        ("Day 3 - Day 4", "Epic 2: Core Dashboard & Overview", "US-03, US-04", "5 Points", "Develop root app.py, Overview page, Tailwind KPI metric cards, and sidebar clinic/date filters."),
        ("Day 5 - Day 6", "Epic 3: Exploratory & Behavioral Analysis", "US-05, US-06, US-07, US-08", "11 Points", "Build lead-time decay charts, specialty bar plots, demographic violin distributions, and 2D arrival heatmaps."),
        ("Day 7 - Day 8", "Epic 4: Statistical Inference & Finance", "US-09, US-10, US-11, US-12", "14 Points", "Implement SciPy Chi-square and statsmodels Z-test engines, odds ratio displays, and doctor financial loss accounting."),
        ("Day 9 - Day 10", "Epic 5: Predictive Risk Triage & QA", "US-13, US-14", "10 Points", "Implement composite risk scoring algorithm, front-desk calling priority table, overbooking simulator, and pytest suite.")
    ]
    sp_table = doc.add_table(rows=6, cols=5)
    sp_widths = [1.2, 1.8, 1.1, 0.9, 2.0]
    sp_headers = ["Sprint Days", "Target Epic", "User Stories", "Points", "Sprint Goals & Milestones"]
    for i, h in enumerate(sp_headers):
        sp_table.rows[0].cells[i].paragraphs[0].text = h
    for r_i, r_data in enumerate(sp_data, start=1):
        for c_i, val in enumerate(r_data):
            sp_table.rows[r_i].cells[c_i].paragraphs[0].text = val
    style_table(sp_table, sp_widths)

    doc.add_paragraph().paragraph_format.space_after = Pt(16)
    doc.save(output_path)
    print(f"Step 11 User Stories Document successfully created at: {output_path}")

def update_rtm_with_user_stories(rtm_path):
    wb = openpyxl.load_workbook(rtm_path)
    ws = wb["Traceability_Matrix"]
    
    # Mapping Req ID (BR) to User Story IDs
    us_mappings = {
        "BR-01": "US-01, US-03, US-04",
        "BR-02": "US-05",
        "BR-03": "US-06",
        "BR-04": "US-07",
        "BR-05": "US-09",
        "BR-06": "US-08",
        "BR-07": "US-10",
        "BR-08": "US-11",
        "BR-09": "US-12",
        "BR-10": "US-13",
        "BR-11": "US-14",
        "BR-12": "US-14"
    }
    
    font_code_rtm = Font(name="Calibri", size=9.0, bold=True, color="1F497D")
    
    # Column 3 is Req ID (BR), Column 10 is User Story (Col J)
    for row in range(5, ws.max_row + 1):
        br_id = ws.cell(row=row, column=3).value
        if br_id in us_mappings:
            cell_us = ws.cell(row=row, column=10)
            cell_us.value = us_mappings[br_id]
            cell_us.font = font_code_rtm
            cell_us.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
    wb.save(rtm_path)
    print(f"RTM Excel successfully updated with User Story IDs at: {rtm_path}")

if __name__ == "__main__":
    doc_out = r"c:\Users\deeps\Downloads\Appointment No Show Analysis\docs\11_User_Stories.docx"
    rtm_out = r"c:\Users\deeps\Downloads\Appointment No Show Analysis\docs\04_Traceability_Matrix.xlsx"
    build_step11_doc(doc_out)
    update_rtm_with_user_stories(rtm_out)
