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

def build_step13_doc(output_path):
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
    sub_run = sub_p.add_run("DEPARTMENT OF MASTER OF COMPUTER APPLICATIONS (MCA)\nMCA 2nd Year — Data Analytics Specialization | Quality Assurance & Test Specification")
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
        ("QA Framework:", "Pytest Automated Suite + Functional Manual Test Protocol (TC-F & TC-T)"),
        ("Deliverable Step:", "Step 13 — Comprehensive Test Plan, Test Cases & Final Traceability Matrix")
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

    # 1. Functional Test Cases
    add_heading_1("1. Functional Test Cases Specification (TC-F-01 through TC-F-16)")
    add_body_p(
        "Sixteen end-to-end functional test cases evaluate user interaction, multi-select filters, "
        "Tailwind KPI calculations, interactive Plotly renderings, predictive triage workflows, and data export:"
    )

    ft_headers = ["Test ID", "Linked Req / US", "UI Screen", "Pre-condition", "Execution Steps", "Test Data", "Expected Test Result", "Pri", "Status"]
    ft_widths = [0.6, 0.9, 0.6, 0.8, 1.2, 0.6, 1.2, 0.4, 0.4]
    ft_table = doc.add_table(rows=17, cols=9)
    for i, h in enumerate(ft_headers):
        ft_table.rows[0].cells[i].paragraphs[0].text = h

    f_tests = [
        ("TC-F-01", "BR-01\nUS-04", "UI-01", "App loaded with 300 records", "Select start date '2026-02-01' and end date '2026-04-30'", "Dates Feb-Apr", "All KPI cards and line charts slice to Feb-Apr dates within 500ms.", "Must", "Not run"),
        ("TC-F-02", "BR-01\nUS-04", "UI-01", "App loaded with default filters", "Select 'ApexCare Sector 62 Clinic' from sidebar dropdown", "Branch=Sec 62", "Dashboard recalculates metrics exclusively for Sector 62.", "Must", "Not run"),
        ("TC-F-03", "BR-01\nUS-03", "UI-01", "Database loaded", "Inspect 4 Tailwind KPI cards on Overview screen", "Entire dataset", "Total Bookings=300, Attendance=81.67%, No-Show=18.33%, Loss=₹41,800.", "Must", "Not run"),
        ("TC-F-04", "BR-01\nUS-04", "UI-01", "Overview page active", "Hover over 30-day rolling average Plotly line chart", "Cursor hover", "Interactive tooltip displays date, rolling show rate %, and sample N.", "Must", "Not run"),
        ("TC-F-05", "BR-02\nUS-05", "UI-02", "Lead-Time screen open", "Inspect lead-time bar chart binned into 5 standard buckets", "Lead time bins", "Bars display attendance rates: 0-2d (~17%), 3-7d (~14%), >30d (~42%).", "Must", "Not run"),
        ("TC-F-06", "BR-03\nUS-06", "UI-02", "Specialty tab active", "Inspect specialty horizontal diverging bar chart", "6 Specialties", "Departments sorted by no-show %; Dermatology and Mental Wellness ranked highest.", "Must", "Not run"),
        ("TC-F-07", "BR-04\nUS-07", "UI-03", "Demographics view open", "Select age group '18-30' and gender 'F' in sidebar", "Age=18-30, F", "Grouped bar chart updates showing attendance metrics for young adult females.", "Should", "Not run"),
        ("TC-F-08", "BR-06\nUS-08", "UI-03", "Temporal tab selected", "Inspect 2D weekday vs time-slot arrival heatmap", "18 matrix cells", "Heatmap renders with color intensity highlighting peak risk cells.", "Should", "Not run"),
        ("TC-F-09", "BR-05\nUS-09", "UI-04", "Stats Inference open", "Review digital reminder two-sample proportion Z-test card", "SMS sent (1/0)", "Z-stat > 2.5, p-value < 0.05, odds ratio with 95% Wald CI displayed.", "Must", "Not run"),
        ("TC-F-10", "BR-07\nUS-10", "UI-04", "Stats Inference open", "Inspect chronic condition compliance metrics", "HTN & Diabetes", "Compliance rate for chronic patients exceeds healthy baseline with RR badge.", "Must", "Not run"),
        ("TC-F-11", "BR-08\nUS-11", "UI-06", "Triage page active", "Inspect repeat offender Pareto distribution chart", "Prior misses", "Habitual repeat offender cohort (>=2 misses) shows ~37.5% no-show rate.", "Should", "Not run"),
        ("TC-F-12", "BR-09\nUS-12", "UI-05", "Financial page open", "Select 'Mental Wellness' department filter", "Spec=Mental Well", "Computes uncollected fees in ₹ and idle consultation doctor-hours.", "Must", "Not run"),
        ("TC-F-13", "BR-10\nUS-13", "UI-06", "Risk Triage active", "Filter priority queue by 'High Risk' tier (score >= 60.0)", "Tier='High'", "Table displays urgent tele-calling queue sorted by descending risk score.", "Must", "Not run"),
        ("TC-F-14", "BR-11\nBR-12", "UI-07", "Corporate tab open", "Move overbooking multiplier slider from 1.00 to 1.15", "Slider = 1.15", "Real-time curve updates showing 15% recovered slots vs lobby wait-time variance.", "Could", "Not run"),
        ("TC-F-15", "BR-01\nUS-04", "UI-01", "Overview page active", "Set impossible date filter (e.g. year 2030) or empty clinic", "Empty slice", "Renders graceful warning banner ('No records match') without crashing.", "Must", "Not run"),
        ("TC-F-16", "BR-01\nUS-01", "UI-08", "Data Explorer open", "Click 'Download Filtered Data as CSV' button", "Active slice", "Browser downloads validated CSV file containing exact filtered records.", "Must", "Not run")
    ]
    for r_i, r_data in enumerate(f_tests, start=1):
        for c_i, val in enumerate(r_data):
            ft_table.rows[r_i].cells[c_i].paragraphs[0].text = val
    style_table(ft_table, ft_widths)

    # 2. Technical Pytest Cases
    add_heading_1("2. Technical Test Cases Specification (TC-T-01 through TC-T-12)")
    add_body_p(
        "Twelve automated Pytest unit and integration test cases cover SQLite persistence, schema assertions, "
        "data cleaning transformations, mathematical KPI formulas, and Streamlit component initialization:"
    )

    tt_headers = ["Test ID", "Target Test Module / Function", "Technical Check / Assertion Executed", "Expected Technical Outcome"]
    tt_widths = [0.8, 2.2, 2.3, 1.4]
    tt_table = doc.add_table(rows=13, cols=4)
    for i, h in enumerate(tt_headers):
        tt_table.rows[0].cells[i].paragraphs[0].text = h

    t_tests = [
        ("TC-T-01", "tests/test_data.py::test_tc_t_01_fact_appointments_volume", "SELECT COUNT(*) FROM fact_appointments;", "Asserts row count >= 150 (Actual = 300)."),
        ("TC-T-02", "tests/test_data.py::test_tc_t_02_analytics_view_columns", "PRAGMA table_info(vw_appointment_analytics);", "Asserts view contains exactly 36 columns."),
        ("TC-T-03", "tests/test_data.py::test_tc_t_03_primary_key_uniqueness", "COUNT(pk) == COUNT(DISTINCT pk) across all 5 tables", "Asserts zero duplicate primary keys."),
        ("TC-T-04", "tests/test_data.py::test_tc_t_04_foreign_key_integrity", "PRAGMA foreign_key_check;", "Asserts zero foreign key constraint violations."),
        ("TC-T-05", "tests/test_data.py::test_tc_t_05_check_constraints_and_ranges", "lead_time_days >= 0 AND status IN ('Show','No-Show')", "Asserts zero negative lead times and valid statuses."),
        ("TC-T-06", "tests/test_data.py::test_tc_t_06_dimension_minimum_rows", "COUNT(*) on clinics, specialties, doctors, patients", "Asserts all master tables meet minimum row baselines."),
        ("TC-T-07", "tests/test_analysis.py::test_tc_t_07_kpi_formulas_on_known_mock", "KPI calculations on tiny mock (8 shows, 2 no-shows)", "Asserts attendance=80.0%, no-show=20.0%, sum=100%."),
        ("TC-T-08", "tests/test_analysis.py::test_tc_t_08_financial_and_capacity_formulas", "Summation: fee * is_no_show; (duration/60) * is_no_show", "Asserts exact fee sum (₹2,100) and hours (1.083h)."),
        ("TC-T-09", "tests/test_analysis.py::test_tc_t_09_two_sample_proportion_z_test", "Z-test formula on reminder vs non-reminder mock", "Asserts Z-stat > 2.5 and relative risk reduction=60%."),
        ("TC-T-10", "tests/test_analysis.py::test_tc_t_10_lead_time_decay_bucketing", "Interval classification function across 0, 2, 5, 12, 31 days", "Asserts exact bucket strings ('0-2 days' to '>30 days')."),
        ("TC-T-11", "tests/test_analysis.py::test_tc_t_11_risk_scoring_and_tier_thresholds", "Score cutoffs: >=60 -> High, 30-59 -> Med, <30 -> Low", "Asserts correct tier string assignments."),
        ("TC-T-12", "tests/test_app.py::test_tc_t_12_streamlit_config_exists", "Verify config.toml, ui.py imports, and data artifacts", "Asserts all configuration files and modules exist cleanly.")
    ]
    for r_i, r_data in enumerate(t_tests, start=1):
        for c_i, val in enumerate(r_data):
            tt_table.rows[r_i].cells[c_i].paragraphs[0].text = val
    style_table(tt_table, tt_widths)

    # 3. Test Suites Overview
    add_heading_1("3. Pytest Automated Test Scripts Architecture")
    add_body_p(
        "The automated test framework is partitioned across three dedicated test modules in the tests/ directory:"
    )
    add_bullet_p("tests/test_data.py: ", "Contains TC-T-01 to TC-T-06. Connects to data/app.db with foreign keys enabled, verifying row counts, primary key uniqueness, foreign key integrity, domain ranges, and dimension baselines.")
    add_bullet_p("tests/test_analysis.py: ", "Contains TC-T-07 to TC-T-11. Tests all mathematical and statistical formulas against controlled mock datasets with known analytical answers.")
    add_bullet_p("tests/test_app.py: ", "Contains TC-T-12. Validates Streamlit theme configuration, component library imports, and dataset artifact existence.")

    doc.add_paragraph().paragraph_format.space_after = Pt(16)
    doc.save(output_path)
    print(f"Step 13 Test Cases Document successfully created at: {output_path}")

def update_final_rtm_with_tests(rtm_path):
    wb = openpyxl.load_workbook(rtm_path)
    ws = wb["Traceability_Matrix"]
    
    test_mappings = {
        "BR-01": "TC-F-01, TC-F-02, TC-F-03, TC-F-04,\nTC-F-15, TC-F-16, TC-T-01, TC-T-02, TC-T-07",
        "BR-02": "TC-F-05, TC-T-05, TC-T-10",
        "BR-03": "TC-F-06, TC-T-06",
        "BR-04": "TC-F-07, TC-T-02, TC-T-06",
        "BR-05": "TC-F-09, TC-T-09",
        "BR-06": "TC-F-08, TC-T-02",
        "BR-07": "TC-F-10, TC-T-09",
        "BR-08": "TC-F-11, TC-T-02",
        "BR-09": "TC-F-12, TC-T-08",
        "BR-10": "TC-F-13, TC-T-11",
        "BR-11": "TC-F-14, TC-T-02",
        "BR-12": "TC-F-14, TC-T-08"
    }
    
    font_code_rtm = Font(name="Calibri", size=8.5, bold=True, color="1F497D")
    
    # Column 3 is Req ID (BR), Column 12 is Test Case (Col L)
    for row in range(5, ws.max_row + 1):
        br_id = ws.cell(row=row, column=3).value
        if br_id in test_mappings:
            cell_tc = ws.cell(row=row, column=12)
            cell_tc.value = test_mappings[br_id]
            cell_tc.font = font_code_rtm
            cell_tc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            
    wb.save(rtm_path)
    print(f"FINAL RTM Excel successfully updated with Test Case IDs at: {rtm_path}")

if __name__ == "__main__":
    doc_out = r"c:\Users\deeps\Downloads\Appointment No Show Analysis\docs\13_Test_Cases.docx"
    rtm_out = r"c:\Users\deeps\Downloads\Appointment No Show Analysis\docs\04_Traceability_Matrix.xlsx"
    build_step13_doc(doc_out)
    update_final_rtm_with_tests(rtm_out)
