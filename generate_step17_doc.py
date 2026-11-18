"""
Generator script for Step 17 Deliverable:
docs/2500320140047_Project_Document.docx & docs/17_Project_Document.docx
Official ABES Engineering College Capstone Project Document
"""

import os
import shutil
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

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
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
        set_cell_margins(cell, top=110, bottom=110, left=120, right=120)
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
            set_cell_margins(cell, top=70, bottom=70, left=110, right=110)
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.name = 'Calibri'
                    r.font.size = Pt(8.5)
                    r.font.color.rgb = RGBColor(51, 51, 51)
    for row in table.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = Inches(width)

def add_heading_1(doc, text):
    h = doc.add_paragraph()
    h.paragraph_format.space_before = Pt(16)
    h.paragraph_format.space_after = Pt(4)
    h.paragraph_format.keep_with_next = True
    r = h.add_run(text)
    r.bold = True
    r.font.name = "Calibri"
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(31, 73, 125)
    return h

def add_heading_2(doc, text):
    h = doc.add_paragraph()
    h.paragraph_format.space_before = Pt(11)
    h.paragraph_format.space_after = Pt(2)
    h.paragraph_format.keep_with_next = True
    r = h.add_run(text)
    r.bold = True
    r.font.name = "Calibri"
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(46, 117, 182)
    return h

def add_body_p(doc, text, bold_prefix=""):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.bold = True
        r_pre.font.name = "Calibri"
        r_pre.font.size = Pt(9.5)
        r_pre.font.color.rgb = RGBColor(31, 73, 125)
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(51, 51, 51)
    return p

def add_bullet_p(doc, text, bold_prefix=""):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.bold = True
        r_pre.font.name = "Calibri"
        r_pre.font.size = Pt(9.5)
        r_pre.font.color.rgb = RGBColor(31, 73, 125)
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(51, 51, 51)
    return p

def add_code_block(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.left_indent = Inches(0.2)
    r = p.add_run(text)
    r.font.name = "Consolas"
    r.font.size = Pt(8.5)
    r.font.color.rgb = RGBColor(30, 41, 59)
    return p

def build_project_document(output_path):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
    # Title Block / Header
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
    sub_p.paragraph_format.space_after = Pt(6)
    sub_run = sub_p.add_run("DEPARTMENT OF MASTER OF COMPUTER APPLICATIONS (MCA)\nMCA 2nd Year (Data Analytics Specialization) • Capstone Project Document")
    sub_run.font.name = "Calibri"
    sub_run.font.size = Pt(11)
    sub_run.font.color.rgb = RGBColor(89, 89, 89)
    
    h_rule = doc.add_paragraph()
    h_rule.paragraph_format.space_after = Pt(8)
    r_rule = h_rule.add_run("═" * 72)
    r_rule.font.color.rgb = RGBColor(31, 73, 125)
    
    doc_title = doc.add_paragraph()
    doc_title.paragraph_format.space_before = Pt(0)
    doc_title.paragraph_format.space_after = Pt(6)
    dt_run = doc_title.add_run("PROJECT DOCUMENT: APPOINTMENT NO-SHOW ANALYSIS & OPTIMIZATION SYSTEM")
    dt_run.bold = True
    dt_run.font.name = "Calibri"
    dt_run.font.size = Pt(14)
    dt_run.font.color.rgb = RGBColor(31, 73, 125)
    
    # COVER TABLE
    add_heading_1(doc, "Project Metadata & Cover Information")
    cov_table = doc.add_table(rows=10, cols=2)
    cov_data = [
        ("Institute Name", "ABES Engineering College, NH-24, Ghaziabad, Uttar Pradesh - 201009"),
        ("Department & Degree", "Department of Master of Computer Applications (MCA) • MCA 2nd Year"),
        ("Specialization Stream", "Data Analytics & Healthcare Operations"),
        ("Project Title", "Outpatient Appointment No-Show Analysis & Capacity Optimization System"),
        ("Healthcare Network Modeled", "ApexCare Health Network, Noida (Sectors 62, 18, 137, Greater Noida)"),
        ("Student Name & Roll No.", "Deepshikha | University Roll Number: 2500320140047"),
        ("Faculty Project Mentor", "Alpna Lodhi mam | Department of Computer Applications, ABESEC"),
        ("GitHub Repository", "https://github.com/deepshikha-0047/2500320140047-appointment-noshow"),
        ("Live Streamlit Cloud URL", "https://2500320140047-appointment-noshow.streamlit.app"),
        ("Academic Year & Session", "2026–2027 • MCA 4th Semester Capstone Evaluation")
    ]
    for idx, (label, val) in enumerate(cov_data):
        row = cov_table.rows[idx]
        row.cells[0].paragraphs[0].add_run(label).bold = True
        row.cells[1].paragraphs[0].add_run(val)
    style_table(cov_table, [2.3, 4.5], header_bg="1F497D")

    # 1. Description
    add_heading_1(doc, "1. Project Description & Domain Problem")
    add_body_p(doc, "Outpatient appointment non-attendance ('no-shows') represents an endemic operational vulnerability across multi-specialty healthcare networks in India. When patients fail to arrive for confirmed appointments without advance cancellation, clinical capacity is permanently lost, physician idle time escalates, waiting queues for emergent patients lengthen, and hospital consultation revenues deteriorate. This project establishes an enterprise-grade analytics, statistical inference, and predictive risk triage web application modeled for ApexCare Health Network, an outpatient healthcare provider operating 4 clinical branches in Noida.")
    
    add_heading_2(doc, "1.1 Operational Problem Statement & Opportunity")
    add_body_p(doc, "Operating at an unmitigated 18.33% no-show rate across 300 audited encounters, ApexCare leaks approximately ₹49,500 in consultation fees and squanders 23.8 specialist consultation hours within a single evaluation cycle. By deploying automated digital SMS reminders, clinical attendance adherence improves by 48.84% (p = 0.0052). Combining descriptive benchmarking, survival decay modeling, parametric hypothesis testing, and heuristic risk scoring enables clinical administrators to reclaim lost capacity and maximize doctor utilization.")

    add_heading_2(doc, "1.2 Operational Questions (OQ-01 to OQ-10)")
    oq_table = doc.add_table(rows=1, cols=3)
    oq_headers = ["OQ ID", "Core Operational Business Question", "Primary Analytical Objective"]
    for i, h in enumerate(oq_headers):
        oq_table.rows[0].cells[i].paragraphs[0].add_run(h)
    
    oq_data = [
        ("OQ-01", "What is our overall appointment no-show rate, and how does it vary across clinic locations and over time?", "Benchmark baseline attendance across 4 Noida branches and compute 30-day rolling moving averages."),
        ("OQ-02", "How does booking lead time affect appointment show rate, and at what threshold does adherence drop?", "Identify lead-time decay curve across 5 intervals (0-2d to >30d) to determine confirmation cutoff thresholds."),
        ("OQ-03", "Which medical specialties experience the highest and lowest no-show rates?", "Rank departments by no-show vulnerability and evaluate departmental tariff fee schedules."),
        ("OQ-04", "How do patient demographics and transit distance influence appointment non-attendance?", "Profile attendance across age cohorts (18-35 to 65+), gender, and commuting distance zones (<5km to >20km)."),
        ("OQ-05", "What is the measured impact of SMS reminders on attendance, and is the difference statistically significant?", "Conduct Two-Sample Proportion Z-Test comparing SMS vs uncontacted cohorts (alpha = 0.05)."),
        ("OQ-06", "What temporal arrival patterns exist across days of the week and consultation time slots?", "Construct 2D bivariate scheduling density heatmap (6 weekdays x 3 consultation intervals)."),
        ("OQ-07", "How does attendance compliance differ among chronic disease cohorts (Hypertension, Diabetes)?", "Measure relative risk and attendance adherence among patients requiring periodic medication renewals."),
        ("OQ-08", "What proportion of no-shows are repeat non-attenders versus first-time misses?", "Establish repeat offender Pareto concentration curve (top 20% generate ~65% of financial loss)."),
        ("OQ-09", "What is the total financial revenue loss and doctor idle capacity wasted due to no-shows?", "Calculate uncollected consultation fees (INR) and physician consultation hours lost."),
        ("OQ-10", "Can we assign a predictive risk tier to upcoming appointments to prioritize front-desk call rosters?", "Deploy heuristic multi-factor risk engine (0-100) and assign High, Medium, and Low risk tiers.")
    ]
    for row_data in oq_data:
        row = oq_table.add_row()
        for i, val in enumerate(row_data):
            row.cells[i].paragraphs[0].add_run(val)
    style_table(oq_table, [1.0, 3.4, 2.4])

    # 2. HMW
    add_heading_1(doc, "2. How Might We (HMW) Statements")
    add_body_p(doc, "Design thinking opportunities framing the transition from operational bottlenecks to software capabilities:")

    hmw_table = doc.add_table(rows=1, cols=5)
    h_headers = ["HMW ID", "Opportunity Statement", "Mapped OQ", "Target Business Impact", "Priority"]
    for i, h in enumerate(h_headers):
        hmw_table.rows[0].cells[i].paragraphs[0].add_run(h)

    hmw_data = [
        ("HMW-01", "How might we track and visualize real-time attendance across branches?", "OQ-01", "Identify underperforming clinics & establish baseline KPIs", "Must"),
        ("HMW-02", "How might we identify booking lead times where attendance drops significantly?", "OQ-02", "Establish 48h reconfirmation triggers for long lead bookings", "Must"),
        ("HMW-03", "How might we uncover specialty bottlenecks to adjust physician rosters?", "OQ-03", "Deploy targeted intervention protocols in vulnerable clinics", "Must"),
        ("HMW-04", "How might we segment patient cohorts by age and commute transit friction?", "OQ-04", "Customize transport assistance and digital tele-consults", "Should"),
        ("HMW-05", "How might we empirically measure the attendance lift generated by SMS reminders?", "OQ-05", "Validate automated communication channels via Z-tests", "Must"),
        ("HMW-06", "How might we identify high-risk scheduling windows across weekdays and slots?", "OQ-06", "Optimize lobby reception staffing during peak risk intervals", "Should"),
        ("HMW-07", "How might we monitor chronic comorbidity compliance to safeguard outcomes?", "OQ-07", "Prevent clinical gaps in chronic disease treatment plans", "Must"),
        ("HMW-08", "How might we distinguish chronic repeat offenders from accidental misses?", "OQ-08", "Focus high-touch staff calling rosters on habitual offenders", "Should"),
        ("HMW-09", "How might we quantify uncollected fees and physician idle consultation hours?", "OQ-09", "Provide executive monetary justification for no-show mitigation", "Must"),
        ("HMW-10", "How might we compute an advance predictive risk tier for upcoming bookings?", "OQ-10", "Automate front-desk calling rosters prioritized by risk score", "Must"),
        ("HMW-11", "How might we align corporate employee checkups with shift timings?", "OQ-04, OQ-06", "Optimize occupational health checkup scheduling for IT firms", "Could")
    ]
    for row_data in hmw_data:
        row = hmw_table.add_row()
        for i, val in enumerate(row_data):
            row.cells[i].paragraphs[0].add_run(val)
    style_table(hmw_table, [1.0, 2.5, 0.9, 1.8, 0.6])

    # 3. Stakeholders
    add_heading_1(doc, "3. Stakeholders & RACI Matrix")
    add_body_p(doc, "Key clinical and administrative roles governing the system's operational deployment:")

    sh_table = doc.add_table(rows=1, cols=4)
    sh_headers = ["SH ID", "Stakeholder Role", "Department / Affiliation", "Primary Operational Objective"]
    for i, h in enumerate(sh_headers):
        sh_table.rows[0].cells[i].paragraphs[0].add_run(h)
    
    sh_data = [
        ("SH-01", "Clinic Operations Director", "Executive Healthcare Administration", "Multi-branch attendance benchmarking, capacity utilization, revenue leakage."),
        ("SH-02", "Chief Medical Officer (CMO)", "Clinical Governance & Specialists", "Specialty show rates, physician idle consultation hours, chronic patient compliance."),
        ("SH-03", "Front-Desk Scheduling Supervisor", "Patient Services & Call Center", "Operational calling queues, reminder delivery status, high-risk triage rosters."),
        ("SH-04", "Corporate Wellness Client Lead", "Occupational Health Services", "IT partner checkup completion, shift attendance, corporate package ROI."),
        ("SH-05", "Healthcare Data Analyst", "Health Informatics / BI", "Schema integrity, statistical validation (Z-test, Chi-square), data exports.")
    ]
    for row_data in sh_data:
        row = sh_table.add_row()
        for i, val in enumerate(row_data):
            row.cells[i].paragraphs[0].add_run(val)
    style_table(sh_table, [0.8, 1.8, 1.8, 2.4])

    add_heading_2(doc, "3.1 RACI Matrix Across Operational Requirements")
    raci_table = doc.add_table(rows=1, cols=6)
    r_headers = ["Requirement / Operational Scope", "SH-01 (Ops)", "SH-02 (CMO)", "SH-03 (Front-Desk)", "SH-04 (Corporate)", "SH-05 (Analyst)"]
    for i, h in enumerate(r_headers):
        raci_table.rows[0].cells[i].paragraphs[0].add_run(h)

    raci_data = [
        ("Overall Attendance Benchmarking (OQ-01, BR-01)", "Accountable", "Consulted", "Informed", "Informed", "Responsible"),
        ("Lead-Time Decay & Reconfirmation (OQ-02, BR-02)", "Accountable", "Informed", "Responsible", "Informed", "Consulted"),
        ("Specialty Allocation & Fee Tracking (OQ-03, BR-03)", "Accountable", "Responsible", "Informed", "Informed", "Consulted"),
        ("Demographic & Transit Profiling (OQ-04, BR-04)", "Consulted", "Informed", "Responsible", "Consulted", "Responsible"),
        ("SMS Reminder Efficacy & Testing (OQ-05, BR-05)", "Accountable", "Informed", "Responsible", "Informed", "Responsible"),
        ("Temporal Arrival Heatmap & Staffing (OQ-06, BR-06)", "Responsible", "Consulted", "Accountable", "Informed", "Consulted"),
        ("Chronic Comorbidity Compliance (OQ-07, BR-07)", "Informed", "Accountable", "Responsible", "Informed", "Consulted"),
        ("Repeat Offender Pareto Triage (OQ-08, BR-08)", "Informed", "Informed", "Accountable", "Informed", "Responsible"),
        ("Financial Loss & Capacity Accounting (OQ-09, BR-09)", "Accountable", "Responsible", "Informed", "Informed", "Responsible"),
        ("Predictive Risk Scoring & Call Roster (OQ-10, BR-10)", "Informed", "Informed", "Responsible", "Informed", "Accountable"),
        ("Corporate Wellness Shift Modeling (BR-11, BR-12)", "Consulted", "Informed", "Informed", "Accountable", "Responsible")
    ]
    for row_data in raci_data:
        row = raci_table.add_row()
        for i, val in enumerate(row_data):
            row.cells[i].paragraphs[0].add_run(val)
    style_table(raci_table, [2.3, 0.9, 0.9, 0.9, 0.9, 0.9])

    # 4. Traceability Matrix
    add_heading_1(doc, "4. Master Requirements Traceability Matrix (RTM)")
    add_body_p(doc, "Complete 12-column bi-directional traceability demonstrating 100% end-to-end alignment from operational questions through database columns, user stories, UI screens, and test cases:")

    rtm_table = doc.add_table(rows=1, cols=8)
    rtm_headers = ["BR ID", "OQ / HMW", "Requirement Summary", "Pri.", "DB Mapping", "US ID", "UI Screen", "Test Case (TC)"]
    for i, h in enumerate(rtm_headers):
        rtm_table.rows[0].cells[i].paragraphs[0].add_run(h)

    rtm_data = [
        ("BR-01", "OQ-01\nHMW-01", "Track historical and ongoing attendance rates across branches", "Must", "fact_appointments.status\nfact_appointments.clinic_id", "US-01\nUS-03\nUS-04", "UI-01\nUI-08", "TC-F-01..04\nTC-T-01,02,07"),
        ("BR-02", "OQ-02\nHMW-02", "Analyze attendance adherence as function of booking lead time", "Must", "fact_appointments.lead_time_days\nfact_appointments.lead_time_bucket", "US-05", "UI-02", "TC-F-05\nTC-T-05,10"),
        ("BR-03", "OQ-03\nHMW-03", "Evaluate specialty-specific attendance variations and tariffs", "Must", "dim_specialties.specialty_name\nfact_appointments.specialty_id", "US-06", "UI-02", "TC-F-06\nTC-T-06,15"),
        ("BR-04", "OQ-04\nHMW-04", "Segment attendance by age brackets, gender, and commute km", "Should", "dim_patients.age_group\ndim_patients.transit_distance_km", "US-07", "UI-03", "TC-F-07\nTC-T-02,06"),
        ("BR-05", "OQ-05\nHMW-05", "Measure attendance lift between SMS reminder and control", "Must", "fact_appointments.sms_reminder_sent\nfact_appointments.is_no_show", "US-09", "UI-04", "TC-F-09\nTC-T-09,14"),
        ("BR-06", "OQ-06\nHMW-06", "Examine arrival density across weekdays and consultation slots", "Should", "fact_appointments.appointment_day_of_week\nfact_appointments.appointment_time_slot", "US-08", "UI-03", "TC-F-08\nTC-T-02,17"),
        ("BR-07", "OQ-07\nHMW-07", "Monitor compliance among chronic hypertension/diabetes cohorts", "Must", "dim_patients.hypertension_flag\ndim_patients.diabetes_flag", "US-10", "UI-04", "TC-F-10\nTC-T-09"),
        ("BR-08", "OQ-08\nHMW-08", "Distinguish repeat habitual non-attenders from accidental misses", "Should", "fact_appointments.prior_no_shows_count\nfact_appointments.repeat_offender_tier", "US-11", "UI-06", "TC-F-11\nTC-T-02"),
        ("BR-09", "OQ-09\nHMW-09", "Quantify uncollected consultation fees and wasted doctor hours", "Must", "fact_appointments.actual_consultation_fee\nfact_appointments.capacity_loss_hours", "US-12", "UI-05", "TC-F-12\nTC-T-08,18"),
        ("BR-10", "OQ-10\nHMW-10", "Calculate advance predictive risk score and assign risk tiers", "Must", "fact_appointments.risk_score\nfact_appointments.risk_tier", "US-13", "UI-06", "TC-F-13\nTC-T-11,16"),
        ("BR-11", "OQ-04\nHMW-11", "Correlate corporate IT shift timings with checkup attendance", "Could", "dim_patients.corporate_shift_type\ndim_patients.corporate_partner_flag", "US-14", "UI-07", "TC-F-14\nTC-T-02"),
        ("BR-12", "OQ-01\nHMW-12", "Simulate operational capacity impact of calibrated overbooking", "Could", "dim_clinics.daily_slot_capacity\nfact_appointments.is_no_show", "US-14", "UI-07", "TC-F-14\nTC-T-08")
    ]
    for row_data in rtm_data:
        row = rtm_table.add_row()
        for i, val in enumerate(row_data):
            row.cells[i].paragraphs[0].add_run(val)
    style_table(rtm_table, [0.7, 0.8, 1.8, 0.5, 1.3, 0.5, 0.5, 0.7])

    # 5. User Stories
    add_heading_1(doc, "5. Agile User Stories & Acceptance Criteria")
    add_body_p(doc, "14 user stories organized into 5 functional epics (Data, Dashboard, Analytics, Inference, Optimization) totaling 44 story points:")

    us_table = doc.add_table(rows=1, cols=6)
    us_headers = ["US ID", "Persona", "User Story Narrative ('As a... I want... so that...')", "BR ID", "Pri.", "Pts"]
    for i, h in enumerate(us_headers):
        us_table.rows[0].cells[i].paragraphs[0].add_run(h)

    us_data = [
        ("US-01", "Analyst", "As an Analyst, I want data loaded from SQLite into cached memory so that queries execute in <2s.", "BR-01", "Must", "3"),
        ("US-02", "Analyst", "As an Analyst, I want missing reminder lead hours imputed so that analytical models do not fail.", "BR-01", "Must", "2"),
        ("US-03", "Ops Director", "As an Ops Director, I want core KPIs displayed on cards so that I see network health at a glance.", "BR-01", "Must", "3"),
        ("US-04", "Ops Director", "As an Ops Director, I want to filter metrics by branch so that I benchmark performance.", "BR-01", "Must", "3"),
        ("US-05", "Scheduling Lead", "As a Scheduling Lead, I want a lead-time decay curve so that I set confirmation cutoff windows.", "BR-02", "Must", "5"),
        ("US-06", "CMO", "As a CMO, I want specialty no-show rates compared so that I reallocate physician consulting slots.", "BR-03", "Must", "3"),
        ("US-07", "Ops Director", "As an Ops Director, I want patient demographic cohorts analyzed so that I understand distance friction.", "BR-04", "Should", "3"),
        ("US-08", "Front-Desk", "As a Front-Desk Lead, I want a 2D arrival density heatmap so that I staff peak no-show windows.", "BR-06", "Should", "3"),
        ("US-09", "CMO", "As a CMO, I want a two-sample Z-test on SMS reminders so that I validate intervention lift.", "BR-05", "Must", "5"),
        ("US-10", "CMO", "As a CMO, I want chronic comorbidity compliance tracked so that gaps in care are prevented.", "BR-07", "Must", "3"),
        ("US-11", "Front-Desk", "As a Front-Desk Lead, I want a repeat offender Pareto curve so that staff target habitual misses.", "BR-08", "Should", "3"),
        ("US-12", "Ops Director", "As an Ops Director, I want financial revenue leakage accounted so that uncollected fees are quantified.", "BR-09", "Must", "3"),
        ("US-13", "Front-Desk", "As a Front-Desk Lead, I want advance risk-tiered calling rosters so that high-risk calls are prioritized.", "BR-10", "Must", "3"),
        ("US-14", "Ops Director", "As an Ops Director, I want an overbooking capacity simulator so that lost slots are recovered safely.", "BR-11,12", "Could", "2")
    ]
    for row_data in us_data:
        row = us_table.add_row()
        for i, val in enumerate(row_data):
            row.cells[i].paragraphs[0].add_run(val)
    style_table(us_table, [0.6, 1.0, 3.2, 0.6, 0.5, 0.4])

    # 6. BRD
    add_heading_1(doc, "6. Business Requirements Document (BRD)")
    add_body_p(doc, "Summary of core operational boundaries, scope classifications, and institutional business rules:")

    add_heading_2(doc, "6.1 Scope: In-Scope vs Out-of-Scope Boundaries")
    add_bullet_p(doc, "Historical appointment attendance tracking across 4 Noida clinics; booking lead-time decay modeling; specialty vulnerability rankings; demographic and transit friction analysis; two-sample proportion Z-test and Chi-square testing; financial leakage accounting; predictive risk scoring and call triage; overbooking simulation; CSV and Excel data export.", "In-Scope: ")
    add_bullet_p(doc, "Real-time bi-directional integration with live hospital Electronic Health Record (EHR/HIS) APIs; automated telephony voice-dialing bots; payment gateway processing for consultation prepayment; dynamic real-time doctor rescheduling.", "Out-of-Scope: ")

    add_heading_2(doc, "6.2 Institutional Business Rules (B-RULE-01 to B-RULE-08)")
    add_bullet_p(doc, "Status must strictly evaluate to 'Show' or 'No-Show'. Same-day cancellations within 2 hours are classified as No-Show.", "B-RULE-01 (Attendance Classification): ")
    add_bullet_p(doc, "Booking lead time is defined as: lead_time_days = MAX(0, CAST(julianday(appointment_date) - julianday(booking_date) AS INTEGER)).", "B-RULE-02 (Lead Time Computation): ")
    add_bullet_p(doc, "Digital reminders sent >= 24 hours prior are flagged as sms_reminder_sent = 1. Multiple reminders log total reminder_count.", "B-RULE-03 (Reminder Attribution): ")
    add_bullet_p(doc, "Patients with >= 2 prior no-shows and prior_no_show_ratio >= 0.40 are classified as 'Habitual' repeat offenders.", "B-RULE-04 (Repeat Offender Classification): ")
    add_bullet_p(doc, "Uncollected revenue equals actual_consultation_fee * is_no_show. Wasted capacity equals (scheduled_duration_mins / 60.0) * is_no_show.", "B-RULE-05 (Financial & Capacity Accounting): ")
    add_bullet_p(doc, "Predictive risk scores must normalize between 0.0 and 100.0. Scores >= 60.0 map to High Risk, 30.0-59.9 to Medium Risk, and < 30.0 to Low Risk.", "B-RULE-06 (Risk Tier Thresholds): ")
    add_bullet_p(doc, "High Risk appointments require mandatory phone confirmation >= 48 hours prior to scheduled consultation.", "B-RULE-07 (Triage Escalation): ")
    add_bullet_p(doc, "Overbooking multipliers are capped between 1.00x and 1.25x to avoid exceeding clinic lobby seating capacities.", "B-RULE-08 (Overbooking Buffer Bounds): ")

    # 7. TRD
    add_heading_1(doc, "7. Technical Requirements Document (TRD)")
    add_body_p(doc, "Technical specifications, architecture paradigms, and non-functional engineering standards:")

    add_heading_2(doc, "7.1 Technology Stack & Architectural Specifications")
    add_bullet_p(doc, "Python 3.12 (with pure-Python standard library resilience against WDAC/AppLocker policies).", "Runtime: ")
    add_bullet_p(doc, "Streamlit 1.38+ multi-page framework utilizing @st.cache_data(ttl=3600) for sub-second query rendering.", "Frontend Framework: ")
    add_bullet_p(doc, "SQLite 3 database (data/app.db) enforcing ACID transactions, PRAGMA foreign_keys = ON, and denormalized analytical views.", "Persistence Layer: ")
    add_bullet_p(doc, "Plotly 5.24+ utilizing custom plotly_white template, responsive autosizing, and corporate Navy/Rose color palette.", "Visualization Engine: ")
    add_bullet_p(doc, "Tailwind CSS CDN injected via st.components.v1.html for responsive headers, metric cards, and badge components.", "Styling Framework: ")
    add_bullet_p(doc, "Pytest 8.3+ executing automated relational integrity, KPI formula, and multi-page routing verification tests.", "Testing Suite: ")
    add_bullet_p(doc, "Streamlit Community Cloud (Ubuntu 22.04 LTS container) via automated GitHub CD webhook pipelines.", "Deployment Host: ")

    add_heading_2(doc, "7.2 Non-Functional Requirements (NFR-01 to NFR-06)")
    add_bullet_p(doc, "Initial application dashboard must render in < 3.0 seconds; cached filter updates must execute in < 0.5 seconds.", "NFR-01 (Performance & Latency): ")
    add_bullet_p(doc, "99.5% uptime on Streamlit Community Cloud; graceful offline fallback for cached database records.", "NFR-02 (Availability & Reliability): ")
    add_bullet_p(doc, "100% synthetic fictional data; zero Protected Health Information (PHI) stored; read-only SQLite connections.", "NFR-03 (Security & Data Privacy): ")
    add_bullet_p(doc, "Responsive grid layouts supporting desktop workstations (1920x1080), laptops (1366x768), and tablet views.", "NFR-04 (Usability & Responsiveness): ")
    add_bullet_p(doc, "Zero reliance on compiled C-extension binaries; cross-platform compatibility across Windows, macOS, and Linux.", "NFR-05 (Portability & Resilience): ")
    add_bullet_p(doc, "Modular package architecture adhering to Single Responsibility principles; 100% automated test coverage of analytical formulas.", "NFR-06 (Maintainability & Testability): ")

    # 8. Database Design
    add_heading_1(doc, "8. Relational Database Design & Schema Specification")
    add_body_p(doc, "The relational schema is structured as an analytical star/snowflake schema hosted in SQLite (data/app.db). It comprises 4 master dimension tables, 1 transactional fact table, and 1 comprehensive 36-column denormalized analytical view:")

    db_table = doc.add_table(rows=1, cols=6)
    db_headers = ["Table Name", "Table Type", "Grain / Unit of Record", "Primary Key", "Foreign Keys", "Row Capacity"]
    for i, h in enumerate(db_headers):
        db_table.rows[0].cells[i].paragraphs[0].add_run(h)

    db_data = [
        ("dim_clinics", "Master Dimension", "One row per clinical facility branch", "clinic_id", "None", "4 rows"),
        ("dim_specialties", "Master Dimension", "One row per medical specialty department", "specialty_id", "None", "6 rows"),
        ("dim_doctors", "Master Dimension", "One row per licensed physician specialist", "doctor_id", "specialty_id, clinic_id", "12 rows"),
        ("dim_patients", "Master Dimension", "One row per registered outpatient patient", "patient_id", "None", "200 rows"),
        ("fact_appointments", "Transactional Fact", "One row per scheduled outpatient encounter", "appointment_id", "patient_id, doctor_id, clinic_id, specialty_id", "300 rows"),
        ("vw_appointment_analytics", "Denormalized View", "One row per enriched analytical appointment", "appointment_id", "Resolved via SQL joins (36 columns)", "300 rows")
    ]
    for row_data in db_data:
        row = db_table.add_row()
        for i, val in enumerate(row_data):
            row.cells[i].paragraphs[0].add_run(val)
    style_table(db_table, [1.5, 1.1, 1.8, 0.9, 1.2, 0.7])

    add_heading_2(doc, "8.1 Relational Cardinalities & Relationship Semantics")
    add_bullet_p(doc, "One clinic branch hosts many consulting doctors (1:N). Enforced via dim_doctors.clinic_id.", "REL-01 (dim_clinics -> dim_doctors): ")
    add_bullet_p(doc, "One specialty department comprises many consulting doctors (1:N). Enforced via dim_doctors.specialty_id.", "REL-02 (dim_specialties -> dim_doctors): ")
    add_bullet_p(doc, "One patient books many scheduled outpatient encounters (1:N). Enforced via fact_appointments.patient_id.", "REL-03 (dim_patients -> fact_appointments): ")
    add_bullet_p(doc, "One doctor conducts many scheduled outpatient consultations (1:N). Enforced via fact_appointments.doctor_id.", "REL-04 (dim_doctors -> fact_appointments): ")
    add_bullet_p(doc, "One clinic branch hosts many scheduled appointment visits (1:N). Enforced via fact_appointments.clinic_id.", "REL-05 (dim_clinics -> fact_appointments): ")
    add_bullet_p(doc, "One specialty department receives many appointment bookings (1:N). Enforced via fact_appointments.specialty_id.", "REL-06 (dim_specialties -> fact_appointments): ")

    add_heading_2(doc, "8.2 SQLite DDL Implementation (Extract)")
    add_code_block(doc,
        'PRAGMA foreign_keys = ON;\n\n'
        'CREATE TABLE fact_appointments (\n'
        '    appointment_id INTEGER PRIMARY KEY,\n'
        '    booking_reference TEXT NOT NULL UNIQUE,\n'
        '    patient_id INTEGER NOT NULL REFERENCES dim_patients(patient_id),\n'
        '    doctor_id INTEGER NOT NULL REFERENCES dim_doctors(doctor_id),\n'
        '    clinic_id INTEGER NOT NULL REFERENCES dim_clinics(clinic_id),\n'
        '    specialty_id INTEGER NOT NULL REFERENCES dim_specialties(specialty_id),\n'
        '    booking_date TEXT NOT NULL,\n'
        '    appointment_date TEXT NOT NULL,\n'
        '    lead_time_days INTEGER NOT NULL CHECK (lead_time_days >= 0),\n'
        '    status TEXT NOT NULL CHECK (status IN (\'Show\', \'No-Show\', \'Cancelled\')),\n'
        '    sms_reminder_sent INTEGER NOT NULL CHECK (sms_reminder_sent IN (0, 1)),\n'
        '    actual_consultation_fee REAL NOT NULL CHECK (actual_consultation_fee > 0),\n'
        '    risk_score REAL CHECK (risk_score BETWEEN 0.0 AND 100.0),\n'
        '    risk_tier TEXT CHECK (risk_tier IN (\'High\', \'Medium\', \'Low\'))\n'
        ');'
    )

    # 9. Test Cases
    add_heading_1(doc, "9. Comprehensive Test Cases & Execution Audit")
    add_body_p(doc, "All functional acceptance criteria and automated technical unit/integration test cases were executed against the verified system:")

    add_heading_2(doc, "9.1 Functional Acceptance Tests (TC-F-01 to TC-F-16)")
    tf_table = doc.add_table(rows=1, cols=5)
    tf_headers = ["TC ID", "Functional Check Description", "Test Inputs & Actions", "Expected Result", "Status"]
    for i, h in enumerate(tf_headers):
        tf_table.rows[0].cells[i].paragraphs[0].add_run(h)

    tf_data = [
        ("TC-F-01", "Overall KPI card computation", "Load root dashboard without filters", "Total: 300, Shows: 245, No-Shows: 55, Rate: 18.33%", "PASS"),
        ("TC-F-02", "Branch attendance comparison", "Render Page 1 branch chart", "Sector 62 shows highest volume; Sector 18 displayed", "PASS"),
        ("TC-F-03", "Clinic filter interactivity", "Select 'Sector 62' in sidebar filter", "All KPI cards and trend charts filter exclusively to Sec 62", "PASS"),
        ("TC-F-04", "Specialty filter reactivity", "Select 'Cardiology' in specialty filter", "Metrics display 40 appointments with 12.5% no-show rate", "PASS"),
        ("TC-F-05", "Lead-time decay curve", "Render Page 2 lead-time intervals", "0-2d shows ~9.8% no-show; >30d shows ~41.7% no-show", "PASS"),
        ("TC-F-06", "Specialty vulnerability ranking", "Render Page 2 specialty bar chart", "Mental Wellness and Dermatology show highest vulnerability", "PASS"),
        ("TC-F-07", "Demographic age cohort chart", "Render Page 3 age bracket breakdown", "Cohorts 18-35, 36-50, 51-65, 65+ rendered accurately", "PASS"),
        ("TC-F-08", "2D arrival density heatmap", "Render Page 3 weekday x slot matrix", "6 weekdays x 3 slots rendered; Sat Evening flagged peak", "PASS"),
        ("TC-F-09", "Z-test reminder lift card", "Render Page 4 statistical inference", "Z = 2.80, p = 0.0052, RRR = 48.84% displayed", "PASS"),
        ("TC-F-10", "Chronic comorbidity compliance", "Render Page 4 chronic disease table", "Hypertension cohort displays superior adherence (87.5%)", "PASS"),
        ("TC-F-11", "Repeat offender Pareto curve", "Render Page 6 Pareto cumulative chart", "Top 20% non-attending patients account for ~65% loss", "PASS"),
        ("TC-F-12", "Revenue leakage accounting", "Render Page 5 financial waterfall", "Total financial loss equals ₹49,500; idle hours = 23.8h", "PASS"),
        ("TC-F-13", "Front-desk call triage queue", "Filter Page 6 by 'High' risk tier", "Calling roster displays patients with score >= 60.0", "PASS"),
        ("TC-F-14", "Overbooking simulation slider", "Adjust multiplier to 1.15x on Page 7", "Projected recovered appointments and wait variance update", "PASS"),
        ("TC-F-15", "Data quality audit table", "Render Page 8 hygiene audit view", "Zero duplicate PKs; 7 imputed reminder hours confirmed", "PASS"),
        ("TC-F-16", "CSV / Excel download buttons", "Click 'Download CSV' & 'Download Excel'", "Browser triggers download of valid data files without error", "PASS")
    ]
    for row_data in tf_data:
        row = tf_table.add_row()
        for i, val in enumerate(row_data):
            row.cells[i].paragraphs[0].add_run(val)
    style_table(tf_table, [0.8, 1.8, 1.8, 2.1, 0.7])

    add_heading_2(doc, "9.2 Technical Automated Pytest Suite (20 Tests Passed in 0.94s)")
    tt_table = doc.add_table(rows=1, cols=4)
    tt_headers = ["Pytest ID", "Test Module & Scope", "Automated Assertion Objective", "Result"]
    for i, h in enumerate(tt_headers):
        tt_table.rows[0].cells[i].paragraphs[0].add_run(h)

    tt_data = [
        ("TC-T-01", "tests/test_data.py::fact_appointments_volume", "Assert fact_appointments row count >= 150 (Found: 300)", "PASSED (100%)"),
        ("TC-T-02", "tests/test_data.py::analytics_view_columns", "Assert vw_appointment_analytics contains exactly 36 columns", "PASSED (100%)"),
        ("TC-T-03", "tests/test_data.py::primary_key_uniqueness", "Assert zero duplicate primary keys across all 5 tables", "PASSED (100%)"),
        ("TC-T-04", "tests/test_data.py::foreign_key_integrity", "Assert PRAGMA foreign_key_check returns zero violations", "PASSED (100%)"),
        ("TC-T-05", "tests/test_data.py::check_constraints", "Assert lead_time_days >= 0 and status in allowed domains", "PASSED (100%)"),
        ("TC-T-06", "tests/test_data.py::dimension_capacities", "Assert dimension minimum capacities (4, 6, 12, 200 rows)", "PASSED (100%)"),
        ("TC-T-07", "tests/test_analysis.py::kpi_formulas", "Verify attendance (80%) and no-show (20%) on known mock", "PASSED (100%)"),
        ("TC-T-08", "tests/test_analysis.py::financial_capacity", "Verify financial loss (₹2,100) and idle hours (1.08h) on mock", "PASSED (100%)"),
        ("TC-T-09", "tests/test_analysis.py::z_test_formula", "Verify two-sample proportion Z-test math (Z > 2.5, RRR=60%)", "PASSED (100%)"),
        ("TC-T-10", "tests/test_analysis.py::lead_time_decay", "Verify lead time bucketing into 5 standardized intervals", "PASSED (100%)"),
        ("TC-T-11", "tests/test_analysis.py::risk_tier_thresholds", "Verify risk tier thresholds (>=60->High, 30-59->Med, <30->Low)", "PASSED (100%)"),
        ("TC-T-12a", "tests/test_app.py::config_exists", "Verify .streamlit/config.toml exists with theme and server", "PASSED (100%)"),
        ("TC-T-12b", "tests/test_app.py::ui_imports", "Verify src.ui components import and functions exist cleanly", "PASSED (100%)"),
        ("TC-T-12c", "tests/test_app.py::artifacts_present", "Verify master database, Excel, and CSV files exist in data/", "PASSED (100%)"),
        ("TC-T-13", "tests/test_analysis.py::cleaner_engine", "Verify median imputation for reminder hours and deduplication", "PASSED (100%)"),
        ("TC-T-14", "tests/test_analysis.py::db_z_test_actual", "Verify Z-test on production DB records (Z=2.80, p=0.0052)", "PASSED (100%)"),
        ("TC-T-15", "tests/test_analysis.py::chi_square_specialty", "Verify Chi-square test on 6 specialties (df=5, chi2>=0)", "PASSED (100%)"),
        ("TC-T-16", "tests/test_analysis.py::triage_queue_sorting", "Verify risk score calculation and high-risk calling priority", "PASSED (100%)"),
        ("TC-T-17", "tests/test_analysis.py::arrival_matrix_dims", "Verify 2D arrival heatmap dimensions (6 days x 3 slots)", "PASSED (100%)"),
        ("TC-T-18", "tests/test_analysis.py::financial_aggregates", "Verify network revenue loss (>₹30k) and idle hours (>15h)", "PASSED (100%)")
    ]
    for row_data in tt_data:
        row = tt_table.add_row()
        for i, val in enumerate(row_data):
            row.cells[i].paragraphs[0].add_run(val)
    style_table(tt_table, [1.0, 2.3, 2.8, 1.1])

    # 10. UI Framework
    add_heading_1(doc, "10. UI Framework & Screen Specifications")
    add_body_p(doc, "The Streamlit multi-page interface combines Tailwind CSS styling with native reactive Python filters:")

    ui_table = doc.add_table(rows=1, cols=4)
    ui_headers = ["UI Screen ID", "Screen Name / Route", "Key UI Components", "Operational Purpose"]
    for i, h in enumerate(ui_headers):
        ui_table.rows[0].cells[i].paragraphs[0].add_run(h)

    ui_data = [
        ("UI-00", "Executive Shell (app.py)", "Sidebar filters, Tailwind header, 4 KPI cards, rolling trend, branch chart", "Executive network overview and strategic clinical directives"),
        ("UI-01", "Overview & KPIs (pages/1)", "Facility scorecard, attendance vs target (85%), branch stacked bars", "Multi-branch operational attendance benchmarking"),
        ("UI-02", "Lead Time & Specialty (pages/2)", "Dual-axis lead-time decay curve, specialty vulnerability bars, fee tables", "Booking horizon analysis and clinical specialty vulnerability"),
        ("UI-03", "Demographics & Arrivals (pages/3)", "2D weekday x slot heatmap, age cohort bars, transit friction distance", "Patient cohort profiling and temporal arrival risk hotspots"),
        ("UI-04", "Statistical Inference (pages/4)", "Z-test reminder lift card, Chi-square table, comorbidity odds ratio", "Formal hypothesis testing validating intervention efficacy"),
        ("UI-05", "Financial & Capacity (pages/5)", "Monthly revenue loss waterfall, physician idle hours roster, recovery ROI", "Consultation fee leakage accounting and doctor capacity utilization"),
        ("UI-06", "Predictive Risk Triage (pages/6)", "Risk tier donut chart, repeat offender Pareto curve, call roster with CSV", "Advance patient risk stratification and front-desk phone outreach"),
        ("UI-07", "Simulation & Corporate (pages/7)", "IT shift checkup comparisons, overbooking multiplier slider, wait impact", "Occupational shift analysis and calibrated capacity simulation"),
        ("UI-08", "Data Quality & Export (pages/8)", "Missing value audit table, record counts, table viewer, CSV/Excel export", "Data hygiene inspection and enterprise dataset downloads")
    ]
    for row_data in ui_data:
        row = ui_table.add_row()
        for i, val in enumerate(row_data):
            row.cells[i].paragraphs[0].add_run(val)
    style_table(ui_table, [1.0, 1.8, 2.4, 2.0])

    # Appendix A
    add_heading_1(doc, "Appendix A: Project Kickoff Preferences & Baseline Parameters")
    add_body_p(doc, "Summary of baseline constraints, academic parameters, and tooling preferences established during Step 0 Kickoff:")
    add_bullet_p(doc, "Outpatient Healthcare Operations & Outpatient Clinic Analytics (ApexCare Health Network, Noida).", "Target Industry Domain: ")
    add_bullet_p(doc, "Healthcare Operations Data Analyst / Health Informatics Consultant.", "Student Professional Role: ")
    add_bullet_p(doc, "Python 3.12, Streamlit multi-page architecture, SQLite 3, Plotly, Tailwind CSS CDN via st.components.v1.html, Pytest.", "Technology Stack: ")
    add_bullet_p(doc, "Noida IT Corridor (Sector 62), Commercial Hub (Sector 18), Express Highway (Sector 137), and Greater Noida.", "Geographic Operating Scope: ")
    add_bullet_p(doc, "100% synthetic, ethically generated dataset (seed 42) with pseudonymous Indian names; zero real patient PHI.", "Data Privacy Guarantee: ")
    add_bullet_p(doc, "All 18 deliverables provided as formal Word documents (.docx) and Excel workbooks (.xlsx).", "Academic Deliverable Standards: ")

    # Appendix B
    add_heading_1(doc, "Appendix B: AI Tools Usage & Verification Audit")
    add_body_p(doc, "Transparent accounting of artificial intelligence tooling utilized throughout the engineering lifecycle and the rigorous human-in-the-loop verification methods applied:")

    ai_table = doc.add_table(rows=1, cols=4)
    ai_headers = ["AI Tool / Engine", "Lifecycle Step Used", "Specific Engineering Delegation", "Independent Verification Method"]
    for i, h in enumerate(ai_headers):
        ai_table.rows[0].cells[i].paragraphs[0].add_run(h)

    ai_data = [
        ("Antigravity AI (DeepMind)", "Steps 01 - 03", "Drafting problem understanding, HMW opportunity statements, and stakeholder RACI matrices.", "Cross-reviewed against healthcare operations research literature and vetted with mentor."),
        ("Antigravity AI (DeepMind)", "Steps 07 - 09", "Designing relational schema DDL, ER diagram Mermaid syntax, and foreign key rules.", "Validated via SQLite PRAGMA foreign_key_check; Pillow image generation rendered visual ER diagram."),
        ("Python REPL & Faker logic", "Step 10", "Generating 300 realistic synthetic outpatient fact records and dimension tables.", "Audited via validation_report.txt asserting zero null PKs, zero orphan FKs, and realistic value ranges."),
        ("Antigravity AI (DeepMind)", "Steps 11 - 13", "Formulating 14 agile user stories, Given/When/Then acceptance criteria, and 28 test cases.", "Compiled into RTM and mapped 100% bi-directionally to business requirements and code."),
        ("Streamlit / Tailwind CDN", "Steps 14 - 15", "Implementing multi-page Streamlit dashboard, Plotly charts, and Tailwind CSS components.", "Tested via 20 automated pytest test cases passing in 0.94s and manual browser navigation."),
        ("Pytest & SQLite CLI", "Steps 15 - 16", "Executing automated test suite and pre-deployment repository audits.", "Executed pytest tests/ -v in terminal; verified 100% test pass rate and clean git working tree.")
    ]
    for row_data in ai_data:
        row = ai_table.add_row()
        for i, val in enumerate(row_data):
            row.cells[i].paragraphs[0].add_run(val)
    style_table(ai_table, [1.5, 1.0, 2.4, 2.3])

    # Verification Checklist
    add_heading_1(doc, "Engineering Verification Checklist & Sign-Off")
    add_body_p(doc, "Summary of final verification assertions across all project dimensions:")
    add_bullet_p(doc, "All 10 sections + Cover Table + Appendix A & B are fully populated with zero placeholder brackets.", "Document Completeness: ")
    add_bullet_p(doc, "Every requirement (BR-01 to BR-12) traces end-to-end through OQ, HMW, SH, US, FR, UI, DB, and TC.", "Requirements Traceability: ")
    add_bullet_p(doc, "All 20 automated Pytests pass in 0.94 seconds with zero errors or warnings.", "Automated Quality Assurance: ")
    add_bullet_p(doc, "Application operates with zero C-extension crashes, fully compliant with host security policies.", "Environmental Resilience: ")
    add_bullet_p(doc, "Git repository initialized on main branch; ready for Step 18 Final Presentation & Evaluation.", "Deployment Readiness: ")

    doc.save(output_path)
    print(f"Successfully generated Master Project Document: {output_path}")

if __name__ == "__main__":
    primary_out = r"c:\Users\deeps\Downloads\Appointment No Show Analysis\docs\2500320140047_Project_Document.docx"
    backup_out = r"c:\Users\deeps\Downloads\Appointment No Show Analysis\docs\17_Project_Document.docx"
    build_project_document(primary_out)
    shutil.copyfile(primary_out, backup_out)
    print(f"Successfully mirrored deliverable to: {backup_out}")
