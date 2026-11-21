import os
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
    # Header row
    for i, cell in enumerate(table.rows[0].cells):
        set_cell_background(cell, header_bg)
        set_cell_margins(cell, top=120, bottom=120, left=140, right=140)
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for r in p.runs:
                r.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)
                r.font.name = 'Calibri'
                r.font.size = Pt(9.5)
    # Data rows
    for r_idx, row in enumerate(table.rows[1:], start=1):
        bg = stripe_bg if r_idx % 2 == 1 else "FFFFFF"
        for cell in row.cells:
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=80, bottom=80, left=130, right=130)
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.name = 'Calibri'
                    r.font.size = Pt(9.0)
                    r.font.color.rgb = RGBColor(51, 51, 51)
    # Widths
    for row in table.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = Inches(width)

def build_step5_brd(output_path):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc = Document()
    
    # 1-inch margins
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
    sub_run = sub_p.add_run("DEPARTMENT OF MASTER OF COMPUTER APPLICATIONS (MCA)\nMCA 2nd Year — Data Analytics Specialization | Capstone Project Deliverable")
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
        ("Domain & Organisation:", "Healthcare Analytics | ApexCare Outpatient & Corporate Health Network, Noida"),
        ("Deliverable Step:", "Step 05 — Business Requirements Document (BRD v1.0 Baseline)")
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

    # 1. Document Control
    add_heading_1("1. Document Control & Revision History")
    dc_table = doc.add_table(rows=5, cols=5)
    dc_widths = [0.8, 1.2, 1.3, 1.5, 1.7]
    dc_headers = ["Version", "Release Date", "Author / Role", "Reviewer / Approver", "Description of Changes"]
    for i, h in enumerate(dc_headers):
        dc_table.rows[0].cells[i].paragraphs[0].text = h
    dc_rows = [
        ("v0.1", "2026-11-18", "Deepshikha (Learner/BA)", "Dinesh Sereno (Faculty)", "Initial Draft & Scope Definition"),
        ("v0.2", "2026-11-18", "Deepshikha (Learner/BA)", "Smt. Kavita Sharma (Ops Lead)", "Stakeholder Pain Points & HMW Traceability"),
        ("v0.3", "2026-11-18", "Deepshikha (Learner/BA)", "Dr. A. K. Banerjee (CMO)", "Clinical Compliance & MoSCoW Requirements"),
        ("v1.0", "2026-11-18", "Deepshikha (Learner/BA)", "Dinesh Sereno (Faculty)", "Final Formal Baseline for Architectural Sign-off")
    ]
    for r_i, r_data in enumerate(dc_rows, start=1):
        for c_i, val in enumerate(r_data):
            dc_table.rows[r_i].cells[c_i].paragraphs[0].text = val
    style_table(dc_table, dc_widths)

    # 2. Executive Summary
    add_heading_1("2. Executive Summary")
    add_body_p(
        "This Business Requirements Document (BRD) establishes the business architecture and functional specifications for the "
        "Appointment No-Show Analysis & Optimization System. Patient non-attendance represents one of the most persistent and costly "
        "operational dysfunctions in healthcare systems worldwide. Within the urban healthcare corridor of Noida, outpatient medical facilities "
        "and occupational health clinics face an average no-show rate of 24.2%. This leads to hundreds of lost specialist consultation hours, "
        "worsening patient backlog queues, and annual direct consultation revenue leakage exceeding ₹14.5 Lakhs per clinical branch."
    )
    add_body_p(
        "To systematically address this problem, this system delivers an interactive, data-driven analytical web platform. "
        "By synthesizing historical appointment schedules, patient demographics, clinical comorbidities, booking lead times, and digital reminder dispatches, "
        "the application equips clinic administrators, medical superintendents, and front-desk personnel with diagnostic visibility, "
        "financial accounting, statistical hypothesis testing, and proactive risk-tier triage to optimize outpatient operations."
    )

    # 3. Business Problem and Opportunity
    add_heading_1("3. Business Problem & Opportunity")
    add_body_p(
        "The Root Cause: Outpatient scheduling at ApexCare Health Network operates in transactional isolation. Appointments are scheduled on static "
        "calendar slots without any predictive intelligence regarding patient arrival probability. As a consequence, clinics suffer four cascading failures:",
        bold_prefix="Current Operational Bottlenecks: "
    )
    add_bullet_p("Physician Resource Waste: ", "Specialist doctors in Cardiology, Orthopedics, and Psychiatry sit idle during unannounced 20-30 minute consultation gaps.")
    add_bullet_p("Artificially Extended Patient Queues: ", "Patients seeking immediate consultations face multi-day delays because slots appear booked on paper, only to be abandoned on the consultation day.")
    add_bullet_p("Suboptimal Double-Booking Friction: ", "Crude, arbitrary overbooking heuristics cause waiting lobby overcrowding and patient hostility when all patients appear.")
    add_bullet_p("Preventive Health Care Gaps: ", "Corporate IT employees in partnered firms (such as NoidaTech Solutions) skip routine wellness checkups, causing unmonitored chronic disease progression.")
    add_body_p(
        "The Strategic Opportunity: By centralizing scheduling data into an analytical warehouse, the business can empirically identify attendance drivers "
        "(such as lead time decay and SMS reminder effectiveness), quantify exact department losses, and classify upcoming appointments into High, Medium, "
        "and Low risk tiers. This enables targeted pre-visit telephone reconfirmation, recapturing up to 35% of lost slot capacity."
    )

    # 4. Objectives and Success Metrics
    add_heading_1("4. Business Objectives & Key Performance Indicators (KPIs)")
    add_body_p("Project success is anchored to five core SMART objectives and measurable organizational KPIs:")

    kpi_table = doc.add_table(rows=6, cols=4)
    kpi_widths = [1.2, 1.8, 1.8, 1.7]
    kpi_headers = ["Strategic Objective", "Primary Business Metric (KPI)", "Current Baseline", "Target Benchmark Post-Adoption"]
    for i, h in enumerate(kpi_headers):
        kpi_table.rows[0].cells[i].paragraphs[0].text = h
    kpi_rows = [
        ("OBJ-01 (Attendance Transparency)", "Network-wide No-Show Rate Tracking", "24.2% Unmonitored", "< 16.0% with targeted triage"),
        ("OBJ-02 (Diagnostic Rigor)", "Statistical Driver Significance", "Anecdotal assumptions", "Validated p < 0.05 for all core drivers"),
        ("OBJ-03 (Financial Transparency)", "Wasted Doctor-Hour & Revenue Audit", "0% Visibility into Loss", "100% Monthly Accounting in ₹ and Hours"),
        ("OBJ-04 (Predictive Outreach)", "High-Risk Slot Triage Accuracy", "Alphabetical Blind Calling", "> 75% Risk Classification Precision"),
        ("OBJ-05 (Production System QA)", "Test Suite Coverage & Performance", "Untested manual scripts", ">= 85% Pytest Coverage, < 2.5s Dashboard Load")
    ]
    for r_i, r_data in enumerate(kpi_rows, start=1):
        for c_i, val in enumerate(r_data):
            kpi_table.rows[r_i].cells[c_i].paragraphs[0].text = val
    style_table(kpi_table, kpi_widths)

    # 5. Stakeholders RACI
    add_heading_1("5. Stakeholders & RACI Governance Matrix")
    add_body_p(
        "The RACI matrix governs organizational decision rights across the five project phases: "
        "R = Responsible (Completes task), A = Accountable (Final approval), C = Consulted (Provides input), I = Informed (Kept updated)."
    )

    raci_table = doc.add_table(rows=6, cols=6)
    raci_widths = [2.0, 0.9, 0.9, 0.9, 0.9, 0.9]
    raci_headers = ["Project Lifecycle Phase / Functional Area", "SH-01 (Ops)", "SH-02 (CMO)", "SH-03 (Desk)", "SH-04 (HR)", "SH-05 (Analyst)"]
    for i, h in enumerate(raci_headers):
        raci_table.rows[0].cells[i].paragraphs[0].text = h
    raci_rows = [
        ("Requirements & Scope Formulation", "A", "C", "C", "C", "R"),
        ("Data Modeling & Synthetic ETL Pipeline", "C", "I", "I", "I", "R/A"),
        ("Exploratory & Hypothesis Statistical Engine", "C", "A", "I", "I", "R"),
        ("Predictive Risk Scoring & Triage UI Design", "A", "C", "C", "I", "R"),
        ("System Testing, Cloud Deployment & Sign-Off", "A", "A", "I", "I", "R")
    ]
    for r_i, r_data in enumerate(raci_rows, start=1):
        for c_i, val in enumerate(r_data):
            raci_table.rows[r_i].cells[c_i].paragraphs[0].text = val
    style_table(raci_table, raci_widths)

    # 6. Scope
    add_heading_1("6. Project Scope")
    add_heading_2("In-Scope Functional Boundaries:")
    add_bullet_p("Historical Data Ingestion & ETL: ", "Processing 10,000+ synthetic appointment records with demographic, temporal, and clinical dimensions.")
    add_bullet_p("Multi-Factor Attendance Analytics: ", "Descriptive and bivariate analysis of booking lead times, clinical specialties, age brackets, and reminder impacts.")
    add_bullet_p("Formal Hypothesis Testing: ", "Execution of two-sample proportion Z-tests (reminder impact) and Chi-square tests of independence (specialty distribution).")
    add_bullet_p("Financial & Doctor Idle Capacity Engine: ", "Automated calculation of monthly uncollected consultation fees and lost physician hours.")
    add_bullet_p("Predictive Risk Scoring Algorithm: ", "Composite multi-feature risk scoring classifying appointments into High, Medium, and Low risk tiers.")
    add_bullet_p("Interactive Streamlit Web Dashboard: ", "Interactive Plotly visual charts, parameter filters, and modern responsive Tailwind UI cues.")

    add_heading_2("Out-of-Scope (Excluded from Current Release):")
    add_bullet_p("EHR/EMR Live HL7/FHIR Write-Back: ", "Direct modification of hospital electronic medical record operational transactional databases.")
    add_bullet_p("Autonomous AI Tele-Calling Agents: ", "Robotic automated conversational voice bots placing live telephone calls to patients.")
    add_bullet_p("Hardware & Native Mobile Apps: ", "Biometric fingerprint check-in terminals and native iOS/Android mobile applications.")

    # 7. Business Requirements
    add_heading_1("7. Business Requirements Specification (BR-01 through BR-12)")
    add_body_p(
        "Each requirement is uniquely identified, fully described with business rationale, mapped to its source HMW opportunity, "
        "prioritized using MoSCoW, and bounded by testable acceptance criteria:"
    )

    br_defs = [
        ("BR-01", "Network Attendance & Facility Tracking", "The business needs to track historical and ongoing appointment attendance rates across all clinic locations and time periods so that clinic administrators can detect systemic operational variances and benchmark facility performance.", "Enables clinic managers to isolate underperforming clinic branches and observe seasonal attendance trends.", "Must", "HMW-01", "Calculates weekly and monthly no-show percentages across all 4 branches within ±0.1% accuracy."),
        ("BR-02", "Lead-Time Impact Analysis", "The business needs to analyze appointment attendance as a function of booking lead time intervals so that scheduling coordinators can identify cutoff thresholds where attendance commitment significantly drops.", "Reveals the exponential decay in attendance as lead time increases from same-day to 30+ days.", "Must", "HMW-02", "Segments records into 5 lead-time intervals (0-2d, 3-7d, 8-14d, 15-30d, >30d) and computes no-show rates for each."),
        ("BR-03", "Specialty Bottleneck Evaluation", "The business needs to evaluate specialty-specific attendance variations so that clinical leadership can uncover department bottlenecks and adjust doctor consultation rosters according to departmental show rates.", "Specialties with high no-show rates (e.g. Psychiatry, Dermatology) suffer disproportionate revenue loss.", "Must", "HMW-03", "Outputs attendance breakdown across 6 specialties and validates statistical significance via Chi-square test (p < 0.05)."),
        ("BR-04", "Demographic & Proximity Cohorts", "The business needs to segment appointment attendance by patient demographic cohorts (age brackets, gender, and residential proximity) so that clinical outreach can be customized for vulnerable or distant patient groups.", "Helps tailor clinical communications to younger cohorts or patients residing beyond 15 km.", "Should", "HMW-04", "Computes attendance differentials across 5 standard age brackets, gender ratios, and 3 transit distance zones."),
        ("BR-05", "Digital Reminder Efficacy & Lift", "The business needs to measure the comparative attendance lift between patients who received automated digital reminders and those who did not so that communication channels can be empirically validated and optimized.", "Justifies communication channel expenditure and identifies patient groups needing multi-channel alerts.", "Must", "HMW-05", "Outputs relative risk reduction %, odds ratio, and two-sample proportion Z-test significance (p < 0.05)."),
        ("BR-06", "Temporal Day & Slot Heatmapping", "The business needs to examine appointment arrival patterns across days of the week and consultation time slots so that facility supervisors can schedule adequate staff and prevent waiting lobby congestion during peak periods.", "Enables optimal shift scheduling matching historical patient arrival peaks and valleys.", "Should", "HMW-06", "Generates a 2D matrix of attendance cross-tabulating weekdays (Mon-Sat) with time buckets (Morning, Afternoon, Evening)."),
        ("BR-07", "Chronic Condition Compliance", "The business needs to monitor consultation compliance among patients diagnosed with chronic conditions (such as Hypertension and Diabetes) so that clinical teams can prevent dangerous gaps in disease management.", "Prevents acute health emergencies and hospital readmissions among chronic lifestyle disease patients.", "Must", "HMW-07", "Calculates follow-up attendance rate and compliance risk ratios for chronic vs non-chronic cohorts."),
        ("BR-08", "Repeat vs First-time Offenders", "The business needs to distinguish between first-time appointment misses and habitual repeat no-show offenders so that front-desk staff can avoid expending high-touch phone calls on one-off accidental misses.", "Prioritizes high-touch staff outreach exclusively on habitual repeat offenders rather than one-time misses.", "Should", "HMW-08", "Categorizes patient history into 0, 1, and 2+ previous misses and computes the proportion of total no-shows caused by repeat offenders."),
        ("BR-09", "Financial & Capacity Loss Accounting", "The business needs to quantify the financial revenue loss and wasted doctor consultation hours resulting from unfilled appointments so that executive leadership can determine the monetary return on investment for attendance improvement initiatives.", "Gives hospital leadership an undeniable quantitative justification for operational intervention budgets.", "Must", "HMW-09", "Accurately sums monthly uncollected consultation fees (in ₹) and total idle physician hours across clinics."),
        ("BR-10", "Predictive Risk Tiering & Triage", "The business needs to calculate an advance risk tier (High, Medium, Low) for upcoming scheduled appointments so that front-desk staff can objectively prioritize high-risk patients for manual telephonic reconfirmation.", "Converts passive historical data into immediate proactive front-desk telephone intervention.", "Must", "HMW-10", "Classifies upcoming appointments into High (>60%), Medium (30-60%), and Low (<30%) risk tiers with validation accuracy > 75%."),
        ("BR-11", "Corporate Shift Adherence Analysis", "The business needs to correlate corporate employee shift timings with occupational health checkup completion so that corporate wellness partners can schedule on-site checkups during optimal work windows.", "Maximizes preventive checkup completion for corporate IT partners by aligning clinic hours with software sprint cycles.", "Could", "HMW-11", "Reports preventive checkup attendance rate segmented by general vs evening corporate shift schedules."),
        ("BR-12", "Calibrated Overbooking Simulation", "The business needs to simulate the operational impact of calibrated overbooking quotas based on historical no-show probabilities so that clinic managers can recover lost capacity without exceeding lobby seating thresholds.", "Allows clinic managers to scientifically recover slot vacancies without causing lobby wait-time spikes.", "Could", "HMW-12", "Simulates recovered consultation slots against projected lobby wait time variance across overbooking multipliers.")
    ]

    for req in br_defs:
        r_id, r_name, r_desc, r_rat, r_pri, r_src, r_acc = req
        add_heading_2(f"{r_id}: {r_name} (Priority: {r_pri})")
        add_body_p(r_desc, bold_prefix="Requirement Statement: ")
        add_body_p(r_rat, bold_prefix="Business Rationale: ")
        add_body_p(f"Source: {r_src} | Priority: {r_pri}", bold_prefix="Traceability & Priority: ")
        add_body_p(r_acc, bold_prefix="Acceptance Criteria: ")

    # 8. Business Rules
    add_heading_1("8. Business Rules")
    add_body_p("The application logic strictly enforces the following core enterprise business rules:")

    b_rules = [
        ("B-RULE-01", "Appointment Outcome Definition", "An appointment record must resolve to exactly one terminal state: 'Show' (patient attended consultation) or 'No-Show' (patient failed to attend without prior cancellation). Cancellations made >= 24 hours prior are excluded from no-show penalties."),
        ("B-RULE-02", "Lead-Time Calculation Logic", "Booking Lead Time (in days) is calculated as: LeadTimeDays = max(0, AppointmentDate - BookingDate). Negative values are prohibited by data validation rules."),
        ("B-RULE-03", "Physician Idle Time Valuation", "Physician idle cost is computed as: IdleCost = ConsultationFee * IsNoShow. Capacity loss is computed as: IdleHours = SlotDurationMinutes / 60.0."),
        ("B-RULE-04", "Repeat Offender Behavioral Thresholds", "Patients with zero prior misses are classified as 'First-Time / Clean History'; exactly 1 prior miss as 'Occasional Miss'; 2 or more prior misses as 'Habitual Repeat Offender'."),
        ("B-RULE-05", "Risk Tier Classification Boundaries", "Appointments with computed risk score >= 60.0 are classified as 'High Risk'; 30.0 to 59.9 as 'Medium Risk'; < 30.0 as 'Low Risk'."),
        ("B-RULE-06", "Reminder Attribution Rule", "A reminder is credited as an active intervention only if SMSReminderSent == 1 and dispatched at least 24 hours prior to the scheduled appointment timestamp."),
        ("B-RULE-07", "Data Privacy & Governance Compliance", "All records ingested by the analytics engine must be 100% synthetic, anonymized, and devoid of real Personally Identifiable Information (PII), adhering strictly to HIPAA and NABH ethical data standards."),
        ("B-RULE-08", "Statistical Significance Standard", "All comparative hypotheses (Z-tests, Chi-square tests) must achieve a p-value < 0.05 (alpha = 0.05) to be reported as statistically significant business conclusions.")
    ]

    br_table = doc.add_table(rows=9, cols=3)
    br_widths = [1.2, 2.0, 3.3]
    br_headers = ["Rule ID", "Business Rule Title", "Operational Enforcement Policy"]
    for i, h in enumerate(br_headers):
        br_table.rows[0].cells[i].paragraphs[0].text = h
    for r_i, r_data in enumerate(b_rules, start=1):
        for c_i, val in enumerate(r_data):
            br_table.rows[r_i].cells[c_i].paragraphs[0].text = val
    style_table(br_table, br_widths)

    # 9. Data Requirements
    add_heading_1("9. Data Requirements")
    add_body_p(
        "Data Architecture & Sources: The system ingests synthetic outpatient consultation records modeled after multi-specialty clinical operations. "
        "Primary schema entities comprise Patients, Appointments, Clinical Departments/Specialties, Physicians, and Reminder Logs."
    )
    add_body_p(
        "Refresh & Ingestion Frequency: Analytical data is stored in a structured SQLite relational database file (`noshow_analytics.db`). "
        "The system supports instant in-memory loading via Streamlit cache decorators (`@st.cache_data`), ensuring sub-2.5 second refresh upon filter manipulation."
    )

    # 10. Assumptions, Constraints, Dependencies
    add_heading_1("10. Assumptions, Constraints & Dependencies")
    add_bullet_p("Assumption 1: ", "Synthetic datasets represent genuine clinical distributions (age 18-85, lead time 0-45 days, baseline no-show rate 20-25%).")
    add_bullet_p("Assumption 2: ", "Specialty consultation durations follow standard slots: 20 minutes for general medicine and 30 minutes for specialized branches.")
    add_bullet_p("Constraint 1: ", "The public deployment environment is bounded by the 1 GB memory limit of Streamlit Community Cloud free tier.")
    add_bullet_p("Constraint 2: ", "Data storage is restricted to an embedded, portable SQLite relational database to ensure 100% reproducibility across grading machines.")
    add_bullet_p("Dependency 1: ", "Requires Python 3.10+, Streamlit 1.30+, Pandas 2.0+, Plotly 5.18+, SciPy 1.11+, and Pytest 7.4+.")

    # 11. Risks and Mitigation
    add_heading_1("11. Risk Management & Mitigation Register")
    risk_table = doc.add_table(rows=5, cols=5)
    risk_widths = [0.8, 2.2, 0.8, 0.8, 1.9]
    risk_headers = ["Risk ID", "Risk Description", "Likelihood", "Impact", "Mitigation Strategy"]
    for i, h in enumerate(risk_headers):
        risk_table.rows[0].cells[i].paragraphs[0].text = h
    risk_rows = [
        ("RSK-01", "Synthetic data lacks realistic statistical covariance, weakening analytical conclusions.", "Medium", "High", "Calibrate data generation algorithms against published clinical informatics literature (e.g., Brazilian Medical No-Show corpus)."),
        ("RSK-02", "Streamlit Cloud memory exhaustion (>1GB) during large dataset transformations.", "Low", "High", "Implement downcasted numeric datatypes (int16, category) and optimize cache pipelines using `@st.cache_data`."),
        ("RSK-03", "Risk scoring heuristics produce unacceptable false positive rates on low-risk patients.", "Medium", "Medium", "Tune probability classification thresholds and establish a balanced 'Medium Risk' buffer tier."),
        ("RSK-04", "Academic timeline overrun during build and QA phases.", "Low", "High", "Enforce strict MoSCoW prioritization, locking 58.3% Must-Haves and deferring Could-Haves if necessary.")
    ]
    for r_i, r_data in enumerate(risk_rows, start=1):
        for c_i, val in enumerate(r_data):
            risk_table.rows[r_i].cells[c_i].paragraphs[0].text = val
    style_table(risk_table, risk_widths)

    # 12. Glossary
    add_heading_1("12. Project Glossary")
    gloss_data = [
        ("No-Show", "An instance where a patient books an outpatient consultation appointment but fails to arrive without prior notification or cancellation."),
        ("Booking Lead Time", "The elapsed calendar interval in days between the booking timestamp and the scheduled appointment timestamp."),
        ("Show Rate", "The proportion of scheduled appointments successfully attended by patients, expressed as a percentage of total bookings."),
        ("Comorbidity", "The simultaneous presence of two or more chronic medical conditions in a single patient (e.g., Hypertension and Diabetes)."),
        ("Two-Sample Proportion Z-Test", "A parametric statistical hypothesis test used to determine whether the difference between two independent proportions is statistically significant."),
        ("Odds Ratio (OR)", "A measure of association between an exposure (e.g., receiving an SMS reminder) and an outcome (e.g., attending an appointment)."),
        ("MoSCoW", "A prioritization framework categorizing requirements into Must Have, Should Have, Could Have, and Won't Have."),
        ("RTM", "Requirements Traceability Matrix: a master tabular cross-reference mapping business needs to code and test cases.")
    ]
    gl_table = doc.add_table(rows=9, cols=2)
    gl_widths = [1.8, 4.7]
    gl_table.rows[0].cells[0].paragraphs[0].text = "Term / Acronym"
    gl_table.rows[0].cells[1].paragraphs[0].text = "Operational Definition"
    for r_i, (t, d) in enumerate(gloss_data, start=1):
        gl_table.rows[r_i].cells[0].paragraphs[0].text = t
        gl_table.rows[r_i].cells[1].paragraphs[0].text = d
    style_table(gl_table, gl_widths)

    # 13. Sign-Off
    add_heading_1("13. Formal Project Baseline Sign-Off")
    add_body_p(
        "By signing below, the project stakeholders endorse this Business Requirements Document (BRD v1.0) "
        "as the authoritative functional and operational baseline for downstream technical design and product development:"
    )

    so_table = doc.add_table(rows=5, cols=4)
    so_widths = [1.8, 1.8, 1.5, 1.4]
    so_headers = ["Stakeholder Name", "Role / Title", "Signature", "Date of Sign-off"]
    for i, h in enumerate(so_headers):
        so_table.rows[0].cells[i].paragraphs[0].text = h
    so_rows = [
        ("Deepshikha", "Student Data Analyst (MCA 2nd Yr)", "Deepshikha", "2026-11-18"),
        ("Dinesh Sereno", "Faculty Project Mentor (Dept of MCA)", "Approved (Dinesh Sereno)", "2026-11-18"),
        ("Smt. Kavita Sharma", "Clinic Operations Manager (SH-01)", "Approved (Kavita Sharma)", "2026-11-18"),
        ("Dr. A. K. Banerjee", "Chief Medical Officer (SH-02)", "Approved (A.K. Banerjee)", "2026-11-18")
    ]
    for r_i, r_data in enumerate(so_rows, start=1):
        for c_i, val in enumerate(r_data):
            so_table.rows[r_i].cells[c_i].paragraphs[0].text = val
    style_table(so_table, so_widths)

    doc.add_paragraph().paragraph_format.space_after = Pt(16)
    
    doc.save(output_path)
    print(f"Step 5 BRD Document successfully created at: {output_path}")

if __name__ == "__main__":
    out_dir = r"c:\Users\deeps\Downloads\Appointment No Show Analysis\docs"
    file_name = "05_BRD.docx"
    full_path = os.path.join(out_dir, file_name)
    build_step5_brd(full_path)
