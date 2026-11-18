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
    for i, cell in enumerate(table.rows[0].cells):
        set_cell_background(cell, header_bg)
        set_cell_margins(cell, top=120, bottom=120, left=150, right=150)
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for r in p.runs:
                r.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)
                r.font.name = 'Calibri'
                r.font.size = Pt(10)
    for r_idx, row in enumerate(table.rows[1:], start=1):
        bg = stripe_bg if r_idx % 2 == 1 else "FFFFFF"
        for cell in row.cells:
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=90, bottom=90, left=140, right=140)
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.name = 'Calibri'
                    r.font.size = Pt(9.5)
                    r.font.color.rgb = RGBColor(51, 51, 51)
    for row in table.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = Inches(width)

def build_step3_document(output_path):
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
    sub_run = sub_p.add_run("DEPARTMENT OF MASTER OF COMPUTER APPLICATIONS (MCA)\nMCA 2nd Year — Data Analytics Specialization | Capstone Lifecycle")
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
        ("Domain & Organisation:", "Healthcare Analytics | ApexCare Outpatient & Corporate Health Network, Noida"),
        ("Deliverable Step:", "Step 03 — Business Needs Analysis & Requirements Specification (BR)")
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

    # Section 1 & 2
    add_heading_1("1 & 2. Business Requirements Specification (BR-01 through BR-12)")
    add_body_p(
        "Following business analysis best practices, this needs analysis transitions from human-centered HMW opportunity statements "
        "to formal Business Requirements. Each requirement defines WHAT the enterprise requires and the expected operational BENEFIT, "
        "strictly omitting implementation technologies (such as Streamlit, SQLite, or SQL queries)."
    )
    add_body_p(
        "Standard Business Requirement Template: \"The business needs to <capability> so that <benefit>.\""
    )

    br_table = doc.add_table(rows=13, cols=6)
    br_widths = [0.7, 2.3, 1.0, 0.7, 0.6, 1.2]
    br_headers = ["BR ID", "Business Requirement Statement", "Stakeholder", "Source HMW", "Priority", "How Measured (Testable Metric)"]
    for i, h in enumerate(br_headers):
        br_table.rows[0].cells[i].paragraphs[0].text = h

    br_data = [
        ("BR-01", "The business needs to track historical and ongoing appointment attendance rates across all clinic locations and time periods so that clinic administrators can detect systemic operational variances and benchmark facility performance.", "SH-01", "HMW-01", "Must", "Calculates monthly and weekly attendance % per clinic within ±0.1% accuracy."),
        ("BR-02", "The business needs to analyze appointment attendance as a function of booking lead time intervals so that scheduling coordinators can identify cutoff thresholds where attendance commitment significantly drops.", "SH-01, SH-03", "HMW-02", "Must", "Computes no-show percentages across 5 lead time buckets (0-2d, 3-7d, 8-14d, 15-30d, >30d)."),
        ("BR-03", "The business needs to evaluate specialty-specific attendance variations so that clinical leadership can uncover department bottlenecks and adjust doctor consultation rosters according to departmental show rates.", "SH-02", "HMW-03", "Must", "Outputs attendance breakdown across 6 specialties and validates significance via Chi-square test (p < 0.05)."),
        ("BR-04", "The business needs to segment appointment attendance by patient demographic cohorts (age brackets, gender, and residential proximity) so that clinical outreach can be customized for vulnerable or distant patient groups.", "SH-01, SH-03", "HMW-04", "Should", "Segments attendance by 5 age brackets, gender ratios, and 3 transit distance zones."),
        ("BR-05", "The business needs to measure the comparative attendance lift between patients who received automated digital reminders and those who did not so that communication channels can be empirically validated and optimized.", "SH-03", "HMW-05", "Must", "Calculates relative risk reduction %, odds ratio, and 2-sample proportion Z-test p-value."),
        ("BR-06", "The business needs to examine appointment arrival patterns across days of the week and consultation time slots so that facility supervisors can schedule adequate staff and prevent waiting lobby congestion during peak periods.", "SH-01", "HMW-06", "Should", "Builds a 2D matrix of attendance by Day (Mon-Sat) and Time Interval (Morning/Afternoon/Evening)."),
        ("BR-07", "The business needs to monitor consultation compliance among patients diagnosed with chronic conditions (such as Hypertension and Diabetes) so that clinical teams can prevent dangerous gaps in disease management.", "SH-02", "HMW-07", "Must", "Calculates follow-up attendance rate and compliance risk ratios for chronic vs non-chronic cohorts."),
        ("BR-08", "The business needs to distinguish between first-time appointment misses and habitual repeat no-show offenders so that front-desk staff can avoid expending high-touch phone calls on one-off accidental misses.", "SH-03", "HMW-08", "Should", "Classifies patient attendance history into 0, 1, and 2+ previous misses and computes repeat miss proportion."),
        ("BR-09", "The business needs to quantify the financial revenue loss and wasted doctor consultation hours resulting from unfilled appointments so that executive leadership can determine the monetary return on investment for attendance improvement initiatives.", "SH-01, SH-02", "HMW-09", "Must", "Accurately sums monthly uncollected consultation fees (in ₹) and wasted specialist hours across branches."),
        ("BR-10", "The business needs to calculate an advance risk tier (High, Medium, Low) for upcoming scheduled appointments so that front-desk staff can objectively prioritize high-risk patients for manual telephonic reconfirmation.", "SH-03", "HMW-10", "Must", "Classifies scheduled appointments into 3 risk tiers with analytical accuracy > 75%."),
        ("BR-11", "The business needs to correlate corporate employee shift timings with occupational health checkup completion so that corporate wellness partners can schedule on-site checkups during optimal work windows.", "SH-04", "HMW-11", "Could", "Reports completion rate across general vs evening shift corporate employee cohorts."),
        ("BR-12", "The business needs to simulate the operational impact of calibrated overbooking quotas based on historical no-show probabilities so that clinic managers can recover lost capacity without exceeding lobby seating thresholds.", "SH-01", "HMW-01, HMW-06", "Could", "Simulates recovered consultation slots against projected lobby wait time variance.")
    ]
    for r_idx, r_vals in enumerate(br_data, start=1):
        for c_idx, val in enumerate(r_vals):
            br_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(br_table, br_widths)

    # Section 3
    add_heading_1("3. MoSCoW Priority Breakdown")
    add_body_p(
        "The requirements set adheres strictly to standard agile project governance principles, "
        "maintaining roughly 58.3% (7 out of 12) core 'Must Have' capabilities to guarantee project feasibility within timeline constraints:"
    )
    add_bullet_p("Must Have (7 requirements, 58.3%): ", "BR-01, BR-02, BR-03, BR-05, BR-07, BR-09, BR-10. These represent mission-critical operational, statistical, financial, and predictive capabilities.")
    add_bullet_p("Should Have (3 requirements, 25.0%): ", "BR-04, BR-06, BR-08. High-value demographic, temporal heat-mapping, and repeat-offender behavioral segmentation.")
    add_bullet_p("Could Have (2 requirements, 16.7%): ", "BR-11, BR-12. Specialized corporate shift correlation and overbooking capacity simulation.")
    add_bullet_p("Won't Have (Out-of-Scope / Deferred): ", "Live telephonic AI auto-dialer robots, direct biometric patient check-in kiosks, and bi-directional hospital EHR HL7/FHIR write-back.")

    # Section 4
    add_heading_1("4. Data Needs per Business Requirement")
    add_body_p(
        "To guarantee that every requirement can be completely satisfied by the analytical data store, "
        "the table below documents the specific data attributes and features demanded by each BR:"
    )

    d_table = doc.add_table(rows=13, cols=4)
    d_widths = [0.8, 1.8, 2.5, 1.4]
    d_headers = ["BR ID", "Requirement Summary", "Required Data Attributes / Columns", "Data Type & Schema Role"]
    for i, h in enumerate(d_headers):
        d_table.rows[0].cells[i].paragraphs[0].text = h

    d_data = [
        ("BR-01", "Facility Attendance Tracking", "AppointmentID, AppointmentDate, ClinicBranch, Status", "Discrete IDs, Timestamps, Categorical Branch, Binary Status"),
        ("BR-02", "Lead Time Impact Analysis", "BookingDate, AppointmentDate, LeadTimeDays, Status", "Timestamps, Continuous/Binned Integer Days, Binary Status"),
        ("BR-03", "Specialty Bottleneck Evaluation", "Specialty, DepartmentID, ConsultationFee, Status", "Categorical Medical Specialties, Foreign Keys, Binary Status"),
        ("BR-04", "Demographic & Proximity Cohorts", "PatientAge, AgeGroup, Gender, Neighborhood, DistanceKm, Status", "Continuous Age/Distance, Categorical Bins/Gender, Status"),
        ("BR-05", "Reminder Lift & Efficacy", "SMSReminderSent, ReminderCount, ReminderDate, Status", "Boolean Flag (0/1), Integer Count, Timestamp, Status"),
        ("BR-06", "Day & Time Slot Heatmapping", "AppointmentDate, AppointmentTime, DayOfWeek, TimeSlotBucket, Status", "Date/Time stamps, Derived Day (Mon-Sat), Categorical Slot"),
        ("BR-07", "Chronic Condition Compliance", "Hypertension, Diabetes, ComorbidityCount, Status", "Boolean Diagnostic Flags (0/1), Ordinal Count, Status"),
        ("BR-08", "Repeat vs First-time Offenders", "PatientID, TotalPreviousVisits, PreviousNoShows, RepeatOffenderFlag", "Entity Identifier, Cumulative Counts, Categorical Flag"),
        ("BR-09", "Financial & Capacity Loss", "Status, ConsultationFee, SlotDurationMinutes, DoctorID", "Binary Status, Decimal Currency (INR), Integer Minutes"),
        ("BR-10", "Predictive Risk Stratification", "LeadTimeDays, SMSReminderSent, AgeGroup, PreviousNoShows, Specialty, RiskScore, RiskTier", "Multi-factor Feature Vector, Continuous Score (0-100), Ordinal Tier (H/M/L)"),
        ("BR-11", "Corporate Shift Adherence", "CorporatePartnerID, EmployeeShift (General/Evening), Status", "Corporate Foreign Key, Categorical Shift, Status"),
        ("BR-12", "Overbooking Simulation", "ClinicBranch, SlotCapacity, PredictedNoShowRate, BufferMultiplier", "Facility Identifier, Integer Capacity, Float Probability")
    ]
    for r_idx, r_vals in enumerate(d_data, start=1):
        for c_idx, val in enumerate(r_vals):
            d_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(d_table, d_widths)

    # Section 5
    add_heading_1("5. Gap Analysis (HMW vs. BR Bidirectional Check)")
    add_body_p(
        "A rigorous bidirectional gap analysis ensures complete architectural integrity: zero orphaned HMW statements "
        "and zero unsourced business requirements."
    )

    gap_table = doc.add_table(rows=3, cols=3)
    gap_widths = [1.8, 1.8, 2.9]
    gap_headers = ["Audit Dimension", "Finding", "Resolution & Audit Notes"]
    for i, h in enumerate(gap_headers):
        gap_table.rows[0].cells[i].paragraphs[0].text = h

    gap_data = [
        ("HMW Statements with No BR (Orphaned Opportunities)", "0 Found (Zero Gaps)", "Every HMW from HMW-01 to HMW-11 maps directly to at least one corresponding business requirement (e.g., HMW-10 -> BR-10, HMW-09 -> BR-09)."),
        ("Business Requirements with No Source HMW (Scope Creep)", "0 Found (Zero Scope Creep)", "All 12 Business Requirements (BR-01 to BR-12) explicitly cite their originating HMW statements. No extraneous or unsolicited features exist.")
    ]
    for r_idx, r_vals in enumerate(gap_data, start=1):
        for c_idx, val in enumerate(r_vals):
            gap_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(gap_table, gap_widths)

    add_heading_2("Testability Audit:")
    add_body_p(
        "Every requirement defined in Section 1 specifies a clear, quantifiable measurement metric "
        "(e.g., statistical p-values, accuracy thresholds, percentage bins, currency summation formulas). "
        "This ensures that down the lifecycle in Step 13 (Test Cases) and Step 15 (Pytest Build), automated assertions can "
        "validate 100% of these business requirements without ambiguity."
    )

    doc.add_paragraph().paragraph_format.space_after = Pt(16)
    
    doc.save(output_path)
    print(f"Step 3 Document successfully created at: {output_path}")

if __name__ == "__main__":
    out_dir = r"c:\Users\deeps\Downloads\Appointment No Show Analysis\docs"
    file_name = "03_Needs_Analysis.docx"
    full_path = os.path.join(out_dir, file_name)
    build_step3_document(full_path)
