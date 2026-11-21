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
    # Format header
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
    # Format data rows
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
    # Set widths
    for row in table.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = Inches(width)

def build_step2_document(output_path):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc = Document()
    
    # Page setup
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
    h_table = doc.add_table(rows=6, cols=2)
    h_widths = [2.2, 4.3]
    h_data = [
        ("Project ID & Batch:", "P101 | ABES_OND26 MCA 2ND YEAR_DATA ANALYTICS_BATCH 01"),
        ("Project Title:", "Appointment No-Show Analysis & Optimization System"),
        ("Student Details:", "Deepshikha | Roll No: 2500320140047 | MCA 2nd Year (DA)"),
        ("Faculty Mentor:", "Dinesh Sereno, Assistant Professor, Dept. of MCA"),
        ("Domain & Organisation:", "Healthcare Analytics | ApexCare Outpatient & Corporate Health Network, Noida"),
        ("Deliverable Step:", "Step 02 — How Might We (HMW) Design-Thinking Opportunity Framework")
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

    # Section 1
    add_heading_1("1. Stakeholder Pain Point Synthesis")
    add_body_p(
        "Using the foundational operational analysis established in Step 1, the design-thinking facilitation process "
        "synthesizes the granular frustrations experienced by each distinct stakeholder group across clinical, operational, and technical dimensions:"
    )

    add_heading_2("SH-01: Clinic Operations Manager (Smt. Kavita Sharma)")
    add_bullet_p("Unpredictable Facility Throughput: ", "Lacks advance visibility into daily attendance variance, forcing administrative staff to reactively manage unpredictable patient surges or vacant waiting rooms.")
    add_bullet_p("Blunt Overbooking Penalties: ", "Attempts to compensate for no-shows through arbitrary 20% flat overbooking, leading to severe lobby congestion, patient friction, and staff stress on days when attendance spikes.")
    add_bullet_p("Unquantified Departmental Losses: ", "Unable to present clear monthly financial metrics demonstrating exact revenue loss and clinical overhead wasted per facility to executive management.")

    add_heading_2("SH-02: Chief Medical Officer / Medical Superintendent (Dr. A. K. Banerjee)")
    add_bullet_p("Physician Idle Time & Dissatisfaction: ", "Specialist doctors frequently sit idle during 20–30 minute unannounced gaps, leading to physician frustration and diminished clinical productivity.")
    add_bullet_p("Compromised Chronic Care Continuity: ", "Patients suffering from chronic comorbidities (Hypertension, Diabetes) frequently skip routine monitoring visits unnoticed, escalating emergency complications.")
    add_bullet_p("Uneven Specialty Utilization: ", "Lacks granular analytics to discern why certain clinical specialties experience 30%+ no-shows while others remain oversubscribed.")

    add_heading_2("SH-03: Front Desk & Outpatient Scheduling Staff (Mr. Rohit Verma)")
    add_bullet_p("Overwhelming & Untargeted Call Rosters: ", "Staff have limited time (1-2 hours daily) for telephonic appointment reconfirmations, currently dialing patients alphabetically rather than by no-show risk.")
    add_bullet_p("Ineffective Generic Reminders: ", "Standard broadcast SMS notifications are sent indiscriminately, yielding minimal impact on long lead-time bookings or habitual no-show patients.")
    add_bullet_p("High Lobby Friction: ", "Faces patient anger when overbooked slots result in 45+ minute lobby delays, causing operational burnout.")

    add_heading_2("SH-04: Corporate Partner HR & Wellness Lead (Ms. Neha Gupta - NoidaTech IT)")
    add_bullet_p("Poor Preventive Checkup Adherence: ", "Corporate employees in high-stress software engineering roles repeatedly miss booked annual wellness checks due to sudden project meetings.")
    add_bullet_p("Suboptimal Slot Alignment: ", "Wellness appointment slots do not correspond with employee work schedules (e.g., late evenings or weekends), depressing attendance.")

    add_heading_2("SH-05: Data Analyst & Academic Evaluator (Deepshikha & Dinesh Sereno)")
    add_bullet_p("Risk of Disconnected Feature Engineering: ", "Ensuring that complex statistical transformations and risk calculations map directly back to actionable clinic workflows.")
    add_bullet_p("Need for Mathematical & Hypothesis Rigor: ", "Must prove statistical significance beyond descriptive correlation using rigorous two-sample hypothesis testing and odds ratios.")

    # Section 2 & 3
    add_heading_1("2 & 3. How Might We (HMW) Statements & Master Traceability Table")
    add_body_p(
        "Design-thinking transforms each validated pain point into an open, actionable 'How Might We' question. "
        "These statements frame opportunities without presupposing specific UI controls or premature technical solutions:"
    )

    hmw_table = doc.add_table(rows=12, cols=5)
    hmw_widths = [0.8, 2.5, 1.2, 1.4, 0.6]
    hmw_headers = ["HMW ID", "How Might We ...?", "Stakeholder", "Stakeholder Pain Point", "Linked OQ"]
    for i, h in enumerate(hmw_headers):
        hmw_table.rows[0].cells[i].paragraphs[0].text = h

    hmw_data = [
        ("HMW-01", "How might we empower clinic administrators to detect systemic attendance trends and branch-level disparities over time before scheduling rosters are locked?", "SH-01: Operations Manager", "Unpredictable clinic throughput and vacant waiting areas", "OQ-01"),
        ("HMW-02", "How might we enable scheduling staff to identify appointments whose booking lead time introduces a high decay in patient attendance commitment?", "SH-01, SH-03: Front Desk", "Lead-time forgetfulness causing chronic no-shows", "OQ-02"),
        ("HMW-03", "How might we help medical leadership uncover specialty-specific bottlenecks where patient hesitation or consultation barriers drive non-attendance?", "SH-02: Chief Medical Officer", "Uneven specialty utilization and doctor idle time", "OQ-03"),
        ("HMW-04", "How might we assist clinical coordinators in understanding how patient age, gender, and geographical transit distance impact physical clinic arrival?", "SH-01, SH-03: Operations / Desk", "Demographic attendance disparities and transit friction", "OQ-04"),
        ("HMW-05", "How might we help communications staff evaluate the true behavioral efficacy of digital reminders (SMS/WhatsApp) across different patient cohorts?", "SH-03: Front Desk Staff", "Ineffective generic broadcast reminder dispatches", "OQ-05"),
        ("HMW-06", "How might we guide clinic supervisors to optimize slot capacity across high-risk weekdays and time-of-day blocks without overloading the waiting lobby?", "SH-01: Operations Manager", "Peak slot crowding and lobby friction from overbooking", "OQ-06"),
        ("HMW-07", "How might we alert clinical teams when high-risk chronic comorbidity patients (Hypertension/Diabetes) fall out of routine consultation compliance?", "SH-02: Chief Medical Officer", "Compromised continuity of care for chronic lifestyle diseases", "OQ-07"),
        ("HMW-08", "How might we distinguish chronic repeat no-show offenders from accidental one-time misses so receptionists don't waste limited outreach time?", "SH-03: Front Desk Staff", "Alphabetical dialing wasting staff calling hours", "OQ-08"),
        ("HMW-09", "How might we quantify the exact monetary leakage and wasted physician consultation hours so hospital executives can justify targeted intervention budgets?", "SH-01, SH-02: Exec Management", "Unquantified departmental financial losses", "OQ-09"),
        ("HMW-10", "How might we equip front-desk operators with an early-warning risk tier for upcoming appointments to prioritize high-yield reconfirmation calls?", "SH-03: Front Desk Staff", "Lack of objective prioritization for daily patient calling rosters", "OQ-10"),
        ("HMW-11", "How might we provide corporate HR wellness leads with data evidence to align wellness clinic hours with employee work patterns to boost preventive checkups?", "SH-04: Corporate Partner HR", "Low employee participation in occupational wellness drives", "OQ-06, OQ-04")
    ]
    for r_idx, r_vals in enumerate(hmw_data, start=1):
        for c_idx, val in enumerate(r_vals):
            hmw_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(hmw_table, hmw_widths)

    # Section 4
    add_heading_1("4. Priority Selection (Top 5 HMW Opportunities)")
    add_body_p(
        "From the synthesized opportunity set, five HMW statements are selected as core high-impact priorities based on "
        "operational leverage, clinical risk mitigation, and feasibility within our analytics platform:"
    )

    p_table = doc.add_table(rows=6, cols=3)
    p_widths = [1.0, 2.5, 3.0]
    p_headers = ["Priority Rank", "HMW Statement", "One-Line Strategic Rationale"]
    for i, h in enumerate(p_headers):
        p_table.rows[0].cells[i].paragraphs[0].text = h

    p_data = [
        ("PRIORITY 1\n(HMW-10)", "How might we equip front-desk operators with an early-warning risk tier for upcoming appointments to prioritize high-yield reconfirmation calls?", "Directly solves the primary operational bottleneck by transforming passive scheduling into high-ROI targeted front-desk intervention before slot vacancy occurs."),
        ("PRIORITY 2\n(HMW-02)", "How might we enable scheduling staff to identify appointments whose booking lead time introduces a high decay in patient attendance commitment?", "Lead time is the single largest controllable behavioral driver of appointment forgetfulness and early attrition in outpatient healthcare."),
        ("PRIORITY 3\n(HMW-09)", "How might we quantify the exact monetary leakage and wasted physician consultation hours so hospital executives can justify targeted intervention budgets?", "Provides executive leadership with undeniable financial evidence and capacity metrics necessary to sponsor workflow reforms and track ROI."),
        ("PRIORITY 4\n(HMW-05)", "How might we help communications staff evaluate the true behavioral efficacy of digital reminders across different patient cohorts?", "Validates whether existing communication investments produce statistically significant attendance lift, preventing wasted messaging costs."),
        ("PRIORITY 5\n(HMW-07)", "How might we alert clinical teams when high-risk chronic comorbidity patients (Hypertension/Diabetes) fall out of routine consultation compliance?", "Protects patient health outcomes by preventing disease escalation in vulnerable populations, elevating the project from pure finance to clinical impact.")
    ]
    for r_idx, r_vals in enumerate(p_data, start=1):
        for c_idx, val in enumerate(r_vals):
            p_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(p_table, p_widths)

    # Section 5
    add_heading_1("5. Traceability & OQ Coverage Verification")
    add_body_p(
        "Design-thinking integrity mandates that every single Operational Question established in Step 1 is mapped and addressed "
        "by at least one HMW opportunity statement. The verification matrix below confirms 100% comprehensive coverage:"
    )

    cov_table = doc.add_table(rows=11, cols=4)
    cov_widths = [0.8, 2.5, 1.2, 2.0]
    cov_headers = ["OQ ID", "Operational Question Core Subject", "Covering HMW(s)", "Coverage Verification Status"]
    for i, h in enumerate(cov_headers):
        cov_table.rows[0].cells[i].paragraphs[0].text = h

    cov_data = [
        ("OQ-01", "Overall baseline attendance rate & temporal trend", "HMW-01", "VERIFIED (Direct 1:1 Coverage)"),
        ("OQ-02", "Booking lead time correlation with non-attendance", "HMW-02", "VERIFIED (Direct 1:1 Coverage)"),
        ("OQ-03", "Specialty-specific no-show distributions", "HMW-03", "VERIFIED (Direct 1:1 Coverage)"),
        ("OQ-04", "Demographic factors (age, gender, transit distance)", "HMW-04, HMW-11", "VERIFIED (Multi-HMW Coverage)"),
        ("OQ-05", "SMS/WhatsApp digital reminder efficacy lift", "HMW-05", "VERIFIED (Direct 1:1 Coverage)"),
        ("OQ-06", "Day-of-week and time-slot arrival friction", "HMW-06, HMW-11", "VERIFIED (Multi-HMW Coverage)"),
        ("OQ-07", "Chronic comorbidity compliance (Diabetes/Hypertension)", "HMW-07", "VERIFIED (Direct 1:1 Coverage)"),
        ("OQ-08", "Repeat no-show behavioral distribution (80/20 rule)", "HMW-08", "VERIFIED (Direct 1:1 Coverage)"),
        ("OQ-09", "Monetary loss and doctor-hour capacity waste", "HMW-09", "VERIFIED (Direct 1:1 Coverage)"),
        ("OQ-10", "Composite risk tiering engine for advance triage", "HMW-10", "VERIFIED (Direct 1:1 Coverage)")
    ]
    for r_idx, r_vals in enumerate(cov_data, start=1):
        for c_idx, val in enumerate(r_vals):
            cov_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(cov_table, cov_widths)

    doc.add_paragraph().paragraph_format.space_after = Pt(16)
    
    doc.save(output_path)
    print(f"Step 2 Document successfully created at: {output_path}")

if __name__ == "__main__":
    out_dir = r"c:\Users\deeps\Downloads\Appointment No Show Analysis\docs"
    file_name = "02_HMW_Statements.docx"
    full_path = os.path.join(out_dir, file_name)
    build_step2_document(full_path)
