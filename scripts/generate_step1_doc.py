import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
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
        set_cell_margins(cell, top=120, bottom=120, left=160, right=160)
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
            set_cell_margins(cell, top=90, bottom=90, left=150, right=150)
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.name = 'Calibri'
                    r.font.size = Pt(9.5)
                    r.font.color.rgb = RGBColor(51, 51, 51)
    # Set widths
    for row in table.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = Inches(width)

def build_step1_document(output_path):
    doc = Document()
    
    # Page setup - 1 inch margins
    sections = doc.sections
    for section in sections:
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
        ("Deliverable Step:", "Step 01 — Project Understanding Document (Comprehensive Specification)")
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
    add_heading_1("1. Project Description")
    add_body_p(
        "The Appointment No-Show Analysis & Optimization System is an enterprise data analytics web application engineered "
        "to investigate, diagnose, and mitigate patient appointment non-attendance across multi-specialty outpatient clinics and corporate health networks. "
        "In healthcare operations, when a patient books a consultation slot with a medical specialist but fails to attend without prior cancellation (a 'no-show'), "
        "the clinical infrastructure suffers immediate operational friction: specialized physicians sit idle during allocated consultation blocks, clinic staffing "
        "remains suboptimally allocated, waiting queues for genuinely ailing patients lengthen unnecessarily, and the healthcare provider experiences substantial revenue leakage."
    )
    add_body_p(
        "This project models the real-world operational challenges of ApexCare Health Network, a healthcare network operating four modern outpatient centers in Noida "
        "(Sector 62, Sector 18, Expressway Campus, and Greater Noida) that also manages occupational health and preventive wellness consultations for leading corporate IT partners "
        "(including a representative 500-employee IT firm in Noida). The application processes historical appointment records containing multi-dimensional features—such as booking lead time, "
        "patient demographics (age groups, gender, residential proximity), chronic comorbidity indicators (Hypertension, Diabetes), digital reminder dispatch logs, and appointment slot timings. "
        "Developed using Python, Streamlit, Plotly, SQLite, and modern UI styling principles, the platform delivers descriptive, exploratory, statistical, and risk-scoring insights to enable clinic administrators "
        "to move from reactive overbooking to proactive, data-backed operational intervention."
    )

    # Section 2
    add_heading_1("2. Background & Problem Context")
    add_body_p(
        "Currently, outpatient scheduling at ApexCare Health Network and affiliated occupational health clinics operates via legacy desk-based and online booking tools "
        "that capture appointments in transactional silos. Clinic receptionists and desk coordinators manually schedule appointments on fixed 20-to-30 minute consultation blocks. "
        "However, historical data reveals that between 20% and 26% of scheduled patients consistently fail to show up for their scheduled consultations."
    )
    add_body_p(
        "Why This Is an Urgent Problem for Healthcare Management:", "Key Bottlenecks: "
    )
    add_bullet_p("Physician Idle Time & Resource Waste: ", "Specialist physicians (e.g., Cardiologists, Orthopedic Surgeons, Clinical Psychologists) are paid high fixed or hourly retainers. An unannounced no-show causes 20–30 minutes of completely non-recoverable clinical idle time.")
    add_bullet_p("Inflated Patient Waitlist Times: ", "Routine patients seeking urgent consultations often face 5-to-10 day waiting periods, unaware that up to one in four slots on the daily roster will end up unused.")
    add_bullet_p("Arbitrary and Blunt Overbooking: ", "To counteract empty slots, clinics currently resort to blunt double-booking heuristics (e.g., booking 2 patients for every 9:00 AM slot). When both patients unexpectedly show up, the waiting lobby becomes overcrowded, doctor consultation times get compressed, and patient satisfaction drops steeply.")
    add_bullet_p("Lack of Granular Visibility: ", "Clinic managers have no analytical dashboards to determine which patient cohorts (e.g., long lead-time bookings, young adults, patients without SMS alerts, specific weekdays) represent high risk, preventing targeted communication or smart reconfirmation workflows.")

    # Section 3
    add_heading_1("3. Problem Statement")
    add_body_p(
        "\"Outpatient healthcare clinics and corporate wellness centers in the Noida urban corridor experience an unmitigated 24.2% patient appointment no-show rate, "
        "resulting in severe physician scheduling inefficiencies, avoidable operating revenue loss of over ₹14.5 Lakhs annually per facility, and extended patient wait times, "
        "due to the complete absence of a centralized, data-driven analytical system to identify no-show drivers, quantify operational waste, and guide targeted clinical intervention strategies.\""
    )

    # Section 4
    add_heading_1("4. Intended Audience & End Users")
    add_body_p(
        "The application is architected specifically for healthcare administrators, medical directors, and front-desk scheduling coordinators. "
        "The following matrix outlines primary user personas, application access frequency, and core functional responsibilities:"
    )
    
    u_table = doc.add_table(rows=5, cols=4)
    u_widths = [1.5, 1.2, 1.3, 2.5]
    u_headers = ["User Role", "Target Persona", "Access Frequency", "Primary Operational Tasks in Application"]
    for i, h in enumerate(u_headers):
        u_table.rows[0].cells[i].paragraphs[0].text = h
    
    u_rows = [
        ("Clinic Operations Manager", "Smt. Kavita Sharma (Admin Lead)", "Daily (Morning & Evening)", "Reviews overall attendance KPIs, monitors weekly no-show trends, evaluates capacity utilization rates, and adjusts staffing rosters."),
        ("Medical Superintendent / CMO", "Dr. A. K. Banerjee (Chief Medical Officer)", "Weekly / Monthly Review", "Assesses specialty-wise clinical utilization, physician idle cost metrics, and chronic disease patient follow-up compliance."),
        ("Outpatient Scheduling Supervisor", "Mr. Rohit Verma (Front Desk Lead)", "Real-time / Daily", "Identifies high-risk appointment cohorts for pre-visit telephonic reconfirmation, SMS reminder management, and buffer management."),
        ("Corporate Wellness Coordinator", "Ms. Neha Gupta (NoidaTech HR Wellness)", "Bi-Weekly", "Monitors occupational health check attendance among corporate employees, tracks seasonal checkup adherence, and schedules on-site wellness camps.")
    ]
    for r_idx, r_data in enumerate(u_rows, start=1):
        for c_idx, val in enumerate(r_data):
            u_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(u_table, u_widths)

    # Section 5
    add_heading_1("5. Project Stakeholders")
    add_body_p(
        "The project governance model engages five primary stakeholder groups whose strategic and operational decisions are driven by the analytics platform:"
    )

    sh_table = doc.add_table(rows=6, cols=6)
    sh_widths = [0.8, 1.3, 1.2, 1.3, 0.7, 1.2]
    sh_headers = ["SH ID", "Stakeholder", "Role", "Interest", "Influence", "Decision Made with the App"]
    for i, h in enumerate(sh_headers):
        sh_table.rows[0].cells[i].paragraphs[0].text = h
        
    sh_rows = [
        ("SH-01", "Clinic Operations Manager", "Operational Custodian", "Maximize clinic throughput, reduce slot vacancies, optimize staff distribution.", "High", "Decides daily buffer slot allocations, overbooking quotas, and staff shift distribution."),
        ("SH-02", "Chief Medical Officer (CMO)", "Clinical Executive", "Doctor time utilization, clinical continuity for chronic disease patients.", "High", "Decides specialist roster schedules, consultation slot duration policies, and medical follow-up guidelines."),
        ("SH-03", "Front Desk & Scheduling Staff", "End User / Operator", "Smooth patient queue flow, reduced lobby friction, predictable intake.", "Medium", "Prioritizes daily phone calls and re-verification reminders to patients flagged as high-risk."),
        ("SH-04", "Corporate Partner HR / Wellness", "Client Sponsor", "Ensuring IT workforce completes annual preventive and occupational health checkups.", "Medium", "Adjusts corporate wellness communication campaigns and decides slot timings matching IT shifts."),
        ("SH-05", "Data Analyst / Academic Evaluator", "Developer & Auditor", "Rigorous statistical analysis, reproducible ETL, data integrity, and viva defense.", "High", "Validates statistical significance of findings, feature engineering validity, and model evaluation metrics.")
    ]
    for r_idx, r_data in enumerate(sh_rows, start=1):
        for c_idx, val in enumerate(r_data):
            sh_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(sh_table, sh_widths)

    # Section 6
    add_heading_1("6. Operational Questions (OQs)")
    add_body_p(
        "To ensure business relevance, the analytics platform is designed to answer ten specific Operational Questions (OQ-01 through OQ-10). "
        "Every single question is strictly verifiable and mathematically computable from the underlying database features:"
    )

    oq_table = doc.add_table(rows=11, cols=4)
    oq_widths = [0.8, 2.5, 1.8, 1.4]
    oq_headers = ["OQ ID", "Operational Question", "Required Data Attributes / Features", "Analytical Method"]
    for i, h in enumerate(oq_headers):
        oq_table.rows[0].cells[i].paragraphs[0].text = h

    oq_rows = [
        ("OQ-01", "What is the baseline appointment no-show rate across all clinic locations, and how does it trend over time?", "AppointmentID, Status, AppointmentDate, ClinicBranch", "Proportion metrics, Monthly/Weekly Time-series aggregation"),
        ("OQ-02", "How does booking lead time (interval between booking date and visit date) correlate with patient no-show probability?", "BookingDate, AppointmentDate, LeadTimeDays, Status", "Binned lead-time distribution, Pearson/Spearman correlation, Odds ratio"),
        ("OQ-03", "Which medical specialties (e.g., General Medicine, Dermatology, Orthopedics, Mental Health) exhibit the highest no-show rates?", "Specialty, Status, AppointmentID", "Categorical aggregation, Cross-tabulation, Chi-square independence test"),
        ("OQ-04", "How do patient demographic dimensions (age groups, gender, neighborhood distance) influence attendance patterns?", "PatientAge, AgeGroup, Gender, Neighborhood, DistanceKm, Status", "Demographic stratification, cohort comparison, bar/violin distribution plots"),
        ("OQ-05", "Does dispatching digital SMS/WhatsApp appointment reminders measurably reduce patient no-show occurrence?", "SMSReminderSent (0/1), ReminderCount, Status", "Two-sample proportion Z-test, Relative risk reduction, Percentage lift"),
        ("OQ-06", "How do day-of-week (Monday–Saturday) and time-slot intervals (Morning vs. Afternoon vs. Evening) impact patient attendance?", "AppointmentDayOfWeek, TimeSlotBucket, HourOfDay, Status", "Two-dimensional heatmaps, Day/Time slot pivot matrices"),
        ("OQ-07", "Are patients diagnosed with chronic health conditions (Hypertension, Diabetes) more or less compliant with scheduled appointments?", "Hypertension (0/1), Diabetes (0/1), ComorbidityScore, Status", "Sub-group rate comparison, Risk ratio calculation, Multi-bar charts"),
        ("OQ-08", "What proportion of missed appointments stems from chronic repeat no-show patients versus first-time or one-off misses?", "PatientID, PreviousAppointmentsCount, PreviousNoShowsCount, Status", "Patient-level behavioral frequency distribution, Pareto analysis (80/20 rule)"),
        ("OQ-09", "What is the quantified financial loss and doctor-hour capacity waste caused by appointment no-shows per month across departments?", "Status, ConsultationFee, SlotDurationMinutes, Specialty", "Financial summation formula: sum(ConsultationFee * IsNoShow), Capacity hour sum"),
        ("OQ-10", "Can a composite risk scoring model segment upcoming appointments into High, Medium, and Low risk to prioritize front-desk confirmation outreach?", "LeadTimeDays, SMSReminderSent, AgeGroup, PreviousNoShows, Specialty", "Weighted rule-based / logistic scoring algorithm with risk tier stratification")
    ]
    for r_idx, r_data in enumerate(oq_rows, start=1):
        for c_idx, val in enumerate(r_data):
            oq_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(oq_table, oq_widths)

    add_heading_2("Data Answerability Verification Check:")
    add_body_p(
        "Verification Confirmation: All ten operational questions (OQ-01 to OQ-10) map 1:1 to discrete, synthetically generated relational schema fields. "
        "No question requires external, uncollectible, or subjective data. Statistical significance tests (Chi-Square for categorical associations and Z-tests for reminder efficacy) "
        "can be executed programmatically in Python using SciPy and Statsmodels."
    )

    # Section 7
    add_heading_1("7. SMART Project Objectives")
    add_body_p("The project is governed by five Specific, Measurable, Achievable, Relevant, and Time-bound (SMART) objectives:")
    add_bullet_p("OBJ-01 (Descriptive Analysis): ", "Ingest and process 10,000+ synthetic outpatient appointment records to establish an exact baseline no-show rate within ±0.1% accuracy across 4 clinics and 6 specialties.")
    add_bullet_p("OBJ-02 (Diagnostic & Statistical Rigor): ", "Quantify the statistical impact of lead time, reminder notifications (SMS), and patient comorbidities on attendance, delivering p-values (<0.05) and odds ratios for all core factors.")
    add_bullet_p("OBJ-03 (Financial & Capacity Loss Quantification): ", "Implement an automated monetary and doctor-hour loss calculation engine that computes monthly revenue leakage and doctor idle time across all departments with 100% mathematical precision.")
    add_bullet_p("OBJ-04 (Predictive Risk Stratification): ", "Develop and validate a composite No-Show Risk Scoring engine that categorizes appointments into High (>60%), Medium (30–60%), and Low (<30%) probability tiers, achieving an analytical classification accuracy above 75%.")
    add_bullet_p("OBJ-05 (Interactive Dashboard Delivery): ", "Architect, package, test with automated pytest test suites (>85% code coverage), and deploy a production-ready Streamlit application on Streamlit Community Cloud within the 18-step project lifecycle.")

    # Section 8
    add_heading_1("8. Project Scope")
    add_body_p("A well-defined scope boundary guarantees project feasibility within academic timelines while maintaining commercial-grade depth:")

    sc_table = doc.add_table(rows=7, cols=3)
    sc_widths = [1.2, 2.7, 2.6]
    sc_headers = ["Functional Area", "In-Scope (Deliverables Built)", "Out-of-Scope (Future / Excluded)"]
    for i, h in enumerate(sc_headers):
        sc_table.rows[0].cells[i].paragraphs[0].text = h

    sc_rows = [
        ("Data Pipeline & ETL", "Synthetic data generation script (Faker/NumPy) modeling 10,000+ records; SQLite database storage; data cleansing and feature engineering in Python (Pandas).", "Live database streaming pipelines (Apache Kafka), hospital EMR FHIR real-time integration."),
        ("Analytics & Statistics", "Univariate, bivariate, and multivariate statistical EDA; Hypothesis testing (Chi-square, 2-proportion Z-test); Correlation matrices.", "Deep learning / neural network black-box models; Natural Language Processing on doctor prescription notes."),
        ("Predictive Risk Tiering", "Feature-weighted risk index / logistic heuristic scoring engine for immediate high-risk slot identification.", "Autonomous robot tele-calling agents; automated bi-directional SMS conversational chatbots."),
        ("Visualization & UI", "Multi-page interactive Streamlit web dashboard; interactive Plotly charts (heatmaps, bar plots, line charts, box plots); modern Tailwind CSS design cues.", "Native iOS/Android mobile apps; proprietary desktop software installations."),
        ("Testing & Quality", "Unit tests and integration test suites using pytest; automated data validation assertions; edge case handling.", "Hardware stress testing; distributed cluster load testing."),
        ("Deployment & Artifacts", "Public cloud deployment on Streamlit Community Cloud; complete ABES MCA-formatted Project Document and Viva Presentation slide deck.", "On-premise Kubernetes multi-node cluster deployment; proprietary commercial license management.")
    ]
    for r_idx, r_data in enumerate(sc_rows, start=1):
        for c_idx, val in enumerate(r_data):
            sc_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(sc_table, sc_widths)

    # Section 9
    add_heading_1("9. Assumptions and Constraints")
    add_heading_2("Project Assumptions:")
    add_bullet_p("A-01 (Data Governance): ", "All patient appointment records analyzed are 100% synthetic, generated in compliance with healthcare data privacy principles (HIPAA/NABH), containing zero Real Personally Identifiable Information (PII).")
    add_bullet_p("A-02 (Operational Realism): ", "Synthetic data distributions (e.g., baseline no-show rate of 20-25%, age distribution 18-85, lead time 0-45 days) are calibrated according to published medical informatics literature and hospital benchmarks.")
    add_bullet_p("A-03 (Consultation Model): ", "Consultation durations follow standard outpatient slots (20 minutes for general medicine, 30 minutes for specialized cardiology/orthopedics) with fixed specialty fee schedules.")
    add_bullet_p("A-04 (Organizational Scope): ", "The system represents a multi-facility suburban network consisting of 4 clinical locations and 6 distinct medical specialties across Noida and NCR.")

    add_heading_2("Project Constraints:")
    add_bullet_p("C-01 (Deployment Environment): ", "The public analytical web application must run entirely within the free-tier memory (1 GB RAM) and CPU limits of Streamlit Community Cloud.")
    add_bullet_p("C-02 (Database Architecture): ", "Data storage is restricted to a lightweight, zero-configuration embedded SQLite database file (`noshow_analytics.db`) to ensure seamless reproducibility across grading environments.")
    add_bullet_p("C-03 (Academic Governance): ", "The project must strictly adhere to the ABES Engineering College MCA 2nd Year project guidelines, meeting syllabus standards for Data Analytics, software documentation, and viva examination rubrics.")

    # Section 10
    add_heading_1("10. Measurable Success Criteria")
    add_body_p("Project completion and quality will be evaluated against five quantifiable success benchmarks:")

    crit_table = doc.add_table(rows=6, cols=4)
    crit_widths = [1.2, 1.8, 1.8, 1.7]
    crit_headers = ["Category", "Success Metric", "Target Benchmark", "Evaluation & Verification Method"]
    for i, h in enumerate(crit_headers):
        crit_table.rows[0].cells[i].paragraphs[0].text = h

    crit_rows = [
        ("Analytical Coverage", "OQ Resolution Rate", "100% (All 10 OQs answered)", "Direct verification against interactive dashboard tabs and statistical reports."),
        ("Statistical Validity", "Hypothesis Rigor", "p < 0.05 for key drivers", "Formal hypothesis test tables generated with p-values and confidence intervals."),
        ("Software Quality", "Automated Test Coverage", ">= 85% coverage with pytest", "Passing pytest execution with terminal execution report and zero failed assertions."),
        ("User Experience", "Dashboard Load Time", "< 2.5 seconds on initial query", "Streamlit performance caching (`@st.cache_data`) benchmarked across 10,000 records."),
        ("Academic Compliance", "ABES MCA Rubric Score", "100% compliance with ABES template", "Final project report and slide deck fully formatted to departmental dissertation specifications.")
    ]
    for r_idx, r_data in enumerate(crit_rows, start=1):
        for c_idx, val in enumerate(r_data):
            crit_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(crit_table, crit_widths)

    doc.add_paragraph().paragraph_format.space_after = Pt(16)
    
    # Save document
    doc.save(output_path)
    print(f"Document successfully created at: {output_path}")

if __name__ == "__main__":
    out_dir = r"c:\Users\deeps\Downloads\Appointment No Show Analysis"
    file_name = "01_Appointment_No_Show_Project_Understanding.docx"
    full_path = os.path.join(out_dir, file_name)
    build_step1_document(full_path)
