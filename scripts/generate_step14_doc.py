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

def build_step14_doc(output_path):
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
    sub_run = sub_p.add_run("DEPARTMENT OF MASTER OF COMPUTER APPLICATIONS (MCA)\nMCA 2nd Year — Data Analytics Specialization | Build Plan & Execution Architecture")
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
        ("Execution Plan:", "18 Atomic Build Tasks (BUILD-01 to BUILD-18) | Pytest Verification"),
        ("Deliverable Step:", "Step 14 — Master Implementation Roadmap & AI Execution Plan (BUILD_PLAN)")
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
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(38, 38, 38)
        return p

    # 1. Project Summary
    add_heading_1("1. Project Summary & Master Documentation Links")
    add_body_p(
        "The Appointment No-Show Analysis & Optimization System is an enterprise-grade healthcare analytics web application "
        "designed to diagnose, quantify, and mitigate outpatient appointment non-attendance across 4 clinical facilities in Noida. "
        "Operating at an unmitigated 18.3%-24.2% no-show rate, outpatient healthcare networks leak over ₹14.5 Lakhs annually per branch "
        "while wasting hundreds of specialist physician hours. This platform ingests historical appointment records, patient clinical comorbidities "
        "(Hypertension, Diabetes), booking lead times, and digital reminder logs. Developed using Python, Streamlit, SQLite, and modern Tailwind UI styling, "
        "the system delivers descriptive KPIs, formal hypothesis testing (p < 0.05), revenue leakage accounting, and advance risk-tier triage to optimize daily clinical operations."
    )

    doc_links = [
        ("Step 01", "Project Understanding Document", "01_Appointment_No_Show_Project_Understanding.docx"),
        ("Step 02", "How Might We (HMW) Statements", "02_HMW_Statements.docx"),
        ("Step 03", "Business Needs Analysis (BR)", "03_Needs_Analysis.docx"),
        ("Step 04", "Master Requirements Traceability Matrix (RTM)", "04_Traceability_Matrix.xlsx"),
        ("Step 05", "Business Requirements Document (BRD v1.0)", "05_BRD.docx"),
        ("Step 07", "Technical Requirements Document (TRD / Tech Spec)", "07_TRD.docx"),
        ("Step 08", "Relational Database Schema Design", "08_Database_Design.xlsx"),
        ("Step 09", "Entity-Relationship Diagram & Semantics", "09_ER_Diagram.docx"),
        ("Step 10", "Synthetic Dataset Specification", "10_Dataset_Specification.docx"),
        ("Step 11", "Agile User Stories & Acceptance Criteria", "11_User_Stories.docx"),
        ("Step 12", "UI Framework & Wireframe Blueprints", "12_UI_Framework.docx"),
        ("Step 13", "Comprehensive Test Plan & Test Cases", "13_Test_Cases.docx")
    ]
    l_table = doc.add_table(rows=13, cols=3)
    l_widths = [1.2, 2.8, 2.5]
    l_headers = ["Lifecycle Step", "Document Title", "Saved Deliverable File"]
    for i, h in enumerate(l_headers):
        l_table.rows[0].cells[i].paragraphs[0].text = h
    for r_i, r_data in enumerate(doc_links, start=1):
        for c_i, val in enumerate(r_data):
            l_table.rows[r_i].cells[c_i].paragraphs[0].text = val
    style_table(l_table, l_widths)

    # 2. Final Folder Structure
    add_heading_1("2. Final Target Repository Structure")
    add_body_p("The production code repository is organized into cleanly bounded modules:")
    add_code_block(
        "Appointment_No_Show_Analysis/\n"
        "├── app.py                         # Root entrypoint & executive overview\n"
        "├── requirements.txt               # Pinned package dependencies\n"
        "├── README.md                      # Operations manual & setup guide\n"
        "├── BUILD_PLAN.md                  # Implementation roadmap (AI build guide)\n"
        "├── .streamlit/config.toml         # UI theme & server configuration\n"
        "├── data/app.db                    # SQLite production database\n"
        "├── pages/                         # Multi-page views (1 to 8)\n"
        "├── src/                           # Backend analytics, stats, models & UI\n"
        "├── tests/                         # Automated Pytest test suites\n"
        "└── docs/                          # Word & Excel formal lifecycle documents"
    )

    # 3. Build Tasks BUILD-01 to BUILD-18
    add_heading_1("3. Build Tasks Specification (BUILD-01 through BUILD-18)")
    add_body_p(
        "Eighteen sequential, atomic build tasks govern the implementation phase in Step 15. "
        "Each task specifies its goal, files touched, linked user stories/features, verification method, and required passing tests:"
    )

    b_tasks = [
        ("BUILD-01", "Environment Setup & Dependencies", "requirements.txt, .streamlit/config.toml", "Setup", "Verify imports of openpyxl, docx, and config.toml presence.", "TC-T-12"),
        ("BUILD-02", "Data Ingestion & Cleaning Engine", "src/data/loader.py, src/data/cleaner.py", "US-01, US-02 | FR-01", "Query SQLite vw_appointment_analytics; verify 300 rows and 36 cols.", "TC-T-01, TC-T-02"),
        ("BUILD-03", "Analytics Calculation Modules", "src/analytics/descriptive.py, lead_time.py, financial.py", "US-03, US-05, US-12 | FR-01, FR-02, FR-09", "Assert KPI calculations and revenue loss match sum(fee*is_no_show).", "TC-T-07, TC-T-08, TC-T-10"),
        ("BUILD-04", "Statistical Inference Engine", "src/stats/hypothesis.py", "US-09, US-10 | FR-03, FR-05", "Execute Z-test on SMS reminder; assert Z > 2.5 and p < 0.05.", "TC-T-09"),
        ("BUILD-05", "Predictive Risk Scoring Engine", "src/models/risk_engine.py", "US-11, US-13 | FR-08, FR-10", "Verify multi-factor score and tier boundaries (High >= 60, Med 30-59, Low < 30).", "TC-T-11"),
        ("BUILD-06", "UI Components & Chart Factory", "src/ui.py, src/ui/charts.py", "US-03, US-08 | FR-01, FR-06", "Test Plotly chart factory functions and Tailwind CDN components.", "TC-T-12"),
        ("BUILD-07", "Root Application Landing Page", "app.py", "US-01, US-03, US-04 | FR-01", "Launch streamlit run app.py; confirm root hero header & KPI row.", "TC-T-12"),
        ("BUILD-08", "Page 1: Overview & Executive KPIs", "pages/1_Overview_and_KPIs.py", "US-03, US-04 | FR-01", "Verify multi-branch stacked bars and 30-day moving average curve.", "TC-F-01 to TC-F-04"),
        ("BUILD-09", "Page 2: Lead Time & Specialty", "pages/2_Lead_Time_and_Specialty.py", "US-05, US-06 | FR-02, FR-03", "Verify lead-time decay curve (>30d ~42%) and specialty comparisons.", "TC-F-05, TC-F-06"),
        ("BUILD-10", "Page 3: Demographics & Heatmap", "pages/3_Demographics_and_Arrivals.py", "US-07, US-08 | FR-04, FR-06", "Inspect age cohort bars and 2D arrival heatmap (Mon-Sat vs Day/Time).", "TC-F-07, TC-F-08"),
        ("BUILD-11", "Page 4: Statistical Inference", "pages/4_Statistical_Inference.py", "US-09, US-10 | FR-05, FR-07", "Inspect Z-test reminder lift card, Chi-square table, and odds ratio CI.", "TC-F-09, TC-F-10"),
        ("BUILD-12", "Page 5: Financial & Capacity Loss", "pages/5_Financial_and_Capacity.py", "US-12 | FR-09", "Verify exact INR fee loss summation and doctor consultation idle hours.", "TC-F-12"),
        ("BUILD-13", "Page 6: Predictive Risk Triage", "pages/6_Predictive_Risk_Triage.py", "US-11, US-13 | FR-08, FR-10", "Inspect High-Risk triage calling queue table sorted by descending risk.", "TC-F-11, TC-F-13"),
        ("BUILD-14", "Page 7: Corporate & Simulation", "pages/7_Simulation_and_Corporate.py", "US-14 | FR-11, FR-12", "Adjust overbooking slider (1.00-1.25); verify recovered slots update.", "TC-F-14"),
        ("BUILD-15", "Page 8: Data Quality & Export", "pages/8_Data_Quality_and_Export.py", "US-01, US-02 | FR-01", "Inspect missing value tables and test CSV and Excel export downloads.", "TC-F-16"),
        ("BUILD-16", "Automated Pytest Suite Execution", "tests/test_*.py", "All US & Features", "Run python -m pytest tests/ -v; confirm 14 passed out of 14.", "All TC-T-01 to TC-T-12"),
        ("BUILD-17", "Production Documentation & README", "README.md", "Governance", "Review comprehensive README with setup, commands, and viva guide.", "Manual Audit"),
        ("BUILD-18", "Cloud Deployment Packaging", "Repository root", "Deployment", "Verify relative paths, zero secrets, and Streamlit Community Cloud readiness.", "Cold Boot Smoke Test")
    ]

    bt_table = doc.add_table(rows=19, cols=6)
    bt_widths = [0.8, 1.4, 1.4, 1.1, 1.3, 0.5]
    bt_headers = ["Task ID", "Build Task Title", "Touched Files", "Linked US / Feature", "Verification Method", "Tests"]
    for i, h in enumerate(bt_headers):
        bt_table.rows[0].cells[i].paragraphs[0].text = h
    for r_i, r_data in enumerate(b_tasks, start=1):
        for c_i, val in enumerate(r_data):
            bt_table.rows[r_i].cells[c_i].paragraphs[0].text = val
    style_table(bt_table, bt_widths)

    # 4. Definition of Done
    add_heading_1("4. Definition of Done (DoD)")
    add_bullet_p("100% Must-Have Requirements Functional: ", "All 7 core Must BRs (BR-01, BR-02, BR-03, BR-05, BR-07, BR-09, BR-10) are fully operational.")
    add_bullet_p("Automated Test Suite Pass Rate: ", "All 14 pytest test cases in tests/ pass with zero assertion failures.")
    add_bullet_p("Sub-2.5s Performance: ", "All Streamlit pages load in < 2.5 seconds on initial boot and < 500ms on cached filter changes.")
    add_bullet_p("Zero PII & HIPAA/NABH Compliance: ", "Dataset contains zero real personal data and zero external security credentials.")

    # 5. Commands
    add_heading_1("5. Standard CLI Execution Commands")
    add_body_p("Commands to install, run, and test the project:")
    add_code_block(
        "# 1. Environment Installation\n"
        "python -m venv .venv\n"
        ".venv\\Scripts\\activate\n"
        "pip install -r requirements.txt\n\n"
        "# 2. Dataset Generation\n"
        "python src/generate_data.py\n\n"
        "# 3. Run Automated Tests\n"
        "python -m pytest tests/ -v\n\n"
        "# 4. Launch Application\n"
        "streamlit run app.py"
    )

    # 6. Viva Notes
    add_heading_1("6. Top 10 Viva Examination Questions & High-Scoring Answers")
    viva_qa = [
        ("Q1: Why is appointment non-attendance formulated as an optimization and risk problem rather than basic retrospective reporting?",
         "Retrospective reporting merely logs past empty consultation slots after revenue and doctor hours are already wasted. Formulating it as an optimization and predictive risk problem allows the hospital to classify tomorrow's appointments into High, Medium, and Low risk bands, enabling front-desk personnel to proactively call high-risk patients before the slot is lost."),
        ("Q2: Why did you choose a Dimensional Snow-Star schema in SQLite with a 36-column analytical view?",
         "Normalizing master dimensions (clinics, specialties, doctors, patients) ensures transactional integrity and avoids data redundancy. The denormalized 36-column analytical view allows Streamlit to ingest all multi-dimensional features in a single fast query, cached in memory via @st.cache_data for sub-15ms filter updates."),
        ("Q3: Why was a Two-Sample Proportion Z-Test chosen to evaluate digital reminder efficacy (OQ-05)?",
         "Digital reminder impact compares two independent binomial proportions (Reminder Sent vs. Not Sent). With a sample size exceeding 100 observations (satisfying np >= 5 and n(1-p) >= 5), the sampling distribution of proportion differences is normal by the Central Limit Theorem, making the Z-test the most statistically powerful parametric test."),
        ("Q4: Why was Pearson's Chi-Square Test chosen for specialty attendance analysis (OQ-03)?",
         "Both variables are discrete categorical: Specialty has 6 levels and Appointment Outcome has 2 levels (Show vs. No-Show). Pearson's Chi-Square Test of Independence is the standard statistical test to determine whether attendance distributions are independent across an R x C contingency table."),
        ("Q5: How did you model booking lead time, and what behavioral pattern did the analysis reveal?",
         "Lead time was computed as elapsed calendar days between booking and appointment dates and binned into 5 standard buckets. Analysis revealed an exponential decay curve: bookings within 2 days had ~17% no-shows, whereas bookings exceeding 30 days suffered a 41.7% no-show rate."),
        ("Q6: Why are consultation fees snapshotted in fact_appointments rather than joined dynamically from dim_specialties?",
         "This prevents Slowly Changing Dimension (SCD Type 2) distortion. If a medical department increases its tariff in the future, dynamically joining standard fees would retroactively distort past financial accounting. Capturing actual_consultation_fee at booking preserves historical financial accounting fidelity."),
        ("Q7: What is the operational purpose of the repeat offender Pareto analysis (OQ-08)?",
         "In healthcare operations, the 80/20 rule often applies: a small, concentrated cohort of habitual repeat offenders drives a disproportionate share of total empty slots. Grouping patients by prior missed visits allows front-desk staff to focus phone reconfirmation calls specifically on high-frequency offenders rather than accidental one-time misses."),
        ("Q8: How does your UI architecture integrate Tailwind CSS into Streamlit without style collisions?",
         "Streamlit does not compile Tailwind utility classes. Rather than using unsafe HTML in st.markdown which can bleed into global DOM styles, we wrap semantic HTML with Tailwind CDN inside st.components.v1.html(..., height=...). This isolates components within sandboxed virtual iframes."),
        ("Q9: How do unit tests in tests/test_analysis.py verify mathematical formulas without depending on database state?",
         "Tests like test_tc_t_07 and test_tc_t_08 execute against tiny, hardcoded mock datasets with known analytical answers (e.g. 8 shows and 2 no-shows). This decouples algorithmic verification from database mutations, ensuring tests are deterministic and reproducible."),
        ("Q10: What is the business return on investment (ROI) of deploying this analytics platform?",
         "By identifying high-risk appointment cohorts and enabling targeted pre-visit telephone reconfirmation, outpatient facilities can recapture up to 35% of previously lost consultation slots. For a 4-branch clinic network leaking ₹14.5 Lakhs annually per branch, recapturing one-third of wasted capacity recovers over ₹20 Lakhs in annual network revenue.")
    ]
    for q_text, a_text in viva_qa:
        add_body_p(q_text, bold_prefix="•  ")
        add_body_p(a_text, bold_prefix="   Answer: ")

    doc.add_paragraph().paragraph_format.space_after = Pt(16)
    doc.save(output_path)
    print(f"Step 14 Build Plan Document successfully created at: {output_path}")

if __name__ == "__main__":
    doc_out = r"c:\Users\deeps\Downloads\Appointment No Show Analysis\docs\14_BUILD_PLAN.docx"
    build_step14_doc(doc_out)
