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

def build_step7_trd(output_path):
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
    sub_run = sub_p.add_run("DEPARTMENT OF MASTER OF COMPUTER APPLICATIONS (MCA)\nMCA 2nd Year — Data Analytics Specialization | Capstone Technical Specification")
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
        ("Technical Stack:", "Python 3.12, Streamlit, Pandas, NumPy, SciPy, Statsmodels, Plotly, SQLite, Pytest, Tailwind CSS"),
        ("Deliverable Step:", "Step 07 — Technical Requirements Document (TRD / Tech Spec v1.0)")
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
        r.font.size = Pt(9.0)
        r.font.color.rgb = RGBColor(38, 38, 38)
        # Background light gray
        return p

    # 1. Architecture
    add_heading_1("1. End-to-End System Architecture")
    add_body_p(
        "The technical architecture follows a modular, decoupled data pipeline designed for high performance "
        "and sub-2.5 second analytical rendering on Streamlit Community Cloud. The data flow transitions through five discrete stages: "
        "Relational Persistence (SQLite) -> Cached Ingestion -> Cleaning & Validation -> Statistical & Analytics Engines -> Multi-Page Presentation Layer."
    )
    add_heading_2("Architecture Data Flow (Mermaid Diagram Specification):")
    add_code_block(
        "flowchart TD\n"
        "    subgraph Layer 1: Data Persistence\n"
        "        DB[(SQLite: data/app.db)] -->|SQL / SQLite3 Driver| LDR[src/data/loader.py]\n"
        "    end\n"
        "    subgraph Layer 2: Ingestion & Caching\n"
        "        LDR -->|@st.cache_data / TTL=3600| MEM[In-Memory Pandas DataFrames]\n"
        "        MEM --> CLN[src/data/cleaner.py: Validation & Feature Engineering]\n"
        "    end\n"
        "    subgraph Layer 3: Analytics & Statistical Computation Engines\n"
        "        CLN --> ENG1[src/analytics/descriptive.py: Attendance & Branch Trends]\n"
        "        CLN --> ENG2[src/analytics/lead_time.py: Decay Curve & Buckets]\n"
        "        CLN --> ENG3[src/stats/hypothesis.py: Chi-Square & Z-Tests]\n"
        "        CLN --> ENG4[src/analytics/financial.py: Revenue & Capacity Accounting]\n"
        "        CLN --> ENG5[src/models/risk_engine.py: Multi-Factor Risk Scoring]\n"
        "    end\n"
        "    subgraph Layer 4: Presentation & UI Assembly\n"
        "        ENG1 & ENG2 & ENG3 & ENG4 & ENG5 --> UIC[src/ui/components.py: Tailwind Cards via st.components.v1.html]\n"
        "        ENG1 & ENG2 & ENG3 & ENG4 & ENG5 --> UIG[src/ui/charts.py: Plotly Interactive Visuals]\n"
        "        UIC & UIG --> P0[app.py: Executive Overview Dashboard]\n"
        "        UIC & UIG --> P1[pages/1_Overview_and_KPIs.py]\n"
        "        UIC & UIG --> P2[pages/2_Lead_Time_and_Specialty.py]\n"
        "        UIC & UIG --> P3[pages/3_Statistical_Inference.py]\n"
        "        UIC & UIG --> P4[pages/4_Financial_and_Capacity.py]\n"
        "        UIC & UIG --> P5[pages/5_Predictive_Risk_Triage.py]\n"
        "    end"
    )

    # 2. Folder Structure
    add_heading_1("2. Standard Repository & Folder Structure")
    add_body_p(
        "The codebase enforces enterprise Python packaging standards, separating UI routing, business logic, "
        "data ingestion, statistical analysis, tests, and documentation into strictly bounded directories:"
    )
    add_code_block(
        "Appointment_No_Show_Analysis/\n"
        "├── app.py                         # Application Entrypoint & Executive Overview\n"
        "├── requirements.txt               # Pinned Python package dependencies\n"
        "├── README.md                      # Comprehensive project documentation & quickstart\n"
        "├── pages/                         # Streamlit multi-page routing directory\n"
        "│   ├── 1_Overview_and_KPIs.py     # Network baseline attendance & branch trends\n"
        "│   ├── 2_Lead_Time_and_Specialty.py # Lead-time decay curves & specialty distributions\n"
        "│   ├── 3_Statistical_Inference.py # Chi-Square test, Z-tests, odds ratios & p-values\n"
        "│   ├── 4_Financial_and_Capacity.py# Revenue loss & physician idle doctor-hours accounting\n"
        "│   └── 5_Predictive_Risk_Triage.py# Pre-visit risk tier classification & calling roster\n"
        "├── src/                           # Core business logic & backend modules\n"
        "│   ├── __init__.py\n"
        "│   ├── database/                  # SQLite connection & schema DDL\n"
        "│   │   ├── __init__.py\n"
        "│   │   ├── connection.py          # SQLite context manager & query helpers\n"
        "│   │   └── schema.py              # DDL schema creation & indexing scripts\n"
        "│   ├── data/                      # Ingestion, validation, cleaning, and cache\n"
        "│   │   ├── __init__.py\n"
        "│   │   ├── loader.py              # Cached SQLite ingestion (@st.cache_data)\n"
        "│   │   ├── cleaner.py             # Schema validation, missing values, typing\n"
        "│   │   └── generator.py           # Controlled synthetic dataset generator\n"
        "│   ├── analytics/                 # Analytical calculation modules\n"
        "│   │   ├── __init__.py\n"
        "│   │   ├── descriptive.py         # Proportion calculations, moving averages\n"
        "│   │   ├── lead_time.py           # Interval bucketing & decay functions\n"
        "│   │   ├── specialty.py           # Specialty cross-tabs & fee schedules\n"
        "│   │   ├── demographics.py        # Age cohorts, gender ratios, transit bands\n"
        "│   │   ├── temporal.py            # Day-of-week & time-slot 2D heatmaps\n"
        "│   │   ├── comorbidities.py       # Hypertension & Diabetes compliance ratios\n"
        "│   │   ├── behavioral.py          # Repeat offender frequency & Pareto curves\n"
        "│   │   └── financial.py           # Revenue leakage & doctor-hours waste formulas\n"
        "│   ├── stats/                     # Formal statistical inference\n"
        "│   │   ├── __init__.py\n"
        "│   │   └── hypothesis.py          # Chi-Square, 2-sample Z-test, odds ratios\n"
        "│   ├── models/                    # Heuristic & predictive risk engine\n"
        "│   │   ├── __init__.py\n"
        "│   │   └── risk_engine.py         # Multi-factor score calculator & tiering\n"
        "│   └── ui/                        # Presentation & visual components\n"
        "│       ├── __init__.py\n"
        "│       ├── components.py          # Tailwind CSS KPI cards & HTML containers\n"
        "│       └── charts.py              # Reusable Plotly chart factory functions\n"
        "├── data/                          # Data persistence layer\n"
        "│   ├── app.db                     # Embedded SQLite database file\n"
        "│   └── raw_synthetic_records.csv  # Baseline generated CSV archive\n"
        "├── tests/                         # Automated test suite (Pytest & AppTest)\n"
        "│   ├── __init__.py\n"
        "│   ├── test_data_loader.py        # Schema integrity & loader assertions\n"
        "│   ├── test_analytics.py          # Descriptive, financial, and lead time tests\n"
        "│   ├── test_statistics.py         # Z-test, Chi-square & odds ratio tests\n"
        "│   ├── test_risk_engine.py        # Risk score calculation & tier cutoff tests\n"
        "│   └── test_ui_app.py             # Streamlit AppTest integration suites\n"
        "└── docs/                          # Project lifecycle documentation repository\n"
        "    ├── 01_Project_Understanding.docx\n"
        "    ├── 02_HMW_Statements.docx\n"
        "    ├── 03_Needs_Analysis.docx\n"
        "    ├── 04_Traceability_Matrix.xlsx\n"
        "    ├── 05_BRD.docx\n"
        "    └── 07_TRD.docx"
    )

    # 3. Functional Requirements
    add_heading_1("3. Functional Requirements Specification (FR-01 to FR-12)")
    add_body_p(
        "Each functional requirement specifies the concrete software module, computational logic, "
        "and linked business requirements:"
    )

    fr_table = doc.add_table(rows=13, cols=4)
    fr_widths = [0.8, 2.5, 1.2, 2.0]
    fr_headers = ["FR ID", "Functional Requirement Description", "Linked BR", "Responsible Python Module"]
    for i, h in enumerate(fr_headers):
        fr_table.rows[0].cells[i].paragraphs[0].text = h

    fr_data = [
        ("FR-01", "Multi-Branch Attendance & Trend Aggregator: Ingests appointments, filters by date/branch, and computes daily/weekly/monthly attendance rates and moving averages.", "BR-01, BR-12", "src/analytics/descriptive.py"),
        ("FR-02", "Lead-Time Interval Discretization & Decay Engine: Bins lead times into 5 standardized intervals and computes non-attendance probabilities and Spearman rank correlations.", "BR-02", "src/analytics/lead_time.py"),
        ("FR-03", "Specialty Cross-Tabulator & Chi-Square Engine: Aggregates no-show proportions across 6 specialties and evaluates independence via Scipy stats Chi-square.", "BR-03", "src/analytics/specialty.py\nsrc/stats/hypothesis.py"),
        ("FR-04", "Demographic & Proximity Stratification: Segments attendance across age brackets (18-30, 31-45, 46-60, 61-75, 75+), gender, and transit distance zones.", "BR-04, BR-11", "src/analytics/demographics.py"),
        ("FR-05", "Digital Reminder Efficacy & Z-Test Evaluator: Executes two-sample proportion Z-test and computes relative risk reduction and odds ratios with 95% Wald CI.", "BR-05", "src/stats/hypothesis.py"),
        ("FR-06", "Temporal Day-of-Week & Time-Slot Heatmapper: Cross-tabulates attendance across weekdays (Mon-Sat) and 3 time intervals (Morning, Afternoon, Evening).", "BR-06", "src/analytics/temporal.py"),
        ("FR-07", "Chronic Comorbidity Risk Compliance Engine: Evaluates attendance rates for patients with Hypertension and Diabetes against non-chronic baselines.", "BR-07", "src/analytics/comorbidities.py"),
        ("FR-08", "Repeat Offender Behavioral & Pareto Profiler: Categorizes patients by prior missed visits (0, 1, 2+) and computes cumulative Pareto loss shares.", "BR-08", "src/analytics/behavioral.py"),
        ("FR-09", "Financial Revenue Loss & Physician Idle-Hour Engine: Multiplies unannounced no-shows by specialty consultation fees and sums doctor consultation idle minutes.", "BR-09", "src/analytics/financial.py"),
        ("FR-10", "Composite Predictive Risk Scoring & Triage Engine: Calculates a weighted composite risk index (0-100) and assigns appointments into High, Medium, Low risk tiers.", "BR-10", "src/models/risk_engine.py"),
        ("FR-11", "Corporate Shift Correlation Module: Computes occupational wellness attendance rates across general vs. evening corporate employee shifts.", "BR-11", "src/analytics/corporate.py"),
        ("FR-12", "Calibrated Overbooking Simulation Engine: Models simulated capacity recovery and lobby wait-time variance across varying overbooking multipliers.", "BR-12", "src/analytics/simulation.py")
    ]
    for r_i, r_data in enumerate(fr_data, start=1):
        for c_i, val in enumerate(r_data):
            fr_table.rows[r_i].cells[c_i].paragraphs[0].text = val
    style_table(fr_table, fr_widths)

    # 4. Data Layer
    add_heading_1("4. Data Persistence, Caching & Validation Layer")
    add_body_p(
        "Persistence Engine: Data is housed in an embedded SQLite 3 database file (`data/app.db`) stored directly within the repository. "
        "This eliminates external database connection latencies and credential exposure, ensuring 100% plug-and-play execution."
    )
    add_body_p(
        "Streamlit Caching Pipeline: To ensure instant UI responsiveness, data ingestion is decorated with `@st.cache_data(ttl=3600, show_spinner=False)`. "
        "The cache key is hashed on the database file modification timestamp. In-memory queries execute in under 15 milliseconds."
    )
    add_body_p(
        "Data Validation & Type Enforcement: Ingested DataFrames are validated through strict assertions: non-null primary keys, "
        "valid status enumerations ('Show', 'No-Show'), non-negative lead times (`LeadTimeDays >= 0`), valid age bounds (`18 <= PatientAge <= 95`), "
        "and boolean conversions for comorbidity and reminder flags."
    )

    # 5. Statistical Methods
    add_heading_1("5. Statistical Methodology for Operational Questions (OQ-01 to OQ-10)")
    add_body_p(
        "Every operational question is grounded in established statistical, probabilistic, and econometric formulations:"
    )

    stat_table = doc.add_table(rows=11, cols=4)
    stat_widths = [0.8, 1.8, 2.5, 1.4]
    stat_headers = ["OQ ID", "Operational Question", "Mathematical / Statistical Method", "Python Implementation"]
    for i, h in enumerate(stat_headers):
        stat_table.rows[0].cells[i].paragraphs[0].text = h

    stat_data = [
        ("OQ-01", "Network Attendance Rates & Trends", "Proportion metric p = sum(NoShow) / N; 7-day and 30-day rolling moving averages", "pandas.DataFrame.rolling()\nnp.mean()"),
        ("OQ-02", "Lead-Time Decay Function", "Interval binning; Spearman rank correlation coefficient rho; Logistic non-linear curve fitting", "scipy.stats.spearmanr\npandas.cut()"),
        ("OQ-03", "Specialty No-Show Independence", "Pearson's Chi-Square Test of Independence; contingency table cross-tabulation; Cramer's V", "scipy.stats.chi2_contingency"),
        ("OQ-04", "Demographic & Proximity Cohorts", "Stratified group aggregation; Kruskal-Wallis non-parametric H-test across age cohorts", "scipy.stats.kruskal\npandas.groupby()"),
        ("OQ-05", "Reminder Efficacy & Lift", "Two-Sample Proportion Z-Test; Relative Risk Reduction (RRR); Odds Ratio with 95% Wald CI", "statsmodels.stats.proportion.\nproportions_ztest"),
        ("OQ-06", "Temporal Day & Slot Patterns", "Two-Way cross-tabulation pivot matrix; cell frequency Chi-square test; bivariate heatmap", "pandas.pivot_table()\nplotly.express.density_heatmap"),
        ("OQ-07", "Chronic Comorbidity Compliance", "Subgroup relative risk ratio RR = p_chronic / p_non_chronic; Mantel-Haenszel common odds", "scipy.stats.fisher_exact"),
        ("OQ-08", "Repeat Offender Behavioral Loss", "Discrete frequency distribution; Lorenz curve formulation; 80/20 Pareto cumulative proportion", "numpy.cumsum()\npandas.qcut()"),
        ("OQ-09", "Monetary & Capacity Loss", "Summation: Fee * IsNoShow; Summation: SlotDurationMinutes / 60.0 * IsNoShow", "numpy.sum()\npandas.DataFrame.sum()"),
        ("OQ-10", "Predictive Risk Scoring Engine", "Weighted multi-attribute logistic scoring model: Score = sum(w_k * f_ik); Min-Max normalization", "scikit-learn / custom weighted\nnumpy vector dot product")
    ]
    for r_i, r_data in enumerate(stat_data, start=1):
        for c_i, val in enumerate(r_data):
            stat_table.rows[r_i].cells[c_i].paragraphs[0].text = val
    style_table(stat_table, stat_widths)

    # 6. UI Notes
    add_heading_1("6. User Interface Architecture & Styling Specifications")
    add_bullet_p("Tailwind CSS CDN Integration: ", "Integrated via `st.components.v1.html(..., height=...)` to inject modern Tailwind CSS utilities (rounded cards, drop shadows, responsive flex containers, gradient badges).")
    add_bullet_p("Corporate Healthcare Color Palette: ", "Primary Brand Navy: #1F497D; Vibrant Accent Blue: #2E75B6; Soft Slate Neutral: #F2F5F9; Card Background: #FFFFFF; Alert Red (High Risk): #E53E3E; Success Green: #38A169.")
    add_bullet_p("Plotly Visualization Standard: ", "All analytical charts are styled with a clean white template (`template='plotly_white'`), interactive tooltips, high-contrast markers, and responsive container auto-sizing (`use_container_width=True`).")
    add_bullet_p("Responsive Layout: ", "Streamlit wide-screen layout (`layout='wide'`) with a persistent sidebar filter drawer containing Date Range pickers, Multi-Branch selectors, and Specialty dropdowns.")

    # 7. Performance, Error Handling & Security
    add_heading_1("7. Performance, Error Handling & Security Architecture")
    add_bullet_p("Performance Optimization: ", "Sub-second query execution achieved by pre-indexing SQLite columns (`AppointmentDate`, `ClinicBranch`, `Specialty`, `Status`). Categorical columns in Pandas are downcast to `category` dtypes, reducing memory footprint by over 65%.")
    add_bullet_p("Graceful Error Handling: ", "All analytical functions wrap database interactions in `try-except-finally` blocks. If an invalid filter produces empty slices, the UI renders user-friendly warning banners (`st.warning('No records found matching selected filter criteria.')`) without crashing the session.")
    add_bullet_p("Data Privacy & Security: ", "100% synthetic dataset containing zero real Personally Identifiable Information (PII). No patient names, phone numbers, or government IDs exist in the database. Read-only database connections prevent SQL injection write mutations.")

    # 8. requirements.txt
    add_heading_1("8. Pinned Package Dependencies (requirements.txt)")
    add_body_p("To prevent environment drift and ensure 100% deterministic builds on Streamlit Community Cloud, dependencies are strictly pinned:")
    add_code_block(
        "python==3.12.*\n"
        "streamlit==1.38.0\n"
        "pandas==2.2.2\n"
        "numpy==1.26.4\n"
        "scipy==1.13.1\n"
        "statsmodels==0.14.2\n"
        "plotly==5.24.0\n"
        "pytest==8.3.2\n"
        "openpyxl==3.1.5\n"
        "python-docx==1.2.0"
    )

    # 9. Deployment Requirements
    add_heading_1("9. Deployment Requirements (Streamlit Community Cloud & GitHub)")
    add_bullet_p("Hosting Platform: ", "Streamlit Community Cloud connected to GitHub public/private repository.")
    add_bullet_p("Deployment Branch: ", "Production deployments trigger automatically on push to the `main` branch.")
    add_bullet_p("Entrypoint Specification: ", "Primary app file designated as `app.py` in root directory.")
    add_bullet_p("Environment Configuration: ", "Standard `.streamlit/config.toml` configuring server theme, dark/light modes, and headless port binding.")

    # 10. Technical Risks
    add_heading_1("10. Technical Risks & Mitigation Strategies")
    trisk_table = doc.add_table(rows=5, cols=4)
    trisk_widths = [0.8, 2.2, 1.8, 1.7]
    trisk_headers = ["Risk ID", "Technical Risk Description", "Architectural Impact", "Technical Mitigation Strategy"]
    for i, h in enumerate(trisk_headers):
        trisk_table.rows[0].cells[i].paragraphs[0].text = h
    trisk_rows = [
        ("T-RSK-01", "Streamlit 1 GB RAM memory limit exceeded during large aggregations.", "App container crash & reload", "Enforce strict DataFrame downcasting (`int32` to `int16`, `object` to `category`), pre-aggregating metrics in SQLite."),
        ("T-RSK-02", "Plotly re-rendering latency on complex heatmaps.", "UI sluggishness > 3 seconds", "Pre-compute pivot matrices and utilize Plotly WebGL or SVG rasterization where appropriate."),
        ("T-RSK-03", "Tailwind CDN blockages or iframe rendering glitches.", "Unstyled HTML card components", "Embed minimal fallback CSS directly in HTML strings so cards remain styled even if CDN requests are throttled."),
        ("T-RSK-04", "Streamlit AppTest session state mismatch during automated pytest execution.", "False test assertion failures", "Use official `streamlit.testing.v1.AppTest` mock sessions with initialized session states.")
    ]
    for r_i, r_data in enumerate(trisk_rows, start=1):
        for c_i, val in enumerate(r_data):
            trisk_table.rows[r_i].cells[c_i].paragraphs[0].text = val
    style_table(trisk_table, trisk_widths)

    doc.add_paragraph().paragraph_format.space_after = Pt(16)
    
    doc.save(output_path)
    print(f"Step 7 TRD Document successfully created at: {output_path}")

if __name__ == "__main__":
    out_dir = r"c:\Users\deeps\Downloads\Appointment No Show Analysis\docs"
    file_name = "07_TRD.docx"
    full_path = os.path.join(out_dir, file_name)
    build_step7_trd(full_path)
