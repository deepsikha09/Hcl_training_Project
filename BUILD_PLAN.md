# BUILD_PLAN.md — Implementation Roadmap & AI Execution Guide

**Project:** Appointment No-Show Analysis & Optimization System  
**Project ID:** P101  
**Batch ID:** ABES_OND26 MCA 2ND YEAR_DATA ANALYTICS_BATCH 01  
**Organization Modeled:** ApexCare Health Network, Noida  
**Student:** Deepshikha | Roll No: `2500320140047` | MCA 2nd Year Data Analytics, ABES Engineering College  
**Faculty Mentor:** Dinesh Sereno  
**Technology Stack:** Python 3.12, Streamlit 1.38+, Pandas, NumPy, SciPy, Statsmodels, Plotly, SQLite 3, Pytest, Tailwind CSS CDN  

---

## 1. Project Summary & Master Documentation Links

The Appointment No-Show Analysis & Optimization System is an enterprise-grade healthcare analytics web application designed to diagnose, quantify, and mitigate outpatient appointment non-attendance across 4 clinical facilities in Noida. Operating at an unmitigated 18.3%–24.2% no-show rate, outpatient healthcare networks leak over ₹14.5 Lakhs annually per branch while wasting hundreds of specialist physician hours. This platform ingests historical appointment records, patient clinical comorbidities (Hypertension, Diabetes), booking lead times, and digital reminder logs. Developed using Python, Streamlit, SQLite, and modern Tailwind UI styling, the system delivers descriptive KPIs, formal hypothesis testing ($p < 0.05$), revenue leakage accounting, and advance risk-tier triage to optimize daily clinical operations.

### Master Project Documentation Index:
1. **Step 01:** [Project Understanding Document](docs/01_Appointment_No_Show_Project_Understanding.docx)
2. **Step 02:** [How Might We (HMW) Statements](docs/02_HMW_Statements.docx)
3. **Step 03:** [Business Needs Analysis (BR)](docs/03_Needs_Analysis.docx)
4. **Step 04:** [Master Requirements Traceability Matrix (RTM)](docs/04_Traceability_Matrix.xlsx)
5. **Step 05:** [Business Requirements Document (BRD v1.0)](docs/05_BRD.docx)
6. **Step 07:** [Technical Requirements Document (TRD / Tech Spec)](docs/07_TRD.docx)
7. **Step 08:** [Relational Database Schema Design](docs/08_Database_Design.xlsx)
8. **Step 09:** [Entity-Relationship Diagram & Semantics](docs/09_ER_Diagram.docx)
9. **Step 10:** [Synthetic Dataset Specification](docs/10_Dataset_Specification.docx)
10. **Step 11:** [Agile User Stories & Acceptance Criteria](docs/11_User_Stories.docx)
11. **Step 12:** [UI Framework & Wireframe Blueprints](docs/12_UI_Framework.docx)
12. **Step 13:** [Comprehensive Test Plan & Test Cases](docs/13_Test_Cases.docx)

---

## 2. Final Repository & Folder Structure

```text
Appointment_No_Show_Analysis/
├── app.py                             # Root application entrypoint & executive overview
├── requirements.txt                   # Pinned package dependencies
├── README.md                          # Production operations manual & deployment guide
├── BUILD_PLAN.md                      # Step-by-step implementation guide (this document)
├── .streamlit/
│   └── config.toml                    # Server & theme configuration (Navy/Slate)
├── data/
│   ├── app.db                         # Production SQLite 3 database (tables, indexes, view)
│   ├── Appointment_No_Show_data.xlsx  # Multi-sheet master dataset workbook
│   ├── dim_clinics.csv                # Clinic branches master CSV
│   ├── dim_specialties.csv            # Medical departments master CSV
│   ├── dim_doctors.csv                # Specialists master CSV
│   ├── dim_patients.csv               # Registered patients master CSV
│   ├── fact_appointments.csv          # Transactional fact appointments CSV
│   ├── vw_appointment_analytics.csv   # 36-column denormalized view CSV
│   └── validation_report.txt          # Integrity and pattern audit log
├── pages/
│   ├── 1_Overview_and_KPIs.py         # UI-01: Baseline KPIs & temporal moving average
│   ├── 2_Lead_Time_and_Specialty.py   # UI-02: Lead-time decay curves & specialty bars
│   ├── 3_Demographics_and_Arrivals.py # UI-03: Age cohorts & 2D arrival heatmap
│   ├── 4_Statistical_Inference.py     # UI-04: Z-test, Chi-square & odds ratios
│   ├── 5_Financial_and_Capacity.py    # UI-05: Revenue leakage & doctor idle hours
│   ├── 6_Predictive_Risk_Triage.py    # UI-06: Advance risk scoring & phone calling queue
│   ├── 7_Simulation_and_Corporate.py  # UI-07: Overbooking simulator & corporate shifts
│   └── 8_Data_Quality_and_Export.py   # UI-08: Data hygiene inspector & CSV/Excel export
├── src/
│   ├── __init__.py
│   ├── ui.py                          # Reusable Tailwind CSS components (headers, KPI cards)
│   ├── data/
│   │   ├── __init__.py
│   │   ├── loader.py                  # Cached SQLite data loading (@st.cache_data)
│   │   └── cleaner.py                 # Imputation, outlier handling, and typing
│   ├── analytics/
│   │   ├── __init__.py
│   │   ├── descriptive.py             # KPIs, moving averages, branch breakdowns
│   │   ├── lead_time.py               # Discretization buckets & decay calculations
│   │   ├── specialty.py               # Specialty cross-tabs & fee schedules
│   │   ├── demographics.py            # Age cohorts, gender ratios, transit bands
│   │   ├── temporal.py                # 2D day/time pivot matrices
│   │   └── financial.py               # Revenue loss & doctor-hour waste formulas
│   ├── stats/
│   │   ├── __init__.py
│   │   └── hypothesis.py              # Z-tests, Chi-square test, odds ratios
│   ├── models/
│   │   ├── __init__.py
│   │   └── risk_engine.py             # Heuristic multi-factor risk scoring engine
│   └── ui/
│       ├── __init__.py
│       └── charts.py                  # Plotly white-template reusable chart factory
├── tests/
│   ├── __init__.py
│   ├── test_data.py                   # TC-T-01 to TC-T-06: Schema, PK, FK integrity
│   ├── test_analysis.py               # TC-T-07 to TC-T-11: KPI formulas & stats
│   └── test_app.py                    # TC-T-12: Config, imports, and packaging
└── docs/                              # Project Word specifications (Steps 01-14)
```

---

## 3. Step-by-Step Build Tasks (`BUILD-01` to `BUILD-18`)

### `BUILD-01`: Environment Setup, Dependencies & Streamlit Configuration
- **Goal:** Establish deterministic Python environment, pin package dependencies, and set corporate UI theme.
- **Files:** `requirements.txt`, `.streamlit/config.toml`
- **User Stories & Features:** Foundational setup for `US-01` to `US-14`.
- **Exact AI Prompt:**
  ```text
  Create requirements.txt pinning exact packages for Python 3.12 (streamlit, pandas, numpy, scipy, statsmodels, plotly, pytest, openpyxl, python-docx). Create .streamlit/config.toml configuring primaryColor="#1F497D", backgroundColor="#F8FAFC", secondaryBackgroundColor="#FFFFFF", textColor="#1E293B", and headless server mode.
  ```
- **How to Check:** Run `python -c "import openpyxl, docx; print('OK')"` and verify `.streamlit/config.toml` exists.
- **Tests That Must Pass:** `tests/test_app.py::test_tc_t_12_streamlit_config_exists`

---

### `BUILD-02`: Database Ingestion, Caching & Data Cleaning Engine
- **Goal:** Implement high-performance data loader reading SQLite `vw_appointment_analytics` with `@st.cache_data` and clean missing values/outliers.
- **Files:** `src/data/loader.py`, `src/data/cleaner.py`
- **User Stories & Features:** `US-01`, `US-02` | `FR-01` | `BR-01`
- **Exact AI Prompt:**
  ```text
  Write src/data/loader.py to connect to data/app.db, query vw_appointment_analytics using sqlite3, and cache the resulting DataFrame using @st.cache_data(ttl=3600). Write src/data/cleaner.py to handle missing values in reminder_lead_hours (impute with median), downcast categorical columns to Pandas 'category', and flag outliers (e.g. age > 100, lead time > 60).
  ```
- **How to Check:** In Python, call `load_data()` and verify it returns a 300-row DataFrame with 36 columns and zero nulls in required fields.
- **Tests That Must Pass:** `tests/test_data.py::test_tc_t_01_fact_appointments_volume`, `tests/test_data.py::test_tc_t_02_analytics_view_columns`

---

### `BUILD-03`: Descriptive, Lead-Time & Financial Analytics Modules
- **Goal:** Implement core business logic for KPI summaries, lead-time decay curve bucketing, and financial revenue leakage accounting.
- **Files:** `src/analytics/descriptive.py`, `src/analytics/lead_time.py`, `src/analytics/financial.py`
- **User Stories & Features:** `US-03`, `US-04`, `US-05`, `US-12` | `FR-01`, `FR-02`, `FR-09` | `BR-01`, `BR-02`, `BR-09`
- **Exact AI Prompt:**
  ```text
  Implement src/analytics/descriptive.py to compute Total Appointments, Total Shows, Total No-Shows, Attendance Rate %, No-Show Rate %, and 30-day moving averages. Implement src/analytics/lead_time.py to bin lead times into 5 intervals ('0-2 days', '3-7 days', '8-14 days', '15-30 days', '>30 days') and calculate group show rates. Implement src/analytics/financial.py to sum uncollected consultation fees (in INR) and doctor consultation hours wasted.
  ```
- **How to Check:** Verify functions against mock inputs and confirm total revenue loss matches `SUM(actual_consultation_fee * is_no_show)`.
- **Tests That Must Pass:** `tests/test_analysis.py::test_tc_t_07_kpi_formulas_on_known_mock`, `tests/test_analysis.py::test_tc_t_08_financial_and_capacity_formulas_on_known_data`, `tests/test_analysis.py::test_tc_t_10_lead_time_decay_bucketing_logic`

---

### `BUILD-04`: Statistical Inference & Hypothesis Testing Engine
- **Goal:** Implement formal parametric and non-parametric statistical tests (Z-test, Chi-square, odds ratios).
- **Files:** `src/stats/hypothesis.py`
- **User Stories & Features:** `US-09`, `US-10` | `FR-03`, `FR-05` | `BR-03`, `BR-05`
- **Exact AI Prompt:**
  ```text
  Write src/stats/hypothesis.py to calculate a Two-Sample Proportion Z-Test comparing no-show rates between SMS reminder (1) and No-Reminder (0) cohorts, returning z_stat, p_value, relative risk reduction %, and odds ratio with 95% Wald CI. Implement Chi-Square Test of Independence for specialty attendance distributions returning chi2_stat, p_value, and degrees of freedom.
  ```
- **How to Check:** Call functions on `data/app.db` records and assert reminder lift has $p < 0.05$ and $Z > 2.5$.
- **Tests That Must Pass:** `tests/test_analysis.py::test_tc_t_09_two_sample_proportion_z_test_formula`

---

### `BUILD-05`: Predictive Risk Scoring & Triage Engine
- **Goal:** Compute multi-factor heuristic risk scores and assign appointments to High, Medium, and Low risk tiers.
- **Files:** `src/models/risk_engine.py`
- **User Stories & Features:** `US-11`, `US-13` | `FR-08`, `FR-10` | `BR-08`, `BR-10`
- **Exact AI Prompt:**
  ```text
  Write src/models/risk_engine.py to compute normalized risk_score (0.0 to 100.0) based on weighted factors: lead time (+0.25 for >14d), reminder status (-0.20 if sent), specialty risk (+0.15 for mental wellness/dermatology), and repeat history (+0.35 for habitual). Map scores into risk_tier: >=60.0 -> 'High', 30.0-59.9 -> 'Medium', <30.0 -> 'Low'. Generate prioritized front-desk calling rosters.
  ```
- **How to Check:** Verify boundary thresholds (e.g., 60.0 -> High, 59.9 -> Medium, 29.9 -> Low).
- **Tests That Must Pass:** `tests/test_analysis.py::test_tc_t_11_risk_scoring_and_tier_thresholds`

---

### `BUILD-06`: Reusable UI Components & Plotly Chart Factory
- **Goal:** Provide responsive Tailwind cards, headers, and standard `plotly_white` chart generators.
- **Files:** `src/ui.py`, `src/ui/charts.py`
- **User Stories & Features:** `US-03`, `US-08` | `FR-01`, `FR-06` | All BRs
- **Exact AI Prompt:**
  ```text
  Write src/ui/charts.py with reusable Plotly functions: plot_rolling_trend(df), plot_lead_time_decay(df), plot_specialty_bars(df), plot_arrival_heatmap(df), plot_pareto_curve(df), and plot_risk_donut(df). Enforce template='plotly_white', brand colors (#1F497D, #2E75B6, #EF4444), and responsive autosizing.
  ```
- **How to Check:** Call each chart generator with sample DataFrame and assert valid `plotly.graph_objects.Figure` is returned.
- **Tests That Must Pass:** `tests/test_app.py::test_tc_t_12_ui_component_module_imports`

---

### `BUILD-07`: Root Application Landing Page (`app.py`)
- **Goal:** Build main navigation shell, global sidebar filters, session state persistence, and overview summary.
- **Files:** `app.py`
- **User Stories & Features:** `US-01`, `US-03`, `US-04` | `FR-01` | `BR-01`
- **Exact AI Prompt:**
  ```text
  Write app.py with st.set_page_config(page_title="ApexCare Analytics", layout="wide"). Initialize global session state for date range, clinic branch, and specialty. Render Tailwind header banner via src.ui.render_tailwind_header(), display 4 core KPI cards, render 30-day rolling attendance trend chart, and provide executive guidance.
  ```
- **How to Check:** Launch `streamlit run app.py` and verify root page renders in browser without errors.
- **Tests That Must Pass:** `tests/test_app.py::test_tc_t_12_streamlit_config_exists`

---

### `BUILD-08`: Page 1 — Overview & Executive KPIs
- **Goal:** Implement multi-branch attendance comparison and historical time-series analytics.
- **Files:** `pages/1_Overview_and_KPIs.py`
- **User Stories & Features:** `US-03`, `US-04` | `FR-01` | `BR-01`
- **Exact AI Prompt:**
  ```text
  Create pages/1_Overview_and_KPIs.py. Ingest cleaned data, apply global sidebar filters, render 4 Tailwind metric cards, render stacked horizontal branch show vs no-show bar chart, and display rolling 30-day moving average trendline.
  ```
- **How to Check:** Navigate to Page 1 in Streamlit; toggle clinic branch filter and verify charts update instantly.

---

### `BUILD-09`: Page 2 — Lead Time & Specialty Breakdown
- **Goal:** Visualize lead-time attendance decay curves and specialty-wise show rates.
- **Files:** `pages/2_Lead_Time_and_Specialty.py`
- **User Stories & Features:** `US-05`, `US-06` | `FR-02`, `FR-03` | `BR-02`, `BR-03`
- **Exact AI Prompt:**
  ```text
  Create pages/2_Lead_Time_and_Specialty.py. Render binned lead-time bar chart with decay trendline, display specialty-level horizontal diverging bar chart, and render department consultation tariff distribution boxplot.
  ```
- **How to Check:** Verify lead-time bins show monotonic no-show increases (>30 days having ~42% no-show).

---

### `BUILD-10`: Page 3 — Demographics & Temporal Arrival Heatmap
- **Goal:** Deliver patient demographic cohort profiling and 2D arrival heatmaps.
- **Files:** `pages/3_Demographics_and_Arrivals.py`
- **User Stories & Features:** `US-07`, `US-08` | `FR-04`, `FR-06` | `BR-04`, `BR-06`
- **Exact AI Prompt:**
  ```text
  Create pages/3_Demographics_and_Arrivals.py. Render age-group clustered bar charts, distance zone boxplots, and a 2D bivariate density heatmap cross-tabulating weekdays (Mon-Sat) with consultation intervals (Morning, Afternoon, Evening).
  ```
- **How to Check:** Inspect heatmap matrix to verify that peak-risk operational cells are visually highlighted.

---

### `BUILD-11`: Page 4 — Statistical Inference Dashboard
- **Goal:** Surface hypothesis testing cards with $p$-values, confidence intervals, and odds ratios.
- **Files:** `pages/4_Statistical_Inference.py`
- **User Stories & Features:** `US-09`, `US-10` | `FR-03`, `FR-05`, `FR-07` | `BR-03`, `BR-05`, `BR-07`
- **Exact AI Prompt:**
  ```text
  Create pages/4_Statistical_Inference.py. Display Two-Sample Proportion Z-Test summary card for SMS reminder lift with p-value and 95% CI. Display Chi-Square Test of Independence table for medical specialties. Display chronic comorbidity relative risk ratio comparison with clinical interpretation badges.
  ```
- **How to Check:** Verify $p$-values match `scipy`/`statsmodels` outputs with $p < 0.05$.

---

### `BUILD-12`: Page 5 — Financial & Capacity Loss Accounting
- **Goal:** Compute uncollected consultation fees and wasted physician doctor-hours.
- **Files:** `pages/5_Financial_and_Capacity.py`
- **User Stories & Features:** `US-12` | `FR-09` | `BR-09`
- **Exact AI Prompt:**
  ```text
  Create pages/5_Financial_and_Capacity.py. Render monthly financial leakage bar chart in INR, display doctor-wise idle consultation hours wasted, and provide department-level revenue recovery opportunity table.
  ```
- **How to Check:** Verify financial summation equals `SUM(actual_consultation_fee * is_no_show)` exactly.

---

### `BUILD-13`: Page 6 — Predictive Risk Triage Queue
- **Goal:** Provide front-desk staff with an advance calling queue sorted by risk score.
- **Files:** `pages/6_Predictive_Risk_Triage.py`
- **User Stories & Features:** `US-11`, `US-13` | `FR-08`, `FR-10` | `BR-08`, `BR-10`
- **Exact AI Prompt:**
  ```text
  Create pages/6_Predictive_Risk_Triage.py. Display risk tier distribution donut chart (High/Medium/Low), render repeat offender Pareto cumulative loss curve, and render an interactive call roster table with search, risk color badges, and CSV export.
  ```
- **How to Check:** Verify High Risk tier shows patients with long lead times, no reminders, or habitual history.

---

### `BUILD-14`: Page 7 — Simulation & Corporate Wellness
- **Goal:** Model corporate shift attendance adherence and simulate calibrated overbooking buffers.
- **Files:** `pages/7_Simulation_and_Corporate.py`
- **User Stories & Features:** `US-14` | `FR-11`, `FR-12` | `BR-11`, `BR-12`
- **Exact AI Prompt:**
  ```text
  Create pages/7_Simulation_and_Corporate.py. Render corporate IT employee checkup attendance bar chart comparing General vs Evening shifts. Implement an interactive overbooking multiplier slider (1.00 to 1.25) calculating projected recovered slots vs simulated lobby wait-time variance.
  ```
- **How to Check:** Adjust slider and verify recovered capacity updates in real time.

---

### `BUILD-15`: Page 8 — Data Quality Audit & Raw Data Explorer
- **Goal:** Surface data hygiene metrics, outlier inspections, and clean data export tools.
- **Files:** `pages/8_Data_Quality_and_Export.py`
- **User Stories & Features:** `US-01`, `US-02` | `FR-01` | `BR-01`
- **Exact AI Prompt:**
  ```text
  Create pages/8_Data_Quality_and_Export.py. Render missing value summary audit table, display outlier inspection cards, provide paginated table explorer with column search, and offer downloadable CSV and Excel buttons for the active dataset slice.
  ```
- **How to Check:** Click download buttons and confirm valid files are saved to browser.

---

### `BUILD-16`: Full Automated Test Suite Execution & Verification
- **Goal:** Execute entire pytest test suite ensuring 100% pass rate.
- **Files:** `tests/test_data.py`, `tests/test_analysis.py`, `tests/test_app.py`
- **Exact AI Prompt:**
  ```text
  Run pytest tests/ -v. Verify that all 14 tests pass with zero errors. Output terminal execution log confirming 100% test pass rate.
  ```
- **How to Check:** Run `python -m pytest tests/` in terminal; expect `14 passed`.

---

### `BUILD-17`: Production Documentation & README Manual
- **Goal:** Provide complete repository setup guide, architecture summary, and user instructions.
- **Files:** `README.md`
- **Exact AI Prompt:**
  ```text
  Write comprehensive README.md with project overview, architecture diagram, installation commands, directory structure, running the application, running pytest test suite, and viva summary.
  ```
- **How to Check:** Verify README renders cleanly in Markdown viewer with clickable file links.

---

### `BUILD-18`: Cloud Deployment Packaging & Production Verification
- **Goal:** Verify production packaging for Streamlit Community Cloud.
- **Files:** `.gitignore`, `requirements.txt`, `app.py`
- **Exact AI Prompt:**
  ```text
  Ensure .gitignore excludes .venv/, __pycache__/, and temporary cache files. Confirm all paths in code are relative to Path(__file__).parent. Verify app starts without external credentials.
  ```
- **How to Check:** Test cold boot with `streamlit run app.py` on clean port.

---

## 4. Definition of Done (DoD)

A build task or sprint milestone is classified as **DONE** only when:
1. **Business Coverage:** All 7 Must-Have Business Requirements (`BR-01`, `BR-02`, `BR-03`, `BR-05`, `BR-07`, `BR-09`, `BR-10`) are 100% functional.
2. **Automated Quality:** All 14 pytest test cases in `tests/` pass with zero failures and zero warnings.
3. **Performance Ceiling:** App pages load in under 2.5 seconds on initial load and under 500ms on cached filter changes.
4. **Zero PII & Governance:** Data is 100% synthetic, anonymized, and HIPAA/NABH compliant.
5. **Multi-Page Stability:** No unhandled tracebacks or blank screens when slicing empty filter selections.
6. **Documentation & Repo:** `README.md` is complete and repository is ready for public cloud deployment.

---

## 5. Standard CLI Execution Commands

### 1. Environment Installation:
```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Generate Synthetic Dataset & SQLite Warehouse:
```bash
python src/generate_data.py
```

### 3. Run Automated Pytest Test Suite:
```bash
python -m pytest tests/ -v
```

### 4. Launch Streamlit Web Application:
```bash
streamlit run app.py
```

---

## 6. Ten Core Viva Examination Questions & Model Answers

1. **Q1: Why is appointment non-attendance formulated as an optimization and risk-scoring problem rather than basic descriptive reporting?**  
   *Answer:* Descriptive reporting is retrospective—it merely records past slot vacancies after revenue and physician hours are already lost. Formulating it as an optimization and predictive risk problem enables us to classify tomorrow's scheduled visits into High, Medium, and Low risk bands, allowing front-desk personnel to intervene proactively via targeted phone reconfirmations before the slot is wasted.

2. **Q2: Why did you choose a Dimensional Snow-Star schema in SQLite with a materialized 36-column analytical view?**  
   *Answer:* Normalizing master dimensions (`dim_clinics`, `dim_specialties`, `dim_doctors`, `dim_patients`) guarantees transactional consistency, prevents data redundancy, and enables foreign key constraint enforcement. The denormalized 36-column view (`vw_appointment_analytics`) enables Streamlit to ingest all analytical attributes in a single fast query, cached in memory via `@st.cache_data` for sub-15ms slicing.

3. **Q3: What statistical test did you choose to evaluate digital reminder efficacy (`OQ-05`), and why?**  
   *Answer:* We implemented a Two-Sample Proportion Z-Test. Since we are comparing two independent binomial proportions (Reminder Sent vs. Not Sent) across a sample size exceeding 100 observations ($np \ge 5$ and $n(1-p) \ge 5$), the sampling distribution of the difference in proportions is normal by the Central Limit Theorem, making the Z-test the most statistically sound parametric test.

4. **Q4: Why was Pearson's Chi-Square Test chosen for specialty attendance analysis (`OQ-03`) instead of ANOVA or t-tests?**  
   *Answer:* Both the independent variable (Medical Specialty with 6 categorical levels) and the dependent outcome (Show vs. No-Show, a binary categorical state) are discrete categorical variables. Pearson's Chi-Square Test of Independence is the standard statistical test to determine whether attendance distribution is independent of clinical specialty across an $R \times C$ contingency table.

5. **Q5: How did you model booking lead time, and what behavioral pattern did the analysis reveal?**  
   *Answer:* Booking lead time was computed as elapsed calendar days: $\text{LeadTimeDays} = \text{AppointmentDate} - \text{BookingDate}$, discretized into 5 standard operational buckets (0–2d, 3–7d, 8–14d, 15–30d, >30d). The analysis revealed an exponential non-attendance decay curve, where same-week bookings experienced ~14% no-shows, whereas bookings exceeding 30 days exhibited a 41.7% no-show rate.

6. **Q6: Why are consultation fees snapshotted in `fact_appointments` rather than joined dynamically from `dim_specialties`?**  
   *Answer:* This prevents Slowly Changing Dimension (SCD) distortion. If a clinical department updates its consultation tariff, joining the current fee from `dim_specialties` would retroactively distort historical revenue loss accounting. Capturing `actual_consultation_fee` at the moment of booking preserves historical financial accounting fidelity.

7. **Q7: What is the purpose of the repeat offender Pareto analysis (`OQ-08`)?**  
   *Answer:* In outpatient operations, the 80/20 rule often applies: a small, concentrated cohort of chronic repeat no-show patients accounts for a disproportionate volume of clinic vacancies. By grouping patients into First-Timers, Occasional Misses, and Habitual Offenders ($\ge 2$ prior misses), receptionists can prioritize high-touch phone outreach specifically on habitual offenders rather than expending limited staff hours calling accidental one-time misses.

8. **Q8: How does your UI architecture integrate Tailwind CSS into Streamlit without style collisions?**  
   *Answer:* Standard Streamlit does not compile Tailwind utility classes. Rather than using `st.markdown(unsafe_allow_html=True)` which can collide with Streamlit's global DOM, we wrap custom semantic HTML and the Tailwind CDN inside `st.components.v1.html(..., height=...)`. This creates an isolated virtual iframe sandbox where modern Tailwind cards, headers, and badges render cleanly.

9. **Q9: How do unit tests in `tests/test_analysis.py` verify mathematical formulas without depending on database state?**  
   *Answer:* Tests such as `test_tc_t_07` and `test_tc_t_08` execute against tiny, hardcoded mock datasets with known analytical answers (e.g., 8 shows and 2 no-shows). This guarantees that unit tests evaluate algorithmic correctness independently of database mutations, ensuring tests are deterministic and reproducible.

10. **Q10: What is the business return on investment (ROI) of deploying this analytics platform for clinic management?**  
    *Answer:* By identifying high-risk appointment cohorts and enabling targeted pre-visit telephone reconfirmation, outpatient facilities can recapture up to 35% of previously lost consultation slots. For a 4-branch clinic network leaking ₹14.5 Lakhs annually per branch, recapturing one-third of wasted capacity recovers over ₹20 Lakhs in annual network consultation revenue while dramatically reducing patient waitlist queues.
