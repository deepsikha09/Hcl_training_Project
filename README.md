# ApexCare Outpatient Appointment No-Show Analysis & Optimization System

[![Python Version](https://img.shields.io/badge/Python-3.12%20%7C%203.15-blue.svg)](https://www.python.org/)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.38%2B-red.svg)](https://streamlit.io/)
[![Database](https://img.shields.io/badge/Database-SQLite%203-lightgrey.svg)](https://www.sqlite.org/)
[![Pytest](https://img.shields.io/badge/Tests-20%20Passed%20%28100%25%29-brightgreen.svg)](https://pytest.org/)
[![Styling](https://img.shields.io/badge/Styling-Tailwind%20CSS%20CDN-38B2AC.svg)](https://tailwindcss.com/)

**Academic Program:** MCA 2nd Year (Data Analytics Specialization), ABES Engineering College  
**Project ID:** P101  
**Batch ID:** ABES_OND26 MCA 2ND YEAR_DATA ANALYTICS_BATCH 01  
**Learner:** Deepshikha | Roll No: `2500320140047`  
**Faculty Mentor:** Dinesh Sereno  
**Modeled Healthcare Network:** ApexCare Health Network, Noida (Sectors 62, 18, 137, Greater Noida)  
**GitHub Repository:** [https://github.com/deepsikha09/Hcl_training_Project](https://github.com/deepsikha09/Hcl_training_Project)  
**Live Web Application:** [https://2500320140047-appointment-noshow.streamlit.app](https://2500320140047-appointment-noshow.streamlit.app)  

---

## 1. Executive Summary & Clinical Context

Outpatient appointment non-attendance ("no-shows") is a chronic operational vulnerability across multi-specialty healthcare networks, directly eroding clinical revenue, inflating physician idle capacity, and delaying patient care pathways. Modeled for ApexCare Health Network across 4 clinic branches in Noida, this enterprise analytics platform diagnoses historical booking patterns, quantifies financial leakage, evaluates digital SMS reminder lift via formal hypothesis testing ($p < 0.05$), and deploys an advance risk-scoring triage engine for front-desk clinical staff.

### Key Operational Findings:
1. **Network No-Show Rate:** 18.33% baseline across 300 audited outpatient encounters (55 missed consultations).
2. **Revenue Leakage:** ₹49,500 uncollected consultation fees and 23.8 specialist consultation hours lost within the sample period.
3. **SMS Reminder Efficacy:** Statistically significant **48.84% Relative Risk Reduction (RRR)** in no-shows ($Z = 2.80, p = 0.0052, OR = 2.30$).
4. **Lead-Time Decay:** Monotonic increase in non-attendance from **9.8%** (0-2 days advance booking) to **41.7%** (>30 days horizon).
5. **Specialty Vulnerabilities:** Dermatology (25.0%) and Mental Wellness (22.0%) display elevated non-attendance relative to Cardiology (12.5%).

---

## 2. System Architecture & Data Flow

```mermaid
flowchart TD
    DB[(SQLite app.db)] --> |vw_appointment_analytics| LOADER[src.data.loader]
    LOADER --> |@st.cache_data| CLEANER[src.data.cleaner]
    CLEANER --> |Clean Ingested Records| ENGINES{Analytics & Risk Engines}
    
    ENGINES --> DESC[src.analytics.descriptive]
    ENGINES --> LEAD[src.analytics.lead_time]
    ENGINES --> STATS[src.stats.hypothesis]
    ENGINES --> FIN[src.analytics.financial]
    ENGINES --> RISK[src.models.risk_engine]
    ENGINES --> TEMP[src.analytics.temporal]
    
    DESC --> UI_FACTORY[src.ui.charts & src.ui.components]
    LEAD --> UI_FACTORY
    STATS --> UI_FACTORY
    FIN --> UI_FACTORY
    RISK --> UI_FACTORY
    TEMP --> UI_FACTORY
    
    UI_FACTORY --> APP[app.py - Executive Dashboard]
    UI_FACTORY --> P1[pages/1_Overview_and_KPIs.py]
    UI_FACTORY --> P2[pages/2_Lead_Time_and_Specialty.py]
    UI_FACTORY --> P3[pages/3_Demographics_and_Arrivals.py]
    UI_FACTORY --> P4[pages/4_Statistical_Inference.py]
    UI_FACTORY --> P5[pages/5_Financial_and_Capacity.py]
    UI_FACTORY --> P6[pages/6_Predictive_Risk_Triage.py]
    UI_FACTORY --> P7[pages/7_Simulation_and_Corporate.py]
    UI_FACTORY --> P8[pages/8_Data_Quality_and_Export.py]
```

---

## 3. Repository Directory Structure

```text
Appointment_No_Show_Analysis/
├── app.py                             # Root application entrypoint & executive overview
├── requirements.txt                   # Pinned package dependencies
├── README.md                          # Production operations manual (this document)
├── BUILD_PLAN.md                      # 18-task build roadmap & AI execution plan
├── .streamlit/
│   └── config.toml                    # Theme (Navy/Slate) & headless server config
├── data/
│   ├── app.db                         # Production SQLite 3 database (5 tables + view)
│   ├── Appointment_No_Show_data.xlsx  # Multi-sheet master dataset workbook
│   ├── dim_clinics.csv                # 4 Clinic branches master CSV
│   ├── dim_specialties.csv            # 6 Medical departments master CSV
│   ├── dim_doctors.csv                # 12 Specialists master CSV
│   ├── dim_patients.csv               # 200 Registered patients master CSV
│   ├── fact_appointments.csv          # 300 Transactional fact records CSV
│   ├── vw_appointment_analytics.csv   # 36-column denormalized analytical view
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
│   ├── ui.py                          # Reusable Tailwind CSS & Chart re-export facade
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
│       ├── components.py              # Tailwind CSS header banners & metric cards
│       └── charts.py                  # Plotly white-template reusable chart factory
├── tests/
│   ├── __init__.py
│   ├── test_data.py                   # TC-T-01 to TC-T-06: Schema, PK, FK integrity
│   ├── test_analysis.py               # TC-T-07 to TC-T-11, TC-T-13 to TC-T-18
│   └── test_app.py                    # TC-T-12: Config, imports, and packaging
└── docs/                              # Formal Academic Project Deliverables (Steps 01-15)
```

---

## 4. Installation & Setup Guide

### Step 1: Clone or Open Repository
```bash
cd "Appointment No Show Analysis"
```

### Step 2: Create and Activate Virtual Environment
```bash
# Windows PowerShell
python -m venv .venv
.venv\Scripts\activate

# macOS / Linux
python3 -m venv .venv
source .venv/bin/activate
```

### Step 3: Install Required Dependencies
```bash
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

---

## 5. Execution & Verification

### Running the Streamlit Multi-Page Web Application
```bash
streamlit run app.py
```
*The interactive dashboard will automatically open in your default browser at `http://localhost:8501`.*

### Running the Automated Pytest Test Suite
```bash
python -m pytest tests/ -v
```
*Expected output: **20 passed in < 1.0s (100% pass rate)**.*

---

## 6. Academic Viva Examination Guide (Top 10 Questions)

1. **Q: What is the primary business problem addressed by this project?**  
   *A:* Outpatient appointment no-shows cause uncollected consultation revenue, physician idle time, and clinical schedule disruption across ApexCare's 4 Noida branches.

2. **Q: How is the database organized?**  
   *A:* A star/snowflake relational schema in SQLite comprising 4 dimension tables (`dim_clinics`, `dim_specialties`, `dim_doctors`, `dim_patients`) and 1 transactional fact table (`fact_appointments`), denormalized via the 36-column view `vw_appointment_analytics`.

3. **Q: What statistical test was used to evaluate reminder efficacy?**  
   *A:* A Two-Sample Proportion Z-Test comparing no-show rates between SMS-reminder and no-reminder cohorts ($Z = 2.80, p = 0.0052$), demonstrating a statistically significant 48.84% relative risk reduction.

4. **Q: What is the lead-time decay phenomenon observed in the data?**  
   *A:* As booking lead time expands from 0-2 days to >30 days, outpatient no-show propensity monotonically increases from 9.8% to 41.7% due to scheduling oblivion and symptom resolution.

5. **Q: How does the predictive risk scoring engine work?**  
   *A:* It calculates a normalized heuristic risk score (0-100) weighting lead time (>14 days: +25), reminder absence (+20), vulnerable specialty (+15), and repeat offender history (+35), mapping scores into High, Medium, and Low risk tiers.

6. **Q: Why was Tailwind CSS embedded into Streamlit via `st.components.v1.html`?**  
   *A:* To deliver enterprise healthcare branding, bespoke typography, and responsive KPI card layouts while retaining Streamlit's native reactive python state management.

7. **Q: How does the calibrated overbooking simulator operate?**  
   *A:* It models an overbooking multiplier (e.g. 1.12x) against historical show rates (81.7%) to calculate reclaimed consultation slots and revenue while estimating lobby wait-time variance.

8. **Q: How was relational integrity validated during development?**  
   *A:* Automated Pytest fixtures execute `PRAGMA foreign_key_check` and assert primary key uniqueness ($total == distinct$) with zero violations.

9. **Q: What role does data imputation play in the cleaning module?**  
   *A:* Missing values in `reminder_lead_hours` (2.3% of records) are dynamically imputed with the median lead time (24.0 hours) to preserve analytical continuity.

10. **Q: How is patient privacy maintained in this project?**  
    *A:* 100% synthetic fictional data was generated with pseudonymous Indian names and synthetic Medical Record Numbers (MRNs), containing zero Protected Health Information (PHI).

---

## 7. License & Academic Attribution
Developed as an academic capstone project (Project ID: **P101**, Batch: **ABES_OND26 MCA 2ND YEAR_DATA ANALYTICS_BATCH 01**) at **ABES Engineering College, Ghaziabad** under the guidance of **Dinesh Sereno**. Free for academic and educational evaluation.
