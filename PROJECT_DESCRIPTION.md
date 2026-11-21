# ApexCare Outpatient Appointment No-Show Analysis & Optimization System
## Comprehensive Project Description & Technical Specification

[![Python Version](https://img.shields.io/badge/Python-3.12%20%7C%203.15-blue.svg)](https://www.python.org/)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.38%2B-red.svg)](https://streamlit.io/)
[![Database](https://img.shields.io/badge/Database-SQLite%203-lightgrey.svg)](https://www.sqlite.org/)
[![Pytest](https://img.shields.io/badge/Tests-20%20Passed%20%28100%25%29-brightgreen.svg)](https://pytest.org/)
[![Styling](https://img.shields.io/badge/Styling-Tailwind%20CSS%20CDN-38B2AC.svg)](https://tailwindcss.com/)

---

### Project Identification & Academic Metadata
* **Academic Institution:** ABES Engineering College, Ghaziabad
* **Department:** Department of Master of Computer Applications (MCA)
* **Project ID:** P101
* **Batch ID:** ABES_OND26 MCA 2ND YEAR_DATA ANALYTICS_BATCH 01
* **Specialization:** MCA 2nd Year — Data Analytics Specialization (Capstone Project)
* **Student Developer:** Deepshikha | Roll No: `2500320140047`
* **Faculty Project Mentor:** Dinesh Sereno, Assistant Professor, Dept. of MCA
* **Modeled Clinical Network:** ApexCare Outpatient & Corporate Health Network, Noida (Sectors 62, 18, 137, and Greater Noida)
* **Live Streamlit Application:** [https://2500320140047-appointment-noshow.streamlit.app](https://2500320140047-appointment-noshow.streamlit.app)
* **GitHub Repository:** [https://github.com/deepsikha09/Hcl_training_Project](https://github.com/deepsikha09/Hcl_training_Project)

---

## 1. Executive Summary & Abstract

Patient non-attendance without prior cancellation ("no-shows") is a chronic operational crisis in outpatient clinical medicine. When scheduled consultation slots are abandoned, clinical departments suffer permanent and irrecoverable capacity loss, specialist doctor time is wasted, revenue evaporates, and waiting queues for other patients are artificially prolonged.

Modeled for the **ApexCare Health Network** across four multi-specialty clinical facilities in the Noida/NCR healthcare corridor, this project delivers an enterprise-grade analytics, statistical inference, and proactive decision-support web platform. The system ingests historical and prospective appointment schedules, clinical comorbidities, booking lead times, transit distances, and digital communication logs to diagnose attendance drivers, audit financial leakage, test mathematical hypotheses, and generate proactive telephone triage calling rosters for front-desk reception staff.

### Key Empirical Findings:
1. **Network Attendance Baseline:** **18.33% No-Show Rate** across 300 audited outpatient encounters (55 missed consultations).
2. **Financial Revenue Leakage:** **₹49,500** in uncollected consultation fees within the audited sample alone, scaling to multi-lakh annual leakage per facility.
3. **Physician Capacity Loss:** **23.83 idle specialist doctor-hours** wasted across consultation gaps.
4. **Digital SMS Reminder Efficacy:** Statistically significant **48.84% Relative Risk Reduction (RRR)** in no-shows ($Z = 2.80, p = 0.0052, OR = 2.30$).
5. **Lead-Time Booking Decay:** Monotonic increase in non-attendance from **9.8%** (0–2 days horizon) to **41.7%** (>30 days horizon).
6. **Specialty Vulnerability:** Dermatology (**25.0%**) and Mental Wellness (**22.0%**) experience double the non-attendance rate of Cardiology (**12.5%**).
7. **Corporate Shift Disparity:** General shift employees achieve **82.5%** preventive checkup completion compared to **65.2%** for evening/night shifts.
8. **Calibrated Overbooking Recovery:** A 1.15x overbooking multiplier safely recovers **18 appointment slots** while keeping average lobby wait times under **14 minutes**.

---

## 2. Clinical Operational Context & Problem Formulation

ApexCare operates four distinct facilities in Noida:
* **Sector 62 (ApexCare City Center):** High-volume residential and IT corridor hospital.
* **Sector 18 (ApexCare Express Clinic):** Commercial walk-in hub with rapid consultation turnover.
* **Sector 137 (ApexCare Multi-Specialty Hospital):** Tertiary center catering to multi-specialty referrals.
* **Greater Noida (ApexCare Wellness & Diagnostic Hub):** Preventative health and corporate wellness center.

### Operational Dysfunctions Addressed:
1. **Physician Idle Capacity:** Specialist physicians in Cardiology, Orthopedics, Pediatrics, and Psychiatry sit idle during 20–30 minute unannounced gaps.
2. **Phantom Waiting Lists:** Patients seeking urgent consultations face multi-week waiting lists because slots appear booked on calendar grids, only to be abandoned on the consultation day.
3. **Uncalibrated Overbooking Friction:** Arbitrary double-booking creates waiting lobby congestion, patient hostility, and clinician burnout when all patients arrive.
4. **Chronic Care Disruption:** Corporate IT workers enrolled in occupational health programs skip routine checkups, allowing hypertension and diabetes to progress unmonitored.

---

## 3. System Architecture & 4-Tier Decoupled Design

```
+-----------------------------------------------------------------------------------------+
|                             TIER 1: RELATIONAL PERSISTENCE                              |
|  data/app.db (SQLite 3) -> 5 Tables (dim_clinics, dim_specialties, dim_doctors,        |
|                            dim_patients, fact_appointments) + vw_appointment_analytics  |
+--------------------------------------------+--------------------------------------------+
                                             | SQL Query via sqlite3
                                             v
+-----------------------------------------------------------------------------------------+
|                             TIER 2: INGESTION & CACHING LAYER                           |
|  src/data/loader.py: load_data() -> @st.cache_data(ttl=3600) In-Memory Pandas DataFrame |
|  src/data/cleaner.py: clean_dataset() -> Impute reminder_lead_hours, downcast category  |
+--------------------------------------------+--------------------------------------------+
                                             | Clean Records Stream
                                             v
+-----------------------------------------------------------------------------------------+
|                             TIER 3: ANALYTICS & STATISTICAL ENGINES                     |
|  * src/analytics/descriptive.py  * src/analytics/lead_time.py  * src/analytics/specialty.py|
|  * src/analytics/demographics.py * src/analytics/temporal.py   * src/analytics/financial.py|
|  * src/stats/hypothesis.py (Two-Sample Z-Test, Chi-Square, Odds Ratio with 95% Wald CI) |
|  * src/models/risk_engine.py (Composite Risk Score [0-100] & Triage Calling Queue)     |
+--------------------------------------------+--------------------------------------------+
                                             | Aggregated Metrics & Charts
                                             v
+-----------------------------------------------------------------------------------------+
|                             TIER 4: PRESENTATION & UI ASSEMBLY                          |
|  * app.py: Executive Overview Dashboard & High-Level KPI Summary Cards                  |
|  * pages/1_Overview_and_KPIs.py         * pages/2_Lead_Time_and_Specialty.py            |
|  * pages/3_Demographics_and_Arrivals.py * pages/4_Statistical_Inference.py             |
|  * pages/5_Financial_and_Capacity.py    * pages/6_Predictive_Risk_Triage.py             |
|  * pages/7_Simulation_and_Corporate.py  * pages/8_Data_Quality_and_Export.py            |
|  * src/ui/components.py (Tailwind Cards) * src/ui/charts.py (Plotly Factory)           |
+-----------------------------------------------------------------------------------------+
```

---

## 4. Relational Database Schema (Star Schema)

The embedded SQLite database (`data/app.db`) enforces referential integrity (`PRAGMA foreign_keys = ON`):

| Entity Name | Nature | Rows | Primary Key | Key Attributes |
| :--- | :--- | :---: | :--- | :--- |
| `dim_clinics` | Dimension | 4 | `clinic_id` | `clinic_name`, `location_sector`, `contact_number` |
| `dim_specialties` | Dimension | 6 | `specialty_id` | `specialty_name`, `consultation_fee` (₹600–₹1,500), `slot_duration_minutes` (20–45m) |
| `dim_doctors` | Dimension | 12 | `doctor_id` | `doctor_name`, `specialty_id` (FK), `clinic_id` (FK), `experience_years` |
| `dim_patients` | Dimension | 200 | `patient_id` | `patient_name`, `age`, `gender`, `distance_km`, `hypertension`, `diabetes`, `corporate_shift` |
| `fact_appointments` | Fact | 300 | `appointment_id` | `patient_id` (FK), `doctor_id` (FK), `appointment_date`, `booking_date`, `time_slot`, `sms_reminder_sent`, `status` |
| `vw_appointment_analytics` | View | 300 | — | Denormalized 36-column analytical view pre-joining all dimensional and computed fields |

---

## 5. Statistical Inference & Mathematical Formulations

### 5.1 Two-Sample Proportion Z-Test (SMS Reminder Lift)
Tests the hypothesis that automated SMS reminders reduce appointment non-attendance:
$$\hat{p} = \frac{x_1 + x_2}{n_1 + n_2}, \quad SE = \sqrt{\hat{p}(1-\hat{p})\left(\frac{1}{n_1} + \frac{1}{n_2}\right)}, \quad Z = \frac{p_1 - p_2}{SE}$$
* Two-tailed p-value computed using complementary error function: $p = \text{erfc}(|Z| / \sqrt{2})$.
* **Empirical Result:** $Z = 2.80$, $p = 0.0052$ ($< 0.05$), $RRR = 48.84\%$, $OR = 2.30$ (Statistically Significant).

### 5.2 Multi-Factor Predictive Risk Scoring Algorithm
Calculates an advance propensity risk score ($0.0 \le \text{Score} \le 100.0$):
$$\text{Score} = 10.0 + \Delta_{\text{lead}} + \Delta_{\text{reminder}} + \Delta_{\text{specialty}} + \Delta_{\text{repeat}} + \Delta_{\text{distance}}$$
* **Lead Time Factor:** $>14$ days ($+25$), $8-14$ days ($+15$), $3-7$ days ($+5$).
* **Reminder Factor:** No SMS sent ($+20$).
* **Specialty Factor:** Mental Wellness / Dermatology ($+15$), General Medicine ($+5$).
* **Repeat Offender Factor:** Habitual ($+35$), Occasional ($+15$).
* **Commute Friction:** Distance $>15$ km ($+10$).
* **Operational Triage Tiers:**
  * **High Risk ($\ge 60.0$):** Mandatory morning telephonic reconfirmation.
  * **Medium Risk ($30.0 - 59.9$):** Automated multi-channel reminder dispatch.
  * **Low Risk ($< 30.0$):** Standard confirmation protocol.

---

## 6. Multi-Page Application Feature Guide

1. **`app.py` (Executive Overview Dashboard):**
   * Top-level KPI cards (Total Bookings, Attendance %, Revenue Loss, Idle Hours).
   * Cross-clinic attendance comparison chart and recent activity log.
2. **`pages/1_Overview_and_KPIs.py`:**
   * Branch-level attendance breakdown and 7-day rolling moving averages.
3. **`pages/2_Lead_Time_and_Specialty.py`:**
   * Non-linear attendance decay curves across 5 lead-time intervals and specialty rankings.
4. **`pages/3_Demographics_and_Arrivals.py`:**
   * Age bracket segmentation, gender ratios, distance bands, and 2D arrival heatmaps (Day of Week vs Time Slot).
5. **`pages/4_Statistical_Inference.py`:**
   * Interactive hypothesis testing engine (Z-Test, Chi-Square, Odds Ratio forest plot, comorbidity compliance).
6. **`pages/5_Financial_and_Capacity.py`:**
   * Lost consultation revenue (₹) and wasted doctor hours audit by department and specialist.
7. **`pages/6_Predictive_Risk_Triage.py`:**
   * Priority calling roster for front-desk staff to reconfirm high-risk patient appointments.
8. **`pages/7_Simulation_and_Corporate.py`:**
   * Overbooking capacity recovery simulator and corporate employee shift adherence analysis.
9. **`pages/8_Data_Quality_and_Export.py`:**
   * Data hygiene validator (missing values, duplicates, outliers) with CSV/Excel export.

---

## 7. Verification & Automated Test Suite

Validated with 20 automated Pytest test cases across three test suites (100% pass rate):
* **`tests/test_data.py` (TC-T-01 to TC-T-06):** Verifies row count thresholds, 36-column view integrity, PK uniqueness, FK referential constraints, and value range bounds.
* **`tests/test_analysis.py` (TC-T-07 to TC-T-18):** Verifies KPI arithmetic, financial loss calculations, statistical Z-test routines, lead-time discretization, and risk scoring cutoffs.
* **`tests/test_app.py` (TC-T-12):** Verifies Streamlit configuration files (`.streamlit/config.toml`), module imports, and packaging sanity.

---

## 8. Quickstart & Local Setup Guide

```bash
# 1. Clone the repository
git clone https://github.com/deepsikha09/Hcl_training_Project.git
cd Hcl_training_Project

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run automated tests
python -m pytest -v

# 4. Launch the application
streamlit run app.py
```
Open your browser at `http://localhost:8501`.
