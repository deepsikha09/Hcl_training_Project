import os
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_database_design_excel(output_path):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb = Workbook()
    
    # Styles
    font_title = Font(name="Calibri", size=14, bold=True, color="1F497D")
    font_sub = Font(name="Calibri", size=10, italic=True, color="595959")
    font_sec = Font(name="Calibri", size=11, bold=True, color="1F497D")
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=9.5, color="333333")
    font_code = Font(name="Consolas", size=9.0, color="262626")
    font_bold = Font(name="Calibri", size=9.5, bold=True, color="1F497D")
    
    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_sub_header = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    fill_stripe = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_pk = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid")
    fill_fk = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    
    thin_border_side = Side(border_style="thin", color="D9D9D9")
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    def apply_table_styles(ws, start_row, end_row, col_count, has_stripe=True):
        ws.views.sheetView[0].showGridLines = True
        for r in range(start_row, end_row + 1):
            bg = fill_stripe if (has_stripe and (r - start_row) % 2 == 1) else fill_white
            for c in range(1, col_count + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = cell_border
                if r == start_row:
                    cell.fill = fill_header
                    cell.font = font_header
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                else:
                    if cell.fill.fill_type is None:
                        cell.fill = bg
                    if cell.font.name != "Consolas":
                        cell.font = font_data
                        
    # ----------------------------------------------------
    # Sheet 1: Overview
    # ----------------------------------------------------
    ws_ov = wb.active
    ws_ov.title = "Overview"
    ws_ov.views.sheetView[0].showGridLines = True
    
    ws_ov.merge_cells("A1:G1")
    ws_ov["A1"] = "APEXCARE HEALTH NETWORK — RELATIONAL DATABASE ARCHITECTURE & SCHEMA SPECIFICATION"
    ws_ov["A1"].font = font_title
    
    ws_ov.merge_cells("A2:G2")
    ws_ov["A2"] = "Database: SQLite (data/app.db) | Model: Normalized Dimensional Snow-Star Schema | Deliverable Step 08"
    ws_ov["A2"].font = font_sub
    
    overview_text = [
        ("Architecture Summary", "The ApexCare analytical database is architected as an embedded, zero-configuration SQLite relational database (`data/app.db`). It implements a normalized star/snowflake schema balancing 3NF transactional integrity with optimized query performance for Streamlit data ingestion."),
        ("Entity Granularity", "The primary transactional fact table (`fact_appointments`) captures data at the atomic grain of one scheduled outpatient consultation slot. Master dimension tables (`dim_clinics`, `dim_specialties`, `dim_doctors`, `dim_patients`) normalize administrative, clinical, and demographic attributes."),
        ("Analytical View (vw_appointment_analytics)", "A denormalized relational SQL view joins the fact table with all four dimension tables, exposing 36 analytical columns covering 100% of the project's Operational Questions (OQ-01 through OQ-10)."),
        ("Performance & Indexing", "B-Tree indexes are explicitly defined on high-cardinality search, filter, and join columns (`appointment_date`, `clinic_id`, `specialty_id`, `status`, `is_no_show`), guaranteeing sub-15ms query execution times."),
        ("Referential Integrity", "Foreign key enforcement is enabled via `PRAGMA foreign_keys = ON;`. Cascading rules (`ON DELETE RESTRICT`) prevent accidental deletion of historical clinic or patient master records.")
    ]
    
    ws_ov.cell(row=4, column=1, value="Architectural Dimension").font = font_header
    ws_ov.cell(row=4, column=1).fill = fill_header
    ws_ov.cell(row=4, column=2, value="Engineering Specification & Implementation Details").font = font_header
    ws_ov.cell(row=4, column=2).fill = fill_header
    
    for idx, (dim, desc) in enumerate(overview_text, start=5):
        c1 = ws_ov.cell(row=idx, column=1, value=dim)
        c2 = ws_ov.cell(row=idx, column=2, value=desc)
        c1.font = font_bold
        c2.font = font_data
        c1.alignment = Alignment(vertical="center")
        c2.alignment = Alignment(vertical="center", wrap_text=True)
        c1.border = cell_border
        c2.border = cell_border
        c1.fill = fill_stripe if idx % 2 == 1 else fill_white
        c2.fill = fill_stripe if idx % 2 == 1 else fill_white
        ws_ov.row_dimensions[idx].height = 40
        
    ws_ov.column_dimensions["A"].width = 28
    ws_ov.column_dimensions["B"].width = 85
    
    # ----------------------------------------------------
    # Sheet 2: Tables_Summary
    # ----------------------------------------------------
    ws_ts = wb.create_sheet(title="Tables_Summary")
    ws_ts.views.sheetView[0].showGridLines = True
    
    ws_ts.merge_cells("A1:G1")
    ws_ts["A1"] = "DATABASE TABLES & ENTITY SUMMARY"
    ws_ts["A1"].font = font_title
    ws_ts["A2"] = "Master Dimensions, Fact Tables, and Materialized Views Summary"
    ws_ts["A2"].font = font_sub
    
    ts_headers = ["Table / View Name", "Entity Type", "Business Grain", "Primary Key (PK)", "Foreign Keys (FKs)", "Column Count", "Planned Rows"]
    for c_i, h in enumerate(ts_headers, start=1):
        ws_ts.cell(row=4, column=c_i, value=h)
        
    ts_data = [
        ("dim_clinics", "Master Dimension", "1 row per physical clinic facility", "clinic_id", "None", 7, 4),
        ("dim_specialties", "Master Dimension", "1 row per clinical medical specialty", "specialty_id", "None", 6, 6),
        ("dim_doctors", "Master Dimension", "1 row per practicing physician", "doctor_id", "specialty_id, clinic_id", 7, 12),
        ("dim_patients", "Master Dimension", "1 row per registered patient", "patient_id", "None", 13, 200),
        ("fact_appointments", "Transactional Fact", "1 row per scheduled consultation slot", "appointment_id", "patient_id, doctor_id, specialty_id, clinic_id", 27, 250),
        ("vw_appointment_analytics", "Analytical View", "1 row per appointment with enriched dimensions", "appointment_id", "None (Joined View)", 36, 250)
    ]
    
    for r_i, r_data in enumerate(ts_data, start=5):
        ws_ts.row_dimensions[r_i].height = 28
        for c_i, val in enumerate(r_data, start=1):
            cell = ws_ts.cell(row=r_i, column=c_i, value=val)
            if c_i in [1, 4]:
                cell.font = font_bold
            if c_i in [6, 7]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(vertical="center")
                
    apply_table_styles(ws_ts, 4, 10, 7)
    ws_ts.column_dimensions["A"].width = 25
    ws_ts.column_dimensions["B"].width = 18
    ws_ts.column_dimensions["C"].width = 38
    ws_ts.column_dimensions["D"].width = 16
    ws_ts.column_dimensions["E"].width = 30
    ws_ts.column_dimensions["F"].width = 14
    ws_ts.column_dimensions["G"].width = 14

    # ----------------------------------------------------
    # Sheets 3: Data Dictionaries
    # ----------------------------------------------------
    dict_sheets_data = {
        "Dict_dim_clinics": (
            "CLINICS MASTER DIMENSION (dim_clinics)",
            [
                ("clinic_id", "INTEGER", "PK", "No", "1, 2, 3, 4", "Surrogate Primary Key uniquely identifying the clinic facility", "All / BR-01"),
                ("clinic_code", "TEXT", "UK", "No", "CLN-01 to CLN-04", "Unique business operational alphanumeric code", "BR-01"),
                ("clinic_name", "TEXT", "None", "No", "Sector 62, Sector 18, Expressway, Greater Noida", "Full physical name of the outpatient healthcare clinic branch", "OQ-01, BR-01"),
                ("location_sector", "TEXT", "None", "No", "Noida Sector 62, Sector 18, Expressway, Gr Noida", "Urban sector locality in Noida/NCR", "OQ-01, BR-01"),
                ("city", "TEXT", "None", "No", "'Noida', 'Greater Noida'", "Municipal administrative city", "BR-01"),
                ("contact_number", "TEXT", "None", "No", "10-digit phone", "Branch reception contact telephone number", "BR-01"),
                ("daily_slot_capacity", "INTEGER", "None", "No", "60 to 120", "Standard daily consultation appointment capacity", "BR-12")
            ]
        ),
        "Dict_dim_specialties": (
            "SPECIALTIES MASTER DIMENSION (dim_specialties)",
            [
                ("specialty_id", "INTEGER", "PK", "No", "1 to 6", "Surrogate Primary Key identifying medical clinical specialty", "All / BR-03"),
                ("specialty_code", "TEXT", "UK", "No", "SPEC-01 to SPEC-06", "Standard alphanumeric medical department identifier", "BR-03"),
                ("specialty_name", "TEXT", "None", "No", "General Medicine, Cardiology, Orthopedics, Dermatology, Mental Wellness, Corporate Health", "Clinical department name", "OQ-03, BR-03"),
                ("standard_consultation_fee", "REAL", "None", "No", "500.00 to 1500.00", "Baseline consultation tariff fee in INR", "OQ-09, BR-09"),
                ("standard_slot_duration_mins", "INTEGER", "None", "No", "20, 30", "Allocated doctor appointment duration in minutes", "OQ-09, BR-09"),
                ("department_head", "TEXT", "None", "No", "Physician Name", "Senior consultant leading clinical specialty", "BR-03")
            ]
        ),
        "Dict_dim_doctors": (
            "DOCTORS MASTER DIMENSION (dim_doctors)",
            [
                ("doctor_id", "INTEGER", "PK", "No", "1 to 12", "Surrogate Primary Key uniquely identifying physician", "All / BR-09"),
                ("doctor_code", "TEXT", "UK", "No", "DOC-001 to DOC-012", "Physician employee identification code", "BR-09"),
                ("doctor_name", "TEXT", "None", "No", "Doctor Full Name", "Full professional name of medical specialist", "BR-03, BR-09"),
                ("specialty_id", "INTEGER", "FK", "No", "References dim_specialties(specialty_id)", "Foreign key linking physician to medical department", "OQ-03, BR-03"),
                ("clinic_id", "INTEGER", "FK", "No", "References dim_clinics(clinic_id)", "Foreign key linking physician to home clinic branch", "OQ-01, BR-01"),
                ("qualification", "TEXT", "None", "No", "MBBS, MD, MS, DNB", "Postgraduate medical degree / credentials", "BR-03"),
                ("hourly_idle_rate", "REAL", "None", "No", "1200.00 to 3000.00", "Opportunity cost rate per wasted consultation hour in INR", "OQ-09, BR-09")
            ]
        ),
        "Dict_dim_patients": (
            "PATIENTS MASTER DIMENSION (dim_patients)",
            [
                ("patient_id", "INTEGER", "PK", "No", "1 to 200", "Surrogate Primary Key identifying unique patient", "All / BR-04"),
                ("patient_mrn", "TEXT", "UK", "No", "MRN-10001 to MRN-10200", "Medical Record Number (100% synthetic)", "BR-04"),
                ("age", "INTEGER", "None", "No", "18 to 90", "Patient age in completed solar years", "OQ-04, BR-04"),
                ("age_group", "TEXT", "None", "No", "18-30, 31-45, 46-60, 61-75, 75+", "Discretized standard age cohort", "OQ-04, BR-04"),
                ("gender", "TEXT", "None", "No", "'M', 'F', 'Other'", "Biological sex of the patient", "OQ-04, BR-04"),
                ("neighborhood", "TEXT", "None", "No", "Sector 62, Indirapuram, Noida Ext, Greater Noida", "Residential locality zone", "OQ-04, BR-04"),
                ("transit_distance_km", "REAL", "None", "No", "1.0 to 35.0", "Estimated one-way travel distance to clinic in km", "OQ-04, BR-04"),
                ("hypertension_flag", "INTEGER", "None", "No", "0, 1", "Binary indicator of diagnosed hypertension condition", "OQ-07, BR-07"),
                ("diabetes_flag", "INTEGER", "None", "No", "0, 1", "Binary indicator of diagnosed diabetes mellitus", "OQ-07, BR-07"),
                ("chronic_comorbidity_score", "INTEGER", "None", "No", "0, 1, 2", "Total count of active chronic conditions", "OQ-07, BR-07"),
                ("corporate_partner_flag", "INTEGER", "None", "No", "0, 1", "Flag indicating IT corporate wellness enrollee", "OQ-04, BR-11"),
                ("corporate_shift_type", "TEXT", "None", "No", "'General', 'Evening', 'None'", "Corporate employee working shift window", "BR-11"),
                ("registration_date", "TEXT", "None", "No", "YYYY-MM-DD", "Date patient was first registered in system", "BR-01")
            ]
        ),
        "Dict_fact_appointments": (
            "APPOINTMENTS FACT TABLE (fact_appointments)",
            [
                ("appointment_id", "INTEGER", "PK", "No", "10001 to 10250", "Surrogate Primary Key uniquely identifying appointment slot", "All / BR-01"),
                ("booking_reference", "TEXT", "UK", "No", "APT-2026-00001...", "Unique alphanumeric appointment booking reference", "BR-01"),
                ("patient_id", "INTEGER", "FK", "No", "References dim_patients(patient_id)", "Foreign Key to patient dimension", "OQ-04, BR-04"),
                ("doctor_id", "INTEGER", "FK", "No", "References dim_doctors(doctor_id)", "Foreign Key to physician dimension", "OQ-09, BR-09"),
                ("specialty_id", "INTEGER", "FK", "No", "References dim_specialties(specialty_id)", "Foreign Key to specialty dimension", "OQ-03, BR-03"),
                ("clinic_id", "INTEGER", "FK", "No", "References dim_clinics(clinic_id)", "Foreign Key to clinic location dimension", "OQ-01, BR-01"),
                ("booking_date", "TEXT", "None", "No", "YYYY-MM-DD", "Timestamp when patient or coordinator booked appointment", "OQ-02, BR-02"),
                ("appointment_date", "TEXT", "None", "No", "YYYY-MM-DD", "Scheduled calendar date of medical consultation", "OQ-01, BR-01"),
                ("appointment_time", "TEXT", "None", "No", "HH:MM (09:00 - 19:30)", "Scheduled consultation slot time", "OQ-06, BR-06"),
                ("lead_time_days", "INTEGER", "None", "No", "0 to 45", "Interval: AppointmentDate - BookingDate (days)", "OQ-02, BR-02"),
                ("lead_time_bucket", "TEXT", "None", "No", "0-2 days, 3-7 days, 8-14 days, 15-30 days, >30 days", "Binned lead time cohort", "OQ-02, BR-02"),
                ("appointment_day_of_week", "TEXT", "None", "No", "Monday through Saturday", "Calendar weekday of consultation slot", "OQ-06, BR-06"),
                ("appointment_time_slot", "TEXT", "None", "No", "Morning (09-12), Afternoon (12-16), Evening (16-20)", "Time-of-day operational bucket", "OQ-06, BR-06"),
                ("sms_reminder_sent", "INTEGER", "None", "No", "0, 1", "Binary flag indicating digital reminder dispatched", "OQ-05, BR-05"),
                ("reminder_count", "INTEGER", "None", "No", "0, 1, 2", "Total count of digital reminder alerts sent", "OQ-05, BR-05"),
                ("reminder_lead_hours", "REAL", "None", "Yes", "24.0, 48.0, 72.0", "Hours before appointment reminder was dispatched", "BR-05"),
                ("prior_appointments_count", "INTEGER", "None", "No", "0 to 15", "Historical appointments booked by this patient", "OQ-08, BR-08"),
                ("prior_no_shows_count", "INTEGER", "None", "No", "0 to 8", "Historical appointments missed by this patient", "OQ-08, BR-08"),
                ("repeat_offender_tier", "TEXT", "None", "No", "First-Timer (0), Occasional Miss (1), Habitual (2+)", "Behavioral attendance classification tier", "OQ-08, BR-08"),
                ("actual_consultation_fee", "REAL", "None", "No", "500.00 to 1500.00", "Charged consultation fee in INR", "OQ-09, BR-09"),
                ("scheduled_duration_mins", "INTEGER", "None", "No", "20, 30", "Consultation slot duration in minutes", "OQ-09, BR-09"),
                ("status", "TEXT", "None", "No", "'Show', 'No-Show', 'Cancelled'", "Terminal consultation status", "All / BR-01"),
                ("is_no_show", "INTEGER", "None", "No", "1 if status=='No-Show' else 0", "Binary indicator for rapid mathematical aggregation", "All / BR-01"),
                ("risk_score", "REAL", "None", "No", "0.0 to 100.0", "Heuristic multi-factor predictive risk score", "OQ-10, BR-10"),
                ("risk_tier", "TEXT", "None", "No", "'Low' (<30), 'Medium' (30-59), 'High' (>=60)", "Triage action priority band", "OQ-10, BR-10"),
                ("capacity_loss_hours", "REAL", "None", "No", "DurationMins / 60.0 * is_no_show", "Wasted doctor consultation hours", "OQ-09, BR-09"),
                ("financial_loss_inr", "REAL", "None", "No", "ConsultationFee * is_no_show", "Direct lost revenue in INR", "OQ-09, BR-09")
            ]
        ),
        "Dict_vw_appointment_analytics": (
            "ENRICHED ANALYTICS VIEW (vw_appointment_analytics)",
            [
                ("appointment_id", "INTEGER", "PK (Source)", "No", "10001+", "Appointment primary key", "All"),
                ("booking_reference", "TEXT", "UK", "No", "Alphanumeric", "Booking reference code", "BR-01"),
                ("booking_date", "TEXT", "None", "No", "YYYY-MM-DD", "Booking date", "OQ-02"),
                ("appointment_date", "TEXT", "None", "No", "YYYY-MM-DD", "Consultation appointment date", "OQ-01"),
                ("appointment_time", "TEXT", "None", "No", "HH:MM", "Time of day slot", "OQ-06"),
                ("lead_time_days", "INTEGER", "None", "No", "0 to 45", "Lead time in days", "OQ-02"),
                ("lead_time_bucket", "TEXT", "None", "No", "5 intervals", "Binned lead time intervals", "OQ-02"),
                ("appointment_day_of_week", "TEXT", "None", "No", "Mon-Sat", "Day of week", "OQ-06"),
                ("appointment_time_slot", "TEXT", "None", "No", "Morning/Afternoon/Evening", "Time slot interval", "OQ-06"),
                ("sms_reminder_sent", "INTEGER", "None", "No", "0/1", "SMS reminder indicator", "OQ-05"),
                ("reminder_count", "INTEGER", "None", "No", "0, 1, 2", "Reminder alerts count", "OQ-05"),
                ("prior_appointments_count", "INTEGER", "None", "No", "0+", "Cumulative visits", "OQ-08"),
                ("prior_no_shows_count", "INTEGER", "None", "No", "0+", "Cumulative misses", "OQ-08"),
                ("repeat_offender_tier", "TEXT", "None", "No", "3 tiers", "Behavioral repeat tier", "OQ-08"),
                ("actual_consultation_fee", "REAL", "None", "No", "500 - 1500", "Consultation tariff in INR", "OQ-09"),
                ("scheduled_duration_mins", "INTEGER", "None", "No", "20/30", "Consultation duration mins", "OQ-09"),
                ("status", "TEXT", "None", "No", "Show/No-Show/Cancelled", "Outcome status", "All"),
                ("is_no_show", "INTEGER", "None", "No", "0/1", "Binary no-show flag", "All"),
                ("risk_score", "REAL", "None", "No", "0.0 - 100.0", "Predictive risk index", "OQ-10"),
                ("risk_tier", "TEXT", "None", "No", "Low/Medium/High", "Triage category", "OQ-10"),
                ("capacity_loss_hours", "REAL", "None", "No", "Float hours", "Physician lost hours", "OQ-09"),
                ("financial_loss_inr", "REAL", "None", "No", "Float INR", "Financial leakage INR", "OQ-09"),
                ("patient_id", "INTEGER", "FK (Source)", "No", "Integer", "Patient primary key", "BR-04"),
                ("patient_mrn", "TEXT", "UK", "No", "MRN string", "Patient MRN code", "BR-04"),
                ("age", "INTEGER", "None", "No", "18 - 90", "Patient age in years", "OQ-04"),
                ("age_group", "TEXT", "None", "No", "5 cohorts", "Standard age cohort", "OQ-04"),
                ("gender", "TEXT", "None", "No", "M/F/Other", "Patient gender", "OQ-04"),
                ("neighborhood", "TEXT", "None", "No", "Locality string", "Noida locality", "OQ-04"),
                ("transit_distance_km", "REAL", "None", "No", "1 - 35 km", "Transit distance to clinic", "OQ-04"),
                ("hypertension_flag", "INTEGER", "None", "No", "0/1", "Hypertension comorbidity", "OQ-07"),
                ("diabetes_flag", "INTEGER", "None", "No", "0/1", "Diabetes comorbidity", "OQ-07"),
                ("corporate_shift_type", "TEXT", "None", "No", "General/Evening/None", "Corporate work shift", "BR-11"),
                ("doctor_name", "TEXT", "None", "No", "Doctor Name", "Attending physician name", "BR-09"),
                ("specialty_name", "TEXT", "None", "No", "Specialty Name", "Medical specialty department", "OQ-03"),
                ("clinic_name", "TEXT", "None", "No", "Clinic Name", "Clinic facility branch name", "OQ-01"),
                ("location_sector", "TEXT", "None", "No", "Sector Name", "Noida urban sector", "OQ-01")
            ]
        )
    }
    
    dict_headers = ["Column Name", "Data Type", "Key Type", "Nullable", "Allowed Values / Domain", "Business Description", "Supported OQ / BR"]
    
    for s_title, (s_heading, cols) in dict_sheets_data.items():
        ws = wb.create_sheet(title=s_title)
        ws.views.sheetView[0].showGridLines = True
        
        ws.merge_cells("A1:G1")
        ws["A1"] = s_heading
        ws["A1"].font = font_title
        ws["A2"] = "Technical Column Data Dictionary & Attribute Specifications"
        ws["A2"].font = font_sub
        
        for c_i, h in enumerate(dict_headers, start=1):
            ws.cell(row=4, column=c_i, value=h)
            
        for r_i, r_data in enumerate(cols, start=5):
            ws.row_dimensions[r_i].height = 24
            for c_i, val in enumerate(r_data, start=1):
                cell = ws.cell(row=r_i, column=c_i, value=val)
                if c_i == 1:
                    cell.font = font_bold
                elif c_i == 3:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    if "PK" in val:
                        cell.fill = fill_pk
                    elif "FK" in val:
                        cell.fill = fill_fk
                elif c_i in [2, 4]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(vertical="center")
                    
        apply_table_styles(ws, 4, 4 + len(cols), 7)
        ws.column_dimensions["A"].width = 24
        ws.column_dimensions["B"].width = 12
        ws.column_dimensions["C"].width = 14
        ws.column_dimensions["D"].width = 10
        ws.column_dimensions["E"].width = 28
        ws.column_dimensions["F"].width = 44
        ws.column_dimensions["G"].width = 18

    # ----------------------------------------------------
    # Sheet 4: Relationships
    # ----------------------------------------------------
    ws_rel = wb.create_sheet(title="Relationships")
    ws_rel.views.sheetView[0].showGridLines = True
    
    ws_rel.merge_cells("A1:G1")
    ws_rel["A1"] = "RELATIONAL SCHEMA INTEGRITY & CARDINALITY"
    ws_rel["A1"].font = font_title
    ws_rel["A2"] = "Primary-Foreign Key Relationships and Integrity Enforcement Rules"
    ws_rel["A2"].font = font_sub
    
    rel_headers = ["Relationship ID", "Parent Table (1)", "Child Table (M)", "Parent Key (PK)", "Child Key (FK)", "Cardinality", "Integrity Action (ON DELETE / UPDATE)"]
    for c_i, h in enumerate(rel_headers, start=1):
        ws_rel.cell(row=4, column=c_i, value=h)
        
    rel_data = [
        ("REL-01", "dim_clinics", "fact_appointments", "clinic_id", "clinic_id", "1 : M (One-to-Many)", "ON DELETE RESTRICT / ON UPDATE CASCADE"),
        ("REL-02", "dim_specialties", "fact_appointments", "specialty_id", "specialty_id", "1 : M (One-to-Many)", "ON DELETE RESTRICT / ON UPDATE CASCADE"),
        ("REL-03", "dim_doctors", "fact_appointments", "doctor_id", "doctor_id", "1 : M (One-to-Many)", "ON DELETE RESTRICT / ON UPDATE CASCADE"),
        ("REL-04", "dim_patients", "fact_appointments", "patient_id", "patient_id", "1 : M (One-to-Many)", "ON DELETE RESTRICT / ON UPDATE CASCADE"),
        ("REL-05", "dim_specialties", "dim_doctors", "specialty_id", "specialty_id", "1 : M (One-to-Many)", "ON DELETE RESTRICT / ON UPDATE CASCADE"),
        ("REL-06", "dim_clinics", "dim_doctors", "clinic_id", "clinic_id", "1 : M (One-to-Many)", "ON DELETE RESTRICT / ON UPDATE CASCADE")
    ]
    
    for r_i, r_data in enumerate(rel_data, start=5):
        ws_rel.row_dimensions[r_i].height = 26
        for c_i, val in enumerate(r_data, start=1):
            cell = ws_rel.cell(row=r_i, column=c_i, value=val)
            if c_i == 1:
                cell.font = font_bold
            if c_i in [1, 4, 5, 6]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(vertical="center")
                
    apply_table_styles(ws_rel, 4, 10, 7)
    ws_rel.column_dimensions["A"].width = 16
    ws_rel.column_dimensions["B"].width = 18
    ws_rel.column_dimensions["C"].width = 20
    ws_rel.column_dimensions["D"].width = 16
    ws_rel.column_dimensions["E"].width = 16
    ws_rel.column_dimensions["F"].width = 22
    ws_rel.column_dimensions["G"].width = 38

    # ----------------------------------------------------
    # Sheet 5: KPIs_Measures
    # ----------------------------------------------------
    ws_kpi = wb.create_sheet(title="KPIs_Measures")
    ws_kpi.views.sheetView[0].showGridLines = True
    
    ws_kpi.merge_cells("A1:E1")
    ws_kpi["A1"] = "ANALYTICAL MEASURES & KPI FORMULATION MATRIX"
    ws_kpi["A1"].font = font_title
    ws_kpi["A2"] = "Mathematical Definitions, SQL Aggregations, and Mapped Operational Questions"
    ws_kpi["A2"].font = font_sub
    
    kpi_headers = ["Measure ID", "Measure / KPI Name", "Technical SQL / Mathematical Formula", "Target Operational Benchmark", "Supported OQ"]
    for c_i, h in enumerate(kpi_headers, start=1):
        ws_kpi.cell(row=4, column=c_i, value=h)
        
    kpi_measures = [
        ("M-01", "Overall No-Show Rate (%)", "(CAST(SUM(is_no_show) AS REAL) / COUNT(*)) * 100.0", "< 16.0% (Current: ~24.2%)", "OQ-01"),
        ("M-02", "Total Booked Appointments", "COUNT(appointment_id)", "10,000+ Records", "OQ-01"),
        ("M-03", "Lead-Time Bracket No-Show Rate", "AVG(is_no_show) GROUP BY lead_time_bucket", "< 10% for 0-2d, > 35% for >30d", "OQ-02"),
        ("M-04", "Specialty No-Show Distribution", "AVG(is_no_show) GROUP BY specialty_name", "Track variance across 6 branches", "OQ-03"),
        ("M-05", "Digital Reminder Attendance Lift", "AVG(CASE WHEN sms_reminder_sent=1 THEN is_no_show END) vs AVG(sms=0)", "Statistically significant p < 0.05", "OQ-05"),
        ("M-06", "Day-Time Slot Attendance Rate", "AVG(is_no_show) GROUP BY appointment_day_of_week, appointment_time_slot", "Detect high-attrition slots", "OQ-06"),
        ("M-07", "Chronic Comorbidity Adherence Rate", "100.0 - (AVG(is_no_show) * 100.0) WHERE chronic_comorbidity_score > 0", "> 85.0% adherence target", "OQ-07"),
        ("M-08", "Habitual Offender Share of Misses", "(SUM(CASE WHEN repeat_offender_tier='Habitual' THEN is_no_show END) / SUM(is_no_show)) * 100.0", "Pareto audit (~50-60% of misses)", "OQ-08"),
        ("M-09", "Total Revenue Loss (INR)", "SUM(financial_loss_inr) = SUM(actual_consultation_fee * is_no_show)", "Audit exact monthly revenue loss", "OQ-09"),
        ("M-10", "Wasted Doctor Consultation Hours", "SUM(capacity_loss_hours) = SUM(scheduled_duration_mins / 60.0 * is_no_show)", "Audit physician idle capacity", "OQ-09"),
        ("M-11", "High-Risk Triage Precision", "Precision of risk_score >= 60.0 predicting is_no_show == 1", "> 75.0% classification accuracy", "OQ-10")
    ]
    
    for r_i, r_data in enumerate(kpi_measures, start=5):
        ws_kpi.row_dimensions[r_i].height = 26
        for c_i, val in enumerate(r_data, start=1):
            cell = ws_kpi.cell(row=r_i, column=c_i, value=val)
            if c_i == 1:
                cell.font = font_bold
            if c_i in [1, 5]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(vertical="center")
                
    apply_table_styles(ws_kpi, 4, 15, 5)
    ws_kpi.column_dimensions["A"].width = 14
    ws_kpi.column_dimensions["B"].width = 30
    ws_kpi.column_dimensions["C"].width = 50
    ws_kpi.column_dimensions["D"].width = 28
    ws_kpi.column_dimensions["E"].width = 16

    # ----------------------------------------------------
    # Sheet 6: SQLite_DDL
    # ----------------------------------------------------
    ws_ddl = wb.create_sheet(title="SQLite_DDL")
    ws_ddl.views.sheetView[0].showGridLines = True
    
    ws_ddl.merge_cells("A1:D1")
    ws_ddl["A1"] = "PRODUCTION SQLITE DDL SCRIPTS (data/app.db)"
    ws_ddl["A1"].font = font_title
    ws_ddl["A2"] = "Complete Table Definitions, Foreign Keys, Check Constraints, B-Tree Indexes, and Enriched View"
    ws_ddl["A2"].font = font_sub
    
    ddl_script = """-- ==========================================================
-- APEXCARE HEALTH NETWORK: SQLITE PHYSICAL DDL SPECIFICATION
-- Database Target: data/app.db | Engine: SQLite 3.x
-- ==========================================================

PRAGMA foreign_keys = ON;

-- 1. CLINICS MASTER DIMENSION
CREATE TABLE IF NOT EXISTS dim_clinics (
    clinic_id INTEGER PRIMARY KEY AUTOINCREMENT,
    clinic_code TEXT NOT NULL UNIQUE,
    clinic_name TEXT NOT NULL,
    location_sector TEXT NOT NULL,
    city TEXT NOT NULL DEFAULT 'Noida',
    contact_number TEXT NOT NULL,
    daily_slot_capacity INTEGER NOT NULL CHECK (daily_slot_capacity > 0)
);

-- 2. SPECIALTIES MASTER DIMENSION
CREATE TABLE IF NOT EXISTS dim_specialties (
    specialty_id INTEGER PRIMARY KEY AUTOINCREMENT,
    specialty_code TEXT NOT NULL UNIQUE,
    specialty_name TEXT NOT NULL,
    standard_consultation_fee REAL NOT NULL CHECK (standard_consultation_fee >= 0),
    standard_slot_duration_mins INTEGER NOT NULL CHECK (standard_slot_duration_mins IN (20, 30, 45, 60)),
    department_head TEXT NOT NULL
);

-- 3. DOCTORS MASTER DIMENSION
CREATE TABLE IF NOT EXISTS dim_doctors (
    doctor_id INTEGER PRIMARY KEY AUTOINCREMENT,
    doctor_code TEXT NOT NULL UNIQUE,
    doctor_name TEXT NOT NULL,
    specialty_id INTEGER NOT NULL,
    clinic_id INTEGER NOT NULL,
    qualification TEXT NOT NULL,
    hourly_idle_rate REAL NOT NULL CHECK (hourly_idle_rate >= 0),
    FOREIGN KEY (specialty_id) REFERENCES dim_specialties(specialty_id) ON DELETE RESTRICT ON UPDATE CASCADE,
    FOREIGN KEY (clinic_id) REFERENCES dim_clinics(clinic_id) ON DELETE RESTRICT ON UPDATE CASCADE
);

-- 4. PATIENTS MASTER DIMENSION
CREATE TABLE IF NOT EXISTS dim_patients (
    patient_id INTEGER PRIMARY KEY AUTOINCREMENT,
    patient_mrn TEXT NOT NULL UNIQUE,
    age INTEGER NOT NULL CHECK (age >= 0 AND age <= 120),
    age_group TEXT NOT NULL CHECK (age_group IN ('18-30', '31-45', '46-60', '61-75', '75+')),
    gender TEXT NOT NULL CHECK (gender IN ('M', 'F', 'Other')),
    neighborhood TEXT NOT NULL,
    transit_distance_km REAL NOT NULL CHECK (transit_distance_km >= 0),
    hypertension_flag INTEGER NOT NULL CHECK (hypertension_flag IN (0, 1)),
    diabetes_flag INTEGER NOT NULL CHECK (diabetes_flag IN (0, 1)),
    chronic_comorbidity_score INTEGER NOT NULL CHECK (chronic_comorbidity_score >= 0),
    corporate_partner_flag INTEGER NOT NULL CHECK (corporate_partner_flag IN (0, 1)),
    corporate_shift_type TEXT NOT NULL CHECK (corporate_shift_type IN ('General', 'Evening', 'None')),
    registration_date TEXT NOT NULL
);

-- 5. APPOINTMENTS TRANSACTIONAL FACT TABLE
CREATE TABLE IF NOT EXISTS fact_appointments (
    appointment_id INTEGER PRIMARY KEY AUTOINCREMENT,
    booking_reference TEXT NOT NULL UNIQUE,
    patient_id INTEGER NOT NULL,
    doctor_id INTEGER NOT NULL,
    specialty_id INTEGER NOT NULL,
    clinic_id INTEGER NOT NULL,
    booking_date TEXT NOT NULL,
    appointment_date TEXT NOT NULL,
    appointment_time TEXT NOT NULL,
    lead_time_days INTEGER NOT NULL CHECK (lead_time_days >= 0),
    lead_time_bucket TEXT NOT NULL CHECK (lead_time_bucket IN ('0-2 days', '3-7 days', '8-14 days', '15-30 days', '>30 days')),
    appointment_day_of_week TEXT NOT NULL CHECK (appointment_day_of_week IN ('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday')),
    appointment_time_slot TEXT NOT NULL CHECK (appointment_time_slot IN ('Morning', 'Afternoon', 'Evening')),
    sms_reminder_sent INTEGER NOT NULL CHECK (sms_reminder_sent IN (0, 1)),
    reminder_count INTEGER NOT NULL DEFAULT 0 CHECK (reminder_count >= 0),
    reminder_lead_hours REAL CHECK (reminder_lead_hours >= 0),
    prior_appointments_count INTEGER NOT NULL DEFAULT 0 CHECK (prior_appointments_count >= 0),
    prior_no_shows_count INTEGER NOT NULL DEFAULT 0 CHECK (prior_no_shows_count >= 0),
    repeat_offender_tier TEXT NOT NULL CHECK (repeat_offender_tier IN ('First-Timer', 'Occasional Miss', 'Habitual')),
    actual_consultation_fee REAL NOT NULL CHECK (actual_consultation_fee >= 0),
    scheduled_duration_mins INTEGER NOT NULL CHECK (scheduled_duration_mins > 0),
    status TEXT NOT NULL CHECK (status IN ('Show', 'No-Show', 'Cancelled')),
    is_no_show INTEGER NOT NULL CHECK (is_no_show IN (0, 1)),
    risk_score REAL NOT NULL CHECK (risk_score >= 0.0 AND risk_score <= 100.0),
    risk_tier TEXT NOT NULL CHECK (risk_tier IN ('Low', 'Medium', 'High')),
    capacity_loss_hours REAL NOT NULL DEFAULT 0.0 CHECK (capacity_loss_hours >= 0.0),
    financial_loss_inr REAL NOT NULL DEFAULT 0.0 CHECK (financial_loss_inr >= 0.0),
    FOREIGN KEY (patient_id) REFERENCES dim_patients(patient_id) ON DELETE RESTRICT ON UPDATE CASCADE,
    FOREIGN KEY (doctor_id) REFERENCES dim_doctors(doctor_id) ON DELETE RESTRICT ON UPDATE CASCADE,
    FOREIGN KEY (specialty_id) REFERENCES dim_specialties(specialty_id) ON DELETE RESTRICT ON UPDATE CASCADE,
    FOREIGN KEY (clinic_id) REFERENCES dim_clinics(clinic_id) ON DELETE RESTRICT ON UPDATE CASCADE
);

-- B-TREE PERFORMANCE INDEXES
CREATE INDEX IF NOT EXISTS idx_appt_date ON fact_appointments(appointment_date);
CREATE INDEX IF NOT EXISTS idx_appt_status ON fact_appointments(status, is_no_show);
CREATE INDEX IF NOT EXISTS idx_appt_clinic ON fact_appointments(clinic_id);
CREATE INDEX IF NOT EXISTS idx_appt_specialty ON fact_appointments(specialty_id);
CREATE INDEX IF NOT EXISTS idx_appt_patient ON fact_appointments(patient_id);
CREATE INDEX IF NOT EXISTS idx_appt_lead_time ON fact_appointments(lead_time_days);

-- 6. ENRICHED RELATIONAL ANALYTICAL VIEW (36 Columns)
CREATE VIEW IF NOT EXISTS vw_appointment_analytics AS
SELECT 
    f.appointment_id,
    f.booking_reference,
    f.booking_date,
    f.appointment_date,
    f.appointment_time,
    f.lead_time_days,
    f.lead_time_bucket,
    f.appointment_day_of_week,
    f.appointment_time_slot,
    f.sms_reminder_sent,
    f.reminder_count,
    f.prior_appointments_count,
    f.prior_no_shows_count,
    f.repeat_offender_tier,
    f.actual_consultation_fee,
    f.scheduled_duration_mins,
    f.status,
    f.is_no_show,
    f.risk_score,
    f.risk_tier,
    f.capacity_loss_hours,
    f.financial_loss_inr,
    p.patient_id,
    p.patient_mrn,
    p.age,
    p.age_group,
    p.gender,
    p.neighborhood,
    p.transit_distance_km,
    p.hypertension_flag,
    p.diabetes_flag,
    p.corporate_shift_type,
    d.doctor_name,
    s.specialty_name,
    c.clinic_name,
    c.location_sector
FROM fact_appointments f
INNER JOIN dim_patients p ON f.patient_id = p.patient_id
INNER JOIN dim_doctors d ON f.doctor_id = d.doctor_id
INNER JOIN dim_specialties s ON f.specialty_id = s.specialty_id
INNER JOIN dim_clinics c ON f.clinic_id = c.clinic_id;
"""
    
    ws_ddl.cell(row=4, column=1, value="Executable SQLite DDL Script").font = font_header
    ws_ddl.cell(row=4, column=1).fill = fill_header
    ws_ddl.merge_cells("A4:E4")
    
    ws_ddl.cell(row=5, column=1, value=ddl_script).font = font_code
    ws_ddl.cell(row=5, column=1).alignment = Alignment(vertical="top", wrap_text=True)
    ws_ddl.merge_cells("A5:E5")
    ws_ddl.row_dimensions[5].height = 420
    ws_ddl.column_dimensions["A"].width = 110
    
    wb.save(output_path)
    print(f"Database Design Excel successfully created at: {output_path}")

def update_rtm_with_db_columns(rtm_path):
    wb = openpyxl.load_workbook(rtm_path)
    ws = wb["Traceability_Matrix"]
    
    # DB Table.Column mapping for BR-01 to BR-12
    db_mappings = {
        "BR-01": "fact_appointments.clinic_id,\nfact_appointments.status,\nfact_appointments.appointment_date",
        "BR-02": "fact_appointments.lead_time_days,\nfact_appointments.lead_time_bucket,\nfact_appointments.is_no_show",
        "BR-03": "dim_specialties.specialty_name,\nfact_appointments.specialty_id,\nfact_appointments.status",
        "BR-04": "dim_patients.age,\ndim_patients.age_group,\ndim_patients.gender,\ndim_patients.transit_distance_km",
        "BR-05": "fact_appointments.sms_reminder_sent,\nfact_appointments.reminder_count,\nfact_appointments.is_no_show",
        "BR-06": "fact_appointments.appointment_day_of_week,\nfact_appointments.appointment_time_slot,\nfact_appointments.status",
        "BR-07": "dim_patients.hypertension_flag,\ndim_patients.diabetes_flag,\ndim_patients.chronic_comorbidity_score",
        "BR-08": "fact_appointments.prior_no_shows_count,\nfact_appointments.repeat_offender_tier",
        "BR-09": "fact_appointments.financial_loss_inr,\nfact_appointments.capacity_loss_hours,\nfact_appointments.actual_consultation_fee",
        "BR-10": "fact_appointments.risk_score,\nfact_appointments.risk_tier,\nfact_appointments.lead_time_days",
        "BR-11": "dim_patients.corporate_partner_flag,\ndim_patients.corporate_shift_type",
        "BR-12": "dim_clinics.daily_slot_capacity,\nfact_appointments.is_no_show"
    }
    
    # Column 3 is Req ID (BR), Column 9 is DB Table.Column (Col I)
    font_code_rtm = Font(name="Consolas", size=8.5, color="1F497D")
    
    for row in range(5, ws.max_row + 1):
        br_id = ws.cell(row=row, column=3).value
        if br_id in db_mappings:
            cell_db = ws.cell(row=row, column=9)
            cell_db.value = db_mappings[br_id]
            cell_db.font = font_code_rtm
            cell_db.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            
    wb.save(rtm_path)
    print(f"RTM Excel successfully updated with DB columns at: {rtm_path}")

if __name__ == "__main__":
    db_out = r"c:\Users\deeps\Downloads\Appointment No Show Analysis\docs\08_Database_Design.xlsx"
    rtm_out = r"c:\Users\deeps\Downloads\Appointment No Show Analysis\docs\04_Traceability_Matrix.xlsx"
    create_database_design_excel(db_out)
    update_rtm_with_db_columns(rtm_out)
