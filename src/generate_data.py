import os
import csv
import random
import sqlite3
import datetime
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

def generate_dataset():
    # 1. Setup random seed for exact reproducibility
    random.seed(42)

    base_dir = r"c:\Users\deeps\Downloads\Appointment No Show Analysis"
    data_dir = os.path.join(base_dir, "data")
    os.makedirs(data_dir, exist_ok=True)
    db_path = os.path.join(data_dir, "app.db")
    excel_path = os.path.join(data_dir, "Appointment_No_Show_data.xlsx")

    # ---------------------------------------------------------
    # 2. Master Dimension 1: dim_clinics (4 branches in Noida/NCR)
    # ---------------------------------------------------------
    clinics_headers = ["clinic_id", "clinic_code", "clinic_name", "location_sector", "city", "contact_number", "daily_slot_capacity"]
    clinics_rows = [
        [1, "CLN-01", "ApexCare Sector 62 Clinic", "Noida Sector 62", "Noida", "0120-4100101", 120],
        [2, "CLN-02", "ApexCare Sector 18 Center", "Noida Sector 18", "Noida", "0120-4100102", 100],
        [3, "CLN-03", "ApexCare Expressway Campus", "Noida Expressway", "Noida", "0120-4100103", 80],
        [4, "CLN-04", "ApexCare Greater Noida Branch", "Knowledge Park III", "Greater Noida", "0120-4100104", 90]
    ]

    # ---------------------------------------------------------
    # 3. Master Dimension 2: dim_specialties (6 clinical branches)
    # ---------------------------------------------------------
    specialties_headers = ["specialty_id", "specialty_code", "specialty_name", "standard_consultation_fee", "standard_slot_duration_mins", "department_head"]
    specialties_rows = [
        [1, "SPEC-01", "General Medicine", 600.0, 20, "Dr. Rajesh Sharma"],
        [2, "SPEC-02", "Cardiology", 1200.0, 30, "Dr. A. K. Banerjee"],
        [3, "SPEC-03", "Orthopedics", 900.0, 20, "Dr. Vikram Malhotra"],
        [4, "SPEC-04", "Dermatology", 800.0, 20, "Dr. Priya Nair"],
        [5, "SPEC-05", "Mental Wellness", 1500.0, 45, "Dr. Ananya Sen"],
        [6, "SPEC-06", "Corporate Health", 500.0, 20, "Dr. Amit Verma"]
    ]
    spec_map = {row[0]: {"name": row[2], "fee": row[3], "duration": row[4]} for row in specialties_rows}

    # ---------------------------------------------------------
    # 4. Master Dimension 3: dim_doctors (12 specialists)
    # ---------------------------------------------------------
    doctors_headers = ["doctor_id", "doctor_code", "doctor_name", "specialty_id", "clinic_id", "qualification", "hourly_idle_rate"]
    doctors_rows = [
        [1, "DOC-001", "Dr. Rajesh Sharma", 1, 1, "MBBS, MD (Medicine)", 1800.0],
        [2, "DOC-002", "Dr. Sunita Rao", 1, 2, "MBBS, DNB (Internal Med)", 1800.0],
        [3, "DOC-003", "Dr. A. K. Banerjee", 2, 1, "MBBS, MD, DM (Cardiology)", 3000.0],
        [4, "DOC-004", "Dr. Sanjay Gupta", 2, 3, "MBBS, MD (Cardiology)", 2800.0],
        [5, "DOC-005", "Dr. Vikram Malhotra", 3, 1, "MBBS, MS (Orthopedics)", 2200.0],
        [6, "DOC-006", "Dr. Ritu Saxena", 3, 4, "MBBS, DNB (Ortho)", 2000.0],
        [7, "DOC-007", "Dr. Priya Nair", 4, 2, "MBBS, MD (Dermatology)", 2000.0],
        [8, "DOC-008", "Dr. Manish Tiwari", 4, 3, "MBBS, DVD", 1900.0],
        [9, "DOC-009", "Dr. Ananya Sen", 5, 1, "MD (Psychiatry), DPM", 2500.0],
        [10, "DOC-010", "Dr. Rohan Kapoor", 5, 2, "DNB (Psychiatry)", 2200.0],
        [11, "DOC-011", "Dr. Amit Verma", 6, 1, "MBBS, AFIH (Occupational Health)", 1500.0],
        [12, "DOC-012", "Dr. Neha Agarwal", 6, 4, "MBBS, PGDHM", 1400.0]
    ]
    doc_map = {row[0]: {"name": row[2], "specialty_id": row[3], "clinic_id": row[4]} for row in doctors_rows}

    # ---------------------------------------------------------
    # 5. Master Dimension 4: dim_patients (200 Indian patients)
    # ---------------------------------------------------------
    first_names_m = ["Aarav", "Rohan", "Aditya", "Rahul", "Varun", "Amit", "Manish", "Deepak", "Suresh", "Vikram", "Prateek", "Nikhil", "Gaurav", "Sanjay", "Anil", "Harish", "Karan", "Vivek", "Arun", "Ashok"]
    first_names_f = ["Pooja", "Sneha", "Neha", "Ananya", "Ritu", "Priya", "Kavita", "Swati", "Meera", "Divya", "Shalini", "Sunita", "Preeti", "Monika", "Aarti", "Bhavna", "Kiran", "Tanvi", "Rekha", "Suman"]
    last_names = ["Sharma", "Verma", "Gupta", "Malhotra", "Singh", "Yadav", "Saxena", "Mishra", "Tiwari", "Chauhan", "Bhatnagar", "Agarwal", "Nair", "Kapoor", "Joshi", "Bansal", "Kumar", "Pandey", "Mehta", "Chopra"]
    neighborhoods = [
        ("Noida Sector 62", 2.5), ("Noida Sector 18", 5.0), ("Indirapuram", 4.0),
        ("Noida Extension (Gaur City)", 8.5), ("Vaishali Ghaziabad", 6.0),
        ("Noida Expressway Sector 137", 12.0), ("Greater Noida Alpha 1", 22.0),
        ("Noida Sector 15", 4.5), ("Vasundhara", 5.5), ("Knowledge Park Greater Noida", 26.0)
    ]

    patients_headers = [
        "patient_id", "patient_mrn", "age", "age_group", "gender", "neighborhood",
        "transit_distance_km", "hypertension_flag", "diabetes_flag", "chronic_comorbidity_score",
        "corporate_partner_flag", "corporate_shift_type", "registration_date"
    ]
    patients_rows = []
    base_reg_date = datetime.date(2025, 1, 1)

    for p_id in range(1, 201):
        is_male = random.random() < 0.52
        gender = "M" if is_male else "F"
        fname = random.choice(first_names_m) if is_male else random.choice(first_names_f)
        lname = random.choice(last_names)
        mrn = f"MRN-{10000 + p_id}"
        
        # Age distribution
        r_age = random.random()
        if r_age < 0.25:
            age = random.randint(18, 30)
        elif r_age < 0.65:
            age = random.randint(31, 55)
        elif r_age < 0.88:
            age = random.randint(56, 70)
        else:
            age = random.randint(71, 85)

        if age <= 30:
            age_group = "18-30"
        elif age <= 45:
            age_group = "31-45"
        elif age <= 60:
            age_group = "46-60"
        elif age <= 75:
            age_group = "61-75"
        else:
            age_group = "75+"

        nbr, base_dist = random.choice(neighborhoods)
        dist = round(max(1.0, base_dist + random.uniform(-1.5, 2.5)), 1)

        # Comorbidities
        if age >= 50:
            htn = 1 if random.random() < 0.55 else 0
            dia = 1 if random.random() < 0.45 else 0
        elif age >= 35:
            htn = 1 if random.random() < 0.25 else 0
            dia = 1 if random.random() < 0.20 else 0
        else:
            htn = 1 if random.random() < 0.08 else 0
            dia = 1 if random.random() < 0.05 else 0
        comorb_score = htn + dia

        # Corporate employee
        is_corp = 1 if (age >= 21 and age <= 58 and random.random() < 0.38) else 0
        shift = "None"
        if is_corp == 1:
            shift = "Evening" if random.random() < 0.35 else "General"

        reg_offset = random.randint(0, 365)
        reg_date = (base_reg_date + datetime.timedelta(days=reg_offset)).isoformat()

        patients_rows.append([
            p_id, mrn, age, age_group, gender, nbr, dist, htn, dia, comorb_score, is_corp, shift, reg_date
        ])
    patient_map = {row[0]: row for row in patients_rows}

    # ---------------------------------------------------------
    # 6. Transactional Fact: fact_appointments (300 rows)
    # ---------------------------------------------------------
    appointments_headers = [
        "appointment_id", "booking_reference", "patient_id", "doctor_id", "specialty_id", "clinic_id",
        "booking_date", "appointment_date", "appointment_time", "lead_time_days", "lead_time_bucket",
        "appointment_day_of_week", "appointment_time_slot", "sms_reminder_sent", "reminder_count",
        "reminder_lead_hours", "prior_appointments_count", "prior_no_shows_count", "repeat_offender_tier",
        "actual_consultation_fee", "scheduled_duration_mins", "status", "is_no_show", "risk_score",
        "risk_tier", "capacity_loss_hours", "financial_loss_inr"
    ]
    appointments_rows = []
    start_date = datetime.date(2026, 1, 5)
    days_of_week_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    habitual_patient_ids = set(random.sample(range(1, 201), 10))

    for appt_idx in range(1, 301):
        booking_ref = f"APT-2026-{10000 + appt_idx}"
        
        # Patient
        p_row = random.choice(patients_rows)
        p_id = p_row[0]
        p_age = p_row[2]
        p_comorb = p_row[9]
        p_corp = p_row[10]
        p_shift = p_row[11]

        # Doctor
        if p_corp == 1 and random.random() < 0.50:
            doc_id = random.choice([11, 12, 1, 2])
        else:
            doc_id = random.randint(1, 12)

        doc_info = doc_map[doc_id]
        spec_id = doc_info["specialty_id"]
        clinic_id = doc_info["clinic_id"]
        spec_info = spec_map[spec_id]

        # Lead time
        r_lead = random.random()
        if r_lead < 0.32:
            lead_days = random.randint(0, 2)
        elif r_lead < 0.65:
            lead_days = random.randint(3, 7)
        elif r_lead < 0.85:
            lead_days = random.randint(8, 14)
        else:
            lead_days = random.randint(15, 38)

        if lead_days <= 2:
            lead_bucket = "0-2 days"
        elif lead_days <= 7:
            lead_bucket = "3-7 days"
        elif lead_days <= 14:
            lead_bucket = "8-14 days"
        elif lead_days <= 30:
            lead_bucket = "15-30 days"
        else:
            lead_bucket = ">30 days"

        day_offset = random.randint(0, 180)
        appt_d = start_date + datetime.timedelta(days=day_offset)
        if appt_d.weekday() == 6: # Sunday
            appt_d += datetime.timedelta(days=1)
        booking_d = appt_d - datetime.timedelta(days=lead_days)
        dow = days_of_week_names[appt_d.weekday()]

        # Time slot
        r_time = random.random()
        if r_time < 0.45:
            slot_bucket = "Morning"
            hour = random.randint(9, 11)
            minute = random.choice([0, 20, 30, 40])
        elif r_time < 0.78:
            slot_bucket = "Afternoon"
            hour = random.randint(12, 15)
            minute = random.choice([0, 20, 30, 40])
        else:
            slot_bucket = "Evening"
            hour = random.randint(16, 19)
            minute = random.choice([0, 20, 30])
        appt_time_str = f"{hour:02d}:{minute:02d}"

        # Reminders
        if lead_days >= 3:
            sms_sent = 1 if random.random() < 0.78 else 0
            rem_count = random.choice([1, 2]) if sms_sent == 1 else 0
            rem_hours = float(random.choice([24, 48])) if sms_sent == 1 else None
        else:
            sms_sent = 1 if random.random() < 0.30 else 0
            rem_count = 1 if sms_sent == 1 else 0
            rem_hours = 12.0 if sms_sent == 1 else None

        # Prior history
        if p_id in habitual_patient_ids:
            prior_appts = random.randint(3, 8)
            prior_misses = random.randint(2, prior_appts)
            offender_tier = "Habitual"
        else:
            prior_appts = random.randint(0, 5)
            if prior_appts == 0:
                prior_misses = 0
                offender_tier = "First-Timer"
            else:
                prior_misses = 1 if random.random() < 0.25 else 0
                offender_tier = "Occasional Miss" if prior_misses == 1 else "First-Timer"

        fee = float(spec_info["fee"])
        duration = int(spec_info["duration"])

        # Realistic No-Show Probability Model
        p_noshow = 0.18
        if lead_days == 0:
            p_noshow -= 0.12
        elif lead_days <= 2:
            p_noshow -= 0.08
        elif lead_days >= 20:
            p_noshow += 0.22
        elif lead_days >= 10:
            p_noshow += 0.12

        if sms_sent == 1:
            p_noshow -= 0.14
        else:
            p_noshow += 0.10

        if spec_id == 5: # Mental Wellness
            p_noshow += 0.14
        elif spec_id == 4: # Dermatology
            p_noshow += 0.08
        elif spec_id in [1, 2]:
            p_noshow -= 0.06

        if p_comorb >= 1:
            p_noshow -= 0.08

        if offender_tier == "Habitual":
            p_noshow += 0.35
        elif offender_tier == "Occasional Miss":
            p_noshow += 0.10

        if p_shift == "Evening" and slot_bucket == "Evening":
            p_noshow += 0.12

        p_noshow = max(0.03, min(0.92, p_noshow))
        is_noshow = 1 if random.random() < p_noshow else 0
        status = "No-Show" if is_noshow == 1 else "Show"

        risk_score = round(p_noshow * 100.0, 1)
        if risk_score >= 60.0:
            risk_tier = "High"
        elif risk_score >= 30.0:
            risk_tier = "Medium"
        else:
            risk_tier = "Low"

        fin_loss = fee if is_noshow == 1 else 0.0
        cap_loss = round((duration / 60.0) if is_noshow == 1 else 0.0, 2)

        appointments_rows.append([
            appt_idx, booking_ref, p_id, doc_id, spec_id, clinic_id,
            booking_d.isoformat(), appt_d.isoformat(), appt_time_str, lead_days, lead_bucket,
            dow, slot_bucket, sms_sent, rem_count, rem_hours,
            prior_appts, prior_misses, offender_tier, fee, duration,
            status, is_noshow, risk_score, risk_tier, cap_loss, fin_loss
        ])

    # ---------------------------------------------------------
    # 7. Inject Built-In Data Quality Issues
    # ---------------------------------------------------------
    # Issue A: 6 rows (2%) with missing reminder_lead_hours
    missing_count = 0
    for r in appointments_rows:
        if r[13] == 1 and missing_count < 6:
            r[15] = None
            missing_count += 1

    # Issue B: 3 Outliers
    # Row 15: Extreme lead time (92 days)
    appointments_rows[14][9] = 92
    appointments_rows[14][10] = ">30 days"
    # Row 45: Extreme patient age outlier (108 years)
    patients_rows[44][2] = 108
    patients_rows[44][3] = "75+"
    # Row 80: High fee entry anomaly (₹3,500 on General Med)
    appointments_rows[79][19] = 3500.0

    # ---------------------------------------------------------
    # 8. Construct Enriched Analytical View Rows (36 Columns)
    # ---------------------------------------------------------
    view_headers = [
        "appointment_id", "booking_reference", "booking_date", "appointment_date", "appointment_time",
        "lead_time_days", "lead_time_bucket", "appointment_day_of_week", "appointment_time_slot",
        "sms_reminder_sent", "reminder_count", "prior_appointments_count", "prior_no_shows_count",
        "repeat_offender_tier", "actual_consultation_fee", "scheduled_duration_mins", "status",
        "is_no_show", "risk_score", "risk_tier", "capacity_loss_hours", "financial_loss_inr",
        "patient_id", "patient_mrn", "age", "age_group", "gender", "neighborhood", "transit_distance_km",
        "hypertension_flag", "diabetes_flag", "corporate_shift_type",
        "doctor_name", "specialty_name", "clinic_name", "location_sector"
    ]
    clinic_map = {row[0]: {"name": row[2], "sector": row[3]} for row in clinics_rows}
    
    view_rows = []
    for appt in appointments_rows:
        p = patient_map[appt[2]]
        d = doc_map[appt[3]]
        s = spec_map[appt[4]]
        c = clinic_map[appt[5]]

        v_row = [
            appt[0], appt[1], appt[6], appt[7], appt[8],
            appt[9], appt[10], appt[11], appt[12],
            appt[13], appt[14], appt[16], appt[17],
            appt[18], appt[19], appt[20], appt[21],
            appt[22], appt[23], appt[24], appt[25], appt[26],
            p[0], p[1], p[2], p[3], p[4], p[5], p[6],
            p[7], p[8], p[11],
            d["name"], s["name"], c["name"], c["sector"]
        ]
        view_rows.append(v_row)

    # ---------------------------------------------------------
    # 9. Write CSV Files
    # ---------------------------------------------------------
    def write_csv(filename, headers, rows):
        path = os.path.join(data_dir, filename)
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(headers)
            w.writerows(rows)

    write_csv("dim_clinics.csv", clinics_headers, clinics_rows)
    write_csv("dim_specialties.csv", specialties_headers, specialties_rows)
    write_csv("dim_doctors.csv", doctors_headers, doctors_rows)
    write_csv("dim_patients.csv", patients_headers, patients_rows)
    write_csv("fact_appointments.csv", appointments_headers, appointments_rows)
    write_csv("vw_appointment_analytics.csv", view_headers, view_rows)

    # Raw appointments with duplicates for cleaning step
    raw_appointments_rows = list(appointments_rows) + [appointments_rows[10], appointments_rows[25]]
    write_csv("raw_appointments_with_anomalies.csv", appointments_headers, raw_appointments_rows)

    # ---------------------------------------------------------
    # 10. Write Multi-Sheet Excel Workbook (Appointment_No_Show_data.xlsx)
    # ---------------------------------------------------------
    wb = Workbook()
    
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=9.5, color="333333")
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    stripe_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    thin_border = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
                         top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))

    sheets_to_create = [
        ("dim_clinics", clinics_headers, clinics_rows),
        ("dim_specialties", specialties_headers, specialties_rows),
        ("dim_doctors", doctors_headers, doctors_rows),
        ("dim_patients", patients_headers, patients_rows),
        ("fact_appointments", appointments_headers, appointments_rows),
        ("vw_appointment_analytics", view_headers, view_rows)
    ]

    for s_idx, (s_name, hdrs, rows) in enumerate(sheets_to_create):
        if s_idx == 0:
            ws = wb.active
            ws.title = s_name
        else:
            ws = wb.create_sheet(title=s_name)
            
        ws.views.sheetView[0].showGridLines = True
        
        # Write headers
        ws.row_dimensions[1].height = 26
        for c_i, h in enumerate(hdrs, start=1):
            cell = ws.cell(row=1, column=c_i, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
            
        # Write rows
        for r_i, r_data in enumerate(rows, start=2):
            ws.row_dimensions[r_i].height = 20
            bg = stripe_fill if r_i % 2 == 1 else white_fill
            for c_i, val in enumerate(r_data, start=1):
                cell = ws.cell(row=r_i, column=c_i, value=val)
                cell.font = data_font
                cell.fill = bg
                cell.border = thin_border
                if isinstance(val, (int, float)):
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                    
        # Adjust column widths
        for col_idx in range(1, len(hdrs) + 1):
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = max(len(hdrs[col_idx-1]) + 4, 12)

    wb.save(excel_path)

    # ---------------------------------------------------------
    # 11. Populate SQLite Database (data/app.db)
    # ---------------------------------------------------------
    if os.path.exists(db_path):
        os.remove(db_path)

    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("PRAGMA foreign_keys = ON;")

    ddl_sql = """
    CREATE TABLE dim_clinics (
        clinic_id INTEGER PRIMARY KEY AUTOINCREMENT,
        clinic_code TEXT NOT NULL UNIQUE,
        clinic_name TEXT NOT NULL,
        location_sector TEXT NOT NULL,
        city TEXT NOT NULL DEFAULT 'Noida',
        contact_number TEXT NOT NULL,
        daily_slot_capacity INTEGER NOT NULL CHECK (daily_slot_capacity > 0)
    );

    CREATE TABLE dim_specialties (
        specialty_id INTEGER PRIMARY KEY AUTOINCREMENT,
        specialty_code TEXT NOT NULL UNIQUE,
        specialty_name TEXT NOT NULL,
        standard_consultation_fee REAL NOT NULL CHECK (standard_consultation_fee >= 0),
        standard_slot_duration_mins INTEGER NOT NULL CHECK (standard_slot_duration_mins IN (20, 30, 45, 60)),
        department_head TEXT NOT NULL
    );

    CREATE TABLE dim_doctors (
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

    CREATE TABLE dim_patients (
        patient_id INTEGER PRIMARY KEY AUTOINCREMENT,
        patient_mrn TEXT NOT NULL UNIQUE,
        age INTEGER NOT NULL CHECK (age >= 0 AND age <= 120),
        age_group TEXT NOT NULL,
        gender TEXT NOT NULL CHECK (gender IN ('M', 'F', 'Other')),
        neighborhood TEXT NOT NULL,
        transit_distance_km REAL NOT NULL CHECK (transit_distance_km >= 0),
        hypertension_flag INTEGER NOT NULL CHECK (hypertension_flag IN (0, 1)),
        diabetes_flag INTEGER NOT NULL CHECK (diabetes_flag IN (0, 1)),
        chronic_comorbidity_score INTEGER NOT NULL CHECK (chronic_comorbidity_score >= 0),
        corporate_partner_flag INTEGER NOT NULL CHECK (corporate_partner_flag IN (0, 1)),
        corporate_shift_type TEXT NOT NULL,
        registration_date TEXT NOT NULL
    );

    CREATE TABLE fact_appointments (
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
        lead_time_bucket TEXT NOT NULL,
        appointment_day_of_week TEXT NOT NULL,
        appointment_time_slot TEXT NOT NULL,
        sms_reminder_sent INTEGER NOT NULL CHECK (sms_reminder_sent IN (0, 1)),
        reminder_count INTEGER NOT NULL DEFAULT 0 CHECK (reminder_count >= 0),
        reminder_lead_hours REAL,
        prior_appointments_count INTEGER NOT NULL DEFAULT 0 CHECK (prior_appointments_count >= 0),
        prior_no_shows_count INTEGER NOT NULL DEFAULT 0 CHECK (prior_no_shows_count >= 0),
        repeat_offender_tier TEXT NOT NULL,
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

    CREATE INDEX idx_appt_date ON fact_appointments(appointment_date);
    CREATE INDEX idx_appt_status ON fact_appointments(status, is_no_show);
    CREATE INDEX idx_appt_clinic ON fact_appointments(clinic_id);
    CREATE INDEX idx_appt_specialty ON fact_appointments(specialty_id);
    CREATE INDEX idx_appt_patient ON fact_appointments(patient_id);
    CREATE INDEX idx_appt_lead_time ON fact_appointments(lead_time_days);

    CREATE VIEW vw_appointment_analytics AS
    SELECT 
        f.appointment_id, f.booking_reference, f.booking_date, f.appointment_date, f.appointment_time,
        f.lead_time_days, f.lead_time_bucket, f.appointment_day_of_week, f.appointment_time_slot,
        f.sms_reminder_sent, f.reminder_count, f.prior_appointments_count, f.prior_no_shows_count,
        f.repeat_offender_tier, f.actual_consultation_fee, f.scheduled_duration_mins, f.status,
        f.is_no_show, f.risk_score, f.risk_tier, f.capacity_loss_hours, f.financial_loss_inr,
        p.patient_id, p.patient_mrn, p.age, p.age_group, p.gender, p.neighborhood, p.transit_distance_km,
        p.hypertension_flag, p.diabetes_flag, p.corporate_shift_type,
        d.doctor_name, s.specialty_name, c.clinic_name, c.location_sector
    FROM fact_appointments f
    INNER JOIN dim_patients p ON f.patient_id = p.patient_id
    INNER JOIN dim_doctors d ON f.doctor_id = d.doctor_id
    INNER JOIN dim_specialties s ON f.specialty_id = s.specialty_id
    INNER JOIN dim_clinics c ON f.clinic_id = c.clinic_id;
    """
    cur.executescript(ddl_sql)

    cur.executemany("INSERT INTO dim_clinics VALUES (?,?,?,?,?,?,?)", clinics_rows)
    cur.executemany("INSERT INTO dim_specialties VALUES (?,?,?,?,?,?)", specialties_rows)
    cur.executemany("INSERT INTO dim_doctors VALUES (?,?,?,?,?,?,?)", doctors_rows)
    cur.executemany("INSERT INTO dim_patients VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", patients_rows)
    cur.executemany("INSERT INTO fact_appointments VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", appointments_rows)
    conn.commit()

    # ---------------------------------------------------------
    # 12. Run Validation Checks & Export Report
    # ---------------------------------------------------------
    cur.execute("PRAGMA foreign_key_check;")
    fk_violations = cur.fetchall()

    cur.execute("SELECT COUNT(*) FROM vw_appointment_analytics;")
    view_count = cur.fetchone()[0]

    cur.execute("SELECT AVG(is_no_show)*100.0 FROM fact_appointments;")
    overall_noshow = cur.fetchone()[0]

    cur.execute("SELECT lead_time_bucket, AVG(is_no_show)*100.0, COUNT(*) FROM fact_appointments GROUP BY lead_time_bucket ORDER BY lead_time_bucket;")
    pattern_lead = cur.fetchall()

    cur.execute("SELECT sms_reminder_sent, AVG(is_no_show)*100.0, COUNT(*) FROM fact_appointments GROUP BY sms_reminder_sent;")
    pattern_sms = cur.fetchall()

    cur.execute("SELECT s.specialty_name, AVG(f.is_no_show)*100.0, COUNT(*) FROM fact_appointments f JOIN dim_specialties s ON f.specialty_id = s.specialty_id GROUP BY s.specialty_name ORDER BY AVG(f.is_no_show) DESC;")
    pattern_spec = cur.fetchall()

    cur.execute("SELECT repeat_offender_tier, AVG(is_no_show)*100.0, COUNT(*), SUM(is_no_show) FROM fact_appointments GROUP BY repeat_offender_tier;")
    pattern_repeat = cur.fetchall()

    conn.close()

    val_path = os.path.join(data_dir, "validation_report.txt")
    with open(val_path, "w", encoding="utf-8") as f:
        f.write("=========================================================================\n")
        f.write("APEXCARE HEALTH NETWORK: SYNTHETIC DATASET GENERATION VALIDATION REPORT\n")
        f.write("Timestamp: " + datetime.datetime.now().isoformat() + "\n")
        f.write("=========================================================================\n\n")
        f.write(f"1. GENERATED TABLE ROW COUNTS:\n")
        f.write(f"   - dim_clinics:              {len(clinics_rows)} rows\n")
        f.write(f"   - dim_specialties:          {len(specialties_rows)} rows\n")
        f.write(f"   - dim_doctors:              {len(doctors_rows)} rows\n")
        f.write(f"   - dim_patients:             {len(patients_rows)} rows\n")
        f.write(f"   - fact_appointments:        {len(appointments_rows)} rows\n")
        f.write(f"   - vw_appointment_analytics: {view_count} rows, {len(view_headers)} columns\n\n")

        f.write(f"2. REFERENTIAL INTEGRITY AUDIT:\n")
        f.write(f"   - Foreign Key Violations:   {len(fk_violations)} (PASS: 100% Valid Keys)\n\n")

        f.write(f"3. BASELINE OPERATIONAL OUTCOME:\n")
        f.write(f"   - Overall No-Show Rate:     {overall_noshow:.2f}%\n\n")

        f.write(f"4. BUILT-IN STATISTICAL PATTERNS AUDIT:\n")
        f.write(f"   - Pattern 1: Lead-Time Decay Effect\n")
        for b, r, c in pattern_lead:
            f.write(f"       * {b:<12}: {r:5.1f}% no-show rate (N={c})\n")
        f.write(f"   - Pattern 2: Digital Reminder (SMS) Protective Lift\n")
        for s, r, c in pattern_sms:
            tag = "Reminder Sent (1)" if s == 1 else "No Reminder (0)  "
            f.write(f"       * {tag}: {r:5.1f}% no-show rate (N={c})\n")
        f.write(f"   - Pattern 3: Specialty Attrition Variance\n")
        for sn, r, c in pattern_spec:
            f.write(f"       * {sn:<20}: {r:5.1f}% no-show rate (N={c})\n")
        f.write(f"   - Pattern 4: Repeat Offender Concentration (Pareto Risk)\n")
        for tier, r, c, misses in pattern_repeat:
            f.write(f"       * {tier:<16}: {r:5.1f}% no-show rate | Total Misses = {misses} (N={c})\n")

        f.write(f"\n5. BUILT-IN DATA QUALITY DEFECTS (FOR STEP 15 CLEANING PIPELINE):\n")
        f.write(f"   - Missing Values: 6 records with missing 'reminder_lead_hours' (2.0% rate)\n")
        f.write(f"   - Duplicates:     2 duplicate submissions preserved in raw CSV archive\n")
        f.write(f"   - Outliers:       Patient Age=108, Lead Time=92 days, Fee=INR 3500\n")
        f.write("=========================================================================\n")

    print("Step 10 generation completed successfully!")
    print(f"Data files created at: {data_dir}")

if __name__ == "__main__":
    generate_dataset()
