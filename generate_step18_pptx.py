"""
Generator script for Step 18 Deliverable:
docs/2500320140047_Presentation.pptx & docs/18_Presentation.pptx
Official ABES Engineering College 11-Slide Capstone Presentation Deck
"""

import os
import shutil
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# Corporate Palette
COLOR_NAVY = RGBColor(31, 73, 125)        # #1F497D
COLOR_ACCENT = RGBColor(46, 117, 182)     # #2E75B6
COLOR_DARK = RGBColor(30, 41, 59)         # #1E293B
COLOR_MUTED = RGBColor(100, 116, 139)     # #64748B
COLOR_CARD_BG = RGBColor(255, 255, 255)   # #FFFFFF
COLOR_BORDER = RGBColor(226, 232, 240)    # #E2E8F0
COLOR_BG = RGBColor(248, 250, 252)        # #F8FAFC
COLOR_GREEN = RGBColor(16, 185, 129)      # #10B981
COLOR_RED = RGBColor(239, 68, 68)         # #EF4444
COLOR_AMBER = RGBColor(245, 158, 11)      # #F59E0B

def add_header(slide, title_text, category_text="ABES ENGINEERING COLLEGE • MCA CAPSTONE 2026"):
    """Adds a standard executive header across content slides."""
    # Top banner bar
    top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.15))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = COLOR_NAVY
    top_bar.line.color.rgb = COLOR_NAVY

    # Category / Breadcrumb
    cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.12), Inches(11.7), Inches(0.3))
    tf_cat = cat_box.text_frame
    tf_cat.word_wrap = True
    p_cat = tf_cat.paragraphs[0]
    p_cat.text = category_text.upper()
    p_cat.font.size = Pt(9.5)
    p_cat.font.bold = True
    p_cat.font.color.rgb = RGBColor(147, 197, 253)

    # Main Title
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.38), Inches(11.7), Inches(0.65))
    tf_title = title_box.text_frame
    tf_title.word_wrap = True
    p_title = tf_title.paragraphs[0]
    p_title.text = title_text
    p_title.font.size = Pt(20)
    p_title.font.bold = True
    p_title.font.color.rgb = RGBColor(255, 255, 255)

def add_speaker_notes(slide, notes_text):
    """Attaches formal spoken rehearsal notes to the slide."""
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = notes_text.strip()

def create_card(slide, left, top, width, height, title="", title_color=COLOR_NAVY):
    """Draws a modern rounded or rectangular card container."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_CARD_BG
    card.line.color.rgb = COLOR_BORDER
    card.line.width = Pt(1.5)

    if title:
        tb = slide.shapes.add_textbox(Inches(left + 0.2), Inches(top + 0.15), Inches(width - 0.4), Inches(0.45))
        p = tb.text_frame.paragraphs[0]
        p.text = title
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = title_color

    return card

def build_presentation(output_path):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # =========================================================================
    # SLIDE 1: Title Slide
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    # Background fill
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = COLOR_NAVY
    bg1.line.color.rgb = COLOR_NAVY

    # Institution Header
    h1_box = s1.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.333), Inches(0.8))
    p1_inst = h1_box.text_frame.paragraphs[0]
    p1_inst.text = "ABES ENGINEERING COLLEGE, GHAZIABAD"
    p1_inst.font.size = Pt(16)
    p1_inst.font.bold = True
    p1_inst.font.color.rgb = RGBColor(147, 197, 253)

    p1_dept = h1_box.text_frame.add_paragraph()
    p1_dept.text = "DEPARTMENT OF MASTER OF COMPUTER APPLICATIONS (MCA) • MCA 2nd Year Capstone"
    p1_dept.font.size = Pt(11)
    p1_dept.font.color.rgb = RGBColor(203, 213, 225)

    # Project Title
    title_box1 = s1.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.333), Inches(1.8))
    p1_t = title_box1.text_frame.paragraphs[0]
    p1_t.text = "Outpatient Appointment No-Show Analysis\n& Capacity Optimization System"
    p1_t.font.size = Pt(32)
    p1_t.font.bold = True
    p1_t.font.color.rgb = RGBColor(255, 255, 255)

    p1_sub = title_box1.text_frame.add_paragraph()
    p1_sub.text = "Operational Diagnostics, Parametric Hypothesis Testing & Predictive Risk Triage"
    p1_sub.font.size = Pt(14)
    p1_sub.font.color.rgb = RGBColor(226, 232, 240)

    # Details Card on Slide 1
    card1 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(4.3), Inches(11.333), Inches(2.2))
    card1.fill.solid()
    card1.fill.fore_color.rgb = RGBColor(15, 23, 42)
    card1.line.color.rgb = COLOR_ACCENT

    meta_tb1 = s1.shapes.add_textbox(Inches(1.3), Inches(4.5), Inches(10.7), Inches(1.8))
    tf1_meta = meta_tb1.text_frame
    
    p = tf1_meta.paragraphs[0]
    p.text = "Healthcare Network Modeled: ApexCare Health Network, Noida (Sectors 62, 18, 137, Greater Noida)"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)

    p = tf1_meta.add_paragraph()
    p.text = "Student Developer: Deepshikha  |  University Roll No: 2500320140047  |  MCA 2nd Year (Data Analytics)"
    p.font.size = Pt(11)
    p.font.color.rgb = RGBColor(203, 213, 225)

    p = tf1_meta.add_paragraph()
    p.text = "Faculty Project Mentor: Alpna Lodhi mam  |  Academic Session: 2026–2027"
    p.font.size = Pt(11)
    p.font.color.rgb = RGBColor(203, 213, 225)

    p = tf1_meta.add_paragraph()
    p.text = "Live App URL: https://2500320140047-appointment-noshow.streamlit.app"
    p.font.size = Pt(10)
    p.font.color.rgb = RGBColor(96, 165, 250)

    add_speaker_notes(s1, 
        "Good morning respected evaluators and faculty members. My name is Deepshikha, Roll Number 2500320140047, "
        "and today I am presenting my MCA Capstone Project: the Outpatient Appointment No-Show Analysis and Capacity "
        "Optimization System, completed under the mentorship of Alpna Lodhi mam. This project tackles the endemic problem "
        "of patient no-shows across ApexCare Health Network's four outpatient clinics in Noida using Python, Streamlit, "
        "and statistical inference. Over the next five minutes, I will walk you through the operational challenge, "
        "our analytical solution, key findings, and live demonstration."
    )

    # =========================================================================
    # SLIDE 2: Problem + Top 2 HMW
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "Operational Problem & Design Thinking Opportunities")

    create_card(s2, 0.8, 1.45, 5.6, 5.4, "The Outpatient Healthcare Challenge")
    tb2_left = s2.shapes.add_textbox(Inches(1.0), Inches(2.05), Inches(5.2), Inches(4.5))
    tf2_l = tb2_left.text_frame
    tf2_l.word_wrap = True
    
    bullets2_l = [
        ("Endemic Non-Attendance: ", "18.33% unmitigated no-show rate across audited encounters (55 missed consultations)."),
        ("Financial Revenue Leakage: ", "₹49,500 in consultation fees permanently lost within a single sample cycle."),
        ("Physician Idleness: ", "23.8 specialist doctor hours squandered, inflating clinic overhead and patient queues."),
        ("Chronic Care Risks: ", "Unscheduled gaps in care for patients with Hypertension and Diabetes."),
        ("Operational Blindspots: ", "Absence of predictive triage rosters forces front-desk to call patients blindly.")
    ]
    for idx, (b_title, b_text) in enumerate(bullets2_l):
        p = tf2_l.paragraphs[0] if idx == 0 else tf2_l.add_paragraph()
        p.space_after = Pt(8)
        r1 = p.add_run()
        r1.text = "• " + b_title
        r1.font.bold = True
        r1.font.size = Pt(11)
        r1.font.color.rgb = COLOR_NAVY
        r2 = p.add_run()
        r2.text = b_text
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = COLOR_DARK

    create_card(s2, 6.8, 1.45, 5.7, 5.4, "Top 2 How Might We (HMW) Statements")
    tb2_right = s2.shapes.add_textbox(Inches(7.0), Inches(2.05), Inches(5.3), Inches(4.5))
    tf2_r = tb2_right.text_frame
    tf2_r.word_wrap = True

    bullets2_r = [
        ("HMW-02 (Lead-Time Cutoff Thresholds):", "How might we identify booking lead times where attendance drops significantly, so that scheduling coordinators can trigger advance confirmation protocols before appointments lapse?"),
        ("HMW-05 (SMS Reminder Efficacy Proof):", "How might we empirically measure the attendance lift generated by automated SMS reminders, so that clinical administrators can justify digital infrastructure investments using parametric hypothesis testing?"),
        ("Strategic Alignment:", "These two design opportunities directly address booking decay (OQ-02) and communication intervention lift (OQ-05), forming the analytical backbone of our system.")
    ]
    for idx, (b_title, b_text) in enumerate(bullets2_r):
        p = tf2_r.paragraphs[0] if idx == 0 else tf2_r.add_paragraph()
        p.space_after = Pt(12)
        r1 = p.add_run()
        r1.text = b_title + "\n"
        r1.font.bold = True
        r1.font.size = Pt(11)
        r1.font.color.rgb = COLOR_ACCENT
        r2 = p.add_run()
        r2.text = b_text
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = COLOR_DARK

    add_speaker_notes(s2,
        "Moving to the business problem: outpatient no-shows create severe financial and operational damage. "
        "At ApexCare's four Noida clinics, the unmitigated no-show rate stands at 18.33 percent. In our audited dataset, "
        "this caused forty-nine thousand five hundred rupees in uncollected consultation fees and twenty-three point eight "
        "doctor consultation hours wasted. To tackle this, we formulated two primary How Might We questions: first, "
        "how might we detect lead-time decay thresholds to trigger reconfirmation protocols; and second, how might we "
        "statistically prove the attendance lift created by automated digital SMS reminders."
    )

    # =========================================================================
    # SLIDE 3: Key Stakeholders & RACI Influence
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "Key Stakeholder Profiles & Governance Matrix")

    sh_cards_data = [
        ("SH-01: Clinic Operations Director", "Executive Healthcare Administration", [
            ("Power / Influence:", "High Power / High Interest (Key Driver)"),
            ("Accountability:", "Multi-branch attendance benchmarking & revenue recovery"),
            ("System Need:", "Executive KPI dashboard, capacity utilization metrics")
        ]),
        ("SH-02: Chief Medical Officer (CMO)", "Clinical Governance & Medical Staff", [
            ("Power / Influence:", "High Power / High Interest (Key Driver)"),
            ("Accountability:", "Specialty physician rosters & chronic disease compliance"),
            ("System Need:", "Department show-rate comparisons, doctor idle hours")
        ]),
        ("SH-03: Front-Desk Scheduling Lead", "Patient Services & Front-Desk Reception", [
            ("Power / Influence:", "Medium Power / High Interest (Primary User)"),
            ("Accountability:", "Patient call rosters & advance appointment confirmation"),
            ("System Need:", "Prioritized triage calling queue, reminder logs")
        ])
    ]

    for idx, (sh_title, sh_sub, items) in enumerate(sh_cards_data):
        cx = 0.8 + idx * 3.95
        create_card(s3, cx, 1.45, 3.75, 4.2, sh_title)
        tb = s3.shapes.add_textbox(Inches(cx + 0.15), Inches(2.05), Inches(3.45), Inches(3.4))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p_sub = tf.paragraphs[0]
        p_sub.text = sh_sub
        p_sub.font.size = Pt(9.5)
        p_sub.font.italic = True
        p_sub.font.color.rgb = COLOR_MUTED
        p_sub.space_after = Pt(8)

        for l_lbl, l_val in items:
            p = tf.add_paragraph()
            p.space_after = Pt(6)
            r1 = p.add_run()
            r1.text = l_lbl + " "
            r1.font.bold = True
            r1.font.size = Pt(10)
            r1.font.color.rgb = COLOR_NAVY
            r2 = p.add_run()
            r2.text = l_val
            r2.font.size = Pt(9.5)
            r2.font.color.rgb = COLOR_DARK

    # Bottom RACI Summary Box
    create_card(s3, 0.8, 5.85, 11.7, 1.15, "RACI Governance Summary")
    tb3_bot = s3.shapes.add_textbox(Inches(1.0), Inches(6.25), Inches(11.3), Inches(0.65))
    tf3_b = tb3_bot.text_frame
    p_b = tf3_b.paragraphs[0]
    p_b.text = "• Accountable: SH-01 (Capacity/Revenue) & SH-02 (Clinical Quality)  |  • Responsible: SH-03 (Triage Calls) & SH-05 (Analytics/Integrity)\n• Consulted: SH-04 (Corporate IT Wellness Partners)  |  • Informed: External Patients & Clinical Staff"
    p_b.font.size = Pt(10)
    p_b.font.color.rgb = COLOR_DARK

    add_speaker_notes(s3,
        "Here we see our stakeholder governance model. The Clinic Operations Director, SH-01, is accountable for "
        "overall capacity and revenue reclamation. The Chief Medical Officer, SH-02, governs specialty scheduling and "
        "chronic care compliance for patients with hypertension and diabetes. The Front-Desk Scheduling Supervisor, "
        "SH-03, is the primary day-to-day user who executes patient outreach using our advance risk triage rosters. "
        "Our RACI matrix ensures clear operational accountability across all clinical requirements."
    )

    # =========================================================================
    # SLIDE 4: Traceability Matrix (3-4 Must Rows)
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "Requirements Traceability Matrix (Top Must-Have Requirements)")

    # Table
    rows_cnt = 5
    cols_cnt = 6
    tbl_shape4 = s4.shapes.add_table(rows_cnt, cols_cnt, Inches(0.8), Inches(1.5), Inches(11.733), Inches(4.5))
    t4 = tbl_shape4.table
    col_widths4 = [Inches(1.1), Inches(1.1), Inches(3.8), Inches(1.3), Inches(2.2), Inches(2.233)]
    for i, w in enumerate(col_widths4):
        t4.columns[i].width = w

    headers4 = ["HMW ID", "BR ID", "Business Requirement Summary", "User Story", "UI Screen", "Test Case (TC)"]
    for i, h in enumerate(headers4):
        cell = t4.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_NAVY
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.bold = True
        p.font.size = Pt(10)
        p.font.color.rgb = RGBColor(255, 255, 255)

    data4 = [
        ("HMW-01", "BR-01 (Must)", "Track historical and ongoing attendance rates across 4 clinic locations", "US-01, US-03", "UI-01 (Overview)\nUI-08 (Data Quality)", "TC-F-01..04\nTC-T-01, 02, 07"),
        ("HMW-02", "BR-02 (Must)", "Analyze attendance decay across 5 lead-time intervals (0-2d to >30d)", "US-05", "UI-02 (Lead Time &\nSpecialty)", "TC-F-05\nTC-T-05, 10"),
        ("HMW-05", "BR-05 (Must)", "Empirically measure attendance lift between SMS reminder and control cohorts", "US-09", "UI-04 (Statistical\nInference)", "TC-F-09\nTC-T-09, 14"),
        ("HMW-10", "BR-10 (Must)", "Calculate advance predictive risk score (0-100) and assign High/Med/Low tiers", "US-13", "UI-06 (Predictive\nRisk Triage)", "TC-F-13\nTC-T-11, 16")
    ]
    for r_idx, row_vals in enumerate(data4, start=1):
        for c_idx, val in enumerate(row_vals):
            cell = t4.cell(r_idx, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(241, 245, 249) if r_idx % 2 == 1 else RGBColor(255, 255, 255)
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.size = Pt(9.5)
            p.font.color.rgb = COLOR_DARK

    # Bottom Note
    nb4 = s4.shapes.add_textbox(Inches(0.8), Inches(6.2), Inches(11.7), Inches(0.8))
    p4_note = nb4.text_frame.paragraphs[0]
    p4_note.text = "Traceability Audit Result: 100% bi-directional mapping verified across all 12 Business Requirements (BR-01 to BR-12), linking design opportunities, user stories, database fields, UI screens, and automated tests."
    p4_note.font.size = Pt(10)
    p4_note.font.bold = True
    p4_note.font.color.rgb = COLOR_NAVY

    add_speaker_notes(s4,
        "Traceability is the hallmark of professional engineering. This matrix highlights four of our core Must-Have requirements: "
        "BR-01 for multi-branch tracking; BR-02 for lead-time decay modeling; BR-05 for empirical SMS reminder testing; "
        "and BR-10 for predictive risk triage. Notice how every single requirement traces cleanly from the original HMW statement "
        "to the agile user story, the dedicated UI screen, and both functional and technical automated test cases."
    )

    # =========================================================================
    # SLIDE 5: Data Layer & Data Hygiene
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "Relational Healthcare Data Layer & Data Hygiene")

    create_card(s5, 0.8, 1.45, 5.6, 5.4, "Relational Dataset Architecture")
    tb5_l = s5.shapes.add_textbox(Inches(1.0), Inches(2.05), Inches(5.2), Inches(4.5))
    tf5_l = tb5_l.text_frame
    tf5_l.word_wrap = True

    data_bullets_l = [
        ("Fact Appointments Volume: ", "300 transactional encounters covering Q3-Q4 2026."),
        ("Entity Capacity: ", "200 Patients, 12 Doctors, 6 Specialties, 4 Clinic Branches."),
        ("Denormalized Analytics View: ", "vw_appointment_analytics enriched with 36 columns."),
        ("Database Engine: ", "SQLite 3 (data/app.db) enforcing PRAGMA foreign_keys = ON."),
        ("Zero PHI / HIPAA Compliant: ", "100% synthetic, ethically generated dataset (seed 42).")
    ]
    for idx, (b_title, b_text) in enumerate(data_bullets_l):
        p = tf5_l.paragraphs[0] if idx == 0 else tf5_l.add_paragraph()
        p.space_after = Pt(10)
        r1 = p.add_run()
        r1.text = "• " + b_title
        r1.font.bold = True
        r1.font.size = Pt(11)
        r1.font.color.rgb = COLOR_NAVY
        r2 = p.add_run()
        r2.text = b_text
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = COLOR_DARK

    create_card(s5, 6.8, 1.45, 5.7, 5.4, "Top 3 Data Quality Cleaning Fixes")
    tb5_r = s5.shapes.add_textbox(Inches(7.0), Inches(2.05), Inches(5.3), Inches(4.5))
    tf5_r = tb5_r.text_frame
    tf5_r.word_wrap = True

    fixes = [
        ("1. Missing Value Imputation (reminder_lead_hours):", "Identified 7 rows (2.3%) with unrecorded reminder delivery lead hours. Imputed using sample median (24.0 hours) to ensure analytical continuity without biasing variances."),
        ("2. Primary Key Uniqueness & Deduplication:", "Enforced strict zero-duplicate constraint across all 5 tables (total == distinct count). Audited 300 unique appointment IDs with zero collision."),
        ("3. Domain Bounds & Range Validation:", "Validated lead_time_days >= 0 (eliminated negative booking dates), confirmed consultation fees strictly within tariff schedule (₹500-₹1500), and enforced status in {'Show', 'No-Show'}.")
    ]
    for idx, (f_title, f_text) in enumerate(fixes):
        p = tf5_r.paragraphs[0] if idx == 0 else tf5_r.add_paragraph()
        p.space_after = Pt(10)
        r1 = p.add_run()
        r1.text = f_title + "\n"
        r1.font.bold = True
        r1.font.size = Pt(11)
        r1.font.color.rgb = COLOR_ACCENT
        r2 = p.add_run()
        r2.text = f_text
        r2.font.size = Pt(10)
        r2.font.color.rgb = COLOR_DARK

    add_speaker_notes(s5,
        "Turning to our data layer: our SQLite database hosts five relational tables denormalized into a thirty-six column "
        "analytical view. The dataset contains three hundred appointment encounters across four clinics and twelve specialists. "
        "During data ingestion, our cleaner module executed three critical quality fixes: first, imputing missing reminder "
        "lead hours with the median of twenty-four hours; second, verifying zero duplicate primary keys; and third, validating "
        "strict range constraints such as non-negative lead times and allowed fee schedules."
    )

    # =========================================================================
    # SLIDE 6: Solution + Tech Flow
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, "End-to-End System Architecture & Technical Flow")

    # Flow Cards
    steps6 = [
        ("1. Ingestion Layer", "SQLite 3 Database", "data/app.db\n• 5 Normalized Tables\n• 36-col Analytical View\n• PRAGMA foreign_keys ON"),
        ("2. Data Layer", "Caching & Cleaning", "src.data.loader & cleaner\n• @st.cache_data (ttl=3600)\n• Median Imputation\n• Typing & Outlier Bounds"),
        ("3. Analytical Core", "Mathematical Engines", "src.analytics & stats\n• Z-test & Chi-Square\n• Lead-time decay\n• Risk scoring engine"),
        ("4. Presentation", "Multi-Page UI", "Streamlit Multi-Page\n• 8 Dedicated screens\n• Tailwind CSS CDN\n• Interactive Plotly charts")
    ]
    for idx, (st_num, st_name, st_desc) in enumerate(steps6):
        cx = 0.8 + idx * 2.95
        create_card(s6, cx, 1.45, 2.8, 4.3, st_num)
        tb = s6.shapes.add_textbox(Inches(cx + 0.15), Inches(2.05), Inches(2.5), Inches(3.5))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p1 = tf.paragraphs[0]
        p1.text = st_name
        p1.font.bold = True
        p1.font.size = Pt(11)
        p1.font.color.rgb = COLOR_ACCENT
        p1.space_after = Pt(8)

        p2 = tf.add_paragraph()
        p2.text = st_desc
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = COLOR_DARK

    # Bottom Resilience Banner
    create_card(s6, 0.8, 5.95, 11.7, 1.1, "Architectural Resilience Highlight")
    tb6_bot = s6.shapes.add_textbox(Inches(1.0), Inches(6.3), Inches(11.3), Inches(0.65))
    tf6_b = tb6_bot.text_frame
    p_b6 = tf6_b.paragraphs[0]
    p_b6.text = "• Environment Hardening: Built with zero dependence on vulnerable compiled C-extensions, eliminating 100% of Windows WDAC/AppLocker DLL blocks.\n• Responsive Performance: Sub-second query response time powered by Streamlit memory caching and pure-python statistical mathematics."
    p_b6.font.size = Pt(9.5)
    p_b6.font.color.rgb = COLOR_DARK

    add_speaker_notes(s6,
        "Here is our system architecture and data flow. Data flows from our SQLite database into a cached loader "
        "and data cleaner. The cleaned records feed our modular analytics and hypothesis testing engines. The presentation "
        "layer delivers an interactive multi-page experience using Streamlit and Tailwind CSS. Crucially, we engineered the entire "
        "backend using pure-Python standard algorithms, completely bypassing Windows AppLocker and WDAC binary DLL blocks, "
        "guaranteeing rock-solid stability and instant cold-start execution."
    )

    # =========================================================================
    # SLIDE 7: Dashboard (Screenshot + UI Screens)
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, "Interactive Streamlit Multi-Page Application")

    # Embed Screenshot if exists
    img_path = "docs/18_Dashboard_Preview.png"
    if os.path.exists(img_path):
        s7.shapes.add_picture(img_path, Inches(0.8), Inches(1.45), width=Inches(7.2))
    else:
        create_card(s7, 0.8, 1.45, 7.2, 5.4, "Live Dashboard Preview")

    # Right Card: UI Screens Summary
    create_card(s7, 8.2, 1.45, 4.3, 5.4, "8 Functional Analytics Screens")
    tb7_r = s7.shapes.add_textbox(Inches(8.35), Inches(2.0), Inches(4.0), Inches(4.6))
    tf7_r = tb7_r.text_frame
    tf7_r.word_wrap = True

    screens = [
        ("UI-00 / Root:", "Executive KPIs & 30-day rolling trend"),
        ("UI-01:", "Overview & Multi-Branch Benchmarking"),
        ("UI-02:", "Lead-Time Decay Curve & Specialty Rankings"),
        ("UI-03:", "Demographics & 2D Weekday Arrival Heatmap"),
        ("UI-04:", "Statistical Inference (Z-test & Chi-square)"),
        ("UI-05:", "Financial Leakage & Doctor Idle Hours"),
        ("UI-06:", "Predictive Risk Triage & Front-Desk Roster"),
        ("UI-07:", "Overbooking Capacity Simulator & Shifts"),
        ("UI-08:", "Data Hygiene Inspector & CSV/Excel Export")
    ]
    for idx, (s_code, s_desc) in enumerate(screens):
        p = tf7_r.paragraphs[0] if idx == 0 else tf7_r.add_paragraph()
        p.space_after = Pt(4)
        r1 = p.add_run()
        r1.text = s_code + " "
        r1.font.bold = True
        r1.font.size = Pt(10)
        r1.font.color.rgb = COLOR_NAVY
        r2 = p.add_run()
        r2.text = s_desc
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = COLOR_DARK

    add_speaker_notes(s7,
        "Slide seven showcases our live multi-page web application. On the left is an exact visual snapshot of our "
        "executive dashboard featuring Tailwind KPI cards, temporal rolling trends, and stacked branch comparisons. "
        "On the right, you see the eight specialized subpages we built—ranging from lead-time decay and bivariate arrival heatmaps "
        "to formal hypothesis testing, doctor idle capacity accounting, and front-desk phone calling queues with CSV export."
    )

    # =========================================================================
    # SLIDE 8: Key Findings (3 Numbers + Evidence)
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "Key Analytical Findings & Statistical Evidence")

    findings = [
        ("FINDING 1: SMS REMINDER LIFT", "48.84% RRR", COLOR_GREEN, [
            ("Statistical Test:", "Two-Sample Proportion Z-Test"),
            ("Test Statistic:", "Z = +2.7973, p = 0.00515 (p < 0.01)"),
            ("Odds Ratio (OR):", "2.30 [95% CI: 1.27 to 4.16]"),
            ("Clinical Evidence:", "Patients receiving SMS reminders had a 13.6% no-show rate vs 26.6% for uncontacted patients. Null hypothesis decisively rejected.")
        ]),
        ("FINDING 2: LEAD-TIME DECAY", "41.7% Loss", COLOR_RED, [
            ("Behavioral Metric:", "Booking Lead-Time Horizon Decay"),
            ("Short Horizon (0-2d):", "9.8% No-Show Rate (High Commitment)"),
            ("Long Horizon (>30d):", "41.7% No-Show Rate (Oblivion / Resolved)"),
            ("Operational Impact:", "Monotonic non-attendance escalation dictates mandatory 48-hour call confirmations for bookings >14 days.")
        ]),
        ("FINDING 3: PARETO LOSS RISK", "65% of Loss", COLOR_AMBER, [
            ("Concentration Law:", "Repeat Offender Pareto Analysis"),
            ("Top 20% Non-Attenders:", "Generate ~65% of Total Revenue Loss"),
            ("Total Financial Loss:", "₹49,500 uncollected (23.8 doctor hours)"),
            ("Actionable Triage:", "Targeting the top risk tier via front-desk phone outreach captures the vast majority of financial leakage.")
        ])
    ]

    for idx, (f_title, f_stat, f_col, f_items) in enumerate(findings):
        cx = 0.8 + idx * 3.95
        create_card(s8, cx, 1.45, 3.75, 5.4, f_title)
        
        # Stat Banner inside card
        sb = s8.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx + 0.2), Inches(2.0), Inches(3.35), Inches(0.8))
        sb.fill.solid()
        sb.fill.fore_color.rgb = f_col
        sb.line.color.rgb = f_col
        p_sb = sb.text_frame.paragraphs[0]
        p_sb.alignment = PP_ALIGN.CENTER
        p_sb.text = f_stat
        p_sb.font.bold = True
        p_sb.font.size = Pt(22)
        p_sb.font.color.rgb = RGBColor(255, 255, 255)

        tb = s8.shapes.add_textbox(Inches(cx + 0.15), Inches(2.95), Inches(3.45), Inches(3.7))
        tf = tb.text_frame
        tf.word_wrap = True

        for l_lbl, l_val in f_items:
            p = tf.add_paragraph()
            p.space_after = Pt(6)
            r1 = p.add_run()
            r1.text = l_lbl + " "
            r1.font.bold = True
            r1.font.size = Pt(9.5)
            r1.font.color.rgb = COLOR_NAVY
            r2 = p.add_run()
            r2.text = l_val
            r2.font.size = Pt(9.0)
            r2.font.color.rgb = COLOR_DARK

    add_speaker_notes(s8,
        "Here are our three core empirical findings: First, automated SMS reminders generate a forty-eight point eight percent "
        "Relative Risk Reduction in no-shows. With a Z-statistic of 2.80 and a p-value of 0.0052, this is statistically significant "
        "beyond the one percent level. Second, we proved lead-time decay: appointments booked zero to two days out have under ten "
        "percent no-shows, but bookings over thirty days out escalate to forty-one point seven percent. Third, Pareto analysis reveals "
        "that the top twenty percent of repeat non-attending patients drive sixty-five percent of cumulative financial losses."
    )

    # =========================================================================
    # SLIDE 9: Testing & Quality Assurance
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_header(s9, "Quality Assurance & Automated Testing Verification")

    create_card(s9, 0.8, 1.45, 5.0, 5.4, "Automated Test Suite Summary")
    tb9_l = s9.shapes.add_textbox(Inches(1.0), Inches(2.05), Inches(4.6), Inches(4.5))
    tf9_l = tb9_l.text_frame
    tf9_l.word_wrap = True

    t_bullets_l = [
        ("Pytest Execution: ", "20 Automated Tests Executed."),
        ("Pass Rate: ", "100% Passed (20 passed in 0.94s)."),
        ("Functional Tests: ", "16 Acceptance Tests (TC-F-01..16) Verified."),
        ("Schema Verification: ", "Zero foreign key violations via PRAGMA."),
        ("Formula Verification: ", "Exact match on known mock datasets."),
        ("Regression Shield: ", "Automated tests guard every calculation.")
    ]
    for idx, (b_title, b_text) in enumerate(t_bullets_l):
        p = tf9_l.paragraphs[0] if idx == 0 else tf9_l.add_paragraph()
        p.space_after = Pt(8)
        r1 = p.add_run()
        r1.text = "• " + b_title
        r1.font.bold = True
        r1.font.size = Pt(11)
        r1.font.color.rgb = COLOR_GREEN if "Pass" in b_title else COLOR_NAVY
        r2 = p.add_run()
        r2.text = b_text
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = COLOR_DARK

    create_card(s9, 6.1, 1.45, 6.4, 5.4, "3 Key Automated Verification Tests")
    
    # Table for 3 test rows
    t_shape9 = s9.shapes.add_table(4, 4, Inches(6.3), Inches(2.1), Inches(6.0), Inches(4.3))
    t9 = t_shape9.table
    w9 = [Inches(1.1), Inches(1.8), Inches(2.1), Inches(1.0)]
    for i, w in enumerate(w9):
        t9.columns[i].width = w

    h9 = ["Test ID", "Tested Component", "Automated Assertion", "Status"]
    for i, h in enumerate(h9):
        cell = t9.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_NAVY
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.bold = True
        p.font.size = Pt(9.5)
        p.font.color.rgb = RGBColor(255, 255, 255)

    rows9 = [
        ("TC-T-04", "Foreign Key Integrity\n(tests/test_data.py)", "PRAGMA foreign_key_check returns zero orphan records", "PASSED\n(100%)"),
        ("TC-T-09", "Z-Test Math Logic\n(tests/test_analysis.py)", "Proportion Z-score > 2.5 on mock cohort data", "PASSED\n(100%)"),
        ("TC-T-11", "Risk Scoring Tiers\n(tests/test_analysis.py)", "Threshold checks: 85.5->High, 59.9->Med, 29.9->Low", "PASSED\n(100%)")
    ]
    for r_idx, r_data in enumerate(rows9, start=1):
        for c_idx, val in enumerate(r_data):
            cell = t9.cell(r_idx, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(241, 245, 249) if r_idx % 2 == 1 else RGBColor(255, 255, 255)
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.size = Pt(9.0)
            p.font.color.rgb = COLOR_GREEN if c_idx == 3 else COLOR_DARK

    add_speaker_notes(s9,
        "Testing and quality assurance were built into every step. Our test suite includes sixteen manual functional acceptance "
        "tests and twenty automated Pytests, all passing in under one second with a one hundred percent pass rate. This table shows "
        "three vital automated tests: TC-T-04 validating foreign key integrity; TC-T-09 proving our mathematical Z-test formula; "
        "and TC-T-11 asserting exact boundary classifications for our predictive risk tiers."
    )

    # =========================================================================
    # SLIDE 10: Recommendations, Limitations & Next Steps
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_header(s10, "Strategic Recommendations, Limitations & Roadmap")

    cols10_data = [
        ("Operational Directives", COLOR_GREEN, [
            ("Mandatory Call Windows: ", "Enforce 48h phone confirmations for lead times >14 days."),
            ("Automated Reminders: ", "Mandate 24h SMS notifications across all branches."),
            ("Calibrated Overbooking: ", "Apply 1.12x capacity buffer on high-volume outpatient slots.")
        ]),
        ("Technical Limitations", COLOR_AMBER, [
            ("Synthetic Cohort: ", "Evaluated on 300 synthetic records modeled for Noida operations."),
            ("Single-Tenant Storage: ", "SQLite relational database optimized for analytics read performance."),
            ("Batch Refresh: ", "Daily data refresh cycle rather than streaming telemetry.")
        ]),
        ("Future Engineering Roadmap", COLOR_ACCENT, [
            ("EHR Interoperability: ", "Live HL7/FHIR API integration with hospital HIS platforms."),
            ("Automated Telephony: ", "Deploy interactive voice response (IVR) voice bots."),
            ("PostgreSQL Migration: ", "Scale to multi-tenant distributed cloud database.")
        ])
    ]

    for idx, (c_title, c_col, c_items) in enumerate(cols10_data):
        cx = 0.8 + idx * 3.95
        create_card(s10, cx, 1.45, 3.75, 5.4, c_title, title_color=c_col)
        tb = s10.shapes.add_textbox(Inches(cx + 0.15), Inches(2.05), Inches(3.45), Inches(4.5))
        tf = tb.text_frame
        tf.word_wrap = True

        for l_lbl, l_val in c_items:
            p = tf.add_paragraph()
            p.space_after = Pt(10)
            r1 = p.add_run()
            r1.text = "• " + l_lbl
            r1.font.bold = True
            r1.font.size = Pt(10)
            r1.font.color.rgb = COLOR_NAVY
            r2 = p.add_run()
            r2.text = l_val
            r2.font.size = Pt(9.5)
            r2.font.color.rgb = COLOR_DARK

    add_speaker_notes(s10,
        "Based on these findings, we provide three concrete operational recommendations: implement mandatory phone calls "
        "for appointments booked more than fourteen days in advance; mandate automated SMS reminders across all clinics; "
        "and introduce a calibrated twelve percent overbooking buffer to reclaim lost slots. While our system was built on a synthetic "
        "three-hundred encounter dataset in SQLite, our future roadmap includes integrating live HL7/FHIR hospital APIs and migrating "
        "to a distributed cloud PostgreSQL database."
    )

    # =========================================================================
    # SLIDE 11: Live Demo & Artifacts
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    add_header(s11, "Live Demonstration & Deliverable Artifacts")

    # 2 Large Action Cards
    create_card(s11, 0.8, 1.45, 5.7, 4.4, "Live Cloud Web Application")
    tb11_l = s11.shapes.add_textbox(Inches(1.0), Inches(2.1), Inches(5.3), Inches(3.5))
    tf11_l = tb11_l.text_frame
    tf11_l.word_wrap = True

    p = tf11_l.paragraphs[0]
    p.text = "Hosted on Streamlit Community Cloud:"
    p.font.size = Pt(11)
    p.font.color.rgb = COLOR_MUTED
    p.space_after = Pt(6)

    p_url = tf11_l.add_paragraph()
    p_url.text = "https://2500320140047-appointment-noshow.streamlit.app"
    p_url.font.size = Pt(11)
    p_url.font.bold = True
    p_url.font.color.rgb = COLOR_ACCENT
    p_url.space_after = Pt(14)

    p_feat = tf11_l.add_paragraph()
    p_feat.text = "• 8 Interactive Multi-Page Analytical Screens\n• Live Global Sidebar Filters (Clinic, Specialty, Shift)\n• Real-Time Plotly Visualizations & Metric Cards\n• One-Click In-Memory CSV & Excel Data Export"
    p_feat.font.size = Pt(10)
    p_feat.font.color.rgb = COLOR_DARK

    create_card(s11, 6.8, 1.45, 5.7, 4.4, "GitHub Open-Source Repository")
    tb11_r = s11.shapes.add_textbox(Inches(7.0), Inches(2.1), Inches(5.3), Inches(3.5))
    tf11_r = tb11_r.text_frame
    tf11_r.word_wrap = True

    p = tf11_r.paragraphs[0]
    p.text = "Complete Source Code & Documentation Suite:"
    p.font.size = Pt(11)
    p.font.color.rgb = COLOR_MUTED
    p.space_after = Pt(6)

    p_git = tf11_r.add_paragraph()
    p_git.text = "https://github.com/deepshikha-0047/2500320140047-appointment-noshow"
    p_git.font.size = Pt(11)
    p_git.font.bold = True
    p_git.font.color.rgb = COLOR_ACCENT
    p_git.space_after = Pt(14)

    p_repo = tf11_r.add_paragraph()
    p_repo.text = "• Complete Codebase (app.py, pages/, src/, tests/, data/)\n• Production SQLite Database (data/app.db)\n• Full Academic Deliverables Suite (Steps 01 through 17)\n• 20 Automated Pytest Test Cases (100% Pass Rate)"
    p_repo.font.size = Pt(10)
    p_repo.font.color.rgb = COLOR_DARK

    # Bottom Call to Action Card
    create_card(s11, 0.8, 6.05, 11.7, 1.05, "Live Evaluation & Q&A")
    tb11_b = s11.shapes.add_textbox(Inches(1.0), Inches(6.35), Inches(11.3), Inches(0.65))
    tf11_b = tb11_b.text_frame
    p_c = tf11_b.paragraphs[0]
    p_c.text = "Thank you! The application is live and accessible. I am now delighted to demonstrate the live platform and welcome your questions."
    p_c.font.size = Pt(11)
    p_c.font.bold = True
    p_c.font.color.rgb = COLOR_NAVY

    add_speaker_notes(s11,
        "In conclusion, our platform is fully deployed and accessible on Streamlit Community Cloud, and our complete "
        "codebase and academic deliverables are hosted on GitHub. I would like to express my sincere gratitude to Alpna Lodhi mam "
        "and the MCA department for their guidance. I am now excited to switch to the live application for a demonstration "
        "and welcome any questions from the panel. Thank you very much."
    )

    prs.save(output_path)
    print(f"Successfully generated 11-Slide Master Presentation Deck: {output_path}")

if __name__ == "__main__":
    primary_out = r"c:\Users\deeps\Downloads\Appointment No Show Analysis\docs\2500320140047_Presentation.pptx"
    backup_out = r"c:\Users\deeps\Downloads\Appointment No Show Analysis\docs\18_Presentation.pptx"
    build_presentation(primary_out)
    shutil.copyfile(primary_out, backup_out)
    print(f"Successfully mirrored presentation to: {backup_out}")
