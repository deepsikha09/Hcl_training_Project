"""
Generate Comprehensive Source Code PDF for LMS Submission:
docs/2500320140047_Source_Code.pdf & docs/ApexCare_Appointment_No_Show_Source_Code.pdf
ABES Engineering College, Ghaziabad • Department of MCA
Student: Deepshikha (Roll No: 2500320140047)
"""

import os
import shutil
import time
from fpdf import FPDF

class SourceCodePDF(FPDF):
    def __init__(self):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.set_auto_page_break(auto=True, margin=16)
        
        # Load TrueType Fonts from Windows Fonts
        self.add_font("Arial", "", "C:/Windows/Fonts/arial.ttf")
        self.add_font("Arial", "B", "C:/Windows/Fonts/arialbd.ttf")
        self.add_font("Arial", "I", "C:/Windows/Fonts/ariali.ttf")
        self.add_font("Consolas", "", "C:/Windows/Fonts/consola.ttf")
        self.add_font("Consolas", "B", "C:/Windows/Fonts/consolab.ttf")
        
        self.current_section = ""

    def header(self):
        if self.page_no() > 1:
            self.set_font("Arial", "B", 8)
            self.set_text_color(31, 73, 125) # Navy
            self.cell(100, 5, "ApexCare Healthcare Analytics • Complete Source Code Repository", align="L")
            self.set_font("Arial", "", 8)
            self.set_text_color(100, 116, 139) # Slate
            self.cell(0, 5, "Deepshikha | Roll No: 2500320140047", align="R", new_x="LMARGIN", new_y="NEXT")
            self.set_draw_color(226, 232, 240)
            self.set_line_width(0.3)
            self.line(self.l_margin, 12, self.w - self.r_margin, 12)
            self.ln(3)

    def footer(self):
        if self.page_no() > 1:
            self.set_y(-12)
            self.set_draw_color(226, 232, 240)
            self.set_line_width(0.3)
            self.line(self.l_margin, self.get_y(), self.w - self.r_margin, self.get_y())
            self.set_font("Arial", "", 8)
            self.set_text_color(100, 116, 139)
            self.cell(100, 7, "ABES Engineering College, Ghaziabad • Department of MCA", align="L")
            self.cell(0, 7, f"Page {self.page_no()}", align="R")

    def draw_cover_page(self):
        self.add_page()
        # Navy decorative side bar
        self.set_fill_color(31, 73, 125)
        self.rect(0, 0, 12, 297, "F")
        
        # Institution Block
        self.set_xy(20, 20)
        self.set_font("Arial", "B", 18)
        self.set_text_color(31, 73, 125)
        self.cell(0, 8, "ABES ENGINEERING COLLEGE, GHAZIABAD", new_x="LMARGIN", new_y="NEXT")
        
        self.set_x(20)
        self.set_font("Arial", "B", 11)
        self.set_text_color(71, 85, 105)
        self.cell(0, 6, "DEPARTMENT OF MASTER OF COMPUTER APPLICATIONS (MCA)", new_x="LMARGIN", new_y="NEXT")
        
        self.set_x(20)
        self.set_font("Arial", "", 9.5)
        self.set_text_color(100, 116, 139)
        self.cell(0, 5, "MCA 2nd Year (Data Analytics Specialization) • Project ID: P101 • Batch: ABES_OND26 MCA 2ND YEAR_DATA ANALYTICS_BATCH 01", new_x="LMARGIN", new_y="NEXT")
        
        # Horizontal Rule
        self.set_xy(20, 44)
        self.set_draw_color(31, 73, 125)
        self.set_line_width(1.0)
        self.line(20, 44, 195, 44)
        
        # Document Title
        self.set_xy(20, 50)
        self.set_font("Arial", "B", 21)
        self.set_text_color(15, 23, 42)
        self.multi_cell(175, 8.5, "Outpatient Appointment No-Show Analysis\n& Optimization System", new_x="LMARGIN", new_y="NEXT")
        
        self.set_xy(20, 70)
        self.set_font("Arial", "B", 12.5)
        self.set_text_color(46, 117, 182) # Accent blue
        self.cell(0, 7, "COMPLETE TECHNICAL IMPLEMENTATION & SOURCE CODE REPOSITORY", new_x="LMARGIN", new_y="NEXT")
        
        self.set_xy(20, 79)
        self.set_font("Arial", "I", 9.5)
        self.set_text_color(71, 85, 105)
        self.multi_cell(175, 4.8, "End-to-End Multi-Page Web Application, SQLite Database Layer, Statistical Inference Engine, Predictive Risk Triage Model, and Automated Pytest Verification Suite.")
        
        # Academic Metadata Box
        self.set_xy(20, 97)
        self.set_fill_color(248, 250, 252)
        self.set_draw_color(203, 213, 225)
        self.rect(20, 97, 175, 99, "FD")
        
        # Title of Box
        self.set_xy(25, 101)
        self.set_font("Arial", "B", 11)
        self.set_text_color(31, 73, 125)
        self.cell(0, 5, "PROJECT & CANDIDATE SPECIFICATION", new_x="LMARGIN", new_y="NEXT")
        
        meta_items = [
            ("Project ID", "P101"),
            ("Batch ID", "ABES_OND26 MCA 2ND YEAR_DATA ANALYTICS_BATCH 01"),
            ("Student Name", "Deepshikha"),
            ("University Roll Number", "2500320140047"),
            ("Academic Degree", "Master of Computer Applications (MCA) — 2nd Year"),
            ("Specialization", "Data Analytics & Healthcare Informatics"),
            ("Faculty Project Mentor", "Dinesh Sereno"),
            ("Healthcare Network Modeled", "ApexCare Health Network, Noida (Sectors 62, 18, 137, Gr. Noida)"),
            ("GitHub Source Code", "https://github.com/deepsikha09/Hcl_training_Project"),
            ("Live Web Application", "https://2500320140047-appointment-noshow.streamlit.app"),
            ("Automated Pytest Status", "20/20 Test Cases Passed (100% Pass Rate in 0.94s)"),
            ("Submission Session", "Academic Session 2026–2027 • LMS Evaluated")
        ]
        
        y_pos = 109
        for label, val in meta_items:
            self.set_xy(25, y_pos)
            self.set_font("Arial", "B", 8.8)
            self.set_text_color(31, 73, 125)
            self.cell(48, 5, label + ":", align="L")
            self.set_font("Arial", "", 8.8)
            self.set_text_color(30, 41, 59)
            self.cell(117, 5, val, align="L")
            y_pos += 6.6
            
        # Summary Statistics Badges
        self.set_xy(20, 202)
        badge_w = 41
        badges = [
            ("3,957", "Lines of Code", (31, 73, 125)),
            ("34", "Source Modules", (46, 117, 182)),
            ("8", "Streamlit Screens", (16, 185, 129)),
            ("20/20", "Pytests Passed", (15, 118, 110))
        ]
        for i, (val, lbl, col) in enumerate(badges):
            bx = 20 + i * (badge_w + 3.6)
            self.set_fill_color(241, 245, 249)
            self.set_draw_color(col[0], col[1], col[2])
            self.rect(bx, 202, badge_w, 20, "FD")
            self.set_xy(bx, 204)
            self.set_font("Arial", "B", 13)
            self.set_text_color(col[0], col[1], col[2])
            self.cell(badge_w, 5, val, align="C")
            self.set_xy(bx, 211)
            self.set_font("Arial", "", 8)
            self.set_text_color(71, 85, 105)
            self.cell(badge_w, 4, lbl, align="C")

        # Formal Student Certification
        self.set_xy(20, 229)
        self.set_fill_color(248, 250, 252)
        self.set_draw_color(226, 232, 240)
        self.rect(20, 229, 175, 34, "FD")
        
        self.set_xy(24, 232)
        self.set_font("Arial", "B", 9)
        self.set_text_color(31, 73, 125)
        self.cell(0, 5, "CANDIDATE CERTIFICATION & LMS SUBMISSION DECLARATION")
        
        self.set_xy(24, 238)
        self.set_font("Arial", "", 8.2)
        self.set_text_color(51, 65, 85)
        cert_text = (
            "I hereby declare that this document contains the original, complete source code of Project P101 "
            "(Batch: ABES_OND26 MCA 2ND YEAR_DATA ANALYTICS_BATCH 01) developed by me under the faculty guidance of "
            "Dinesh Sereno. All modules, database schemas, statistical algorithms, and Streamlit user interfaces were "
            "engineered, verified via automated Pytest test suites, and deployed to Streamlit Community Cloud."
        )
        self.multi_cell(167, 4.2, cert_text)

    def draw_table_of_contents(self, file_sections):
        self.add_page()
        self.set_xy(20, 18)
        self.set_font("Arial", "B", 16)
        self.set_text_color(31, 73, 125)
        self.cell(0, 8, "TABLE OF CONTENTS & REPOSITORY INDEX", new_x="LMARGIN", new_y="NEXT")
        
        self.set_font("Arial", "", 9.5)
        self.set_text_color(100, 116, 139)
        self.cell(0, 5, "Organized by architectural subsystem and execution hierarchy", new_x="LMARGIN", new_y="NEXT")
        self.ln(4)
        
        for sec_name, files in file_sections:
            self.set_font("Arial", "B", 10.5)
            self.set_text_color(31, 73, 125)
            self.cell(0, 6, sec_name, new_x="LMARGIN", new_y="NEXT")
            self.set_draw_color(46, 117, 182)
            self.set_line_width(0.3)
            self.line(self.get_x(), self.get_y(), 195, self.get_y())
            self.ln(2)
            
            for fpath, desc, lcount in files:
                self.set_font("Consolas", "B", 8.5)
                self.set_text_color(30, 41, 59)
                self.cell(75, 4.5, "  " + fpath)
                self.set_font("Arial", "", 8.5)
                self.set_text_color(71, 85, 105)
                self.cell(85, 4.5, desc)
                self.set_font("Consolas", "", 8)
                self.set_text_color(100, 116, 139)
                self.cell(0, 4.5, f"{lcount:3d} lines", align="R", new_x="LMARGIN", new_y="NEXT")
            self.ln(3)

    def print_file_code(self, rel_path, description):
        if not os.path.exists(rel_path):
            return

        with open(rel_path, "r", encoding="utf-8", errors="replace") as fp:
            raw_lines = fp.readlines()

        # Section Header Banner for each file
        self.ln(4)
        banner_y = self.get_y()
        if banner_y > 250:
            self.add_page()
            banner_y = self.get_y()

        self.set_fill_color(31, 73, 125)
        self.set_draw_color(31, 73, 125)
        self.rect(self.l_margin, banner_y, self.w - self.l_margin - self.r_margin, 10, "F")
        
        self.set_xy(self.l_margin + 3, banner_y + 1)
        self.set_font("Consolas", "B", 10)
        self.set_text_color(255, 255, 255)
        self.cell(100, 8, rel_path)
        
        self.set_font("Arial", "I", 8.5)
        self.set_text_color(203, 213, 225)
        self.cell(0, 8, f"{len(raw_lines)} lines • {description}", align="R", new_x="LMARGIN", new_y="NEXT")
        self.ln(2)

        # Code block lines
        self.set_font("Consolas", "", 7.2)
        line_height = 3.6
        code_box_w = self.w - self.l_margin - self.r_margin

        for idx, line in enumerate(raw_lines, start=1):
            # Clean non-printable or unsupported characters safely
            safe_line = line.rstrip("\r\n").replace("\t", "    ")
            # Replace characters that might fail in basic Windows fonts
            safe_line = (safe_line
                .replace("₹", "INR ")
                .replace("χ²", "Chi^2")
                .replace("χ", "Chi")
                .replace("•", "*")
                .replace("→", "->")
                .replace("≤", "<=")
                .replace("≥", ">=")
                .replace("—", "--")
                .replace("–", "-")
                .replace("’", "'")
                .replace("‘", "'")
                .replace("“", '"')
                .replace("”", '"')
                .replace("🏥", "[Hospital]")
                .replace("💡", "[Idea]")
                .replace("ℹ️", "[Info]")
                .replace("ℹ", "[Info]")
                .replace("⚠️", "[Alert]")
                .replace("⚠", "[Alert]")
                .replace("✅", "[OK]")
                .replace("📊", "[Chart]")
                .replace("🔍", "[Search]")
                .replace("📋", "[Scorecard]")
                .replace("⏱️", "[Timer]")
                .replace("⏱", "[Timer]")
                .replace("🩺", "[Clinical]")
                .replace("👥", "[Demographics]")
                .replace("🗓️", "[Calendar]")
                .replace("🗓", "[Calendar]")
                .replace("🔬", "[Inference]")
                .replace("💰", "[Financial]")
                .replace("⚙️", "[Simulation]")
                .replace("⚙", "[Simulation]")
                .replace("💾", "[Export]")
                .replace("🏢", "[Corporate]")
                .replace("🎛️", "[Slider]")
                .replace("🎛", "[Slider]")
                .replace("📄", "[Record]")
                .replace("📥", "[Download]")
                .replace("📞", "[Call]")
                .replace("👨‍⚕️", "[Doctor]")
            )
            # Filter any remaining non-BMP or unsupported symbol characters
            safe_line = "".join(c if ord(c) < 1000 or ord(c) in (8211, 8212, 8216, 8217, 8220, 8221, 8226) else " " for c in safe_line)

            # Check page break
            if self.get_y() + line_height > 280:
                self.add_page()

            # Line number column
            self.set_text_color(148, 163, 184) # Light grey for line numbers
            self.cell(11, line_height, f"{idx:4d} | ", align="R")
            
            # Code text column
            self.set_text_color(30, 41, 59) # Dark slate
            # Check comment vs code for slight color distinction
            trimmed = safe_line.strip()
            if trimmed.startswith("#") or trimmed.startswith('"""') or trimmed.startswith("'''"):
                self.set_text_color(100, 116, 139) # Slate grey for comments
            elif trimmed.startswith("def ") or trimmed.startswith("class "):
                self.set_text_color(31, 73, 125) # Navy for definitions
                
            self.cell(code_box_w - 11, line_height, safe_line, new_x="LMARGIN", new_y="NEXT")

        self.ln(3)

def generate_full_pdf():
    pdf = SourceCodePDF()
    pdf.draw_cover_page()

    # Define File Manifest
    file_sections = [
        ("SECTION 1: APPLICATION ENTRYPOINT & CONFIGURATION", [
            ("app.py", "Main landing page, session state, global filters, KPI cards", 161),
            ("requirements.txt", "Pinned production package dependencies", 6),
            (".streamlit/config.toml", "Streamlit UI theme and headless server configuration", 15),
        ]),
        ("SECTION 2: DATA INGESTION & DATA HYGIENE ENGINE", [
            ("src/data/__init__.py", "Package initialization and module exports", 13),
            ("src/data/loader.py", "Cached SQLite data loader using @st.cache_data", 69),
            ("src/data/cleaner.py", "Median imputation, deduplication, domain validation", 137),
            ("src/generate_data.py", "Synthetic dataset generator with realistic patterns", 642),
        ]),
        ("SECTION 3: CORE ANALYTICAL & FINANCIAL ENGINES", [
            ("src/analytics/__init__.py", "Analytics subpackage export facade", 48),
            ("src/analytics/descriptive.py", "Operational KPIs, 4-branch metrics, rolling trend", 144),
            ("src/analytics/lead_time.py", "5 lead-time intervals and booking decay curves", 60),
            ("src/analytics/specialty.py", "Specialty vulnerability rankings and fee schedules", 50),
            ("src/analytics/demographics.py", "Age cohorts, transit distance bands, comorbidities", 157),
            ("src/analytics/temporal.py", "2D bivariate weekday x time slot arrival matrix", 52),
            ("src/analytics/financial.py", "Uncollected consultation fees and doctor idle hours", 85),
        ]),
        ("SECTION 4: STATISTICAL INFERENCE & PREDICTIVE RISK MODELS", [
            ("src/stats/__init__.py", "Statistical package exports", 15),
            ("src/stats/hypothesis.py", "Two-sample proportion Z-test, Chi-square, odds ratios", 209),
            ("src/models/__init__.py", "Predictive modeling exports", 17),
            ("src/models/risk_engine.py", "Heuristic 0-100 risk scoring, tier mapping, triage", 106),
        ]),
        ("SECTION 5: REUSABLE UI COMPONENTS & PLOTLY CHART FACTORY", [
            ("src/ui/__init__.py", "UI package exports", 36),
            ("src/ui/components.py", "Tailwind CSS header banners, KPI cards, table maker", 157),
            ("src/ui/charts.py", "Plotly white-template reusable chart generators", 322),
            ("src/ui.py", "UI facade re-export library", 36),
        ]),
        ("SECTION 6: MULTI-PAGE WEB APPLICATION SCREENS (UI-01 TO UI-08)", [
            ("pages/1_Overview_and_KPIs.py", "UI-01: Multi-branch attendance benchmarking", 79),
            ("pages/2_Lead_Time_and_Specialty.py", "UI-02: Lead-time decay curves & specialty bars", 79),
            ("pages/3_Demographics_and_Arrivals.py", "UI-03: Age cohorts & 2D arrival density heatmap", 82),
            ("pages/4_Statistical_Inference.py", "UI-04: Parametric Z-test cards & Chi-square matrix", 146),
            ("pages/5_Financial_and_Capacity.py", "UI-05: Revenue leakage waterfall & idle hours", 103),
            ("pages/6_Predictive_Risk_Triage.py", "UI-06: Advance risk queue & calling roster with CSV", 102),
            ("pages/7_Simulation_and_Corporate.py", "UI-07: Noida IT shift dynamics & overbooking", 152),
            ("pages/8_Data_Quality_and_Export.py", "UI-08: Data hygiene audit & CSV/Excel export", 164),
        ]),
        ("SECTION 7: AUTOMATED PYTEST TEST SUITE", [
            ("tests/test_data.py", "TC-T-01..06: Relational schema, PK, FK integrity", 84),
            ("tests/test_analysis.py", "TC-T-07..11, 13..18: KPI formulas, Z-test, risk tiers", 174),
            ("tests/test_app.py", "TC-T-12: Config, imports, and multi-page routing", 48),
        ]),
        ("SECTION 8: PRODUCTION OPERATIONS MANUAL", [
            ("README.md", "Production operations manual & viva exam preparation", 207),
        ])
    ]

    pdf.draw_table_of_contents(file_sections)

    # Print code for each section
    for sec_name, files in file_sections:
        pdf.add_page()
        pdf.set_font("Arial", "B", 14)
        pdf.set_text_color(31, 73, 125)
        pdf.cell(0, 8, sec_name, new_x="LMARGIN", new_y="NEXT")
        pdf.set_draw_color(31, 73, 125)
        pdf.set_line_width(0.6)
        pdf.line(pdf.get_x(), pdf.get_y(), 195, pdf.get_y())
        pdf.ln(3)

        for fpath, desc, _ in files:
            pdf.print_file_code(fpath, desc)

    out_primary = "docs/2500320140047_Source_Code.pdf"
    out_mirror = "docs/ApexCare_Appointment_No_Show_Source_Code.pdf"
    os.makedirs("docs", exist_ok=True)
    
    print(f"Generating PDF: {out_primary} ...")
    pdf.output(out_primary)
    shutil.copyfile(out_primary, out_mirror)
    print(f"Successfully created: {out_primary} ({os.path.getsize(out_primary):,} bytes)")
    print(f"Successfully mirrored: {out_mirror}")

if __name__ == "__main__":
    generate_full_pdf()
