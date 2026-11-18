import os
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_rtm_excel(output_path):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb = Workbook()
    
    # ----------------------------------------------------
    # Sheet 1: Traceability Matrix
    # ----------------------------------------------------
    ws_rtm = wb.active
    ws_rtm.title = "Traceability_Matrix"
    ws_rtm.views.sheetView[0].showGridLines = True
    
    # Styling definitions
    font_title = Font(name="Calibri", size=14, bold=True, color="1F497D")
    font_sub = Font(name="Calibri", size=10, italic=True, color="595959")
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=9.5, color="333333")
    font_bold_data = Font(name="Calibri", size=9.5, bold=True, color="1F497D")
    
    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_future_header = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    fill_stripe = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    # Priority fills
    fill_must = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid") # Soft Orange
    fill_should = PatternFill(start_color="EBF1F5", end_color="EBF1F5", fill_type="solid") # Soft Blue
    fill_could = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") # Soft Green
    
    thin_border_side = Side(border_style="thin", color="D9D9D9")
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    # Title Rows
    ws_rtm.merge_cells("A1:L1")
    ws_rtm["A1"] = "APEXCARE HEALTH NETWORK — REQUIREMENTS TRACEABILITY MATRIX (RTM)"
    ws_rtm["A1"].font = font_title
    ws_rtm["A1"].alignment = Alignment(vertical="center")
    
    ws_rtm.merge_cells("A2:L2")
    ws_rtm["A2"] = "Project: Appointment No-Show Analysis | Student: Deepshikha (Roll: 2500320140047) | Mentor: Alpna Lodhi mam | Step: 04"
    ws_rtm["A2"].font = font_sub
    ws_rtm["A2"].alignment = Alignment(vertical="center")
    
    headers = [
        "OQ ID", "HMW ID", "Req ID (BR)", "Requirement Description", "Stakeholder (SH)",
        "Priority", "Status", "Feature (PRD)", "DB Table.Column", "User Story", "UI Screen", "Test Case (TC-F / TC-T)"
    ]
    
    header_row_idx = 4
    ws_rtm.row_dimensions[header_row_idx].height = 28
    
    for col_idx, header in enumerate(headers, start=1):
        cell = ws_rtm.cell(row=header_row_idx, column=col_idx, value=header)
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        # Distinguish current vs future columns
        if col_idx <= 7:
            cell.fill = fill_header
        else:
            cell.fill = fill_future_header
        cell.border = cell_border
        
    rtm_data = [
        ("OQ-01", "HMW-01", "BR-01", "The business needs to track historical and ongoing appointment attendance rates across all clinic locations and time periods so that clinic administrators can detect systemic operational variances and benchmark facility performance.", "SH-01", "Must", "Identified"),
        ("OQ-02", "HMW-02", "BR-02", "The business needs to analyze appointment attendance as a function of booking lead time intervals so that scheduling coordinators can identify cutoff thresholds where attendance commitment significantly drops.", "SH-01, SH-03", "Must", "Identified"),
        ("OQ-03", "HMW-03", "BR-03", "The business needs to evaluate specialty-specific attendance variations so that clinical leadership can uncover department bottlenecks and adjust doctor consultation rosters according to departmental show rates.", "SH-02", "Must", "Identified"),
        ("OQ-04", "HMW-04", "BR-04", "The business needs to segment appointment attendance by patient demographic cohorts (age brackets, gender, and residential proximity) so that clinical outreach can be customized for vulnerable or distant patient groups.", "SH-01, SH-03", "Should", "Identified"),
        ("OQ-05", "HMW-05", "BR-05", "The business needs to measure the comparative attendance lift between patients who received automated digital reminders and those who did not so that communication channels can be empirically validated and optimized.", "SH-03", "Must", "Identified"),
        ("OQ-06", "HMW-06", "BR-06", "The business needs to examine appointment arrival patterns across days of the week and consultation time slots so that facility supervisors can schedule adequate staff and prevent waiting lobby congestion during peak periods.", "SH-01", "Should", "Identified"),
        ("OQ-07", "HMW-07", "BR-07", "The business needs to monitor consultation compliance among patients diagnosed with chronic conditions (such as Hypertension and Diabetes) so that clinical teams can prevent dangerous gaps in disease management.", "SH-02", "Must", "Identified"),
        ("OQ-08", "HMW-08", "BR-08", "The business needs to distinguish between first-time appointment misses and habitual repeat no-show offenders so that front-desk staff can avoid expending high-touch phone calls on one-off accidental misses.", "SH-03", "Should", "Identified"),
        ("OQ-09", "HMW-09", "BR-09", "The business needs to quantify the financial revenue loss and wasted doctor consultation hours resulting from unfilled appointments so that executive leadership can determine the monetary return on investment for attendance improvement initiatives.", "SH-01, SH-02", "Must", "Identified"),
        ("OQ-10", "HMW-10", "BR-10", "The business needs to calculate an advance risk tier (High, Medium, Low) for upcoming scheduled appointments so that front-desk staff can objectively prioritize high-risk patients for manual telephonic reconfirmation.", "SH-03", "Must", "Identified"),
        ("OQ-04, OQ-06", "HMW-11", "BR-11", "The business needs to correlate corporate employee shift timings with occupational health checkup completion so that corporate wellness partners can schedule on-site checkups during optimal work windows.", "SH-04", "Could", "Identified"),
        ("OQ-01, OQ-06", "HMW-01, HMW-06", "BR-12", "The business needs to simulate the operational impact of calibrated overbooking quotas based on historical no-show probabilities so that clinic managers can recover lost capacity without exceeding lobby seating thresholds.", "SH-01", "Could", "Identified")
    ]
    
    current_row = header_row_idx + 1
    for r_idx, row_values in enumerate(rtm_data):
        row_num = current_row + r_idx
        ws_rtm.row_dimensions[row_num].height = 42
        bg = fill_stripe if r_idx % 2 == 1 else fill_white
        
        for c_idx in range(1, 13):
            cell = ws_rtm.cell(row=row_num, column=c_idx)
            cell.border = cell_border
            cell.font = font_data
            cell.fill = bg
            
            if c_idx <= len(row_values):
                val = row_values[c_idx - 1]
                cell.value = val
                
                # Alignments and specific styles
                if c_idx in [1, 2, 3, 5, 7]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    if c_idx == 3:
                        cell.font = font_bold_data
                elif c_idx == 6:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    if val == "Must":
                        cell.fill = fill_must
                    elif val == "Should":
                        cell.fill = fill_should
                    elif val == "Could":
                        cell.fill = fill_could
                elif c_idx == 4:
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                # Empty future lifecycle columns
                cell.value = ""
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
    # Enable AutoFilter on the active headers
    ws_rtm.auto_filter.ref = f"A{header_row_idx}:L{current_row + len(rtm_data) - 1}"
    
    # Set explicit column widths
    col_widths = {
        "A": 12, # OQ ID
        "B": 14, # HMW ID
        "C": 12, # Req ID (BR)
        "D": 45, # Requirement Description
        "E": 14, # Stakeholder
        "F": 12, # Priority
        "G": 13, # Status
        "H": 18, # Feature (PRD)
        "I": 18, # DB Table.Column
        "J": 15, # User Story
        "K": 15, # UI Screen
        "L": 22  # Test Case
    }
    for col_letter, width in col_widths.items():
        ws_rtm.column_dimensions[col_letter].width = width
        
    # Freeze Panes at header
    ws_rtm.freeze_panes = f"A{header_row_idx + 1}"
    
    # ----------------------------------------------------
    # Sheet 2: Coverage Analysis & Audit
    # ----------------------------------------------------
    ws_cov = wb.create_sheet(title="Coverage_Analysis")
    ws_cov.views.sheetView[0].showGridLines = True
    
    ws_cov.merge_cells("A1:F1")
    ws_cov["A1"] = "REQUIREMENTS TRACEABILITY AUDIT & COVERAGE SUMMARY"
    ws_cov["A1"].font = font_title
    ws_cov["A1"].alignment = Alignment(vertical="center")
    
    ws_cov.merge_cells("A2:F2")
    ws_cov["A2"] = "Verification of Bidirectional Linkages Across Operational Questions, HMWs, and Business Requirements"
    ws_cov["A2"].font = font_sub
    ws_cov["A2"].alignment = Alignment(vertical="center")
    
    # Table 1: Key Metrics Summary
    ws_cov.cell(row=4, column=1, value="Metric Category").font = font_header
    ws_cov.cell(row=4, column=1).fill = fill_header
    ws_cov.cell(row=4, column=2, value="Count / Value").font = font_header
    ws_cov.cell(row=4, column=2).fill = fill_header
    ws_cov.cell(row=4, column=3, value="Audit Benchmark").font = font_header
    ws_cov.cell(row=4, column=3).fill = fill_header
    ws_cov.cell(row=4, column=4, value="Status").font = font_header
    ws_cov.cell(row=4, column=4).fill = fill_header
    
    metrics = [
        ("Total Operational Questions (OQs)", 10, "10 Defined", "100% Covered"),
        ("Total 'How Might We' Statements (HMWs)", 11, "11 Formulated", "100% Covered"),
        ("Total Business Requirements (BRs)", 12, "10-15 Target", "12 Active"),
        ("Must Have Requirements", "7 (58.3%)", "~60% Target", "Optimal Feasibility"),
        ("Should Have Requirements", "3 (25.0%)", "20-25% Target", "Optimal Buffer"),
        ("Could Have Requirements", "2 (16.7%)", "10-15% Target", "Controlled Scope"),
        ("Orphaned HMWs (Without BR)", 0, "0 Allowed", "PASS (Zero Gaps)"),
        ("Orphaned BRs (Without Source HMW)", 0, "0 Allowed", "PASS (Zero Scope Creep)"),
        ("Uncovered OQs", 0, "0 Allowed", "PASS (100% Solvable)")
    ]
    
    fill_pass = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") # Soft Green
    font_pass = Font(name="Calibri", size=9.5, bold=True, color="006100")
    
    for r_i, (cat, val, bench, stat) in enumerate(metrics, start=5):
        c1 = ws_cov.cell(row=r_i, column=1, value=cat)
        c2 = ws_cov.cell(row=r_i, column=2, value=val)
        c3 = ws_cov.cell(row=r_i, column=3, value=bench)
        c4 = ws_cov.cell(row=r_i, column=4, value=stat)
        
        for c in [c1, c2, c3, c4]:
            c.border = cell_border
            c.font = font_data
            c.alignment = Alignment(vertical="center")
            
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c4.alignment = Alignment(horizontal="center", vertical="center")
        if "PASS" in stat or "100%" in stat or "Optimal" in stat:
            c4.fill = fill_pass
            c4.font = font_pass
            
    # Table 2: OQ Coverage Breakdown
    start_oq_row = 16
    ws_cov.cell(row=start_oq_row, column=1, value="OQ ID").font = font_header
    ws_cov.cell(row=start_oq_row, column=1).fill = fill_header
    ws_cov.cell(row=start_oq_row, column=2, value="Operational Question Summary").font = font_header
    ws_cov.cell(row=start_oq_row, column=2).fill = fill_header
    ws_cov.cell(row=start_oq_row, column=3, value="Covering HMW").font = font_header
    ws_cov.cell(row=start_oq_row, column=3).fill = fill_header
    ws_cov.cell(row=start_oq_row, column=4, value="Fulfilling BR").font = font_header
    ws_cov.cell(row=start_oq_row, column=4).fill = fill_header
    ws_cov.cell(row=start_oq_row, column=5, value="Traceability Status").font = font_header
    ws_cov.cell(row=start_oq_row, column=5).fill = fill_header
    
    oq_breakdown = [
        ("OQ-01", "Overall baseline attendance rate & temporal trend", "HMW-01", "BR-01, BR-12", "VERIFIED"),
        ("OQ-02", "Booking lead time correlation with non-attendance", "HMW-02", "BR-02", "VERIFIED"),
        ("OQ-03", "Specialty-specific no-show distributions", "HMW-03", "BR-03", "VERIFIED"),
        ("OQ-04", "Demographic factors (age, gender, transit distance)", "HMW-04, HMW-11", "BR-04, BR-11", "VERIFIED"),
        ("OQ-05", "SMS/WhatsApp digital reminder efficacy lift", "HMW-05", "BR-05", "VERIFIED"),
        ("OQ-06", "Day-of-week and time-slot arrival friction", "HMW-06, HMW-11", "BR-06, BR-11, BR-12", "VERIFIED"),
        ("OQ-07", "Chronic comorbidity compliance (Diabetes/Hypertension)", "HMW-07", "BR-07", "VERIFIED"),
        ("OQ-08", "Repeat no-show behavioral distribution (80/20 rule)", "HMW-08", "BR-08", "VERIFIED"),
        ("OQ-09", "Monetary loss and doctor-hour capacity waste", "HMW-09", "BR-09", "VERIFIED"),
        ("OQ-10", "Composite risk tiering engine for advance triage", "HMW-10", "BR-10", "VERIFIED")
    ]
    
    for r_i, (oq, desc, hmw, br, st) in enumerate(oq_breakdown, start=start_oq_row+1):
        c1 = ws_cov.cell(row=r_i, column=1, value=oq)
        c2 = ws_cov.cell(row=r_i, column=2, value=desc)
        c3 = ws_cov.cell(row=r_i, column=3, value=hmw)
        c4 = ws_cov.cell(row=r_i, column=4, value=br)
        c5 = ws_cov.cell(row=r_i, column=5, value=st)
        
        for c in [c1, c2, c3, c4, c5]:
            c.border = cell_border
            c.font = font_data
            c.alignment = Alignment(vertical="center")
            
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c3.alignment = Alignment(horizontal="center", vertical="center")
        c4.alignment = Alignment(horizontal="center", vertical="center")
        c5.alignment = Alignment(horizontal="center", vertical="center")
        c5.fill = fill_pass
        c5.font = font_pass
        
    ws_cov.column_dimensions["A"].width = 14
    ws_cov.column_dimensions["B"].width = 45
    ws_cov.column_dimensions["C"].width = 18
    ws_cov.column_dimensions["D"].width = 20
    ws_cov.column_dimensions["E"].width = 18
    
    # Save Workbook
    wb.save(output_path)
    print(f"RTM Excel successfully created at: {output_path}")

if __name__ == "__main__":
    out_dir = r"c:\Users\deeps\Downloads\Appointment No Show Analysis\docs"
    file_name = "04_Traceability_Matrix.xlsx"
    full_path = os.path.join(out_dir, file_name)
    create_rtm_excel(full_path)
