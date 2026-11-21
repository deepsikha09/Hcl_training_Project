"""
Script to update metadata across all project scripts, docs, and deliverables:
- Batch ID: ABES_OND26 MCA 2ND YEAR_DATA ANALYTICS_BATCH 01
- Project ID: P101
- Project Mentor: Dinesh Sereno
- GitHub Repo: https://github.com/deepsikha09/Hcl_training_Project
"""

import os
import re
import glob

def update_file(filepath):
    with open(filepath, "r", encoding="utf-8") as f:
        content = f.read()

    original = content
    
    # 1. Update Faculty Mentor mentions
    content = content.replace("Dinesh Sereno, Assistant Professor, Dept. of MCA", "Dinesh Sereno, Assistant Professor, Dept. of MCA")
    content = content.replace("Dinesh Sereno | Department of Computer Applications, ABES Engineering College", "Dinesh Sereno | Department of Computer Applications, ABES Engineering College")
    content = content.replace("Dinesh Sereno | Department of Computer Applications, ABESEC", "Dinesh Sereno | Department of Computer Applications, ABESEC")
    content = content.replace("Dinesh Sereno (Assistant Professor, Dept of MCA)", "Dinesh Sereno (Assistant Professor, Dept of MCA)")
    content = content.replace("Dinesh Sereno (Faculty)", "Dinesh Sereno (Faculty)")
    content = content.replace("Approved (Dinesh Sereno)", "Approved (Dinesh Sereno)")
    content = content.replace('"Dinesh Sereno", "Faculty Project Mentor (Dept of MCA)"', '"Dinesh Sereno", "Faculty Project Mentor (Dept of MCA)"')
    content = content.replace('"Dinesh Sereno", "Department of Computer Applications, ABESEC"', '"Dinesh Sereno", "Department of Computer Applications, ABESEC"')
    content = content.replace("Dinesh Sereno", "Dinesh Sereno")
    content = content.replace("Dinesh Sereno", "Dinesh Sereno")

    # 2. Update Repo URL mentions
    content = content.replace("deepsikha09/Hcl_training_Project", "deepsikha09/Hcl_training_Project")
    content = content.replace("https://github.com/deepsikha09/Hcl_training_Project.git", "https://github.com/deepsikha09/Hcl_training_Project.git")
    content = content.replace("https://github.com/deepsikha09/Hcl_training_Project", "https://github.com/deepsikha09/Hcl_training_Project")

    # 3. Add Project ID & Batch to standard h_table in step doc scripts
    old_htable = """    h_table = doc.add_table(rows=6, cols=2)
    h_widths = [2.2, 4.3]
    h_data = [
        ("Project ID & Batch:", "P101 | ABES_OND26 MCA 2ND YEAR_DATA ANALYTICS_BATCH 01"),
        ("Project Title:", "Appointment No-Show Analysis & Optimization System"),"""
    new_htable = """    h_table = doc.add_table(rows=6, cols=2)
    h_widths = [2.2, 4.3]
    h_data = [
        ("Project ID & Batch:", "P101 | ABES_OND26 MCA 2ND YEAR_DATA ANALYTICS_BATCH 01"),
        ("Project Title:", "Appointment No-Show Analysis & Optimization System"),"""
    if old_htable in content:
        content = content.replace(old_htable, new_htable)

    # 4. Step 4 Excel
    old_s4 = 'ws_rtm["A2"] = "Project ID: P101 | Batch: ABES_OND26 MCA 2ND YEAR_DATA ANALYTICS_BATCH 01 | Student: Deepshikha (Roll: 2500320140047) | Mentor: Dinesh Sereno | Step: 04"'
    new_s4 = 'ws_rtm["A2"] = "Project ID: P101 | Batch: ABES_OND26 MCA 2ND YEAR_DATA ANALYTICS_BATCH 01 | Student: Deepshikha (Roll: 2500320140047) | Mentor: Dinesh Sereno | Step: 04"'
    if old_s4 in content:
        content = content.replace(old_s4, new_s4)

    # 5. Step 15 Doc
    old_s15 = """    meta_table = doc.add_table(rows=8, cols=2)
    meta_data = [
        ("Project ID", "P101"),
        ("Batch ID", "ABES_OND26 MCA 2ND YEAR_DATA ANALYTICS_BATCH 01"),
        ("Document Identifier", "APEX-BLD-RPT-STEP15-v1.0"),
        ("Project Title", "Outpatient Appointment No-Show Analysis & Capacity Optimization System"),
        ("Healthcare Network Modeled", "ApexCare Health Network, Noida (Sectors 62, 18, 137, Greater Noida)"),
        ("Learner / Developer", "Deepshikha | Roll No: 2500320140047 | MCA 2nd Year (Data Analytics)"),
        ("Faculty Project Mentor", "Dinesh Sereno | Department of Computer Applications, ABES Engineering College"),
        ("Build & Test Status", "18/18 Build Tasks Complete • 20/20 Pytests Passing (100% Pass Rate) • Ready for Step 16")
    ]"""
    new_s15 = """    meta_table = doc.add_table(rows=8, cols=2)
    meta_data = [
        ("Project ID", "P101"),
        ("Batch ID", "ABES_OND26 MCA 2ND YEAR_DATA ANALYTICS_BATCH 01"),
        ("Document Identifier", "APEX-BLD-RPT-STEP15-v1.0"),
        ("Project Title", "Outpatient Appointment No-Show Analysis & Capacity Optimization System"),
        ("Healthcare Network Modeled", "ApexCare Health Network, Noida (Sectors 62, 18, 137, Greater Noida)"),
        ("Learner / Developer", "Deepshikha | Roll No: 2500320140047 | MCA 2nd Year (Data Analytics)"),
        ("Faculty Project Mentor", "Dinesh Sereno | Department of Computer Applications, ABES Engineering College"),
        ("Build & Test Status", "18/18 Build Tasks Complete • 20/20 Pytests Passing (100% Pass Rate) • Ready for Step 16")
    ]"""
    if old_s15 in content:
        content = content.replace(old_s15, new_s15)

    # 6. Step 16 Doc
    old_s16 = """    meta_table = doc.add_table(rows=9, cols=2)
    meta_data = [
        ("Project ID", "P101"),
        ("Batch ID", "ABES_OND26 MCA 2ND YEAR_DATA ANALYTICS_BATCH 01"),
        ("Document Identifier", "APEX-DEP-GUIDE-STEP16-v1.0"),
        ("Project Title", "Outpatient Appointment No-Show Analysis & Capacity Optimization System"),
        ("Healthcare Network Modeled", "ApexCare Health Network, Noida (Sectors 62, 18, 137, Greater Noida)"),
        ("Learner / Developer", "Deepshikha | Roll No: 2500320140047 | MCA 2nd Year (Data Analytics)"),
        ("Faculty Project Mentor", "Dinesh Sereno | Department of Computer Applications, ABESEC"),
        ("Target Cloud Host", "Streamlit Community Cloud (share.streamlit.io) • Ubuntu 22.04 LTS Container"),
        ("Deployment Target Coordinates", "Repo: deepsikha09/Hcl_training_Project | Branch: main | URL: https://2500320140047-appointment-noshow.streamlit.app")
    ]"""
    new_s16 = """    meta_table = doc.add_table(rows=9, cols=2)
    meta_data = [
        ("Project ID", "P101"),
        ("Batch ID", "ABES_OND26 MCA 2ND YEAR_DATA ANALYTICS_BATCH 01"),
        ("Document Identifier", "APEX-DEP-GUIDE-STEP16-v1.0"),
        ("Project Title", "Outpatient Appointment No-Show Analysis & Capacity Optimization System"),
        ("Healthcare Network Modeled", "ApexCare Health Network, Noida (Sectors 62, 18, 137, Greater Noida)"),
        ("Learner / Developer", "Deepshikha | Roll No: 2500320140047 | MCA 2nd Year (Data Analytics)"),
        ("Faculty Project Mentor", "Dinesh Sereno | Department of Computer Applications, ABESEC"),
        ("Target Cloud Host", "Streamlit Community Cloud (share.streamlit.io) • Ubuntu 22.04 LTS Container"),
        ("Deployment Target Coordinates", "Repo: deepsikha09/Hcl_training_Project | Branch: main | URL: https://2500320140047-appointment-noshow.streamlit.app")
    ]"""
    if old_s16 in content:
        content = content.replace(old_s16, new_s16)

    if content != original:
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(content)
        print(f"Updated: {filepath}")
    else:
        print(f"No changes: {filepath}")

def main():
    py_files = sorted(glob.glob("scripts/*.py"))
    for py_file in py_files:
        update_file(py_file)

if __name__ == "__main__":
    main()
