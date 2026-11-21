"""
Create Clean ZIP Packages for LMS Submission:
1. 2500320140047_Deepshikha_LMS_Complete_Submission.zip (All documents)
2. 2500320140047_Deepshikha_Project_Codebase.zip (Complete project source code)
"""

import os
import zipfile

BASE_DIR = r"c:\Users\deeps\Downloads\Appointment No Show Analysis"
DOCS_DIR = os.path.join(BASE_DIR, "LMS_Submission_Documents")

# 1. Zip all LMS submission documents
zip_docs_path = os.path.join(BASE_DIR, "2500320140047_Deepshikha_LMS_Complete_Submission.zip")
print(f"Creating documents ZIP: {zip_docs_path} ...")
with zipfile.ZipFile(zip_docs_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for fname in sorted(os.listdir(DOCS_DIR)):
        fpath = os.path.join(DOCS_DIR, fname)
        if os.path.isfile(fpath):
            zf.write(fpath, arcname=fname)
print(f"Documents ZIP created successfully ({os.path.getsize(zip_docs_path):,} bytes).")

# 2. Zip project codebase (excluding git, venv, pycache, zip files, LMS_Submission_Documents)
zip_code_path = os.path.join(BASE_DIR, "2500320140047_Deepshikha_Project_Codebase.zip")
print(f"Creating codebase ZIP: {zip_code_path} ...")

exclude_dirs = {'.git', '.venv', 'venv', '__pycache__', '.pytest_cache', 'LMS_Submission_Documents'}
exclude_exts = {'.zip'}

with zipfile.ZipFile(zip_code_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for root, dirs, files in os.walk(BASE_DIR):
        # prune exclude directories
        dirs[:] = [d for d in dirs if d not in exclude_dirs]
        for file in files:
            ext = os.path.splitext(file)[1].lower()
            if ext in exclude_exts:
                continue
            abs_path = os.path.join(root, file)
            rel_path = os.path.relpath(abs_path, BASE_DIR)
            zf.write(abs_path, arcname=rel_path)

print(f"Codebase ZIP created successfully ({os.path.getsize(zip_code_path):,} bytes).")
