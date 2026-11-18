"""
Automated Pytest Suite: UI Components & Application Packaging (TC-T-12)
Verifies Streamlit configuration, ui.py components, and multi-page routing.
"""

import os
import sys
from pathlib import Path
import pytest

# Add parent directory to path so src modules can be imported
BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

def test_tc_t_12_streamlit_config_exists():
    """TC-T-12a: Verify .streamlit/config.toml exists and contains valid server and theme sections."""
    config_path = BASE_DIR / ".streamlit" / "config.toml"
    assert config_path.exists(), f"Configuration file not found at {config_path}"
    with open(config_path, "r", encoding="utf-8") as f:
        content = f.read()
    assert "[theme]" in content, "Missing [theme] section in config.toml"
    assert "primaryColor" in content, "Missing primaryColor definition"
    assert "[server]" in content, "Missing [server] section in config.toml"

def test_tc_t_12_ui_component_module_imports():
    """TC-T-12b: Verify src/ui.py component library imports cleanly and functions exist."""
    import src.ui as ui
    assert hasattr(ui, "render_tailwind_header"), "Missing render_tailwind_header function in src/ui.py"
    assert hasattr(ui, "render_kpi_card_row"), "Missing render_kpi_card_row function in src/ui.py"
    assert hasattr(ui, "render_insight_card"), "Missing render_insight_card function in src/ui.py"

def test_tc_t_12_dataset_artifacts_present():
    """TC-T-12c: Verify master data directory contains app.db, Excel file, and CSV exports."""
    data_dir = BASE_DIR / "data"
    required_files = [
        "app.db",
        "Appointment_No_Show_data.xlsx",
        "dim_clinics.csv",
        "dim_specialties.csv",
        "dim_doctors.csv",
        "dim_patients.csv",
        "fact_appointments.csv",
        "vw_appointment_analytics.csv",
        "validation_report.txt"
    ]
    for rf in required_files:
        p = data_dir / rf
        assert p.exists(), f"Required data artifact missing: {rf}"
