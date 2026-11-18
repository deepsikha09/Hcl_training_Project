"""
Automated Pytest Suite: Data Layer & Relational Integrity (TC-T-01 to TC-T-06)
Verifies SQLite app.db schema, row volumes, primary keys, foreign keys, and cleaning.
"""

import os
import sqlite3
from pathlib import Path
import pytest

BASE_DIR = Path(__file__).resolve().parent.parent
DB_PATH = BASE_DIR / "data" / "app.db"

@pytest.fixture(scope="module")
def db_conn():
    """Provides an active connection to the production SQLite database with FKs enabled."""
    assert os.path.exists(DB_PATH), f"Database file not found at {DB_PATH}"
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON;")
    yield conn
    conn.close()

def test_tc_t_01_fact_appointments_volume(db_conn):
    """TC-T-01: Verify fact_appointments row count meets or exceeds the planned minimum (150+)."""
    cur = db_conn.cursor()
    cur.execute("SELECT COUNT(*) FROM fact_appointments;")
    count = cur.fetchone()[0]
    assert count >= 150, f"Expected at least 150 fact rows, found {count}"

def test_tc_t_02_analytics_view_columns(db_conn):
    """TC-T-02: Verify enriched analytical view vw_appointment_analytics contains exactly 36 columns."""
    cur = db_conn.cursor()
    cur.execute("PRAGMA table_info(vw_appointment_analytics);")
    columns = cur.fetchall()
    assert len(columns) == 36, f"Expected exactly 36 columns in view, found {len(columns)}"

def test_tc_t_03_primary_key_uniqueness(db_conn):
    """TC-T-03: Verify primary keys are 100% unique across all tables (zero duplicates)."""
    cur = db_conn.cursor()
    tables = [
        ("dim_clinics", "clinic_id"),
        ("dim_specialties", "specialty_id"),
        ("dim_doctors", "doctor_id"),
        ("dim_patients", "patient_id"),
        ("fact_appointments", "appointment_id")
    ]
    for table, pk in tables:
        cur.execute(f"SELECT COUNT({pk}), COUNT(DISTINCT {pk}) FROM {table};")
        total, distinct = cur.fetchone()
        assert total == distinct, f"Duplicate primary keys detected in {table}!"

def test_tc_t_04_foreign_key_integrity(db_conn):
    """TC-T-04: Execute PRAGMA foreign_key_check to verify zero relational orphan records."""
    cur = db_conn.cursor()
    cur.execute("PRAGMA foreign_key_check;")
    violations = cur.fetchall()
    assert len(violations) == 0, f"Foreign key constraint violations detected: {violations}"

def test_tc_t_05_check_constraints_and_ranges(db_conn):
    """TC-T-05: Verify data domain ranges (lead_time_days >= 0, status in Show/No-Show)."""
    cur = db_conn.cursor()
    # Negative lead time check
    cur.execute("SELECT COUNT(*) FROM fact_appointments WHERE lead_time_days < 0;")
    assert cur.fetchone()[0] == 0, "Negative lead time detected in database!"
    
    # Valid status check
    cur.execute("SELECT DISTINCT status FROM fact_appointments;")
    statuses = {row[0] for row in cur.fetchall()}
    assert statuses.issubset({"Show", "No-Show", "Cancelled"}), f"Invalid status found: {statuses}"

def test_tc_t_06_dimension_minimum_rows(db_conn):
    """TC-T-06: Verify master dimension tables satisfy minimum planned capacity."""
    cur = db_conn.cursor()
    cur.execute("SELECT COUNT(*) FROM dim_clinics;")
    assert cur.fetchone()[0] >= 4, "Clinics master row count below 4"
    
    cur.execute("SELECT COUNT(*) FROM dim_specialties;")
    assert cur.fetchone()[0] >= 6, "Specialties master row count below 6"
    
    cur.execute("SELECT COUNT(*) FROM dim_doctors;")
    assert cur.fetchone()[0] >= 12, "Doctors master row count below 12"
    
    cur.execute("SELECT COUNT(*) FROM dim_patients;")
    assert cur.fetchone()[0] >= 100, "Patients master row count below 100"
