"""
Database Ingestion & Analytical Caching Module
Queries SQLite database views and master tables with Streamlit caching.
"""

import os
import sqlite3
from typing import List, Dict, Any, Optional

DEFAULT_DB_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "data", "app.db"))

def get_connection(db_path: Optional[str] = None) -> sqlite3.Connection:
    """Creates a connection to SQLite with row-factory enabled."""
    target_path = db_path or DEFAULT_DB_PATH
    if not os.path.exists(target_path):
        raise FileNotFoundError(f"Database file not found at: {target_path}")
    conn = sqlite3.connect(target_path)
    conn.row_factory = sqlite3.Row
    return conn

def _fetch_view_records(db_path: Optional[str] = None) -> List[Dict[str, Any]]:
    """Internal function to query vw_appointment_analytics without caching."""
    conn = get_connection(db_path)
    try:
        cur = conn.cursor()
        cur.execute("SELECT * FROM vw_appointment_analytics ORDER BY appointment_id ASC;")
        rows = cur.fetchall()
        # Convert sqlite3.Row to standard dicts
        records = [dict(r) for r in rows]
        return records
    finally:
        conn.close()

# Try caching with Streamlit if running within Streamlit runtime
try:
    import streamlit as st
    @st.cache_data(ttl=3600, show_spinner=False)
    def load_data(db_path: Optional[str] = None) -> List[Dict[str, Any]]:
        """Cached loader for analytical records (36-column view)."""
        return _fetch_view_records(db_path)
except Exception:
    def load_data(db_path: Optional[str] = None) -> List[Dict[str, Any]]:
        """Fallback loader for non-Streamlit environments."""
        return _fetch_view_records(db_path)

def get_table_names(db_path: Optional[str] = None) -> List[str]:
    """Returns list of all tables in the SQLite database."""
    conn = get_connection(db_path)
    try:
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%';")
        tables = [r[0] for r in cur.fetchall()]
        return tables
    finally:
        conn.close()

def get_table_records(table_name: str, db_path: Optional[str] = None, limit: Optional[int] = None) -> List[Dict[str, Any]]:
    """Returns all or limited records from a given table."""
    conn = get_connection(db_path)
    try:
        cur = conn.cursor()
        sql = f"SELECT * FROM {table_name}"
        if limit:
            sql += f" LIMIT {int(limit)}"
        cur.execute(sql)
        rows = cur.fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()
