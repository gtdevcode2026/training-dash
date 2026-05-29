"""
SQLite persistence for transcript completion dates and full row metadata.

- completion_records: latest snapshot per (email, training) with completion date, IDs, zone,
  snapshot week, and how the date was derived (LMS vs inferred).
- Migrates legacy completion_backfill into completion_records when present.
"""

from __future__ import annotations

import sqlite3
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Optional

import pandas as pd


def _norm_email(s: object) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return str(s).strip().lower()


def varied_completion_date(snapshot_date_iso: str, email: object, training_name: str) -> str:
    """
    Pick a deterministic completion day (Mon–Sun) within the week of the snapshot date,
    so rows in the same weekly export don't all share one date. Stable for (email, training, week).
    """
    s = (snapshot_date_iso or "")[:10]
    try:
        snap = date.fromisoformat(s)
    except ValueError:
        snap = date.today()
    mon = snap
    while mon.weekday() != 0:
        mon -= timedelta(days=1)
    em = _norm_email(email)
    tn = (training_name or "").strip()
    h = hash((em, tn, s)) & 0xFFFFFFFF
    off = h % 7
    return (mon + timedelta(days=off)).isoformat()


def _parse_date(val: object) -> Optional[str]:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, pd.Timestamp):
        if pd.isna(val):
            return None
        return val.date().isoformat()
    if isinstance(val, datetime):
        return val.date().isoformat()
    if isinstance(val, date):
        return val.isoformat()
    ts = pd.to_datetime(val, errors="coerce")
    if pd.isna(ts):
        return None
    return pd.Timestamp(ts).date().isoformat()


def _ensure_schema(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS completion_records (
            work_email_norm TEXT NOT NULL,
            training_name TEXT NOT NULL,
            global_employee_id TEXT,
            employee_id TEXT,
            employee_email TEXT,
            transcript_status TEXT NOT NULL,
            transcript_completion_date TEXT,
            zone TEXT,
            date_source TEXT NOT NULL,
            snapshot_date TEXT NOT NULL,
            input_month TEXT,
            input_week TEXT,
            updated_at TEXT NOT NULL,
            PRIMARY KEY (work_email_norm, training_name)
        )
        """
    )
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_completion_snapshot ON completion_records(snapshot_date)"
    )
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_completion_month_week ON completion_records(input_month, input_week)"
    )
    conn.commit()
    _migrate_legacy_backfill(conn)


def _migrate_legacy_backfill(conn: sqlite3.Connection) -> None:
    cur = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name='completion_backfill'"
    )
    if not cur.fetchone():
        return
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    try:
        conn.execute(
            """
            INSERT OR REPLACE INTO completion_records (
                work_email_norm, training_name, global_employee_id, employee_id, employee_email,
                transcript_status, transcript_completion_date, zone,
                date_source, snapshot_date, input_month, input_week, updated_at
            )
            SELECT
                work_email_norm,
                training_name,
                NULL,
                NULL,
                NULL,
                'Completed',
                assigned_completion_date,
                NULL,
                'inferred_legacy',
                assigned_completion_date,
                NULL,
                NULL,
                ?
            FROM completion_backfill
            WHERE NOT EXISTS (
                SELECT 1 FROM completion_records cr
                WHERE cr.work_email_norm = completion_backfill.work_email_norm
                  AND cr.training_name = completion_backfill.training_name
            )
            """,
            (now,),
        )
        conn.commit()
    except sqlite3.OperationalError:
        conn.rollback()


def resolve_completion_date(
    db_path: Optional[Path],
    email: object,
    training_name: str,
    status_mapped: str,
    raw_completion: object,
    snapshot_date_iso: Optional[str] = None,
) -> tuple[str, str]:
    """
    Returns (completion_iso_or_empty, date_source_tag).
    date_source: lms | inferred_stored | inferred_first | inferred_no_email | not_completed
    Inferred dates are spread across Mon–Sun of the snapshot week (see varied_completion_date).
    """
    if (status_mapped or "").strip() != "Completed":
        return "", "not_completed"

    parsed = _parse_date(raw_completion)
    if parsed:
        return parsed, "lms"

    snap = snapshot_date_iso or date.today().isoformat()

    if not db_path:
        return varied_completion_date(snap, email, training_name), "inferred_first"

    em = _norm_email(email)
    tn = (training_name or "").strip()
    if not em or not tn:
        return varied_completion_date(snap, "", training_name), "inferred_no_email"

    conn = sqlite3.connect(str(db_path))
    try:
        _ensure_schema(conn)
        row = conn.execute(
            """
            SELECT transcript_completion_date FROM completion_records
            WHERE work_email_norm=? AND training_name=?
            """,
            (em, tn),
        ).fetchone()
        if row and row[0]:
            return str(row[0]).strip(), "inferred_stored"
        return varied_completion_date(snap, email, training_name), "inferred_first"
    finally:
        conn.close()


def persist_completion_records(
    db_path: Path,
    proc: pd.DataFrame,
    snapshot_date: str,
    month_folder: str,
    week_folder: str,
    sources: list[str],
) -> None:
    """
    Upserts one row per learner/training from processed output. sources[i] aligns with proc row order.
    Stores Completed and Not Completed rows (completion date empty when not completed).
    """
    if proc.empty or len(sources) != len(proc):
        return

    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    conn = sqlite3.connect(str(db_path))
    try:
        _ensure_schema(conn)
        for (_, row), src in zip(proc.iterrows(), sources):
            em = _norm_email(row.get("Work Email Address"))
            tn = str(row.get("Training Name") or "").strip()
            if not em or not tn:
                continue  # primary key requires email + training
            geid = str(row.get("Global Employee ID") or "").strip()
            eid = str(row.get("Employee ID") or "").strip()
            email_disp = str(row.get("Work Email Address") or "").strip()
            status = str(row.get("Transcript Status") or "").strip()
            comp = str(row.get("Transcript Completion Date") or "").strip()
            zone = str(row.get("Zone") or "").strip()

            conn.execute(
                """
                INSERT INTO completion_records (
                    work_email_norm, training_name, global_employee_id, employee_id, employee_email,
                    transcript_status, transcript_completion_date, zone,
                    date_source, snapshot_date, input_month, input_week, updated_at
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
                ON CONFLICT(work_email_norm, training_name) DO UPDATE SET
                    global_employee_id=excluded.global_employee_id,
                    employee_id=excluded.employee_id,
                    employee_email=excluded.employee_email,
                    transcript_status=excluded.transcript_status,
                    transcript_completion_date=excluded.transcript_completion_date,
                    zone=excluded.zone,
                    date_source=excluded.date_source,
                    snapshot_date=excluded.snapshot_date,
                    input_month=excluded.input_month,
                    input_week=excluded.input_week,
                    updated_at=excluded.updated_at
                """,
                (
                    em,
                    tn,
                    geid or None,
                    eid or None,
                    email_disp or None,
                    status,
                    comp or None,
                    zone or None,
                    src,
                    snapshot_date,
                    month_folder or None,
                    week_folder or None,
                    now,
                ),
            )
        conn.commit()
    finally:
        conn.close()
