"""SQLite run history for training calculations (separate from completion_records)."""

from __future__ import annotations

import json
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional

DEFAULT_DB_PATH = Path(__file__).resolve().parents[2] / "data" / "training_status.db"


def _utc_now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def ensure_schema(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS calculation_runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            training_id TEXT NOT NULL,
            run_date TEXT NOT NULL,
            assign_date TEXT,
            base_file TEXT,
            tool_file TEXT,
            output_file TEXT,
            total_rows INTEGER,
            completed INTEGER,
            pending INTEGER,
            not_found INTEGER,
            metrics_json TEXT,
            notes TEXT,
            created_at TEXT NOT NULL
        )
        """
    )
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_runs_training ON calculation_runs(training_id, created_at DESC)"
    )
    conn.commit()


def get_connection(db_path: Optional[Path] = None) -> sqlite3.Connection:
    path = db_path or DEFAULT_DB_PATH
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path))
    conn.row_factory = sqlite3.Row
    ensure_schema(conn)
    return conn


def record_run(
    training_id: str,
    *,
    run_date: str,
    assign_date: str = "",
    base_file: str = "",
    tool_file: str = "",
    output_file: str = "",
    total_rows: int = 0,
    completed: int = 0,
    pending: int = 0,
    not_found: int = 0,
    metrics: Optional[dict[str, Any]] = None,
    notes: str = "",
    db_path: Optional[Path] = None,
) -> int:
    conn = get_connection(db_path)
    try:
        cur = conn.execute(
            """
            INSERT INTO calculation_runs (
                training_id, run_date, assign_date, base_file, tool_file, output_file,
                total_rows, completed, pending, not_found, metrics_json, notes, created_at
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,
            (
                training_id,
                run_date,
                assign_date or None,
                base_file or None,
                tool_file or None,
                output_file or None,
                total_rows,
                completed,
                pending,
                not_found,
                json.dumps(metrics or {}),
                notes or None,
                _utc_now(),
            ),
        )
        conn.commit()
        return int(cur.lastrowid)
    finally:
        conn.close()


def list_runs(
    training_id: Optional[str] = None,
    limit: int = 20,
    db_path: Optional[Path] = None,
) -> list[dict[str, Any]]:
    conn = get_connection(db_path)
    try:
        if training_id:
            rows = conn.execute(
                """
                SELECT * FROM calculation_runs
                WHERE training_id = ?
                ORDER BY id DESC LIMIT ?
                """,
                (training_id, limit),
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM calculation_runs ORDER BY id DESC LIMIT ?",
                (limit,),
            ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()
