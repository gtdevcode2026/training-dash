"""Load LMS tool exports and userbase workbooks for dashboard calculations."""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

from training_processor import (
    REQUIRED_TRAINING_COLS,
    _find_header_row_index,
    _strip_columns,
    _training_columns_present,
    read_training_table,
)


def load_tool_excel(path: Path) -> pd.DataFrame:
    """LMS export with header detection (Employee ID, Transcript Status, Work Email)."""
    return read_training_table(path.resolve())


def _row_looks_like_userbase_header(values) -> bool:
    has_email = False
    has_id = False
    for v in values:
        if pd.isna(v):
            continue
        t = str(v).strip().lower()
        if not t:
            continue
        if "email" in t:
            has_email = True
        if re.search(r"employee\s*id|emp\s*id|local\s*employee|global\s*employee", t):
            has_id = True
    return has_email and has_id


def load_userbase_excel(path: Path, scan_rows: int = 35) -> pd.DataFrame:
    """Phishing tracking / BSC userbase — header usually row 1, scanned if needed."""
    path = path.resolve()
    raw = pd.read_excel(path, header=None, nrows=scan_rows, engine="openpyxl")
    header_row = 0
    for i in range(len(raw)):
        if _row_looks_like_userbase_header(raw.iloc[i].values):
            header_row = i
            break
    df = _strip_columns(pd.read_excel(path, header=header_row, engine="openpyxl"))
    status_col = None
    for name in ("Employee Status", "employee status"):
        for c in df.columns:
            if str(c).strip().lower() == name.lower():
                status_col = c
                break
    if status_col:
        mask = df[status_col].astype(str).str.strip().str.lower() != "terminated"
        df = df.loc[mask].copy()
    return df


def find_header_row_index(path: Path) -> int:
    preview = _strip_columns(pd.read_excel(path, header=0, nrows=0, engine="openpyxl"))
    if _training_columns_present(preview):
        return 0
    return _find_header_row_index(path)
