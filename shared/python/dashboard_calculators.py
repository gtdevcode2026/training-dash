"""
Dashboard calculation logic (mirrors shared/javascript/app.js).

Each function returns (output_dataframe, metrics_dict).
Userbase trainings write all original columns plus appended fields.
"""

from __future__ import annotations

import re
from datetime import date, datetime
from typing import Any, Callable, Optional

import pandas as pd

from excel_io import load_tool_excel, load_userbase_excel
from pathlib import Path
from training_processor import _find_column, _norm_id, _norm_key, _strip_columns, map_macro_zone, map_transcript_status

BSC_APPEND = [
    "start date_extracted",
    "Training completion date",
    "training completion status",
]
TRACKING_APPEND = [
    "Training Start Date",
    "Training completion date",
    "training completion status",
]

OUTPUT_COLS = [
    "Employee ID",
    "Work Email Address",
    "Zone",
    "Transcript Status",
    "Training Start Date",
    "Transcript Completed Date",
]

NOT_FOUND_TEXT = "Learner does not have access to the content"

PHISHING_APPEND = [
    "Transcript Status",
    "Training Start Date",
    "Transcript Completed Date",
]


def _norm(s: object) -> str:
    return _norm_key(s).lower()


def _parse_calendar_year(raw) -> Optional[int]:
    """Calendar year from LMS date cell (start/complete/hire). Avoids 2024–2026 as Excel serial → 1905."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None
    if isinstance(raw, (datetime, date)):
        return raw.year
    if isinstance(raw, pd.Timestamp):
        if pd.isna(raw):
            return None
        return int(raw.year)
    if isinstance(raw, (int, float)):
        n = float(raw)
        if n == int(n) and 1990 <= int(n) <= 2036:
            return int(n)
        if 25000 < n < 60000:
            ts = pd.to_datetime(n, unit="D", origin="1899-12-30", errors="coerce")
            if not pd.isna(ts) and 1990 <= int(ts.year) <= 2036:
                return int(ts.year)
        return None
    s = str(raw).strip()
    if not s:
        return None
    if re.fullmatch(r"20\d{2}", s):
        return int(s)
    ts = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.isna(ts):
        return None
    if int(ts.year) < 1990:
        return None
    return int(ts.year)


def _fmt_date_iso(val: object) -> str:
    """Date → YYYY-MM-DD (ISO).
    Mirrors standalone: pd.to_datetime(..., errors='coerce') → strftime or '' on NaT.
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, pd.Timestamp):
        if pd.isna(val):
            return ""
        return val.strftime("%Y-%m-%d")
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, date):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, (int, float)):
        n = float(val)
        if n == int(n) and 1990 <= int(n) <= 2036:
            return f"{int(n)}-01-01"
        if 25000 < n < 60000:
            ts = pd.to_datetime(n, unit="D", origin="1899-12-30", errors="coerce")
            if not pd.isna(ts):
                return ts.strftime("%Y-%m-%d")
        return ""
    s = str(val).strip()
    if not s:
        return ""
    ts = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.isna(ts):
        return ""
    return ts.strftime("%Y-%m-%d")


def _fmt_date(val: object) -> str:
    """Training Start Date, Completed Date, assignment date → dd/mm/yyyy."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, pd.Timestamp):
        if pd.isna(val):
            return ""
        return val.strftime("%d/%m/%Y")
    if isinstance(val, datetime):
        return val.strftime("%d/%m/%Y")
    if isinstance(val, date):
        return val.strftime("%d/%m/%Y")
    if isinstance(val, (int, float)):
        n = float(val)
        if n == int(n) and 1990 <= int(n) <= 2036:
            return f"01/01/{int(n)}"
        if 25000 < n < 60000:
            ts = pd.to_datetime(n, unit="D", origin="1899-12-30", errors="coerce")
            if not pd.isna(ts) and 1990 <= int(ts.year) <= 2036:
                return ts.strftime("%d/%m/%Y")
    s = str(val).strip()
    if not s:
        return ""
    if re.fullmatch(r"20\d{2}", s):
        return f"01/01/{s}"
    ts = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.isna(ts):
        return s
    if int(ts.year) < 1990:
        return s
    return ts.strftime("%d/%m/%Y")


def _new_joiner_year_tab(calendar_year: int) -> Optional[str]:
    if calendar_year == 2024:
        return "2024"
    if calendar_year == 2025:
        return "2025"
    if calendar_year == 2026:
        return "2026"
    return None


def _hire_date_cell_to_string(raw) -> str:
    """Minimal conversion — keep LMS text; only format obvious Excel serials / years."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return ""
    if isinstance(raw, pd.Timestamp):
        if pd.isna(raw):
            return ""
        return raw.strftime("%d/%m/%Y")
    if isinstance(raw, (datetime, date)):
        d = raw if isinstance(raw, datetime) else datetime.combine(raw, datetime.min.time())
        return d.strftime("%d/%m/%Y")
    if isinstance(raw, (int, float)):
        n = float(raw)
        if n == int(n) and 1990 <= int(n) <= 2036:
            return f"01/01/{int(n)}"
        if 25000 < n < 60000:
            ts = pd.to_datetime(n, unit="D", origin="1899-12-30", errors="coerce")
            if not pd.isna(ts) and 1990 <= int(ts.year) <= 2036:
                return ts.strftime("%d/%m/%Y")
        return str(raw).strip()
    s = str(raw).strip()
    if re.fullmatch(r"\d{5}(\.\d+)?", s):
        n = float(s)
        if 25000 < n < 60000:
            ts = pd.to_datetime(n, unit="D", origin="1899-12-30", errors="coerce")
            if not pd.isna(ts) and 1990 <= int(ts.year) <= 2036:
                return ts.strftime("%d/%m/%Y")
    return s


def _extract_hire_year_from_string(s: str) -> Optional[int]:
    """Find 2024, 2025, or 2026 in hire cell text (no strict date parse)."""
    t = str(s or "").strip()
    if not t:
        return None
    if t in ("2024", "2025", "2026"):
        return int(t)
    m = re.search(r"/(2024|2025|2026)\s*$", t)
    if m:
        return int(m.group(1))
    m = re.search(r"(2024|2025|2026)-\d{1,2}-\d{1,2}", t)
    if m:
        return int(m.group(1))
    m = re.search(r"(2024|2025|2026)\s*$", t)
    if m and re.search(r"[-/]", t):
        return int(m.group(1))
    hits = re.findall(r"2024|2025|2026", t)
    if hits:
        return int(hits[-1])
    return None


def _normalize_original_hire_date(raw) -> Optional[tuple[str, str]]:
    """Return (display text, year_tab) for New Joiner or None if year not 2024–2026."""
    display = _hire_date_cell_to_string(raw)
    if not display:
        return None
    year = _extract_hire_year_from_string(display)
    tab = _new_joiner_year_tab(year) if year else None
    if not tab:
        return None
    return display, tab


def _year_bucket_from_date(sd: str) -> str:
    """Match yearBucketFromStartDate in app.js (Training Start Date → 2024 / 2025 / 2026 tab)."""
    current_year = date.today().year
    cal = _parse_calendar_year(sd)
    year = cal if cal is not None else current_year
    if year <= 2024:
        return "2024"
    if year == 2025:
        return "2025"
    return "2026"


def _derive_status(raw_status: str, completed_date: str) -> str:
    if completed_date and str(completed_date).strip():
        return "Completed"
    raw = str(raw_status or "").strip()
    if re.search(r"in progress", raw, re.I):
        return "In Progress"
    if raw:
        return raw
    return "Not Started"


def _find_col(df: pd.DataFrame, exact: list[str], *patterns: str) -> Optional[str]:
    for name in exact:
        c = _find_column(df, name)
        if c:
            return c
    for pat in patterns:
        rx = re.compile(pat, re.I)
        for c in df.columns:
            if rx.search(str(c)):
                return c
    return None


def _tool_cols(df: pd.DataFrame) -> dict[str, Optional[str]]:
    return {
        "id": _find_col(df, ["Employee ID"], r"^employee\s*id$", r"employee.?id|emp.?id"),
        "email": _find_col(
            df, ["Work Email Address"], r"^work\s*email", r"work.?email|email.?address"
        ),
        "status": _find_col(df, ["Transcript Status"], r"transcript.?status"),
        "start": _find_col(df, ["Training Start Date"], r"training.?start"),
        "complete": _find_col(
            df,
            ["Transcript Completed Date", "Transcript Completion Date", "Completion Date"],
            r"transcript.?completed|completion.?date",
        ),
        "zone": _find_col(
            df, ["Macro Entity Level 2 (Zone)", "Zone"], r"macro.?entity.?level.?2", r"^zone$"
        ),
        "emp_status": _find_col(df, ["Employee Status"], r"employee\s*status"),
        "bu3": _find_col(
            df,
            [
                "Macro entity level three BU Description",
                "Macro Entity Level 3 (BU Description)",
                "Macro Entity Level 3 BU Description",
            ],
            r"macro.?entity.?level.?three.*bu",
            r"macro.?entity.?level.?3.*bu",
            r"macro.?entity.?level.?3",
        ),
        "hire": _find_col(
            df,
            [
                "Original Hire Date",
                "New J orginal hire date",
                "New J original hire date",
                "New Joiner Original Hire Date",
            ],
            r"^original\s*hire\s*date$",
            r"new\s*j\s*orginal\s*hire",
            r"new\s*j\s*original\s*hire",
            r"new\s*joiner.*original\s*hire",
        ),
    }


def _last_real_header_index(headers: list[str]) -> int:
    last = -1
    for i, h in enumerate(headers):
        if str(h).strip() and not str(h).startswith("__col_"):
            last = i
    return last


def _append_columns(headers: list[str], append: list[str]) -> list[str]:
    out = headers[: _last_real_header_index(headers) + 1]
    seen = {str(h).strip().lower() for h in out}
    for col in append:
        if col.lower() not in seen:
            out.append(col)
            seen.add(col.lower())
    return out


def _build_lookup(df: pd.DataFrame, email_col: Optional[str], id_col: Optional[str]) -> dict[str, dict]:
    lookup: dict[str, dict] = {}
    for _, row in df.iterrows():
        d = row.to_dict()
        if email_col:
            e = _norm(row.get(email_col))
            if e:
                lookup[e] = d
        if id_col:
            i = _norm_id(row.get(id_col))
            if i:
                lookup[i] = d
    return lookup


def _build_lookup_list(
    df: pd.DataFrame, email_col: Optional[str], id_col: Optional[str]
) -> dict[str, list[dict]]:
    """Like _build_lookup but keeps ALL tool rows per key (email + id) for multi-row resolution."""
    lookup: dict[str, list[dict]] = {}
    for _, row in df.iterrows():
        d = row.to_dict()
        if email_col:
            e = _norm(row.get(email_col))
            if e:
                lookup.setdefault(e, []).append(d)
        if id_col:
            i = _norm_id(row.get(id_col))
            if i:
                lookup.setdefault(i, []).append(d)
    return lookup


def _gather_tool_rows_multi(lookup_list: dict[str, list[dict]], values: dict[str, str]) -> list[dict]:
    """Collect all tool rows matching a userbase user across email + global/local/emp ids (deduped)."""
    out: list[dict] = []
    seen: set[int] = set()
    keys = [_norm(values.get("email") or "")]
    keys += [_norm_id(values.get(k) or "") for k in ("global_id", "local_id", "emp_id")]
    for key in keys:
        if not key:
            continue
        for d in lookup_list.get(key, []):
            if id(d) not in seen:
                seen.add(id(d))
                out.append(d)
    return out


def _pick_latest_tool_row(rows: list[dict], cols: dict) -> Optional[dict]:
    """Latest tool row by Training Start Date (newest wins); None if empty."""
    if not rows:
        return None
    return max(
        rows,
        key=lambda r: _parse_date_for_compare(_fmt_date(r.get(cols["start"]))) or datetime.min,
    )


def _row_is_completed(row: dict, cols: dict) -> bool:
    """A tool row whose derived status is Completed (has a completion date or 'completed' status)."""
    completed_date = _fmt_date(row.get(cols["complete"]))
    raw_status = str(row.get(cols["status"]) or "").strip()
    return _derive_status(raw_status, completed_date) == "Completed"


def _pick_completed_tool_row(rows: list[dict], cols: dict) -> Optional[dict]:
    """Completed-wins: if the user completed on ANY transcript row, return the latest such row
    (by completion date, then start date); None if no row is completed. This keeps a user who
    finished the training Completed even when a newer re-assignment row is Not Started/Terminated."""
    completed_rows = [r for r in rows if _row_is_completed(r, cols)]
    if not completed_rows:
        return None
    return max(
        completed_rows,
        key=lambda r: (
            _parse_date_for_compare(_fmt_date(r.get(cols["complete"]))) or datetime.min,
            _parse_date_for_compare(_fmt_date(r.get(cols["start"]))) or datetime.min,
        ),
    )


def _find_match(lookup: dict[str, dict], email: str, emp_id: str) -> Optional[dict]:
    if email:
        r = lookup.get(_norm(email))
        if r:
            return r
    if emp_id:
        r = lookup.get(_norm_id(emp_id))
        if r:
            return r
    return None


def _tool_rows_for_user(
    tool_df: pd.DataFrame, cols: dict, email: str, emp_id: str
) -> list[dict]:
    out: list[dict] = []
    for _, row in tool_df.iterrows():
        r_email = _norm(row.get(cols["email"])) if cols["email"] else ""
        r_id = _norm_id(row.get(cols["id"])) if cols["id"] else ""
        match = (email and r_email == _norm(email)) or (emp_id and r_id == _norm_id(emp_id))
        if not match:
            continue
        out.append(row.to_dict())
    return out


def _parse_date_for_compare(s: str) -> Optional[datetime]:
    if not s:
        return None
    parts = str(s).split("/")
    if len(parts) == 3:
        try:
            return datetime(int(parts[2]), int(parts[1]), int(parts[0]))
        except ValueError:
            pass
    ts = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.isna(ts):
        return None
    return ts.to_pydatetime()


def _is_terminated_tool_row(row: dict, emp_status_col: Optional[str]) -> bool:
    """Tool export marks this employee Terminated (mirrors app.js isTerminatedRow)."""
    if not emp_status_col:
        return False
    s = _norm(row.get(emp_status_col))
    return bool(re.search(r"\bterminated\b", s)) or s == "terminated"


def _resolve_bsc_training(tool_rows: list[dict], cols: dict, assign_date: str) -> dict[str, str]:
    start_extracted = _fmt_date(assign_date) or assign_date or ""
    status = "Not Started"
    completed = ""

    if not tool_rows:
        return {
            "status": "Not Found",
            "start_date_extracted": start_extracted,
            "completed_date": completed,
        }

    # Multiple matches: use the latest training-start row only.
    row = max(
        tool_rows,
        key=lambda r: _parse_date_for_compare(_fmt_date(r.get(cols["start"]))) or datetime.min,
    )

    # Terminated in the tool export overrides everything: mark the userbase row Terminated.
    if _is_terminated_tool_row(row, cols.get("emp_status")):
        return {
            "status": "Terminated",
            "start_date_extracted": start_extracted,
            "completed_date": "Terminated",
        }

    raw_status = str(row.get(cols["status"]) or "").strip()
    completed = _fmt_date(row.get(cols["complete"]))
    status = _derive_status(raw_status, completed)

    return {
        "status": status,
        "start_date_extracted": start_extracted,
        "completed_date": completed,
    }


def _make_output_rows(
    tool_df: pd.DataFrame,
    cols: dict,
    assign_date: str,
    *,
    zone_col: Optional[str] = None,
    map_zone: bool = True,
) -> pd.DataFrame:
    rows = []
    for _, tr in tool_df.iterrows():
        emp_id = tr.get(cols["id"])
        email = tr.get(cols["email"])
        raw_status = str(tr.get(cols["status"]) or "").strip()
        completed_date = _fmt_date(tr.get(cols["complete"]))
        start_date = _fmt_date(tr.get(cols["start"]))
        if _norm(raw_status) != "completed" and assign_date and not start_date:
            start_date = _fmt_date(assign_date)
        if zone_col:
            zone_raw = tr.get(zone_col)
        else:
            zone_raw = tr.get(cols["zone"]) if cols["zone"] else ""
        zone = (
            map_macro_zone(zone_raw) or str(zone_raw or "").strip()
            if map_zone
            else str(zone_raw or "").strip()
        )
        status = _derive_status(raw_status, completed_date)
        rows.append(
            {
                "Employee ID": _norm_id(emp_id) or str(emp_id or "").strip(),
                "Work Email Address": str(email or "").strip(),
                "Zone": zone,
                "Transcript Status": status,
                "Training Start Date": start_date,
                "Transcript Completed Date": completed_date,
            }
        )
    return pd.DataFrame(rows, columns=OUTPUT_COLS)


def _metrics_from_output(df: pd.DataFrame, status_col: str = "Transcript Status") -> dict[str, Any]:
    total = len(df)
    completed = sum(1 for s in df[status_col].astype(str) if _norm(s) == "completed")
    return {
        "total": total,
        "completed": completed,
        "pending": total - completed,
        "not_found": 0,
    }


def calculate_app_sec(tool_path, assign_date: str = "") -> tuple[pd.DataFrame, dict]:
    tool_df = load_tool_excel(tool_path)
    cols = _tool_cols(tool_df)
    out = _make_output_rows(tool_df, cols, assign_date)
    return out, _metrics_from_output(out)


def calculate_cyber_ot(tool_path, assign_date: str = "") -> tuple[pd.DataFrame, dict]:
    return calculate_app_sec(tool_path, assign_date)


def calculate_growth(tool_path, assign_date: str = "") -> tuple[pd.DataFrame, dict]:
    tool_df = load_tool_excel(tool_path)
    cols = _tool_cols(tool_df)
    if not cols["bu3"]:
        raise ValueError('Missing column "Macro entity level three BU Description"')
    bu = tool_df[cols["bu3"]].astype(str).str.strip().str.upper().str.replace(r"\s+", " ", regex=True)
    filtered = tool_df.loc[bu == "GLOBAL GROWTH"].copy()
    out = _make_output_rows(filtered, cols, assign_date, zone_col=cols["bu3"], map_zone=False)
    m = _metrics_from_output(out)
    m["excluded"] = len(tool_df) - len(filtered)
    m["total_in_file"] = len(tool_df)
    return out, m


def _is_global_core_zone(mapped_zone: str, raw_zone: object) -> bool:
    z = _norm(mapped_zone)
    r = _norm(raw_zone)
    return z in ("global core", "global") or r in ("global core", "global")


def _resolve_new_joiner_zone(tr: pd.Series, cols: dict, bu_col: Optional[str]) -> str:
    raw = tr.get(cols["zone"]) if cols["zone"] else ""
    zone = map_macro_zone(raw) or str(raw or "").strip()
    if bu_col and _is_global_core_zone(zone, raw):
        bu = re.sub(r"\s+", " ", str(tr.get(bu_col) or "").strip().upper())
        if bu == "GLOBAL GROWTH":
            return "Growth"
    return zone


def calculate_new_joiner(tool_path, assign_date: str = "") -> tuple[pd.DataFrame, dict]:
    tool_df = load_tool_excel(tool_path)
    status_col = _find_col(tool_df, ["Employee Status"], r"employee\s*status")
    if status_col:
        tool_df = tool_df[
            tool_df[status_col].astype(str).str.strip().str.lower() != "terminated"
        ].copy()
    cols = _tool_cols(tool_df)
    if not cols["hire"]:
        raise ValueError('Missing column "Original Hire Date" in tool export')
    bu_col = cols["bu3"]
    rows = []
    for _, tr in tool_df.iterrows():
        emp_id = tr.get(cols["id"])
        email = tr.get(cols["email"])
        hire_norm = _normalize_original_hire_date(tr.get(cols["hire"]))
        if not hire_norm:
            continue
        hire, _year_tab = hire_norm
        raw_status = str(tr.get(cols["status"]) or "").strip()
        completed_date = _fmt_date(tr.get(cols["complete"]))
        zone = _resolve_new_joiner_zone(tr, cols, bu_col)
        status = _derive_status(raw_status, completed_date)
        rows.append(
            {
                "Employee ID": _norm_id(emp_id) or str(emp_id or "").strip(),
                "Work Email Address": str(email or "").strip(),
                "Zone": zone,
                "Transcript Status": status,
                "Training Start Date": hire,
                "Transcript Completed Date": completed_date,
            }
        )
    out = pd.DataFrame(rows, columns=OUTPUT_COLS)
    return out, _metrics_from_output(out)


def _resolve_base_id_email(base_df: pd.DataFrame) -> tuple[Optional[str], Optional[str]]:
    id_col = _find_col(
        base_df, ["Emp ID", "Local Employee ID", "Employee ID"], r"emp.?id", r"local.?id"
    )
    email_col = _find_col(base_df, ["Employee Email", "Email - Primary Work"], r"email")
    return id_col, email_col


def _resolve_userbase_match_cols(base_df: pd.DataFrame) -> dict[str, Optional[str]]:
    return {
        "email": _find_col(
            base_df,
            ["Employee Email", "Email - Primary Work"],
            r"employee\s*email",
            r"^email$",
            r"email",
        ),
        "global_id": _find_col(base_df, ["Global Employee ID"], r"global\s*employee\s*id"),
        "local_id": _find_col(
            base_df, ["Local Employee ID"], r"local\s*employee\s*id", r"local.?id"
        ),
        "emp_id": _find_col(
            base_df, ["Emp ID", "Employee ID", "Local Employee ID"], r"emp.?id", r"^employee\s*id$"
        ),
    }


def _userbase_match_values(row_dict: dict, cols: dict[str, Optional[str]]) -> dict[str, str]:
    return {
        "email": str(row_dict.get(cols["email"]) or "").strip() if cols["email"] else "",
        "global_id": str(row_dict.get(cols["global_id"]) or "").strip() if cols["global_id"] else "",
        "local_id": str(row_dict.get(cols["local_id"]) or "").strip() if cols["local_id"] else "",
        "emp_id": str(row_dict.get(cols["emp_id"]) or "").strip() if cols["emp_id"] else "",
    }


def _userbase_dedupe_key(values: dict[str, str]) -> str:
    return (
        _norm(values.get("email") or "")
        or _norm_id(values.get("global_id") or "")
        or _norm_id(values.get("local_id") or "")
        or _norm_id(values.get("emp_id") or "")
    )


def _find_match_multi(lookup: dict[str, dict], values: dict[str, str]) -> Optional[dict]:
    email = values.get("email") or ""
    if email:
        hit = lookup.get(_norm(email))
        if hit:
            return hit
    for key in (values.get("global_id"), values.get("local_id"), values.get("emp_id")):
        if not key:
            continue
        hit = lookup.get(_norm_id(key))
        if hit:
            return hit
    return None


def _enrich_userbase_tracking(
    base_df: pd.DataFrame,
    tool_df: pd.DataFrame,
    assign_date: str,
    in_scope: Callable[[dict], bool],
    *,
    start_from_assign_date: bool = False,
    multi_row_latest: bool = False,
    keep_terminated: bool = False,
) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    match_cols = _resolve_userbase_match_cols(base_df)
    cols = _tool_cols(tool_df)
    lookup = _build_lookup(tool_df, cols["email"], cols["id"])
    lookup_list = _build_lookup_list(tool_df, cols["email"], cols["id"]) if multi_row_latest else None
    zone_col = _find_col(base_df, ["Zone", "Macro Entity Level 2 (Zone)"], r"^zone$", r"macro.?entity.?level.?2")

    headers = _append_columns(list(base_df.columns.astype(str)), TRACKING_APPEND)
    enriched_rows: list[dict] = []
    stats_rows: list[dict] = []
    completed = not_completed = not_found = terminated = matched = 0

    for _, base_row in base_df.iterrows():
        row_dict = {str(k): base_row[k] for k in base_df.columns}
        enriched = {h: "" for h in headers}
        enriched.update(row_dict)

        if not in_scope(row_dict):
            enriched_rows.append(enriched)
            continue

        match_values = _userbase_match_values(row_dict, match_cols)
        match_key = _userbase_dedupe_key(match_values)

        if match_key:
            force_completed = False
            if multi_row_latest:
                gathered = _gather_tool_rows_multi(lookup_list, match_values)
                # Completed wins: if the user completed on ANY matched row, keep them Completed
                # even when a newer re-assignment row is Not Started/Terminated.
                tool_row = _pick_completed_tool_row(gathered, cols)
                if tool_row is not None:
                    force_completed = True
                else:
                    tool_row = _pick_latest_tool_row(gathered, cols)
            else:
                tool_row = _find_match_multi(lookup, match_values)
            if tool_row:
                raw_status = str(tool_row.get(cols["status"]) or "").strip()
                if start_from_assign_date:
                    start_date = _fmt_date(assign_date) or ""
                else:
                    start_date = _fmt_date(tool_row.get(cols["start"])) or assign_date or ""
                zone = row_dict.get(zone_col) if zone_col else ""
                if not zone and cols["zone"]:
                    zone = tool_row.get(cols["zone"])
                if force_completed:
                    completed_date = _fmt_date(tool_row.get(cols["complete"]))
                    status = "Completed"
                elif keep_terminated and _is_terminated_tool_row(tool_row, cols.get("emp_status")):
                    status = "Terminated"
                    completed_date = "Terminated"
                else:
                    completed_date = _fmt_date(tool_row.get(cols["complete"]))
                    status = _derive_status(raw_status, completed_date)
            else:
                status = "Not Found"
                start_date = completed_date = ""

            enriched["Training Start Date"] = start_date
            enriched["Training completion date"] = completed_date
            enriched["training completion status"] = status

        enriched_rows.append(enriched)
        stats_rows.append(enriched)

        status_val = _norm(str(enriched.get("training completion status") or ""))
        if status_val == "completed":
            completed += 1
            matched += 1
        elif status_val == "not found":
            not_found += 1
        elif status_val == "terminated":
            terminated += 1
            matched += 1
        else:
            not_completed += 1
            matched += 1

    out = pd.DataFrame(enriched_rows, columns=headers)
    stats_df = pd.DataFrame(stats_rows, columns=headers) if stats_rows else pd.DataFrame(columns=headers)
    metrics = {
        "total": len(stats_rows),
        "completed": completed,
        "pending": not_completed + not_found,
        "not_completed": not_completed,
        "not_found": not_found,
        "terminated": terminated,
        "matched": matched,
        "userbase_rows": len(base_df),
    }
    return out, stats_df, metrics


def _in_phishing_normal_scope(row: dict) -> bool:
    # User uploads already-filtered data for Phishing Normal.
    return True


def _in_band4_scope(row: dict) -> bool:
    # User uploads already-filtered data for Band 4+.
    return True


def _enrich_phishing_band4_tracking(
    base_df: pd.DataFrame,
    tool_df: pd.DataFrame,
) -> tuple[pd.DataFrame, dict]:
    """
    Matching logic from TrainingPhisheduser_Automation.py:
    - Sort tool rows by Training Start Date desc; keep latest record per email/ID.
    - Match priority: email → Global Employee ID → Local Employee ID → Emp ID.
    - Not found  → NOT_FOUND_TEXT in all three output columns.
    - Terminated → "Terminated" in all three output columns.
    - Otherwise  → raw Transcript Status, formatted start date, formatted completed date.
    """
    match_cols = _resolve_userbase_match_cols(base_df)
    cols = _tool_cols(tool_df)

    # Sort descending so the first row encountered per key is the latest
    tool_sorted = tool_df.copy()
    if cols["start"]:
        tool_sorted["_sd"] = pd.to_datetime(tool_sorted[cols["start"]], errors="coerce")
        tool_sorted = tool_sorted.sort_values("_sd", ascending=False)

    # Build email and ID lookups keeping only the latest row per key
    email_lookup: dict[str, dict] = {}
    id_lookup: dict[str, dict] = {}
    for _, row in tool_sorted.iterrows():
        d = row.to_dict()
        if cols["email"]:
            e = _norm(row.get(cols["email"]))
            if e and e not in email_lookup:
                email_lookup[e] = d
        if cols["id"]:
            i = _norm_id(row.get(cols["id"]))
            if i and i not in id_lookup:
                id_lookup[i] = d

    headers = _append_columns(list(base_df.columns.astype(str)), PHISHING_APPEND)
    enriched_rows: list[dict] = []
    completed = not_completed = not_found = terminated = matched = 0

    for _, base_row in base_df.iterrows():
        row_dict = {str(k): base_row[k] for k in base_df.columns}
        enriched = {h: "" for h in headers}
        enriched.update(row_dict)

        mv = _userbase_match_values(row_dict, match_cols)

        # Match: email first, then Global/Local/Emp ID
        rec: Optional[dict] = None
        e = _norm(mv.get("email") or "")
        if e:
            rec = email_lookup.get(e)
        if rec is None:
            for key in (mv.get("global_id"), mv.get("local_id")):
                if not key:
                    continue
                rec = id_lookup.get(_norm_id(key))
                if rec:
                    break

        if rec is None:
            enriched["Transcript Status"] = NOT_FOUND_TEXT
            enriched["Training Start Date"] = NOT_FOUND_TEXT
            enriched["Transcript Completed Date"] = NOT_FOUND_TEXT
            not_found += 1
        elif cols.get("emp_status") and _norm(rec.get(cols["emp_status"])) == "terminated":
            enriched["Transcript Status"] = "Terminated"
            enriched["Training Start Date"] = "Terminated"
            enriched["Transcript Completed Date"] = "Terminated"
            terminated += 1
            matched += 1
        else:
            raw_status = str(rec.get(cols["status"]) or "").strip() if cols["status"] else ""
            start_date = _fmt_date_iso(rec.get(cols["start"])) if cols["start"] else ""
            comp_date = _fmt_date_iso(rec.get(cols["complete"])) if cols["complete"] else ""
            enriched["Transcript Status"] = raw_status
            enriched["Training Start Date"] = start_date
            enriched["Transcript Completed Date"] = comp_date
            if _norm(raw_status) == "completed":
                completed += 1
            else:
                not_completed += 1
            matched += 1

        enriched_rows.append(enriched)

    out = pd.DataFrame(enriched_rows, columns=headers)
    metrics = {
        "total": len(enriched_rows),
        "completed": completed,
        "pending": not_completed + not_found,
        "not_completed": not_completed,
        "not_found": not_found,
        "terminated": terminated,
        "matched": matched,
        "userbase_rows": len(base_df),
    }
    return out, metrics


def _load_base_unfiltered(base_path) -> pd.DataFrame:
    """Load userbase exactly as the standalone does — full file, no terminated filtering.
    Terminated status is determined from the tool export, not the base file.
    """
    path = Path(base_path).resolve()
    scan_rows = 35
    raw = pd.read_excel(path, header=None, nrows=scan_rows, engine="openpyxl")

    def _looks_like_header(values) -> bool:
        has_email = has_id = False
        for v in values:
            if pd.isna(v):
                continue
            t = str(v).strip().lower()
            if "email" in t:
                has_email = True
            if re.search(r"employee\s*id|emp\s*id|local\s*employee|global\s*employee", t):
                has_id = True
        return has_email and has_id

    header_row = 0
    for i in range(len(raw)):
        if _looks_like_header(raw.iloc[i].values):
            header_row = i
            break
    return _strip_columns(pd.read_excel(path, header=header_row, engine="openpyxl"))


def calculate_phishing_normal(base_path, tool_path, assign_date: str = "") -> tuple[pd.DataFrame, dict]:
    base_df = _load_base_unfiltered(base_path)
    tool_df = load_tool_excel(tool_path)
    out, metrics = _enrich_phishing_band4_tracking(base_df, tool_df)
    return out, metrics


def calculate_band4(base_path, tool_path, assign_date: str = "") -> tuple[pd.DataFrame, dict]:
    base_df = _load_base_unfiltered(base_path)
    tool_df = load_tool_excel(tool_path)
    out, metrics = _enrich_phishing_band4_tracking(base_df, tool_df)
    return out, metrics


def calculate_bsc(base_path, tool_path, assign_date: str = "") -> tuple[pd.DataFrame, dict]:
    base_df = _load_base_unfiltered(base_path)
    tool_df = load_tool_excel(tool_path)
    out, metrics = _enrich_phishing_band4_tracking(base_df, tool_df)
    return out, metrics


TRAINING_CALCULATORS = {
    "app_sec": ("application_security", calculate_app_sec, False),
    "application_security": ("application_security", calculate_app_sec, False),
    "cyber_ot": ("cyber_ot", calculate_cyber_ot, False),
    "growth": ("growth_group", calculate_growth, False),
    "growth_group": ("growth_group", calculate_growth, False),
    "new_joiner": ("new_joiner", calculate_new_joiner, False),
    "bsc": ("bsc", calculate_bsc, True),
    "phishing_normal": ("phishing_normal", calculate_phishing_normal, True),
    "phishingNormal": ("phishing_normal", calculate_phishing_normal, True),
    "band4": ("band4_senior_management", calculate_band4, True),
    "band4_senior_management": ("band4_senior_management", calculate_band4, True),
}
