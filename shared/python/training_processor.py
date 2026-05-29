"""
Shared logic for BSC / CyberOT / Phished / DirPhished training exports.
Reads source + Datamart from the script directory; writes results under output/ with a dated
filename and copies inputs to processed/<date>/ without modifying originals.
"""

from __future__ import annotations

import os
import re
import shutil
import sys
from datetime import date
from pathlib import Path
from typing import Optional

import pandas as pd
from rich.console import Console
from rich.progress import (
    BarColumn,
    Progress,
    SpinnerColumn,
    TaskProgressColumn,
    TextColumn,
    TimeElapsedColumn,
)

from completion_state import persist_completion_records, resolve_completion_date

# Zone fallback: values in "Macro Entity Level 2 (Zone)" -> normalized zone code
MACRO_ZONE_MAP = {
    "GLOBAL": "Global Core",
    "ZONE MIDDLE AMERICAS": "MAZ",
    "ZONE EUROPE": "EUR",
    "ZONE ASIA PACIFIC": "APAC",
    "ZONE AFRICA": "AFR",
    "ZONE SOUTH AMERICA": "SAZ",
    "ZONE NORTH AMERICA": "NAZ",
}

# Expected Transcript Status labels from the source. Other stray LMS values are handled in
# map_transcript_status before this lookup.
TRANSCRIPT_STATUS_MAP = {
    "completed": "Completed",
    "in progress": "Not Completed",
    "not started": "Not Completed",
}

# Header text must match your LMS / Datamart exports (Cornerstone transcript + Workday-style dump).
REQUIRED_TRAINING_COLS = ("Employee ID", "Transcript Status", "Work Email Address")
OPTIONAL_TRAINING_COLS = (
    "Employee Status",
    "Macro Entity Level 2 (Zone)",
    "Transcript Completion Date",
    "Completion Date",
)
COMPLETION_DATE_ALIASES = ("Transcript Completion Date", "Completion Date")
DATAMART_REQUIRED_COLS = ("Global Employee ID", "Employee Email", "Zone")
ASSIGN_NEW_JOINERS_REQUIRED_COLS = (
    "Global Employee ID",
    "Employee Name",
    "Zone",
    "Employee Email",
    "First Hire Date",
)


def _strip_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    return out


def _find_column(df: pd.DataFrame, name: str) -> Optional[str]:
    """Exact match after strip; else case-insensitive match."""
    cols = list(df.columns)
    for c in cols:
        if str(c).strip() == name:
            return c
    name_lower = name.lower()
    for c in cols:
        if str(c).strip().lower() == name_lower:
            return c
    return None


def _training_columns_present(df: pd.DataFrame) -> bool:
    return all(_find_column(df, n) is not None for n in REQUIRED_TRAINING_COLS)


def read_learning_record_name(path: Path, max_rows: int = 80) -> str:
    """
    LMS exports put the course title on the row labeled 'Learning Record' (column A),
    value in column B (see ``samples actual/`` reference workbooks).

    The **Training Name** column in processed output always comes from this cell — not
    from the file name, so odd or special characters in the filename do not affect it.

    Returns the training name string, or '' if not found.
    """
    raw = pd.read_excel(path, header=None, nrows=max_rows, engine="openpyxl")
    label_ok = frozenset(
        {
            "learning record",
            "learning record:",
        }
    )
    for i in range(len(raw)):
        a = raw.iloc[i, 0]
        if pd.isna(a):
            continue
        lab = str(a).strip().lower().rstrip(":")
        if lab not in label_ok:
            continue
        if raw.shape[1] < 2:
            return ""
        val = raw.iloc[i, 1]
        if pd.isna(val):
            return ""
        return str(val).strip()
    return ""


def _find_header_row_index(path: Path, max_rows: int = 120) -> int:
    """LMS exports often put the real column header row below report titles."""
    raw = pd.read_excel(path, header=None, nrows=max_rows, engine="openpyxl")
    required = {n.lower() for n in REQUIRED_TRAINING_COLS}
    for i in range(len(raw)):
        cells: set[str] = set()
        for x in raw.iloc[i].values:
            if pd.isna(x):
                continue
            t = str(x).strip()
            if t:
                cells.add(t.lower())
        if required.issubset(cells):
            return i
    return 0


def read_training_table(path: Path) -> pd.DataFrame:
    """Load the first sheet that has Employee ID / Transcript / Work Email (after header detection)."""
    preview = _strip_columns(pd.read_excel(path, header=0, nrows=0, engine="openpyxl"))
    if _training_columns_present(preview):
        return _strip_columns(pd.read_excel(path, header=0, engine="openpyxl"))
    hr = _find_header_row_index(path)
    return _strip_columns(pd.read_excel(path, header=hr, engine="openpyxl"))


def load_training_excel(path: Path, progress: Progress, task_id: int) -> pd.DataFrame:
    progress.update(task_id, description="[cyan]Detecting header row...")
    preview = _strip_columns(pd.read_excel(path, header=0, nrows=0, engine="openpyxl"))
    if _training_columns_present(preview):
        progress.advance(task_id)
        progress.update(task_id, description="[cyan]Reading training rows...")
        df = _strip_columns(pd.read_excel(path, header=0, engine="openpyxl"))
        progress.advance(task_id)
        return df
    hr = _find_header_row_index(path)
    progress.advance(task_id)
    progress.update(task_id, description="[cyan]Reading training rows...")
    df = _strip_columns(pd.read_excel(path, header=hr, engine="openpyxl"))
    progress.advance(task_id)
    return df


def _norm_key(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return str(s).strip()


def _norm_id(s) -> str:
    t = _norm_key(s)
    if not t:
        return ""
    if t.endswith(".0") and t.replace(".0", "").isdigit():
        t = t[:-2]
    return t


def map_transcript_status(raw) -> str:
    k = _norm_key(raw).lower()
    if k == "refresher":
        k = "completed"
    return TRANSCRIPT_STATUS_MAP.get(k, _norm_key(raw) or "")


def map_macro_zone(raw) -> str:
    k = _norm_key(raw).upper()
    k = re.sub(r"\s+", " ", k).strip()
    return MACRO_ZONE_MAP.get(k, "")


def _stem_normalized(name: str) -> str:
    return re.sub(r"\s+", "", Path(name).stem).lower()


def _is_datamart_candidate(path: Path) -> bool:
    if not path.suffix.lower() == ".xlsx":
        return False
    if path.name.startswith("~$"):
        return False
    stem = _stem_normalized(path.name)
    return "datamart" in stem


def _norm_alnum(s: str) -> str:
    """Lowercase alphanumerics only — for fuzzy key vs learning-record matching."""
    return re.sub(r"[^a-z0-9]+", "", (s or "").lower())


def resolve_week_training_source(
    week_dir: Path,
    key: str,
    *,
    exclude: Optional[set[Path]] = None,
) -> Optional[Path]:
    """
    Find the LMS training-export workbook for a logical pipeline key (e.g. ``CyberOT``).

    1. Prefer ``{key}.xlsx`` in ``week_dir`` when present (legacy layout).
    2. Otherwise scan ``*.xlsx`` in that folder (excluding Datamart workbooks and Excel
       temp files) and pick the best match using **Learning Record** text from inside
       the file (via :func:`read_learning_record_name`) and the file stem — so names
       like ``20260311 Introduction to CyberOT.xlsx`` work.

    Returns ``None`` if no suitable file exists.
    """
    exclude = exclude or set()
    week_dir = week_dir.resolve()
    ex = {p.resolve() for p in exclude}

    exact = week_dir / f"{key}.xlsx"
    if exact.is_file() and exact.resolve() not in ex:
        return exact.resolve()

    key_a = _norm_alnum(key)
    key_parts = [p for p in re.split(r"[_\s]+", key.lower()) if len(p) > 2]

    candidates: list[Path] = []
    try:
        for f in sorted(week_dir.glob("*.xlsx")):
            if not f.is_file():
                continue
            if f.name.startswith("~$"):
                continue
            if _is_datamart_candidate(f):
                continue
            rp = f.resolve()
            if rp in ex:
                continue
            candidates.append(f)
    except OSError:
        return None

    best: tuple[int, str, Path] | None = None
    for f in candidates:
        lr = read_learning_record_name(f)
        lr_a = _norm_alnum(lr)
        stem_a = _norm_alnum(f.stem)
        score = 0
        tie = f.name.lower()
        if key_a and len(key_a) >= 2:
            if key_a in lr_a:
                score = 100 + min(len(key_a), len(lr_a))
            elif key_a in stem_a:
                score = 50 + min(len(key_a), len(stem_a))
        if score == 0 and lr and key.lower() in lr.lower():
            score = 40
        if score == 0 and key_parts and lr and all(p in lr.lower() for p in key_parts):
            score = 35
        if score == 0 and key_parts:
            for p in key_parts:
                if p in stem_a:
                    score = max(score, 25 + len(p))
        if score > 0:
            cand = (score, tie, f.resolve())
            if best is None or cand[0] > best[0] or (cand[0] == best[0] and cand[1] < best[1]):
                best = cand
    return best[2] if best else None


def find_datamart_file(folder: Path) -> Path:
    env_path = os.environ.get("DATAMART_PATH", "").strip()
    if env_path:
        p = Path(env_path).expanduser()
        if not p.is_absolute():
            p = folder.resolve() / p
        if p.is_file():
            return p.resolve()
        raise FileNotFoundError(f"DATAMART_PATH not found: {p}")

    folder = folder.resolve()
    candidates: list[Path] = []
    try:
        for f in folder.glob("*.xlsx"):
            if _is_datamart_candidate(f):
                candidates.append(f.resolve())
    except OSError:
        pass

    uniq = list({p for p in candidates})
    if not uniq:
        raise FileNotFoundError(
            f"No Datamart Excel file found in {folder} (name should contain "
            "'Datamart' ignoring spaces, e.g. Datamart.xlsx). "
            "Place the workbook next to the scripts, or set DATAMART_PATH."
        )
    for p in uniq:
        if p.name.lower() == "datamart.xlsx":
            return p.resolve()
    if len(uniq) > 1:
        uniq.sort(key=lambda p: p.name.lower())
        names = [x.name for x in uniq]
        raise FileNotFoundError(
            f"Multiple Datamart candidates in {folder}: {names}. "
            "Keep one file, rename one to Datamart.xlsx, or set DATAMART_PATH."
        )
    return uniq[0].resolve()


def read_datamart_table(path: Path) -> pd.DataFrame:
    """Load the first worksheet that looks like the employee Datamart."""
    xl = pd.ExcelFile(path, engine="openpyxl")
    for sheet in xl.sheet_names:
        df = _strip_columns(pd.read_excel(path, sheet_name=sheet, engine="openpyxl"))
        if _find_column(df, "Global Employee ID") and _find_column(df, "Employee Email"):
            return df
    return _strip_columns(pd.read_excel(path, sheet_name=0, engine="openpyxl"))


def load_datamart_excel(path: Path, progress: Progress, task_id: int) -> pd.DataFrame:
    progress.update(task_id, description="[cyan]Loading Datamart...")
    df = read_datamart_table(path)
    progress.advance(task_id)
    return df


def get_run_date_str() -> str:
    return date.today().strftime("%Y-%m-%d")


def ensure_output_and_processed_dirs(
    script_dir: Path, run_date: str, output_dir_name: str = "output"
) -> tuple[Path, Path]:
    script_dir = script_dir.resolve()
    out_dir = script_dir / output_dir_name
    processed_run_dir = script_dir / "processed" / run_date
    out_dir.mkdir(parents=True, exist_ok=True)
    processed_run_dir.mkdir(parents=True, exist_ok=True)
    return out_dir, processed_run_dir


def _input_week_context(script_dir: Path) -> tuple[str, str]:
    """If script_dir is like .../input/YYYY-MM/week-NN, return (month, week); else ('', '')."""
    script_dir = script_dir.resolve()
    week = script_dir.name
    month = script_dir.parent.name if script_dir.parent else ""
    if re.match(r"^\d{4}-\d{2}$", month) and re.match(r"^week-\d{2}$", week):
        return month, week
    return "", ""


def output_file_with_date(base_filename: str, run_date: str) -> str:
    p = Path(base_filename)
    return f"{p.stem}_{run_date}{p.suffix}"


def _safe_archive_filename(name: str) -> str:
    """Strip characters that are invalid in Windows paths; keep a safe basename."""
    bad = '<>:"/\\|?*\x00'
    out = "".join("_" if c in bad or ord(c) < 32 else c for c in name)
    out = out.strip(" .")
    return out or "archived.xlsx"


def _unique_dest_path(dest_dir: Path, filename: str) -> Path:
    """Return destination path that does not exist yet (avoid overwriting)."""
    dest_dir = dest_dir.resolve()
    dest = dest_dir / filename
    if not dest.exists():
        return dest
    p = Path(filename)
    stem, suf = p.stem, p.suffix
    n = 1
    while True:
        cand = dest_dir / f"{stem}_{n}{suf}"
        if not cand.exists():
            return cand
        n += 1


def archive_inputs_to_processed(processed_run_dir: Path, *paths: Path) -> None:
    """
    Copy inputs used for this run into ``processed/<run_date>/``.

    **Never modifies** the originals; only reads bytes from the source paths and writes
    copies. Destination names are sanitized (invalid path characters) and made unique
    so two uploads with tricky names cannot overwrite each other.
    """
    seen: set[str] = set()
    processed_run_dir = processed_run_dir.resolve()
    processed_run_dir.mkdir(parents=True, exist_ok=True)
    for src in paths:
        src = src.resolve()
        if not src.is_file():
            continue
        key = str(src)
        if key in seen:
            continue
        seen.add(key)
        safe_name = _safe_archive_filename(src.name)
        dest = _unique_dest_path(processed_run_dir, safe_name)
        if dest.parent.resolve() != processed_run_dir:
            raise ValueError(f"Unsafe archive destination: {dest}")
        shutil.copy2(src, dest)


def resolve_source_path(script_dir: Path, source_arg: str) -> Path:
    p = Path(source_arg)
    if p.is_absolute():
        return p.resolve()
    return (script_dir / p).resolve()


def _find_completion_column(df: pd.DataFrame) -> Optional[str]:
    for name in COMPLETION_DATE_ALIASES:
        c = _find_column(df, name)
        if c is not None:
            return c
    return None


def process_training_export(
    script_dir: Path,
    source_path: Path,
    output_filename: str,
    run_date: Optional[str] = None,
    show_progress: bool = True,
    state_db_path: Optional[Path] = None,
    output_dir_name: str = "output",
) -> tuple[Path, Path]:
    script_dir = script_dir.resolve()
    source_path = source_path.resolve()
    run_date = run_date or get_run_date_str()
    dated_name = output_file_with_date(output_filename, run_date)

    if not source_path.is_file():
        raise FileNotFoundError(f"Source file not found: {source_path}")

    datamart_path = find_datamart_file(script_dir)
    out_dir, processed_run_dir = ensure_output_and_processed_dirs(
        script_dir, run_date, output_dir_name=output_dir_name
    )
    out_path = out_dir / dated_name

    console = Console(stderr=True)

    def _body(progress: Progress) -> None:
        training_name = read_learning_record_name(source_path)
        if not training_name:
            training_name = source_path.stem

        t_load = progress.add_task("load", total=2)
        bsc = load_training_excel(source_path, progress, t_load)
        t_dm = progress.add_task("dm", total=1)
        dm = load_datamart_excel(datamart_path, progress, t_dm)

        col_status = _find_column(bsc, "Employee Status")
        col_emp_id = _find_column(bsc, "Employee ID")
        col_transcript = _find_column(bsc, "Transcript Status")
        col_email = _find_column(bsc, "Work Email Address")
        col_macro = _find_column(bsc, "Macro Entity Level 2 (Zone)")
        col_completion = _find_completion_column(bsc)

        missing = [
            n
            for n, c in [
                ("Employee ID", col_emp_id),
                ("Transcript Status", col_transcript),
                ("Work Email Address", col_email),
            ]
            if c is None
        ]
        if missing:
            raise ValueError(
                f"{source_path.name}: missing required column(s): {', '.join(missing)} "
                "(after header detection). Check the export layout."
            )

        dm_id = _find_column(dm, "Global Employee ID")
        dm_email = _find_column(dm, "Employee Email")
        dm_zone = _find_column(dm, "Zone")
        if not all([dm_id, dm_email, dm_zone]):
            raise ValueError(
                "Datamart must contain columns: "
                + ", ".join(DATAMART_REQUIRED_COLS)
            )

        t_lookup = progress.add_task("[cyan]Building zone lookups...", total=1)
        dm_clean = dm[[dm_id, dm_email, dm_zone]].dropna(how="all")
        id_to_zone: dict[str, str] = {}
        email_to_zone: dict[str, str] = {}
        id_to_geid: dict[str, str] = {}
        email_to_geid: dict[str, str] = {}
        for _, row in dm_clean.iterrows():
            z = row[dm_zone]
            if pd.isna(z) or str(z).strip() == "":
                continue
            z = str(z).strip()
            iid = _norm_id(row[dm_id])
            geid_val = _norm_key(row[dm_id])
            if iid and iid not in id_to_zone:
                id_to_zone[iid] = z
            if iid and geid_val and iid not in id_to_geid:
                id_to_geid[iid] = geid_val
            em = _norm_key(row[dm_email]).lower()
            if em and em not in email_to_zone:
                email_to_zone[em] = z
            if em and geid_val and em not in email_to_geid:
                email_to_geid[em] = geid_val
        progress.advance(t_lookup)

        if col_status:
            is_term = bsc[col_status].astype(str).str.strip().str.lower() == "terminated"
            terminated = bsc.loc[is_term].copy()
            active = bsc.loc[~is_term].copy()
        else:
            terminated = pd.DataFrame(columns=bsc.columns)
            active = bsc.copy()
            console.print(
                f"[yellow]Warning:[/yellow] '{source_path.name}' has no column 'Employee Status'; "
                "no rows moved to 'terminated employees'."
            )

        dm_term_col = _find_column(dm, "Termination Date")
        try:
            run_date_d = date.fromisoformat(str(run_date)[:10])
        except (TypeError, ValueError):
            run_date_d = date.today()

        id_to_term: dict[str, date] = {}
        email_to_term: dict[str, date] = {}
        if dm_term_col:
            for _, r in dm.iterrows():
                td = r[dm_term_col]
                if pd.isna(td):
                    continue
                ts = pd.to_datetime(td, errors="coerce")
                if pd.isna(ts):
                    continue
                tdt = ts.date()
                iid = _norm_id(r[dm_id])
                em = _norm_key(r[dm_email]).lower()
                if iid:
                    id_to_term[iid] = tdt
                if em:
                    email_to_term[em] = tdt

        if dm_term_col and (id_to_term or email_to_term):

            def _row_terminated_by_datamart(row) -> bool:
                eid = _norm_id(row[col_emp_id])
                em = _norm_key(row[col_email]).lower()
                td: Optional[date] = None
                if eid and eid in id_to_term:
                    td = id_to_term[eid]
                elif em and em in email_to_term:
                    td = email_to_term[em]
                if td is None:
                    return False
                return td <= run_date_d

            mask_dm_term = active.apply(_row_terminated_by_datamart, axis=1)
            if mask_dm_term.any():
                dm_term_rows = active.loc[mask_dm_term].copy()
                active = active.loc[~mask_dm_term].copy()
                terminated = (
                    pd.concat([terminated, dm_term_rows], ignore_index=True)
                    if len(terminated)
                    else dm_term_rows
                )

        proc = pd.DataFrame(
            {
                "Employee ID": active[col_emp_id].map(lambda x: _norm_id(x) or x),
                "Training Name": training_name,
                "Transcript Status": active[col_transcript].map(map_transcript_status),
                "Work Email Address": active[col_email],
            }
        )

        n_active = len(active)
        t_zone = progress.add_task("[cyan]Assigning zones & IDs...", total=max(n_active, 1))
        zones: list[str] = []
        geids: list[str] = []
        for _, row in active.iterrows():
            eid = _norm_id(row[col_emp_id])
            email = _norm_key(row[col_email]).lower()
            z = ""
            if eid and eid in id_to_zone:
                z = id_to_zone[eid]
            elif email and email in email_to_zone:
                z = email_to_zone[email]
            if not z and col_macro is not None:
                z = map_macro_zone(row[col_macro])
            zones.append(z)
            g = ""
            if eid and eid in id_to_geid:
                g = id_to_geid[eid]
            elif email and email in email_to_geid:
                g = email_to_geid[email]
            geids.append(g)
            progress.advance(t_zone)
        proc["Zone"] = zones
        proc["Global Employee ID"] = geids

        completion_dates: list[str] = []
        completion_sources: list[str] = []
        for i, row in active.iterrows():
            st = proc.at[i, "Transcript Status"]
            raw_c = row[col_completion] if col_completion is not None else None
            iso, src = resolve_completion_date(
                state_db_path,
                row[col_email],
                training_name,
                st,
                raw_c,
                run_date,
            )
            completion_dates.append(iso)
            completion_sources.append(src)
        proc["Transcript Completion Date"] = completion_dates

        proc = proc[
            [
                "Global Employee ID",
                "Employee ID",
                "Training Name",
                "Transcript Status",
                "Transcript Completion Date",
                "Work Email Address",
                "Zone",
            ]
        ]

        if state_db_path is not None:
            mo, wk = _input_week_context(script_dir)
            persist_completion_records(
                state_db_path,
                proc,
                run_date,
                mo,
                wk,
                completion_sources,
            )

        terminated_out = _strip_columns(terminated) if len(terminated) else pd.DataFrame()

        t_write = progress.add_task("[cyan]Writing output workbook...", total=1)
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            proc.to_excel(writer, sheet_name="processed data", index=False)
            terminated_out.to_excel(writer, sheet_name="terminated employees", index=False)
        progress.advance(t_write)

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(bar_width=40),
        TaskProgressColumn(),
        TimeElapsedColumn(),
        console=console,
        disable=not show_progress,
    ) as progress:
        _body(progress)

    archive_inputs_to_processed(processed_run_dir, source_path, datamart_path)

    return out_path, processed_run_dir


def count_new_joiners_in_datamart(datamart_path: Path, hire_start: date, as_of: date) -> int:
    """
    Rows in Datamart with First Hire Date between hire_start and as_of (inclusive).
    Mirrors the window used for new-joiner assignment reporting.
    """
    dm = read_datamart_table(datamart_path)
    c_hire = _find_column(dm, "First Hire Date")
    if c_hire is None:
        return 0
    n = 0
    for _, row in dm.iterrows():
        ts = pd.to_datetime(row[c_hire], errors="coerce")
        if pd.isna(ts):
            continue
        d = ts.date()
        if hire_start <= d <= as_of:
            n += 1
    return n


def resolve_training_work_dir() -> Path:
    """Use trainings/<name>/ when CLI lives in that module's scripts/ folder."""
    entry = Path(sys.argv[0]).resolve().parent
    module_root = entry.parent
    if (module_root / "input").is_dir() and (module_root / "output").is_dir():
        return module_root
    return entry


def main_for_source(default_source: str, default_output: str) -> None:
    script_dir = resolve_training_work_dir()
    source_arg = sys.argv[1] if len(sys.argv) > 1 else default_source
    source_path = resolve_source_path(script_dir, source_arg)
    try:
        out, processed_run_dir = process_training_export(
            script_dir, source_path, default_output
        )
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    print(f"Wrote: {out}")
    print(f"Archived inputs used for this run to: {processed_run_dir}")


if __name__ == "__main__":
    main_for_source("BSC.xlsx", "BSC_training_output.xlsx")
