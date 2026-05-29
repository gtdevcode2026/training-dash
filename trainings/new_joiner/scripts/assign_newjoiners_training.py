"""
New joiners report from Datamart: employees with First Hire Date from a given start date through today.
Input file is read-only; output is a new workbook with a full copy plus the filtered extract.
"""

from __future__ import annotations

import re
import sys
from datetime import date, datetime
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[3]
_SHARED = _ROOT / "shared" / "python"
if str(_SHARED) not in sys.path:
    sys.path.insert(0, str(_SHARED))

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

from training_processor import (
    ASSIGN_NEW_JOINERS_REQUIRED_COLS,
    archive_inputs_to_processed,
    ensure_output_and_processed_dirs,
    find_datamart_file,
    get_run_date_str,
    output_file_with_date,
    read_datamart_table,
    resolve_training_work_dir,
    _find_column,
)


def _parse_hire_date(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return pd.NaT
    if isinstance(val, (datetime, pd.Timestamp)):
        ts = pd.Timestamp(val)
        if pd.isna(ts):
            return pd.NaT
        return ts.normalize()
    if isinstance(val, date):
        return pd.Timestamp(val)
    s = str(val).strip()
    if not s:
        return pd.NaT
    if re.fullmatch(r"\d+(\.\d+)?", s):
        try:
            return pd.to_datetime(float(s), unit="d", origin="1899-12-30", utc=False)
        except Exception:
            pass
    return pd.to_datetime(s, errors="coerce")


def main() -> None:
    script_dir = resolve_training_work_dir()
    console = Console(stderr=True)

    try:
        dm_path = find_datamart_file(script_dir)
    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    raw_in = input("Enter start date for First Hire Date filter (YYYY-MM-DD): ").strip()
    try:
        start = pd.Timestamp(datetime.strptime(raw_in, "%Y-%m-%d").date())
    except ValueError:
        print("Error: Use format YYYY-MM-DD, e.g. 2024-01-15", file=sys.stderr)
        sys.exit(1)

    today = pd.Timestamp.now().normalize()

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(bar_width=40),
        TaskProgressColumn(),
        TimeElapsedColumn(),
        console=console,
    ) as progress:
        t1 = progress.add_task("[cyan]Loading Datamart...", total=1)
        dm = read_datamart_table(dm_path)
        progress.advance(t1)

        found = {n: _find_column(dm, n) for n in ASSIGN_NEW_JOINERS_REQUIRED_COLS}
        missing = [n for n, c in found.items() if c is None]
        if missing:
            print(
                f"Error: Datamart missing column(s): {', '.join(missing)}",
                file=sys.stderr,
            )
            sys.exit(1)

        t2 = progress.add_task("[cyan]Parsing hire dates and filtering...", total=1)
        c_id = found["Global Employee ID"]
        c_name = found["Employee Name"]
        c_zone = found["Zone"]
        c_email = found["Employee Email"]
        c_hire = found["First Hire Date"]

        hire_parsed = dm[c_hire].map(_parse_hire_date)
        dm_work = dm.copy()
        dm_work["_hire_parsed"] = hire_parsed

        mask = (dm_work["_hire_parsed"] >= start) & (dm_work["_hire_parsed"] <= today)
        filtered = dm_work.loc[mask, [c_id, c_name, c_zone, c_email]].copy()
        filtered.columns = [
            "Global Employee ID",
            "Employee Name",
            "Zone",
            "Employee Email",
        ]
        progress.advance(t2)

        run_date = get_run_date_str()
        out_dir, processed_run_dir = ensure_output_and_processed_dirs(script_dir, run_date)

        dated_name = output_file_with_date("assign_newjoiners_training.xlsx", run_date)
        out_path = out_dir / dated_name
        full_copy = dm.copy()

        t3 = progress.add_task("[cyan]Writing workbook...", total=1)
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            full_copy.to_excel(writer, sheet_name="Datamart copy", index=False)
            filtered.to_excel(writer, sheet_name="New joiners", index=False)
        progress.advance(t3)

    archive_inputs_to_processed(processed_run_dir, dm_path)

    print(f"Wrote: {out_path}")
    print(f"Archived Datamart snapshot to: {processed_run_dir}")
    print(
        f"Rows with First Hire Date from {start.date()} through {today.date()}: {len(filtered)}"
    )


if __name__ == "__main__":
    main()
