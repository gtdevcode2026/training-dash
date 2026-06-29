#!/usr/bin/env python3
"""
Run a training calculation from the command line.

Examples:
  python tool/run_calculation.py bsc --base trainings/bsc/input/userbase.xlsx --tool trainings/bsc/input/tool.xlsx
  python tool/run_calculation.py phishing_normal --base ... --tool ... --assign-date 2026-05-29
  python tool/run_calculation.py --list
  python tool/run_calculation.py --history
"""

from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[1]
_SHARED = _ROOT / "shared" / "python"
if str(_SHARED) not in sys.path:
    sys.path.insert(0, str(_SHARED))

from dashboard_calculators import TRAINING_CALCULATORS  # noqa: E402
from run_database import DEFAULT_DB_PATH, list_runs, record_run  # noqa: E402
from training_processor import get_run_date_str, output_file_with_date  # noqa: E402


def _run(training_key: str, base: Path | None, tool: Path, assign_date: str, out_dir: Path | None) -> Path:
    import pandas as pd

    entry = TRAINING_CALCULATORS.get(training_key)
    if not entry:
        raise SystemExit(f"Unknown training: {training_key}. Use --list.")

    slug, fn, needs_base = entry
    if needs_base and not base:
        raise SystemExit(f"{training_key} requires --base (userbase Excel).")

    run_date = get_run_date_str()
    if needs_base:
        df, metrics = fn(base, tool, assign_date)
    else:
        df, metrics = fn(tool, assign_date)

    target_dir = out_dir or (_ROOT / "trainings" / slug / "output")
    target_dir.mkdir(parents=True, exist_ok=True)
    filename = output_file_with_date(f"{slug}_Updated_Userbase.xlsx" if needs_base else f"{slug}_Output.xlsx", run_date)
    out_path = target_dir / filename

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        sheet = "Userbase" if needs_base else "All Users"
        df.to_excel(writer, sheet_name=sheet, index=False)

    pending = metrics.get("pending", metrics.get("not_completed", 0) + metrics.get("not_found", 0))
    record_run(
        training_key,
        run_date=run_date,
        assign_date=assign_date,
        base_file=str(base.resolve()) if base else "",
        tool_file=str(tool.resolve()),
        output_file=str(out_path.resolve()),
        total_rows=int(metrics.get("total", len(df))),
        completed=int(metrics.get("completed", 0)),
        pending=int(pending),
        not_found=int(metrics.get("not_found", 0)),
        metrics=metrics,
        notes=f"CLI run {training_key}",
    )
    return out_path


def main() -> None:
    parser = argparse.ArgumentParser(description="ABInBev training status calculator (SQLite-backed)")
    parser.add_argument("training", nargs="?", help="Training id (bsc, phishing_normal, band4, …)")
    parser.add_argument("--base", type=Path, help="Userbase / tracking input Excel")
    parser.add_argument("--tool", type=Path, help="LMS tool export Excel")
    parser.add_argument("--assign-date", default=date.today().isoformat(), help="YYYY-MM-DD")
    parser.add_argument("--output-dir", type=Path, help="Override output folder")
    parser.add_argument("--list", action="store_true", help="List supported trainings")
    parser.add_argument("--history", action="store_true", help="Show recent runs from SQLite")
    parser.add_argument("--db", type=Path, default=DEFAULT_DB_PATH, help="SQLite database path")
    args = parser.parse_args()

    if args.list:
        print("Supported trainings:")
        seen = set()
        for key, (slug, _, needs_base) in sorted(TRAINING_CALCULATORS.items()):
            if slug in seen:
                continue
            seen.add(slug)
            inputs = "base + tool" if needs_base else "tool only"
            print(f"  {slug:30} ({inputs})")
        print(f"\nSQLite database: {DEFAULT_DB_PATH}")
        print("Completion dates (LMS): same DB, table completion_records")
        return

    if args.history:
        rows = list_runs(limit=25, db_path=args.db)
        if not rows:
            print("No runs recorded yet.")
            return
        for r in rows:
            print(
                f"#{r['id']} {r['training_id']} {r['created_at']} "
                f"total={r['total_rows']} completed={r['completed']} pending={r['pending']} "
                f"not_found={r['not_found']}"
            )
            if r["output_file"]:
                print(f"    -> {r['output_file']}")
        return

    if not args.training or not args.tool:
        parser.error("Provide training id and --tool (and --base for userbase trainings).")

    out = _run(args.training, args.base, args.tool, args.assign_date, args.output_dir)
    print(f"Wrote: {out}")
    print(f"Logged run in: {args.db}")


if __name__ == "__main__":
    main()
