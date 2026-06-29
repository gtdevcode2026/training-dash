# ABInBev Training Status Automation

Simple **HTML dashboard** (browser-only) + **Python CLI** with **SQLite** run history.

## Project structure

```
training_status_new/
├── dashboard/index.html       # Main UI — open directly in Chrome/Edge
├── dash2.html                 # Shortcut → dashboard
├── docs/TRAINING_CALCULATION_STEPS.md   # Exact steps per training
├── tool/run_calculation.py    # Python CLI
├── data/training_status.db    # Created on first CLI run (SQLite)
├── shared/
│   ├── javascript/app.js      # In-browser calculations
│   └── python/                # Calculators, SQLite, legacy CLI helpers
└── trainings/<module>/        # input/, output/, scripts/ per training
```

## Quick start — HTML (no server)

1. Open **`dashboard/index.html`** in Chrome or Edge (double-click the file).
2. Upload Excel files on each tab and click **Calculate**.
3. Download the updated userbase or tool output Excel.

No localhost, no Flask, no install required for the dashboard.

**Calculation details:** see [docs/TRAINING_CALCULATION_STEPS.md](docs/TRAINING_CALCULATION_STEPS.md).

## Quick start — Python + SQLite

```powershell
cd c:\Users\manka\Downloads\training_status_new
pip install -r requirements.txt
python tool/run_calculation.py --list
```

Example (BSC):

```powershell
python tool/run_calculation.py bsc `
  --base trainings\bsc\input\userbase.xlsx `
  --tool trainings\bsc\input\BSC.xlsx `
  --assign-date 2026-05-29
```

- Output: `trainings/bsc/output/bsc_Updated_Userbase_YYYY-MM-DD.xlsx`
- Run logged in `data/training_status.db` (`calculation_runs` table)
- History: `python tool/run_calculation.py --history`

## Training index

| Folder | Dashboard tab | Needs userbase? | CLI |
|--------|---------------|-----------------|-----|
| `application_security` | App Security | No | `run_calculation.py app_sec --tool …` |
| `cyber_ot` | Cyber OT | No | `run_calculation.py cyber_ot --tool …` |
| `growth_group` | Growth Group | No | `run_calculation.py growth --tool …` |
| `new_joiner` | New Joiner | No | Dashboard or `assign_newjoiners_training.py` |
| `bsc` | BSC | Yes | `run_calculation.py bsc --base … --tool …` |
| `phishing_normal` | Phishing Normal | Yes | `run_calculation.py phishing_normal --base … --tool …` |
| `band4_senior_management` | Band 4+ | Yes | `run_calculation.py band4 --base … --tool …` |

## SQLite

| Table | Purpose |
|-------|---------|
| `calculation_runs` | CLI run history (counts, paths, timestamps) |
| `completion_records` | Legacy completion-date backfill (existing Python exports) |

Database path: **`data/training_status.db`**

## Optional legacy server

`server/upload_server.py` and `scripts/start_server.ps1` are **not required**. They only saved uploads to disk when the dashboard was served over HTTP.
add 