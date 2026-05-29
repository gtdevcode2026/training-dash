# ABInBev Training Status Automation

Professional layout: one folder per training module, shared libraries, and a single browser dashboard.

## Project structure

```
training_status_new/
├── README.md
├── requirements.txt
├── dashboard/
│   └── index.html          # Open in a browser (main UI)
├── shared/
│   ├── python/             # training_processor.py, completion_state.py
│   └── javascript/         # Dashboard logic (config, Excel utils, processors, UI)
└── trainings/
    ├── application_security/
    ├── cyber_ot/
    ├── growth_group/
    ├── new_joiner/
    ├── bsc/
    ├── phishing_normal/
    └── band4_senior_management/
```

Each training folder contains:

| Path | Purpose |
|------|---------|
| `config.json` | Module metadata (dashboard tab id, LMS name, default filenames) |
| `README.md` | How to run this training |
| `scripts/` | Python CLI for that training (where applicable) |
| `input/` | Drop Datamart + LMS exports before CLI runs |
| `output/` | Dated Excel results |
| `processed/` | Dated copies of inputs used per run |
| `samples/` | Optional reference workbooks |

## Quick start (dashboard)

1. Open `dashboard/index.html` in Chrome or Edge.
2. Use the tabs for each training; upload tool/base Excel files as described on each panel.
3. Click **Calculate** or **Process All Sections**.

## Quick start (Python CLI)

```powershell
cd c:\Users\manka\Downloads\training_status_new
pip install -r requirements.txt
```

Example for Cyber OT:

```powershell
cd trainings\cyber_ot\scripts
# Copy CyberOT.xlsx and Datamart.xlsx into ..\input\
python process_cyberot.py "..\input\CyberOT.xlsx"
```

Outputs go under `trainings/cyber_ot/output/` and snapshots under `trainings/cyber_ot/processed/YYYY-MM-DD/`.

## Training index

| Folder | Dashboard tab | CLI script |
|--------|---------------|------------|
| `application_security` | App Security | Dashboard only |
| `cyber_ot` | Cyber OT | `process_cyberot.py` |
| `growth_group` | Growth Group | Dashboard only |
| `new_joiner` | New Joiner | `assign_newjoiners_training.py` |
| `bsc` | BSC | `process_bsc.py` |
| `phishing_normal` | Phishing Normal | `process_phished.py` |
| `band4_senior_management` | Band 4+ | `process_dirphished.py` |

## Legacy note

The previous single-file dashboard was `dash2.html` at the repo root. It is replaced by `dashboard/index.html` plus `shared/javascript/`. You can delete `dash2.html` after verifying the new layout.
