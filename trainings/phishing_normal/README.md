# Phishing Normal Users

Dashboard section ID: `phishingNormal`  
LMS training name: Security & Compliance Awareness : Phishing #ABI

## Folders

| Folder | Purpose |
|--------|---------|
| `input/` | Place Datamart and weekly LMS exports here before running scripts |
| `output/` | Dated Excel outputs from CLI scripts |
| `processed/` | Dated snapshots of inputs used per run |
| `samples/` | Optional reference workbooks (not used automatically) |
| `scripts/` | Training-specific automation |

## Browser dashboard

Open `../../dashboard/index.html` and use the **Phishing Normal Users** tab.

**Calculate** enriches the uploaded **Phishing Tracking Input** as-is (no extra BSC/Band4 filtering in code). The **userbase is the master list** (absolute Total). Match tool file by **Employee Email**, **Global Employee ID**, and **Local Employee ID** (any hit counts). Appended columns after the last filled userbase header:

- `Training Start Date`
- `Training completion date`
- `training completion status`

| Metric | Meaning |
|--------|---------|
| **Total (userbase)** | Unique users in uploaded userbase (deduped by email → global ID → local ID) |
| **Completed** | Found in tool and completed |
| **Not completed** | Found in tool but not completed |
| **Not found** | In userbase but not in tool |

**Download Updated Userbase** — full userbase with original columns + appended fields (single sheet).  
**Download Pending Users** — uploaded userbase rows that are not completed.

## CLI (Python)

From this folder:

```powershell
cd scripts
pip install -r ../../requirements.txt
python process_phished.py
```

Place `Phished.xlsx` and a Datamart `*.xlsx` in `../input/` (or pass paths as arguments where supported).
