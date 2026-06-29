# Band 4+ Senior Management

Dashboard section ID: `band4`  
LMS training name: Security & Compliance Awareness Training (Sr. Management) #ABI

## Folders

| Folder | Purpose |
|--------|---------|
| `input/` | Place Datamart and weekly LMS exports here before running scripts |
| `output/` | Dated Excel outputs from CLI scripts |
| `processed/` | Dated snapshots of inputs used per run |
| `samples/` | Optional reference workbooks (not used automatically) |
| `scripts/` | Training-specific automation |

## Browser dashboard

Open `../../dashboard/index.html` and use the **Band 4+ Senior Management** tab.

**Calculate** enriches the uploaded **Phishing Tracking Input** userbase as-is (no extra Band 4 filter in code). The **userbase is the master list** (absolute Total). Each user is matched to the tool file by **Employee Email**, **Global Employee ID**, and **Local Employee ID** (any hit counts). Appended columns (after the last filled userbase header):

- `Training Start Date`
- `Training completion date`
- `training completion status` — Completed, Not completed (in tool), or Not found (not in tool)

| Metric | Meaning |
|--------|---------|
| **Total (userbase)** | Unique users in uploaded userbase (deduped by email → global ID → local ID) |
| **Completed** | Found in tool and completed |
| **Not completed** | Found in tool but not completed |
| **Not found** | In userbase but not in tool |

**Download Updated Userbase** exports the full workbook with all original columns plus appended fields (single sheet). **Download Pending Users** exports uploaded rows that are not completed.

## CLI (Python)

From this folder:

```powershell
cd scripts
pip install -r ../../requirements.txt
python process_dirphished.py
```

Place `DirPhished.xlsx` and a Datamart `*.xlsx` in `../input/` (or pass paths as arguments where supported).
