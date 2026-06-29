# BSC

Dashboard section ID: `bsc`  
LMS training name: Don't take the bait — Avoid Financial Impact of Phishing #ABI

## Folders

| Folder | Purpose |
|--------|---------|
| `input/` | Place Datamart and weekly LMS exports here before running scripts |
| `output/` | Dated Excel outputs from CLI scripts |
| `processed/` | Dated snapshots of inputs used per run |
| `samples/` | Optional reference workbooks (not used automatically) |
| `scripts/` | Training-specific automation |

## Browser dashboard

Open `../../dashboard/index.html` and use the **BSC** tab.

**Calculate** walks every **userbase** row, matches the tool file by **Emp ID** and **email**, and appends (after the last filled userbase column):

- `start date_extracted` — BSC annual/phishing logic (original start date columns are left unchanged)
- `Training completion date`
- `training completion status`

**Download Updated Userbase** exports the full userbase with all original columns plus appended fields (single sheet). **Download Pending Users** exports rows that are not completed.

## CLI (Python)

From this folder:

```powershell
cd scripts
pip install -r ../../requirements.txt
python process_bsc.py
```

Place `BSC.xlsx` and a Datamart `*.xlsx` in `../input/` (or pass paths as arguments where supported).
