# New Joiner

Dashboard section ID: `newJoiner`  
LMS training name: IT Security Awareness - English #ABI

## Folders

| Folder | Purpose |
|--------|---------|
| `input/` | Place Datamart and weekly LMS exports here before running scripts |
| `output/` | Dated Excel outputs from CLI scripts |
| `processed/` | Dated snapshots of inputs used per run |
| `samples/` | Optional reference workbooks (not used automatically) |
| `scripts/` | Training-specific automation |

## Browser dashboard

Open `../../dashboard/index.html` and use the **New Joiner** tab.

**Year tabs (2024 / 2025 / 2026)** use **`Original Hire Date`** only (normalized to dd/mm/yyyy). Rows outside 2024–2026 are excluded. Terminated employees (`Employee Status = Terminated`) are removed at parse.

## CLI (Python)

From this folder:

```powershell
cd scripts
pip install -r ../../requirements.txt
python assign_newjoiners_training.py
```

Place `Datamart.xlsx` and a Datamart `*.xlsx` in `../input/` (or pass paths as arguments where supported).
