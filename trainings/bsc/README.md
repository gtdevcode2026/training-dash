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

## CLI (Python)

From this folder:

```powershell
cd scripts
pip install -r ../../requirements.txt
python process_bsc.py
```

Place `BSC.xlsx` and a Datamart `*.xlsx` in `../input/` (or pass paths as arguments where supported).
