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

## CLI (Python)

From this folder:

```powershell
cd scripts
pip install -r ../../requirements.txt
python process_dirphished.py
```

Place `DirPhished.xlsx` and a Datamart `*.xlsx` in `../input/` (or pass paths as arguments where supported).
