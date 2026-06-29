# Dashboard upload storage

When you run the local server (`scripts/start_server.ps1`) and open the dashboard in the browser, every Excel you upload is copied here automatically.

## Layout (per training module)

```
trainings/<module_slug>/
├── config.json
├── input/              ← manual drops for Python CLI
├── output/             ← CLI results
├── processed/          ← CLI archived inputs
└── uploads/            ← dashboard auto-save
    ├── base/
    │   ├── .gitkeep
    │   ├── _latest.xlsx      ← most recent base upload
    │   └── YYYY-MM-DD/
    │       └── HHMMSS_<original-filename>.xlsx
    └── tool/
        ├── .gitkeep
        ├── _latest.xlsx
        └── YYYY-MM-DD/
            └── HHMMSS_<original-filename>.xlsx
```

## Module slugs

| Dashboard tab   | Folder                         |
|-----------------|--------------------------------|
| App Security    | `application_security`         |
| Cyber OT        | `cyber_ot`                     |
| Growth Group    | `growth_group`                 |
| New Joiner      | `new_joiner`                   |
| BSC             | `bsc`                          |
| Phishing Normal | `phishing_normal`              |
| Band 4+         | `band4_senior_management`      |

Tool-only tabs still have both `base/` and `tool/` folders; only the folders you use will get files.
