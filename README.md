# CMF Production Scheduling System

Production scheduling system for **Canadian Metal Fabricators Ltd.**  
Tracks work orders from order entry through shipping across all shop floor stations.

---

## What This Is

A Python-driven Excel workbook (`CMF WIP - Schedule.xlsx`) with three tabs that automatically refresh from a single source of truth.

```
MAIN  →  CALENDAR  →  TODAY (TV Display)
```

| Tab | Who Uses It | Purpose |
|-----|-------------|---------|
| **MAIN** | Engineer / Boss | Enter orders, fill routing dates, update current step |
| **CALENDAR** | Production Manager | See what's due when — plan the week |
| **TODAY** | Shop Floor (TV) | Kanban board showing today's + overdue work by station |

---

## Quick Start

### Mac
```bash
python3 cmf_schedule_builder.py
```

### Windows
Double-click **`Refresh Schedule.bat`**

> First time on Windows: script auto-installs `openpyxl` and `Pillow` via pip.

---

## Files

| File | Purpose | Run When |
|------|---------|---------|
| `cmf_schedule_builder.py` | Rebuilds CALENDAR + TODAY from MAIN | Daily |
| `cmf_migrate_main.py` | One-time migration from old WIP format | Once (already done) |
| `cmf_merge.py` | Merges boss's WIP updates + engineer's routing | When boss sends new WIP |
| `Refresh Schedule.bat` | Windows double-click runner | Daily (Windows) |
| `CMF WIP - Schedule.xlsx` | The working Excel file | Always open |

---

## Daily Workflow

```
1. Engineer opens MAIN → fills in process due dates (blue columns L–AI)
   and updates CURRENT STEP dropdown as parts move through the shop

2. Run: Refresh Schedule.bat  (Windows)
        python3 cmf_schedule_builder.py  (Mac)

3. CALENDAR refreshes → PM plans the day
   TODAY refreshes    → display on factory TV
```

---

## When Boss Sends a New WIP File

```bash
# Put the new WIP file in the same folder, then:
python3 cmf_merge.py
python3 cmf_schedule_builder.py
```

The merge script pulls new/updated orders from the boss's file and overlays the engineer's routing data.

---

## Requirements

- Python 3.8+
- `openpyxl` — Excel read/write
- `Pillow` — image handling

```bash
pip install openpyxl Pillow
```
