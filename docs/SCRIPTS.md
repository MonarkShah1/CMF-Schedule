# Script Reference

---

## cmf_schedule_builder.py
**Run daily. Reads MAIN → rebuilds CALENDAR + TODAY.**

### Key functions

| Function | What it does |
|----------|-------------|
| `parse_main(ws)` | Reads all part rows from MAIN, builds process-due entries + image map. Filters completed steps using date comparison. Computes `processes_left` sorted by actual due date. |
| `build_calendar(wb, entries, row_to_img)` | Generates CALENDAR tab. 4-week rolling date-bucket. Always-visible overdue section. |
| `build_today(wb, entries, row_to_img)` | Generates TODAY kanban. Active stations as column groups. TwoCellAnchor images. |
| `protect_workbook_images(wb)` | Wraps all loaded image BytesIO refs in `_UnclosableBytesIO` before any processing. |
| `copy_image(src, ws, col, row)` | Copies image using TwoCellAnchor so it hides with filtered rows. |

### Constants to know

```python
COL_CURRENT_STEP = 36   # col AJ — where part is now
COL_STATUS       = 37   # col AK — COMPLETE / ON HOLD / etc.
BLUE_FIRST       = 12   # col L  — first process due date col
BLUE_LAST        = 35   # col AI — DELIVERY DATE

STEP_TO_COL = { "LASER CUT": 14, "WELD": 25, ... }  # used for date-based filtering
COL_TO_PROCESS = { 14: "LASER CUT", 25: "WELD", ... } # used for PROCESSES LEFT
```

### STATIONS list
Controls which columns feed into which station display in CALENDAR and TODAY.
Each entry: `(key, display_name, header_hex, row_hex, [col_indices])`

---

## cmf_migrate_main.py
**Run once. Converts old WIP format → new Schedule format.**

- Reads `CMF WIP.xlsx` (old 15-col format)
- Pre-groups rows by WO — ensures every WO gets 1 navy header + ≥1 grey part row
- Copies screenshots using TwoCellAnchor
- Outputs `CMF WIP - Schedule.xlsx`

### Why this exists
The old MAIN had one row per part-step (denormalized). The new format has one row per part with all process due dates as columns (normalized). Migration converts between the two.

---

## cmf_merge.py
**Run when boss sends an updated WIP file.**

Takes two inputs:
- `CMF WIP (1).xlsx` — boss's latest orders (old format)
- `CMF WIP - Schedule (N).xlsx` — engineer's routing (new format)

Merge logic:
1. Re-migrates all orders from boss's file (gets new WOs, updated dates)
2. Builds routing lookup from engineer's file by **(WO#, Part#)** key
3. Part-level match wins over WO-level match
4. Delivery date pre-filled from old DUE DATE if not in routing
5. Images pulled from boss's file first (more up-to-date), then engineer's file as fallback

---

## Refresh Schedule.bat
**Windows double-click runner.**

1. Checks Python is installed
2. Installs `openpyxl` and `Pillow` if missing (silent)
3. Runs `cmf_schedule_builder.py`
4. Pauses so you can read any errors

**Update the file paths** at the top of the .bat if you move the folder.

---

## Common Errors

| Error | Cause | Fix |
|-------|-------|-----|
| `ValueError: Invalid format string` | `%-m` Mac-only date format ran on Windows | Use `f"{dv.month}/{dv.day}/..."` |
| `ValueError: I/O operation on closed file` | openpyxl `_data()` closes BytesIO after reading | Ensure `_UnclosableBytesIO` is used |
| `KeyError: [Content_Types].xml` | Excel had file open during script save | Close Excel before running script |
| `No MAIN sheet found` | Wrong file or migration not run | Run `cmf_migrate_main.py` first |
