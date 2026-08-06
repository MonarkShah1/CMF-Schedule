# Script Reference

---

## cmf_schedule_builder.py
**Run daily. Reads MAIN → rebuilds CALENDAR + PURCHASING.**

### Key functions

| Function | What it does |
|----------|-------------|
| `parse_main(ws)` | Reads all part rows from MAIN, builds process-due entries + image map. Filters completed steps using date comparison. Computes `processes_left` sorted by actual due date. |
| `build_calendar(wb, entries, row_to_img)` | Generates CALENDAR tab. 4-week rolling date-bucket with exact process names shown inside grouped stations. |
| `build_purchasing(wb, entries, row_to_img)` | Generates PURCHASING tab. Flat list of all active outside/vendor steps with photos. |
| `protect_workbook_images(wb)` | Wraps all loaded image BytesIO refs in `_UnclosableBytesIO` before any processing. |
| `normalize_image_anchors(wb)` | Forces existing images to save as move-and-size-with-cells so Excel sorting keeps screenshots on the correct rows. |
| `copy_image(src, ws, col, row)` | Copies image using TwoCellAnchor so it hides with filtered rows. |

### Constants to know

```python
COL_CURRENT_STEP = 37   # col AK — where part is now
COL_STATUS       = 38   # col AL — COMPLETE / ON HOLD / etc.
BLUE_FIRST       = 12   # col L  — first process due date col
BLUE_LAST        = 36   # col AJ — DELIVERY DATE

STEP_TO_COL = { "LASER CUT": 14, "WELD": 25, ... }  # used for date-based filtering
COL_TO_PROCESS = { 14: "LASER CUT", 25: "WELD", ... } # used for PROCESSES LEFT
```

### STATIONS list
Controls which columns feed into which station display in CALENDAR and PURCHASING.
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

## cmf_dedupe_media.py
**Runs automatically at the end of `cmf_schedule_builder.py`.**

openpyxl stores one copy of an image per placement, so a screenshot shown on
the MAIN row plus every CALENDAR and PURCHASING card is written dozens of times
(measured: 920 stored parts for 139 unique images). This pass hashes the media,
keeps one copy of each, and repoints the drawing relationships.

Measured on the real workbook: **17.7 MB → 4.0 MB (77% smaller)**, with all 920
image placements intact and 0 broken. Safe to re-run; stable across cycles.

```bash
python3 cmf_dedupe_media.py                 # the working file
python3 cmf_dedupe_media.py other.xlsx      # any workbook
```

---

## cmf_merge.py
**Run when boss sends an updated WIP file.**

Takes two inputs:
- `CMF WIP (1).xlsx` — boss's latest orders (old format)
- `CMF WIP - Schedule (N).xlsx` — engineer's routing (new format)

Inputs are chosen by **most recent modification time**, and the script prints
every candidate plus the one it picked. It warns when the newest file is more
than 48 hours old, or when two candidates were modified within 2 minutes —
both signs the wrong export is about to be used.

> Earlier versions picked by the highest `(N)` suffix, so a freshly downloaded
> `CMF WIP.xlsx` silently lost to a stale `CMF WIP (1).xlsx`. See
> [`POSTMORTEM.md`](POSTMORTEM.md).

### Flags

| Flag | Effect |
|------|--------|
| `--dry-run` | Writes `Merge Review.xlsx` and stops. Nothing is overwritten. |

### Outputs

- `CMF WIP - Schedule.xlsx` — merged result (a backup is written first)
- `Merge Review.xlsx` — new WOs, removed WOs, ship-date changes, parts needing routing

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
