# System Architecture

## Data Flow

```
CMF WIP.xlsx  (boss's original file — old format)
      │
      ▼
cmf_migrate_main.py  ──────────────────────────────────┐
      │                                                 │
      │  Re-run if boss sends updated WIP:              │
      │  cmf_merge.py                                   │
      ▼                                                 │
CMF WIP - Schedule.xlsx  (working file)                 │
  ├── MAIN        ← engineer edits manually             │
  ├── CALENDAR    ← script generates                    │
  └── PURCHASING  ← script generates                    │
                                                        │
cmf_schedule_builder.py  ◄──────────────────────────────┘
      reads MAIN, writes CALENDAR + PURCHASING
```

---

## MAIN Sheet Structure

**39 columns (A → AM)**

### Order Info (A–E) — dark navy header rows per WO
| Col | Field | Notes |
|-----|-------|-------|
| A | WO # | Work order number |
| B | PO # | Customer PO number |
| C | COMPANY | Customer name |
| D | JOB NAME | Job description (on header row) |
| E | CUST SHIP | Customer requested ship date |

### Part Info (F–K) — grey alternating part rows
| Col | Field | Notes |
|-----|-------|-------|
| F | PART # | Part number / code |
| G | DESCRIPTION | Part description |
| H | QTY | Quantity |
| I | SCREENSHOT | Engineer pastes photo here |
| J | MATERIAL | Material type |
| K | THICKNESS | Material thickness |

### Process Due Dates (L–AJ) — blue columns, engineer fills in
Each column = a due date for that process step. Leave blank if part doesn't go through that step.

| Col | Process | Col | Process |
|-----|---------|-----|---------|
| L | MATERIALS | X | GRIND |
| M | ENGINEERING | Y | WELD |
| N | LASER CUT | Z | MACHINE (O) |
| O | LASER CUT (O) | AA | PLATING (O) |
| P | TUBE LASER (O) | AB | PAINT |
| Q | SAW | AC | PAINT (O) |
| R | BANDSAW (O) | AD | SPECIAL |
| S | BEND | AE | SPECIAL (O) |
| T | CLEAN | AF | WHOLE JOB (O) |
| U | CSK | AG | HARDWARE |
| V | DRILL | AH | ASSEMBLY |
| W | TAPPING | AI | SHIP TO VENDOR |
|  |  | AJ | DELIVERY DATE |

### Status (AK–AM)
| Col | Field | Notes |
|-----|-------|-------|
| AK | CURRENT STEP | Dropdown — where the part is now |
| AL | STATUS | e.g. COMPLETE, ON HOLD |
| AM | NOTES | Old routing reference / misc notes |

### Row Types
- **Navy row** — WO header (one per work order). Boss scans these to see open POs.
- **Grey/white row** — Part row (one per part). Engineer fills in routing.

### Conditional Formatting
- Date < today → **red** (overdue)
- Date ≤ today + 3 → **yellow** (due soon)

---

## CALENDAR Tab

Horizontal station board with one row per scheduled process entry.

Each station block contains:

- PROCESS
- DUE
- SHIP
- COMPANY
- WO #
- PART #
- QTY
- UPDATED BY
- PHOTO

### How it works
- Rolling **4-week Mon–Fri** window from today
- **Always-visible overdue section** at top (red = items overdue, green = all clear)
- Each calendar row = one process step with a due date
- Grouped stations still show the **exact process name** on each row (for example `▶ DRILL` or `⏳ PAINT (O)`)
- **Delivery Date is NOT a calendar event** — it's a column on every row
- Completed processes are filtered out using the **date-based filter**:
  - Find the due date of CURRENT STEP
  - Only show processes due on or after that date
- PROCESSES LEFT sorted by actual due dates (not column order)

---

## PURCHASING Tab

Flat list of every active outside/vendor process:

- SHIP TO VENDOR
- LASER CUT (O)
- TUBE LASER (O)
- BANDSAW (O)
- MACHINE (O)
- PLATING (O)
- PAINT (O)
- SPECIAL (O)
- WHOLE JOB (O)

Columns:

- STATUS
- DUE
- SHIP
- PROCESS
- COMPANY
- WO #
- PART #
- DESCRIPTION
- QTY
- UPDATED BY
- PHOTO

## LOG Tab

Audit trail for completed process marks:

- Completion rows are created only when CALENDAR/PURCHASING has a green card and UPDATED BY is filled in
- UPDATED BY records who marked the process complete
- PM REVIEW, REVIEWED BY, OWNER / STATION, ISSUE NOTES, and RESOLUTION support end-of-day sign-off
- Date banner rows summarize completion and issue counts for the day

## SETTINGS Tab

Editable source list for employee dropdowns used by CALENDAR, PURCHASING, and LOG.

---

## Key Design Decisions

### Why `TwoCellAnchor` for images?
Standard floating images in Excel don't hide when rows are filtered. `TwoCellAnchor` with `editAs='twoCell'` makes images resize to zero height when a row is hidden, fixing the filter bug.

### Why normalize image anchors on every save?
Excel row sorting only keeps screenshots with the correct part when the image is saved as **move and size with cells**. Normalizing all two-cell anchors before save prevents the old screenshot scrambling issue from returning.

### Why date-based filtering (not column-order)?
Step order varies by job. A part might go WELD → MACHINE(O) → ASSM even though MACHINE(O) is after WELD in our column layout. Using the actual due dates means PROCESSES LEFT and the calendar filter are always accurate to what the engineer planned.

### Why delivery date as a column, not a calendar event?
Showing delivery dates in the calendar created noise — every part had an entry. Delivery date is more useful as a reference column on each process row so the PM can see the deadline alongside the specific step.

### Image preservation across script runs
`_UnclosableBytesIO` subclasses `io.BytesIO` and no-ops `close()`. This prevents openpyxl's `_data()` method from permanently closing the buffer, which would cause `ValueError: I/O operation on closed file` on save.
