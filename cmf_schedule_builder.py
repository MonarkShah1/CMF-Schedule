#!/usr/bin/env python3
"""
cmf_schedule_builder.py  —  Run daily (or any time)
Refreshes CALENDAR and TODAY (Kanban TV display) from MAIN.

Flow:  MAIN → CALENDAR (4-week date-bucket) → TODAY (Kanban by station)

Usage: python3 cmf_schedule_builder.py
"""

import os, io
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from datetime import datetime, timedelta
from collections import defaultdict

FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CMF WIP - Schedule.xlsx')

# ── Palette ───────────────────────────────────────────────────────────────────
C_NAVY     = "1F3864";  C_WHITE    = "FFFFFF"
C_STEEL_DK = "1F4E79";  C_LGRAY    = "F2F2F2"
C_DGRAY    = "404040";  C_YELLOW   = "FFF2CC"
C_RED_LT   = "FFD7D7";  C_GREEN_LT = "E2EFDA"

def _fill(h): return PatternFill(start_color=h, end_color=h, fill_type="solid")
def _font(bold=False, color=C_DGRAY, size=10): return Font(name="Calibri", bold=bold, color=color, size=size)
def _border(c="CCCCCC"): s=Side(style="thin",color=c); return Border(left=s,right=s,top=s,bottom=s)
def _align(h="center",v="center",wrap=False): return Alignment(horizontal=h,vertical=v,wrap_text=wrap)

# ── MAIN column positions (must match cmf_migrate_main.py) ───────────────────
COL_WO           = 1    # A
COL_PO           = 2    # B
COL_COMPANY      = 3    # C
COL_PARTNO       = 6    # F
COL_DESC         = 7    # G
COL_QTY          = 8    # H
COL_SCREENSHOT   = 9    # I
COL_CURRENT_STEP = 37   # AK  (shifted after SHIP TO VENDOR col inserted at AI/35)
COL_STATUS       = 38   # AL
HDR_ROW          = 2
DATA_START       = 3

# Maps CURRENT STEP value → its process column index.
# Any process col < curr_col is considered done and hidden from calendar/today.
# Maps process column index → short display name (for PROCESSES LEFT)
COL_TO_PROCESS = {
    12:"MATERIALS",  13:"ENGINEERING", 14:"LASER CUT",  15:"LASER(O)",
    16:"TUBE LSR(O)",17:"SAW",         18:"BANDSAW(O)", 19:"BEND",
    20:"CLEAN",      21:"CSK",         22:"DRILL",      23:"TAPPING",
    24:"GRIND",      25:"WELD",        26:"MACHINE(O)", 27:"PLATING(O)",
    28:"PAINT",      29:"PAINT(O)",    30:"SPECIAL",    31:"SPECIAL(O)",
    32:"W.JOB(O)",   33:"HARDWARE",    34:"ASSEMBLY",   35:"SHIP VNDR",
}

STEP_TO_COL = {
    "MATERIALS": 12,     "ENGINEERING": 13,
    "LASER CUT": 14,     "LASER CUT (O)": 15,  "TUBE LASER (O)": 16,
    "SAW": 17,           "BANDSAW (O)": 18,
    "BEND": 19,          "CLEAN": 20,
    "CSK": 21,           "DRILL": 22,           "TAPPING": 23,
    "GRIND": 24,         "WELD": 25,
    "MACHINE (O)": 26,   "PLATING (O)": 27,
    "PAINT": 28,         "PAINT (O)": 29,
    "SPECIAL": 30,       "SPECIAL (O)": 31,
    "WHOLE JOB (O)": 32, "WHOLE JOB": 32,
    "HARDWARE": 33,      "ASSEMBLY": 34,
    "SHIP TO VENDOR": 35,
    "SHIP": 36,          "DELIVERY DATE": 36,
    "COMPLETE": 99,      "DONE": 99,  "SHIPPED": 99,
}

IMG_W = 85   # copy same size from migration
IMG_H = 48

# Station definitions: (key, display_name, hdr_hex, row_hex, [main_col_indices_1based])
# Each station maps to one or more process-due-date columns.
# ORDER: in-house processes first, then outside (vendor) processes.
# In-house and outside variants of the same process type are separate stations
# so they never share a calendar column.
STATIONS = [
    # ── IN-HOUSE ──────────────────────────────────────────────────────────────
    # key            display name               hdr       row        cols
    ("MATERIALS",   "MATERIALS",            "1D4ED8","BAE6FD", [12]),        # L   Sky Blue
    ("ENGINEERING", "ENGINEERING",          "3730A3","C7D2FE", [13]),        # M   Indigo
    ("LASER_IN",    "LASER CUTTING",        "C2410C","FED7AA", [14]),        # N   Orange
    ("SAW_IN",      "SAW",                  "92400E","FDE68A", [17]),        # Q   Amber
    ("FORMING",     "FORMING / BENDING",    "A16207","FEF08A", [19]),        # S   Yellow
    ("CLEAN",       "CLEANING",             "065F46","A7F3D0", [20]),        # T   Mint Green
    ("ACCESSORIES", "CSK / DRILL / TAPPING","6B21A8","E9D5FF", [21,22,23]), # UVW Purple
    ("GRIND",       "GRINDING",             "334155","CBD5E1", [24]),        # X   Slate
    ("WELDING",     "WELDING",              "B91C1C","FECACA", [25]),        # Y   Red
    ("SPECIAL_IN",  "SPECIAL",              "374151","D1D5DB", [30]),        # AD  Gray
    ("HARDWARE",    "HARDWARE",             "0E7490","A5F3FC", [33]),        # AG  Cyan
    ("ASSEMBLY",    "ASSEMBLY",             "166534","D9F99D", [34]),        # AH  Lime
    ("SHIP_VENDOR", "SHIP TO VENDOR",       "0F766E","99F6E4", [35]),        # AI  Teal
    # ── OUTSIDE (VENDOR) ──────────────────────────────────────────────────────
    ("LASER_OUT",   "LASER CUT / TUBE (OUTSIDE)","EA580C","FFEDD5", [15,16]),# OP  Orange-lite
    ("SAW_OUT",     "BANDSAW (OUTSIDE)",    "B45309","FEF3C7", [18]),        # R   Amber-lite
    ("OUTSIDE",     "OUTSIDE MACHINE",      "5B21B6","DDD6FE", [26]),       # Z   Violet
    ("PAINT",       "PAINT / POWDER COAT / PLATE","15803D","BBF7D0",[27,28,29]),# AAABAC Green
    ("SPECIAL_OUT", "SPECIAL / WHOLE JOB (OUTSIDE)","4B5563","E5E7EB",[31,32]),# AEAF Gray-lite
    # ── REFERENCE — not rendered as a calendar station ────────────────────────
    ("SHIPPING",    "DELIVERY DATE",        "1F3864","E2EFDA", [36]),        # AJ  Navy
]

STATION_KEYS  = [s[0] for s in STATIONS]
STATION_NAME  = {s[0]: s[1] for s in STATIONS}
STATION_HDR_C = {s[0]: s[2] for s in STATIONS}
STATION_ROW_C = {s[0]: s[3] for s in STATIONS}


# ── Image helpers ─────────────────────────────────────────────────────────────
class _UnclosableBytesIO(io.BytesIO):
    def close(self): pass   # prevent openpyxl _data() from permanently closing


def _protect_image(img):
    ref = img.ref
    if isinstance(ref, io.BytesIO) and not isinstance(ref, _UnclosableBytesIO):
        try:
            ref.seek(0)
            img.ref = _UnclosableBytesIO(ref.read())
        except Exception:
            pass


def protect_workbook_images(wb):
    for ws in wb.worksheets:
        for img in ws._images:
            _protect_image(img)


def copy_image(raw_bytes, dest_ws, col_1idx, row_1idx):
    """
    Place a copy of raw_bytes as an image in dest_ws at (col_1idx, row_1idx).
    Accepts raw bytes (as stored in row_to_img) — never touches MAIN images.
    Uses TwoCellAnchor so image moves with its row.
    """
    from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
    try:
        buf = _UnclosableBytesIO(raw_bytes)
        img = XLImage(buf)
        c0, r0 = col_1idx - 1, row_1idx - 1
        anchor        = TwoCellAnchor()
        anchor.editAs = 'twoCell'
        anchor._from  = AnchorMarker(col=c0,   colOff=0, row=r0,   rowOff=0)
        anchor.to     = AnchorMarker(col=c0+1, colOff=0, row=r0+1, rowOff=0)
        img.anchor    = anchor
        dest_ws.add_image(img)
        return True
    except Exception:
        return False


# ── Auto-migration: insert SHIP TO VENDOR column if workbook predates it ──────
def ensure_ship_vendor_col(ws):
    """
    If MAIN was created before the SHIP TO VENDOR column was added, insert a
    blank column at position 35 (AI) so that DELIVERY DATE shifts to AJ and
    all subsequent column numbers stay consistent with the current constants.
    Safe to call every run — it's a no-op if the column already exists.
    """
    header_val = str(ws.cell(HDR_ROW, 35).value or "").strip().upper().replace("\n", " ")
    if "SHIP TO" in header_val:
        return  # already migrated

    print("  AUTO-MIGRATION: inserting SHIP TO VENDOR column at AI (col 35) ...")
    ws.insert_cols(35)

    hdr_cell = ws.cell(HDR_ROW, 35)
    hdr_cell.value     = "SHIP TO\nVENDOR"
    hdr_cell.font      = Font(name="Calibri", bold=True, color="FFFFFF", size=9)
    hdr_cell.fill      = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    hdr_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    hdr_cell.border    = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )
    ws.column_dimensions[get_column_letter(35)].width = 6
    print("  AUTO-MIGRATION: done — no further action needed.")


# ── Parse MAIN ────────────────────────────────────────────────────────────────
def parse_main(ws):
    """Returns (entries, row_to_image).
    entries:     list of dicts — one per process-due-date filled in MAIN.
    row_to_img:  {main_row_1indexed: raw_bytes}  — bytes copied out of MAIN
                 so the original image objects are never touched during
                 calendar/today builds or saves.  MAIN photos are safe.
    """
    entries    = []
    row_to_img = {}

    # Read image bytes once — store raw bytes, NOT the image object.
    # This completely decouples MAIN images from anything the script writes.
    for img in ws._images:
        try:
            excel_row = img.anchor._from.row + 1
            ref = img.ref
            if hasattr(ref, 'read'):
                ref.seek(0)
                raw = ref.read()
                ref.seek(0)   # reset so save can still read it
            else:
                raw = img._data()
            if raw:
                row_to_img[excel_row] = raw   # bytes only, no image reference
        except Exception:
            pass

    for row in range(DATA_START, ws.max_row + 1):
        wo      = ws.cell(row, COL_WO).value
        if not wo: continue

        part_no = ws.cell(row, COL_PARTNO).value
        desc    = ws.cell(row, COL_DESC).value
        qty     = ws.cell(row, COL_QTY).value
        company = ws.cell(row, COL_COMPANY).value
        po      = ws.cell(row, COL_PO).value
        status  = str(ws.cell(row, COL_STATUS).value or "").strip().upper()

        # Skip WO header rows and completed items
        has_part  = bool(part_no or desc)
        has_dates = any(ws.cell(row, c).value
                        for s in STATIONS for c in s[4])
        if not has_part and not has_dates: continue
        if status in ("COMPLETE","DONE","SHIPPED"): continue

        # Date-based filter: find the due date of the CURRENT STEP, then only
        # show processes whose due date is >= that date (still to come).
        curr_raw       = str(ws.cell(row, COL_CURRENT_STEP).value or "").strip().upper()
        curr_step_col  = STEP_TO_COL.get(curr_raw)
        curr_step_date = None
        if curr_step_col and curr_step_col < 99:
            v = ws.cell(row, curr_step_col).value
            curr_step_date = v if isinstance(v, datetime) else None
        if curr_raw in ("COMPLETE", "DONE", "SHIPPED"):
            continue

        # Delivery date (col AJ = 36, shifted from AI=35 after SHIP TO VENDOR inserted)
        delivery_date = ws.cell(row, 36).value
        if not isinstance(delivery_date, datetime):
            delivery_date = None

        # Compute "PROCESSES LEFT" sorted by actual due date (cols 12–35 inclusive)
        remaining = []
        for c in range(12, 36):
            v = ws.cell(row, c).value
            if v and isinstance(v, datetime) and c in COL_TO_PROCESS and (curr_step_date is None or v >= curr_step_date):
                remaining.append((v, COL_TO_PROCESS[c]))
        remaining.sort(key=lambda x: x[0])
        processes_left = " → ".join(name for _, name in remaining) or "—"

        for s_key, s_name, _, _, col_indices in STATIONS:
            if s_key == "SHIPPING":
                continue
            for col_idx in col_indices:
                due = ws.cell(row, col_idx).value
                if not due or not isinstance(due, datetime): continue
                if curr_step_date and due < curr_step_date:
                    continue    # already done

                # Step status label — 2 states only:
                #   ▶ READY  = operator can work on this right now
                #   ⏳ <step> = still waiting; shows what's blocking
                if curr_step_date is None:
                    step_status = ""               # not started — no label
                elif due == curr_step_date:
                    step_status = "▶ READY"        # this IS the current step
                else:
                    step_status = f"⏳ {curr_raw[:10]}"  # blocked by current step

                entries.append(dict(
                    main_row      = row,
                    date          = due,
                    s_key         = s_key,
                    station       = s_name,
                    wo            = wo,
                    po            = po,
                    company       = company or "",
                    part_no       = str(part_no or "").strip(),
                    desc          = str(desc    or "").strip(),
                    qty           = qty,
                    current_step  = curr_raw or "—",
                    processes_left= processes_left,
                    delivery_date = delivery_date,
                    step_status   = step_status,
                ))

    return entries, row_to_img


# ── CALENDAR ─────────────────────────────────────────────────────────────────
#   STATUS | DUE | COMPANY | WO# | QTY | PHOTO
# ── CALENDAR — Horizontal Kanban ─────────────────────────────────────────────
# Columns = stations, items stacked under each station sorted by date.
# Makes it easy for PM to see workstation loading and queue depth at a glance.
#
#  Row 1: Title
#  Row 2: Overdue status banner
#  Row 3: Station headers (colored, merged across sub-cols)
#  Row 4: Sub-headers: STATUS | DUE | COMPANY | WO # | QTY | PHOTO per station
#  Row 5+: Items sorted date-ascending (overdue at top in coral)

CAL_SUB    = [("STATUS", 8), ("DUE", 6), ("SHIP", 6), ("COMPANY", 11), ("WO #", 6), ("QTY", 7), ("PHOTO", 11)]
CAL_SUB_N  = len(CAL_SUB)   # 6 data cols
CAL_DIV_W  = 1               # divider col between stations
CAL_BLOCK  = CAL_SUB_N + CAL_DIV_W  # data cols + divider per station
CAL_ROW_H  = 52              # tall enough for photo thumbnails
CAL_IMG_W  = 68              # photo width in pixels
CAL_IMG_H  = 44              # photo height in pixels

CAL_STATION_COLOR = {s[0]: s[3] for s in STATIONS}


def build_calendar(wb, entries, row_to_img):
    if "CALENDAR" in wb.sheetnames: del wb["CALENDAR"]
    ws = wb.create_sheet("CALENDAR")
    today = datetime.now().date()

    # Active stations in defined display order (skip SHIPPING — delivery date is a column)
    active = [s for s in STATIONS
              if s[0] != "SHIPPING" and any(e["s_key"] == s[0] for e in entries)]

    if not active:
        ws["A1"] = "No scheduled work found."
        print("  CALENDAR: no entries")
        return

    # Per-station items sorted by date (overdue first, then ascending)
    station_items = {
        s[0]: sorted([e for e in entries if e["s_key"] == s[0]], key=lambda e: e["date"])
        for s in active
    }
    max_items  = max(len(v) for v in station_items.values())
    total_cols = len(active) * CAL_BLOCK
    last_col   = get_column_letter(total_cols)

    # Assign column start position for each station and set column widths
    station_start = {}
    col = 1
    for s in active:
        station_start[s[0]] = col
        for _, w in CAL_SUB:
            ws.column_dimensions[get_column_letter(col)].width = w
            col += 1
        ws.column_dimensions[get_column_letter(col)].width = CAL_DIV_W
        col += 1

    # ── Row 1: Title ──────────────────────────────────────────────────────────
    ws.row_dimensions[1].height = 30
    ws.merge_cells(f"A1:{last_col}1")
    c = ws["A1"]
    c.value     = f"CMF  PRODUCTION CALENDAR   —   {today.strftime('%B %d, %Y')}"
    c.font      = _font(bold=True, color=C_WHITE, size=14)
    c.fill      = _fill(C_NAVY); c.alignment = _align()

    # ── Row 2: Overdue status banner ─────────────────────────────────────────
    overdue_count = sum(1 for e in entries if e["date"].date() < today)
    ws.row_dimensions[2].height = 26
    ws.merge_cells(f"A2:{last_col}2")
    c = ws["A2"]
    if overdue_count:
        c.value = (f"  ⚠  {overdue_count} OVERDUE process step(s)  —  "
                   f"shown in red at the top of each station column")
        c.font  = _font(bold=True, color=C_WHITE, size=11)
        c.fill  = _fill("C00000")
    else:
        c.value = "  ✅  NO OVERDUE PROCESSES  —  all steps on track"
        c.font  = _font(bold=True, color=C_WHITE, size=11)
        c.fill  = _fill("375623")
    c.alignment = _align(h="left")

    # ── Row 3: Station headers ────────────────────────────────────────────────
    ws.row_dimensions[3].height = 36
    for s in active:
        sc       = station_start[s[0]]
        last_sub = get_column_letter(sc + CAL_SUB_N - 1)
        ws.merge_cells(f"{get_column_letter(sc)}3:{last_sub}3")
        c = ws.cell(3, sc)
        n = len(station_items[s[0]])
        od = sum(1 for e in station_items[s[0]] if e["date"].date() < today)
        c.value     = f"  {STATION_NAME[s[0]].upper()}  ({n})" + (f"  ⚠{od}" if od else "")
        c.font      = _font(bold=True, color=C_WHITE, size=12)
        c.fill      = _fill(STATION_HDR_C[s[0]]); c.alignment = _align(h="left")
        # Divider
        ws.cell(3, sc + CAL_SUB_N).fill = _fill("D0D0D0")

    # ── Row 4: Sub-headers ────────────────────────────────────────────────────
    ws.row_dimensions[4].height = 16
    for s in active:
        sc = station_start[s[0]]
        for i, (lbl, _) in enumerate(CAL_SUB):
            c = ws.cell(4, sc + i, lbl)
            c.font      = _font(bold=True, size=9, color=C_DGRAY)
            c.fill      = _fill(STATION_ROW_C[s[0]])
            c.border    = _border("AAAAAA"); c.alignment = _align()
        ws.cell(4, sc + CAL_SUB_N).fill = _fill("D0D0D0")

    # ── Rows 5+: Item cards ───────────────────────────────────────────────────
    for item_idx in range(max_items):
        row = 5 + item_idx
        ws.row_dimensions[row].height = CAL_ROW_H

        for s in active:
            sc    = station_start[s[0]]
            items = station_items[s[0]]

            if item_idx < len(items):
                e       = items[item_idx]
                is_od   = e["date"].date() < today
                days    = (e["date"].date() - today).days
                is_soon = not is_od and days <= 7
                bg      = "FCA5A5" if is_od else (C_YELLOW if is_soon else STATION_ROW_C[s[0]])
                due_fg  = "CC0000" if is_od else C_DGRAY
                status  = e.get("step_status", "")
                due_str = f"{e['date'].month}/{e['date'].day}"

                # STATUS | DUE | SHIP | COMPANY | WO# | QTY | PHOTO
                dd = e.get("delivery_date")
                ship_str = f"{dd.month}/{dd.day}" if dd else "—"
                vals = [status, due_str, ship_str, e["company"], e["wo"], e["qty"], None]
                for i, val in enumerate(vals):
                    c = ws.cell(row, sc + i, val)
                    c.border    = _border("CCCCCC")
                    c.alignment = _align(h="left" if i == 3 else "center")

                    if i == 0:      # STATUS label — coloured independently
                        if status == "▶ READY":
                            c.fill = _fill("DBEAFE")
                            c.font = _font(bold=True, size=9, color="1E40AF")
                        elif status.startswith("⏳"):
                            c.fill = _fill("F3F4F6")
                            c.font = _font(size=9, color="6B7280")
                        else:
                            c.fill = _fill(bg); c.font = _font(size=9)
                    elif i == 1:    # DUE date
                        c.fill = _fill(bg)
                        c.font = _font(bold=True, size=10, color=due_fg)
                    elif i == 2:    # SHIP date
                        c.fill = _fill(bg)
                        c.font = _font(size=9, color="6B7280")
                    elif i == 3:    # COMPANY
                        c.fill = _fill(bg)
                        c.font = _font(bold=True, size=10)
                    else:           # WO#, QTY, and PHOTO placeholder
                        c.fill = _fill(bg); c.font = _font(size=10)

                # Screenshot in PHOTO column (sub-col index 6 = sc+6)
                if e["main_row"] in row_to_img:
                    copy_image(row_to_img[e["main_row"]], ws, sc + 6, row)
            else:
                # Empty slot
                for i in range(CAL_SUB_N):
                    c = ws.cell(row, sc + i)
                    c.fill   = _fill(STATION_ROW_C[s[0]])
                    c.border = _border("DDDDDD")

            ws.cell(row, sc + CAL_SUB_N).fill = _fill("D0D0D0")

    ws.freeze_panes = "A5"
    print(f"  CALENDAR: {len(entries)} entries across {len(active)} stations  ({overdue_count} overdue)")


# ── TODAY — Kanban board ──────────────────────────────────────────────────────
# Each active station = a column group: WO | COMPANY | PART# | QTY | DUE | PHOTO
BLOCK_W      = [8, 16, 14, 7, 10, 13]   # col widths per sub-column
BLOCK_LABELS = ["WO #","COMPANY","PART #","QTY","DUE","PHOTO"]
BLOCK_NC     = len(BLOCK_W)     # 6 sub-columns
DIVIDER_W    = 2
ITEM_H       = 90    # row height in points — tall for in-cell photo
KANBAN_IMG_W = 90
KANBAN_IMG_H = 70


def build_today(wb, entries, row_to_img):
    if "TODAY" in wb.sheetnames: del wb["TODAY"]
    ws = wb.create_sheet("TODAY")
    today = datetime.now().date()

    due_entries = [e for e in entries if e["date"].date() <= today]
    if not due_entries:
        ws.row_dimensions[1].height = 60
        ws.merge_cells("A1:H1")
        c = ws["A1"]
        c.value     = f"  ✅  PRODUCTION SCHEDULE  —  {datetime.now().strftime('%A %B %d, %Y').upper()}  —  SHOP IS CLEAR"
        c.font      = _font(bold=True, size=18, color="375623")
        c.fill      = _fill("E2EFDA"); c.alignment = _align(h="left")
        print("  TODAY: clear"); return

    # Active stations in display order
    active = [s for s in STATIONS if any(e["s_key"] == s[0] for e in due_entries)]

    # Assign column starts
    station_start = {}
    col = 1
    for s in active:
        station_start[s[0]] = col
        for offset, w in enumerate(BLOCK_W):
            ws.column_dimensions[get_column_letter(col + offset)].width = w
        ws.column_dimensions[get_column_letter(col + BLOCK_NC)].width = DIVIDER_W
        col += BLOCK_NC + 1
    total_cols = col - 1

    station_items = {}
    for s in active:
        items = [e for e in due_entries if e["s_key"] == s[0]]
        items.sort(key=lambda e: (0 if e["date"].date() < today else 1, e["date"]))
        station_items[s[0]] = items

    max_items = max(len(v) for v in station_items.values())

    # ── Row 1: title ──────────────────────────────────────────────────────────
    ws.row_dimensions[1].height = 50
    ws.merge_cells(f"A1:{get_column_letter(total_cols)}1")
    c = ws["A1"]
    od = sum(1 for e in due_entries if e["date"].date() < today)
    c.value     = (f"  CMF  PRODUCTION SCHEDULE   —   "
                   f"{datetime.now().strftime('%A, %B %d, %Y').upper()}"
                   f"   |   {od} OVERDUE  •  {len(due_entries)-od} DUE TODAY")
    c.font      = _font(bold=True, color=C_WHITE, size=18)
    c.fill      = _fill(C_NAVY); c.alignment = _align(h="left")

    # ── Row 2: station headers ────────────────────────────────────────────────
    ws.row_dimensions[2].height = 44
    for s in active:
        sc = station_start[s[0]]
        ws.merge_cells(f"{get_column_letter(sc)}2:{get_column_letter(sc+BLOCK_NC-1)}2")
        c = ws.cell(2, sc)
        n  = len(station_items[s[0]])
        od_n = sum(1 for e in station_items[s[0]] if e["date"].date() < today)
        c.value     = (f"  {STATION_NAME[s[0]].upper()}  ({n})"
                       + (f"  ⚠ {od_n} OVERDUE" if od_n else ""))
        c.font      = _font(bold=True, color=C_WHITE, size=15)
        c.fill      = _fill(STATION_HDR_C[s[0]]); c.alignment = _align(h="left")
        ws.cell(2, sc+BLOCK_NC).fill = _fill("D0D0D0")

    # ── Row 3: sub-headers ────────────────────────────────────────────────────
    ws.row_dimensions[3].height = 18
    for s in active:
        sc = station_start[s[0]]
        for offset, lbl in enumerate(BLOCK_LABELS):
            c = ws.cell(3, sc+offset, lbl)
            c.font = _font(bold=True, size=9, color=C_DGRAY)
            c.fill = _fill(STATION_ROW_C[s[0]]); c.border = _border("AAAAAA")
            c.alignment = _align()
        ws.cell(3, sc+BLOCK_NC).fill = _fill("D0D0D0")

    # ── Rows 4+: item cards ───────────────────────────────────────────────────
    for idx in range(max_items):
        row = 4 + idx
        ws.row_dimensions[row].height = ITEM_H

        for s in active:
            sc    = station_start[s[0]]
            items = station_items[s[0]]

            if idx < len(items):
                e     = items[idx]
                is_od = e["date"].date() < today
                bg    = "FCA5A5" if is_od else C_YELLOW
                due_fg= "CC0000" if is_od else "806000"

                vals = [e["wo"], e["company"],
                        e["part_no"] or e["desc"], e["qty"],
                        e["date"].strftime("%b %d").upper() if isinstance(e["date"], datetime) else "",
                        None]
                for offset, val in enumerate(vals):
                    c = ws.cell(row, sc+offset, val)
                    c.fill   = _fill(bg); c.border = _border("CCCCCC")
                    c.alignment = _align(h="left" if offset in (1,2) else "center")
                    if offset == 0:   c.font = _font(bold=True, size=13, color=C_DGRAY)
                    elif offset == 1: c.font = _font(bold=True, size=12, color=C_DGRAY)
                    elif offset == 4: c.font = _font(bold=True, size=12, color=due_fg)
                    else:             c.font = _font(size=12, color=C_DGRAY)

                # Photo — TwoCellAnchor so it filters/hides with the row
                if e["main_row"] in row_to_img:
                    copy_image(row_to_img[e["main_row"]], ws, sc + 5, row)
            else:
                bg = STATION_ROW_C[s[0]]
                for offset in range(BLOCK_NC):
                    c = ws.cell(row, sc+offset)
                    c.fill = _fill(bg); c.border = _border("DDDDDD")

            ws.cell(row, sc+BLOCK_NC).fill = _fill("D0D0D0")

    ws.freeze_panes = "A4"
    img_count = sum(1 for e in due_entries if e["main_row"] in row_to_img)
    print(f"  TODAY: {len(due_entries)} items across {len(active)} stations ({img_count} with photos)")


# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print(f"Opening  {FILE}")
    if not os.path.exists(FILE):
        print("ERROR: Run cmf_migrate_main.py first."); raise SystemExit(1)

    wb = openpyxl.load_workbook(FILE)
    protect_workbook_images(wb)

    if "MAIN" not in wb.sheetnames:
        print("ERROR: No MAIN sheet. Run cmf_migrate_main.py first."); raise SystemExit(1)

    print("Checking MAIN column layout ...")
    ensure_ship_vendor_col(wb["MAIN"])

    print("Parsing MAIN ...")
    entries, row_to_img = parse_main(wb["MAIN"])
    print(f"  {len(entries)} process-due entries  |  {len(row_to_img)} screenshots")

    print("Building CALENDAR ...")
    build_calendar(wb, entries, row_to_img)

    print("Building TODAY (Kanban) ...")
    build_today(wb, entries, row_to_img)

    for name in ["TODAY","CALENDAR","MAIN"]:
        if name in wb.sheetnames:
            wb.move_sheet(name, offset=-len(wb.sheetnames))

    print(f"Saving   {FILE}")
    wb.save(FILE)
    print("Done ✓\n")
    print("  MAIN     → engineer fills routing (blue cols L–AI) + screenshots (col I)")
    print("  CALENDAR → PM planning view, 4-week rolling, photos inline")
    print("  TODAY    → Kanban TV display, stations as columns, photos in cards")
