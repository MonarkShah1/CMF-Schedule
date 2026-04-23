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
COL_CURRENT_STEP = 36   # AJ
COL_STATUS       = 37   # AK
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
    32:"W.JOB(O)",   33:"HARDWARE",    34:"ASSEMBLY",
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
    "SHIP": 35,          "DELIVERY DATE": 35,
    "COMPLETE": 99,      "DONE": 99,  "SHIPPED": 99,
}

IMG_W = 85   # copy same size from migration
IMG_H = 48

# Station definitions: (key, display_name, hdr_hex, row_hex, [main_col_indices_1based])
# Each station maps to one or more process-due-date columns
# Bold, saturated row colors — each station immediately distinct at a glance.
# Row colors use medium saturation (not pale pastels, not too dark for text).
# Headers use a deeper shade of the same hue.
STATIONS = [
    # key          display name              hdr       row(bold)  cols
    ("MATERIALS",   "MATERIALS",            "1D4ED8","BAE6FD", [12]),        # L   Sky Blue
    ("ENGINEERING", "ENGINEERING",          "3730A3","C7D2FE", [13]),        # M   Indigo
    ("LASER",       "LASER CUTTING",        "C2410C","FED7AA", [14,15,16]),  # NOP Orange
    ("SAW",         "SAW / BANDSAW",        "92400E","FDE68A", [17,18]),     # QR  Amber
    ("FORMING",     "FORMING / BENDING",    "A16207","FEF08A", [19]),        # S   Yellow
    ("CLEAN",       "CLEANING",             "065F46","A7F3D0", [20]),        # T   Mint Green
    ("ACCESSORIES", "CSK / DRILL / TAPPING","6B21A8","E9D5FF", [21,22,23]), # UVW Purple
    ("GRIND",       "GRINDING",             "334155","CBD5E1", [24]),        # X   Slate
    ("WELDING",     "WELDING",              "B91C1C","FECACA", [25]),        # Y   Red
    ("OUTSIDE",     "OUTSIDE PROCESS",      "5B21B6","DDD6FE", [26,27]),    # ZAA Violet
    ("PAINT",       "PAINT / POWDER COAT",  "15803D","BBF7D0", [28,29]),    # ABAC Green
    ("SPECIAL",     "SPECIAL / WHOLE JOB",  "374151","D1D5DB", [30,31,32]), # ADAEAF Gray
    ("HARDWARE",    "HARDWARE",             "0E7490","A5F3FC", [33]),        # AG  Cyan
    ("ASSEMBLY",    "ASSEMBLY",             "166534","D9F99D", [34]),        # AH  Lime
    ("SHIPPING",    "DELIVERY DATE",        "1F3864","E2EFDA", [35]),        # AI  Navy
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


def copy_image(src_img, dest_ws, col_1idx, row_1idx):
    """
    Copy image using TwoCellAnchor(editAs='twoCell') so it moves AND collapses
    with its row when filtering — fixes the floating-image filter bug.
    col_1idx, row_1idx are 1-based Excel coordinates.
    """
    from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
    try:
        ref = src_img.ref
        if hasattr(ref, 'read'):
            ref.seek(0); raw = ref.read()
        else:
            raw = src_img._data()
        buf = _UnclosableBytesIO(raw)
        img = XLImage(buf)
        c0, r0 = col_1idx - 1, row_1idx - 1
        anchor          = TwoCellAnchor()
        anchor.editAs   = 'twoCell'           # image hides when row is filtered/hidden
        anchor._from    = AnchorMarker(col=c0,   colOff=0, row=r0,   rowOff=0)
        anchor.to       = AnchorMarker(col=c0+1, colOff=0, row=r0+1, rowOff=0)
        img.anchor = anchor
        dest_ws.add_image(img)
        return True
    except Exception:
        return False


# ── Parse MAIN ────────────────────────────────────────────────────────────────
def parse_main(ws):
    """Returns (entries, row_to_image).
    entries: list of dicts — one per process-due-date filled in MAIN.
    row_to_image: {main_row_1indexed: Image}
    """
    entries    = []
    row_to_img = {}

    # Build image map
    for img in ws._images:
        try:
            row_to_img[img.anchor._from.row + 1] = img
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

        # Delivery date (col AI = 35) — shown as a column, not a calendar event
        delivery_date = ws.cell(row, 35).value
        if not isinstance(delivery_date, datetime):
            delivery_date = None

        # Compute "PROCESSES LEFT": remaining steps sorted by their actual due date
        remaining = []
        for c in range(12, 35):
            v = ws.cell(row, c).value
            if v and isinstance(v, datetime) and c in COL_TO_PROCESS and (curr_step_date is None or v >= curr_step_date):
                remaining.append((v, COL_TO_PROCESS[c]))
        remaining.sort(key=lambda x: x[0])   # sort by date, not column order
        processes_left = " → ".join(name for _, name in remaining) or "—"

        for s_key, s_name, _, _, col_indices in STATIONS:
            if s_key == "SHIPPING":
                continue    # delivery date is now a column, not a calendar event
            for col_idx in col_indices:
                due = ws.cell(row, col_idx).value
                if not due or not isinstance(due, datetime): continue
                if curr_step_date and due < curr_step_date:
                    continue    # already done
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
                ))

    return entries, row_to_img


# ── CALENDAR ─────────────────────────────────────────────────────────────────
#   STATION | COMPANY | WO# | PART# | QTY | CURRENT STEP | PROCESSES LEFT | DELIVERY DATE | PHOTO
CAL_WIDTHS = [14,       20,     8,    18,    6,    16,             28,              11,              12]
CAL_LABELS = ["STATION","COMPANY","WO #","PART #","QTY","CURRENT STEP","PROCESSES LEFT","DELIVERY DATE","PHOTO"]
CAL_NCOLS  = len(CAL_WIDTHS)
CAL_LAST   = get_column_letter(CAL_NCOLS)
CAL_IMG_W  = 60;  CAL_IMG_H = 38
CAL_ROW_H_IMG = 42;  CAL_ROW_H = 18

CAL_STATION_COLOR = {s[0]: s[3] for s in STATIONS}


def build_calendar(wb, entries, row_to_img):
    if "CALENDAR" in wb.sheetnames: del wb["CALENDAR"]
    ws = wb.create_sheet("CALENDAR")
    today = datetime.now().date()

    for col, w in enumerate(CAL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(col)].width = w

    # Title
    ws.row_dimensions[1].height = 30
    ws.merge_cells(f"A1:{CAL_LAST}1")
    c = ws["A1"]
    c.value = f"CMF  PRODUCTION CALENDAR   —   {today.strftime('%B %d, %Y')}"
    c.font  = _font(bold=True, color=C_WHITE, size=14)
    c.fill  = _fill(C_NAVY); c.alignment = _align()

    # Column headers
    ws.row_dimensions[2].height = 18
    for col, lbl in enumerate(CAL_LABELS, 1):
        c = ws.cell(2, col, lbl)
        c.font = _font(bold=True, size=10, color=C_WHITE)
        c.fill = _fill(C_STEEL_DK); c.border = _border(); c.alignment = _align()

    cur = 3

    def banner(row, text, bg, fg, h=22):
        ws.row_dimensions[row].height = h
        ws.merge_cells(f"A{row}:{CAL_LAST}{row}")
        c = ws.cell(row, 1, text)
        c.font = _font(bold=True, color=fg, size=12)
        c.fill = _fill(bg); c.alignment = _align(h="left")

    def item(row, entry, od):
        has_img = entry["main_row"] in row_to_img
        ws.row_dimensions[row].height = CAL_ROW_H_IMG if has_img else CAL_ROW_H
        bg  = "FCA5A5" if od else CAL_STATION_COLOR.get(entry["s_key"], C_LGRAY)
        dv  = entry.get("delivery_date")
        dv_str = f"{dv.month}/{dv.day}/{str(dv.year)[2:]}" if isinstance(dv, datetime) else (str(dv)[:8] if dv else "—")
        # Cols: STATION | COMPANY | WO# | PART# | QTY | CURRENT STEP | PROCESSES LEFT | DELIVERY DATE | PHOTO
        vals = [
            STATION_NAME.get(entry["s_key"], entry["s_key"]),  # 1 station
            entry["company"],                                   # 2 company
            entry["wo"],                                        # 3 WO#
            entry["part_no"],                                   # 4 Part#
            entry["qty"],                                       # 5 QTY
            entry.get("current_step", "—"),                    # 6 current step
            entry.get("processes_left", "—"),                  # 7 processes left
            dv_str,                                             # 8 delivery date
            None,                                               # 9 photo
        ]
        for col, val in enumerate(vals, 1):
            c = ws.cell(row, col, val)
            c.fill = _fill(bg); c.border = _border("BBBBBB")
            c.font = _font(size=10, bold=(col == 8 and od))  # bold delivery date if overdue
            if col == 8 and od:  # overdue delivery date → red text
                c.font = _font(size=10, bold=True, color="CC0000")
            c.alignment = _align(h="left" if col in (1, 2, 7) else "center")
        if has_img:
            copy_image(row_to_img[entry["main_row"]], ws, 9, row)

    # ── Overdue section — always visible so PM can see status at a glance ──────
    overdue = sorted([e for e in entries if e["date"].date() < today], key=lambda e: e["date"])
    if overdue:
        banner(cur, f"  ⚠  OVERDUE  —  {len(overdue)} process step(s) past due  —  check what is holding these up", "C00000", C_WHITE, 30)
        cur += 1
        # Column sub-header for overdue section
        ws.row_dimensions[cur].height = 16
        for col, lbl in enumerate(CAL_LABELS, 1):
            c = ws.cell(cur, col, lbl)
            c.font = _font(bold=True, size=9, color="CC0000")
            c.fill = _fill("FFE8E8"); c.border = _border("EE9999"); c.alignment = _align()
        cur += 1
        for e in overdue:
            item(cur, e, True); cur += 1
    else:
        # Always show the overdue banner even when clear
        banner(cur, "  ✅  NO OVERDUE PROCESSES  —  all steps on track", "375623", C_WHITE, 26)
    cur += 1

    by_date = defaultdict(list)
    for e in entries:
        if e["date"].date() >= today:
            by_date[e["date"].date()].append(e)

    shown, d = 0, today
    while shown < 28:   # 4 full weeks including weekends
        shown += 1
        is_weekend = d.weekday() >= 5   # Sat=5, Sun=6
        day_entries = sorted(by_date.get(d, []),
                             key=lambda e: STATION_KEYS.index(e["s_key"]) if e["s_key"] in STATION_KEYS else 99)
        if not day_entries and d != today:
            d += timedelta(days=1); continue
        is_today   = (d == today)
        # Weekends get a warm gray banner so they're visually distinct from weekdays
        bg  = C_STEEL_DK if is_today else ("78716C" if is_weekend else "6D6D6D")
        txt = f"  {'TODAY  ·  ' if is_today else ''}{d.strftime('%A  %B %d').upper()}"
        if is_weekend and not is_today:
            txt += "  (WEEKEND)"
        banner(cur, txt, bg, C_WHITE, 28 if is_today else 22); cur += 1
        if day_entries:
            for e in day_entries: item(cur, e, False); cur += 1
        else:
            ws.row_dimensions[cur].height = 16
            ws.merge_cells(f"A{cur}:{CAL_LAST}{cur}")
            c = ws.cell(cur, 1, "     — nothing scheduled today —")
            c.font = _font(size=10, color="888888")
            c.fill = _fill("F5F5F5"); c.alignment = _align(h="left")
            cur += 1
            cur += 1
        d += timedelta(days=1)

    ws.freeze_panes = "A3"
    print(f"  CALENDAR: {len(entries)} entries  ({len(overdue)} overdue)")


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
