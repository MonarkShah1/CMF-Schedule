#!/usr/bin/env python3
"""
cmf_schedule_builder.py  —  Run daily (or any time)
Refreshes CALENDAR and PURCHASING from MAIN.

Flow:  MAIN → CALENDAR (production) + PURCHASING (outside/vendor work)

Usage: python3 cmf_schedule_builder.py
"""

import os, io, copy, hashlib
from collections import OrderedDict
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor
from datetime import datetime, timedelta
from collections import defaultdict

FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CMF WIP - Schedule.xlsx')

# ── Palette ───────────────────────────────────────────────────────────────────
C_NAVY     = "1F3864";  C_WHITE    = "FFFFFF"; C_GOLD = "FFD966"
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
COL_THICKNESS    = 11   # K
COL_CURRENT_STEP = 37   # AK  (shifted after SHIP TO VENDOR col inserted at AI/35)
COL_STATUS       = 38   # AL
HDR_ROW          = 2
DATA_START       = 3
MAIN_LAST_COL    = 39  # AM
BLUE_FIRST       = 12  # L
BLUE_LAST        = 36  # AJ

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

STEP_LIST = (
    "MATERIALS,ENGINEERING,LASER CUT,LASER CUT (O),TUBE LASER (O),"
    "SAW,BANDSAW (O),BEND,CLEAN,CSK,DRILL,TAPPING,GRIND,WELD,"
    "MACHINE (O),PLATING (O),PAINT,PAINT (O),"
    "SPECIAL,SPECIAL (O),WHOLE JOB (O),HARDWARE,ASSEMBLY,SHIP TO VENDOR,SHIP,RECEIVING,"
    "COMPLETE,ON HOLD"
)

# Maps process column → canonical CURRENT STEP name used in the dropdown.
# Used by apply_completions() to write the correct step name after advancing.
_COL_TO_DROPDOWN_STEP = {
    12: "MATERIALS",     13: "ENGINEERING",
    14: "LASER CUT",     15: "LASER CUT (O)",  16: "TUBE LASER (O)",
    17: "SAW",           18: "BANDSAW (O)",
    19: "BEND",          20: "CLEAN",
    21: "CSK",           22: "DRILL",           23: "TAPPING",
    24: "GRIND",         25: "WELD",
    26: "MACHINE (O)",   27: "PLATING (O)",
    28: "PAINT",         29: "PAINT (O)",
    30: "SPECIAL",       31: "SPECIAL (O)",
    32: "WHOLE JOB (O)", 33: "HARDWARE",
    34: "ASSEMBLY",      35: "SHIP TO VENDOR",
    36: "SHIP",
}

IMG_W = 85   # copy same size from migration
IMG_H = 48
IMG_PAD_X = 2
IMG_PAD_Y = 2
PART_ROW_H = 52
HEADER_ROW_H = 24
MAIN_COLS = [
    ("A",  "WO #",              8, False),
    ("B",  "PO #",             17, False),
    ("C",  "COMPANY",          20, False),
    ("D",  "JOB NAME",         28, False),
    ("E",  "CUST SHIP",        12, False),
    ("F",  "PART #",           16, False),
    ("G",  "DESCRIPTION",      26, False),
    ("H",  "QTY",               7, False),
    ("I",  "SCREENSHOT",       13, False),
    ("J",  "MATERIAL",         10, False),
    ("K",  "THICKNESS",         9, False),
    ("L",  "MATL",              6, True),
    ("M",  "ENG",               6, True),
    ("N",  "LASER",             6, True),
    ("O",  "LASER\n(O)",        6, True),
    ("P",  "TUBE\nLSR(O)",      6, True),
    ("Q",  "SAW",               6, True),
    ("R",  "BAND\n(O)",         6, True),
    ("S",  "BEND",              6, True),
    ("T",  "CLEAN",             6, True),
    ("U",  "CSK",               6, True),
    ("V",  "DRILL",             6, True),
    ("W",  "TAP",               6, True),
    ("X",  "GRIND",             6, True),
    ("Y",  "WELD",              6, True),
    ("Z",  "MACH\n(O)",         6, True),
    ("AA", "PLAT\n(O)",         6, True),
    ("AB", "PAINT",             6, True),
    ("AC", "PAINT\n(O)",        6, True),
    ("AD", "SPEC",              6, True),
    ("AE", "SPEC\n(O)",         6, True),
    ("AF", "W.JOB\n(O)",        6, True),
    ("AG", "HW",                6, True),
    ("AH", "ASSM",              6, True),
    ("AI", "SHIP TO\nVENDOR",   8, True),
    ("AJ", "DELIVERY\nDATE",    8, True),
    ("AK", "CURRENT\nSTEP",    18, False),
    ("AL", "STATUS",           12, False),
    ("AM", "NOTES",            35, False),
]

# Station definitions: (key, display_name, hdr_hex, row_hex, [main_col_indices_1based])
# Each horizontal section maps to exactly one process column so the calendar can
# show a separate block for each individual step.
STATIONS = [
    # ── IN-HOUSE ──────────────────────────────────────────────────────────────
    # key            display name               hdr       row        cols
    ("MATERIALS",     "MATERIALS",        "1D4ED8","BAE6FD", [12]),
    ("ENGINEERING",   "ENGINEERING",      "3730A3","C7D2FE", [13]),
    ("LASER_CUT",     "LASER CUT",        "C2410C","FED7AA", [14]),
    ("SAW",           "SAW",              "92400E","FDE68A", [17]),
    ("BEND",          "BEND",             "A16207","FEF08A", [19]),
    ("CLEAN",         "CLEAN",            "065F46","A7F3D0", [20]),
    ("CSK",           "CSK",              "6B21A8","E9D5FF", [21]),
    ("DRILL",         "DRILL",            "7C3AED","E9D5FF", [22]),
    ("TAPPING",       "TAPPING",          "8B5CF6","E9D5FF", [23]),
    ("GRIND",         "GRIND",            "334155","CBD5E1", [24]),
    ("WELD",          "WELD",             "B91C1C","FECACA", [25]),
    ("SPECIAL",       "SPECIAL",          "374151","D1D5DB", [30]),
    ("HARDWARE",      "HARDWARE",         "0E7490","A5F3FC", [33]),
    ("ASSEMBLY",      "ASSEMBLY",         "166534","D9F99D", [34]),
    ("SHIP_VENDOR",   "SHIP TO VENDOR",   "0F766E","99F6E4", [35]),
    # ── OUTSIDE (VENDOR) ──────────────────────────────────────────────────────
    ("LASER_CUT_O",   "LASER CUT (O)",    "EA580C","FFEDD5", [15]),
    ("TUBE_LASER_O",  "TUBE LASER (O)",   "F97316","FFEDD5", [16]),
    ("BANDSAW_O",     "BANDSAW (O)",      "B45309","FEF3C7", [18]),
    ("MACHINE_O",     "MACHINE (O)",      "5B21B6","DDD6FE", [26]),
    ("PLATING_O",     "PLATING (O)",      "15803D","BBF7D0", [27]),
    ("PAINT",         "PAINT",            "16A34A","BBF7D0", [28]),
    ("PAINT_O",       "PAINT (O)",        "22C55E","DCFCE7", [29]),
    ("SPECIAL_O",     "SPECIAL (O)",      "4B5563","E5E7EB", [31]),
    ("WHOLE_JOB_O",   "WHOLE JOB (O)",    "6B7280","E5E7EB", [32]),
    # ── REFERENCE — not rendered as a calendar station ────────────────────────
    ("SHIPPING",      "DELIVERY DATE",     "1F3864","E2EFDA", [36]),
]

STATION_KEYS  = [s[0] for s in STATIONS]
STATION_NAME  = {s[0]: s[1] for s in STATIONS}
STATION_HDR_C = {s[0]: s[2] for s in STATIONS}
STATION_ROW_C = {s[0]: s[3] for s in STATIONS}
PURCHASING_STATION_KEYS = {
    "SHIP_VENDOR", "LASER_CUT_O", "TUBE_LASER_O", "BANDSAW_O",
    "MACHINE_O", "PLATING_O", "PAINT_O", "SPECIAL_O", "WHOLE_JOB_O",
}
CALENDAR_STATION_KEYS = {
    s[0] for s in STATIONS
    if s[0] not in PURCHASING_STATION_KEYS and s[0] != "SHIPPING"
}

SETTINGS_SHEET = "SETTINGS"
SETTINGS_NAME_FIRST_ROW = 2
SETTINGS_NAME_LAST_ROW = 101
SETTINGS_NAME_RANGE = f"'{SETTINGS_SHEET}'!$A${SETTINGS_NAME_FIRST_ROW}:$A${SETTINGS_NAME_LAST_ROW}"


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


# ── SETTINGS sheet ───────────────────────────────────────────────────────────
def ensure_settings_sheet(wb):
    """Create/repair the editable employee-name source for dropdowns."""
    if SETTINGS_SHEET in wb.sheetnames:
        ws = wb[SETTINGS_SHEET]
    else:
        ws = wb.create_sheet(SETTINGS_SHEET)

    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 54

    c = ws["A1"]
    c.value = "EMPLOYEE NAMES"
    c.font = _font(bold=True, color=C_WHITE, size=11)
    c.fill = _fill(C_NAVY)
    c.alignment = _align(h="left")

    c = ws["B1"]
    c.value = "Add one name per row below. CALENDAR, PURCHASING, and LOG dropdowns use this list."
    c.font = _font(bold=True, color=C_DGRAY, size=9)
    c.fill = _fill(C_YELLOW)
    c.alignment = _align(h="left", wrap=True)

    for row in range(SETTINGS_NAME_FIRST_ROW, SETTINGS_NAME_LAST_ROW + 1):
        ws.cell(row, 1).border = _border("DDDDDD")

    ws.freeze_panes = "A2"
    return ws


def _employee_name_validation():
    return DataValidation(
        type="list",
        formula1=f"={SETTINGS_NAME_RANGE}",
        allow_blank=True,
        showErrorMessage=False,
    )


def _set_two_cell_anchor(img, col_1idx, row_1idx, width_px, height_px, pad_x=IMG_PAD_X, pad_y=IMG_PAD_Y):
    from openpyxl.drawing.spreadsheet_drawing import AnchorMarker
    c0, r0 = col_1idx - 1, row_1idx - 1
    anchor = TwoCellAnchor(editAs='twoCell')
    anchor._from = AnchorMarker(
        col=c0,
        colOff=pixels_to_EMU(pad_x),
        row=r0,
        rowOff=pixels_to_EMU(pad_y),
    )
    anchor.to = AnchorMarker(
        col=c0,
        colOff=pixels_to_EMU(pad_x + width_px),
        row=r0,
        rowOff=pixels_to_EMU(pad_y + height_px),
    )
    img.anchor = anchor


def normalize_image_anchors(wb):
    """Force all existing images to stay inside their cell and sort with rows."""
    normalized = 0
    for ws in wb.worksheets:
        for img in ws._images:
            anchor = getattr(img, "anchor", None)
            if not isinstance(anchor, TwoCellAnchor):
                continue
            row_1idx = anchor._from.row + 1
            col_1idx = anchor._from.col + 1
            if ws.title == "MAIN":
                _set_two_cell_anchor(img, col_1idx, row_1idx, IMG_W, IMG_H)
                normalized += 1
            elif anchor.editAs != "twoCell":
                anchor.editAs = "twoCell"
                normalized += 1
    if normalized:
        print(f"  IMAGE-ANCHOR: normalized {normalized} image(s) for Excel sorting")


def copy_image(raw_bytes, dest_ws, col_1idx, row_1idx):
    """
    Place a copy of raw_bytes as an image in dest_ws at (col_1idx, row_1idx).
    Accepts raw bytes (as stored in row_to_img) — never touches MAIN images.
    Uses TwoCellAnchor so image moves with its row.
    """
    try:
        buf = _UnclosableBytesIO(raw_bytes)
        img = XLImage(buf)
        _set_two_cell_anchor(img, col_1idx, row_1idx, CAL_IMG_W, CAL_IMG_H)
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


def _is_blank(val):
    return val is None or (isinstance(val, str) and val.strip() == "")


def _main_row_kind(ws, row):
    wo = ws.cell(row, COL_WO).value
    part_area_has_data = any(not _is_blank(ws.cell(row, c).value) for c in range(6, MAIN_LAST_COL + 1))
    meta_area_has_data = any(not _is_blank(ws.cell(row, c).value) for c in range(1, 6))
    if not part_area_has_data and meta_area_has_data and not _is_blank(wo):
        return "header"
    if part_area_has_data:
        return "part"
    return "blank"


def repair_main_sheet(ws, title_text="CMF  WORK IN PROGRESS"):
    """Restore MAIN sheet layout/styles so admins can edit without preserving format manually."""
    title_range = f"A1:{MAIN_COLS[-1][0]}1"
    if str(title_range) not in {str(r) for r in ws.merged_cells.ranges}:
        ws.merge_cells(title_range)

    ws.row_dimensions[1].height = 30
    title = ws["A1"]
    title.value = f"{title_text}   —   Updated {datetime.now().strftime('%B %d, %Y')}"
    title.font = _font(bold=True, color=C_WHITE, size=14)
    title.fill = _fill(C_NAVY)
    title.alignment = _align()

    ws.row_dimensions[HDR_ROW].height = 42
    for idx, (letter, label, width, is_blue) in enumerate(MAIN_COLS, 1):
        ws.column_dimensions[letter].width = width
        c = ws.cell(HDR_ROW, idx, label)
        c.border = _border()
        c.alignment = _align(wrap=True)
        if is_blue:
            c.font = _font(bold=True, size=8, color=C_STEEL_DK)
            c.fill = _fill("DEEAF1")
        else:
            c.font = _font(bold=True, size=8, color=C_WHITE)
            c.fill = _fill(C_NAVY)

    current_header = {"wo": None, "po": None, "company": None}
    part_alt = False
    autofilled = 0
    styled_parts = 0
    styled_headers = 0

    for row in range(DATA_START, ws.max_row + 1):
        kind = _main_row_kind(ws, row)

        if kind == "header":
            current_header = {
                "wo": ws.cell(row, 1).value,
                "po": ws.cell(row, 2).value,
                "company": ws.cell(row, 3).value,
            }
            styled_headers += 1
            ws.row_dimensions[row].height = HEADER_ROW_H
            for col in range(1, MAIN_LAST_COL + 1):
                c = ws.cell(row, col)
                c.fill = _fill(C_NAVY)
                c.border = _border("2E4D7B")
                c.alignment = _align(h="left")
                c.font = _font(bold=True, color=(C_GOLD if col == 1 else C_WHITE), size=11)
                if col == 5 and not _is_blank(c.value):
                    c.number_format = "m/d/yy"
            continue

        if kind != "part":
            continue

        if current_header["wo"] is not None:
            if _is_blank(ws.cell(row, 1).value):
                ws.cell(row, 1).value = current_header["wo"]; autofilled += 1
            if _is_blank(ws.cell(row, 2).value):
                ws.cell(row, 2).value = current_header["po"]; autofilled += 1
            if _is_blank(ws.cell(row, 3).value):
                ws.cell(row, 3).value = current_header["company"]; autofilled += 1

        part_alt = not part_alt
        styled_parts += 1
        bg = C_LGRAY if part_alt else C_WHITE
        ws.row_dimensions[row].height = PART_ROW_H

        for col in range(1, MAIN_LAST_COL + 1):
            c = ws.cell(row, col)
            is_blue = BLUE_FIRST <= col <= BLUE_LAST
            c.border = _border()
            c.alignment = _align(h="left" if col in (3, 6, 7, 39) else "center")
            c.font = _font(bold=(col in (1, 2, 3)), size=10)
            c.fill = _fill("DEEAF1" if is_blue and _is_blank(c.value) else (C_WHITE if is_blue else bg))

            if col in range(BLUE_FIRST, BLUE_LAST + 1) or col == 5:
                c.number_format = "m/d/yy"

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A{HDR_ROW}:{MAIN_COLS[-1][0]}{ws.max_row}"

    ws.conditional_formatting._cf_rules.clear()
    if ws.max_row >= DATA_START:
        first_col = get_column_letter(BLUE_FIRST)
        last_col = get_column_letter(BLUE_LAST)
        rng = f"{first_col}{DATA_START}:{last_col}{ws.max_row}"
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f"AND({first_col}{DATA_START}<>\"\",ISNUMBER({first_col}{DATA_START}),{first_col}{DATA_START}<TODAY())"],
            fill=_fill(C_RED_LT), font=Font(color="CC0000", bold=True),
        ))
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f"AND({first_col}{DATA_START}<>\"\",ISNUMBER({first_col}{DATA_START}),{first_col}{DATA_START}>=TODAY(),{first_col}{DATA_START}<TODAY()+4)"],
            fill=_fill(C_YELLOW), font=Font(color="806000", bold=True),
        ))

    ws.data_validations.dataValidation = []
    if ws.max_row >= DATA_START:
        dv = DataValidation(type="list", formula1=f'"{STEP_LIST}"', allow_blank=True, showErrorMessage=False)
        dv.sqref = f"AK{DATA_START}:AK{ws.max_row}"
        ws.add_data_validation(dv)

    print(f"  MAIN-REPAIR: styled {styled_headers} header row(s), {styled_parts} part row(s), autofilled {autofilled} key cell(s)")


def extract_row_images(ws):
    row_to_img = {}
    for img in ws._images:
        try:
            excel_row = img.anchor._from.row + 1
            ref = img.ref
            if hasattr(ref, 'read'):
                ref.seek(0)
                raw = ref.read()
                ref.seek(0)
            else:
                raw = img._data()
            if raw:
                row_to_img[excel_row] = raw
        except Exception:
            pass
    return row_to_img


def extract_wip_groups(ws):
    row_to_img = extract_row_images(ws)
    groups = OrderedDict()
    current_header = {"wo": None, "po": None, "company": None, "job_name": None, "cust_ship": None}

    for row in range(DATA_START, ws.max_row + 1):
        kind = _main_row_kind(ws, row)
        if kind == "blank":
            continue

        vals = [ws.cell(row, c).value for c in range(1, MAIN_LAST_COL + 1)]

        if kind == "header":
            current_header = {
                "wo": vals[0],
                "po": vals[1],
                "company": vals[2],
                "job_name": vals[3],
                "cust_ship": vals[4],
            }
            wo_key = str(vals[0]).strip() if vals[0] is not None else None
            if wo_key and wo_key not in groups:
                groups[wo_key] = {"header": current_header.copy(), "parts": []}
            elif wo_key:
                groups[wo_key]["header"].update({k: v for k, v in current_header.items() if not _is_blank(v)})
            continue

        wo = vals[0] if not _is_blank(vals[0]) else current_header["wo"]
        if _is_blank(wo):
            continue
        wo_key = str(wo).strip()
        if wo_key not in groups:
            groups[wo_key] = {
                "header": {
                    "wo": wo,
                    "po": vals[1] if not _is_blank(vals[1]) else current_header["po"],
                    "company": vals[2] if not _is_blank(vals[2]) else current_header["company"],
                    "job_name": current_header["job_name"],
                    "cust_ship": current_header["cust_ship"],
                },
                "parts": [],
            }

        header = groups[wo_key]["header"]
        for field, idx in [("po", 1), ("company", 2)]:
            if _is_blank(header[field]) and not _is_blank(vals[idx]):
                header[field] = vals[idx]
        if _is_blank(header["job_name"]) and not _is_blank(current_header["job_name"]):
            header["job_name"] = current_header["job_name"]
        if _is_blank(header["cust_ship"]) and not _is_blank(current_header["cust_ship"]):
            header["cust_ship"] = current_header["cust_ship"]

        groups[wo_key]["parts"].append({
            "values": vals,
            "image": row_to_img.get(row),
        })

    return groups


def part_is_ready_for_main(part):
    values = part["values"]
    current_raw = str(values[COL_CURRENT_STEP - 1] or "").strip().upper()
    has_process_dates = any(isinstance(values[c - 1], datetime) for c in range(BLUE_FIRST, 36))
    return bool(current_raw) and has_process_dates


def repartition_main_and_admin(main_groups, admin_groups):
    new_main = OrderedDict()
    new_admin = OrderedDict()

    def push(target, wo_key, header, part):
        if wo_key not in target:
            target[wo_key] = {"header": header.copy(), "parts": []}
        target[wo_key]["parts"].append(part)

    for source in (main_groups, admin_groups):
        for wo_key, group in source.items():
            header = group["header"]
            for part in group["parts"]:
                target = new_main if part_is_ready_for_main(part) else new_admin
                push(target, wo_key, header, part)

    return new_main, new_admin


def rebuild_wip_sheet(wb, sheet_name, title_text, groups, index):
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name, index)

    out_row = DATA_START
    for _, group in groups.items():
        if not group["parts"]:
            continue
        header = group["header"]
        ws.cell(out_row, 1, header["wo"])
        ws.cell(out_row, 2, header["po"])
        ws.cell(out_row, 3, header["company"])
        ws.cell(out_row, 4, header["job_name"])
        ws.cell(out_row, 5, header["cust_ship"])
        out_row += 1

        for part in group["parts"]:
            vals = list(part["values"])
            if _is_blank(vals[0]): vals[0] = header["wo"]
            if _is_blank(vals[1]): vals[1] = header["po"]
            if _is_blank(vals[2]): vals[2] = header["company"]
            for col_idx, val in enumerate(vals, 1):
                ws.cell(out_row, col_idx, val)
            if part.get("image"):
                copy_image(part["image"], ws, COL_SCREENSHOT, out_row)
            out_row += 1

    repair_main_sheet(ws, title_text=title_text)
    return ws


def sync_admin_input_flow(wb):
    main_groups = extract_wip_groups(wb["MAIN"])
    admin_groups = extract_wip_groups(wb["ADMIN INPUT"]) if "ADMIN INPUT" in wb.sheetnames else OrderedDict()
    new_main, new_admin = repartition_main_and_admin(main_groups, admin_groups)

    rebuild_wip_sheet(wb, "MAIN", "CMF  WORK IN PROGRESS", new_main, 0)
    rebuild_wip_sheet(wb, "ADMIN INPUT", "CMF  ADMIN INPUT  —  New Orders / Routing Review", new_admin, 1)

    main_parts = sum(len(g["parts"]) for g in new_main.values())
    admin_parts = sum(len(g["parts"]) for g in new_admin.values())
    print(f"  ADMIN-FLOW: MAIN has {main_parts} routed part(s) | ADMIN INPUT has {admin_parts} pending part(s)")


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
        thick   = ws.cell(row, COL_THICKNESS).value
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
                    continue    # already done: earlier date
                if (curr_step_date and due == curr_step_date
                        and curr_step_col and col_idx < curr_step_col):
                    continue    # already done: same date, earlier process column

                process_name = _COL_TO_DROPDOWN_STEP.get(col_idx, COL_TO_PROCESS.get(col_idx, s_name))
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
                    thickness     = str(thick   or "").strip(),
                    qty           = qty,
                    current_step  = curr_raw or "—",
                    processes_left= processes_left,
                    delivery_date = delivery_date,
                    process_name  = process_name,
                    step_status   = step_status,
                ))

    return entries, row_to_img


def _raw_image_bytes(img):
    ref = getattr(img, "ref", None)
    if hasattr(ref, "read"):
        ref.seek(0)
        raw = ref.read()
        ref.seek(0)
        return raw
    return img._data()


def _image_digest(raw):
    return hashlib.sha256(raw).hexdigest()[:16] if raw else None


def _worksheet_image_digests_by_anchor(ws):
    images = defaultdict(list)
    for img in getattr(ws, "_images", []):
        try:
            anchor = getattr(img, "anchor", None)
            if not anchor or not hasattr(anchor, "_from"):
                continue
            row = anchor._from.row + 1
            col = anchor._from.col + 1
            digest = _image_digest(_raw_image_bytes(img))
            if digest:
                images[(row, col)].append(digest)
        except Exception:
            continue
    return images


def audit_board_images(wb, row_to_img, sheet_names=("CALENDAR", "PURCHASING")):
    """Verify generated board photos still match each card's MAIN row ID."""
    expected_by_main_row = {
        main_row: _image_digest(raw)
        for main_row, raw in row_to_img.items()
        if raw
    }
    problems = []

    for sheet_name in sheet_names:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        anchored_images = _worksheet_image_digests_by_anchor(ws)
        expected_photo_cells = set()
        audited_cards = 0

        for row in range(5, ws.max_row + 1):
            col = 1
            while col <= ws.max_column:
                id_val = str(ws.cell(row, col).value or "")
                if "|" not in id_val:
                    col += 1
                    continue

                sub_n = _board_subcol_count(ws, col)
                photo_offset = _board_header_offset(ws, col, sub_n, "PHOTO")
                if photo_offset is None:
                    problems.append(f"{sheet_name} R{row} C{col}: missing PHOTO header")
                    col += sub_n + CAL_DIV_W
                    continue

                try:
                    main_row = int(id_val.split("|", 1)[0])
                except ValueError:
                    problems.append(f"{sheet_name} R{row} C{col}: invalid card ID {id_val!r}")
                    col += sub_n + CAL_DIV_W
                    continue

                photo_cell = (row, col + photo_offset)
                expected_photo_cells.add(photo_cell)
                expected = expected_by_main_row.get(main_row)
                actual = anchored_images.get(photo_cell, [])
                audited_cards += 1

                if expected and expected not in actual:
                    got = ",".join(actual) if actual else "missing"
                    problems.append(
                        f"{sheet_name} R{row} photo for MAIN row {main_row}: expected {expected}, got {got}"
                    )
                elif not expected and actual:
                    problems.append(
                        f"{sheet_name} R{row} photo for MAIN row {main_row}: unexpected image {','.join(actual)}"
                    )

                col += sub_n + CAL_DIV_W

        extra_images = sorted(set(anchored_images) - expected_photo_cells)
        for row, col in extra_images[:10]:
            problems.append(f"{sheet_name} R{row} C{col}: image is not attached to a generated PHOTO cell")
        if len(extra_images) > 10:
            problems.append(f"{sheet_name}: {len(extra_images) - 10} additional non-card image(s)")

        print(f"  IMAGE-AUDIT: {sheet_name} checked {audited_cards} card(s)")

    if problems:
        sample = "\n    ".join(problems[:20])
        more = f"\n    ... {len(problems) - 20} more" if len(problems) > 20 else ""
        raise RuntimeError(f"Board image audit failed:\n    {sample}{more}")


# ── CALENDAR ─────────────────────────────────────────────────────────────────
#   STATUS | DUE | COMPANY | WO# | PART# | THICK | QTY | UPDATED BY | PHOTO
# ── CALENDAR — Horizontal Kanban ─────────────────────────────────────────────
# Columns = stations, items stacked under each station sorted by date.
# Makes it easy for PM to see workstation loading and queue depth at a glance.
#
#  Row 1: Title
#  Row 2: Overdue status banner
#  Row 3: Station headers (colored, merged across sub-cols)
#  Row 4: Sub-headers: STATUS | DUE | COMPANY | WO # | PART # | THICK | QTY | PHOTO per station
#  Row 5+: Items sorted date-ascending (overdue at top in coral)

CAL_SUB    = [("_ID", 0.5), ("STATUS", 8), ("DUE", 6), ("SHIP", 6), ("COMPANY", 11), ("WO #", 6), ("PART #", 14), ("THICK", 7), ("QTY", 7), ("UPDATED BY", 14), ("PHOTO", 11)]
CAL_SUB_N  = len(CAL_SUB)   # sub-cols per card (first is hidden row ID)
CAL_UPDATED_BY_OFFSET = next(i for i, (lbl, _) in enumerate(CAL_SUB) if lbl == "UPDATED BY")
CAL_PHOTO_OFFSET = next(i for i, (lbl, _) in enumerate(CAL_SUB) if lbl == "PHOTO")
CAL_DIV_W  = 1               # divider col between stations
CAL_BLOCK  = CAL_SUB_N + CAL_DIV_W  # data cols + divider per station
CAL_ROW_H  = 52              # tall enough for photo thumbnails
CAL_IMG_W  = 68              # photo width in pixels
CAL_IMG_H  = 44              # photo height in pixels

CAL_STATION_COLOR = {s[0]: s[3] for s in STATIONS}
PENDING_GREEN_HIGHLIGHTS = {}

_TEMPLATE_FILL_HEXES = {
    C_NAVY, C_WHITE, C_GOLD, C_STEEL_DK, C_LGRAY, C_DGRAY, C_YELLOW,
    C_RED_LT, C_GREEN_LT, "FCA5A5", "DBEAFE", "F3F4F6", "D0D0D0",
    "DDDDDD", "CCCCCC", "AAAAAA",
}
_TEMPLATE_FILL_HEXES.update(s[2] for s in STATIONS)
_TEMPLATE_FILL_HEXES.update(s[3] for s in STATIONS)


def _board_due_display(due_dt, today):
    if not isinstance(due_dt, datetime):
        return "—"
    days = (due_dt.date() - today).days
    date_part = f"{due_dt.month}/{due_dt.day}"
    if 0 <= days <= 6:
        return f"{due_dt.strftime('%a')} {date_part}"
    return date_part


def _build_horizontal_board(wb, sheet_name, title, entries, row_to_img, station_keys=None):
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    today = datetime.now().date()

    # Active stations in defined display order (skip SHIPPING — delivery date is a column)
    active = [s for s in STATIONS
              if s[0] != "SHIPPING"
              and (station_keys is None or s[0] in station_keys)
              and any(e["s_key"] == s[0] for e in entries)]

    if not active:
        ws["A1"] = "No scheduled work found."
        print(f"  {sheet_name}: no entries")
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
    c.value     = f"{title}   —   {today.strftime('%B %d, %Y')}"
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
            if lbl == "_ID":        # hidden ID column — blank header, no border
                c = ws.cell(4, sc + i, "")
                c.fill = _fill(STATION_ROW_C[s[0]])
                c.font = _font(size=6, color=STATION_ROW_C[s[0]])
            else:
                c = ws.cell(4, sc + i, lbl)
                c.font      = _font(bold=True, size=9, color=C_DGRAY)
                c.fill      = _fill(STATION_ROW_C[s[0]])
                c.border    = _border("AAAAAA"); c.alignment = _align()
        ws.cell(4, sc + CAL_SUB_N).fill = _fill("D0D0D0")

    # ── Rows 5+: Item cards ───────────────────────────────────────────────────
    employee_dv = _employee_name_validation()
    ws.add_data_validation(employee_dv)

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
                due_str = _board_due_display(e["date"], today)

                # _ID | STATUS | DUE | SHIP | COMPANY | WO# | PART# | THICK | QTY | UPDATED BY | PHOTO
                dd = e.get("delivery_date")
                ship_str = f"{dd.month}/{dd.day}" if dd else "—"
                id_val = f"{e['main_row']}|{e['s_key']}"
                vals = [id_val, status, due_str, ship_str, e["company"], e["wo"], e["part_no"], e["thickness"], e["qty"], None, None]
                pending_green_fill = PENDING_GREEN_HIGHLIGHTS.get(id_val)
                for i, val in enumerate(vals):
                    c = ws.cell(row, sc + i, val)

                    if i == 0:      # hidden ID — invisible text, no border
                        c.fill = pending_green_fill or _fill(bg)
                        c.font = _font(size=6, color="92D050" if pending_green_fill else bg)
                        continue

                    c.border    = _border("CCCCCC")
                    c.alignment = _align(h="left" if i == 4 else "center")

                    if i == 1:      # STATUS label — coloured independently
                        if pending_green_fill:
                            c.fill = pending_green_fill
                            c.font = _font(bold=True, size=9, color=C_DGRAY)
                        elif status == "▶ READY":
                            c.fill = _fill("DBEAFE")
                            c.font = _font(bold=True, size=9, color="1E40AF")
                        elif status.startswith("⏳"):
                            c.fill = _fill("F3F4F6")
                            c.font = _font(size=9, color="6B7280")
                        else:
                            c.fill = _fill(bg); c.font = _font(size=9)
                    elif i == 2:    # DUE date
                        c.fill = pending_green_fill or _fill(bg)
                        c.font = _font(bold=True, size=10, color=due_fg)
                    elif i == 3:    # SHIP date
                        c.fill = pending_green_fill or _fill(bg)
                        c.font = _font(size=9, color="6B7280")
                    elif i == 4:    # COMPANY
                        c.fill = pending_green_fill or _fill(bg)
                        c.font = _font(bold=True, size=10)
                    elif i == CAL_UPDATED_BY_OFFSET:
                        c.fill = pending_green_fill or _fill(bg)
                        c.font = _font(size=9)
                        employee_dv.add(c)
                    elif i == 6:    # PART #
                        c.fill = pending_green_fill or _fill(bg)
                        c.font = _font(size=9)
                    else:           # WO#, THICK, QTY, and PHOTO placeholder
                        c.fill = pending_green_fill or _fill(bg); c.font = _font(size=10)

                # Screenshot in PHOTO column
                if e["main_row"] in row_to_img:
                    copy_image(row_to_img[e["main_row"]], ws, sc + CAL_PHOTO_OFFSET, row)
            else:
                # Empty slot — skip border on hidden ID column
                for i in range(CAL_SUB_N):
                    c = ws.cell(row, sc + i)
                    c.fill = _fill(STATION_ROW_C[s[0]])
                    if i > 0:
                        c.border = _border("DDDDDD")

            ws.cell(row, sc + CAL_SUB_N).fill = _fill("D0D0D0")

    ws.freeze_panes = "A5"
    visible_count = sum(len(items) for items in station_items.values())
    print(f"  {sheet_name}: {visible_count} entries across {len(active)} stations  ({overdue_count} overdue)")


def build_calendar(wb, entries, row_to_img):
    _build_horizontal_board(
        wb,
        "CALENDAR",
        "CMF  PRODUCTION CALENDAR",
        [e for e in entries if e["s_key"] in CALENDAR_STATION_KEYS],
        row_to_img,
        station_keys=CALENDAR_STATION_KEYS,
    )


def build_purchasing(wb, entries, row_to_img):
    if "TODAY" in wb.sheetnames:
        del wb["TODAY"]
    _build_horizontal_board(
        wb,
        "PURCHASING",
        "CMF  PURCHASING / OUTSIDE WORK",
        [e for e in entries if e["s_key"] in PURCHASING_STATION_KEYS],
        row_to_img,
        station_keys=PURCHASING_STATION_KEYS,
    )


# ── Green-completion sync ─────────────────────────────────────────────────────

def _is_user_green(cell):
    """Detect any shade of green fill applied by the user.

    Handles:
    - Theme fills: index 6 = 'Olive Green, Accent 3' in the default Office theme.
      Excel saves theme-color picks without an explicit RGB value; openpyxl
      returns type='theme' with theme=6 for these cells.
    - Explicit RGB fills: green is the dominant channel, including light
      Excel greens. Known schedule template fills are excluded first so normal
      board colors do not count as operator completions.
    """
    try:
        fill = cell.fill
        if not fill or fill.fill_type != "solid":
            return False
        fg = fill.fgColor

        # ── Theme color (Olive Green Accent 3 = index 6 in default theme) ─
        if fg.type == "theme":
            return int(fg.theme) == 6

        # ── Explicit RGB ──────────────────────────────────────────────────
        rgb = str(fg.rgb or "")
        if not rgb or rgb in ("00000000", "FF000000", "FFFFFFFF"):
            return False
        if len(rgb) == 8:   # ARGB  e.g. "FF9BBB59" or "009BBB59"
            r = int(rgb[2:4], 16); g = int(rgb[4:6], 16); b = int(rgb[6:8], 16)
        elif len(rgb) == 6: # RGB   e.g. "9BBB59"
            r = int(rgb[0:2], 16); g = int(rgb[2:4], 16); b = int(rgb[4:6], 16)
        else:
            return False
        clean_rgb = f"{r:02X}{g:02X}{b:02X}"
        if clean_rgb in _TEMPLATE_FILL_HEXES:
            return False
        return g >= 80 and g >= r + 8 and g >= b + 8
    except Exception:
        return False


def _green_fill_for_preserve(cell):
    if not _is_user_green(cell):
        return None
    try:
        return copy.copy(cell.fill)
    except Exception:
        return _fill("92D050")


def _board_subcol_count(ws, start_col):
    """Return the number of sub-columns in a station block from row-4 headers."""
    offset = 1
    while start_col + offset <= ws.max_column:
        header = str(ws.cell(4, start_col + offset).value or "").strip()
        if not header:
            break
        offset += 1
    return max(offset, 1)


def _board_header_offset(ws, start_col, sub_n, label):
    wanted = label.strip().upper()
    for offset in range(1, sub_n):
        if str(ws.cell(4, start_col + offset).value or "").strip().upper() == wanted:
            return offset
    return None


def _norm_match_key(val):
    """Normalize visible board identifiers so they survive Excel type changes."""
    if val is None:
        return None
    s = str(val).strip()
    try:
        f = float(s)
        if f == int(f):
            s = str(int(f))
    except ValueError:
        pass
    return s or None


def _board_value(ws, row, start_col, sub_n, label):
    offset = _board_header_offset(ws, start_col, sub_n, label)
    if offset is None:
        return None
    return ws.cell(row, start_col + offset).value


def _scan_green_completions_from_board(ws, sheet_name):
    """Scan one horizontal board for rows highlighted green.

    Each station block's first sub-column (_ID) stores "main_row|s_key".
    If any visible cell in that block has a user-green fill and UPDATED BY is
    filled in, we record it.
    """
    found = []
    skipped_missing_name = []

    for row in range(5, ws.max_row + 1):
        col = 1
        while col <= ws.max_column:
            id_val = str(ws.cell(row, col).value or "")
            if "|" in id_val:
                sub_n = _board_subcol_count(ws, col)
                # Any visible cell in this block green?
                green_fill = None
                for offset in range(1, sub_n):
                    green_fill = _green_fill_for_preserve(ws.cell(row, col + offset))
                    if green_fill:
                        break
                if green_fill:
                    updated_by = ""
                    updated_by_offset = _board_header_offset(ws, col, sub_n, "UPDATED BY")
                    if updated_by_offset is not None:
                        updated_by = str(ws.cell(row, col + updated_by_offset).value or "").strip()
                    if not updated_by:
                        skipped_missing_name.append((row, col, id_val, green_fill))
                        col += sub_n + CAL_DIV_W
                        continue

                    parts = id_val.split("|", 1)
                    if len(parts) == 2:
                        try:
                            found.append({
                                "main_row": int(parts[0]),
                                "s_key": parts[1],
                                "source_sheet": sheet_name,
                                "updated_by": updated_by,
                                "visible_wo": _board_value(ws, row, col, sub_n, "WO #"),
                                "visible_part": _board_value(ws, row, col, sub_n, "PART #"),
                                "visible_company": _board_value(ws, row, col, sub_n, "COMPANY"),
                                "visible_qty": _board_value(ws, row, col, sub_n, "QTY"),
                            })
                        except ValueError:
                            pass
                col += sub_n + CAL_DIV_W   # jump to next station block
                continue
            col += 1

    return found, skipped_missing_name


def _build_main_row_index(ws_main):
    by_wo_part = {}
    by_wo_company_qty = defaultdict(list)
    by_wo = defaultdict(list)

    for row in range(DATA_START, ws_main.max_row + 1):
        kind = _main_row_kind(ws_main, row)
        if kind != "part":
            continue
        wo = _norm_match_key(ws_main.cell(row, COL_WO).value)
        if not wo:
            continue
        part = _norm_match_key(ws_main.cell(row, COL_PARTNO).value)
        company = _norm_match_key(ws_main.cell(row, COL_COMPANY).value)
        qty = _norm_match_key(ws_main.cell(row, COL_QTY).value)

        by_wo[wo].append(row)
        if part:
            by_wo_part.setdefault((wo, part), row)
        by_wo_company_qty[(wo, company, qty)].append(row)

    return {
        "by_wo_part": by_wo_part,
        "by_wo_company_qty": by_wo_company_qty,
        "by_wo": by_wo,
    }


def _resolve_completion_main_row(comp, ws_main, row_index):
    """Prefer visible card identity over the hidden row ID.

    The board's hidden _ID includes an old MAIN row number. That row number can
    become stale when MAIN is edited, rebuilt, or repartitioned, while the
    visible WO/part/qty on the card still describes the operator's intent.
    """
    wo = _norm_match_key(comp.get("visible_wo"))
    part = _norm_match_key(comp.get("visible_part"))
    company = _norm_match_key(comp.get("visible_company"))
    qty = _norm_match_key(comp.get("visible_qty"))

    if wo and part:
        row = row_index["by_wo_part"].get((wo, part))
        if row:
            return row

    if wo:
        candidates = row_index["by_wo_company_qty"].get((wo, company, qty), [])
        if len(candidates) == 1:
            return candidates[0]
        candidates = row_index["by_wo"].get(wo, [])
        if len(candidates) == 1:
            return candidates[0]

    main_row = comp["main_row"]
    if DATA_START <= main_row <= ws_main.max_row:
        return main_row
    return None


def sync_green_completions(wb):
    """Scan existing CALENDAR and PURCHASING boards for green completions.

    Purchasing uses the same horizontal board layout and hidden row ID as
    CALENDAR, so vendor/outside-work completions can update MAIN the same way.

    Must be called BEFORE build_calendar()/build_purchasing() delete the sheets.
    Returns list of {"main_row": int, "s_key": str}.
    """
    PENDING_GREEN_HIGHLIGHTS.clear()
    seen = set()
    found = []
    row_index = _build_main_row_index(wb["MAIN"]) if "MAIN" in wb.sheetnames else None

    for sheet_name in ("CALENDAR", "PURCHASING"):
        if sheet_name not in wb.sheetnames:
            continue

        sheet_found, skipped_missing_name = _scan_green_completions_from_board(wb[sheet_name], sheet_name)
        print(f"  GREEN-SYNC: {len(sheet_found)} completion(s) detected in {sheet_name}")
        if skipped_missing_name:
            print(f"  GREEN-SYNC: {len(skipped_missing_name)} green row(s) skipped in {sheet_name} — UPDATED BY is blank")
            for _, _, id_val, green_fill in skipped_missing_name:
                if green_fill:
                    PENDING_GREEN_HIGHLIGHTS[id_val] = green_fill

        for comp in sheet_found:
            if row_index is not None:
                resolved_row = _resolve_completion_main_row(comp, wb["MAIN"], row_index)
                if resolved_row is None:
                    continue
                if resolved_row != comp["main_row"]:
                    print(
                        f"  GREEN-SYNC: remapped stale ID row {comp['main_row']} → {resolved_row} "
                        f"for WO {comp.get('visible_wo')} [{STATION_NAME.get(comp['s_key'], comp['s_key'])}]"
                    )
                comp["main_row"] = resolved_row

            key = (comp["main_row"], comp["s_key"])
            if key in seen:
                continue
            seen.add(key)
            found.append({
                "main_row": comp["main_row"],
                "s_key": comp["s_key"],
                "updated_by": comp["updated_by"],
                "updated_from": comp["source_sheet"],
                "visible_wo": comp.get("visible_wo"),
                "visible_part": comp.get("visible_part"),
            })

    print(f"  GREEN-SYNC: {len(found)} unique completion(s) detected")
    return found


def apply_completions(ws_main, completions):
    """Advance CURRENT_STEP in MAIN for every green-marked item.

    For each completed station:
      1. Find the latest scheduled date among that station's columns.
      2. Scan remaining process columns for the next earliest date.
      3. Write the canonical step name to CURRENT_STEP (col AK).
      4. If nothing remains, mark COMPLETE.

    Returns list of record dicts for the LOG sheet.
    """
    if not completions:
        return []

    today = datetime.now()

    grouped = OrderedDict()
    for comp in completions:
        main_row = comp["main_row"]
        s_key    = comp["s_key"]

        station_def = next((s for s in STATIONS if s[0] == s_key), None)
        if station_def is None:
            continue
        col_indices     = station_def[4]

        # Pull context from MAIN
        wo           = ws_main.cell(main_row, COL_WO).value
        company      = str(ws_main.cell(main_row, COL_COMPANY).value or "")
        part_no      = str(ws_main.cell(main_row, COL_PARTNO).value or "")
        desc         = str(ws_main.cell(main_row, COL_DESC).value or "")
        delivery_raw = ws_main.cell(main_row, 36).value
        delivery_dt  = delivery_raw if isinstance(delivery_raw, datetime) else None

        # Completed date = latest date in this station's column(s)
        completed_date = None
        for c in col_indices:
            v = ws_main.cell(main_row, c).value
            if isinstance(v, datetime):
                if completed_date is None or v > completed_date:
                    completed_date = v
        if completed_date is None:
            continue

        grouped.setdefault(main_row, []).append({
            "comp": comp,
            "station_def": station_def,
            "completed_date": completed_date,
            "wo": wo,
            "company": company,
            "part_no": part_no,
            "desc": desc,
            "delivery_dt": delivery_dt,
        })

    records = []
    for main_row, events in grouped.items():
        completed_cols = set()
        latest_completed_date = None
        for event in events:
            col_indices = event["station_def"][4]
            completed_cols.update(col_indices)
            completed_date = event["completed_date"]
            if latest_completed_date is None or completed_date > latest_completed_date:
                latest_completed_date = completed_date

        # Next step: earliest remaining process date after all green-marked
        # completions for this row. Same-date steps at a higher column number
        # than the latest completed station are still "next" (not done).
        max_completed_col = max(
            c
            for event in events
            if event["completed_date"] == latest_completed_date
            for c in event["station_def"][4]
        )
        remaining = []
        for c in range(12, 36):
            if c in completed_cols:
                continue
            v = ws_main.cell(main_row, c).value
            if not isinstance(v, datetime) or c not in _COL_TO_DROPDOWN_STEP:
                continue
            if v > latest_completed_date or (v == latest_completed_date and c > max_completed_col):
                remaining.append((v, c))
        remaining.sort()

        if remaining:
            _, next_col = remaining[0]
            next_step   = _COL_TO_DROPDOWN_STEP[next_col]
            ws_main.cell(main_row, COL_CURRENT_STEP).value = next_step
        else:
            next_step = "COMPLETE"
            ws_main.cell(main_row, COL_CURRENT_STEP).value = "COMPLETE"
            ws_main.cell(main_row, COL_STATUS).value       = "COMPLETE"

        for event in events:
            comp = event["comp"]
            s_key = comp["s_key"]
            label = STATION_NAME.get(s_key, s_key)
            print(f"    WO {event['wo']}  [{label}]  →  {next_step}")

            records.append(dict(
                date_completed = today,
                wo             = event["wo"],
                company        = event["company"],
                part_no        = event["part_no"],
                desc           = event["desc"],
                process        = label,
                due_date       = event["completed_date"],
                next_step      = next_step,
                delivery_date  = event["delivery_dt"],
                updated_by     = comp.get("updated_by", ""),
                updated_from   = comp.get("updated_from", ""),
                owner_station  = label,
            ))

    return records


# ── Completion LOG sheet ──────────────────────────────────────────────────────

_LOG_COLS = [
    ("DATE",          10),
    ("WO #",           8),
    ("COMPANY",       20),
    ("PART #",        14),
    ("DESCRIPTION",   26),
    ("PROCESS DONE",  22),
    ("WAS DUE",       10),
    ("NEXT STEP",     18),
    ("DELIVERY DATE", 14),
    ("UPDATED BY",    16),
    ("UPDATED FROM",  14),
    ("PM REVIEW",     14),
    ("REVIEWED BY",   16),
    ("OWNER / STATION", 18),
    ("ISSUE NOTES",   32),
    ("RESOLUTION",    14),
]
_LOG_NCOLS    = len(_LOG_COLS)
_LOG_LAST_COL = get_column_letter(_LOG_NCOLS)
_LOG_BASE_LAST_COL = "I"
_LOG_PM_REVIEW_COL = 12
_LOG_REVIEWED_BY_COL = 13
_LOG_RESOLUTION_COL = 16

_PM_REVIEW_LIST = "OK,ISSUE,NEEDS CHECK"
_RESOLUTION_LIST = "Open,Resolved"


def _unmerge_ranges_on_row(ws, row):
    for merge_range in list(ws.merged_cells.ranges):
        if merge_range.min_row <= row <= merge_range.max_row:
            try:
                ws.unmerge_cells(str(merge_range))
            except KeyError:
                try:
                    ws.merged_cells.ranges.discard(merge_range)
                except AttributeError:
                    if merge_range in ws.merged_cells.ranges:
                        ws.merged_cells.ranges.remove(merge_range)


def _is_log_banner_row(ws, row):
    val = str(ws.cell(row, 1).value or "").strip()
    return bool(val) and row >= 3 and ws.cell(row, 2).value is None


def _add_log_validations(ws):
    ws.data_validations.dataValidation = []

    review_dv = DataValidation(type="list", formula1=f'"{_PM_REVIEW_LIST}"', allow_blank=True, showErrorMessage=False)
    reviewer_dv = _employee_name_validation()
    resolution_dv = DataValidation(type="list", formula1=f'"{_RESOLUTION_LIST}"', allow_blank=True, showErrorMessage=False)

    ws.add_data_validation(review_dv)
    ws.add_data_validation(reviewer_dv)
    ws.add_data_validation(resolution_dv)

    for row in range(3, max(ws.max_row, 300) + 1):
        if _is_log_banner_row(ws, row):
            reviewer_dv.add(ws.cell(row, 11))  # banner SIGNED BY cell
            continue
        review_dv.add(ws.cell(row, _LOG_PM_REVIEW_COL))
        reviewer_dv.add(ws.cell(row, _LOG_REVIEWED_BY_COL))
        resolution_dv.add(ws.cell(row, _LOG_RESOLUTION_COL))


def _refresh_log_banner_summaries(ws):
    for row in range(3, ws.max_row + 1):
        if not _is_log_banner_row(ws, row):
            continue

        next_banner = ws.max_row + 1
        for r in range(row + 1, ws.max_row + 1):
            if _is_log_banner_row(ws, r):
                next_banner = r
                break

        completion_count = 0
        issue_count = 0
        for r in range(row + 1, next_banner):
            if ws.cell(r, 1).value is None and ws.cell(r, 2).value is None:
                continue
            completion_count += 1
            if str(ws.cell(r, _LOG_PM_REVIEW_COL).value or "").strip().upper() == "ISSUE":
                issue_count += 1

        _write_log_date_banner(ws, row, str(ws.cell(row, 1).value or "").strip(), completion_count, issue_count)


def _init_log_sheet(ws):
    _unmerge_ranges_on_row(ws, 1)
    ws.row_dimensions[1].height = 30
    ws.merge_cells(f"A1:{_LOG_LAST_COL}1")
    c = ws["A1"]
    c.value     = "CMF  PRODUCTION LOG  —  Completed Processes"
    c.font      = _font(bold=True, color=C_WHITE, size=13)
    c.fill      = _fill(C_NAVY); c.alignment = _align(h="left")

    ws.row_dimensions[2].height = 16
    for i, (lbl, w) in enumerate(_LOG_COLS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
        c = ws.cell(2, i, lbl)
        c.font      = _font(bold=True, color=C_WHITE, size=9)
        c.fill      = _fill(C_STEEL_DK)
        c.border    = _border(); c.alignment = _align()

    ws.freeze_panes = "A3"
    _add_log_validations(ws)


def _write_log_date_banner(ws, row, date_str, completion_count=0, issue_count=0):
    _unmerge_ranges_on_row(ws, row)
    clean_date = str(date_str or "").strip()
    ws.merge_cells(f"A{row}:{_LOG_BASE_LAST_COL}{row}")
    c = ws.cell(row, 1, f"  {clean_date}")
    c.font      = _font(bold=True, color=C_WHITE, size=10)
    c.fill      = _fill(C_STEEL_DK); c.alignment = _align(h="left")
    ws.row_dimensions[row].height = 18

    signoff_vals = [
        "SIGNED BY",
        ws.cell(row, 11).value,
        "COMPLETIONS",
        completion_count,
        "ISSUES",
        issue_count,
    ]
    for offset, val in enumerate(signoff_vals, 10):
        c = ws.cell(row, offset, val)
        c.fill = _fill(C_STEEL_DK if offset in (10, 12, 14) else C_LGRAY)
        c.font = _font(bold=True if offset in (10, 12, 14) else False, color=C_WHITE if offset in (10, 12, 14) else C_DGRAY, size=9)
        c.border = _border("1F4E79")
        c.alignment = _align(h="left" if offset == 11 else "center")


def _write_log_row(ws, row, rec, alt=False):
    bg = C_LGRAY if alt else C_WHITE
    ws.row_dimensions[row].height = 15
    vals = [
        rec["date_completed"].strftime("%Y-%m-%d"),
        rec["wo"],
        rec["company"],
        rec["part_no"],
        rec["desc"],
        rec["process"],
        rec["due_date"].strftime("%m/%d") if isinstance(rec["due_date"], datetime) else "—",
        rec["next_step"],
        rec["delivery_date"].strftime("%m/%d/%Y") if rec["delivery_date"] else "—",
        rec.get("updated_by", ""),
        rec.get("updated_from", ""),
        "NEEDS CHECK",
        "",
        rec.get("owner_station", rec.get("process", "")),
        "",
        "Open",
    ]
    for i, val in enumerate(vals, 1):
        c = ws.cell(row, i, val)
        c.fill      = _fill(bg); c.border = _border()
        c.font      = _font(size=9, bold=(i in (2, 6, 8, 10, 12)))
        c.alignment = _align(h="left" if i in (3, 4, 5, 8, 10, 13, 15, 16) else "center", wrap=(i in (5, 16)))


def record_completions_to_log(wb, records):
    """Prepend today's completions to the LOG sheet (newest entries at top).

    Structure:
      Row 1: title banner
      Row 2: column headers (frozen)
      Row 3+: date banners + completion rows, newest date at top
    """
    if not records:
        return

    today_str = datetime.now().strftime("%B %d, %Y")

    if "LOG" not in wb.sheetnames:
        ws = wb.create_sheet("LOG")
        _init_log_sheet(ws)
        insert_at    = 3   # first entry goes right after header rows
        needs_banner = True
    else:
        ws = wb["LOG"]
        _init_log_sheet(ws)

        # Check if row 3 already has today's date banner (same-day re-run)
        banner_val = str(ws.cell(3, 1).value or "").strip()
        if banner_val == today_str or banner_val == f"  {today_str}":
            # Append within today's section — find where it ends
            r = 4
            while r <= ws.max_row + 1:
                try:
                    rgb = ws.cell(r, 1).fill.fgColor.rgb or "00000000"
                except Exception:
                    rgb = "00000000"
                # Next date banner uses C_STEEL_DK ("1F4E79") — stop here
                if "1F4E79" in rgb:
                    break
                if ws.cell(r, 1).value is None and ws.cell(r, 2).value is None:
                    break
                r += 1
            insert_at    = r
            needs_banner = False
        else:
            insert_at    = 3   # prepend before existing entries
            needs_banner = True

    n_insert = len(records) + (1 if needs_banner else 0)
    ws.insert_rows(insert_at, amount=n_insert)

    r = insert_at
    if needs_banner:
        _write_log_date_banner(ws, r, today_str)
        r += 1

    for i, rec in enumerate(records):
        _write_log_row(ws, r + i, rec, i % 2 == 0)

    _refresh_log_banner_summaries(ws)
    _add_log_validations(ws)
    print(f"  LOG: {len(records)} completion(s) recorded — {today_str}")


# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print(f"Opening  {FILE}")
    if not os.path.exists(FILE):
        print("ERROR: Run cmf_migrate_main.py first."); raise SystemExit(1)

    wb = openpyxl.load_workbook(FILE)
    protect_workbook_images(wb)

    if "MAIN" not in wb.sheetnames:
        print("ERROR: No MAIN sheet. Run cmf_migrate_main.py first."); raise SystemExit(1)

    print("Checking SETTINGS employee list ...")
    ensure_settings_sheet(wb)
    if "LOG" in wb.sheetnames:
        _init_log_sheet(wb["LOG"])
        _refresh_log_banner_summaries(wb["LOG"])
        _add_log_validations(wb["LOG"])

    print("Checking MAIN column layout ...")
    ensure_ship_vendor_col(wb["MAIN"])
    if "ADMIN INPUT" in wb.sheetnames:
        ensure_ship_vendor_col(wb["ADMIN INPUT"])

    print("Syncing green completions from CALENDAR + PURCHASING ...")
    completions = sync_green_completions(wb)

    if completions:
        print("Advancing CURRENT_STEP in MAIN ...")
        records = apply_completions(wb["MAIN"], completions)
        print("Recording to LOG ...")
        record_completions_to_log(wb, records)
    else:
        print("  No green completions found — MAIN unchanged")

    print("Syncing ADMIN INPUT flow ...")
    sync_admin_input_flow(wb)

    print("Parsing MAIN ...")
    entries, row_to_img = parse_main(wb["MAIN"])
    print(f"  {len(entries)} process-due entries  |  {len(row_to_img)} screenshots")

    print("Building CALENDAR ...")
    build_calendar(wb, entries, row_to_img)

    print("Building PURCHASING ...")
    build_purchasing(wb, entries, row_to_img)

    print("Auditing board images ...")
    audit_board_images(wb, row_to_img)

    print("Normalizing image anchors ...")
    normalize_image_anchors(wb)

    if "MERGE CHANGES" in wb.sheetnames:
        del wb["MERGE CHANGES"]

    desired_order = ["ADMIN INPUT", "MAIN", "CALENDAR", "PURCHASING", "LOG", "SETTINGS"]
    ordered = [wb[name] for name in desired_order if name in wb.sheetnames]
    remaining = [ws for ws in wb.worksheets if ws.title not in desired_order]
    wb._sheets = ordered + remaining

    print(f"Saving   {FILE}")
    wb.save(FILE)
    print("Done ✓\n")
    print("  MAIN       → engineer fills routing (blue cols L–AJ) + screenshots (col I)")
    print("  CALENDAR   → one horizontal section per production step")
    print("  PURCHASING → same horizontal layout for outside/vendor steps")
