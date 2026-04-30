#!/usr/bin/env python3
"""
cmf_migrate_main.py  —  Run ONCE
Migrates CMF WIP.xlsx MAIN to the new per-part routing structure.

Output: CMF WIP - Schedule.xlsx

MAIN column layout (38 cols  A–AL):
  A  WO#   B  PO#   C  COMPANY   D  JOB NAME   E  CUST SHIP
  F  PART#  G  DESC  H  QTY      I  SCREENSHOT  J  MATERIAL  K  THICKNESS
  L  MATERIALS        M  ENGINEERING    N  LASER CUT      O  LASER CUT (O)
  P  TUBE LASER (O)  Q  SAW            R  BANDSAW (O)    S  BEND
  T  CLEAN            U  CSK            V  DRILL          W  TAPPING
  X  GRIND            Y  WELD           Z  MACHINE (O)   AA  PLATING (O)
  AB PAINT           AC  PAINT (O)     AD  SPECIAL       AE  SPECIAL (O)
  AF WHOLE JOB (O)  AG  HARDWARE       AH  ASSEMBLY      AI  SHIP DUE
  AJ CURRENT STEP   AK  STATUS         AL  NOTES

Every WO gets:
  • 1 navy WO HEADER row  (boss scans open orders)
  • ≥1 grey PART row      (engineer fills in routing dates + screenshot)
"""

import os, io
from collections import OrderedDict
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.drawing.image import Image as XLImage
from datetime import datetime

SRC  = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CMF WIP.xlsx')
DEST = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CMF WIP - Schedule.xlsx')

# ── Palette ───────────────────────────────────────────────────────────────────
C_NAVY    = "1F3864";  C_GOLD    = "FFD966"
C_STEEL_DK= "1F4E79";  C_BLUE_LT = "DEEAF1"
C_WHITE   = "FFFFFF";  C_LGRAY   = "F2F2F2"
C_DGRAY   = "404040";  C_YELLOW  = "FFF2CC"
C_RED_LT  = "FFD7D7"

def _fill(h): return PatternFill(start_color=h, end_color=h, fill_type="solid")
def _font(bold=False, color=C_DGRAY, size=10): return Font(name="Calibri", bold=bold, color=color, size=size)
def _border(c="CCCCCC"): s=Side(style="thin",color=c); return Border(left=s,right=s,top=s,bottom=s)
def _align(h="center",v="center",wrap=False): return Alignment(horizontal=h,vertical=v,wrap_text=wrap)

# ── Screenshot sizing (in-cell look) ─────────────────────────────────────────
IMG_W      = 85    # pixels
IMG_H      = 48    # pixels
COL_I_W    = 13    # chars  ≈ 91 px
PART_ROW_H = 52    # points ≈ 69 px  (≥ IMG_H for consistent in-cell look)
IMG_PAD_X  = 2
IMG_PAD_Y  = 2

# ── Column definitions (A–AL = 38 columns) ───────────────────────────────────
COLS = [
    # (letter, header, width, is_blue)
    ("A",  "WO #",           8,  False),
    ("B",  "PO #",          17,  False),
    ("C",  "COMPANY",       20,  False),
    ("D",  "JOB NAME",      28,  False),
    ("E",  "CUST SHIP",     12,  False),
    ("F",  "PART #",        16,  False),
    ("G",  "DESCRIPTION",   26,  False),
    ("H",  "QTY",            7,  False),
    ("I",  "SCREENSHOT",    COL_I_W, False),
    ("J",  "MATERIAL",      10,  False),
    ("K",  "THICKNESS",      9,  False),
    # ── Process due dates (compact — width 6 fits "M/D/YY" dates) ──────────
    ("L",  "MATL",         6,  True ),
    ("M",  "ENG",          6,  True ),
    ("N",  "LASER",        6,  True ),
    ("O",  "LASER\n(O)",   6,  True ),
    ("P",  "TUBE\nLSR(O)", 6,  True ),
    ("Q",  "SAW",          6,  True ),
    ("R",  "BAND\n(O)",    6,  True ),
    ("S",  "BEND",         6,  True ),
    ("T",  "CLEAN",        6,  True ),
    ("U",  "CSK",          6,  True ),
    ("V",  "DRILL",        6,  True ),
    ("W",  "TAP",          6,  True ),
    ("X",  "GRIND",        6,  True ),
    ("Y",  "WELD",         6,  True ),
    ("Z",  "MACH\n(O)",    6,  True ),
    ("AA", "PLAT\n(O)",    6,  True ),
    ("AB", "PAINT",        6,  True ),
    ("AC", "PAINT\n(O)",   6,  True ),
    ("AD", "SPEC",         6,  True ),
    ("AE", "SPEC\n(O)",    6,  True ),
    ("AF", "W.JOB\n(O)",   6,  True ),
    ("AG", "HW",           6,  True ),
    ("AH", "ASSM",         6,  True ),
    ("AI", "SHIP TO\nVENDOR", 6, True ),   # new — date when parts ship to vendor
    ("AJ", "DELIVERY\nDATE",  8, True ),   # shifted from AI
    # ── Status ────────────────────────────────────────────────────────────
    ("AK", "CURRENT\nSTEP", 18,  False),   # shifted from AJ
    ("AL", "STATUS",        12,  False),   # shifted from AK
    ("AM", "NOTES",         35,  False),   # shifted from AL
]
NCOLS = len(COLS)     # 39
LAST  = "AM"
HDR   = 2
DATA  = 3

# Blue column range for conditional formatting
BLUE_FIRST = 12   # col L (1-indexed)
BLUE_LAST  = 36   # col AJ (1-indexed, shifted from 35)

STEP_LIST = (
    "MATERIALS,ENGINEERING,LASER CUT,LASER CUT (O),TUBE LASER (O),"
    "SAW,BANDSAW (O),BEND,CLEAN,CSK,DRILL,TAPPING,GRIND,WELD,"
    "MACHINE (O),PLATING (O),PAINT,PAINT (O),"
    "SPECIAL,SPECIAL (O),WHOLE JOB (O),HARDWARE,ASSEMBLY,SHIP TO VENDOR,SHIP,RECEIVING,"
    "COMPLETE,ON HOLD"
)


# ── Image helpers ─────────────────────────────────────────────────────────────
class _UnclosableBytesIO(io.BytesIO):
    """openpyxl _data() calls fp.close() — this no-ops it so the buffer stays readable."""
    def close(self): pass


def copy_image(src_img, dest_ws, col_1idx, row_1idx):
    from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
    try:
        raw = src_img._data()
        buf = _UnclosableBytesIO(raw)
        img = XLImage(buf)
        c0, r0 = col_1idx - 1, row_1idx - 1
        anchor = TwoCellAnchor(editAs='twoCell')
        anchor._from = AnchorMarker(
            col=c0,
            colOff=pixels_to_EMU(IMG_PAD_X),
            row=r0,
            rowOff=pixels_to_EMU(IMG_PAD_Y),
        )
        anchor.to = AnchorMarker(
            col=c0,
            colOff=pixels_to_EMU(IMG_PAD_X + IMG_W),
            row=r0,
            rowOff=pixels_to_EMU(IMG_PAD_Y + IMG_H),
        )
        img.anchor = anchor
        dest_ws.add_image(img)
        return True
    except Exception:
        return False


# ── Sheet setup ───────────────────────────────────────────────────────────────
def setup_sheet(ws):
    ws.row_dimensions[1].height = 30
    ws.merge_cells(f"A1:{LAST}1")
    c = ws["A1"]
    c.value     = f"CMF  WORK IN PROGRESS   —   Updated {datetime.now().strftime('%B %d, %Y')}"
    c.font      = _font(bold=True, color=C_WHITE, size=14)
    c.fill      = _fill(C_NAVY)
    c.alignment = _align()

    ws.row_dimensions[HDR].height = 42
    for i, (letter, label, width, is_blue) in enumerate(COLS, 1):
        ws.column_dimensions[letter].width = width
        c = ws.cell(HDR, i, label)
        c.border    = _border()
        c.alignment = _align(wrap=True)
        if is_blue:
            c.font = _font(bold=True, size=8, color=C_STEEL_DK)
            c.fill = _fill(C_BLUE_LT)
        else:
            c.font = _font(bold=True, size=8, color=C_WHITE)
            c.fill = _fill(C_NAVY)


def write_wo_header(ws, row, wo, po, company, job_name, cust_ship):
    ws.row_dimensions[row].height = 24
    vals = [wo, po, company, job_name, cust_ship] + [None] * (NCOLS - 5)
    for i, val in enumerate(vals, 1):
        c = ws.cell(row, i, val)
        c.fill      = _fill(C_NAVY)
        c.border    = _border("2E4D7B")
        c.alignment = _align(h="left")
        c.font      = _font(bold=True, color=(C_GOLD if i == 1 else C_WHITE), size=11)
        if isinstance(val, datetime): c.number_format = "M/D/YY"


def write_part_row(ws, row, wo, po, company, pno, desc, qty, mat, thk,
                   current, notes, ship_due, alt):
    ws.row_dimensions[row].height = PART_ROW_H
    bg = C_LGRAY if alt else C_WHITE
    # A  B  C       D     E     F    G     H    I(screenshot)  J    K
    # --- 25 process cols L–AJ (idx 12-36) ---    AK   AL   AM
    # Cols A–AM = 39 total.  One None per blue process col that's blank.
    # AI=col 35 = SHIP TO VENDOR (new, blank).  AJ=col 36 = DELIVERY DATE (ship_due).
    vals = [
        wo, po, company, None, None,                              # A–E   (5)
        pno, desc, qty, None, mat, thk,                           # F–K   (6)  I=screenshot
        None,None,None,None,None,None,None,None,                  # L–S   (8)  MATERIALS→BEND
        None,None,None,None,None,None,                            # T–Y   (6)  CLEAN→WELD
        None,None,None,None,None,None,None,None,None,None,ship_due,# Z–AJ (10 N + ship_due = 11 cols)
        current, None, notes,                                     # AK–AM (3)
    ]
    # Verify length == NCOLS at runtime (silent — remove if confident)
    assert len(vals) == NCOLS, f"vals len {len(vals)} != NCOLS {NCOLS}"
    for i, val in enumerate(vals, 1):
        _, _, _, is_blue = COLS[i-1]
        c = ws.cell(row, i, val)
        c.border    = _border()
        c.fill      = _fill(C_BLUE_LT if (is_blue and val is None) else
                            (C_WHITE   if is_blue else bg))
        c.font      = _font(bold=(i in (1,2,3)), size=10)
        c.alignment = _align(h="left" if i in (3,6,7,39) else "center")
        if isinstance(val, datetime): c.number_format = "M/D/YY"


def apply_cf(ws, last_row):
    first_col = get_column_letter(BLUE_FIRST)
    last_col  = get_column_letter(BLUE_LAST)
    rng = f"{first_col}{DATA}:{last_col}{last_row}"
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f"AND({first_col}{DATA}<>\"\",ISNUMBER({first_col}{DATA}),{first_col}{DATA}<TODAY())"],
        fill=_fill(C_RED_LT), font=Font(color="CC0000", bold=True),
    ))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f"AND({first_col}{DATA}<>\"\",ISNUMBER({first_col}{DATA}),{first_col}{DATA}>=TODAY(),{first_col}{DATA}<=TODAY()+3)"],
        fill=_fill(C_YELLOW), font=Font(color="806000", bold=True),
    ))


def apply_dropdown(ws, last_row):
    dv = DataValidation(type="list", formula1=f'"{STEP_LIST}"',
                        allow_blank=True, showErrorMessage=False)
    dv.sqref = f"AK{DATA}:AK{last_row}"   # CURRENT STEP shifted to AK
    ws.add_data_validation(dv)


# ── Migration ─────────────────────────────────────────────────────────────────
def migrate():
    print(f"Reading  {SRC}")
    src    = openpyxl.load_workbook(SRC)
    old_ws = src["MAIN"]

    # Build old-row → image map (0-indexed anchor row → 1-indexed Excel row)
    old_row_to_img = {}
    for img in old_ws._images:
        excel_row = img.anchor._from.row + 1
        old_row_to_img[excel_row] = img
    print(f"  Found {len(old_row_to_img)} screenshots in old MAIN")

    # Pre-group rows by WO (preserve insertion order)
    wo_groups = OrderedDict()
    in_excel_row = 2
    for row in old_ws.iter_rows(min_row=2, values_only=True):
        wo = row[8]
        if wo:
            if wo not in wo_groups:
                wo_groups[wo] = []
            wo_groups[wo].append((in_excel_row, row))
        in_excel_row += 1
    print(f"  Found {len(wo_groups)} WOs in old MAIN")

    dest = openpyxl.Workbook()
    dest.remove(dest.active)

    ws = dest.create_sheet("MAIN", 0)
    setup_sheet(ws)

    out_row    = DATA
    wo_count   = 0
    part_count = 0
    img_count  = 0
    alt        = False

    for wo, rows in wo_groups.items():
        wo_count += 1

        # Detect if the first row is a true summary (no steps, no current) or a part
        first_er, first_row = rows[0]
        first_is_summary = (first_row[0] is None and first_row[1] is None)

        # Determine header metadata
        if first_is_summary:
            job_name = first_row[10] or first_row[9]
        else:
            job_name = None    # no explicit summary; header just shows WO-level info

        company  = first_row[6]
        po       = first_row[7]
        due      = first_row[3]
        if job_name: job_name = str(job_name).strip()

        # ── Write navy WO header ───────────────────────────────────────────
        write_wo_header(ws, out_row, wo, po, company, job_name, due)
        out_row += 1

        # ── Determine part rows ────────────────────────────────────────────
        if first_is_summary:
            part_rows = rows[1:]           # header was first row; rest are parts
        else:
            part_rows = rows               # ALL rows are parts (no summary row)

        if not part_rows:
            # Only had a summary row, no parts → create one blank part row
            part_rows = [(first_er, first_row)]

        # ── Write grey part rows ───────────────────────────────────────────
        for er, row in part_rows:
            part_count += 1
            alt = not alt

            # For single-row WOs the first row data doubles as the part
            pno  = row[9]
            desc = row[10]
            qty  = row[11]
            mat  = row[14]
            thk  = row[13]
            curr = row[1]     # CURRENT STEP
            notes= row[0]     # old ALL STEPS → NOTES reference for engineer
            ship = row[3]     # old DUE DATE → pre-fill SHIP DUE

            write_part_row(ws, out_row, wo, po, company,
                           pno, desc, qty, mat, thk,
                           curr, notes, ship, alt)

            # Copy screenshot (consistent size — in-cell look)
            if er in old_row_to_img:
                if copy_image(old_row_to_img[er], ws, 9, out_row):
                    img_count += 1

            out_row += 1

    apply_cf(ws, out_row - 1)
    apply_dropdown(ws, out_row - 1)
    ws.freeze_panes = f"A{DATA}"
    ws.auto_filter.ref = f"A{HDR}:{LAST}{out_row - 1}"

    print(f"  {wo_count} WOs, {part_count} parts, {img_count} screenshots migrated")
    print(f"Saving   {DEST}")
    dest.save(DEST)
    print("Done ✓\n")
    print("Next steps:")
    print("  1. Open 'CMF WIP - Schedule.xlsx' and review MAIN")
    print("  2. Engineer fills in blue process-due columns (L–AI) for each part")
    print("  3. Run:  python3 cmf_schedule_builder.py")


if __name__ == "__main__":
    migrate()
