#!/usr/bin/env python3
"""
cmf_merge.py  —  Run ONCE to merge boss's updates with engineer's routing
Inputs:
  CMF WIP (1).xlsx           ← boss's latest orders/parts (old format)
  CMF WIP - Schedule (1).xlsx ← engineer's routing/process dates (new format)
Output:
  CMF WIP - Schedule.xlsx    ← merged result (overwrites current working file)

Matching logic:
  Primary:   WO# + Part#  (exact match)
  Secondary: WO# only     (for WO-level routing with no specific part#)
"""

import os, io
from collections import OrderedDict
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.drawing.image import Image as XLImage
from datetime import datetime

_DIR      = os.path.dirname(os.path.abspath(__file__))
WIP_FILE  = os.path.join(_DIR, 'CMF WIP (1).xlsx')
SCH_FILE  = os.path.join(_DIR, 'CMF WIP - Schedule (1).xlsx')
OUT_FILE  = os.path.join(_DIR, 'CMF WIP - Schedule.xlsx')

# ── Palette ───────────────────────────────────────────────────────────────────
C_NAVY="1F3864"; C_GOLD="FFD966"; C_STEEL_DK="1F4E79"; C_BLUE_LT="DEEAF1"
C_WHITE="FFFFFF"; C_LGRAY="F2F2F2"; C_DGRAY="404040"
C_YELLOW="FFF2CC"; C_RED_LT="FFD7D7"

def _fill(h): return PatternFill(start_color=h,end_color=h,fill_type="solid")
def _font(bold=False,color=C_DGRAY,size=10): return Font(name="Calibri",bold=bold,color=color,size=size)
def _border(c="CCCCCC"): s=Side(style="thin",color=c); return Border(left=s,right=s,top=s,bottom=s)
def _align(h="center",v="center",wrap=False): return Alignment(horizontal=h,vertical=v,wrap_text=wrap)

IMG_W=85; IMG_H=48; COL_I_W=13; PART_ROW_H=52

COLS = [
    ("A","WO #",8,False),("B","PO #",17,False),("C","COMPANY",20,False),
    ("D","JOB NAME",28,False),("E","CUST SHIP",12,False),
    ("F","PART #",16,False),("G","DESCRIPTION",26,False),
    ("H","QTY",7,False),("I","SCREENSHOT",COL_I_W,False),
    ("J","MATERIAL",10,False),("K","THICKNESS",9,False),
    ("L","MATL",6,True),("M","ENG",6,True),
    ("N","LASER",6,True),("O","LASER\n(O)",6,True),
    ("P","TUBE\nLSR(O)",6,True),("Q","SAW",6,True),
    ("R","BAND\n(O)",6,True),("S","BEND",6,True),
    ("T","CLEAN",6,True),("U","CSK",6,True),
    ("V","DRILL",6,True),("W","TAP",6,True),
    ("X","GRIND",6,True),("Y","WELD",6,True),
    ("Z","MACH\n(O)",6,True),("AA","PLAT\n(O)",6,True),
    ("AB","PAINT",6,True),("AC","PAINT\n(O)",6,True),
    ("AD","SPEC",6,True),("AE","SPEC\n(O)",6,True),
    ("AF","W.JOB\n(O)",6,True),("AG","HW",6,True),
    ("AH","ASSM",6,True),("AI","DELIVERY\nDATE",8,True),
    ("AJ","CURRENT\nSTEP",18,False),("AK","STATUS",12,False),
    ("AL","NOTES",35,False),
]
NCOLS=len(COLS); LAST="AL"; HDR=2; DATA=3
BLUE_FIRST=12; BLUE_LAST=35

STEP_LIST=(
    "MATERIALS,ENGINEERING,LASER CUT,LASER CUT (O),TUBE LASER (O),"
    "SAW,BANDSAW (O),BEND,CLEAN,CSK,DRILL,TAPPING,GRIND,WELD,"
    "MACHINE (O),PLATING (O),PAINT,PAINT (O),"
    "SPECIAL,SPECIAL (O),WHOLE JOB (O),HARDWARE,ASSEMBLY,SHIP,RECEIVING,"
    "COMPLETE,ON HOLD"
)


# ── Image helpers ─────────────────────────────────────────────────────────────
class _UnclosableBytesIO(io.BytesIO):
    def close(self): pass

def _copy_image(src_img, dest_ws, col_1idx, row_1idx):
    from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
    try:
        raw=src_img._data(); buf=_UnclosableBytesIO(raw); img=XLImage(buf)
        c0,r0=col_1idx-1,row_1idx-1
        anchor=TwoCellAnchor(); anchor.editAs='twoCell'
        anchor._from=AnchorMarker(col=c0,colOff=0,row=r0,rowOff=0)
        anchor.to=AnchorMarker(col=c0+1,colOff=0,row=r0+1,rowOff=0)
        img.anchor=anchor; dest_ws.add_image(img); return True
    except Exception: return False


def _norm_key(val):
    """Normalise a WO# or part# to a consistent string for matching."""
    if val is None: return None
    s = str(val).strip()
    # If it looks like a float (e.g. "4418.0"), convert to int string
    try:
        f = float(s)
        if f == int(f): s = str(int(f))
    except ValueError: pass
    return s or None


# ── Build routing lookup from Schedule(1) ────────────────────────────────────
def build_routing_lookup(ws_sch):
    """
    Returns:
      by_part: {(wo_str, part_str): {col: date}}   ← exact WO+Part match
      by_wo:   {wo_str: {col: date}}                ← WO-level (no part#)
    """
    by_part = {}
    by_wo   = {}

    for row in range(DATA, ws_sch.max_row + 1):
        wo_raw  = ws_sch.cell(row, 1).value
        pno_raw = ws_sch.cell(row, 6).value
        if not wo_raw: continue

        routing = {}
        for col in range(12, 36):
            v = ws_sch.cell(row, col).value
            if v and isinstance(v, datetime):
                routing[col] = v

        if not routing: continue

        wo_k  = _norm_key(wo_raw)
        pno_k = _norm_key(pno_raw)

        if pno_k:
            # Merge (later rows win for same key)
            key = (wo_k, pno_k)
            existing = by_part.get(key, {})
            existing.update(routing)
            by_part[key] = existing
        else:
            # WO-level routing
            existing = by_wo.get(wo_k, {})
            existing.update(routing)
            by_wo[wo_k] = existing

    print(f"  Routing lookup: {len(by_part)} part-level + {len(by_wo)} WO-level entries")
    return by_part, by_wo


def get_routing(wo, pno, by_part, by_wo):
    """Return merged routing dict for this WO+Part (part-level wins over WO-level)."""
    wo_k  = _norm_key(wo)
    pno_k = _norm_key(pno)
    result = dict(by_wo.get(wo_k, {}))           # start with WO-level
    if pno_k:
        result.update(by_part.get((wo_k, pno_k), {}))  # part-level overrides
    return result


# ── Sheet setup (same as migrate script) ─────────────────────────────────────
def _setup_sheet(ws):
    ws.row_dimensions[1].height = 30
    ws.merge_cells(f"A1:{LAST}1")
    c=ws["A1"]
    c.value=f"CMF  WORK IN PROGRESS   —   Updated {datetime.now().strftime('%B %d, %Y')}"
    c.font=_font(bold=True,color=C_WHITE,size=14)
    c.fill=_fill(C_NAVY); c.alignment=_align()

    ws.row_dimensions[HDR].height = 42
    for i,(letter,label,width,is_blue) in enumerate(COLS,1):
        ws.column_dimensions[letter].width = width
        c=ws.cell(HDR,i,label); c.border=_border(); c.alignment=_align(wrap=True)
        if is_blue: c.font=_font(bold=True,size=8,color=C_STEEL_DK); c.fill=_fill(C_BLUE_LT)
        else:       c.font=_font(bold=True,size=8,color=C_WHITE);     c.fill=_fill(C_NAVY)


def _write_header(ws, row, wo, po, company, job_name, cust_ship):
    ws.row_dimensions[row].height = 24
    vals=[wo,po,company,job_name,cust_ship]+[None]*(NCOLS-5)
    for i,val in enumerate(vals,1):
        c=ws.cell(row,i,val); c.fill=_fill(C_NAVY); c.border=_border("2E4D7B")
        c.alignment=_align(h="left")
        c.font=_font(bold=True,color=(C_GOLD if i==1 else C_WHITE),size=11)
        if isinstance(val,datetime): c.number_format="M/D/YY"


def _write_part(ws, row, wo, po, company, pno, desc, qty, mat, thk,
                current, notes, routing, alt):
    ws.row_dimensions[row].height = PART_ROW_H
    bg = C_LGRAY if alt else C_WHITE

    # Build vals: 38 cols total
    # A-E(5), F-K(6), L-S(8), T-Y(6), Z-AI(10 = 9N+delivery), AJ-AL(3)
    vals = [
        wo, po, company, None, None,                    # A-E
        pno, desc, qty, None, mat, thk,                 # F-K  (I=screenshot)
        None,None,None,None,None,None,None,None,        # L-S  (8)
        None,None,None,None,None,None,                  # T-Y  (6)
        None,None,None,None,None,None,None,None,None,None,   # Z-AI (10, delivery at pos 35)
        current, None, notes,                           # AJ-AL
    ]
    assert len(vals)==NCOLS, f"vals len {len(vals)} != {NCOLS}"

    # Apply routing dates into the correct positions
    for col_idx, date_val in routing.items():
        list_pos = col_idx - 1   # 1-indexed col → 0-indexed list position
        if 0 <= list_pos < NCOLS:
            vals[list_pos] = date_val

    for i,val in enumerate(vals,1):
        _,_,_,is_blue=COLS[i-1]
        c=ws.cell(row,i,val); c.border=_border()
        c.fill=_fill(C_BLUE_LT if (is_blue and val is None) else
                     (C_WHITE  if is_blue else bg))
        c.font=_font(bold=(i in (1,2,3)),size=10)
        c.alignment=_align(h="left" if i in (3,6,7,38) else "center")
        if isinstance(val,datetime): c.number_format="M/D/YY"


def _apply_cf(ws, last_row):
    fc=get_column_letter(BLUE_FIRST); lc=get_column_letter(BLUE_LAST)
    rng=f"{fc}{DATA}:{lc}{last_row}"
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f"AND({fc}{DATA}<>\"\",ISNUMBER({fc}{DATA}),{fc}{DATA}<TODAY())"],
        fill=_fill(C_RED_LT),font=Font(color="CC0000",bold=True)))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f"AND({fc}{DATA}<>\"\",ISNUMBER({fc}{DATA}),{fc}{DATA}>=TODAY(),{fc}{DATA}<=TODAY()+3)"],
        fill=_fill(C_YELLOW),font=Font(color="806000",bold=True)))


def _apply_dropdown(ws, last_row):
    dv=DataValidation(type="list",formula1=f'"{STEP_LIST}"',allow_blank=True,showErrorMessage=False)
    dv.sqref=f"AJ{DATA}:AJ{last_row}"; ws.add_data_validation(dv)


# ── Main merge ────────────────────────────────────────────────────────────────
def merge():
    print(f"Loading WIP(1):      {WIP_FILE}")
    wb_wip = openpyxl.load_workbook(WIP_FILE)
    ws_wip = wb_wip['MAIN']

    print(f"Loading Schedule(1): {SCH_FILE}")
    wb_sch = openpyxl.load_workbook(SCH_FILE)
    ws_sch = wb_sch['MAIN']

    # Build image maps from BOTH sources
    # WIP(1) images are primary (more up-to-date), by old_excel_row
    wip_img_map = {}
    for img in ws_wip._images:
        try: wip_img_map[img.anchor._from.row + 1] = img
        except Exception: pass

    # Schedule(1) images keyed by (WO#, Part#) for secondary lookup
    sch_img_by_key = {}
    for row in range(DATA, ws_sch.max_row + 1):
        if not ws_sch._images: break
    for img in ws_sch._images:
        try:
            er = img.anchor._from.row + 1
            wo  = _norm_key(ws_sch.cell(er, 1).value)
            pno = _norm_key(ws_sch.cell(er, 6).value)
            if wo: sch_img_by_key[(wo, pno)] = img
        except Exception: pass

    print(f"  WIP images: {len(wip_img_map)}  |  Schedule images: {len(sch_img_by_key)}")

    # Build routing lookup from Schedule(1)
    print("Building routing lookup ...")
    by_part, by_wo = build_routing_lookup(ws_sch)

    # Pre-group WIP(1) rows by WO (preserve order)
    wo_groups = OrderedDict()
    in_row_num = 2
    for row_data in ws_wip.iter_rows(min_row=2, values_only=True):
        wo = row_data[8]    # old col I = WO NO (0-indexed 8 = 1-indexed 9)
        if wo:
            if wo not in wo_groups: wo_groups[wo] = []
            wo_groups[wo].append((in_row_num, row_data))
        in_row_num += 1

    print(f"  {len(wo_groups)} WOs from WIP(1)")

    # Build output workbook
    dest = openpyxl.Workbook()
    dest.remove(dest.active)
    ws_out = dest.create_sheet("MAIN", 0)
    _setup_sheet(ws_out)

    out_row = DATA
    wo_count = part_count = img_count = routing_count = 0
    alt = False

    for wo, rows in wo_groups.items():
        wo_count += 1
        first_er, first_row = rows[0]
        first_is_summary = (first_row[0] is None and first_row[1] is None)

        # Header metadata
        job_name = (first_row[10] or first_row[9]) if first_is_summary else None
        if job_name: job_name = str(job_name).strip()
        company  = first_row[6]; po = first_row[7]; due = first_row[3]

        # ── Navy header row ───────────────────────────────────────────────
        _write_header(ws_out, out_row, wo, po, company, job_name, due)
        out_row += 1

        # ── Determine part rows ───────────────────────────────────────────
        part_rows = rows[1:] if first_is_summary else rows
        if not part_rows:
            part_rows = [(first_er, first_row)]   # single-row WO → also a part

        # ── Write part rows ───────────────────────────────────────────────
        for er, row_data in part_rows:
            part_count += 1
            alt = not alt

            # Old format col mapping (1-indexed in WIP, 0-indexed in row_data tuple)
            pno     = row_data[9]    # col J = PART NO (description for some WOs)
            desc    = row_data[10]   # col K = PART NAME (part# for some WOs)
            qty     = row_data[11]
            mat     = row_data[14]
            thk     = row_data[13]
            current = row_data[1]    # col B = CURRENT
            notes   = row_data[0]    # col A = ALL STEPS → reference notes
            ship    = row_data[3]    # col D = DUE DATE → delivery date

            # Look up routing from Schedule(1)
            routing = get_routing(wo, pno, by_part, by_wo)

            # If no routing has delivery date, pre-fill from old due date
            if 35 not in routing and ship:
                routing[35] = ship

            if routing: routing_count += 1

            _write_part(ws_out, out_row, wo, po, company,
                        pno, desc, qty, mat, thk,
                        current, notes, routing, alt)

            # Screenshot: prefer WIP(1) source (most up-to-date)
            img_placed = False
            if er in wip_img_map:
                if _copy_image(wip_img_map[er], ws_out, 9, out_row):
                    img_count += 1; img_placed = True

            # Fallback: check Schedule(1) by key
            if not img_placed:
                wo_k  = _norm_key(wo)
                pno_k = _norm_key(pno)
                for key in [(wo_k, pno_k), (wo_k, None)]:
                    if key in sch_img_by_key:
                        if _copy_image(sch_img_by_key[key], ws_out, 9, out_row):
                            img_count += 1; break

            out_row += 1

    _apply_cf(ws_out, out_row - 1)
    _apply_dropdown(ws_out, out_row - 1)
    ws_out.freeze_panes = f"A{DATA}"
    ws_out.auto_filter.ref = f"A{HDR}:{LAST}{out_row - 1}"

    print(f"\n  Result: {wo_count} WOs | {part_count} parts | "
          f"{routing_count} parts with routing | {img_count} screenshots")
    print(f"Saving → {OUT_FILE}")
    dest.save(OUT_FILE)
    print("Done ✓\n")
    print("Next: run  python3 cmf_schedule_builder.py  to rebuild CALENDAR + TODAY")


if __name__ == "__main__":
    merge()
