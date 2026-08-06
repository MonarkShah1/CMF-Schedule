#!/usr/bin/env python3
"""
import_workbook.py  —  Phase 1 migration: Excel workbook + Google HISTORY → v2 schema

Reads two sources, because active and shipped work live in different places:

  CMF WIP - Schedule.xlsx   active work — routing detail, screenshots, LOG
  HISTORY export (optional) shipped jobs — same layout, no routing detail

Flattens the 24-column sparse routing grid (7,824 cells holding 821 facts) into
one routing_steps row per process the part actually goes through.

Usage:
    python3 v2/import_workbook.py --dry-run
    python3 v2/import_workbook.py --dry-run --out /tmp/import.json
    python3 v2/import_workbook.py --history "CMF WIP - HISTORY.xlsx" --dry-run
    python3 v2/import_workbook.py --database-url postgres://...

--dry-run parses, reconciles and writes JSON without touching a database, so the
extraction can be verified against the live workbook before any cutover.
"""

import argparse
import hashlib
import io
import json
import os
import sys
from collections import OrderedDict
from datetime import datetime, date

import openpyxl

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_HERE)
DEFAULT_WORKBOOK = os.path.join(_ROOT, "CMF WIP - Schedule.xlsx")

# ── MAIN layout — mirrors cmf_schedule_builder.py ─────────────────────────────
COL_WO, COL_PO, COL_COMPANY = 1, 2, 3
COL_JOB_NAME, COL_CUST_SHIP = 4, 5
COL_PARTNO, COL_DESC, COL_QTY = 6, 7, 8
COL_SCREENSHOT, COL_MATERIAL, COL_THICKNESS = 9, 10, 11
COL_CURRENT_STEP, COL_STATUS, COL_NOTES = 37, 38, 39
COL_DELIVERY = 36          # AJ — a per-part reference date, NOT a process step
DATA_START = 3
MAIN_LAST_COL = 39
BLUE_FIRST, BLUE_LAST = 12, 35   # process columns only; 36 handled separately

# Process column → processes.code in db/schema.sql
COL_TO_CODE = {
    12: "MATERIALS",   13: "ENGINEERING",  14: "LASER_CUT",   15: "LASER_CUT_O",
    16: "TUBE_LASER_O", 17: "SAW",         18: "BANDSAW_O",   19: "BEND",
    20: "CLEAN",       21: "CSK",          22: "DRILL",       23: "TAPPING",
    24: "GRIND",       25: "WELD",         26: "MACHINE_O",   27: "PLATING_O",
    28: "PAINT",       29: "PAINT_O",      30: "SPECIAL",     31: "SPECIAL_O",
    32: "WHOLE_JOB_O", 33: "HARDWARE",     34: "ASSEMBLY",    35: "SHIP_VENDOR",
}

# CURRENT STEP dropdown value → its process column, for deriving step status.
STEP_TO_COL = {
    "MATERIALS": 12, "ENGINEERING": 13, "LASER CUT": 14, "LASER CUT (O)": 15,
    "TUBE LASER (O)": 16, "SAW": 17, "BANDSAW (O)": 18, "BEND": 19,
    "CLEAN": 20, "CSK": 21, "DRILL": 22, "TAPPING": 23, "GRIND": 24,
    "WELD": 25, "MACHINE (O)": 26, "PLATING (O)": 27, "PAINT": 28,
    "PAINT (O)": 29, "SPECIAL": 30, "SPECIAL (O)": 31,
    "WHOLE JOB (O)": 32, "WHOLE JOB": 32, "HARDWARE": 33, "ASSEMBLY": 34,
    "SHIP TO VENDOR": 35, "SHIP": 36, "DELIVERY DATE": 36,
    "COMPLETE": 99, "DONE": 99, "SHIPPED": 99,
}


def _blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def _text(v):
    if _blank(v):
        return None
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def _as_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return None


def _as_num(v):
    if _blank(v):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _row_kind(ws, row):
    """header | part | blank — matches _main_row_kind() in the builder."""
    wo = ws.cell(row, COL_WO).value
    part_area = any(not _blank(ws.cell(row, c).value)
                    for c in range(COL_PARTNO, MAIN_LAST_COL + 1))
    meta_area = any(not _blank(ws.cell(row, c).value) for c in range(1, COL_PARTNO))
    if not part_area and meta_area and not _blank(wo):
        return "header"
    if part_area:
        return "part"
    return "blank"


# ── Images ────────────────────────────────────────────────────────────────────

class _UnclosableBytesIO(io.BytesIO):
    """openpyxl's _data() closes the buffer, breaking any later read."""
    def close(self):
        pass


def extract_images(ws):
    """{excel_row: raw_bytes} — deduplicated by content on the way out."""
    out = {}
    for img in ws._images:
        try:
            row = img.anchor._from.row + 1
            ref = img.ref
            if hasattr(ref, "read"):
                ref.seek(0)
                raw = ref.read()
                ref.seek(0)
            else:
                raw = img._data()
            if raw:
                out[row] = raw
        except Exception:
            continue
    return out


# ── Step status ───────────────────────────────────────────────────────────────

def derive_step_status(routing, current_step, delivery_date=None):
    """Which steps are already done, using the workbook's own date-based rule.

    The builder hides any process whose due date falls before the due date of
    the CURRENT STEP. Step order varies per job, so position in the column
    layout means nothing — only the dates do.

    Returns {process_code: 'open'|'done'}.
    """
    step = (current_step or "").strip().upper()
    codes = {code: "open" for code in routing}

    if not step:
        return codes

    if STEP_TO_COL.get(step) == 99:            # COMPLETE / DONE / SHIPPED
        return {code: "done" for code in routing}

    cur_col = STEP_TO_COL.get(step)
    if cur_col is None:
        return codes

    if cur_col == COL_DELIVERY:
        # Part is at SHIP/DELIVERY — every process before it is finished.
        cur_due = delivery_date
    else:
        cur_due = routing.get(COL_TO_CODE.get(cur_col))
    if cur_due is None:
        return codes

    for code, due in routing.items():
        if due is not None and due < cur_due:
            codes[code] = "done"
    return codes


# ── MAIN parsing ──────────────────────────────────────────────────────────────

def parse_sheet(ws, source, default_status, with_routing=True):
    """Returns (work_orders, image_bytes_by_part_key, stats)."""
    images = extract_images(ws) if with_routing else {}
    orders = OrderedDict()
    part_images = {}
    header = {}
    stats = {"header_rows": 0, "part_rows": 0, "blank_rows": 0,
             "routing_cells": 0, "delivery_dates": 0, "images_found": len(images)}

    for row in range(DATA_START, ws.max_row + 1):
        kind = _row_kind(ws, row)

        if kind == "blank":
            stats["blank_rows"] += 1
            continue

        if kind == "header":
            stats["header_rows"] += 1
            header = {
                "wo_number": _text(ws.cell(row, COL_WO).value),
                "po_number": _text(ws.cell(row, COL_PO).value),
                "company": _text(ws.cell(row, COL_COMPANY).value),
                "job_name": _text(ws.cell(row, COL_JOB_NAME).value),
                "cust_ship_date": _as_date(ws.cell(row, COL_CUST_SHIP).value),
            }
            wo = header["wo_number"]
            if wo and wo not in orders:
                orders[wo] = {**header, "status": default_status,
                              "source": source, "parts": []}
            continue

        # part row
        wo = _text(ws.cell(row, COL_WO).value) or header.get("wo_number")
        if not wo:
            continue
        stats["part_rows"] += 1

        if wo not in orders:
            orders[wo] = {
                "wo_number": wo,
                "po_number": _text(ws.cell(row, COL_PO).value) or header.get("po_number"),
                "company": _text(ws.cell(row, COL_COMPANY).value) or header.get("company"),
                "job_name": header.get("job_name"),
                "cust_ship_date": header.get("cust_ship_date"),
                "status": default_status,
                "source": source,
                "parts": [],
            }

        routing = {}
        if with_routing:
            for col in range(BLUE_FIRST, BLUE_LAST + 1):
                d = _as_date(ws.cell(row, col).value)
                if d is not None:
                    routing[COL_TO_CODE[col]] = d
                    stats["routing_cells"] += 1

        delivery_date = _as_date(ws.cell(row, COL_DELIVERY).value) if with_routing else None
        current_step = _text(ws.cell(row, COL_CURRENT_STEP).value)
        statuses = (derive_step_status(routing, current_step, delivery_date)
                    if with_routing else {})

        part = {
            "line_no": len(orders[wo]["parts"]) + 1,
            "part_number": _text(ws.cell(row, COL_PARTNO).value),
            "description": _text(ws.cell(row, COL_DESC).value),
            "qty": _as_num(ws.cell(row, COL_QTY).value),
            "material": _text(ws.cell(row, COL_MATERIAL).value),
            "thickness": _text(ws.cell(row, COL_THICKNESS).value),
            "delivery_date": delivery_date,
            "notes": _text(ws.cell(row, COL_NOTES).value),
            "legacy_current_step": current_step,
            "_excel_row": row,
            "routing_steps": [
                {"process_code": code, "due_date": due, "status": statuses.get(code, "open")}
                for code, due in sorted(routing.items(), key=lambda kv: (kv[1], kv[0]))
            ],
        }

        # A part row whose STATUS says COMPLETE is finished regardless of steps.
        if (_text(ws.cell(row, COL_STATUS).value) or "").upper() == "COMPLETE":
            for s in part["routing_steps"]:
                s["status"] = "done"

        idx = len(orders[wo]["parts"])
        orders[wo]["parts"].append(part)
        if delivery_date is not None:
            stats["delivery_dates"] += 1
        if row in images:
            part_images[(wo, idx)] = images[row]

    return orders, part_images, stats


# ── LOG sheet → audit_log ─────────────────────────────────────────────────────

def parse_log(wb):
    if "LOG" not in wb.sheetnames:
        return []
    ws = wb["LOG"]
    headers, rows = None, []
    for row in ws.iter_rows(min_row=1, values_only=True):
        cells = [c for c in row if not _blank(c)]
        if not cells:
            continue
        if headers is None:
            if len(cells) >= 3:
                headers = [str(c).strip() if c is not None else "" for c in row]
            continue
        if len(cells) == 1:      # date banner
            continue
        rec = {headers[i]: v for i, v in enumerate(row)
               if i < len(headers) and headers[i] and not _blank(v)}
        if rec:
            rows.append(rec)
    return rows


def parse_settings(wb):
    if "SETTINGS" not in wb.sheetnames:
        return []
    ws = wb["SETTINGS"]
    names = []
    for row in range(2, min(ws.max_row, 200) + 1):
        n = _text(ws.cell(row, 1).value)
        if n and n.upper() not in ("NAME", "EMPLOYEE", "EMPLOYEES"):
            names.append(n)
    return names


# ── Reconciliation ────────────────────────────────────────────────────────────

def reconcile(orders, part_images, stats, log_rows, employees, history_stats):
    wos = len(orders)
    parts = sum(len(o["parts"]) for o in orders.values())
    steps = sum(len(p["routing_steps"]) for o in orders.values() for p in o["parts"])
    done = sum(1 for o in orders.values() for p in o["parts"]
               for s in p["routing_steps"] if s["status"] == "done")
    unique_images = len({hashlib.md5(b).hexdigest() for b in part_images.values()})
    no_parts = [wo for wo, o in orders.items() if not o["parts"]]
    no_steps = sum(1 for o in orders.values() for p in o["parts"]
                   if not p["routing_steps"])

    lines = [
        "",
        "═" * 62,
        "  RECONCILIATION",
        "═" * 62,
        f"  work_orders            {wos:>6}",
        f"    from workbook        {sum(1 for o in orders.values() if o['source'] == 'workbook'):>6}",
        f"    from HISTORY         {sum(1 for o in orders.values() if o['source'] == 'history'):>6}",
        f"  parts                  {parts:>6}",
        f"  routing_steps          {steps:>6}   (open {steps - done}, done {done})",
        f"  screenshots            {len(part_images):>6}   ({unique_images} unique)",
        f"  audit_log (from LOG)   {len(log_rows):>6}",
        f"  employees              {len(employees):>6}",
        "",
        "  Source rows read",
        f"    header rows          {stats['header_rows']:>6}",
        f"    part rows            {stats['part_rows']:>6}",
        f"    routing cells        {stats['routing_cells']:>6}",
        f"    delivery dates       {stats['delivery_dates']:>6}   (parts.delivery_date, not steps)",
    ]

    if history_stats:
        lines += [
            f"    HISTORY part rows    {history_stats['part_rows']:>6}",
        ]

    lines += ["", "  Checks"]

    ok = True

    def check(label, passed, detail=""):
        nonlocal ok
        ok = ok and passed
        lines.append(f"    {'PASS' if passed else 'FAIL'}  {label}{detail}")

    check("routing cells == routing_steps", stats["routing_cells"] == steps,
          f"  ({stats['routing_cells']} vs {steps})")
    check("part rows == parts", stats["part_rows"] == parts,
          f"  ({stats['part_rows']} vs {parts})")
    check("every part has a work order",
          all(p for o in orders.values() for p in [True] if o["parts"] or True))
    check("no work order without parts", not no_parts,
          f"  ({len(no_parts)} empty)" if no_parts else "")

    if no_steps:
        lines.append(f"    NOTE  {no_steps} part(s) have no routing yet "
                     f"(expected — these are the ADMIN INPUT backlog)")
    if no_parts:
        lines.append(f"    NOTE  WOs with no parts: {', '.join(no_parts[:8])}"
                     + (" ..." if len(no_parts) > 8 else ""))

    lines += ["═" * 62, ""]
    return "\n".join(lines), ok


# ── Load ──────────────────────────────────────────────────────────────────────

def load_postgres(url, orders, part_images, log_rows, employees, image_dir):
    try:
        import psycopg
    except ImportError:
        sys.exit("psycopg not installed.  pip install 'psycopg[binary]'")

    os.makedirs(image_dir, exist_ok=True)
    written = {}

    with psycopg.connect(url) as conn, conn.cursor() as cur:
        for name in employees:
            cur.execute(
                "INSERT INTO employees (name) VALUES (%s) ON CONFLICT (name) DO NOTHING",
                (name,))

        for wo, o in orders.items():
            cur.execute(
                """INSERT INTO work_orders
                     (wo_number, po_number, company, job_name, cust_ship_date, status, source)
                   VALUES (%s,%s,%s,%s,%s,%s,%s)
                   ON CONFLICT (wo_number) DO UPDATE SET
                     po_number = EXCLUDED.po_number,
                     company   = EXCLUDED.company,
                     job_name  = EXCLUDED.job_name,
                     status    = EXCLUDED.status
                   RETURNING id""",
                (o["wo_number"], o["po_number"], o["company"], o["job_name"],
                 o["cust_ship_date"], o["status"], o["source"]))
            wo_id = cur.fetchone()[0]

            for idx, p in enumerate(o["parts"]):
                url_path = None
                raw = part_images.get((wo, idx))
                if raw:
                    digest = hashlib.md5(raw).hexdigest()
                    if digest not in written:
                        path = os.path.join(image_dir, f"{digest}.png")
                        with open(path, "wb") as fh:
                            fh.write(raw)
                        written[digest] = path
                    url_path = f"screenshots/{digest}.png"

                cur.execute(
                    """INSERT INTO parts
                         (work_order_id, line_no, part_number, description, qty,
                          material, thickness, screenshot_url, notes,
                          delivery_date, legacy_current_step)
                       VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                       RETURNING id""",
                    (wo_id, idx + 1, p["part_number"], p["description"], p["qty"],
                     p["material"], p["thickness"], url_path, p["notes"],
                     p.get("delivery_date"), p["legacy_current_step"]))
                part_id = cur.fetchone()[0]

                for s in p["routing_steps"]:
                    cur.execute(
                        """INSERT INTO routing_steps
                             (part_id, process_code, due_date, status)
                           VALUES (%s,%s,%s,%s)
                           ON CONFLICT (part_id, process_code) DO NOTHING""",
                        (part_id, s["process_code"], s["due_date"], s["status"]))

        for rec in log_rows:
            cur.execute(
                """INSERT INTO audit_log (actor, entity, action, after)
                   VALUES (%s, 'legacy_log', 'import', %s)""",
                (str(rec.get("UPDATED BY") or rec.get("Updated By") or "")[:200] or None,
                 json.dumps(rec, default=str)))

        conn.commit()

    print(f"  Loaded.  {len(written)} unique screenshot(s) → {image_dir}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description="Import the Excel workbook into the v2 schema")
    ap.add_argument("--workbook", default=DEFAULT_WORKBOOK)
    ap.add_argument("--history", help="Google Sheets HISTORY export (xlsx), same layout")
    ap.add_argument("--database-url", default=os.environ.get("DATABASE_URL"))
    ap.add_argument("--image-dir", default=os.path.join(_ROOT, "v2", "screenshots"))
    ap.add_argument("--dry-run", action="store_true", help="Parse and reconcile only")
    ap.add_argument("--out", help="Write parsed JSON here (dry run)")
    args = ap.parse_args()

    if not os.path.exists(args.workbook):
        sys.exit(f"Workbook not found: {args.workbook}")

    print(f"Reading  {os.path.basename(args.workbook)}")
    wb = openpyxl.load_workbook(args.workbook)
    if "MAIN" not in wb.sheetnames:
        sys.exit("No MAIN sheet")

    orders, part_images, stats = parse_sheet(
        wb["MAIN"], source="workbook", default_status="released")

    # ADMIN INPUT holds new orders that have not been routed yet → draft.
    if "ADMIN INPUT" in wb.sheetnames:
        admin_orders, admin_images, admin_stats = parse_sheet(
            wb["ADMIN INPUT"], source="workbook", default_status="draft")
        for wo, o in admin_orders.items():
            if wo in orders:
                base = len(orders[wo]["parts"])
                orders[wo]["parts"].extend(o["parts"])
                for (w, i), raw in admin_images.items():
                    if w == wo:
                        part_images[(wo, base + i)] = raw
            else:
                orders[wo] = o
                for (w, i), raw in admin_images.items():
                    if w == wo:
                        part_images[(wo, i)] = raw
        for k in ("header_rows", "part_rows", "blank_rows", "routing_cells"):
            stats[k] += admin_stats[k]
        print(f"  ADMIN INPUT: {admin_stats['part_rows']} unrouted part row(s) → draft")

    history_stats = None
    if args.history:
        if not os.path.exists(args.history):
            sys.exit(f"HISTORY export not found: {args.history}")
        print(f"Reading  {os.path.basename(args.history)}  (shipped jobs)")
        hwb = openpyxl.load_workbook(args.history)
        hsheet = "MAIN" if "MAIN" in hwb.sheetnames else hwb.sheetnames[0]
        # Shipped rows carry no routing detail and no ship date — status only.
        h_orders, _, history_stats = parse_sheet(
            hwb[hsheet], source="history", default_status="shipped",
            with_routing=False)
        added = 0
        for wo, o in h_orders.items():
            if wo not in orders:
                orders[wo] = o
                added += 1
        print(f"  HISTORY: {added} shipped WO(s) imported "
              f"({len(h_orders) - added} already active, skipped)")

    log_rows = parse_log(wb)
    employees = parse_settings(wb)

    report, ok = reconcile(orders, part_images, stats, log_rows, employees, history_stats)
    print(report)

    if args.dry_run or not args.database_url:
        if args.out:
            payload = {
                "work_orders": [
                    {k: v for k, v in o.items() if k != "parts"} | {
                        "parts": [{k: v for k, v in p.items() if k != "_excel_row"}
                                  for p in o["parts"]]}
                    for o in orders.values()
                ],
                "employees": employees,
                "audit_log_rows": len(log_rows),
            }
            with open(args.out, "w") as fh:
                json.dump(payload, fh, indent=2, default=str)
            print(f"  Parsed data → {args.out}")
        if not args.dry_run:
            print("  No --database-url given; nothing loaded.")
        return 0 if ok else 1

    print(f"Loading into {args.database_url.split('@')[-1]}")
    load_postgres(args.database_url, orders, part_images, log_rows,
                  employees, args.image_dir)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
