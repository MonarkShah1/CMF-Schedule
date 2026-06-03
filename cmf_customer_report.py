"""
CMF Customer Evaluation Report
Generates a PDF summarizing open work and relationship strength per customer.
"""

import openpyxl
from datetime import datetime, date
from collections import defaultdict
import warnings
warnings.filterwarnings("ignore")

from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer,
    HRFlowable, KeepTogether
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.graphics.shapes import Drawing, Rect, String, Line
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics import renderPDF


# ── Colour palette ───────────────────────────────────────────────────────────
CMF_NAVY   = colors.HexColor("#1B2A4A")
CMF_BLUE   = colors.HexColor("#2563EB")
CMF_LIGHT  = colors.HexColor("#EFF6FF")
CMF_ACCENT = colors.HexColor("#F59E0B")
CMF_GREEN  = colors.HexColor("#16A34A")
CMF_RED    = colors.HexColor("#DC2626")
CMF_ORANGE = colors.HexColor("#EA580C")
CMF_GREY   = colors.HexColor("#F1F5F9")
CMF_MUTED  = colors.HexColor("#64748B")
WHITE      = colors.white
BLACK      = colors.black

TODAY = datetime.today().date()


# ── Data extraction ───────────────────────────────────────────────────────────

def normalize(name):
    return name.strip().upper() if name else ""

def parse_date(val):
    if isinstance(val, (datetime,)):
        return val.date()
    if isinstance(val, date):
        return val
    if isinstance(val, str):
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%m/%d"):
            try:
                d = datetime.strptime(val.strip(), fmt)
                if d.year == 1900:
                    d = d.replace(year=2026)
                return d.date()
            except ValueError:
                pass
    return None


def load_main(ws):
    """Return per-customer aggregates from the MAIN sheet."""
    customers = defaultdict(lambda: {
        "open_pos": set(),
        "open_line_items": 0,
        "open_qty": 0,
        "overdue_parts": 0,
        "delivery_dates": [],
        "process_cols_used": [],
    })

    PROCESS_COLS = list(range(11, 35))   # cols L–AJ (0-indexed 11–34)
    DELIVERY_COL = 35                    # AJ
    STATUS_COL   = 37                    # AL

    current_header = {}

    for row in ws.iter_rows(min_row=3, values_only=True):
        wo      = row[0]
        po      = row[1]
        company = normalize(row[2])
        part    = row[5]
        qty     = row[7]
        status  = row[STATUS_COL] if len(row) > STATUS_COL else None
        delivery = parse_date(row[DELIVERY_COL]) if len(row) > DELIVERY_COL else None

        if not company:
            continue

        # Header row (no part number)
        if not part:
            current_header[company] = {"wo": wo, "po": po}
            continue

        # Part row
        if status and str(status).strip().upper() == "COMPLETE":
            continue  # skip completed parts

        customers[company]["open_pos"].add(wo)
        customers[company]["open_line_items"] += 1

        if isinstance(qty, (int, float)):
            customers[company]["open_qty"] += int(qty)

        if delivery:
            customers[company]["delivery_dates"].append(delivery)
            # Count overdue parts: delivery date already passed
            if delivery < TODAY:
                customers[company]["overdue_parts"] += 1

        # Count how many processes are planned (non-empty process cols)
        steps = sum(1 for c in PROCESS_COLS if len(row) > c and row[c] is not None)
        customers[company]["process_cols_used"].append(steps)

    return customers


def load_log(ws):
    """Return per-customer LOG metrics."""
    log = defaultdict(lambda: {"completions": 0, "on_time": 0, "late": 0, "issues": 0})

    for row in ws.iter_rows(min_row=3, values_only=True):
        date_val  = row[0]
        company   = normalize(row[2])
        was_due   = row[6]
        pm_review = row[11]

        if not company or not date_val:
            continue
        if isinstance(date_val, str) and any(m in date_val for m in
                ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]):
            continue

        comp_date = parse_date(date_val)
        if not comp_date:
            continue

        log[company]["completions"] += 1

        if pm_review == "ISSUE":
            log[company]["issues"] += 1

        due_date = parse_date(was_due)
        if due_date:
            if due_date.year == 1900:
                due_date = due_date.replace(year=2026)
            if comp_date <= due_date:
                log[company]["on_time"] += 1
            else:
                log[company]["late"] += 1

    return log


def build_customer_table(main, log):
    """Merge main + log data into a list of customer dicts, sorted by priority score."""
    all_companies = set(main.keys()) | set(log.keys())
    rows = []

    for company in all_companies:
        m = main.get(company, {})
        l = log.get(company, {"completions": 0, "on_time": 0, "late": 0, "issues": 0})

        open_pos   = len(m.get("open_pos", set()))
        open_items = m.get("open_line_items", 0)
        open_qty   = m.get("open_qty", 0)
        overdue    = m.get("overdue_parts", 0)
        history    = l["completions"]
        issues     = l["issues"]

        rated = l["on_time"] + l["late"]
        on_time_pct = round(l["on_time"] / rated * 100) if rated > 0 else None

        dates = sorted(m.get("delivery_dates", []))
        nearest = dates[0] if dates else None
        days_out = (nearest - TODAY).days if nearest else None

        # Priority score (higher = more valuable / urgent)
        score = (
            open_pos   * 15 +
            open_items * 8  +
            min(open_qty, 5000) / 100 +
            history    * 5  +
            (on_time_pct or 50) * 0.3 -
            issues     * 10 -
            overdue    * 12
        )

        rows.append({
            "company":    company,
            "open_pos":   open_pos,
            "open_items": open_items,
            "open_qty":   open_qty,
            "overdue":    overdue,
            "nearest":    nearest,
            "days_out":   days_out,
            "history":    history,
            "on_time_pct": on_time_pct,
            "issues":     issues,
            "score":      round(score, 1),
        })

    # Sort by score descending; only show customers with at least some activity
    rows = [r for r in rows if r["open_pos"] > 0 or r["history"] > 0]
    rows.sort(key=lambda r: r["score"], reverse=True)
    return rows


# ── PDF builder ───────────────────────────────────────────────────────────────

def score_color(score):
    if score >= 120:  return CMF_GREEN
    if score >= 60:   return CMF_BLUE
    if score >= 0:    return CMF_ACCENT
    return CMF_RED

def traffic_light(val, thresholds, colors_list):
    """Return a color based on thresholds (ascending)."""
    for t, c in zip(thresholds, colors_list):
        if val <= t:
            return c
    return colors_list[-1]


def make_header_drawing(width):
    """Return a ReportLab Drawing for the page header banner."""
    h = 72
    d = Drawing(width, h)

    # Background
    d.add(Rect(0, 0, width, h, fillColor=CMF_NAVY, strokeColor=None))

    # Accent bar
    d.add(Rect(0, 0, 6, h, fillColor=CMF_ACCENT, strokeColor=None))

    # Title
    d.add(String(22, 40, "CMF MANUFACTURING",
                 fontName="Helvetica-Bold", fontSize=22,
                 fillColor=WHITE))
    d.add(String(22, 20, "Customer Evaluation Report  ·  As of " + TODAY.strftime("%B %d, %Y"),
                 fontName="Helvetica", fontSize=11,
                 fillColor=colors.HexColor("#CBD5E1")))

    # Right-side note
    note = "CONFIDENTIAL — INTERNAL USE ONLY"
    d.add(String(width - 14, 28, note,
                 fontName="Helvetica-Oblique", fontSize=9,
                 fillColor=colors.HexColor("#94A3B8"),
                 textAnchor="end"))
    return d


def make_bar_chart(rows, width, height):
    """Horizontal summary bar chart of Priority Scores."""
    # Take top 10
    top = rows[:10]
    names  = [r["company"][:18] for r in top]
    scores = [max(r["score"], 0) for r in top]

    d = Drawing(width, height)

    # Light background
    d.add(Rect(0, 0, width, height, fillColor=CMF_GREY, strokeColor=None))
    d.add(String(width / 2, height - 18,
                 "Top Customers by Priority Score",
                 fontName="Helvetica-Bold", fontSize=11,
                 fillColor=CMF_NAVY, textAnchor="middle"))

    if not scores or max(scores) == 0:
        d.add(String(width / 2, height / 2, "No score data",
                     fontName="Helvetica", fontSize=10,
                     fillColor=CMF_MUTED, textAnchor="middle"))
        return d

    chart_left   = 145
    chart_bottom = 22
    chart_width  = width - chart_left - 16
    chart_height = height - 46

    bar_height = chart_height / len(scores)
    max_score  = max(scores) * 1.1 or 1

    for i, (name, score) in enumerate(zip(names, scores)):
        y = chart_bottom + (len(scores) - 1 - i) * bar_height
        bar_w = (score / max_score) * chart_width

        # Bar
        bar_color = score_color(score)
        d.add(Rect(chart_left, y + 2, bar_w, bar_height - 6,
                   fillColor=bar_color, strokeColor=None))

        # Customer label
        d.add(String(chart_left - 6, y + bar_height / 2 - 4,
                     name, fontName="Helvetica", fontSize=8,
                     fillColor=CMF_NAVY, textAnchor="end"))

        # Score label
        d.add(String(chart_left + bar_w + 4, y + bar_height / 2 - 4,
                     str(int(score)), fontName="Helvetica-Bold", fontSize=8,
                     fillColor=CMF_NAVY))

    return d


def build_pdf(output_path, data_rows):
    page_w, page_h = landscape(letter)
    margin = 0.45 * inch

    doc = SimpleDocTemplate(
        output_path,
        pagesize=landscape(letter),
        leftMargin=margin, rightMargin=margin,
        topMargin=margin, bottomMargin=margin,
        title="CMF Customer Evaluation Report",
    )

    styles = getSampleStyleSheet()
    body_w = page_w - 2 * margin

    story = []

    # ── Header ────────────────────────────────────────────────────────────────
    story.append(make_header_drawing(body_w))
    story.append(Spacer(1, 10))

    # ── Intro paragraph ───────────────────────────────────────────────────────
    intro_style = ParagraphStyle("intro",
        fontName="Helvetica", fontSize=9.5,
        textColor=CMF_MUTED, leading=14, spaceAfter=8)
    story.append(Paragraph(
        "This report ranks all active customers by a composite <b>Priority Score</b> that weighs open work volume, "
        "relationship depth (historical completions), on-time delivery performance, and current risk (overdue parts, issues logged). "
        "Use it to identify which customers to <b>go all-in on</b> for premium service focus.",
        intro_style))

    story.append(HRFlowable(width="100%", thickness=1, color=CMF_BLUE, spaceAfter=8))

    # ── Bar chart ─────────────────────────────────────────────────────────────
    chart_h = 170
    story.append(make_bar_chart(data_rows, body_w, chart_h))
    story.append(Spacer(1, 12))

    # ── Main table ────────────────────────────────────────────────────────────
    col_labels = [
        "RANK", "CUSTOMER", "OPEN\nPOs", "OPEN\nLINE ITEMS",
        "OPEN\nQTY", "OVERDUE\nPARTS", "NEAREST\nDELIVERY",
        "DAYS\nOUT", "HIST.\nCOMPL.", "ON-TIME\nRATE",
        "ISSUES\nLOGGED", "PRIORITY\nSCORE"
    ]

    # Column widths (must sum ≤ body_w)
    col_widths = [
        0.42*inch,   # rank
        1.55*inch,   # customer
        0.52*inch,   # open pos
        0.62*inch,   # open line items
        0.55*inch,   # open qty
        0.60*inch,   # overdue
        0.78*inch,   # nearest delivery
        0.52*inch,   # days out
        0.52*inch,   # hist compl
        0.60*inch,   # on-time
        0.55*inch,   # issues
        0.62*inch,   # score
    ]

    header_style = ParagraphStyle("th",
        fontName="Helvetica-Bold", fontSize=7.5,
        textColor=WHITE, leading=9, alignment=TA_CENTER)
    cell_style   = ParagraphStyle("td",
        fontName="Helvetica", fontSize=8.5,
        textColor=CMF_NAVY, leading=10, alignment=TA_CENTER)
    name_style   = ParagraphStyle("name",
        fontName="Helvetica-Bold", fontSize=8.5,
        textColor=CMF_NAVY, leading=10, alignment=TA_LEFT)
    score_bold   = ParagraphStyle("score",
        fontName="Helvetica-Bold", fontSize=9,
        leading=10, alignment=TA_CENTER)

    table_data = [[Paragraph(h, header_style) for h in col_labels]]
    cmd = [
        # Header row styling
        ("BACKGROUND",  (0, 0), (-1, 0), CMF_NAVY),
        ("ROWHEIGHT",   (0, 0), (-1, 0), 26),
        ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN",       (0, 0), (-1, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING",(0, 0), (-1, -1), 4),
        ("TOPPADDING",  (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING",(0, 0),(-1, -1), 3),
        ("GRID",        (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E1")),
        ("LINEBELOW",   (0, 0), (-1, 0),  1.5, CMF_ACCENT),
    ]

    for rank, r in enumerate(data_rows, 1):
        # Format cells
        nearest_str = r["nearest"].strftime("%b %d") if r["nearest"] else "—"

        days = r["days_out"]
        if days is None:
            days_str = "—"
            days_color = CMF_MUTED
        elif days < 0:
            days_str = f"{abs(days)}d LATE"
            days_color = CMF_RED
        elif days <= 3:
            days_str = f"{days}d"
            days_color = CMF_ORANGE
        elif days <= 7:
            days_str = f"{days}d"
            days_color = CMF_ACCENT
        else:
            days_str = f"{days}d"
            days_color = CMF_GREEN

        on_time = f"{r['on_time_pct']}%" if r["on_time_pct"] is not None else "N/A"
        ot_color = (CMF_GREEN if (r["on_time_pct"] or 0) >= 70
                    else CMF_ORANGE if (r["on_time_pct"] or 0) >= 40
                    else CMF_RED if r["on_time_pct"] is not None
                    else CMF_MUTED)

        overdue_color = CMF_RED if r["overdue"] > 0 else CMF_NAVY
        issues_color  = CMF_RED if r["issues"] > 0 else CMF_NAVY

        sc = r["score"]
        sc_color = score_color(sc)

        days_para   = ParagraphStyle("d", parent=cell_style, textColor=days_color, fontName="Helvetica-Bold")
        ot_para     = ParagraphStyle("o", parent=cell_style, textColor=ot_color,   fontName="Helvetica-Bold")
        over_para   = ParagraphStyle("v", parent=cell_style, textColor=overdue_color)
        issue_para  = ParagraphStyle("i", parent=cell_style, textColor=issues_color)
        sc_para     = ParagraphStyle("s", parent=score_bold, textColor=sc_color)

        row_bg = CMF_GREY if rank % 2 == 0 else WHITE

        row = [
            Paragraph(str(rank),              cell_style),
            Paragraph(r["company"].title(),    name_style),
            Paragraph(str(r["open_pos"]),      cell_style),
            Paragraph(str(r["open_items"]),    cell_style),
            Paragraph(f"{r['open_qty']:,}",    cell_style),
            Paragraph(str(r["overdue"]) + (" ⚠" if r["overdue"] > 0 else ""), over_para),
            Paragraph(nearest_str,             cell_style),
            Paragraph(days_str,                days_para),
            Paragraph(str(r["history"]),       cell_style),
            Paragraph(on_time,                 ot_para),
            Paragraph(str(r["issues"]) + (" ⚠" if r["issues"] > 0 else ""), issue_para),
            Paragraph(str(int(sc)),            sc_para),
        ]

        ri = len(table_data)
        table_data.append(row)
        cmd.append(("BACKGROUND", (0, ri), (-1, ri), row_bg))
        cmd.append(("ROWHEIGHT",  (0, ri), (-1, ri), 20))
        # Highlight rank 1–3 in left border
        if rank <= 3:
            cmd.append(("LINEBEFORE", (0, ri), (0, ri), 3,
                        CMF_GREEN if rank == 1 else CMF_BLUE if rank == 2 else CMF_ACCENT))

    tbl = Table(table_data, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(TableStyle(cmd))
    story.append(tbl)
    story.append(Spacer(1, 12))

    # ── Legend / Key ──────────────────────────────────────────────────────────
    legend_style = ParagraphStyle("legend",
        fontName="Helvetica", fontSize=8, textColor=CMF_MUTED, leading=11)

    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#CBD5E1"), spaceBefore=4, spaceAfter=4))

    score_formula = (
        "<b>Priority Score formula:</b>  "
        "(Open POs × 15) + (Open Line Items × 8) + (Open Qty / 100, max 50) + "
        "(Historical Completions × 5) + (On-Time % × 0.3) − (Issues × 10) − (Overdue Parts × 12).  "
        "<b>On-Time Rate</b> is calculated from the production LOG (completions vs. original due dates).  "
        "<b>Hist. Compl.</b> = total process steps recorded in LOG for this customer (relationship depth indicator).  "
        "Customers with <b>zero open work and zero history</b> are excluded."
    )
    story.append(Paragraph(score_formula, legend_style))

    # ── Footer ────────────────────────────────────────────────────────────────
    footer_style = ParagraphStyle("footer",
        fontName="Helvetica-Oblique", fontSize=7.5,
        textColor=CMF_MUTED, alignment=TA_RIGHT, spaceBefore=4)
    story.append(Paragraph(
        f"Generated {TODAY.strftime('%B %d, %Y')}  ·  CMF Manufacturing  ·  Confidential",
        footer_style))

    doc.build(story)
    print(f"PDF written → {output_path}")


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    xlsx_path = "CMF WIP - Schedule.xlsx"
    output    = "CMF_Customer_Evaluation.pdf"

    print("Loading workbook…")
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    print("Parsing MAIN sheet…")
    main_data = load_main(wb["MAIN"])

    if "LOG" in wb.sheetnames:
        print("Parsing LOG sheet…")
        log_data = load_log(wb["LOG"])
    else:
        log_data = {}

    print("Building customer metrics…")
    rows = build_customer_table(main_data, log_data)

    print(f"  → {len(rows)} customers with activity")
    for r in rows[:5]:
        print(f"    #{rows.index(r)+1} {r['company']:30s}  score={r['score']:6.1f}  "
              f"pos={r['open_pos']}  items={r['open_items']}  hist={r['history']}")

    print("Building PDF…")
    build_pdf(output, rows)


if __name__ == "__main__":
    main()
