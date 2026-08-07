"""
CMF Schedule v2 — Phase 2: order entry, routing entry, release gate.

Server-rendered HTML, no build step. One Python process plus Postgres.

The release gate replaces the old ADMIN INPUT / MAIN sheet split. In the
workbook, parts moved between sheets automatically based on whether they had a
CURRENT STEP and a process date — so clearing a field made a part silently
teleport out of MAIN. Here, release is an explicit action with an explicit
reason when it is blocked, and it is reversible.

Run:
    DATABASE_URL=postgresql:///cmf uvicorn v2.app.main:app --reload
"""

import os
from collections import OrderedDict
from urllib.parse import quote
from datetime import datetime, date, timedelta

from fastapi import FastAPI, Form, Request, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from . import auth, db

_HERE = os.path.dirname(os.path.abspath(__file__))

app = FastAPI(title="CMF Schedule")
app.mount("/static", StaticFiles(directory=os.path.join(_HERE, "static")), name="static")
templates = Jinja2Templates(directory=os.path.join(_HERE, "templates"))

ACTOR_COOKIE = "cmf_actor"


@app.middleware("http")
async def require_login(request: Request, call_next):
    """Gate every page behind the team password when one is configured."""
    if auth.enabled() and not auth.is_public(request.url.path):
        if not auth.read_session(request.cookies.get(auth.SESSION_COOKIE, "")):
            if request.method == "GET":
                nxt = request.url.path
                if request.url.query:
                    nxt += "?" + request.url.query
                return RedirectResponse(f"/login?next={quote(nxt, safe='')}",
                                        status_code=303)
            return Response("Session expired — reload and sign in again.",
                            status_code=401)
    return await call_next(request)


@app.get("/login", response_class=HTMLResponse)
def login_form(request: Request, next: str = "/", error: str = ""):
    if not auth.enabled():
        return RedirectResponse("/", status_code=303)
    return templates.TemplateResponse(
        request, "login.html",
        {"request": request, "next": next, "error": error, "actor": ""})


@app.post("/login")
def login(request: Request, password: str = Form(""), name: str = Form(""),
          next: str = Form("/")):
    if not auth.password_ok(password):
        return RedirectResponse(
            f"/login?next={quote(next, safe='')}&error=1", status_code=303)

    target = next if next.startswith("/") and not next.startswith("//") else "/"
    resp = RedirectResponse(target, status_code=303)
    resp.set_cookie(auth.SESSION_COOKIE, auth.make_session(name.strip()),
                    max_age=auth.SESSION_DAYS * 86400, httponly=True,
                    samesite="lax", secure=request.url.scheme == "https")
    if name.strip():
        resp.set_cookie(ACTOR_COOKIE, name.strip(),
                        max_age=auth.SESSION_DAYS * 86400, samesite="lax")
    return resp


@app.post("/logout")
def logout():
    resp = RedirectResponse("/login", status_code=303)
    resp.delete_cookie(auth.SESSION_COOKIE)
    return resp


def actor(request: Request) -> str:
    return request.cookies.get(ACTOR_COOKIE) or "UNSPECIFIED"


def render(request: Request, template: str, **kw):
    """Starlette wants the request first; context carries it too for {{ url_for }}."""
    return templates.TemplateResponse(
        request, template, {"request": request, "actor": actor(request), **kw})


# ── Release gate ──────────────────────────────────────────────────────────────

def release_blockers(work_order_id: int):
    """Why this work order cannot be released yet. Empty list means it can.

    The rule mirrors the workbook's part_is_ready_for_main(): a part is ready
    once it has at least one routing step. The difference is that this is
    reported, not applied silently.
    """
    rows = db.query(
        """SELECT p.id, p.line_no, p.part_number, p.description,
                  count(rs.id) AS step_count
             FROM parts p
             LEFT JOIN routing_steps rs ON rs.part_id = p.id
            WHERE p.work_order_id = %s
            GROUP BY p.id
            ORDER BY p.line_no""",
        (work_order_id,))
    if not rows:
        return [{"reason": "This work order has no parts yet."}]
    return [
        {"part_id": r["id"], "line_no": r["line_no"],
         "part_number": r["part_number"], "description": r["description"],
         "reason": "No routing steps"}
        for r in rows if r["step_count"] == 0
    ]


# ── Identity (placeholder for real auth) ──────────────────────────────────────

@app.post("/whoami")
def set_actor(name: str = Form(...)):
    resp = RedirectResponse("/", status_code=303)
    resp.set_cookie(ACTOR_COOKIE, name, max_age=60 * 60 * 24 * 30)
    return resp


# ── Orders ────────────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
def index(request: Request, status: str = "", q: str = ""):
    where, params = [], []
    if status:
        where.append("w.status = %s::wo_status")
        params.append(status)
    if q:
        where.append("(w.wo_number ILIKE %s OR w.po_number ILIKE %s "
                     "OR w.company ILIKE %s OR w.job_name ILIKE %s)")
        params += [f"%{q}%"] * 4
    clause = ("WHERE " + " AND ".join(where)) if where else ""

    orders = db.query(f"""
        SELECT w.*,
               count(p.id)                                    AS part_count,
               count(p.id) FILTER (WHERE rs.n IS NULL OR rs.n = 0) AS unrouted_parts
          FROM work_orders w
          LEFT JOIN parts p ON p.work_order_id = w.id
          LEFT JOIN LATERAL (
              SELECT count(*) n FROM routing_steps r WHERE r.part_id = p.id
          ) rs ON TRUE
          {clause}
          GROUP BY w.id
          ORDER BY
            CASE w.status WHEN 'draft' THEN 0 WHEN 'released' THEN 1 ELSE 2 END,
            w.cust_ship_date NULLS LAST, w.wo_number
          LIMIT 300""", params)

    counts = {r["status"]: r["n"] for r in db.query(
        "SELECT status::text AS status, count(*) n FROM work_orders GROUP BY status")}
    employees = db.query("SELECT name FROM employees WHERE active ORDER BY name")

    return render(request, "orders.html", orders=orders, counts=counts,
                  status=status, q=q, employees=employees)


@app.get("/orders/new", response_class=HTMLResponse)
def new_order_form(request: Request):
    return render(request, "order_new.html")


@app.post("/orders")
def create_order(request: Request,
                 wo_number: str = Form(...),
                 po_number: str = Form(""),
                 company: str = Form(""),
                 job_name: str = Form(""),
                 cust_ship_date: str = Form("")):
    wo_number = wo_number.strip()
    if not wo_number:
        raise HTTPException(400, "WO number is required")

    if db.query_one("SELECT id FROM work_orders WHERE wo_number = %s", (wo_number,)):
        raise HTTPException(409, f"Work order {wo_number} already exists")

    with db.cursor(commit=True) as cur:
        cur.execute(
            """INSERT INTO work_orders
                 (wo_number, po_number, company, job_name, cust_ship_date,
                  status, source, created_by)
               VALUES (%s,%s,%s,%s,%s,'draft','manual',%s)
               RETURNING id""",
            (wo_number, po_number.strip() or None, company.strip() or None,
             job_name.strip() or None, cust_ship_date or None, actor(request)))
        wo_id = cur.fetchone()["id"]
        db.audit(cur, actor(request), "work_order", wo_id, "create",
                 after={"wo_number": wo_number, "status": "draft"})

    return RedirectResponse(f"/orders/{wo_id}", status_code=303)


@app.get("/orders/{wo_id}", response_class=HTMLResponse)
def order_detail(request: Request, wo_id: int):
    wo = db.query_one("SELECT * FROM work_orders WHERE id = %s", (wo_id,))
    if not wo:
        raise HTTPException(404, "Work order not found")

    parts = db.query(
        """SELECT p.*,
                  count(rs.id)                                AS step_count,
                  count(rs.id) FILTER (WHERE rs.status='done') AS done_count,
                  min(rs.due_date) FILTER (WHERE rs.status='open') AS next_due
             FROM parts p
             LEFT JOIN routing_steps rs ON rs.part_id = p.id
            WHERE p.work_order_id = %s
            GROUP BY p.id ORDER BY p.line_no""", (wo_id,))

    steps = {}
    for r in db.query(
        """SELECT rs.*, pr.name AS process_name, pr.is_outside
             FROM routing_steps rs
             JOIN processes pr ON pr.code = rs.process_code
             JOIN parts p ON p.id = rs.part_id
            WHERE p.work_order_id = %s
            ORDER BY rs.due_date NULLS LAST, pr.sort_order""", (wo_id,)):
        steps.setdefault(r["part_id"], []).append(r)

    return render(request, "order_detail.html", wo=wo, parts=parts, steps=steps,
                  blockers=release_blockers(wo_id))


@app.post("/orders/{wo_id}/parts")
def add_part(request: Request, wo_id: int,
             part_number: str = Form(""),
             description: str = Form(""),
             qty: str = Form(""),
             material: str = Form(""),
             thickness: str = Form(""),
             delivery_date: str = Form("")):
    if not db.query_one("SELECT id FROM work_orders WHERE id = %s", (wo_id,)):
        raise HTTPException(404, "Work order not found")
    if not (part_number.strip() or description.strip()):
        raise HTTPException(400, "A part needs at least a part number or description")

    with db.cursor(commit=True) as cur:
        cur.execute("SELECT COALESCE(max(line_no), 0) + 1 AS n FROM parts "
                    "WHERE work_order_id = %s", (wo_id,))
        line_no = cur.fetchone()["n"]
        cur.execute(
            """INSERT INTO parts (work_order_id, line_no, part_number, description,
                                  qty, material, thickness, delivery_date)
               VALUES (%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id""",
            (wo_id, line_no, part_number.strip() or None,
             description.strip() or None,
             float(qty) if qty.strip() else None,
             material.strip() or None, thickness.strip() or None,
             delivery_date or None))
        part_id = cur.fetchone()["id"]
        db.audit(cur, actor(request), "part", part_id, "create",
                 after={"work_order_id": wo_id, "line_no": line_no})

    return RedirectResponse(f"/orders/{wo_id}", status_code=303)


@app.post("/parts/{part_id}/delete")
def delete_part(request: Request, part_id: int):
    row = db.query_one("SELECT work_order_id FROM parts WHERE id = %s", (part_id,))
    if not row:
        raise HTTPException(404, "Part not found")
    with db.cursor(commit=True) as cur:
        db.audit(cur, actor(request), "part", part_id, "delete", before=dict(row))
        cur.execute("DELETE FROM parts WHERE id = %s", (part_id,))
    return RedirectResponse(f"/orders/{row['work_order_id']}", status_code=303)


# ── Routing (engineer) ────────────────────────────────────────────────────────

@app.get("/parts/{part_id}/routing", response_class=HTMLResponse)
def routing_form(request: Request, part_id: int):
    part = db.query_one(
        """SELECT p.*, w.wo_number, w.company, w.job_name, w.cust_ship_date,
                  w.id AS wo_id, w.status::text AS wo_status
             FROM parts p JOIN work_orders w ON w.id = p.work_order_id
            WHERE p.id = %s""", (part_id,))
    if not part:
        raise HTTPException(404, "Part not found")

    processes = db.query("SELECT * FROM processes ORDER BY sort_order")
    existing = {r["process_code"]: r for r in db.query(
        "SELECT * FROM routing_steps WHERE part_id = %s", (part_id,))}

    return render(request, "routing.html", part=part, processes=processes,
                  existing=existing)


@app.post("/parts/{part_id}/routing")
async def save_routing(request: Request, part_id: int):
    """One row per process that has a due date. Clearing a date removes the step.

    A step already marked done is never silently dropped — its completion
    history is preserved even if the date is cleared.
    """
    part = db.query_one("SELECT work_order_id FROM parts WHERE id = %s", (part_id,))
    if not part:
        raise HTTPException(404, "Part not found")

    form = await request.form()
    delivery = (form.get("delivery_date") or "").strip() or None

    submitted = {}
    for code, value in form.items():
        if not code.startswith("due_"):
            continue
        value = (value or "").strip()
        if value:
            submitted[code[4:]] = value

    changed = 0
    with db.cursor(commit=True) as cur:
        cur.execute("UPDATE parts SET delivery_date = %s, updated_at = now() "
                    "WHERE id = %s", (delivery, part_id))

        cur.execute("SELECT process_code, due_date, status FROM routing_steps "
                    "WHERE part_id = %s", (part_id,))
        existing = {r["process_code"]: r for r in cur.fetchall()}

        for code, due in submitted.items():
            prior = existing.get(code)
            if prior is None:
                cur.execute(
                    """INSERT INTO routing_steps (part_id, process_code, due_date)
                       VALUES (%s,%s,%s)""", (part_id, code, due))
                changed += 1
            elif str(prior["due_date"]) != due:
                cur.execute(
                    """UPDATE routing_steps SET due_date = %s
                        WHERE part_id = %s AND process_code = %s""",
                    (due, part_id, code))
                changed += 1

        for code, prior in existing.items():
            if code not in submitted and prior["status"] != "done":
                cur.execute("DELETE FROM routing_steps WHERE part_id = %s "
                            "AND process_code = %s", (part_id, code))
                changed += 1

        if changed:
            db.audit(cur, actor(request), "part", part_id, "routing_update",
                     before={k: str(v["due_date"]) for k, v in existing.items()},
                     after=submitted)

    return RedirectResponse(f"/orders/{part['work_order_id']}", status_code=303)


# ── Release gate ──────────────────────────────────────────────────────────────

@app.post("/orders/{wo_id}/release")
def release(request: Request, wo_id: int):
    wo = db.query_one("SELECT * FROM work_orders WHERE id = %s", (wo_id,))
    if not wo:
        raise HTTPException(404, "Work order not found")
    if wo["status"] != "draft":
        raise HTTPException(409, f"Only draft orders can be released "
                                 f"(this one is {wo['status']})")

    blockers = release_blockers(wo_id)
    if blockers:
        raise HTTPException(409, "Not ready to release: "
                            + "; ".join(b["reason"] for b in blockers))

    with db.cursor(commit=True) as cur:
        cur.execute("UPDATE work_orders SET status='released', updated_at=now() "
                    "WHERE id = %s", (wo_id,))
        db.audit(cur, actor(request), "work_order", wo_id, "release",
                 before={"status": "draft"}, after={"status": "released"})

    return RedirectResponse(f"/orders/{wo_id}", status_code=303)


@app.post("/orders/{wo_id}/unrelease")
def unrelease(request: Request, wo_id: int):
    """Reversible on purpose. The workbook moved parts between sheets with no
    way back; a mistaken release should not require editing the database."""
    wo = db.query_one("SELECT * FROM work_orders WHERE id = %s", (wo_id,))
    if not wo:
        raise HTTPException(404, "Work order not found")
    if wo["status"] != "released":
        raise HTTPException(409, f"Only released orders can be pulled back "
                                 f"(this one is {wo['status']})")

    with db.cursor(commit=True) as cur:
        cur.execute("UPDATE work_orders SET status='draft', updated_at=now() "
                    "WHERE id = %s", (wo_id,))
        db.audit(cur, actor(request), "work_order", wo_id, "unrelease",
                 before={"status": "released"}, after={"status": "draft"})

    return RedirectResponse(f"/orders/{wo_id}", status_code=303)


@app.get("/health")
def health():
    row = db.query_one("SELECT count(*) AS n FROM work_orders")
    return {"status": "ok", "work_orders": row["n"], "at": datetime.now().isoformat()}


# ── Boards (phase 3) ──────────────────────────────────────────────────────────

BOARD_WINDOW_DAYS = 28   # rolling 4-week window, matching the workbook


def _board_rows(view, days, process=None, company=None):
    """Open steps from v_calendar / v_purchasing, split into overdue and window.

    The views carry no date window on purpose — data has no window, boards do.
    """
    where, params = [], []
    if process:
        where.append("process_code = %s")
        params.append(process)
    if company:
        where.append("company = %s")
        params.append(company)
    clause = ("WHERE " + " AND ".join(where)) if where else ""

    rows = db.query(f"""
        SELECT * FROM {view} {clause}
        ORDER BY due_date NULLS LAST, process_name, wo_number""", params)

    horizon = date.today() + timedelta(days=days)
    overdue = [r for r in rows if r["due_date"] and r["due_date"] < date.today()]
    current = [r for r in rows if r["due_date"] and
               date.today() <= r["due_date"] <= horizon]
    later = [r for r in rows if not r["due_date"] or r["due_date"] > horizon]
    return rows, overdue, current, later


def _group_by_process(rows):
    out = OrderedDict()
    for r in rows:
        out.setdefault(r["process_name"], []).append(r)
    return out


def _group_by_day(rows):
    out = OrderedDict()
    for r in rows:
        out.setdefault(r["due_date"], []).append(r)
    return out


@app.get("/board", response_class=HTMLResponse)
def board(request: Request, process: str = "", company: str = "",
          days: int = BOARD_WINDOW_DAYS, group: str = "day"):
    """CALENDAR — the 10-minute standup board."""
    rows, overdue, current, later = _board_rows("v_calendar", days, process, company)
    grouped = _group_by_day(current) if group == "day" else _group_by_process(current)

    return render(request, "board.html",
                  title="Calendar", view="calendar", board_url="/board",
                  rows=rows, overdue=overdue, grouped=grouped, later=later,
                  group=group, days=days, process=process, company=company,
                  processes=db.query("SELECT * FROM processes WHERE NOT is_outside "
                                     "ORDER BY sort_order"),
                  companies=db.query("SELECT DISTINCT company FROM v_calendar "
                                     "WHERE company IS NOT NULL ORDER BY company"),
                  today=date.today())


@app.get("/purchasing", response_class=HTMLResponse)
def purchasing(request: Request, process: str = "", company: str = "",
               days: int = BOARD_WINDOW_DAYS, group: str = "process"):
    """PURCHASING — all open outside/vendor work."""
    rows, overdue, current, later = _board_rows("v_purchasing", days, process, company)
    grouped = _group_by_day(current) if group == "day" else _group_by_process(current)

    return render(request, "board.html",
                  title="Purchasing", view="purchasing", board_url="/purchasing",
                  rows=rows, overdue=overdue, grouped=grouped, later=later,
                  group=group, days=days, process=process, company=company,
                  processes=db.query("SELECT * FROM processes WHERE is_outside "
                                     "ORDER BY sort_order"),
                  companies=db.query("SELECT DISTINCT company FROM v_purchasing "
                                     "WHERE company IS NOT NULL ORDER BY company"),
                  today=date.today())


@app.get("/screenshot/{screenshot_id}")
def screenshot(screenshot_id: int):
    """Images come from Postgres — no object storage in the deployment."""
    row = db.query_one("SELECT content_type, bytes FROM screenshots WHERE id = %s",
                       (screenshot_id,))
    if not row:
        raise HTTPException(404, "Screenshot not found")
    return Response(content=bytes(row["bytes"]), media_type=row["content_type"],
                    headers={"Cache-Control": "public, max-age=31536000, immutable"})
