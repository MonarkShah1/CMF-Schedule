"""
Phase 3 regression — CALENDAR and PURCHASING boards.

The rules these lock down are ones the workbook got wrong at some point, so a
regression here is a repeat of a bug you have already paid for.
"""

import os
import sys
import uuid
from datetime import date, timedelta

import pytest
from fastapi.testclient import TestClient

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__)))))

from v2.app import db            # noqa: E402
from v2.app.main import app      # noqa: E402


@pytest.fixture
def client():
    return TestClient(app)


@pytest.fixture
def job():
    """A released work order with one part, cleaned up afterwards."""
    number = f"BRD-{uuid.uuid4().hex[:8]}"
    wo = db.execute(
        """INSERT INTO work_orders (wo_number, company, status, source, cust_ship_date)
           VALUES (%s,'BOARD TEST CO','released','manual', CURRENT_DATE + 30)
           RETURNING id""", (number,), returning=True)
    part = db.execute(
        """INSERT INTO parts (work_order_id, line_no, part_number, qty)
           VALUES (%s, 1, 'BRD-PART', 5) RETURNING id""",
        (wo["id"],), returning=True)
    yield {"wo_id": wo["id"], "wo_number": number, "part_id": part["id"]}
    db.execute("DELETE FROM work_orders WHERE id = %s", (wo["id"],))


def _step(part_id, code, days_out):
    db.execute("""INSERT INTO routing_steps (part_id, process_code, due_date)
                  VALUES (%s,%s,%s)""",
               (part_id, code, date.today() + timedelta(days=days_out)))


def _in(view, wo_number, process=None):
    sql = "SELECT * FROM {} WHERE wo_number = %s".format(view)
    rows = db.query(sql, (wo_number,))
    if process:
        rows = [r for r in rows if r["process_code"] == process]
    return rows


# ── The in-house / outside split ──────────────────────────────────────────────

def test_inhouse_step_is_on_calendar_not_purchasing(job):
    _step(job["part_id"], "WELD", 3)
    assert len(_in("v_calendar", job["wo_number"], "WELD")) == 1
    assert _in("v_purchasing", job["wo_number"], "WELD") == []


def test_outside_step_is_on_purchasing_not_calendar(job):
    _step(job["part_id"], "PLATING_O", 3)
    assert len(_in("v_purchasing", job["wo_number"], "PLATING_O")) == 1
    assert _in("v_calendar", job["wo_number"], "PLATING_O") == []


def test_paint_is_inhouse_but_paint_o_is_not(job):
    """PAINT sits under the 'outside' heading in the old STATIONS list but is
    in-house. Getting this backwards sends real work to the wrong board."""
    _step(job["part_id"], "PAINT", 2)
    _step(job["part_id"], "PAINT_O", 4)
    assert len(_in("v_calendar", job["wo_number"], "PAINT")) == 1
    assert len(_in("v_purchasing", job["wo_number"], "PAINT_O")) == 1


def test_no_process_appears_on_both_boards():
    overlap = db.query("""
        SELECT DISTINCT c.process_code FROM v_calendar c
        JOIN v_purchasing p ON p.process_code = c.process_code""")
    assert overlap == []


def test_every_board_row_maps_to_a_real_process():
    orphans = db.query("""
        SELECT rs.process_code FROM routing_steps rs
        LEFT JOIN processes p ON p.code = rs.process_code
        WHERE p.code IS NULL""")
    assert orphans == []


# ── Delivery date must never become a board station ───────────────────────────

def test_delivery_is_not_a_process_at_all():
    assert db.query("SELECT 1 FROM processes WHERE code = 'DELIVERY'") == []


def test_delivery_date_never_appears_as_a_board_row(job):
    db.execute("UPDATE parts SET delivery_date = CURRENT_DATE + 5 WHERE id = %s",
               (job["part_id"],))
    _step(job["part_id"], "SAW", 2)
    rows = _in("v_calendar", job["wo_number"])
    assert [r["process_code"] for r in rows] == ["SAW"]
    assert rows[0]["delivery_date"] is not None   # carried as a column


# ── Completion and status filtering ───────────────────────────────────────────

def test_completed_step_leaves_the_board(job):
    _step(job["part_id"], "BEND", 3)
    assert len(_in("v_calendar", job["wo_number"], "BEND")) == 1
    db.execute("""UPDATE routing_steps SET status='done', completed_by='Sam',
                  completed_at=now() WHERE part_id=%s AND process_code='BEND'""",
               (job["part_id"],))
    assert _in("v_calendar", job["wo_number"], "BEND") == []


def test_shipped_work_orders_leave_the_board(job):
    """History is a status, not a location. A shipped WO drops off the board
    without its rows being moved anywhere."""
    _step(job["part_id"], "SAW", 3)
    assert len(_in("v_calendar", job["wo_number"])) == 1
    db.execute("UPDATE work_orders SET status='shipped' WHERE id=%s", (job["wo_id"],))
    assert _in("v_calendar", job["wo_number"]) == []
    still_there = db.query_one("SELECT status::text s FROM work_orders WHERE id=%s",
                               (job["wo_id"],))
    assert still_there["s"] == "shipped"      # explicit, not missing


def test_draft_work_orders_stay_off_the_board(job):
    _step(job["part_id"], "SAW", 3)
    db.execute("UPDATE work_orders SET status='draft' WHERE id=%s", (job["wo_id"],))
    assert _in("v_calendar", job["wo_number"]) == []


# ── Overdue bucketing ─────────────────────────────────────────────────────────

def test_overdue_view_only_holds_past_due_steps():
    future = db.query("SELECT 1 FROM v_overdue WHERE due_date >= CURRENT_DATE")
    assert future == []


def test_board_splits_overdue_from_window(client, job):
    _step(job["part_id"], "SAW", -5)      # overdue
    _step(job["part_id"], "BEND", 3)      # in window
    _step(job["part_id"], "WELD", 90)     # beyond
    r = client.get("/board?days=28")
    assert r.status_code == 200
    assert "Overdue" in r.text


# ── Screenshots come from Postgres ────────────────────────────────────────────

def test_screenshots_are_deduplicated_by_digest():
    rows = db.query("SELECT count(*) n, count(DISTINCT digest) d FROM screenshots")
    assert rows[0]["n"] == rows[0]["d"]


def test_screenshot_endpoint_serves_bytes(client):
    row = db.query_one("SELECT id, byte_size, content_type FROM screenshots LIMIT 1")
    if not row:
        pytest.skip("no screenshots imported")
    r = client.get(f"/screenshot/{row['id']}")
    assert r.status_code == 200
    assert r.headers["content-type"] == row["content_type"]
    assert len(r.content) == row["byte_size"]


def test_missing_screenshot_404s(client):
    assert client.get("/screenshot/99999999").status_code == 404


# ── Pages and filters render ──────────────────────────────────────────────────

@pytest.mark.parametrize("url", [
    "/board", "/purchasing",
    "/board?group=process", "/purchasing?group=day",
    "/board?days=7", "/board?days=56",
    "/board?process=WELD", "/purchasing?process=PLATING_O",
])
def test_board_pages_render(client, url):
    assert client.get(url).status_code == 200


def test_purchasing_filter_only_offers_outside_processes(client):
    r = client.get("/purchasing")
    assert "PLATING (O)" in r.text
    # WELD is in-house; it must not be offered as a purchasing station filter
    assert 'value="WELD"' not in r.text
