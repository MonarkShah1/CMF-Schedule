"""
Phase 2 tests — run against a real Postgres loaded with the schema.

    DATABASE_URL=postgresql:///cmf pytest v2/tests -q

Each test creates its own work order and cleans up, so the suite can run
against a database that already holds imported data.
"""

import os
import sys
import uuid

import pytest
from fastapi.testclient import TestClient

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__)))))

from v2.app import db                     # noqa: E402
from v2.app.main import app, release_blockers  # noqa: E402


@pytest.fixture
def client():
    return TestClient(app)


@pytest.fixture
def wo():
    """A throwaway draft work order, removed afterwards."""
    number = f"TEST-{uuid.uuid4().hex[:8]}"
    row = db.execute(
        """INSERT INTO work_orders (wo_number, company, status, source)
           VALUES (%s, 'TEST CO', 'draft', 'manual') RETURNING id""",
        (number,), returning=True)
    yield {"id": row["id"], "wo_number": number}
    db.execute("DELETE FROM work_orders WHERE id = %s", (row["id"],))


def _add_part(wo_id, line_no=1, part_number="P-1"):
    return db.execute(
        """INSERT INTO parts (work_order_id, line_no, part_number, qty)
           VALUES (%s,%s,%s,10) RETURNING id""",
        (wo_id, line_no, part_number), returning=True)["id"]


# ── Release gate ──────────────────────────────────────────────────────────────

def test_order_with_no_parts_cannot_release(wo):
    blockers = release_blockers(wo["id"])
    assert blockers and "no parts" in blockers[0]["reason"].lower()


def test_unrouted_part_blocks_release(wo):
    _add_part(wo["id"])
    blockers = release_blockers(wo["id"])
    assert len(blockers) == 1
    assert blockers[0]["reason"] == "No routing steps"


def test_routed_part_clears_the_gate(wo):
    part_id = _add_part(wo["id"])
    db.execute("""INSERT INTO routing_steps (part_id, process_code, due_date)
                  VALUES (%s, 'LASER_CUT', '2026-09-01')""", (part_id,))
    assert release_blockers(wo["id"]) == []


def test_one_unrouted_part_blocks_the_whole_order(wo):
    """Partial routing must not release — the workbook let half-routed jobs
    slip onto the board because the check was per-part, not per-order."""
    a = _add_part(wo["id"], 1, "P-A")
    _add_part(wo["id"], 2, "P-B")
    db.execute("""INSERT INTO routing_steps (part_id, process_code, due_date)
                  VALUES (%s, 'WELD', '2026-09-01')""", (a,))
    blockers = release_blockers(wo["id"])
    assert len(blockers) == 1
    assert blockers[0]["part_number"] == "P-B"


def test_release_endpoint_rejects_when_blocked(client, wo):
    _add_part(wo["id"])
    r = client.post(f"/orders/{wo['id']}/release", follow_redirects=False)
    assert r.status_code == 409
    assert db.query_one("SELECT status FROM work_orders WHERE id = %s",
                        (wo["id"],))["status"] == "draft"


def test_release_then_pull_back_is_reversible(client, wo):
    part_id = _add_part(wo["id"])
    db.execute("""INSERT INTO routing_steps (part_id, process_code, due_date)
                  VALUES (%s, 'BEND', '2026-09-01')""", (part_id,))

    assert client.post(f"/orders/{wo['id']}/release",
                       follow_redirects=False).status_code == 303
    assert db.query_one("SELECT status FROM work_orders WHERE id=%s",
                        (wo["id"],))["status"] == "released"

    assert client.post(f"/orders/{wo['id']}/unrelease",
                       follow_redirects=False).status_code == 303
    assert db.query_one("SELECT status FROM work_orders WHERE id=%s",
                        (wo["id"],))["status"] == "draft"


def test_cannot_release_twice(client, wo):
    part_id = _add_part(wo["id"])
    db.execute("""INSERT INTO routing_steps (part_id, process_code, due_date)
                  VALUES (%s, 'BEND', '2026-09-01')""", (part_id,))
    client.post(f"/orders/{wo['id']}/release", follow_redirects=False)
    assert client.post(f"/orders/{wo['id']}/release",
                       follow_redirects=False).status_code == 409


# ── Routing entry ─────────────────────────────────────────────────────────────

def test_saving_routing_creates_only_dated_steps(client, wo):
    part_id = _add_part(wo["id"])
    r = client.post(f"/parts/{part_id}/routing",
                    data={"due_LASER_CUT": "2026-09-01",
                          "due_WELD": "2026-09-03",
                          "due_PAINT": "",              # blank → no step
                          "delivery_date": "2026-09-10"},
                    follow_redirects=False)
    assert r.status_code == 303

    codes = {s["process_code"] for s in db.query(
        "SELECT process_code FROM routing_steps WHERE part_id = %s", (part_id,))}
    assert codes == {"LASER_CUT", "WELD"}

    part = db.query_one("SELECT delivery_date FROM parts WHERE id=%s", (part_id,))
    assert str(part["delivery_date"]) == "2026-09-10"


def test_delivery_date_is_not_a_routing_step(client, wo):
    """Delivery is a part attribute. As a step it would appear on the calendar
    as its own station — the noise the original design removed."""
    part_id = _add_part(wo["id"])
    client.post(f"/parts/{part_id}/routing",
                data={"due_SAW": "2026-09-01", "delivery_date": "2026-09-20"},
                follow_redirects=False)
    steps = db.query("SELECT process_code FROM routing_steps WHERE part_id=%s",
                     (part_id,))
    assert [s["process_code"] for s in steps] == ["SAW"]


def test_clearing_a_date_removes_the_step(client, wo):
    part_id = _add_part(wo["id"])
    client.post(f"/parts/{part_id}/routing", data={"due_SAW": "2026-09-01"},
                follow_redirects=False)
    client.post(f"/parts/{part_id}/routing", data={"due_SAW": ""},
                follow_redirects=False)
    assert db.query("SELECT 1 FROM routing_steps WHERE part_id=%s", (part_id,)) == []


def test_completed_step_survives_a_cleared_date(client, wo):
    """Completion history is never silently discarded."""
    part_id = _add_part(wo["id"])
    db.execute("""INSERT INTO routing_steps (part_id, process_code, due_date,
                                             status, completed_by, completed_at)
                  VALUES (%s,'WELD','2026-09-01','done','Sam', now())""", (part_id,))
    client.post(f"/parts/{part_id}/routing", data={"due_WELD": ""},
                follow_redirects=False)
    rows = db.query("SELECT status, completed_by FROM routing_steps WHERE part_id=%s",
                    (part_id,))
    assert len(rows) == 1 and rows[0]["completed_by"] == "Sam"


def test_routing_update_is_audited(client, wo):
    part_id = _add_part(wo["id"])
    client.post(f"/parts/{part_id}/routing", data={"due_SAW": "2026-09-01"},
                follow_redirects=False)
    rows = db.query("""SELECT * FROM audit_log
                        WHERE entity='part' AND entity_id=%s
                          AND action='routing_update'""", (part_id,))
    assert len(rows) == 1


# ── Order + part creation ─────────────────────────────────────────────────────

def test_duplicate_wo_number_is_rejected(client, wo):
    r = client.post("/orders", data={"wo_number": wo["wo_number"]},
                    follow_redirects=False)
    assert r.status_code == 409


def test_duplicate_part_numbers_are_allowed(client, wo):
    """Verified against the live workbook: WO 4522 carries 18 rows all
    numbered 'CH'. part_number is not an identifier in this business."""
    for _ in range(3):
        r = client.post(f"/orders/{wo['id']}/parts",
                        data={"part_number": "CH", "description": "47 3/8"},
                        follow_redirects=False)
        assert r.status_code == 303
    rows = db.query("SELECT line_no FROM parts WHERE work_order_id=%s "
                    "ORDER BY line_no", (wo["id"],))
    assert [r["line_no"] for r in rows] == [1, 2, 3]


def test_part_needs_a_number_or_description(client, wo):
    r = client.post(f"/orders/{wo['id']}/parts", data={"part_number": "",
                    "description": ""}, follow_redirects=False)
    assert r.status_code == 400


def test_deleting_a_part_removes_its_routing(client, wo):
    part_id = _add_part(wo["id"])
    db.execute("""INSERT INTO routing_steps (part_id, process_code, due_date)
                  VALUES (%s,'SAW','2026-09-01')""", (part_id,))
    client.post(f"/parts/{part_id}/delete", follow_redirects=False)
    assert db.query("SELECT 1 FROM routing_steps WHERE part_id=%s", (part_id,)) == []


# ── Pages render ──────────────────────────────────────────────────────────────

def test_pages_render(client, wo):
    part_id = _add_part(wo["id"])
    for url in ["/", "/?status=draft", "/orders/new",
                f"/orders/{wo['id']}", f"/parts/{part_id}/routing"]:
        assert client.get(url).status_code == 200, url


def test_health(client):
    r = client.get("/health")
    assert r.status_code == 200 and r.json()["status"] == "ok"


def test_missing_records_404(client):
    assert client.get("/orders/99999999").status_code == 404
    assert client.get("/parts/99999999/routing").status_code == 404
