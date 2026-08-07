#!/usr/bin/env python3
"""
End-to-end regression — drives a real browser through the whole workflow.

Starts the app, then walks the path your team actually walks:

    administrator creates an order
      → blocked: no parts
      → adds a part
      → blocked: not routed, with the line named
    engineer routes the part
      → gate clears
      → releases
    the step appears on the right board, with the right photo
      → marked done, it leaves the board
      → shipped, the order leaves the board entirely

Run on its own:  DATABASE_URL=... python3 v2/tests/e2e.py
Or as part of:   ./v2/run_tests.sh
"""

import os
import re
import subprocess
import sys
import time
import urllib.error
import urllib.request
import uuid

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

PORT = int(os.environ.get("E2E_PORT", "8099"))
BASE = f"http://127.0.0.1:{PORT}"

CHROME_CANDIDATES = [
    os.environ.get("CHROME_PATH", ""),
    "/opt/pw-browsers/chromium-1194/chrome-linux/chrome",
    "/usr/bin/chromium", "/usr/bin/chromium-browser", "/usr/bin/google-chrome",
]

_passed, _failed = 0, 0


def check(label, ok):
    global _passed, _failed
    if ok:
        _passed += 1
        print(f"    \033[32mPASS\033[0m  {label}")
    else:
        _failed += 1
        print(f"    \033[31mFAIL\033[0m  {label}")


def find_chrome():
    from glob import glob
    for c in CHROME_CANDIDATES:
        if c and os.path.exists(c):
            return c
    for pat in ("/opt/pw-browsers/chromium-*/chrome-linux/chrome",
                "/root/.cache/ms-playwright/chromium-*/chrome-linux/chrome"):
        hits = glob(pat)
        if hits:
            return hits[0]
    return None


def wait_for_server(proc, timeout=45):
    deadline = time.time() + timeout
    while time.time() < deadline:
        if proc.poll() is not None:
            return False
        try:
            with urllib.request.urlopen(f"{BASE}/health", timeout=2) as r:
                if r.status == 200:
                    return True
        except (urllib.error.URLError, OSError):
            time.sleep(0.4)
    return False


def main():
    if not os.environ.get("DATABASE_URL"):
        sys.exit("DATABASE_URL is not set")

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("  playwright not installed — skipping browser run")
        print("  pip install playwright")
        return 0

    chrome = find_chrome()
    if not chrome:
        print("  no chromium binary found — skipping browser run")
        return 0

    server = subprocess.Popen(
        [sys.executable, "-m", "uvicorn", "v2.app.main:app",
         "--port", str(PORT), "--host", "127.0.0.1", "--log-level", "warning"],
        cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)

    try:
        if not wait_for_server(server):
            out = server.stdout.read().decode(errors="replace")[-2000:]
            sys.exit(f"server failed to start:\n{out}")

        wo = f"E2E-{uuid.uuid4().hex[:6].upper()}"

        with sync_playwright() as p:
            b = p.chromium.launch(executable_path=chrome, args=["--no-sandbox"])
            pg = b.new_page(viewport={"width": 1400, "height": 900})

            # ── Administrator creates the order ────────────────────────────
            print("\n  administrator — create order")
            pg.goto(f"{BASE}/orders/new")
            pg.fill("input[name=wo_number]", wo)
            pg.fill("input[name=company]", "E2E CUSTOMER")
            pg.fill("input[name=job_name]", "Regression bracket")
            pg.click('button:has-text("Create")')
            pg.wait_for_load_state("networkidle")
            wo_url = pg.url
            wo_id = wo_url.rstrip("/").split("/")[-1]
            check("order created as draft", "pill-draft" in pg.content())
            check("gate blocks: no parts", "no parts yet" in pg.content().lower())

            # ── Administrator adds a part ──────────────────────────────────
            print("\n  administrator — add part")
            pg.fill("input[name=part_number]", "E2E-BRK")
            pg.fill("input[name=description]", "Regression bracket")
            pg.fill("input[name=qty]", "42")
            pg.click('button:has-text("Add part")')
            pg.wait_for_load_state("networkidle")
            check("gate blocks: not routed", "No routing steps" in pg.content())
            check("release button disabled",
                  pg.locator('button:has-text("Release")').is_disabled())

            # ── Engineer routes it ─────────────────────────────────────────
            print("\n  engineer — routing")
            pg.click('a:has-text("Route")')
            pg.wait_for_load_state("networkidle")
            route_url = pg.url
            pg.fill("input[name=due_WELD]", _days(2))          # in-house
            pg.fill("input[name=due_PLATING_O]", _days(5))     # outside
            pg.fill("input[name=delivery_date]", _days(20))
            pg.click('button:has-text("Save routing")')
            pg.wait_for_load_state("networkidle")
            check("gate clears once routed", "Ready to release" in pg.content())

            # ── Release ────────────────────────────────────────────────────
            print("\n  engineer — release")
            pg.click('button:has-text("Release to production")')
            pg.wait_for_load_state("networkidle")
            check("order is released", "pill-released" in pg.content())

            # ── Boards ─────────────────────────────────────────────────────
            print("\n  production — boards")
            pg.goto(f"{BASE}/board?days=56")
            pg.wait_for_load_state("networkidle")
            cal = pg.content()
            check("WELD is on the calendar", wo in cal and "WELD" in cal)
            check("PLATING (O) is NOT on the calendar",
                  not re.search(rf"{wo}.{{0,400}}PLATING", cal, re.S))

            pg.goto(f"{BASE}/purchasing?days=56")
            pg.wait_for_load_state("networkidle")
            pur = pg.content()
            check("PLATING (O) is on purchasing", wo in pur and "PLATING (O)" in pur)
            check("WELD is NOT on purchasing",
                  not re.search(rf"{wo}.{{0,400}}>WELD<", pur, re.S))

            # ── Completing a step removes it ───────────────────────────────
            print("\n  production — completion and shipping")
            from v2.app import db
            db.execute("""UPDATE routing_steps rs SET status='done',
                            completed_by='E2E', completed_at=now()
                          FROM parts p JOIN work_orders w ON w.id = p.work_order_id
                          WHERE rs.part_id = p.id AND w.wo_number = %s
                            AND rs.process_code = 'WELD'""", (wo,))
            pg.goto(f"{BASE}/board?days=56")
            pg.wait_for_load_state("networkidle")
            check("completed step leaves the board", wo not in pg.content())

            # ── Shipping removes the order ─────────────────────────────────
            db.execute("UPDATE work_orders SET status='shipped' WHERE wo_number=%s",
                       (wo,))
            pg.goto(f"{BASE}/purchasing?days=56")
            pg.wait_for_load_state("networkidle")
            check("shipped order leaves the board", wo not in pg.content())
            row = db.query_one("SELECT status::text s FROM work_orders "
                               "WHERE wo_number=%s", (wo,))
            check("shipped order still exists (status, not deletion)",
                  row is not None and row["s"] == "shipped")

            # ── Reversibility ──────────────────────────────────────────────
            print("\n  safety — reversibility and audit")
            db.execute("UPDATE work_orders SET status='released' WHERE wo_number=%s",
                       (wo,))
            pg.goto(wo_url)
            pg.wait_for_load_state("networkidle")
            check("released order offers pull-back",
                  "Pull back to draft" in pg.content())

            audits = db.query("""SELECT action FROM audit_log
                                 WHERE entity='work_order' AND entity_id=%s""",
                              (int(wo_id),))
            check("release was audited",
                  any(a["action"] == "release" for a in audits))

            # ── Screenshots come from Postgres ─────────────────────────────
            shot = db.query_one("SELECT id, byte_size FROM screenshots LIMIT 1")
            if shot:
                resp = pg.request.get(f"{BASE}/screenshot/{shot['id']}")
                check("screenshot serves from Postgres",
                      resp.status == 200 and len(resp.body()) == shot["byte_size"])

            b.close()

        # cleanup
        from v2.app import db
        db.execute("DELETE FROM work_orders WHERE wo_number = %s", (wo,))

    finally:
        server.terminate()
        try:
            server.wait(timeout=10)
        except subprocess.TimeoutExpired:
            server.kill()

    print(f"\n  {_passed} passed, {_failed} failed")
    return 1 if _failed else 0


def _days(n):
    from datetime import date, timedelta
    return (date.today() + timedelta(days=n)).isoformat()


if __name__ == "__main__":
    sys.exit(main())
