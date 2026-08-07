#!/usr/bin/env python3
"""
deploy_setup.py — makes the database ready, and is safe to run on every deploy.

Runs automatically before each release. It is idempotent by design:

  * schema missing      → applies db/schema.sql
  * schema present      → leaves it alone
  * no work orders yet  → imports the workbook from the repo (first deploy only)
  * work orders present → imports nothing, ever

That last rule matters. Once the shop is live, the database is the source of
truth and the workbook in the repo is a stale snapshot. Re-importing it would
overwrite real work, so this refuses to.
"""

import os
import subprocess
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_HERE)
SCHEMA = os.path.join(_ROOT, "db", "schema.sql")
WORKBOOK = os.path.join(_ROOT, "CMF WIP - Schedule.xlsx")
HISTORY = os.environ.get("HISTORY_FILE", "")


def log(msg):
    print(f"[setup] {msg}", flush=True)


def main():
    url = os.environ.get("DATABASE_URL")
    if not url:
        sys.exit("[setup] DATABASE_URL is not set")

    try:
        import psycopg
    except ImportError:
        sys.exit("[setup] psycopg not installed")

    # Render hands out postgres:// URLs; psycopg wants postgresql://
    if url.startswith("postgres://"):
        url = "postgresql://" + url[len("postgres://"):]
        os.environ["DATABASE_URL"] = url

    with psycopg.connect(url) as conn:
        with conn.cursor() as cur:
            cur.execute("""SELECT EXISTS (
                             SELECT 1 FROM information_schema.tables
                             WHERE table_schema='public' AND table_name='work_orders')""")
            has_schema = cur.fetchone()[0]

        if not has_schema:
            log("no schema found — applying db/schema.sql")
            with open(SCHEMA) as fh:
                sql = fh.read()
            with conn.cursor() as cur:
                cur.execute(sql)
            conn.commit()
            log("schema applied")
        else:
            log("schema already present — leaving it alone")

        with conn.cursor() as cur:
            cur.execute("SELECT count(*) FROM work_orders")
            existing = cur.fetchone()[0]

    if existing:
        log(f"database already holds {existing} work order(s) — skipping import")
        log("the database is the source of truth now; the workbook is a snapshot")
        return 0

    if not os.path.exists(WORKBOOK):
        log("no workbook in the repo — starting empty")
        return 0

    log("empty database — importing the workbook (first deploy only)")
    cmd = [sys.executable, os.path.join(_HERE, "import_workbook.py"),
           "--workbook", WORKBOOK, "--database-url", url]
    if HISTORY and os.path.exists(os.path.join(_ROOT, HISTORY)):
        cmd += ["--history", os.path.join(_ROOT, HISTORY)]
        log(f"including history: {HISTORY}")

    result = subprocess.run(cmd)
    if result.returncode != 0:
        sys.exit("[setup] import failed — deploy aborted so nothing half-loaded goes live")

    log("import complete")
    return 0


if __name__ == "__main__":
    sys.exit(main())
