"""Database access. Thin psycopg wrapper — no ORM, the schema is the model."""

import os
from contextlib import contextmanager

import psycopg
from psycopg.rows import dict_row

DATABASE_URL = os.environ.get("DATABASE_URL", "postgresql:///cmf")

_pool = None


def _get_pool():
    global _pool
    if _pool is None:
        from psycopg_pool import ConnectionPool
        _pool = ConnectionPool(DATABASE_URL, min_size=1, max_size=8, open=True,
                               kwargs={"row_factory": dict_row})
    return _pool


@contextmanager
def cursor(commit=False):
    """Yields a dict-row cursor. Rolls back on exception."""
    try:
        pool = _get_pool()
        with pool.connection() as conn:
            with conn.cursor() as cur:
                yield cur
                if commit:
                    conn.commit()
    except ImportError:
        # psycopg_pool is optional; fall back to a plain connection.
        with psycopg.connect(DATABASE_URL, row_factory=dict_row) as conn:
            with conn.cursor() as cur:
                yield cur
                if commit:
                    conn.commit()


def query(sql, params=None):
    with cursor() as cur:
        cur.execute(sql, params or ())
        return cur.fetchall()


def query_one(sql, params=None):
    with cursor() as cur:
        cur.execute(sql, params or ())
        return cur.fetchone()


def execute(sql, params=None, returning=False):
    with cursor(commit=True) as cur:
        cur.execute(sql, params or ())
        return cur.fetchone() if returning else None


def audit(cur, actor, entity, entity_id, action, before=None, after=None):
    """Append-only history. Called inside an existing transaction."""
    import json
    cur.execute(
        """INSERT INTO audit_log (actor, entity, entity_id, action, before, after)
           VALUES (%s, %s, %s, %s, %s, %s)""",
        (actor, entity, entity_id, action,
         json.dumps(before, default=str) if before else None,
         json.dumps(after, default=str) if after else None))
