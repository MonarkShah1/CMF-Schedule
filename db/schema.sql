-- CMF Schedule v2 — Postgres schema
--
-- Replaces the 24-column sparse routing grid in MAIN (7,824 cells holding 821
-- facts, 89.5% empty) with one row per actual process step.
--
-- Apply with:  psql "$DATABASE_URL" -f db/schema.sql

BEGIN;

-- ─────────────────────────────────────────────────────────────────────────────
-- Enums
-- ─────────────────────────────────────────────────────────────────────────────

CREATE TYPE wo_status   AS ENUM ('draft', 'released', 'shipped', 'closed', 'hold');
CREATE TYPE step_status AS ENUM ('open', 'done');


-- ─────────────────────────────────────────────────────────────────────────────
-- processes — lookup, seeded once
--
-- is_outside is the single flag driving the PURCHASING view, replacing the
-- hardcoded outside-process list in cmf_schedule_builder.py.
-- ─────────────────────────────────────────────────────────────────────────────

CREATE TABLE processes (
    code        TEXT    PRIMARY KEY,
    name        TEXT    NOT NULL,
    is_outside  BOOLEAN NOT NULL DEFAULT FALSE,
    sort_order  INT     NOT NULL
);

COMMENT ON COLUMN processes.is_outside IS
    'Vendor/outside work — drives the PURCHASING board';


-- ─────────────────────────────────────────────────────────────────────────────
-- employees — from the SETTINGS sheet; used for "completed by" attribution
-- ─────────────────────────────────────────────────────────────────────────────

CREATE TABLE employees (
    id        BIGSERIAL PRIMARY KEY,
    name      TEXT      NOT NULL UNIQUE,
    role      TEXT,
    active    BOOLEAN   NOT NULL DEFAULT TRUE
);


-- ─────────────────────────────────────────────────────────────────────────────
-- work_orders
--
-- status replaces the ADMIN INPUT / MAIN sheet split *and* the manual move to
-- the HISTORY tab. Nothing is ever relocated — 'shipped' is a state, so a
-- shipped WO can never be confused with one that went missing.
-- ─────────────────────────────────────────────────────────────────────────────

CREATE TABLE work_orders (
    id              BIGSERIAL PRIMARY KEY,
    wo_number       TEXT        NOT NULL UNIQUE,
    po_number       TEXT,
    company         TEXT,
    job_name        TEXT,
    cust_ship_date  DATE,
    status          wo_status   NOT NULL DEFAULT 'draft',
    source          TEXT        NOT NULL DEFAULT 'manual',  -- manual|workbook|history|quickbooks
    notes           TEXT,
    created_by      TEXT,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at      TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX idx_wo_status  ON work_orders (status);
CREATE INDEX idx_wo_company ON work_orders (company);
CREATE INDEX idx_wo_ship    ON work_orders (cust_ship_date);


-- ─────────────────────────────────────────────────────────────────────────────
-- parts
-- ─────────────────────────────────────────────────────────────────────────────

CREATE TABLE parts (
    id              BIGSERIAL PRIMARY KEY,
    work_order_id   BIGINT    NOT NULL REFERENCES work_orders (id) ON DELETE CASCADE,
    line_no         INT       NOT NULL,   -- position on the source sheet; see index below
    part_number     TEXT,
    description     TEXT,
    qty             NUMERIC(12, 2),
    material        TEXT,
    thickness       TEXT,
    screenshot_url  TEXT,
    notes           TEXT,

    -- Reference date shown on every board row, NOT a process step. Modelling it
    -- as a routing_step would put it on the CALENDAR as its own station, which
    -- is the noise the original design deliberately removed (every part would
    -- have an entry).
    delivery_date   DATE,

    -- What the workbook's CURRENT STEP said at import. Retained for traceability
    -- only; live "where is it now" is derived from routing_steps.
    legacy_current_step TEXT,

    created_at      TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at      TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX idx_parts_wo      ON parts (work_order_id);
CREATE INDEX idx_parts_partno  ON parts (part_number);

-- NOT unique. Verified against the live workbook: part_number is not an
-- identifier in this business. WO 4522 carries 18 rows all numbered 'CH' (a
-- profile code) where the distinguishing value lives in DESCRIPTION ('48',
-- '47 1/4', '47 3/8'...). 131 duplicate rows exist across 3 WOs.
--
-- A part row's identity is (work_order_id, line_no) — its position on the
-- original sheet — which is why line_no is carried through the migration.
CREATE UNIQUE INDEX idx_parts_wo_line ON parts (work_order_id, line_no);


-- ─────────────────────────────────────────────────────────────────────────────
-- routing_steps — the core normalization
--
-- One row per process the part actually goes through, replacing 24 sparse
-- columns. ~821 rows for the current book of work.
-- ─────────────────────────────────────────────────────────────────────────────

CREATE TABLE routing_steps (
    id            BIGSERIAL   PRIMARY KEY,
    part_id       BIGINT      NOT NULL REFERENCES parts (id)     ON DELETE CASCADE,
    process_code  TEXT        NOT NULL REFERENCES processes (code),
    due_date      DATE,
    status        step_status NOT NULL DEFAULT 'open',
    completed_by  TEXT,
    completed_at  TIMESTAMPTZ,

    UNIQUE (part_id, process_code),

    -- Status and completion metadata cannot disagree.
    CONSTRAINT completed_fields_consistent CHECK (
        (status = 'done') OR (completed_at IS NULL AND completed_by IS NULL)
    )
);

CREATE INDEX idx_steps_due     ON routing_steps (due_date) WHERE status = 'open';
CREATE INDEX idx_steps_part    ON routing_steps (part_id);
CREATE INDEX idx_steps_process ON routing_steps (process_code);


-- ─────────────────────────────────────────────────────────────────────────────
-- shipments — part-level, multiple releases per PO
--
-- Triggered by invoicing in QuickBooks. invoice_ref is the reconciliation key
-- for the CSV import and makes re-importing an overlapping window a no-op.
-- ─────────────────────────────────────────────────────────────────────────────

CREATE TABLE shipments (
    id            BIGSERIAL   PRIMARY KEY,
    shipped_date  DATE        NOT NULL,
    invoice_ref   TEXT,
    notes         TEXT,
    created_by    TEXT,
    created_at    TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Idempotency for the scheduled CSV import: re-running an overlapping date
-- range cannot create duplicate shipments.
CREATE UNIQUE INDEX idx_shipments_invoice
    ON shipments (invoice_ref)
    WHERE invoice_ref IS NOT NULL;

CREATE INDEX idx_shipments_date ON shipments (shipped_date);


CREATE TABLE shipment_lines (
    id           BIGSERIAL      PRIMARY KEY,
    shipment_id  BIGINT         NOT NULL REFERENCES shipments (id) ON DELETE CASCADE,
    part_id      BIGINT         NOT NULL REFERENCES parts (id),
    qty          NUMERIC(12, 2) NOT NULL CHECK (qty > 0)
);

CREATE INDEX idx_shipment_lines_part ON shipment_lines (part_id);


-- ─────────────────────────────────────────────────────────────────────────────
-- QuickBooks mapping — extraction is easy, matching is the hard part.
--
-- A QB invoice line carries QB's item and customer names, not the schedule's
-- part numbers and work orders. The administrator does this matching mentally
-- today; these tables make it explicit.
-- ─────────────────────────────────────────────────────────────────────────────

CREATE TABLE qb_item_map (
    id           BIGSERIAL PRIMARY KEY,
    qb_item      TEXT      NOT NULL UNIQUE,
    part_number  TEXT,
    notes        TEXT
);

CREATE TABLE qb_customer_map (
    id           BIGSERIAL PRIMARY KEY,
    qb_customer  TEXT      NOT NULL UNIQUE,
    company      TEXT      NOT NULL
);

-- Invoice lines that could not be matched confidently. Nothing leaves the
-- active board from here without a human decision.
CREATE TABLE shipment_review_queue (
    id            BIGSERIAL   PRIMARY KEY,
    invoice_no    TEXT        NOT NULL,
    invoice_date  DATE,
    qb_customer   TEXT,
    qb_item       TEXT,
    po_number     TEXT,
    qty           NUMERIC(12, 2),
    reason        TEXT        NOT NULL,
    resolved      BOOLEAN     NOT NULL DEFAULT FALSE,
    resolved_by   TEXT,
    resolved_at   TIMESTAMPTZ,
    created_at    TIMESTAMPTZ NOT NULL DEFAULT now(),

    UNIQUE (invoice_no, qb_item, qty)
);


-- ─────────────────────────────────────────────────────────────────────────────
-- audit_log — append-only
-- ─────────────────────────────────────────────────────────────────────────────

CREATE TABLE audit_log (
    id         BIGSERIAL   PRIMARY KEY,
    actor      TEXT,
    entity     TEXT        NOT NULL,
    entity_id  BIGINT,
    action     TEXT        NOT NULL,
    before     JSONB,
    after      JSONB,
    at         TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX idx_audit_entity ON audit_log (entity, entity_id);
CREATE INDEX idx_audit_at     ON audit_log (at DESC);


-- ─────────────────────────────────────────────────────────────────────────────
-- Views — the four boards, which used to be ~1,700 lines of Python
-- ─────────────────────────────────────────────────────────────────────────────

-- Shipped vs ordered quantity per part. Partial releases are currently
-- invisible in the spreadsheet; this makes the balance explicit.
CREATE VIEW v_part_progress AS
SELECT
    p.id                                        AS part_id,
    p.work_order_id,
    p.part_number,
    p.qty                                       AS qty_ordered,
    COALESCE(SUM(sl.qty), 0)                    AS qty_shipped,
    p.qty - COALESCE(SUM(sl.qty), 0)            AS qty_remaining,
    (p.qty IS NOT NULL AND COALESCE(SUM(sl.qty), 0) >= p.qty) AS fully_shipped
FROM parts p
LEFT JOIN shipment_lines sl ON sl.part_id = p.id
GROUP BY p.id;


-- CALENDAR: every open in-house step with a due date.
CREATE VIEW v_calendar AS
SELECT
    rs.id           AS step_id,
    rs.due_date,
    pr.code         AS process_code,
    pr.name         AS process_name,
    w.wo_number,
    w.po_number,
    w.company,
    w.cust_ship_date,
    p.id            AS part_id,
    p.part_number,
    p.description,
    p.qty,
    p.delivery_date,
    p.screenshot_url
FROM routing_steps rs
JOIN parts       p  ON p.id = rs.part_id
JOIN work_orders w  ON w.id = p.work_order_id
JOIN processes   pr ON pr.code = rs.process_code
WHERE rs.status = 'open'
  AND w.status IN ('draft', 'released')
  AND pr.is_outside = FALSE;


-- PURCHASING: identical, filtered to vendor work.
CREATE VIEW v_purchasing AS
SELECT
    rs.id           AS step_id,
    rs.due_date,
    pr.code         AS process_code,
    pr.name         AS process_name,
    w.wo_number,
    w.po_number,
    w.company,
    w.cust_ship_date,
    p.id            AS part_id,
    p.part_number,
    p.description,
    p.qty,
    p.delivery_date,
    p.screenshot_url
FROM routing_steps rs
JOIN parts       p  ON p.id = rs.part_id
JOIN work_orders w  ON w.id = p.work_order_id
JOIN processes   pr ON pr.code = rs.process_code
WHERE rs.status = 'open'
  AND w.status IN ('draft', 'released')
  AND pr.is_outside = TRUE;


-- Overdue — the always-visible section at the top of the board.
CREATE VIEW v_overdue AS
SELECT * FROM v_calendar WHERE due_date < CURRENT_DATE
UNION ALL
SELECT * FROM v_purchasing WHERE due_date < CURRENT_DATE;


-- True on-time delivery: actual ship date vs the date promised the customer.
--
-- The old customer report approximated this from LOG step completions, because
-- no ship date was ever recorded anywhere. This only has data from cutover
-- forward; imported HISTORY rows have no shipments and are excluded.
CREATE VIEW v_on_time_delivery AS
SELECT
    w.company,
    COUNT(*)                                                             AS shipments,
    SUM(CASE WHEN s.shipped_date <= w.cust_ship_date THEN 1 ELSE 0 END)  AS on_time,
    ROUND(
        100.0 * SUM(CASE WHEN s.shipped_date <= w.cust_ship_date THEN 1 ELSE 0 END)
        / NULLIF(COUNT(*), 0)
    , 1)                                                                 AS on_time_pct
FROM shipments s
JOIN shipment_lines sl ON sl.shipment_id = s.id
JOIN parts p           ON p.id = sl.part_id
JOIN work_orders w     ON w.id = p.work_order_id
WHERE w.cust_ship_date IS NOT NULL
GROUP BY w.company;


-- ─────────────────────────────────────────────────────────────────────────────
-- Seed: processes
--
-- Column order and is_outside taken from STATIONS / PURCHASING_STATION_KEYS in
-- cmf_schedule_builder.py. Note PAINT is in-house; PAINT (O) is not.
-- ─────────────────────────────────────────────────────────────────────────────

INSERT INTO processes (code, name, is_outside, sort_order) VALUES
    ('MATERIALS',      'MATERIALS',      FALSE,  1),
    ('ENGINEERING',    'ENGINEERING',    FALSE,  2),
    ('LASER_CUT',      'LASER CUT',      FALSE,  3),
    ('LASER_CUT_O',    'LASER CUT (O)',  TRUE,   4),
    ('TUBE_LASER_O',   'TUBE LASER (O)', TRUE,   5),
    ('SAW',            'SAW',            FALSE,  6),
    ('BANDSAW_O',      'BANDSAW (O)',    TRUE,   7),
    ('BEND',           'BEND',           FALSE,  8),
    ('CLEAN',          'CLEAN',          FALSE,  9),
    ('CSK',            'CSK',            FALSE, 10),
    ('DRILL',          'DRILL',          FALSE, 11),
    ('TAPPING',        'TAPPING',        FALSE, 12),
    ('GRIND',          'GRIND',          FALSE, 13),
    ('WELD',           'WELD',           FALSE, 14),
    ('MACHINE_O',      'MACHINE (O)',    TRUE,  15),
    ('PLATING_O',      'PLATING (O)',    TRUE,  16),
    ('PAINT',          'PAINT',          FALSE, 17),
    ('PAINT_O',        'PAINT (O)',      TRUE,  18),
    ('SPECIAL',        'SPECIAL',        FALSE, 19),
    ('SPECIAL_O',      'SPECIAL (O)',    TRUE,  20),
    ('WHOLE_JOB_O',    'WHOLE JOB (O)',  TRUE,  21),
    ('HARDWARE',       'HARDWARE',       FALSE, 22),
    ('ASSEMBLY',       'ASSEMBLY',       FALSE, 23),
    ('SHIP_VENDOR',    'SHIP TO VENDOR', TRUE,  24);

COMMIT;
