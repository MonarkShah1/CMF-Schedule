# CMF Schedule v2 — Target Architecture

Replaces the Excel-on-OneDrive system. The **workflow does not change** — it has
been running three months and works. What changes is the storage layer that
keeps failing.

---

## Decisions

| Question | Decision |
|---|---|
| Hosting | **Cloud-hosted** — managed Postgres + web app |
| Build | **Custom web app** — code lives in this repo |
| Boss's Google Sheet | **Read-only mirror**, one-way push from the app |
| QuickBooks Desktop | **Later phase** — design the ingestion seam now, build it later |

### Why not Excel + OneDrive

An `.xlsx` is a single binary blob. Two people editing means two copies of the
*entire workbook*, and OneDrive cannot merge them. Co-authoring only holds for
simple sheets in Excel Online, breaks on embedded images, and breaks completely
when a Python script rewrites the file underneath an open session. No OneDrive
setting fixes this. Concurrency has to move into a database.

### Why the rebuild is smaller than it sounds

Measured on the live workbook:

| | |
|---|---|
| Part rows | 326 |
| Routing grid cells (326 × 24 processes) | 7,824 |
| Cells actually holding a date | **821** |
| Grid utilization | **10.5% — 89.5% empty** |
| Avg process steps per part | **2.5 of 24** |

The 24-column sparse grid is what most of `cmf_schedule_builder.py` exists to
navigate — `STEP_TO_COL`, `COL_TO_PROCESS`, column-shift migrations, row-kind
detection, date-based filtering. Normalized it is **821 rows in one table**, and
that code stops existing.

---

## Data model

```
work_orders
  id, wo_number, po_number, company, job_name,
  cust_ship_date, status(draft|released|shipped|closed|hold),
  created_at, created_by

parts
  id, work_order_id, part_number, description, qty,
  material, thickness, screenshot_url

processes                          -- lookup, seeded once
  code, name, is_outside, sort_order

routing_steps                      -- ~821 rows, replaces the 24-col grid
  id, part_id, process_code, due_date,
  status(open|done), completed_by, completed_at

shipments                          -- one release; see "Shipping" below
  id, shipped_date, invoice_ref, notes, created_by, created_at

shipment_lines                     -- shipping is PART level, not WO level
  id, shipment_id, part_id, qty

audit_log
  id, actor, entity, entity_id, action, before, after, at
```

`processes.is_outside` is the single flag that drives the PURCHASING view —
replacing the hardcoded outside-process list in the current scripts.

### The four views become queries

| View | Query |
|---|---|
| **CALENDAR** | `routing_steps WHERE status='open'` grouped by `due_date` |
| **PURCHASING** | same, `JOIN processes WHERE is_outside = true` |
| **Processes left** | `WHERE part_id = ? AND status='open' ORDER BY due_date` |
| **Overdue** | `WHERE status='open' AND due_date < today` |

Ordering is free. No column math, no date-based filtering hack, no
`COL_TO_PROCESS` map.

### Two bug classes deleted permanently

- **Completion status** is `routing_steps.status`, a real value — not a cell
  fill colour read back with an RGB threshold. The recurring green-sync fixes
  in the git history become impossible.
- **Screenshots** are one file in object storage referenced by URL — not 920
  stored copies of 139 images (the current file writes one photo up to 153
  times, inflating the workbook 4.3 MB → 17.7 MB per refresh).

---

## Roles

| Role | Can do |
|---|---|
| Administrator | Create work orders and parts, record shipments |
| Engineer | Fill routing, set due dates, release orders to production |
| Production supervisor | Mark steps done |
| Project manager | Mark steps done, review, sign off |
| Purchasing | View + update outside/vendor steps |
| Boss | Read everything |

### The review gate is preserved

Today: new orders land in ADMIN INPUT, and move to MAIN once they have a
CURRENT STEP and at least one process date.

v2: that's `work_orders.status`. `draft` → engineer fills routing → **release**
→ `released`. Same gate, same human check, no sheet-shuffling and no file copy.
Your manual refresh becomes an explicit action instead of a save-and-sync.

---

## Shipping — the open piece

### How it actually works today

- **There is no ship date column.** Nothing gets written when something ships.
  The *row moving to HISTORY* is the entire record that it shipped.
- **The trigger is invoicing.** The administrator invoices in QuickBooks;
  invoiced means shipped, and the row comes off the WIP tracker.
- **Shipping is part level**, and most POs go out in **multiple releases** —
  so a WO is partly shipped for long stretches.
- **HISTORY is the identical layout**, which is why moving a row is a copy-paste.

Two consequences worth naming:

1. **No ship date exists anywhere in your system.** Not in the active tab, not
   in HISTORY. So on-time delivery genuinely cannot be computed from the sheet
   — it isn't a matter of nothing reading it, the data was never captured. The
   only place a real ship date exists today is the QuickBooks invoice date.
2. **Partial releases are invisible.** A row sits in the active tab, fully
   present, whether zero or 80% of its quantity has gone out. Nothing shows
   remaining balance until the last release drops it into HISTORY.

### Design

Shipping is a first-class event recorded at part level, with releases as
first-class rows:

```
Administrator invoices in QuickBooks
        ↓
Opens the WO  →  "Record shipment"
   shipped_date, invoice_ref, then per part: qty this release
        ↓
   shipments + shipment_lines rows written
        ↓
   part fully shipped  when SUM(shipment_lines.qty) >= parts.qty
   WO   status='shipped' when every part is fully shipped
        ↓
   mirror moves the row from the active tab to HISTORY
```

`invoice_ref` is the link back to QuickBooks. Capturing it by hand now costs the
administrator a few seconds and becomes the reconciliation key if the QuickBooks
agent is ever built.

**New capability this unlocks:** a part-level shipped/remaining balance, so the
board can show "40 of 100 shipped" instead of a row that looks untouched until
the day it disappears.

`SHIP` already exists in the current CURRENT STEP dropdown, so shipping is
partly modelled as a routing step today. Keeping a `SHIP` routing step for
scheduling *and* a `shipments` record for what actually left the building is
deliberate — one is a plan, the other is a fact.

### History: shipped POs move to a HISTORY tab

Today a shipped PO's row is **moved** from the active sheet into a HISTORY tab.
Two problems come out of that, both real:

**1. The merge cannot tell "shipped" from "missing."** `cmf_merge.py` rebuilds
MAIN only from the WOs present in the export, and reports anything absent as
`Removed WOs (Missing From Google)`. A WO that shipped and a WO that vanished
because the wrong export was loaded look *identical* to the script — both just
disappear from the schedule. Combined with the stale-file bug fixed earlier
(where a week-old export could be picked silently), this could resurrect shipped
orders and drop live ones in the same run.

**2. Your real delivery performance cannot be computed at all.**
`cmf_customer_report.py` derives on-time percentage from **LOG completions** —
whether *process steps* hit their due dates. That is a proxy for on-time
delivery, not the thing itself. And it can't be corrected from the sheet,
because **no ship date is recorded anywhere** — not in the active tab, not in
HISTORY. The only place a true ship date exists is the QuickBooks invoice.

From cutover onward v2 captures `shipments.shipped_date`, so the real metric
starts accumulating immediately. Backfilling history would require reading
QuickBooks invoice dates (see below).

**In v2, history is a status, not a location.** Nothing moves.

```sql
-- active board
WHERE work_orders.status IN ('draft','released')
-- history
WHERE work_orders.status IN ('shipped','closed')
```

Because the row never moves, a shipped order is explicitly `shipped` rather than
absent — the ambiguity above cannot occur. And real on-time delivery becomes a
one-line query:

```sql
SELECT company,
       AVG(CASE WHEN s.shipped_date <= w.cust_ship_date THEN 1.0 ELSE 0 END) AS on_time_pct
FROM shipments s JOIN work_orders w ON w.id = s.work_order_id
GROUP BY company;
```

That is the metric the customer report has been approximating. The mirror keeps
pushing shipped rows to the HISTORY tab so the boss's sheet looks unchanged.

### Confirmed

| Question | Answer |
|---|---|
| Which column receives the ship update? | **None.** Moving the row to HISTORY *is* the record. |
| Who records the shipment? | The administrator |
| What triggers it? | **Invoicing in QuickBooks** — invoiced means shipped |
| Partial shipments? | **Yes** — most POs ship in multiple releases |
| Ship level? | **Part level** |
| HISTORY layout? | Identical to the active tab — a copy-paste |

---

## Google Sheets mirror

One-way, app → sheet. One-way is the whole point: two-way sync recreates the
conflict problem in a new place.

- Push on change (or every few minutes) via the Sheets API with a service account
- **Mirrors both tabs** — active WIP and HISTORY. When a WO becomes `shipped`,
  the mirror removes it from the active tab and appends it to HISTORY, exactly
  as someone does by hand today.
- Explicit column mapping in config, not positional — the current positional
  mapping is why four columns silently went unread
- The sheet gets protected ranges so edits there don't create a phantom second
  source of truth
- If a push fails, it retries and surfaces in the app; the sheet is never the
  thing that blocks production

The boss sees the same two tabs behaving the same way. The difference is that
moving a row between them is now a *consequence* of recording a shipment, not a
manual step someone has to remember.

The boss keeps his tracker, unchanged, with no retraining.

---

## QuickBooks Desktop — designed for, built later

QB Desktop has no cloud REST API. The available paths are the SDK
(qbXML/QBFC), the Web Connector, ODBC drivers, or scheduled report export —
and all of them want to run **on the same machine or LAN as QuickBooks**. That
is why the earlier attempt failed; Web Connector in particular is notoriously
fiddly.

Since hosting is cloud, the eventual shape is:

```
QB Desktop (shop PC)
   └── small local agent (SDK or ODBC, read-only)
          └── HTTPS push  →  POST /api/orders/import  →  work_orders (status='draft')
```

**What to build now:** the order-ingestion seam — a single documented endpoint
that creates draft work orders, with the manual entry form as its first client.
A QuickBooks agent later becomes a second client of the same endpoint. No
rewrite.

### When you do build it, read invoices first — not orders

Your answer that **invoicing is the ship trigger** changes the priority. The
original plan was to pull *orders* out of QuickBooks. Pulling **invoices** is a
better first integration on every axis:

| | Order ingestion | **Invoice ingestion** |
|---|---|---|
| Direction | Read | **Read only** — never writes to QuickBooks |
| Risk to accounting data | Low | **None** |
| Query | Complex, needs mapping to parts | `InvoiceQuery` by date range |
| Replaces | Admin typing new orders | **Admin manually moving rows to HISTORY** |
| Also gives you | — | **Real ship dates → true on-time delivery** |

A read-only invoice pull via the SDK (`InvoiceQuery`) or a QODBC read is about
the smallest useful QuickBooks Desktop integration that exists, and it closes
the loop that is currently manual:

```
QB Desktop invoice created
   └── local agent polls InvoiceQuery (read-only)
          └── HTTPS  →  POST /api/shipments/from-invoice
                 └── matches invoice_ref → shipment_lines → parts shipped
                        └── mirror moves the row to HISTORY automatically
```

If the agent is never built, nothing breaks — the administrator keeps recording
shipments by hand, exactly as today, and `invoice_ref` is already there as the
matching key.

**Also worth checking:** Intuit has been sunsetting QuickBooks Desktop in favour
of QuickBooks Online, which *does* have a clean REST API. Confirm where your
version sits on that timeline before investing in a Desktop-only integration —
it may decide this for you.

---

## Migration off the current workbook

The existing `CMF WIP - Schedule.xlsx` is the source. A one-time importer:

1. Parse MAIN → `work_orders` + `parts` (reuse the existing row-kind logic)
2. Walk the 24 process columns → **821 `routing_steps` rows**
3. Extract the 139 unique screenshots → object storage, set `parts.screenshot_url`
4. Derive step status from CURRENT STEP + the date-based rule already in
   `parse_main()`
5. Import the LOG sheet → `audit_log` so completion history survives
6. Reconcile: part count, step count, and image count must match the workbook

Run it against a copy, compare against the live sheet, then cut over. Keep the
Excel export so nobody loses the artifact they're used to.

---

## Build order

| Phase | Deliverable |
|---|---|
| 1 | Schema + importer, reconciled against the live workbook |
| 2 | Order entry (admin) + routing entry (engineer) + release gate |
| 3 | CALENDAR + PURCHASING views — the 10-minute standup board |
| 4 | Mark-complete + LOG/sign-off |
| 5 | Shipping + Google Sheets mirror |
| 6 | Excel export button |
| 7 | *(later)* QuickBooks agent against the phase-2 ingestion endpoint |

Phases 1–4 replace what hurts daily. Phase 5 closes the shipping loop. Phase 7
only happens if QuickBooks is still on Desktop and still worth it.

---

## What this fixes

| Today | v2 |
|---|---|
| Merge conflicts, close every Excel window | Concurrent users, handled by Postgres |
| Status stored as cell fill colour | A real status column |
| 920 image copies, 17.7 MB per refresh | One file per image in object storage |
| 7,824-cell grid, 89.5% empty | 821 rows |
| Wrong input file silently merged | No files to pick |
| No audit trail beyond the LOG sheet | Append-only `audit_log` |
| Shop floor needs a PC with Excel | Any browser, including a tablet at standup |
| "Shipped" and "missing" look identical | An explicit `shipped` status |
| Shipped POs moved by hand to HISTORY | A consequence of recording the shipment |
| No ship date captured anywhere | `shipments.shipped_date` from day one |
| Partial releases invisible on the board | Part-level shipped/remaining balance |
| On-time % derived from step completions | Real ship date vs promised date |
| QuickBooks integration impossible | A documented endpoint waiting for it |
