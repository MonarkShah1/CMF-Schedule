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

shipments                          -- see "Shipping" below
  id, work_order_id, shipped_date, qty, reference, notes, created_by

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

**Today:** when a PO ships, someone writes it into the Google Sheet. That is a
*write path into the sheet*. Making the sheet read-only removes it, so the
shipping update has to happen in the app first and mirror outward.

**Design:** shipping is a first-class event, not just a routing step. It needs a
date and quantity, it must survive partial shipments, and it is the natural
trigger for both the sheet mirror and eventual QuickBooks invoicing.

```
Administrator opens the WO  →  "Record shipment"
   shipped_date, qty, reference/packing slip, notes
        ↓
   row written to `shipments`
        ↓
   WO status → 'shipped' when total shipped qty >= ordered qty
        ↓
   mirror pushes ship date + status to the Google Sheet
```

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

**2. Your real delivery performance is invisible to the system.**
`cmf_customer_report.py` computes on-time percentage from **LOG completions** —
whether *process steps* hit their due dates. Actual ship date versus promised
customer ship date lives in the HISTORY tab, which nothing in this codebase
reads. The number on the customer report is a proxy, not on-time delivery.

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

### Confirm before building

1. **Which column** in the boss's sheet receives the ship update? The current
   pipeline reads only columns A, B, D, G, H, I, J, K, L, N, O — **C, E, F, and
   M are ignored entirely.** If shipping lands in one of those, the automation
   has never seen it, which may itself explain some drift.
2. **Who** records the shipment — administrator, or a shipping clerk who'd need
   their own login?
3. **Partial shipments** — does a WO ever ship in multiple releases? The schema
   above assumes yes. If it's always all-or-nothing, `shipments` collapses to
   two fields on `work_orders`.
4. **Ship at PO level or part level?** Modelled at WO/PO level above, matching
   how the sheet appears to work.
5. **What does the HISTORY tab hold** that the active tab doesn't — an actual
   ship date, carrier, invoice number? That determines whether `shipments`
   needs more fields, and it is the data the on-time metric should be using.

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
| On-time % derived from step completions | Real ship date vs promised date |
| QuickBooks integration impossible | A documented endpoint waiting for it |
