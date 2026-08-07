# v2 — Phases 1–2

Replaces the Excel-on-OneDrive system. See
[`../docs/ARCHITECTURE-V2.md`](../docs/ARCHITECTURE-V2.md) for the full design and
[`../docs/POSTMORTEM.md`](../docs/POSTMORTEM.md) for why.

| Phase | Status |
|---|---|
| 1 — schema + importer | **done** |
| 2 — order entry, routing entry, release gate | **done** |
| 3 — CALENDAR + PURCHASING boards | next |
| 4 — mark complete + sign-off | |
| 5 — shipping + Google Sheets mirror | |
| 6 — Excel export | |
| 7 — CSV invoice import + review queue | |

No build step. One Python process plus Postgres — no npm, no bundler.

---

## Try it

```bash
pip install -r v2/requirements.txt

createdb cmf
psql cmf -f db/schema.sql

# Migrate off the workbook — parse and reconcile first, touching nothing
python3 v2/import_workbook.py --dry-run
python3 v2/import_workbook.py --database-url postgresql:///cmf

# With the Google Sheets HISTORY export
python3 v2/import_workbook.py --history "CMF WIP - HISTORY.xlsx" \
                             --database-url postgresql:///cmf

# Run the app
DATABASE_URL=postgresql:///cmf uvicorn v2.app.main:app --reload
#   → http://127.0.0.1:8000

# Tests (needs a database with the schema applied)
DATABASE_URL=postgresql:///cmf pytest v2/tests -q
```

---

## Phase 2 — what it does

| Route | Who | Purpose |
|---|---|---|
| `/` | everyone | Work order list, filter by status, search |
| `/orders/new` | administrator | Create an order — lands as **draft** |
| `/orders/{id}` | administrator | Add and remove parts |
| `/parts/{id}/routing` | engineer | Set due dates per process |
| `/orders/{id}/release` | engineer | **Release gate** — draft → released |

### The release gate

The workbook moved parts between ADMIN INPUT and MAIN automatically, based on
whether they had a CURRENT STEP and a process date. Clear a field and a part
silently teleported out of MAIN — work appeared to vanish.

Here the gate is explicit:

- A work order releases only when **every** part has at least one routing step
- When it can't, the page names the exact lines that are blocking it
- The Release button is disabled rather than failing after the click
- Release is **reversible** — "pull back to draft" needs no database surgery
- Every release, un-release and routing change lands in `audit_log`

### Verified

19 tests pass against a real Postgres, and the full workflow was driven through
a real browser end to end: create order → blocked (no parts) → add part →
blocked (no routing, line named) → route → "ready to release" → released.

Behaviours covered by tests:

- an order with no parts, or any unrouted part, cannot be released
- one unrouted part blocks the whole order, and is named
- release is reversible; double-release is rejected
- blank dates create no routing step; clearing a date removes the step
- **a completed step survives a cleared date** — completion history is never
  silently discarded
- delivery date never becomes a routing step
- duplicate part numbers within a work order are accepted (`CH` × 18 is real)
- deleting a part removes its routing

### Not yet real

Identity is a name in a cookie, used for `audit_log` attribution. It is not
authentication and enforces no permissions. Real auth comes with the boards,
when there is something worth protecting.

---

## Verified against the live workbook

Applied to a real Postgres and loaded end to end:

| | |
|---|---|
| work_orders | 83 |
| parts | 337 |
| routing_steps | **835** |
| screenshots | 313 placements → **139 unique files** |
| audit_log (from LOG) | 155 |
| employees (from SETTINGS) | 11 |

All reconciliation checks pass: routing cells equal routing_steps, part rows
equal parts, no orphans, no work order without parts.

### Reconciles exactly against the old builder

The old builder reported 616 process-due entries. The v2 views return 646 open
steps. That difference is fully accounted for:

```
646  open steps (v2, read from the raw workbook)
 -28  green completions the old builder applied before it counted
 -2   steps due beyond its 4-week window
────
616  matches the old builder exactly
```

The views deliberately carry no date window — that is a UI concern, not a data
one.

---

## Two schema decisions the real data forced

**Delivery date is not a process step.** It is `parts.delivery_date`. Modelling
it as a routing step added 335 rows and would have put DELIVERY on the CALENDAR
as its own station — the exact noise the original design removed, since every
part would have an entry. It shows as a reference column on each board row.

**`part_number` is not unique within a work order.** Verified: WO 4522 carries
18 rows all numbered `CH` — a profile code — where the real distinguishing value
sits in DESCRIPTION (`48`, `47 1/4`, `47 3/8`…). 131 duplicate rows exist across
3 WOs. A part row's identity is `(work_order_id, line_no)`, its position on the
source sheet, so `line_no` is carried through the migration.

> This matters for the QuickBooks CSV import in phase 7: invoice lines cannot be
> matched to parts on part number alone. The review queue is not optional.

---

## What the importer does

1. `MAIN` → `work_orders` + `parts`, reusing the builder's header/part row-kind rule
2. 24 process columns → **one `routing_steps` row per actual step**
3. `ADMIN INPUT` → parts on work orders with `status='draft'` (the unrouted backlog)
4. Step status derived with the workbook's own date rule — any step due before
   the CURRENT STEP's due date is `done`; `STATUS=COMPLETE` marks all steps done
5. Screenshots deduplicated by content hash → one file per unique image
6. `LOG` → `audit_log`, `SETTINGS` → `employees`
7. HISTORY export → `work_orders` with `status='shipped'` and **no shipments**,
   because no ship date was ever recorded

---

## Next

Phase 3 — the CALENDAR and PURCHASING boards for the 10-minute standup.
