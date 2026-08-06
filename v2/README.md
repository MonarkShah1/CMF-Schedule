# v2 — Phase 1: schema + importer

Replaces the Excel-on-OneDrive system. See
[`../docs/ARCHITECTURE-V2.md`](../docs/ARCHITECTURE-V2.md) for the full design and
[`../docs/POSTMORTEM.md`](../docs/POSTMORTEM.md) for why.

Phase 1 is the database schema and the one-time migration off the workbook.
No application yet.

---

## Try it

```bash
pip install openpyxl 'psycopg[binary]'

# Parse and reconcile only — touches no database
python3 v2/import_workbook.py --dry-run

# Inspect exactly what would be loaded
python3 v2/import_workbook.py --dry-run --out /tmp/import.json

# For real
createdb cmf
psql cmf -f db/schema.sql
python3 v2/import_workbook.py --database-url postgresql:///cmf

# With the Google Sheets HISTORY export
python3 v2/import_workbook.py --history "CMF WIP - HISTORY.xlsx" --database-url postgresql:///cmf
```

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

Phase 2 — order entry, routing entry, and the release gate.
