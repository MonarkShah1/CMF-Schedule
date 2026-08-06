# Automation Postmortem — Why the Order-Entry Automation Was Turned Off

Written after reviewing the full repo and running the scripts against the real
workbook. Everything in "What actually broke" was reproduced or measured, not
guessed.

---

## The short version

The thing that was turned off was never really an automation. It was a set of
local Python scripts a person ran by hand, with an AI session acting as the
glue and the repair crew every time something looked wrong. The scripts were
the cheap part. The repair crew was the expensive part.

Three things made it unreliable, and they compound:

1. **The merge script could silently run on last week's orders.** Reproduced.
2. **Every refresh quadrupled the file size.** Measured: 4.3 MB → 17.7 MB.
3. **Completion status was stored as cell colour**, read back with a guess.

None of these announce themselves. They produce a workbook that looks fine and
is subtly wrong — which is exactly the failure mode that forces a human to
re-check everything, and forces an AI session to re-read a 17 MB binary to find
out what happened.

---

## What actually broke

### 1. The merge picked the wrong input file (silent, data-losing)

`cmf_merge.py` chose its input by the `(N)` in the filename, preferring the
highest number. Real-world sequence:

```
Monday    download the boss's sheet  →  CMF WIP.xlsx
          download it again          →  CMF WIP (1).xlsx
Friday    clear Downloads, re-export →  CMF WIP.xlsx     ← today's orders
          run cmf_merge.py           →  picks CMF WIP (1).xlsx   ← Monday's
```

Reproduced directly: with a fresh `CMF WIP.xlsx` and a stale `CMF WIP (1).xlsx`,
the script chose the stale one. No warning, no output naming the file it used.

The merge then **overwrites the working schedule in place** with a result built
from stale orders — taking the engineer's routing with it. There were no
backups anywhere in the repo, so the only recovery was OneDrive version
history, if anyone noticed in time.

This alone is enough to explain "we couldn't make it work consistently."

### 2. Every refresh inflated the workbook ~4x

Measured on the real file:

| | media parts | unique images | size |
|---|---|---|---|
| Committed file | 139 | 139 | 4.3 MB |
| After one refresh | **920** | 139 | **17.7 MB** |

openpyxl writes one stored copy per placed image. The builder puts each
screenshot on the MAIN row *and* on every CALENDAR card *and* every PURCHASING
card for that part — so one photo gets written up to **153 times**. 85% of the
file was redundant copies of the same JPEGs.

Downstream effects, all of which read to a user as "it doesn't work":

- Slow to open, slow to save, Excel feels stuck
- A 17 MB file re-syncing to OneDrive on every refresh — slow, and a good way
  to generate sync conflicts when two people touch it the same day
- Every commit of the workbook added ~17 MB to the repo
- Any AI session handed this repo starts by ingesting a multi-megabyte binary

### 3. Completion status lives in cell colour

`_is_user_green()` decides whether a job is done by reading the fill colour and
testing `g >= 80 and g >= r + 8 and g >= b + 8`, plus a special case for theme
index 6.

Using presentation as the database means it breaks on things nobody would think
to report:

- A different Office theme — index 6 is no longer that green
- A green from the palette that doesn't clear the threshold
- A cell highlighted green for *any other reason* → phantom completion
- Picking the **same** green the template uses → silently ignored, because
  template colours are excluded to avoid false positives

The git history shows this being patched repeatedly:
`fix(green-sync): handle theme colors`, `fix refresh stale green card row
mapping`, `fix refresh completion grouping`, `fix: accept green completions
without updater`. Each fix was correct; the approach guarantees more of them.

### 4. MAIN is destroyed and rebuilt on every run

`sync_admin_input_flow()` deletes the MAIN sheet and recreates it from parsed
values each time. Anything the parser doesn't model is lost — and parts move
between MAIN and ADMIN INPUT automatically based on
`part_is_ready_for_main()`. Clear a CURRENT STEP and the part *silently
teleports* out of MAIN on the next refresh. To a user, work vanished.

### 5. The review report was built and then thrown away

`cmf_merge.py` computed exactly what a reviewer needs — new WOs, removed WOs,
ship-date changes, parts still missing routing — assembled it into a
`build_change_log_sheet()` function, and then **never called it**. 70 lines of
finished, dead code. The tab was removed to keep the boss's workbook clean, and
the report went with it.

So the review step — "I come in and review before pushing to OneDrive" — meant
eyeballing 330 part rows with no diff. That is slow, and it is why review
became the bottleneck.

### 6. No tests, no dry run, no backups

No test files. No `--dry-run`. No `shutil.copy` anywhere. Both scripts
overwrite the working file directly. Every run was a live run against the only
copy.

---

## Where the tokens actually went

There is no Google Sheets API client, no OneDrive/Graph client, no scheduler,
and no `.github/workflows` in this repo. Nothing here runs on its own.

So the "automation" in practice was:

```
admin types orders into Google Sheets
   ↓  (human exports, downloads, renames)
scripts run locally on someone's machine
   ↓  (something looks wrong)
AI session opens the repo to investigate
   ↓  reads a 71,000-character builder + a multi-MB binary workbook
   ↓  patches, reruns, re-reads
you review the whole sheet by hand
   ↓
push to OneDrive
```

The token cost wasn't the order entry. It was using a language model as the
runtime and the debugger for a system that had no backups, no tests, no dry
run, and no diff — so every investigation started from zero and had to re-read
everything. 25 of 44 commits touch the 4.3 MB workbook; `.git` is 32 MB.

**An AI agent should be writing the tooling, not being the tooling.**

---

## What was fixed in this pass

All four changes are in and tested against the real workbook.

| Fix | Result |
|---|---|
| `_pick_latest()` now selects by **modification time**, prints every candidate and which one it used, and warns if the newest file is >48h old or if two files are within 2 minutes of each other | The stale-file failure is gone and the choice is visible |
| **Timestamped backups** before every in-place overwrite, in `backups/`, last 30 kept | Any bad run is recoverable in seconds |
| **`cmf_dedupe_media.py`** runs automatically after save | **17.7 MB → 4.0 MB, 77% smaller.** Verified: all 920 image placements intact, 0 broken, same 139 distinct images, stable across repeat runs |
| **`Merge Review.xlsx`** — the dead change-report code now writes to a standalone file, plus `--dry-run` on the merge | You review a diff, not 330 rows. The working file stays clean |

Verified alongside: a repeat refresh **is** idempotent — second run detects 0
completions and produces identical counts (330 parts, 616 entries, 311
screenshots). The instability was in the inputs and the colour reading, not in
the refresh loop itself.

---

## Moving forward

### Now — the workflow to use

```bash
python3 cmf_merge.py --dry-run   # writes Merge Review.xlsx, changes nothing
#   open Merge Review.xlsx: new WOs, removed WOs, ship-date changes,
#   parts still needing routing
python3 cmf_merge.py             # backs up, then merges
python3 cmf_schedule_builder.py  # backs up, rebuilds, dedupes
#   review, then push to OneDrive
```

Two rules that prevent most of what went wrong:

- **Keep exactly one `CMF WIP*.xlsx` in the folder.** Delete the old export
  before downloading a new one. The script now warns, but one file can't be
  ambiguous.
- **Close Excel before running.** An open file causes
  `KeyError: [Content_Types].xml`.

### Next — worth doing, in order of payoff

1. **Stop storing status in cell colour.** Add a real STATUS column with a
   dropdown (`DONE` / `IN PROGRESS` / `HOLD`). Keep the green fill as
   *conditional formatting driven by that column* — the board looks identical,
   but the data is a value, not a pixel. This kills an entire recurring class
   of bug.
2. **Stop committing the workbook to git.** It's data, not source. Put it on
   OneDrive where it already lives and keep this repo to scripts. That shrinks
   every future AI session's context and stops `.git` from growing 17 MB a run.
3. **Add a handful of tests** over a small fixture workbook — 20 rows, 3 images.
   Test: merge picks the newest file; routing survives a merge; a part with a
   cleared CURRENT STEP behaves as intended; dedupe preserves placements. This
   is what makes a future fix a 5-minute job instead of an investigation.
4. **Then, if you want it truly automatic**, the honest version is a scheduled
   job with the Google Sheets API reading the boss's tracker directly — no
   manual export, no filename ambiguity. That's a real project, but it's only
   worth doing *after* 1–3, because it removes the human who is currently
   catching these failures.

### On using AI here

Use it to build and change the tooling — new columns, a new station, a report,
the test suite. That's a bounded task with a reviewable diff.

Don't use it as the daily runtime. A script that runs the same way every day
costs nothing to run; a model that re-reads the whole system every day costs
what you were paying. The scripts should be boring and deterministic, and the
model should only show up when you want them to do something new.
