# Going Live — a walkthrough

Written for someone who does not write code. You will not need to run anything
on your computer. Total time: **about 20 minutes**, most of it waiting.

Your whole system is defined in a file in this repository (`render.yaml`). The
hosting company reads that file and builds everything automatically. You are
mostly clicking "yes".

---

## What you are about to create

| Thing | What it is | Cost |
|---|---|---|
| **Web service** | The site your team opens in a browser | ~$7/month |
| **Database** | Where the schedule and all screenshots live | ~$6/month |

**About $13/month.** No other services — screenshots live inside the database,
so there is nothing else to sign up for or pay for.

You will need: a **GitHub account** (you have one), a **credit card**, and
**one password you invent** for your team.

---

## Before you start — invent the team password

Everyone in the shop will type this to get in. Write it somewhere safe now.

Make it something you can say out loud across a shop floor but that a stranger
would not guess. Four unrelated words works well: `copper-lathe-tuesday-91`.

> This is one shared password for the whole team, not individual accounts. Every
> action is still recorded against whoever's name is entered, so the log still
> tells you who marked what complete.

---

## Step 1 — Create a Render account

1. Go to **https://render.com**
2. Click **Get Started** → **GitHub**
3. Approve the connection and let it see your repositories
4. Add your credit card under **Billing** (nothing is charged until you create
   the services in step 2)

---

## Step 2 — Deploy everything

1. In the Render dashboard click **New +** (top right) → **Blueprint**
2. Find and select **`MonarkShah1/CMF-Schedule`**
3. Render reads `render.yaml` and shows you what it will create:
   - `cmf-schedule` — the web service
   - `cmf-db` — the database
4. It asks for **one** value: **`APP_PASSWORD`**. Type the team password you
   invented. *(It is hidden and stored encrypted.)*
5. Click **Apply**

Now wait. The first build takes **5–10 minutes**. You will see it install
software, create the database, and load your work orders.

**What to look for in the log** — this is your first deploy loading real data:

```
[setup] no schema found — applying db/schema.sql
[setup] schema applied
[setup] empty database — importing the workbook (first deploy only)
  Loaded.  139 unique screenshot(s) stored in Postgres (4.1 MB)
[setup] import complete
```

When the service goes **green / Live**, click the URL at the top. It looks like
`https://cmf-schedule.onrender.com`.

---

## Step 3 — Check it worked

Open the URL. You should see a sign-in page.

1. Type your name and the team password → **Sign in**
2. You should land on the **Work Orders** list with **83 work orders**
3. Click **Calendar** — process steps with photos
4. Click **Purchasing** — outside/vendor work only
5. Open any work order — parts, routing, and the release button

If all five work, you are live.

---

## Step 4 — Give it to your team

Send them three things:

> **Production schedule:** https://cmf-schedule.onrender.com
> **Password:** *(the one you invented)*
> **Enter your own name when you sign in** — it records who marked work done.

It works on phones and tablets. For the 10-minute standup, open **Calendar** on
a tablet.

---

## Running it alongside the spreadsheet first

**Do not switch everyone over on day one.** Run both for a week:

1. Keep using the Excel/OneDrive workbook as you do today
2. Have the engineer *also* enter routing in the new system
3. At the end of the week, compare the two

When the new system has been right for a week, stop updating the workbook.

The workbook stays in the repository as a snapshot, and the system will **never
re-import it** once real data exists — it checks, and refuses. Your live data
cannot be overwritten by a redeploy.

---

## Everyday questions

**Someone forgot the password.** Render dashboard → `cmf-schedule` →
**Environment** → edit `APP_PASSWORD` → Save. It redeploys in about two minutes
and everyone signs in again with the new one.

**Is it backed up?** Render backs the database up daily on the paid plan.
Screenshots are inside the database, so they are covered too.

**Someone deleted something by mistake.** Nothing is truly deleted — releasing,
un-releasing and routing changes are all written to an audit log. Ask for help
and it can be traced.

**It is slow the first time each morning.** On the `starter` plan it should not
sleep. If it does feel slow, check you are not on the free plan.

**We want our own web address** (like `schedule.canadianmetal.ca`). Render →
`cmf-schedule` → **Settings** → **Custom Domain**. You will need to add one
record with whoever hosts your domain.

---

## Before you make a change later

Any time the system is changed, run the full test suite first:

```bash
./v2/run_tests.sh
```

It builds a throwaway database, loads your real workbook, and runs **53 tests
plus 16 browser checks** covering every phase. It touches nothing live. If it
says `✓ Full regression passed`, the change is safe to deploy.

This is the thing the old system never had, and it is why the old one kept
breaking.

---

## What is still to come

Live now: order entry, routing, the release gate, and both boards.

| Phase | What it adds |
|---|---|
| 4 | Marking work complete from the board + end-of-day sign-off |
| 5 | Recording shipments + one-way mirror to the boss's Google Sheet |
| 6 | Excel export button, so anyone can still get a spreadsheet |
| 7 | QuickBooks invoice CSV import to automate the move to HISTORY |

Until phase 5 lands, keep updating the boss's Google Sheet by hand as you do
today.
