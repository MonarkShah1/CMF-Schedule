# Workflows

---

## Daily — Engineer Updates Routing

1. Open `CMF WIP - Schedule.xlsx`
2. Go to **MAIN** tab
3. For each part, fill in the blue process columns (L–AI) with due dates
4. Update **CURRENT STEP** (col AJ) dropdown as parts move through the shop
5. Paste screenshot into **col I** for visual reference
6. Save and close Excel
7. Double-click `Refresh Schedule.bat`

CALENDAR and TODAY update automatically.

---

## Daily — Production Manager

1. Open `CMF WIP - Schedule.xlsx` → go to **CALENDAR**
2. Check the overdue section at the top first — these are blocked jobs
3. Scan the week's date buckets — see which stations are busy
4. Use **PROCESSES LEFT** column to understand the full remaining path for each part
5. Go to **TODAY** tab to see the TV-ready kanban view

---

## Weekly — Boss Sends Updated WIP

When the boss updates dates or adds new orders in the original WIP Google Sheets:

1. Boss exports as `CMF WIP (1).xlsx` → saves to the folder
2. Engineer's latest Schedule file saved as `CMF WIP - Schedule (N).xlsx`
3. Run:
   ```
   python3 cmf_merge.py
   python3 cmf_schedule_builder.py
   ```
4. New orders appear in MAIN with blank routing (navy header + grey part rows)
5. Engineer fills in routing for the new parts

---

## Adding a New Work Order (Manual)

If needed without running the merge script:

1. Go to **MAIN** tab
2. Insert a **navy header row** — fill cols A (WO#), B (PO#), C (Company), D (Job Name), E (Cust Ship)
3. Add one or more **grey part rows** below — fill cols F (Part#), G (Description), H (Qty)
4. Fill in blue process columns (L–AI) for each part
5. Set CURRENT STEP dropdown (col AJ)
6. Run `Refresh Schedule.bat`

---

## Marking a Job Complete

1. Go to the part rows for that WO in **MAIN**
2. Set **STATUS** (col AK) to `COMPLETE` for each part row
3. Run `Refresh Schedule.bat`

The part will disappear from CALENDAR and TODAY automatically.

---

## Updating Current Step

As a part moves through the shop:

1. Find the part row in **MAIN**
2. Click the **CURRENT STEP** cell (col AJ)
3. Select the new step from the dropdown
4. Run `Refresh Schedule.bat`

The calendar will automatically hide completed processes (anything before the current step's date) and show only remaining work.

---

## TV Display Setup

1. Connect the factory floor TV/monitor to a computer
2. Open `CMF WIP - Schedule.xlsx`
3. Navigate to the **TODAY** tab
4. Set Excel to full-screen / presentation mode (`View → Full Screen`)
5. Run `Refresh Schedule.bat` each morning to update
