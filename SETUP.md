# Setup Guide (Plain-English)

This guide walks you through setting up the **Ek Tara Attendance System** step by step.
You do **not** need to be a programmer. Just follow the steps in order.

Once it's running, this system will:
- Email each teacher a **daily attendance form** every morning.
- Automatically **record the answers** into a spreadsheet for each class.
- Keep **charts and summaries** up to date.
- Send a **weekly report** to the people in charge.

> Need to start over later? See the companion file **`RESET.md`**.

---

## Before you start — what you'll need

- A Google account that will own everything (forms, spreadsheets, emails go out from it).
- The list of teachers and which class each one handles.
- A student list (roster) for each class.

---

## Step 1 — Create 3 folders in Google Drive

Make these three folders (any names you like) and keep their web links handy:

1. **Rosters folder** — where you'll put the student lists.
2. **Attendance folder** — leave this **empty**; the system fills it automatically.
3. **Settings folder** — where you'll put the teacher list and holiday list.

To get a folder's link: open the folder in Google Drive → copy the address from the
browser's address bar.

---

## Step 2 — Add your files to the folders

### In the "Rosters" folder — one file per class

Put one student list per class-section. It can be an **Excel file, a CSV, or a Google
Sheet**.

- **The file name must include the class number and section.** Examples that work:
  - `Class 1 A`
  - `Class 5 B`
  - `Grade 7 C`
- **Inside the file**, use a header row with these columns (in any order):
  **Roll No.**, **Child ID**, **Name of Student**, **Status**. Set Status to
  **Active** or **Inactive**. Inactive students retain their historical attendance
  but are excluded from new forms and current attendance-risk alerts.

### In the "Settings" folder — the teacher list (required)

Create a file named **`TeacherClassMapping`** with these columns:

| Name | Email | Class | Section |
|------|-------|-------|---------|
| Meena | meena@example.com | 1 | A |
| Ravi | ravi@example.com | 5 | B |

This tells the system **which teacher gets the form for which class**.

### In the "Settings" folder — the holiday list (recommended)

Create a file named **`publicHoliday`** listing days when **no forms** should go out
(weekends are skipped automatically). Use a **Date** column. You can enter a single day
or a range like `10/10/2026 - 15/10/2026`.

*(Optional advanced files — `programManagers`, `teacherLeads`, `default sharing scopes`
— control who can view/edit the spreadsheets. You can skip these to start; the system
still works without them.)*

---

## Step 3 — Fill in the settings file (`SetConfig.gs`)

This is the **only file you need to edit**. Open it in the Apps Script editor and change
these lines:

- **`STUDENT_CLASS_LIST_FOLDER_LINK`** → paste the link to your **Rosters** folder.
- **`ATTENDANCE_SHEETS_FOLDER_LINK`** → paste the link to your **Attendance** folder.
- **`CONFIG_FOLDER_LINK`** → paste the link to your **Settings** folder.
- **`ACADEMIC_YEAR`** → the school year, e.g. `"2026-2027"`.
- **`START_MONTH`** / **`END_MONTH`** → the first and last month of the year as numbers
  (June = `6`, April = `4`).
- **`STAKEHOLDER_EMAILS`** → who should get the weekly report, separated by commas.

Leave everything else as it is. Save the file.

---

## Step 4 — Turn on Google's file service (one-time)

In the Apps Script editor, on the left, click the **＋** next to **Services**, find
**Drive API**, and add it. (This lets the system read Excel files.)

The **first time** you run anything, Google will ask you to **allow permissions** — click
through and accept. This is needed so it can send emails and manage your files.

---

## Step 5 — Run the setup, in this exact order

In the Apps Script editor, pick each function from the dropdown and click **Run**. Wait
for it to finish and check the log message before moving to the next.

1. **`manual_validateConfig`** — checks it can reach your 3 folders. Should say success.
2. **`manual_validateSetupConfig`** — checks your settings/teacher/holiday files load. Should say success.
3. **`manual_generateSheets`** — creates one attendance spreadsheet per class.
   - If it says it "ran out of time," just **run it again** — it picks up where it left off.
4. **`manual_installTriggers`** — turns on the daily automation. **Run this only once.**

That's it. The system is now live.

---

## Step 6 — (Optional) Test it

- Run **`automated_sendDailyForms`** once and check that a teacher gets a form email.
- Fill in that form, then run **`automated_syncResponses`** and check the answers appear
  in that class's spreadsheet.

---

## What happens automatically after setup

You don't touch anything day to day. Behind the scenes:

- **6:00 AM every day** — a fresh attendance form is emailed to each teacher.
- **Every hour** — new responses are saved into the spreadsheets, the dashboards
  refresh, and leftover forms / stray response tabs are tidied up automatically.
- **11:00 PM every day** — a final sync runs, the dashboards refresh once more, and
  the day's forms are closed.
- **Every Friday, 5:00 PM** — a summary report is emailed to the stakeholders.

Weekends and the holidays you listed are skipped automatically.

---

## Good to know

- **File names the system creates:** each class spreadsheet is named like
  `Class_1_A_2026-2027`, and each daily form is named like
  `Attendance: Class 1-A (25-Jul-2026)`. You don't create these — they appear on their own.
- **New students mid-year:** update that class's roster file, then run
  **`manual_updateSheets`**. It adds the new students without erasing existing records.
- **Don't run the setup twice:** running `manual_installTriggers` more than once creates
  duplicate automations. If unsure, open the **Triggers** page (clock icon) — you should
  see exactly **4** items. Delete any extras.

---

## Quick checklist

- [ ] 3 folders created and their links pasted into `SetConfig.gs`.
- [ ] One correctly-named roster per class in the Rosters folder.
- [ ] `TeacherClassMapping` (and `publicHoliday`) in the Settings folder.
- [ ] Drive API added; permissions accepted on first run.
- [ ] Ran, in order: validate config → validate setup → generate sheets → install triggers.
- [ ] Triggers page shows exactly 4 items.
- [ ] (Optional) Sent a test form and saw a test response appear in the spreadsheet.
