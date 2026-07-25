# Initial Setup Guide

How to stand up the **Ek Tara Attendance Management System** from scratch: what the
code files are, **where to make configuration changes**, how to prepare the Google
Drive folders and files (with exact naming conventions), and the exact order of
functions to run — including installing the automation triggers.

> Companion doc: once set up, see **`RESET.md`** for how to reset for a fresh start.

---

## 0. What this system does (30-second overview)

- A pool of **class rosters** (one file per class-section) lives in an **input folder**.
- Setup builds one **attendance workbook per class-section** in an **output folder**,
  with a month tab per academic month and a day-column per day.
- Every morning a trigger emails each teacher a **Google Form** to mark attendance;
  an hourly trigger **syncs** responses into the workbook and refreshes an
  **`Analysis_Dashboard`** tab; a nightly trigger closes the forms; a weekly trigger
  emails a stakeholder digest.

---

## 1. The code — files and what each contains

All code lives in the `attendance management system/` folder as Apps Script files.
**Apps Script runs on Google's servers** — you cannot run it locally; you paste/push
these files into an Apps Script project bound to your Google account and run the
functions there.

| File | Role | You edit it during setup? |
|------|------|---------------------------|
| **`SetConfig.gs`** | All global configuration (folder links, file names, academic year, alert thresholds, stakeholder emails). | ✅ **YES — this is the main file you change.** |
| **`SetupOperations.gs`** | One-time / manual admin functions: generate & update workbooks, reset helpers, setup validator. | Rarely (run its functions). |
| **`DailyOperations.gs`** | The 4 automated (trigger) operations + manual form/close/validate helpers + the **trigger installer**. | No (run its functions). |
| **`Utils.gs`** | Shared helpers: file-name parsing, config loaders, workbook builder, sync engine, email/HTML, dashboard. | No. |

### Function naming convention (important)

Every function is prefixed to make its purpose unambiguous:

- **`automated_*`** — run by a **time-based trigger** (one function per trigger). Do not
  run these by hand except to test.
- **`manual_*`** — run **by an admin on demand** from the Apps Script editor.

The 4 automated operations:

| Function | Trigger | Schedule |
|----------|---------|----------|
| `automated_sendDailyForms` | time-based | Daily at **6 AM** |
| `automated_syncResponses` | time-based | **Every hour** |
| `automated_closeForms` | time-based | Daily at **11 PM** |
| `automated_sendWeeklyReport` | time-based | **Fridays at 5 PM** |

---

## 2. WHERE to make code changes — `SetConfig.gs`

**Almost all setup changes happen in `SetConfig.gs`.** Edit these values before your
first run:

| Variable | What to set it to | Example |
|----------|-------------------|---------|
| `INPUT_FOLDER_LINK` | Drive URL of the folder holding class rosters | `https://drive.google.com/drive/folders/…` |
| `ATTENDANCE_SHEETS_FOLDER_LINK` | Drive URL of the (initially empty) output folder | `…/folders/…` |
| `CONFIG_FOLDER_LINK` | Drive URL of the folder holding config files | `…/folders/…` |
| `MAPPING_FILE_NAME` | Base name of the teacher↔class mapping file | `TeacherClassMapping` |
| `HOLIDAY_FILE_NAME` | Base name of the public-holiday file | `publicHoliday` |
| `ACADEMIC_YEAR` | `"<startYear>-<endYear>"` — **drives the workbook name** | `"2026-2027"` |
| `START_MONTH` / `END_MONTH` | First/last academic month (1 = Jan … 12 = Dec) | `6` (June) / `4` (April) |
| `STAKEHOLDER_EMAILS` | Comma-separated recipients of the weekly report | `"a@x.com,b@x.com"` |
| Alert thresholds | `CONSECUTIVE_ABSENT_THRESHOLD_DAYS`, `LATE_*`, `ABSENT_*`, etc. | keep defaults unless needed |
| `ADHOC_*` | Defaults used by `manual_runAdhocForm` (makeup forms) | class, section, teacher, date |

> `ADD_VISUALISATIONS` and `MAX_EXECUTION_TIME` normally stay as-is. Do **not** leave
> `ADD_VISUALISATIONS` undefined — the validators log it and would throw otherwise.

Everything else (thresholds, adhoc defaults) is optional tuning. **No file paths or
names are hard-coded elsewhere** — the workbook name is centralized in
`getWorkbookName()` in `Utils.gs`, so you never edit names in multiple places.

---

## 3. Prepare Google Drive — folders and files (naming conventions)

Create **three folders** and paste their URLs into `SetConfig.gs` (Section 2).

### 3a. Input folder — class rosters (one file per class-section)

- **File type:** `.csv`, `.xlsx`/Excel, or a Google Sheet — all supported.
- **File name must contain the class number and section** so the parser can read
  them. The parser looks for `class <X>` / `grade <X>` plus a standalone section
  letter. Underscores are treated as spaces.

  | Good file names | Parsed as |
  |-----------------|-----------|
  | `Class 1 A.xlsx` | class `1`, section `A` |
  | `Class_5_B` | class `5`, section `B` |
  | `Grade 7 C.csv` | class `7`, section `C` |

- **Roster contents:** a header row plus one row per student. Recognized headers
  (case-insensitive, any order): **Roll No.**, **Child ID**, **Name / Student**.
  - If no header is detected, columns are assumed to be `[Child ID, Name, …]` and
    Roll No. is auto-numbered.
  - Only these three fields are used; extra columns are ignored.

### 3b. Config folder — mapping, holidays, roles

Put these files in the **config folder**. Names are matched loosely (case-insensitive,
substring), but use these exact base names to be safe:

| File (base name) | Required? | Purpose | Expected columns (header row) |
|------------------|-----------|---------|-------------------------------|
| **`TeacherClassMapping`** | ✅ **Required** | Maps each class-section → teacher name + email (who gets the daily form). | `Name`, `Email`, `Class` (or `Grade`), `Section` |
| **`publicHoliday`** | Recommended | Days to skip sending forms. | `Date` (single or `start - end` range), or `Start Date` + `End Date` / `From` + `To` |
| **`programManagers`** | Optional | PM contacts for permissions/escalation. | `ID`, `Email` |
| **`teacherLeads`** | Optional | Lead contacts, linked to a manager. | `ID`, `Email`, `Manager ID` |
| **`default sharing scopes`** | Optional | A Sheet with `Roles` and `Stakeholders` tabs controlling who gets edit/view access on generated workbooks. | Stakeholders tab: `Email` (or `Email ID`), `Scope` (`edit`/`view`) |

- Dates accept `DD/MM/YYYY`, `DD-MM-YYYY`, or `YYYY-MM-DD`; ranges via `"start - end"`
  or `"start to end"`.
- If `default sharing scopes` / managers / leads are absent, setup still works — it
  just falls back to fewer permissions.

### 3c. Output folder — generated workbooks (created for you)

Start **empty**. Setup fills it with:

| Artifact | Naming convention | Example |
|----------|-------------------|---------|
| Attendance workbook | `Class_<classNum>_<SECTION>_<ACADEMIC_YEAR>` | `Class_1_A_2026-2027` |
| Daily form (later, at runtime) | `Attendance: Class <classNum>-<SECTION> (<dd-MMM-yyyy>)` | `Attendance: Class 1-A (25-Jul-2026)` |
| Makeup form | `Attendance (Makeup): Class <classNum>-<SECTION> (<dd-MMM-yyyy>)` | `Attendance (Makeup): Class 5-A (25-Jul-2026)` |

- `<SECTION>` is always **UPPERCASE**; `<ACADEMIC_YEAR>` is the value from
  `SetConfig.gs`.
- Each workbook has a **tab per academic month** (from `START_MONTH` to `END_MONTH`),
  columns `A/B/C = Roll No. / Child ID / Name`, then a **day-column per day**, then
  Present/Absent/Late/Total/Percentage summary rows, plus an `Analysis_Dashboard` tab.

---

## 4. Enable the required Google services

In the Apps Script editor:

1. **Advanced Drive Service** — **Services (＋)** → add **Drive API**. Required by
   `parseAndNormalizeData` to convert Excel rosters (`Drive.Files.copy`). *(CSV and
   Google Sheet rosters work without it, but enable it to be safe.)*
2. On the **first run** of any function, Google prompts for **authorization** — grant
   the scopes (Drive, Sheets, Forms, Gmail/MailApp, Script triggers). This is required
   for the automations to send mail and manage files.

---

## 5. Run order — the exact sequence

Run each from the Apps Script editor (**Select function → Run**). Watch the
**Execution log** after each.

| Step | Function | File | What it does |
|------|----------|------|--------------|
| 1 | `manual_validateConfig` | DailyOperations | Confirms all three folders are reachable. Must pass. |
| 2 | `manual_validateSetupConfig` | SetupOperations | Confirms setup config/roles/holidays load. Must pass. |
| 3 | `manual_generateSheets` | SetupOperations | Builds one `Class_<n>_<SEC>_<year>` workbook per roster. **Re-run if it stops on the 6-min limit** — `FLOW1_TOKEN` resumes automatically. |
| 4 | `manual_installTriggers` | DailyOperations | Installs the 4 `automated_*` triggers. **Run once.** |
| 5 | *(optional test)* `automated_sendDailyForms` | DailyOperations | Manually fire once to confirm a teacher receives a form. |
| 6 | *(optional test)* `automated_syncResponses` | DailyOperations | Manually fire after a test submission to confirm it writes to the workbook + dashboard. |

**Mid-year roster changes** (new students / new classes added later): drop the updated
roster into the input folder and run **`manual_updateSheets`** (uses `FLOW2_TOKEN`;
re-run if it times out). It adds new students without disturbing existing data.

---

## 6. Triggers — how they get installed and the limit

`manual_installTriggers()` (in `DailyOperations.gs`) creates exactly these 4
time-based triggers — **one function per trigger**:

```
automated_sendDailyForms    → daily @ 6 AM
automated_syncResponses     → every 1 hour
automated_closeForms        → daily @ 11 PM
automated_sendWeeklyReport  → Fridays @ 5 PM
```

> ⚠️ **20-trigger limit:** Google caps a project at **20 triggers per user**. This
> design uses only **4** (independent of the number of classes) precisely to stay well
> under the cap — attendance auth is handled inside the hourly sync, **not** via a
> per-form `onFormSubmit` trigger. If you ever re-run `manual_installTriggers` without
> deleting the old ones, you'll create duplicates — always delete first (Triggers page
> → ⋮ → Delete) so the page shows exactly 4.

After setup the system is autonomous: forms go out at 6 AM, sync hourly, close at
11 PM, weekly report Fridays — with all history accumulating in the workbooks.

---

## 7. Post-setup checklist

- [ ] `SetConfig.gs`: three folder links + `ACADEMIC_YEAR` + `STAKEHOLDER_EMAILS` set.
- [ ] Input folder has one correctly-named roster per class-section.
- [ ] Config folder has `TeacherClassMapping` (required) and `publicHoliday`.
- [ ] Advanced Drive service enabled; authorization granted.
- [ ] `manual_validateConfig` and `manual_validateSetupConfig` both pass.
- [ ] `manual_generateSheets` produced one `Class_<n>_<SEC>_<year>` workbook per class.
- [ ] `manual_installTriggers` run once; Triggers page shows exactly **4**.
- [ ] Test form received and a test response synced into the workbook.
