# System Reset — Fresh Start Guide

This document describes how to reset the **Ek Tara Attendance Management System**
back to a clean state so the next day (or next academic year) starts fresh.

There are **two kinds of reset**. Read the top of each section to pick the right one.

- **A. Daily reset** — normal end-of-day cleanup. *This is automatic; you rarely
  need to do anything.* Read it to understand what happens each night.
- **B. Full reset** — wipe everything (workbooks, forms, runtime state, triggers)
  to start the system over from scratch, e.g. a new academic year or a botched setup.

> ⚠️ **Nothing here deletes your source data.** The class rosters in the *input
> folder* and the config files (`TeacherClassMapping`, `publicHoliday`) are never
> touched by a reset. A full reset only removes **generated** artifacts.

---

## Key concepts you need before resetting

The system stores state in **three places**. A true reset must address all three:

| Where | What lives there | Example |
|-------|------------------|---------|
| **Google Drive** | Generated attendance workbooks + generated Forms | `Class_1_A_2026-2027`, `Attendance: Class 1-A (25-Jul-2026)` |
| **Script Properties** | Runtime pointers & dedup state | `ACTIVE_FORM_<ssId>`, `AUTHORIZED_TEACHER_<ssId>`, `NOTIFIED_<ssId>`, `FLOW1_TOKEN`, `FLOW2_TOKEN`, `FLOW3_TOKEN` |
| **Triggers** | The 4 time-based automations | `automated_sendDailyForms`, `automated_syncResponses`, `automated_closeForms`, `automated_sendWeeklyReport` |

### Exact file-name formats (so you know what to delete)

| Artifact | Format | Example |
|----------|--------|---------|
| Attendance workbook | `Class_<classNum>_<SECTION>_<ACADEMIC_YEAR>` | `Class_1_A_2026-2027` |
| Daily form | `Attendance: Class <classNum>-<SECTION> (<dd-MMM-yyyy>)` | `Attendance: Class 1-A (25-Jul-2026)` |
| Makeup form | `Attendance (Makeup): Class <classNum>-<SECTION> (<dd-MMM-yyyy>)` | `Attendance (Makeup): Class 5-A (25-Jul-2026)` |

- `<SECTION>` is always **UPPERCASE** in the file name (`A`, `B`, `C`).
- `<ACADEMIC_YEAR>` comes from `ACADEMIC_YEAR` in `SetConfig.gs` (currently `2026-2027`).
- Generated workbooks and forms all live in the **attendance sheets output folder**
  (`ATTENDANCE_SHEETS_FOLDER_LINK` in `SetConfig.gs`).

### Script Property keys, explained

| Key pattern | Set by | Cleared by | Purpose |
|-------------|--------|-----------|---------|
| `ACTIVE_FORM_<ssId>` | `automated_sendDailyForms`, `manual_runAdhocForm` | `automated_closeForms` (nightly) | Which form is currently live for a workbook |
| `AUTHORIZED_TEACHER_<ssId>` | `automated_sendDailyForms`, `manual_runAdhocForm` | ⚠️ **never** (leaks) → cleared only by full reset | Which teacher email may submit for that workbook |
| `NOTIFIED_<ssId>` | `notifySubmitters` (during sync) | `automated_closeForms` (nightly) | Per-day dedup so each teacher is emailed once/day |
| `FLOW1_TOKEN` | `manual_generateSheets` | `manual_resetExecutionTokens`, full reset | Batch continuation for initial sheet generation |
| `FLOW2_TOKEN` | `manual_updateSheets` | `manual_resetExecutionTokens`, full reset | Batch continuation for mid-year sheet updates |
| `FLOW3_TOKEN` | (reserved) | `manual_resetExecutionTokens`, full reset | Reserved batch continuation token |

---

## A. Daily reset (automatic — for reference)

At **11 PM** every day, the `automated_closeForms` trigger runs and, for each
workbook that has an active form:

1. Stops the form from accepting responses.
2. Deletes `ACTIVE_FORM_<ssId>`.
3. Deletes `NOTIFIED_<ssId>` — so the next day, each teacher again gets exactly
   **one** confirmation email.

The next morning at **6 AM**, `automated_sendDailyForms` runs and creates a brand
new form for the day. **The workbook is reused** (not recreated) — new responses
are written into the correct day-column of the current month's sheet.

✅ **You do not need to do anything for a normal daily fresh start.** As long as the
4 triggers are installed, each day cleans up after itself.

> If forms were **not** closing (e.g. the 11 PM trigger was missing), run
> `manual_closeFormsFlexibly()` once to close any lingering forms, then confirm
> triggers are installed (Step B6 below).

---

## B. Full reset (start completely over)

Use this for a new academic year, or to recover from a broken/half-finished setup.
Do the steps **in order**.

### B1. Turn off automation first

Delete the 4 time-based triggers so nothing runs mid-reset.

- Apps Script editor → **Triggers** (clock icon, left rail) → for each of the 4
  triggers below, click **⋮ → Delete trigger**:
  - `automated_sendDailyForms`
  - `automated_syncResponses`
  - `automated_closeForms`
  - `automated_sendWeeklyReport`

### B2. Close any live forms (optional but tidy)

Run `manual_closeFormsFlexibly()` once (leave its filters as `null`) to stop all
forms from accepting responses before you delete them.

### B3. Delete generated Drive artifacts

In the **attendance sheets output folder** (`ATTENDANCE_SHEETS_FOLDER_LINK`), delete:

- All **workbooks** named `Class_<n>_<SEC>_<ACADEMIC_YEAR>` (e.g. `Class_1_A_2026-2027`).
  - Deleting the workbook also removes its embedded **form-response tab** and the
    generated **`Analysis_Dashboard`** tab — no need to clean those separately.
- All **Forms** titled `Attendance: Class …` and `Attendance (Makeup): Class …`.

> 🔒 **Do NOT delete** the class rosters in the input folder, or `TeacherClassMapping`
> / `publicHoliday` in the config folder. Those are your source data.
>
> 🗑️ Empty **Trash** afterward, otherwise a workbook of the same name can conflict
> when it's regenerated.

### B4. Clear all runtime Script Properties

Run **`manual_resetAllRuntimeState()`** once from the Apps Script editor.

This clears every runtime key in one shot:
`ACTIVE_FORM_*`, `AUTHORIZED_TEACHER_*` (the leaking key), `NOTIFIED_*`,
and `FLOW1_TOKEN` / `FLOW2_TOKEN` / `FLOW3_TOKEN`. Check the execution log for the
`✅ Full runtime reset complete.` line.

> Prefer to do it by hand? Apps Script editor → **Project Settings** (gear) →
> **Script Properties** → delete every key matching the patterns above.
>
> If you *only* want to reset the setup batch tokens (not the daily runtime keys),
> run `manual_resetExecutionTokens()` instead.

### B5. Confirm config is in place, then regenerate

1. Confirm `ACADEMIC_YEAR`, `START_MONTH`, `END_MONTH` in `SetConfig.gs` are correct
   for the fresh start (e.g. new academic year).
2. Run **`manual_validateConfig()`** and **`manual_validateSetupConfig()`** — both
   must log success (all folders accessible, config valid).
3. Run **`manual_generateSheets()`** to build fresh `Class_<n>_<SEC>_<year>` workbooks
   from the current rosters. (If it stops on the 6-min limit, just run it again —
   `FLOW1_TOKEN` resumes where it left off.)

### B6. Re-install triggers

Run **`manual_installTriggers()`** once. This registers the 4 automations again
(6 AM forms, hourly sync, 11 PM close, Friday 5 PM report).

> ⚠️ **20-trigger limit:** Google caps a project at 20 triggers per user. Before
> reinstalling, make sure the old ones were deleted in **B1** — otherwise you can
> hit the cap. After install, the Triggers page should show exactly **4**.

### B7. Verify the fresh start

- **Triggers page** shows exactly the 4 `automated_*` triggers.
- **Output folder** shows one `Class_<n>_<SEC>_<year>` workbook per class, each with
  the month tabs but **no day-column data yet**.
- **Script Properties** shows no `ACTIVE_FORM_*` / `AUTHORIZED_TEACHER_*` /
  `NOTIFIED_*` keys (they'll be created fresh at the next 6 AM run).

At the next 6 AM trigger, daily forms go out and the system runs clean.

---

## Validated run order (quick reference)

**Full reset, start to finish:**

1. Delete the 4 triggers (Triggers page).
2. `manual_closeFormsFlexibly()` — close live forms.
3. Delete generated workbooks + forms in the output folder; empty Trash.
4. `manual_resetAllRuntimeState()` — clear all Script Properties.
5. `manual_validateConfig()` → `manual_validateSetupConfig()` — confirm config.
6. `manual_generateSheets()` — rebuild workbooks (re-run if it times out).
7. `manual_installTriggers()` — reinstall the 4 automations.
8. Verify (B7).

**Just need tomorrow to be clean (no full wipe):** do nothing — the 11 PM
`automated_closeForms` + 6 AM `automated_sendDailyForms` cycle handles it. If a
form failed to close, run `manual_closeFormsFlexibly()` once.
