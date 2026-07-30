# System Reset — Fresh Start Guide

This document describes how to reset the **Ek Tara Attendance Management System**
so it **starts fresh from the following day** — **without losing any stored
attendance data or its visualizations**.

There are **three** levels of reset. Almost always you want **A** or **B**.

- **A. Daily reset** — normal end-of-day cleanup. *Automatic; you rarely do anything.*
- **B. Data-preserving reset** ← **the one you usually want.** Clears runtime state so
  tomorrow starts clean, but **keeps every workbook, every day-column of history, and
  the `Analysis_Dashboard` visualizations.**
- **C. Full wipe (new academic year)** — ⚠️ **DESTROYS data.** Deletes the workbooks
  themselves. Only for a brand-new year or recovering a broken setup.

> ✅ **A and B never delete attendance data.** Historical marks live in the day-columns
> of each workbook's month tabs, and the dashboards live in the `Analysis_Dashboard`
> tab. As long as the **workbook** is not deleted, all of that is preserved and simply
> continues the next day. Only **C** removes workbooks.
>
> ⚠️ Source data (class rosters in the input folder, `TeacherClassMapping`,
> `publicHoliday`) is never touched by any reset.

---

## Where state lives

A reset must consider three places, but **only the workbooks hold data/visualizations**:

| Where | What lives there | Deleting it loses data? |
|-------|------------------|-------------------------|
| **Attendance workbooks** (Drive) | All attendance history (day-columns) + `Analysis_Dashboard` visualizations | ✅ **YES — keep these for A & B** |
| **Generated Form files** (Drive) | The daily/makeup form UI (intermediate; already-synced marks are in the workbook) | ❌ No — safe to delete once synced |
| **Script Properties** | Runtime pointers & dedup flags (no attendance data) | ❌ No |
| **Triggers** | The 4 time-based automations | ❌ No |

### Exact file-name formats

| Artifact | Format | Example |
|----------|--------|---------|
| Attendance workbook (**KEEP for A/B**) | `Class_<classNum>_<SECTION>_<ACADEMIC_YEAR>` | `Class_1_A_2026-2027` |
| Daily form | `Attendance: Class <classNum>-<SECTION> (<dd-MMM-yyyy>)` | `Attendance: Class 1-A (25-Jul-2026)` |
| Makeup form | `Attendance (Makeup): Class <classNum>-<SECTION> (<dd-MMM-yyyy>)` | `Attendance (Makeup): Class 5-A (25-Jul-2026)` |

- `<SECTION>` is always **UPPERCASE** (`A`, `B`, `C`).
- `<ACADEMIC_YEAR>` comes from `ACADEMIC_YEAR` in `SetConfig.gs` (currently `2026-2027`).
- Workbooks and forms all live in the **output folder** (`ATTENDANCE_SHEETS_FOLDER_LINK`).

### Script Property keys (none of these hold attendance data)

| Key pattern | Set by | Cleared by | Purpose |
|-------------|--------|-----------|---------|
| `ACTIVE_FORM_<ssId>` | `automated_sendDailyForms`, `manual_sendOnDemandForm` | `automated_closeForms` (nightly) | Which form is live for a workbook |
| `AUTHORIZED_TEACHER_<ssId>` | `automated_sendDailyForms`, `manual_sendOnDemandForm` | ⚠️ **never** (leaks) → cleared only by reset B/C | Which teacher email may submit |
| `NOTIFIED_<ssId>` | `notifySubmitters` (during sync) | `automated_closeForms` (nightly) | Per-day dedup so each teacher emailed once/day |
| `FLOW1_TOKEN` / `FLOW2_TOKEN` / `FLOW3_TOKEN` | setup batch functions | `manual_resetExecutionTokens`, reset B/C | Batch continuation tokens (setup only) |

---

## A. Daily reset (automatic — for reference)

At **11 PM** the `automated_closeForms` trigger, for each workbook with a live form:

1. Stops the form accepting responses.
2. Deletes `ACTIVE_FORM_<ssId>` and `NOTIFIED_<ssId>`.

At **6 AM** `automated_sendDailyForms` creates a **new** form for the day and
**reuses the existing workbook** — new marks append into today's day-column. **All
prior history and the dashboard are untouched.**

✅ **For a normal fresh start next day you do nothing** — as long as the 4 triggers
are installed, each day cleans up after itself and data carries forward.

---

## B. Data-preserving reset (recommended fresh start) ✅ KEEPS ALL DATA

Use this when you want tomorrow to start clean — e.g. the runtime state got into a
bad shape, forms didn't close, or the `AUTHORIZED_TEACHER_` keys leaked — **but you
want to keep all attendance history and the visualizations.**

**This does NOT delete any workbook, day-column, or `Analysis_Dashboard`.**

### B1. (Recommended) Sync today's responses first

Run `automated_syncResponses()` once so any pending marks from today are written into
the workbooks **before** you clear the runtime pointers. (After the reset, the hourly
sync skips workbooks with no `ACTIVE_FORM_`, so capture today's data now.)

### B2. Close any live forms

Run `manual_closeFormsFlexibly()` (leave filters `null`) to stop all forms accepting
responses.

### B3. Clear runtime Script Properties

Run **`manual_resetAllRuntimeState()`**.

This clears `ACTIVE_FORM_*`, `AUTHORIZED_TEACHER_*` (the leaking key), `NOTIFIED_*`,
and the `FLOW*_TOKEN`s. It touches **no Drive files**, so **every workbook and its
dashboard remain intact.** Confirm the `✅ Full runtime reset complete.` log line.

### B4. (Optional) Tidy up old form files

You may delete the closed **Form files** (`Attendance: Class …` / `Attendance
(Makeup): Class …`) from the output folder to reduce clutter. This is **safe** — the
marks are already in the workbooks. **Do not delete the `Class_… ` workbooks.**

### B5. Confirm triggers are installed

Triggers page should show the 4 `automated_*` triggers. If any are missing, run
`manual_installTriggers()` (delete duplicates first — see the 20-trigger note in C).

✅ **Result:** at the next 6 AM run, fresh forms go out and new marks append to the
**same** workbooks. All history and visualizations are preserved.

---

## C. Full wipe — new academic year ⚠️ DESTROYS ATTENDANCE DATA

Only for starting a **brand-new academic year** or rebuilding a broken setup. This
**deletes the workbooks**, so all attendance history and dashboards are lost.

> 💾 **Export/download or copy the `Class_…` workbooks first** if you want to keep the
> old year's records.

1. Delete the 4 triggers (Triggers page): `automated_sendDailyForms`,
   `automated_syncResponses`, `automated_closeForms`, `automated_sendWeeklyReport`.
2. `manual_closeFormsFlexibly()` — close live forms.
3. In the output folder, delete **all** `Class_<n>_<SEC>_<year>` workbooks (this
   removes their history + `Analysis_Dashboard`) and all `Attendance: …` forms.
   Empty **Trash** so regenerated names don't conflict. *(Do NOT delete rosters or
   config files.)*
4. `manual_resetAllRuntimeState()` — clear all Script Properties.
5. Update `ACADEMIC_YEAR` (and `START_MONTH` / `END_MONTH` if needed) in `SetConfig.gs`.
6. `manual_validateConfig()` → `manual_validateSetupConfig()` — both must pass.
7. `manual_generateSheets()` — build fresh workbooks (re-run if it hits the 6-min
   limit; `FLOW1_TOKEN` resumes).
8. `manual_installTriggers()` — reinstall the 4 automations.

> ⚠️ **20-trigger limit:** Google caps a project at 20 triggers per user. Delete the
> old triggers (step 1) before reinstalling, then confirm exactly **4** appear.

---

## Quick reference

| Goal | Do this |
|------|---------|
| **Normal fresh start tomorrow** (keep data) | Nothing — the 11 PM + 6 AM cycle handles it. |
| **Clean runtime state** but keep all data/visualizations | **B:** `automated_syncResponses()` → `manual_closeFormsFlexibly()` → `manual_resetAllRuntimeState()` |
| **Start a brand-new academic year** (data will be lost) | **C:** delete triggers → delete workbooks/forms → `manual_resetAllRuntimeState()` → validate → `manual_generateSheets()` → `manual_installTriggers()` |
