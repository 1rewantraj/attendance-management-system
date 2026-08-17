# Attendance Management System — Operations Guide & Demo Scripts

This document gives the **exact steps** for the five common admin scenarios, plus a
**video-recording script** for each so demos are consistent.

## Where everything lives
- **Google Drive** holds three folders (links configured in `SetConfig.gs`):
  - **Student Class List folder** (`STUDENT_CLASS_LIST_FOLDER_LINK`) — one roster file per class-section.
  - **Attendance Sheets folder** (`ATTENDANCE_SHEETS_FOLDER_LINK`) — the generated workbooks (output).
  - **Config folder** (`CONFIG_FOLDER_LINK`) — teacher mapping, holidays, sharing scopes, leads, managers.
- **The script**: open the main Google Sheet → **Extensions → Apps Script**.
- **How to run a function**: In the Apps Script editor, pick the function from the
  **function dropdown** (top toolbar) → click **Run** → watch **Execution log** (View → Logs).
- **First run of any function** triggers a Google **authorization** prompt — approve it once.

### Roster file format (important)
Each roster file's columns are detected by header name (order-independent):

| Roll No. | Child ID | Name of the Student | Status |
|----------|----------|---------------------|--------|
| 1        | S1001    | Aarav Sharma        | Active |
| 2        | S1002    | Diya Patel          | Active |

- **Status** accepts `Active` or `Inactive` (blank = Active). Any other value → that
  student row is skipped with a warning.
- The file **name** encodes class + section, e.g. `Class_5_A.csv` or `Class 5 Section A.xlsx`.

---

## Scenario 1 — Generate attendance sheets for all classes (+ install automation)

**Goal:** First-time setup — create every class workbook, then install the triggers that
send forms, sync responses, and email the weekly report automatically.

**Steps:**
1. Confirm all roster files are in the **Student Class List folder**, one per class-section,
   each with the `Roll No. / Child ID / Name / Status` columns.
2. Confirm the **Config folder** contains: `TeacherClassMapping`, `publicHoliday`, plus
   `teacherLeads`, `programManagers`, and the sharing-scopes file.
3. Open **Extensions → Apps Script**.
4. Run **`manual_validateSetupConfig`** — confirms folders are reachable, counts roster files,
   and prints the academic months. Fix any ❌ errors it reports before continuing.
5. Run **`manual_generateSheets`**.
   - It creates one workbook per class-section (named `Class_<n>_<Section>_<AcademicYear>`),
     builds a tab per academic month, and applies teacher/lead/manager permissions.
   - **If it logs `⏳ Execution time limit approaching. Progress saved`**, just run it **again** —
     it resumes where it left off (it skips workbooks that already exist).
   - Repeat until you see `SUCCESS: MASTER BATCH GENERATION COMPLETE`.
6. Run **`manual_installTriggers`** once. This installs 4 time-based triggers:
   - `automated_sendDailyForms` — daily at ~6 AM
   - `automated_syncResponses` — hourly
   - `automated_closeForms` — daily at ~11 PM (final sync + dashboard refresh)
   - `automated_sendWeeklyReport` — Fridays at ~5 PM
7. Verify in the Apps Script left sidebar → **Triggers** (clock icon): 4 triggers listed.

> ⚠️ Run `manual_installTriggers` **only once**. Running it again creates duplicate triggers.

**🎬 Recording script (Scenario 1):**
- "Here are our roster files in the Student Class List folder — one per class." (show folder)
- "And here's the Config folder with the teacher mapping and holidays." (show folder)
- "Open Extensions → Apps Script." (show editor)
- "First I validate the setup." → run `manual_validateSetupConfig`, read the ✅ log lines.
- "Now I generate the sheets." → run `manual_generateSheets`; show a workbook opening in the output folder with monthly tabs.
- "If it times out on a big batch, I just re-run it — it resumes automatically." (mention, no need to force)
- "Finally I install the automation." → run `manual_installTriggers`; open the Triggers panel to show the 4 triggers.
- Close: "That's the full setup — sheets created and automation running."

---

## Scenario 2 — Send an on-demand form for an older date (+ close it)

**Goal:** A teacher needs to submit/fix attendance for a **past date**. Send a one-off form
for that specific date, let them fill it, sync it to the right column, then close the form.

**Steps — send the form:**
1. Open **Extensions → Apps Script** → open `SetConfig.gs`.
2. Set the on-demand defaults:
   ```javascript
   var ONDEMAND_CLASS_NUM   = "5";            // class
   var ONDEMAND_SECTION     = "A";            // section
   var ONDEMAND_DATE        = "15-07-2026";   // DD-MM-YYYY  (empty = today)
   var ONDEMAND_TEACHER_NAME  = "Teacher Name";   // optional override
   var ONDEMAND_TEACHER_EMAIL = "teacher@example.com"; // optional override
   ```
   - **Date format is `DD-MM-YYYY`** (e.g. `15-07-2026` = 15 July 2026). Wrong formats
     throw a clear error.
   - Leave name/email as the placeholders to auto-use the mapped class teacher.
3. **Save** (Ctrl/Cmd-S), then run **`manual_sendOnDemandForm`**.
   - A makeup form is created and **emailed to the teacher**; the log shows `✅ ... form sent to: ...`.
   - The system stores the form's **target date**, so responses land in the correct day column.

**Steps — after the teacher submits:**
4. Run **`manual_syncResponses`** to pull the responses into the sheet **immediately**
   (otherwise the hourly `automated_syncResponses` will do it within the hour).
   - Open the workbook → the target month tab → the target day column shows the marks.

**Steps — close the on-demand form:**
5. Run **`manual_closeOnDemandForm`**.
   - With no filters set it closes **all** currently active forms; it stops accepting
     responses and clears the form's runtime state.
   - **To close just one class's form**, edit the filter variables at the top of
     `manual_closeOnDemandForm` before running:
     ```javascript
     var FILTER_CLASS = "5";     // only class 5
     var FILTER_SECTION = "A";   // only section A
     var FILTER_DATE_STRING = null;
     ```
   - The nightly `automated_closeForms` will also close it automatically at ~11 PM if you
     don't close it manually.

> Note: `manual_sendOnDemandForm` reuses the `ONDEMAND_*` values in `SetConfig.gs`. Remember
> to reset `ONDEMAND_DATE` back to `""` when you're done so you don't accidentally reuse it.

**🎬 Recording script (Scenario 2):**
- "A teacher forgot to mark attendance for July 15. Let's send a makeup form."
- Open `SetConfig.gs`, set `ONDEMAND_CLASS_NUM`, `ONDEMAND_SECTION`, and `ONDEMAND_DATE = "15-07-2026"`. Save.
- Run `manual_sendOnDemandForm`; show the ✅ log and the email arriving in the teacher's inbox.
- "The teacher fills the form." (show filling + submitting the form)
- Run `manual_syncResponses`; open the workbook → July tab → column 15 to show the synced marks.
- "Now we close the form." → run `manual_closeOnDemandForm`; show the `🔒 Closed` log.
- Close: "Attendance for a past date, captured and synced to the right day."

---

## Scenario 3 — Class teacher change or new student added

**Goal:** A section got a **new teacher** (or lead/manager), and/or **new students** joined.
`manual_updateSheets` handles **both** — it re-checks permissions and adds new active students.

> 💡 **Speed tip for scenarios 3–5:** `manual_updateSheets` scans every roster by default.
> To update just one class-section, set both filters in `SetConfig.gs` before running, then
> reset them to `""` afterwards:
> ```javascript
> var UPDATE_SHEETS_CLASS_FILTER   = "5";   // only class 5
> var UPDATE_SHEETS_SECTION_FILTER = "A";   // only section A
> ```

**Steps:**
1. **For a teacher/lead/manager change:** open the **Config folder** and update the relevant
   file so the new person's email is present for that class-section:
   - Teacher → `TeacherClassMapping`
   - Lead → `teacherLeads`  •  Manager → `programManagers`
2. **For new students:** open that section's **roster file** and add the new student rows with
   Status = `Active` (Roll No. can be anything; the script re-numbers new rows).
3. Open **Extensions → Apps Script** → run **`manual_updateSheets`**.
   - **Permissions:** it calls `updatePermissionsIfNeeded` — it **adds** access for any new
     teacher/lead/manager/stakeholder email. It is **additive only**: it does **not revoke**
     a removed teacher's access (remove that manually in Drive sharing if required).
   - **New students:** they're appended to every month tab. For past months / earlier days in
     the current month, those cells are filled with placeholders (no false "absent").
   - Formulas, holidays, and styling are refreshed on every tab.
   - If it logs the `⏳` time-limit message, **re-run** it — it resumes.

**🎬 Recording script (Scenario 3):**
- "Class 5-A has a new teacher and two new students."
- Open `TeacherClassMapping` in the Config folder, change the email for Class 5-A. (show edit)
- Open the `Class_5_A` roster, add two students with Status = Active. (show edit)
- Open Apps Script → run `manual_updateSheets`; read the log lines: `[PERMISSIONS CHECK] ... [ADDING EDIT ACCESS]` and `[+] Found 2 new student(s)`.
- Open the workbook → show the two new rows and the new teacher under File → Share.
- Close: "One function keeps both people-access and the roster in sync."

---

## Scenario 4 — A student is dropping out

**Goal:** Stop tracking a student going forward while **preserving** their past attendance.

**Steps:**
1. Open that section's **roster file** in the Student Class List folder.
2. Find the student's row and set **Status = `Inactive`** (do **not** delete the row — keeping it
   preserves the mapping/history).
3. Open **Extensions → Apps Script** → run **`manual_updateSheets`**.
   - Inactive students are **not added** and existing rows are **retained**, so the student's
     historical attendance stays intact in the workbook.
   - Going forward the student is dropped from **active-roster reporting** (e.g. the weekly
     stakeholder report and alerts, which use the active roster).
4. **Optional (recommended):** so the dropped student no longer appears in *future daily forms*,
   remove their row from the **current/future month tabs** manually in the workbook (the daily
   form is built from the names in the month tab). Leave past months untouched for history.

> Why not auto-delete? The system deliberately keeps inactive students in the workbook so their
> recorded attendance is never lost. Marking `Inactive` is the intended, reversible action —
> set it back to `Active` and re-run `manual_updateSheets` if they return.

**🎬 Recording script (Scenario 4):**
- "Diya is leaving Class 5-A. We keep her past attendance but stop tracking her."
- Open the `Class_5_A` roster → set Diya's Status to `Inactive`. (show edit + save)
- Open Apps Script → run `manual_updateSheets`; note the log shows no new adds and no errors.
- Open the workbook → show her past marks are still there.
- "Optionally, I remove her row from the current month tab so she's off future forms." (show)
- Close: "Clean exit — history preserved, future tracking stopped."

---

## Scenario 5 — A student changes section

**Goal:** Move a student from one section to another (e.g. **5-A → 5-B**), keeping their old
attendance in the old workbook and starting them fresh in the new one. This is a
**dropout from the old section + a join to the new section**.

**Steps:**
1. **Old section roster** (`Class_5_A`): set that student's **Status = `Inactive`**
   (keeps their 5-A history; stops future tracking there). Do not delete the row.
2. **New section roster** (`Class_5_B`): **add** the student as a new row with
   **Status = `Active`** (use the same **Child ID** so identity is consistent).
3. Open **Extensions → Apps Script** → run **`manual_updateSheets`** once.
   - 5-A: student stays inactive/retained (history preserved).
   - 5-B: student is appended to every month tab; past months/earlier days get placeholders.
4. **Optional:** remove the student's row from 5-A's **current/future** month tabs so they
   no longer appear in 5-A's daily forms (leave past months for history).
5. If the student should now appear in **5-B's** daily form for a past date already elapsed,
   use **Scenario 2** (`manual_sendOnDemandForm`) for 5-B to capture that day.

**🎬 Recording script (Scenario 5):**
- "Aarav is moving from 5-A to 5-B."
- Open `Class_5_A` roster → set Aarav's Status to `Inactive`. (save)
- Open `Class_5_B` roster → add Aarav with the same Child ID, Status `Active`. (save)
- Open Apps Script → run `manual_updateSheets`.
- Show 5-A workbook: Aarav's old marks preserved. Show 5-B workbook: Aarav now appears in the month tabs.
- "Optionally remove him from 5-A's current month tab, and use an on-demand form if 5-B needs a past day."
- Close: "Section move done — old history kept, new section tracking him."

---

## Quick function reference

| Function | What it does | Scenario |
|----------|--------------|----------|
| `manual_validateSetupConfig` | Checks folders, roster count, academic months | 1 |
| `manual_generateSheets` | Creates all class workbooks (resumable) | 1 |
| `manual_installTriggers` | Installs the 4 automation triggers (run once) | 1 |
| `manual_sendOnDemandForm` | Sends a form for any date (uses `ONDEMAND_*`) | 2 |
| `manual_syncResponses` | Pulls form responses into sheets now | 2 |
| `manual_closeOnDemandForm` | Closes active form(s); supports `FILTER_*` | 2 |
| `manual_updateSheets` | Adds active students + refreshes permissions/formulas (resumable) | 3, 4, 5 |
| `automated_sendDailyForms` | (auto, 6 AM) daily forms | — |
| `automated_syncResponses` | (auto, hourly) response sync | — |
| `automated_closeForms` | (auto, 11 PM) close + dashboard refresh | — |
| `automated_sendWeeklyReport` | (auto, Fri 5 PM) stakeholder report | — |

## General recording tips
- Record at 1080p; zoom the Apps Script editor font to ~16px so the function dropdown is readable.
- Keep the **Execution log** panel visible when running any function.
- Use a throwaway class (e.g. a test `Class_9_Z`) for edits so real data isn't shown.
- Blur/skip real student names and teacher emails, or use sample data.
- Show the **result in Drive/the workbook** after each function — not just the log.
