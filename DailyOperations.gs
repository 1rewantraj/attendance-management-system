// =========================================================================
// AUTOMATED (6 AM daily trigger): Daily Morning Form Distribution
// =========================================================================
function automated_sendDailyForms() {
  var today = new Date();
  if (isGlobalHolidayOrWeekend(today)) {
    Logger.log("⏭️ Today is a weekend or public holiday. Forms will not be sent.");
    return;
  }

  var todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd-MMM-yyyy");
  var configFolder = getFolderByLink(CONFIG_FOLDER_LINK);
  var rosterFolder = getFolderByLink(STUDENT_CLASS_LIST_FOLDER_LINK);
  var outputFolder = getFolderByLink(ATTENDANCE_SHEETS_FOLDER_LINK);
  
  var rosterFiles = rosterFolder.getFiles();
  var teacherMapping = loadTeacherMapping();
  var masterConfig = buildMasterConfig(configFolder);
  var holidays = getPublicHolidays(configFolder);

  while (rosterFiles.hasNext()) {
    var rosterFile = rosterFiles.next();
    var fileName = rosterFile.getName();
    var parsed = parseClassAndSectionFromText(fileName);
    var classNum = parsed.classNum, section = parsed.section;
    var mapKey = classNum.toLowerCase() + "_" + section.toLowerCase();

    if (!teacherMapping[mapKey]) continue;

    var teacherName = teacherMapping[mapKey].name;
    var teacherEmail = teacherMapping[mapKey].email;

    var ssName = getWorkbookName(classNum, section);
    var existingSsArray = outputFolder.getFilesByName(ssName);
    var ss, ssFile;

    if (existingSsArray.hasNext()) {
      ssFile = existingSsArray.next();
      ss = SpreadsheetApp.openById(ssFile.getId());
    } else {
      ss = SpreadsheetApp.create(ssName);
      ssFile = DriveApp.getFileById(ss.getId());
      ssFile.moveTo(outputFolder);

      var rosterData = parseAndNormalizeData(rosterFile, false);
      var monthsToCreate = getAcademicMonthsList(ACADEMIC_YEAR, START_MONTH, END_MONTH);
      buildAttendanceWorkbook(ss, rosterData, monthsToCreate, holidays);
      applyPermissionsToSpreadsheet(ssFile, classNum, section, masterConfig);
    }

    var monthName = today.toLocaleString('en-US', { month: 'long' });
    var sheet = ss.getSheetByName(monthName);
    if (!sheet) continue;

    var studentNamesRange = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues();
    var studentNames = [];
    for (var r = 0; r < studentNamesRange.length; r++) {
      if (studentNamesRange[r][0].toString().trim() !== "") {
        studentNames.push(studentNamesRange[r][0].toString().trim());
      }
    }

    if (studentNames.length === 0) continue;

    var formTitle = "Attendance: Class " + classNum + "-" + section.toUpperCase() + " (" + todayStr + ")";
    var form = FormApp.create(formTitle);
    form.setDescription("Mark attendance for " + todayStr);
    // VERIFIED = auto-capture the signed-in Google account email as read-only
    // (NOT a free-text box the teacher can type/mistype). This is the email the
    // hourly sync matches against AUTHORIZED_TEACHER_ for authorization.
    form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);
    form.setLimitOneResponsePerUser(true);  // ← Requires Google login + limits to 1 response
    form.setAllowResponseEdits(true);
    form.setConfirmationMessage("✅ Thank you! You will receive a confirmation email shortly.");

    var gridItem = form.addGridItem();
    gridItem.setTitle("Today's Attendance for Class " + classNum + "-" + section.toUpperCase());
    gridItem.setRows(studentNames);
    gridItem.setColumns(['Present', 'Absent', 'Late']);
    gridItem.setRequired(true);

    var formFile = DriveApp.getFileById(form.getId());
    formFile.moveTo(outputFolder);
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

    var liveUrl = form.getPublishedUrl();
    var props = PropertiesService.getScriptProperties();
    props.setProperty('ACTIVE_FORM_' + ss.getId(), form.getId());
    props.setProperty('AUTHORIZED_TEACHER_' + ss.getId(), teacherEmail.toLowerCase());

    // Store the target date for this form (for consistency with on-demand forms)
    // Format: YYYY-MM-DD for reliable parsing
    var todayKey = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
    props.setProperty('FORM_TARGET_DATE_' + form.getId(), todayKey);

    var alertsResult = generateAlertBlocks(ss, sheet, studentNames, today);
    var teacherAlertsHtml = alertsResult.teacherHtml || "";

    var htmlBody = buildSimpleHtmlEmail(todayStr, classNum, section, teacherName, liveUrl, teacherAlertsHtml);

    var mailOptions = {
      to: teacherEmail,
      subject: "📋 Daily Attendance Form: Class " + classNum + "-" + section.toUpperCase() + " (" + todayStr + ")",
      htmlBody: htmlBody
    };

    if (alertsResult.teacherBlob) {
      mailOptions.inlineImages = { teacher_chart: alertsResult.teacherBlob };
    }

    MailApp.sendEmail(mailOptions);
    Logger.log("✅ Sent: " + teacherEmail + " | Class " + classNum + "-" + section);
  }
}

// =========================================================================
// MANUAL (admin run): On-Demand Form Sender (for any date, class, or teacher)
// =========================================================================
function manual_sendOnDemandForm() {
  var CLASS_NUM = ONDEMAND_CLASS_NUM || "5";
  var SECTION = ONDEMAND_SECTION || "A";

  // Parse date: try DD-MM-YYYY format first, then fall back to today
  var targetDate;
  if (ONDEMAND_DATE && ONDEMAND_DATE.trim() !== "") {
    targetDate = parseDDMMYYYY(ONDEMAND_DATE);
    if (!targetDate) {
      throw new Error("Invalid date format '" + ONDEMAND_DATE + "'. Use DD-MM-YYYY format (e.g., '05-07-2026' for July 5, 2026).");
    }
  } else {
    targetDate = new Date(); // Today
  }

  var dateStr = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "dd-MMM-yyyy");

  var configFolder = getFolderByLink(CONFIG_FOLDER_LINK);
  var outputFolder = getFolderByLink(ATTENDANCE_SHEETS_FOLDER_LINK);
  var teacherMapping = loadTeacherMapping();
  var mapKey = CLASS_NUM.toLowerCase() + "_" + SECTION.toLowerCase();

  if (!teacherMapping[mapKey]) throw new Error("No teacher found for " + CLASS_NUM + "-" + SECTION);

  var teacherName = ONDEMAND_TEACHER_NAME || teacherMapping[mapKey].name;
  var teacherEmail = ONDEMAND_TEACHER_EMAIL || teacherMapping[mapKey].email;

  var ssName = getWorkbookName(CLASS_NUM, SECTION);
  var ssFiles = outputFolder.getFilesByName(ssName);
  if (!ssFiles.hasNext()) throw new Error("No workbook found");
  
  var ss = SpreadsheetApp.openById(ssFiles.next().getId());
  var monthName = targetDate.toLocaleString('en-US', { month: 'long' });
  var sheet = ss.getSheetByName(monthName);
  if (!sheet) throw new Error("No sheet for " + monthName);

  var studentNamesRange = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues();
  var studentNames = [];
  for (var r = 0; r < studentNamesRange.length; r++) {
    if (studentNamesRange[r][0].toString().trim() !== "") {
      studentNames.push(studentNamesRange[r][0].toString().trim());
    }
  }

  var formTitle = "Attendance (Makeup): Class " + CLASS_NUM + "-" + SECTION.toUpperCase() + " (" + dateStr + ")";
  var form = FormApp.create(formTitle);
  form.setDescription("Makeup form for " + dateStr);
  // VERIFIED = auto-capture the signed-in account email as read-only (see note
  // in automated_sendDailyForms). Requires Google login by definition.
  form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);
  form.setAllowResponseEdits(true);
  form.setConfirmationMessage("✅ Thank you! You will receive a confirmation email shortly.");

  var gridItem = form.addGridItem();
  gridItem.setTitle("Attendance for " + dateStr);
  gridItem.setRows(studentNames);
  gridItem.setColumns(['Present', 'Absent', 'Late']);
  gridItem.setRequired(true);

  var formFile = DriveApp.getFileById(form.getId());
  formFile.moveTo(outputFolder);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  var liveUrl = form.getPublishedUrl();
  var props = PropertiesService.getScriptProperties();
  props.setProperty('ACTIVE_FORM_' + ss.getId(), form.getId());
  props.setProperty('AUTHORIZED_TEACHER_' + ss.getId(), teacherEmail.toLowerCase());

  // Store the target date for this form so sync knows which day column to update
  // Format: YYYY-MM-DD for reliable parsing
  var targetDateKey = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  props.setProperty('FORM_TARGET_DATE_' + form.getId(), targetDateKey);

  var alertsResult = generateAlertBlocks(ss, sheet, studentNames, targetDate);
  var teacherAlertsHtml = alertsResult.teacherHtml || "";
  var htmlBody = buildSimpleHtmlEmail(dateStr, CLASS_NUM, SECTION, teacherName, liveUrl, teacherAlertsHtml);

  var mailOptions = {
    to: teacherEmail,
    subject: "📋 Makeup Attendance Form: Class " + CLASS_NUM + "-" + SECTION.toUpperCase() + " (" + dateStr + ")",
    htmlBody: htmlBody
  };

  if (alertsResult.teacherBlob) {
    mailOptions.inlineImages = { teacher_chart: alertsResult.teacherBlob };
  }

  MailApp.sendEmail(mailOptions);
  Logger.log("✅ Ad-hoc form sent to: " + teacherEmail);
}

// =========================================================================
// AUTOMATED (hourly trigger): Response Sync
// =========================================================================
function automated_syncResponses() {
  var outputFolder = getFolderByLink(ATTENDANCE_SHEETS_FOLDER_LINK);
  var files = outputFolder.getFiles();
  var props = PropertiesService.getScriptProperties();

  // Track which forms/workbooks are LIVE this pass so the orphan sweep below
  // never deletes a form teachers are still submitting to (nor its active
  // response tab). Keyed by id for O(1) lookup in performOrphanCleanup.
  var activeFormIds = {};
  var activeSsIds = {};

  while (files.hasNext()) {
    var file = files.next();
    if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

    var ssId = file.getId();
    var activeFormId = props.getProperty('ACTIVE_FORM_' + ssId);
    if (!activeFormId) continue;

    activeFormIds[activeFormId] = true;
    activeSsIds[ssId] = true;

    try {
      var ss = SpreadsheetApp.openById(ssId);

      // Check if this form has a stored target date (for on-demand forms)
      var targetDateStr = props.getProperty('FORM_TARGET_DATE_' + activeFormId);
      var targetDate;

      if (targetDateStr) {
        // On-demand form: use the stored date
        targetDate = new Date(targetDateStr);
        Logger.log("  [SYNC] Form " + activeFormId + " is for stored date: " + targetDateStr);
      } else {
        // Regular daily form: use today
        targetDate = new Date();
      }

      var monthName = targetDate.toLocaleString('en-US', { month: 'long' });
      var sheet = ss.getSheetByName(monthName);
      if (!sheet) {
        Logger.log("  [SYNC][WARNING] No sheet for month '" + monthName + "' in " + file.getName());
        continue;
      }

      var dayOfMonthDigit = targetDate.getDate();

      // 1. Process attendance responses into sheet
      executeSheetSyncProcessing(sheet, activeFormId, dayOfMonthDigit, ssId);

      // 2. Ensure Dashboard exists & refresh metrics/charts
      createDashboardIfNotExists(ss);
      updateDashboard(ss);

    } catch (err) {
      Logger.log("Error syncing " + file.getName() + ": " + err.message);
    }
  }

  // 3. Sweep orphaned forms + response tabs from the Drive folder, preserving
  //    anything still live today. This keeps the folder clean between the
  //    nightly automated_closeForms runs (e.g. after a form was manually
  //    closed or its ACTIVE_FORM_ pointer was cleared).
  try {
    var swept = performOrphanCleanup(outputFolder, activeFormIds, activeSsIds);
    if (swept.formsDeleted > 0 || swept.tabsDeleted > 0) {
      Logger.log("🧹 Sync sweep: trashed " + swept.formsDeleted +
                 " orphan form(s), deleted " + swept.tabsDeleted + " orphan tab(s).");
    }
  } catch (sweepErr) {
    Logger.log("Orphan sweep during sync failed: " + sweepErr.message);
  }
}

// =========================================================================
// MANUAL (admin run): Manual Response Sync
// =========================================================================
function manual_syncResponses() {
  Logger.log("=========================================================");
  Logger.log("MANUAL SYNC: Syncing all active form responses");
  Logger.log("=========================================================");

  var outputFolder = getFolderByLink(ATTENDANCE_SHEETS_FOLDER_LINK);
  var files = outputFolder.getFiles();
  var props = PropertiesService.getScriptProperties();
  var syncCount = 0;

  while (files.hasNext()) {
    var file = files.next();
    if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

    var ssId = file.getId();
    var activeFormId = props.getProperty('ACTIVE_FORM_' + ssId);
    if (!activeFormId) continue;

    try {
      var ss = SpreadsheetApp.openById(ssId);

      // Check if this form has a stored target date (for on-demand forms)
      var targetDateStr = props.getProperty('FORM_TARGET_DATE_' + activeFormId);
      var targetDate;

      if (targetDateStr) {
        // On-demand form: use the stored date
        targetDate = new Date(targetDateStr);
        Logger.log("--> Syncing form for " + file.getName() + " (target date: " + targetDateStr + ")");
      } else {
        // Regular daily form: use today
        targetDate = new Date();
        Logger.log("--> Syncing form for " + file.getName() + " (today)");
      }

      var monthName = targetDate.toLocaleString('en-US', { month: 'long' });
      var sheet = ss.getSheetByName(monthName);
      if (!sheet) {
        Logger.log("    [WARNING] No sheet for month '" + monthName + "'. Skipping.");
        continue;
      }

      var dayOfMonthDigit = targetDate.getDate();

      // Process attendance responses into sheet
      executeSheetSyncProcessing(sheet, activeFormId, dayOfMonthDigit, ssId);

      // Ensure Dashboard exists & refresh metrics/charts
      createDashboardIfNotExists(ss);
      updateDashboard(ss);

      syncCount++;

    } catch (err) {
      Logger.log("    [ERROR] Syncing " + file.getName() + ": " + err.message);
    }
  }

  Logger.log("=========================================================");
  Logger.log("SUCCESS: Synced " + syncCount + " form(s)");
  Logger.log("=========================================================");
}

// =========================================================================
// AUTOMATED (11 PM daily trigger): Close All Active Forms
// =========================================================================
function automated_closeForms() {
  var outputFolder = getFolderByLink(ATTENDANCE_SHEETS_FOLDER_LINK);
  var files = outputFolder.getFiles();
  var props = PropertiesService.getScriptProperties();

  while (files.hasNext()) {
    var file = files.next();
    if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

    var ssId = file.getId();
    var activeFormId = props.getProperty('ACTIVE_FORM_' + ssId);
    if (!activeFormId) continue;

    try {
      var form = FormApp.openById(activeFormId);
      var formTitle = form.getTitle();

      // 1. Stop accepting new responses.
      form.setAcceptingResponses(false);

      // 2. FINAL SYNC before we destroy the form: capture any responses that
      //    arrived since the last hourly sync, so no attendance is lost when the
      //    form + its response tab are deleted below. Attendance lives in the
      //    month tabs' day-columns — NOT in the "Form Responses" tab — so once
      //    this sync runs, the form and its response tab are safe to remove.
      var ss = SpreadsheetApp.openById(ssId);

      // Check if this form has a stored target date (for on-demand forms)
      var targetDateStr = props.getProperty('FORM_TARGET_DATE_' + activeFormId);
      var targetDate;

      if (targetDateStr) {
        // On-demand form: use the stored date
        targetDate = new Date(targetDateStr);
      } else {
        // Regular daily form: use today
        targetDate = new Date();
      }

      var monthSheet = ss.getSheetByName(targetDate.toLocaleString('en-US', { month: 'long' }));
      if (monthSheet) {
        executeSheetSyncProcessing(monthSheet, activeFormId, targetDate.getDate(), ssId);
      }

      // 3. Unlink the form's response destination, then delete every
      //    "Form Responses N" tab (they accumulate one per day otherwise).
      try { form.removeDestination(); } catch (e) { /* not always linkable; ignore */ }
      var allSheets = ss.getSheets();
      for (var s = 0; s < allSheets.length; s++) {
        var name = allSheets[s].getName();
        if (name.indexOf('Form Responses') === 0 && ss.getSheets().length > 1) {
          ss.deleteSheet(allSheets[s]);
          Logger.log("   🧹 Deleted response tab: " + name + " (" + file.getName() + ")");
        }
      }

      // 4. Trash the Form file from the output folder.
      DriveApp.getFileById(activeFormId).setTrashed(true);

      // 5. Clear runtime state. NOTIFIED_ reset so tomorrow's submitters get
      //    their one confirmation email again (see notifySubmitters in Utils.gs).
      props.deleteProperty('ACTIVE_FORM_' + ssId);
      props.deleteProperty('NOTIFIED_' + ssId);
      props.deleteProperty('FORM_TARGET_DATE_' + activeFormId);

      Logger.log("🔒 Closed + cleaned up form: " + formTitle);
    } catch (err) {
      Logger.log("Error closing form for " + file.getName() + ": " + err.message);
    }
  }
}

// =========================================================================
// MANUAL (admin run): Flexible Form Closer
// =========================================================================
function manual_closeFormsFlexibly() {
  var FILTER_CLASS = null;
  var FILTER_SECTION = null;
  var FILTER_DATE_STRING = null;

  var outputFolder = getFolderByLink(ATTENDANCE_SHEETS_FOLDER_LINK);
  var files = outputFolder.getFiles();
  var props = PropertiesService.getScriptProperties();

  while (files.hasNext()) {
    var file = files.next();
    if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

    var ssId = file.getId();
    var activeFormId = props.getProperty('ACTIVE_FORM_' + ssId);
    if (!activeFormId) continue;

    try {
      var form = FormApp.openById(activeFormId);
      var formTitle = form.getTitle();

      var shouldClose = true;
      if (FILTER_CLASS && formTitle.toLowerCase().indexOf("class " + FILTER_CLASS.toLowerCase()) === -1) shouldClose = false;
      if (FILTER_SECTION && formTitle.toLowerCase().indexOf(FILTER_SECTION.toLowerCase()) === -1) shouldClose = false;
      if (FILTER_DATE_STRING && formTitle.indexOf(FILTER_DATE_STRING) === -1) shouldClose = false;

      if (shouldClose) {
        form.setAcceptingResponses(false);
        props.deleteProperty('ACTIVE_FORM_' + ssId);
        props.deleteProperty('FORM_TARGET_DATE_' + activeFormId);
        Logger.log("🔒 Closed: " + formTitle);
      }
    } catch (err) {
      Logger.log("Error: " + err.message);
    }
  }
}

// =========================================================================
// SHARED: Orphan Form + Response-Tab Cleanup (folder scan)
// -------------------------------------------------------------------------
// Removes ORPHANS by scanning the output folder, independent of any Script
// Property: attendance Form FILES that are no longer live, and stray
// "Form Responses N" TABS left behind in workbooks.
//
// LIVE PRESERVATION: pass activeFormIds / activeSsIds (maps keyed by id) to
// protect forms/tabs that are still in use TODAY — the hourly sync passes the
// current ACTIVE_FORM_ set so it never deletes a form teachers are still
// submitting to, nor the response tab that form is actively appending to.
// Pass empty maps ({}, {}) to nuke everything (the manual full cleanup).
//
// Data-safe: for each form still linked to a workbook it runs a final sync
// (writing attendance into the month tabs) BEFORE trashing the form. Attendance
// lives in the month tabs — never in the "Form Responses" tab — so nothing is
// lost. Returns { formsDeleted, tabsDeleted }.
// =========================================================================
function performOrphanCleanup(outputFolder, activeFormIds, activeSsIds) {
  activeFormIds = activeFormIds || {};
  activeSsIds = activeSsIds || {};
  var formsDeleted = 0, tabsDeleted = 0;

  // 1. Trash orphan Attendance Form files (skip any that are still live).
  var formFiles = outputFolder.getFilesByType(MimeType.GOOGLE_FORMS);
  while (formFiles.hasNext()) {
    var formFile = formFiles.next();
    var fid = formFile.getId();
    var title = formFile.getName();
    // Only touch attendance forms this system generates.
    if (title.indexOf("Attendance") !== 0) continue;
    // Never delete a form that is still the active form for some workbook.
    if (activeFormIds[fid]) continue;

    try {
      var form = FormApp.openById(fid);
      form.setAcceptingResponses(false);

      // Final sync if the form is linked to a workbook we can resolve.
      var destId = null;
      try { destId = form.getDestinationId(); } catch (e) { destId = null; }
      if (destId) {
        try {
          var ss = SpreadsheetApp.openById(destId);
          var today = new Date();
          var monthSheet = ss.getSheetByName(today.toLocaleString('en-US', { month: 'long' }));
          if (monthSheet) {
            executeSheetSyncProcessing(monthSheet, fid, today.getDate(), destId);
          }
        } catch (e2) {
          Logger.log("   [WARN] Final sync failed for " + title + ": " + e2.message);
        }
      }
      try { form.removeDestination(); } catch (e3) { /* ignore */ }
    } catch (eForm) {
      Logger.log("   [WARN] Could not open form " + title + ": " + eForm.message);
    }

    formFile.setTrashed(true);
    formsDeleted++;
    Logger.log("🧹 Trashed orphan form file: " + title);
  }

  // 2. Delete leftover "Form Responses N" tabs. Skip workbooks that still have
  //    a live form — that form legitimately owns its response tab until close.
  var sheetFiles = outputFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (sheetFiles.hasNext()) {
    var sf = sheetFiles.next();
    if (activeSsIds[sf.getId()]) continue;
    try {
      var wb = SpreadsheetApp.openById(sf.getId());
      var tabs = wb.getSheets();
      for (var s = 0; s < tabs.length; s++) {
        var tabName = tabs[s].getName();
        if (tabName.indexOf('Form Responses') === 0 && wb.getSheets().length > 1) {
          wb.deleteSheet(tabs[s]);
          tabsDeleted++;
          Logger.log("🧹 Deleted orphan response tab '" + tabName + "' in " + sf.getName());
        }
      }
    } catch (eWb) {
      Logger.log("   [WARN] Could not clean workbook " + sf.getName() + ": " + eWb.message);
    }
  }

  return { formsDeleted: formsDeleted, tabsDeleted: tabsDeleted };
}

// =========================================================================
// MANUAL (admin run): Full Orphan Form + Response-Tab Cleanup
// -------------------------------------------------------------------------
// Nukes ALL attendance forms + response tabs in the output folder (passes
// empty active maps), then clears any dangling ACTIVE_FORM_ / NOTIFIED_
// pointers. Use when the daily automation left orphans behind.
// =========================================================================
function manual_cleanupOrphanForms() {
  var outputFolder = getFolderByLink(ATTENDANCE_SHEETS_FOLDER_LINK);

  var res = performOrphanCleanup(outputFolder, {}, {});

  // Clear any lingering ACTIVE_FORM_ / NOTIFIED_ pointers (now dangling).
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  for (var key in all) {
    if (key.indexOf('ACTIVE_FORM_') === 0 || key.indexOf('NOTIFIED_') === 0) {
      props.deleteProperty(key);
    }
  }

  Logger.log("✅ Orphan cleanup complete. Forms trashed: " + res.formsDeleted +
             ", response tabs deleted: " + res.tabsDeleted + ".");
}

// =========================================================================
// AUTOMATED (Friday 5 PM trigger): Weekly Stakeholder Report
// =========================================================================
function automated_sendWeeklyReport() {
  var today = new Date();
  var todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd-MMM-yyyy");
  var configFolder = getFolderByLink(CONFIG_FOLDER_LINK);
  var outputFolder = getFolderByLink(ATTENDANCE_SHEETS_FOLDER_LINK);
  var masterConfig = buildMasterConfig(configFolder);

  var digestContent = "";
  var allCharts = [];
  var files = outputFolder.getFiles();

  while (files.hasNext()) {
    var file = files.next();
    if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

    var ss = SpreadsheetApp.openById(file.getId());
    var monthName = today.toLocaleString('en-US', { month: 'long' });
    var sheet = ss.getSheetByName(monthName);
    if (!sheet) continue;

    var studentNamesRange = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues();
    var studentNames = [];
    for (var r = 0; r < studentNamesRange.length; r++) {
      if (studentNamesRange[r][0].toString().trim() !== "") {
        studentNames.push(studentNamesRange[r][0].toString().trim());
      }
    }

    if (studentNames.length === 0) continue;

    var alertsResult = generateAlertBlocks(ss, sheet, studentNames, today);
    if (alertsResult.stakeholderHtml && alertsResult.stakeholderHtml !== "") {
      digestContent += '<h3 style="color: #2d3748; margin-top: 25px;">' + file.getName() + '</h3>';
      // Each workbook's chart needs a UNIQUE inline-image CID, otherwise all
      // blocks reference the same "stakeholder_chart_cid" and the images
      // collide / render broken. generateAlertBlocks emits the placeholder
      // src="cid:stakeholder_chart_cid"; swap it here for a per-file CID that
      // matches what we register in allCharts below.
      var uniqueCid = 'stakeholder_chart_cid_' + file.getId();
      digestContent += alertsResult.stakeholderHtml.replace('cid:stakeholder_chart_cid', 'cid:' + uniqueCid);
      if (alertsResult.stakeholderBlob) {
        allCharts.push({ cid: uniqueCid, blob: alertsResult.stakeholderBlob });
      }
    }
  }

  if (digestContent === "") {
    Logger.log("No chronic alerts to report.");
    return;
  }

  var htmlBody = buildStakeholderDigestHtml(todayStr, digestContent);
  var recipientList = STAKEHOLDER_EMAILS.split(",").map(function(email) { return email.trim(); });

  if (recipientList.length === 0) {
    Logger.log("No stakeholders configured.");
    return;
  }

  var mailOptions = {
    to: recipientList.join(","),
    subject: "📊 Weekly Attendance Report - " + todayStr,
    htmlBody: htmlBody
  };

  if (allCharts.length > 0) {
    var inlineImagesObj = {};
    for (var c = 0; c < allCharts.length; c++) {
      inlineImagesObj[allCharts[c].cid] = allCharts[c].blob;
    }
    mailOptions.inlineImages = inlineImagesObj;
  }

  MailApp.sendEmail(mailOptions);
  Logger.log("✅ Stakeholder report sent to: " + recipientList.join(", "));
}

// =========================================================================
// AUTOMATED (11 PM daily trigger): End-of-Day Visualization Refresh
// -------------------------------------------------------------------------
// Rebuilds the Analysis_Dashboard for every class workbook once at end of day,
// so charts/metrics reflect the full day's synced attendance (including the
// final sync done by automated_closeForms). The hourly sync also refreshes the
// dashboard, but this guarantees a clean, complete end-of-day snapshot.
// =========================================================================
function automated_refreshVisualizations() {
  var outputFolder = getFolderByLink(ATTENDANCE_SHEETS_FOLDER_LINK);
  var files = outputFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var refreshed = 0;

  while (files.hasNext()) {
    var file = files.next();
    try {
      var ss = SpreadsheetApp.openById(file.getId());
      createDashboardIfNotExists(ss);  // create if missing (also populates it)
      updateDashboard(ss);             // recompute metrics/charts from month tabs
      refreshed++;
      Logger.log("📊 Refreshed dashboard: " + file.getName());
    } catch (err) {
      Logger.log("Error refreshing visualization for " + file.getName() + ": " + err.message);
    }
  }

  Logger.log("✅ End-of-day visualization refresh complete. " + refreshed + " workbook(s) updated.");
}

// =========================================================================
// MANUAL (admin run): Refresh visualizations on demand (same as the automated
// end-of-day refresh, but runnable any time).
// =========================================================================
function manual_refreshVisualizations() {
  automated_refreshVisualizations();
}

// =========================================================================
// MANUAL (one-time admin run): Trigger Installer — registers each automated_* function
// =========================================================================
function manual_installTriggers() {
  ScriptApp.newTrigger('automated_sendDailyForms')
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .create();

  ScriptApp.newTrigger('automated_syncResponses')
    .timeBased()
    .everyHours(1)
    .create();

  ScriptApp.newTrigger('automated_closeForms')
    .timeBased()
    .atHour(23)
    .everyDays(1)
    .create();

  // Runs after automated_closeForms (11 PM) so the end-of-day dashboard reflects
  // the full day's synced attendance, including that final close-time sync.
  ScriptApp.newTrigger('automated_refreshVisualizations')
    .timeBased()
    .atHour(23)
    .nearMinute(45)
    .everyDays(1)
    .create();

  ScriptApp.newTrigger('automated_sendWeeklyReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(17)
    .create();

  Logger.log("✅ All triggers installed successfully. (5 total: 6AM forms, hourly sync, 11PM close, 11:45PM viz refresh, Fri 5PM report)");
}

// =========================================================================
// MANUAL (admin run): Configuration Validator
// =========================================================================
function manual_validateConfig() {
  var errors = [];
  
  try {
    var configFolder = getFolderByLink(CONFIG_FOLDER_LINK);
    Logger.log("✅ Config folder accessible: " + configFolder.getName());
  } catch(e) {
    errors.push("❌ Cannot access CONFIG_FOLDER_LINK: " + e.message);
  }
  
  try {
    var rosterFolder = getFolderByLink(STUDENT_CLASS_LIST_FOLDER_LINK);
    Logger.log("✅ Input folder accessible: " + rosterFolder.getName());
  } catch(e) {
    errors.push("❌ Cannot access STUDENT_CLASS_LIST_FOLDER_LINK: " + e.message);
  }
  
  try {
    var outputFolder = getFolderByLink(ATTENDANCE_SHEETS_FOLDER_LINK);
    Logger.log("✅ Output folder accessible: " + outputFolder.getName());
  } catch(e) {
    errors.push("❌ Cannot access ATTENDANCE_SHEETS_FOLDER_LINK: " + e.message);
  }
  
  if (errors.length > 0) {
    Logger.log("\n⚠️ CONFIGURATION ERRORS:");
    for (var i = 0; i < errors.length; i++) {
      Logger.log("   " + errors[i]);
    }
    throw new Error("Configuration validation failed. See logs.");
  }
  
  Logger.log("\n🎉 CONFIGURATION VALID! All folders are accessible.");
  Logger.log("Academic Year: " + ACADEMIC_YEAR);
  Logger.log("Months: " + START_MONTH + " to " + END_MONTH);
  Logger.log("Visualizations: " + (ADD_VISUALISATIONS ? "Enabled" : "Disabled"));
}
