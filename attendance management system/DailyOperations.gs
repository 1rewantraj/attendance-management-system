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
  var rosterFolder = getFolderByLink(INPUT_FOLDER_LINK);
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
    PropertiesService.getScriptProperties().setProperty('ACTIVE_FORM_' + ss.getId(), form.getId());
    PropertiesService.getScriptProperties().setProperty('AUTHORIZED_TEACHER_' + ss.getId(), teacherEmail.toLowerCase());

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
// MANUAL (admin run): Ad-hoc Makeup Form Generator
// =========================================================================
function manual_runAdhocForm() {
  var CLASS_NUM = ADHOC_CLASS_NUM || "5";
  var SECTION = ADHOC_SECTION || "A";
  var TARGET_DATE_STRING = ADHOC_DATE || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yyyy");
  
  var targetDate = new Date(TARGET_DATE_STRING);
  var dateStr = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "dd-MMM-yyyy");
  
  var configFolder = getFolderByLink(CONFIG_FOLDER_LINK);
  var outputFolder = getFolderByLink(ATTENDANCE_SHEETS_FOLDER_LINK);
  var teacherMapping = loadTeacherMapping();
  var mapKey = CLASS_NUM.toLowerCase() + "_" + SECTION.toLowerCase();

  if (!teacherMapping[mapKey]) throw new Error("No teacher found for " + CLASS_NUM + "-" + SECTION);

  var teacherName = ADHOC_TEACHER_NAME || teacherMapping[mapKey].name;
  var teacherEmail = ADHOC_TEACHER_EMAIL || teacherMapping[mapKey].email;

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
  PropertiesService.getScriptProperties().setProperty('ACTIVE_FORM_' + ss.getId(), form.getId());
  PropertiesService.getScriptProperties().setProperty('AUTHORIZED_TEACHER_' + ss.getId(), teacherEmail.toLowerCase());

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

  while (files.hasNext()) {
    var file = files.next();
    if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

    var ssId = file.getId();
    var activeFormId = props.getProperty('ACTIVE_FORM_' + ssId);
    if (!activeFormId) continue;

    try {
      var ss = SpreadsheetApp.openById(ssId);
      var today = new Date();
      var monthName = today.toLocaleString('en-US', { month: 'long' });
      var sheet = ss.getSheetByName(monthName);
      if (!sheet) continue;

      var dayOfMonthDigit = today.getDate();

      // 1. Process attendance responses into sheet
      executeSheetSyncProcessing(sheet, activeFormId, dayOfMonthDigit, ssId);

      // 2. Ensure Dashboard exists & refresh metrics/charts
      createDashboardIfNotExists(ss);
      updateDashboard(ss);

    } catch (err) {
      Logger.log("Error syncing " + file.getName() + ": " + err.message);
    }
  }
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
      var today = new Date();
      var monthSheet = ss.getSheetByName(today.toLocaleString('en-US', { month: 'long' }));
      if (monthSheet) {
        executeSheetSyncProcessing(monthSheet, activeFormId, today.getDate(), ssId);
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
        Logger.log("🔒 Closed: " + formTitle);
      }
    } catch (err) {
      Logger.log("Error: " + err.message);
    }
  }
}

// =========================================================================
// MANUAL (admin run): Orphan Form + Response-Tab Cleanup
// -------------------------------------------------------------------------
// automated_closeForms only cleans up a workbook whose ACTIVE_FORM_<ssId>
// property is still set. Forms whose pointer was already cleared (or never
// set) become ORPHANS — the Form file lingers in the output folder and its
// "Form Responses N" tab lingers in the workbook. This function finds and
// removes them by SCANNING the folder, independent of any Script Property.
//
// It is data-safe: for each Form still linked to a workbook it runs a final
// sync (writing attendance into the month tabs) BEFORE deleting the form and
// its response tabs. Attendance lives in the month tabs — never in the
// "Form Responses" tab — so nothing is lost.
// =========================================================================
function manual_cleanupOrphanForms() {
  var outputFolder = getFolderByLink(ATTENDANCE_SHEETS_FOLDER_LINK);
  var formsDeleted = 0, tabsDeleted = 0;

  // 1. Delete every Attendance Form file in the output folder. First do a final
  //    sync into its destination workbook so no submitted response is lost.
  var formFiles = outputFolder.getFilesByType(MimeType.GOOGLE_FORMS);
  while (formFiles.hasNext()) {
    var formFile = formFiles.next();
    var title = formFile.getName();
    // Only touch attendance forms this system generates.
    if (title.indexOf("Attendance") !== 0) continue;

    try {
      var form = FormApp.openById(formFile.getId());
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
            executeSheetSyncProcessing(monthSheet, formFile.getId(), today.getDate(), destId);
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
    Logger.log("🧹 Trashed form file: " + title);
  }

  // 2. In every workbook, delete leftover "Form Responses N" tabs.
  var sheetFiles = outputFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (sheetFiles.hasNext()) {
    var sf = sheetFiles.next();
    try {
      var wb = SpreadsheetApp.openById(sf.getId());
      var tabs = wb.getSheets();
      for (var s = 0; s < tabs.length; s++) {
        var tabName = tabs[s].getName();
        if (tabName.indexOf('Form Responses') === 0 && wb.getSheets().length > 1) {
          wb.deleteSheet(tabs[s]);
          tabsDeleted++;
          Logger.log("🧹 Deleted response tab '" + tabName + "' in " + sf.getName());
        }
      }
    } catch (eWb) {
      Logger.log("   [WARN] Could not clean workbook " + sf.getName() + ": " + eWb.message);
    }
  }

  // 3. Clear any lingering ACTIVE_FORM_ / NOTIFIED_ pointers (now dangling).
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  for (var key in all) {
    if (key.indexOf('ACTIVE_FORM_') === 0 || key.indexOf('NOTIFIED_') === 0) {
      props.deleteProperty(key);
    }
  }

  Logger.log("✅ Orphan cleanup complete. Forms trashed: " + formsDeleted +
             ", response tabs deleted: " + tabsDeleted + ".");
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
      digestContent += alertsResult.stakeholderHtml;
      if (alertsResult.stakeholderBlob) {
        allCharts.push({ cid: 'stakeholder_chart_cid_' + file.getId(), blob: alertsResult.stakeholderBlob });
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

  ScriptApp.newTrigger('automated_sendWeeklyReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(17)
    .create();

  Logger.log("✅ All triggers installed successfully.");
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
    var rosterFolder = getFolderByLink(INPUT_FOLDER_LINK);
    Logger.log("✅ Input folder accessible: " + rosterFolder.getName());
  } catch(e) {
    errors.push("❌ Cannot access INPUT_FOLDER_LINK: " + e.message);
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
