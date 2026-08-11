function manual_generateSheets() {
  var startTime = Date.now();
  Logger.log("=========================================================");
  Logger.log("STEP 1: STARTING MASTER BATCH INITIAL GENERATION");
  Logger.log("=========================================================");
  
  try { DriveApp.getRootFolder(); } catch(e) {} 
  
  var inputFolder = getFolderByLink(STUDENT_CLASS_LIST_FOLDER_LINK);
  var outputFolder = getFolderByLink(ATTENDANCE_SHEETS_FOLDER_LINK);
  var configFolder = getFolderByLink(CONFIG_FOLDER_LINK);
  
  var monthsToCreate = getAcademicMonthsList(ACADEMIC_YEAR, START_MONTH, END_MONTH);
  
  Logger.log("--> Loading Master Configurations & Holidays...");
  var masterConfig = buildMasterConfig(configFolder);
  var holidays = getPublicHolidays(configFolder);
  
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('FLOW1_TOKEN');
  var files;
  
  if (token) {
    try {
      files = DriveApp.continueFileIterator(token);
      Logger.log("--> [RESUMING FROM SAVED PROGRESS] Continuing previous batch...");
    } catch(e) {
      files = inputFolder.getFiles(); // Fallback if token expired
    }
  } else {
    files = inputFolder.getFiles();
  }
  
  while (files.hasNext()) {
    // Check if we are running out of time
    if (Date.now() - startTime > MAX_EXECUTION_TIME) {
      props.setProperty('FLOW1_TOKEN', files.getContinuationToken());
      Logger.log("⏳ Execution time limit approaching. Progress saved. Please rerun the script to continue.");
      return; 
    }
    
    var file = files.next();
    var mimeType = file.getMimeType();
    
    if (mimeType !== MimeType.CSV && mimeType !== MimeType.MICROSOFT_EXCEL && mimeType !== "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" && mimeType !== MimeType.GOOGLE_SHEETS) {
      continue;
    }
    
    var fileName = file.getName();
    var info = parseClassAndSectionFromText(fileName);
    var targetSpreadsheetName = getWorkbookName(info.classNum, info.section);
    
    if (outputFolder.getFilesByName(targetSpreadsheetName).hasNext()) {
      Logger.log("[SKIP] " + targetSpreadsheetName + " already exists.");
      continue;
    }
    
    Logger.log("--> Processing roster: " + fileName);
    var rosterData = parseAndNormalizeData(file).filter(function(row) { return row[3] === "active"; });
    if (rosterData.length === 0) {
      Logger.log("    [!] No valid student data found. Skipping.");
      continue;
    }
    
    Logger.log("    Creating new Spreadsheet: " + targetSpreadsheetName);
    var newSs = SpreadsheetApp.create(targetSpreadsheetName);
    var ssFile = DriveApp.getFileById(newSs.getId());
    ssFile.moveTo(outputFolder);
    
    applyPermissionsToSpreadsheet(ssFile, info.classNum, info.section, masterConfig);
    buildAttendanceWorkbook(newSs, rosterData, monthsToCreate, holidays);
  }
  
  props.deleteProperty('FLOW1_TOKEN'); // Clear token on successful completion
  Logger.log("=========================================================");
  Logger.log("SUCCESS: MASTER BATCH GENERATION COMPLETE.");
  Logger.log("=========================================================\n");
}

function manual_updateSheets() {
  var startTime = Date.now();
  Logger.log("=========================================================");
  Logger.log("STEP 2: STARTING MID-YEAR DELTA UPDATE CHECK");
  Logger.log("=========================================================");
  
  var inputFolder = getFolderByLink(STUDENT_CLASS_LIST_FOLDER_LINK);
  var outputFolder = getFolderByLink(ATTENDANCE_SHEETS_FOLDER_LINK);
  var configFolder = getFolderByLink(CONFIG_FOLDER_LINK);
  
  var monthsToCreate = getAcademicMonthsList(ACADEMIC_YEAR, START_MONTH, END_MONTH);

  Logger.log("--> Loading Master Configurations & Holidays...");
  var masterConfig = buildMasterConfig(configFolder);
  var holidays = getPublicHolidays(configFolder);

  if (UPDATE_SHEETS_CLASS_FILTER && UPDATE_SHEETS_SECTION_FILTER) {
    Logger.log("--> [FILTER ACTIVE] Only updating Class " + UPDATE_SHEETS_CLASS_FILTER + "-" + UPDATE_SHEETS_SECTION_FILTER.toString().toUpperCase());
  }

  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('FLOW2_TOKEN');
  var files;

  if (token) {
    try {
      files = DriveApp.continueFileIterator(token);
      Logger.log("--> [RESUMING FROM SAVED PROGRESS] Continuing previous update check...");
    } catch(e) {
      files = inputFolder.getFiles();
    }
  } else {
    files = inputFolder.getFiles();
  }

  var today = new Date();
  var todayYear = today.getFullYear(), todayMonth = today.getMonth(), todayDay = today.getDate();
  
  while (files.hasNext()) {
    if (Date.now() - startTime > MAX_EXECUTION_TIME) {
      props.setProperty('FLOW2_TOKEN', files.getContinuationToken());
      Logger.log("⏳ Execution time limit approaching. Progress saved. Please rerun the script to continue.");
      return; 
    }
    
    var file = files.next();
    var mimeType = file.getMimeType();
    
    if (mimeType !== MimeType.CSV && mimeType !== MimeType.MICROSOFT_EXCEL && mimeType !== "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" && mimeType !== MimeType.GOOGLE_SHEETS) continue;
    
    var info = parseClassAndSectionFromText(file.getName());

    if (UPDATE_SHEETS_CLASS_FILTER && UPDATE_SHEETS_SECTION_FILTER) {
      var classMatches = info.classNum.toString().trim() === UPDATE_SHEETS_CLASS_FILTER.toString().trim();
      var sectionMatches = info.section.toString().trim().toUpperCase() === UPDATE_SHEETS_SECTION_FILTER.toString().trim().toUpperCase();
      if (!classMatches || !sectionMatches) continue;
    }

    var targetSpreadsheetName = getWorkbookName(info.classNum, info.section);
    var ssFiles = outputFolder.getFilesByName(targetSpreadsheetName);

    if (!ssFiles.hasNext()) continue;

    var ssFile = ssFiles.next();
    var ss = SpreadsheetApp.open(ssFile);
    var csvRoster = parseAndNormalizeData(file);
    if (csvRoster.length === 0) continue;

    var statusById = {};
    csvRoster.forEach(function(row) { statusById[row[1].toString().trim()] = row[3]; });

    Logger.log("--> Checking updates for: " + targetSpreadsheetName);

    // Check and update permissions if needed
    Logger.log("    [PERMISSIONS CHECK] Verifying access for teachers, leads, and managers...");
    updatePermissionsIfNeeded(ssFile, info.classNum, info.section, masterConfig);
    
    for (var m = 0; m < monthsToCreate.length; m++) {
      var monthInfo = monthsToCreate[m];
      var sheet = ss.getSheetByName(monthInfo.name);
      if (!sheet) continue;

      var isPastMonth = (monthInfo.year < todayYear) || (monthInfo.year === todayYear && monthInfo.monthIndex < todayMonth);
      var isCurrentMonth = (monthInfo.year === todayYear && monthInfo.monthIndex === todayMonth);
      var lastRow = sheet.getLastRow();
      var existingIds = [];
      var idToRow = {};
      var currentStudentCount = 0;

      // Count existing students: scan Child ID (col B) from row 2 until the
      // student block ends (blank ID, or Roll No. in col A is no longer numeric
      // — i.e. we've hit the footer/CLASS AVERAGE rows). Cols A and B are read
      // in one batch call each; a per-row getValue() call here was the same
      // hang-prone pattern fixed for generateAlertBlocks in ce54c1f.
      if (lastRow > 1) {
        var colAValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        var colBValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
        for (var k = 0; k < colBValues.length; k++) {
          var id = colBValues[k][0].toString().trim();
          if (id !== "" && !isNaN(colAValues[k][0])) {
            existingIds.push(id);
            idToRow[id] = k + 2;
            currentStudentCount++;
          } else break;
        }
      }

      // Add only active students. Inactive students already in the workbook are
      // retained so their historical attendance remains available.
      var newEntries = csvRoster.filter(function(row) {
        return row[3] === "active" && existingIds.indexOf(row[1].toString().trim()) === -1;
      }).map(function(row) {
        return [row[0], row[1], row[2]];
      });

      var daysInMonth = new Date(monthInfo.year, monthInfo.monthIndex + 1, 0).getDate();
      var totalStudentsNow = currentStudentCount;

      // 1. ADD NEW STUDENTS (If any)
      if (newEntries.length > 0) {
        Logger.log("    [+] Found " + newEntries.length + " new student(s) for " + monthInfo.name);
        var insertRowIndex = currentStudentCount + 1;
        sheet.insertRowsAfter(insertRowIndex, newEntries.length);

        for (var e = 0; e < newEntries.length; e++) newEntries[e][0] = (currentStudentCount + e + 1).toString();

        var startRow = insertRowIndex + 1;
        sheet.getRange(startRow, 1, newEntries.length, 3).setValues(newEntries);
        totalStudentsNow = currentStudentCount + newEntries.length;

        if (isPastMonth) {
          var pastBlock = sheet.getRange(startRow, 4, newEntries.length, daysInMonth);
          pastBlock.clearDataValidations();
          fillWithPlaceholders(pastBlock, newEntries.length, daysInMonth);
        } else if (isCurrentMonth && todayDay > 1) {
          var currentMonthPastBlock = sheet.getRange(startRow, 4, newEntries.length, todayDay - 1);
          currentMonthPastBlock.clearDataValidations();
          fillWithPlaceholders(currentMonthPastBlock, newEntries.length, todayDay - 1);
        }
      }

      // 2. ALWAYS REFRESH HOLIDAYS & FORMULAS (Even if no new students were added)
      if (totalStudentsNow > 0) {
        refreshFormulasAndStyles(sheet, totalStudentsNow, daysInMonth, monthInfo.year, monthInfo.monthIndex, monthInfo.name, holidays);
      }

      // 3. LOCK OUT DROPPED-OUT STUDENTS. Runs AFTER refreshFormulasAndStyles,
      // which repaints the whole day grid (including this row) — so the grey
      // lock must be applied last or it gets overwritten. Only the affected
      // student's own row is touched; every other row is left exactly as-is.
      var totalCols = 9 + daysInMonth;
      existingIds.forEach(function(id) {
        var rowNum = idToRow[id];
        if (statusById[id] === "inactive") {
          lockInactiveStudentRow(sheet, rowNum, totalCols, daysInMonth, id, isPastMonth, isCurrentMonth, todayDay);
        } else if (statusById[id] === "active") {
          unlockStudentRow(sheet, rowNum, totalCols, daysInMonth, id, isPastMonth, isCurrentMonth, todayDay);
        }
      });
    }
  }
  
  props.deleteProperty('FLOW2_TOKEN');
  Logger.log("SUCCESS: DELTA SYNCHRONIZATION COMPLETE.\n");
}

//Execution Tokens are meant to safely batch process (do some function on) the reports without hitting Google's 6min Execution Timeout.
// Reset the Execution Token After the full setup is done / To run everything from Scratch
function manual_resetExecutionTokens() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('FLOW1_TOKEN');
  props.deleteProperty('FLOW2_TOKEN');
  props.deleteProperty('FLOW3_TOKEN');
  Logger.log("✅ All execution continuation tokens cleared. Next run will start from scratch.");
}


// =========================================================================
// MANUAL (admin run): FULL SYSTEM RESET — clears ALL runtime state.
// Use this to wipe every Script Property the daily loop writes so the
// system starts completely fresh (see RESET.md for the full procedure).
//
// This does NOT touch Drive files (workbooks/forms) or triggers — those are
// deleted separately per RESET.md. It ONLY clears the Script Properties:
//   ACTIVE_FORM_<ssId>        – which form is live for each workbook
//   AUTHORIZED_TEACHER_<ssId> – who may submit for each workbook (this key
//                               is otherwise never deleted → leaks over time)
//   NOTIFIED_<ssId>           – per-day "already emailed" dedup set
//   FORM_TARGET_DATE_<formId> – target date for on-demand forms
//   FLOW1_TOKEN / FLOW2_TOKEN / FLOW3_TOKEN – setup batch continuation tokens
// =========================================================================
function manual_resetAllRuntimeState() {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  var deleted = 0;

  for (var key in all) {
    if (key.indexOf('ACTIVE_FORM_') === 0 ||
        key.indexOf('AUTHORIZED_TEACHER_') === 0 ||
        key.indexOf('NOTIFIED_') === 0 ||
        key.indexOf('FORM_TARGET_DATE_') === 0 ||
        key === 'FLOW1_TOKEN' ||
        key === 'FLOW2_TOKEN' ||
        key === 'FLOW3_TOKEN') {
      props.deleteProperty(key);
      deleted++;
      Logger.log("🗑️ Cleared: " + key);
    }
  }

  Logger.log("✅ Full runtime reset complete. " + deleted + " Script Properties cleared.");
  Logger.log("ℹ️ Next: delete Drive workbooks/forms + triggers, then re-run manual_installTriggers (see RESET.md).");
}


function manual_validateSetupConfig() {
  Logger.log("=========================================================");
  Logger.log("🔍 SETUP VALIDATION - Checking configuration");
  Logger.log("=========================================================\n");

  var errors = [];
  var warnings = [];

  // Check folders
  try {
    var inputFolder = getFolderByLink(STUDENT_CLASS_LIST_FOLDER_LINK);
    Logger.log("✅ Input folder: " + inputFolder.getName());

    var fileCount = 0;
    var files = inputFolder.getFiles();
    while (files.hasNext()) {
      files.next();
      fileCount++;
    }
    Logger.log("   📁 Contains " + fileCount + " file(s)");

    if (fileCount === 0) {
      warnings.push("Input folder is empty. Add roster files before running setup.");
    }
  } catch(e) {
    errors.push("Cannot access STUDENT_CLASS_LIST_FOLDER_LINK: " + e.message);
  }

  try {
    var outputFolder = getFolderByLink(ATTENDANCE_SHEETS_FOLDER_LINK);
    Logger.log("✅ Output folder: " + outputFolder.getName());

    var sheetCount = 0;
    var sheets = outputFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (sheets.hasNext()) {
      sheets.next();
      sheetCount++;
    }
    Logger.log("   📊 Contains " + sheetCount + " attendance sheet(s)");
  } catch(e) {
    errors.push("Cannot access ATTENDANCE_SHEETS_FOLDER_LINK: " + e.message);
  }

  try {
    var configFolder = getFolderByLink(CONFIG_FOLDER_LINK);
    Logger.log("✅ Config folder: " + configFolder.getName());
  } catch(e) {
    errors.push("Cannot access CONFIG_FOLDER_LINK: " + e.message);
  }

  // Check academic year settings
  Logger.log("\n📅 Academic Settings:");
  Logger.log("   Year: " + ACADEMIC_YEAR);
  Logger.log("   Start: Month " + START_MONTH);
  Logger.log("   End: Month " + END_MONTH);

  var months = getAcademicMonthsList(ACADEMIC_YEAR, START_MONTH, END_MONTH);
  Logger.log("   Total months: " + months.length);

  if (months.length < 1 || months.length > 12) {
    errors.push("Invalid academic year configuration. Check START_MONTH and END_MONTH in Config.gs");
  }

  // Check visualizations
  Logger.log("\n📊 Visualization Settings:");
  Logger.log("   ADD_VISUALISATIONS: " + ADD_VISUALISATIONS);

  // Check execution timeout
  Logger.log("\n⏱️ Execution Settings:");
  Logger.log("   MAX_EXECUTION_TIME: " + (MAX_EXECUTION_TIME / 1000 / 60) + " minutes");

  if (MAX_EXECUTION_TIME > 360000) {
    warnings.push("MAX_EXECUTION_TIME exceeds Google's 6-minute limit. Set to max 360000 (6 minutes)");
  }

  // Check for existing tokens
  var props = PropertiesService.getScriptProperties();
  var hasTokens = false;

  Logger.log("\n🔄 Batch Progress:");
  if (props.getProperty('FLOW1_TOKEN')) {
    Logger.log("   ⚠️ Sheet generation in progress (resume available)");
    hasTokens = true;
  }
  if (props.getProperty('FLOW2_TOKEN')) {
    Logger.log("   ⚠️ Student updates in progress (resume available)");
    hasTokens = true;
  }
  if (props.getProperty('FLOW3_TOKEN')) {
    Logger.log("   ⚠️ Chart regeneration in progress (resume available)");
    hasTokens = true;
  }
  if (!hasTokens) {
    Logger.log("   ✅ No batch operations in progress");
  }

  // Summary
  Logger.log("\n=========================================================");
  if (errors.length > 0) {
    Logger.log("❌ VALIDATION FAILED!");
    Logger.log("\nErrors:");
    for (var i = 0; i < errors.length; i++) {
      Logger.log("   " + (i + 1) + ". " + errors[i]);
    }
  } else {
    Logger.log("✅ VALIDATION PASSED!");
  }

  if (warnings.length > 0) {
    Logger.log("\n⚠️ Warnings:");
    for (var i = 0; i < warnings.length; i++) {
      Logger.log("   " + (i + 1) + ". " + warnings[i]);
    }
  }
  Logger.log("=========================================================\n");

  if (errors.length > 0) {
    throw new Error("Configuration validation failed. See logs above.");
  }
}
