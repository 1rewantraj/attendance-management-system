// =========================================================================
// FILE 2: Attendence_Sheets_Setup.gs (Master Generation & Updates)
// =========================================================================

function manual_generateSheets() {
  var startTime = Date.now();
  Logger.log("=========================================================");
  Logger.log("STEP 1: STARTING MASTER BATCH INITIAL GENERATION");
  Logger.log("=========================================================");
  
  try { DriveApp.getRootFolder(); } catch(e) {} 
  
  var inputFolder = getFolderByLink(INPUT_FOLDER_LINK);
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
    Logger.log(fileName);
    var info = parseClassAndSectionFromText(fileName);
    var targetSpreadsheetName = "Class_" + info.classNum + "_" + info.section + "_" + ACADEMIC_YEAR;
    
    if (outputFolder.getFilesByName(targetSpreadsheetName).hasNext()) {
      Logger.log("[SKIP] " + targetSpreadsheetName + " already exists.");
      continue;
    }
    
    Logger.log("--> Processing roster: " + fileName);
    var rosterData = parseAndNormalizeData(file);
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
  
  var inputFolder = getFolderByLink(INPUT_FOLDER_LINK);
  var outputFolder = getFolderByLink(ATTENDANCE_SHEETS_FOLDER_LINK);
  var configFolder = getFolderByLink(CONFIG_FOLDER_LINK);
  
  var monthsToCreate = getAcademicMonthsList(ACADEMIC_YEAR, START_MONTH, END_MONTH);
  var holidays = getPublicHolidays(configFolder);
  
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
    var targetSpreadsheetName = "Class_" + info.classNum + "_" + info.section + "_" + ACADEMIC_YEAR;
    var ssFiles = outputFolder.getFilesByName(targetSpreadsheetName);
    
    if (!ssFiles.hasNext()) continue;
    
    var ss = SpreadsheetApp.open(ssFiles.next());
    var csvRoster = parseAndNormalizeData(file);
    if (csvRoster.length === 0) continue;
    
    Logger.log("--> Checking updates for: " + targetSpreadsheetName);
    
    for (var m = 0; m < monthsToCreate.length; m++) {
      var monthInfo = monthsToCreate[m];
      var sheet = ss.getSheetByName(monthInfo.name);
      if (!sheet) continue;
      
      var isPastMonth = (monthInfo.year < todayYear) || (monthInfo.year === todayYear && monthInfo.monthIndex < todayMonth);
      var isCurrentMonth = (monthInfo.year === todayYear && monthInfo.monthIndex === todayMonth);
      var lastRow = sheet.getLastRow();
      var existingIds = [];
      var currentStudentCount = 0;
      
      if (lastRow > 1) {
        var colBValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
        for (var k = 0; k < colBValues.length; k++) {
          var id = colBValues[k][0].toString().trim();
          if (id !== "" && !isNaN(sheet.getRange(k + 2, 1).getValue())) {
            existingIds.push(id);
            currentStudentCount++;
          } else break;
        }
      }
      
      var newEntries = csvRoster.filter(function(row) {
        return existingIds.indexOf(row[1].toString().trim()) === -1;
      });
      
      for (var m = 0; m < monthsToCreate.length; m++) {
      var monthInfo = monthsToCreate[m];
      var sheet = ss.getSheetByName(monthInfo.name);
      if (!sheet) continue;
      
      var isPastMonth = (monthInfo.year < todayYear) || (monthInfo.year === todayYear && monthInfo.monthIndex < todayMonth);
      var isCurrentMonth = (monthInfo.year === todayYear && monthInfo.monthIndex === todayMonth);
      var lastRow = sheet.getLastRow();
      var existingIds = [];
      var currentStudentCount = 0;
      
      if (lastRow > 1) {
        var colBValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
        for (var k = 0; k < colBValues.length; k++) {
          var id = colBValues[k][0].toString().trim();
          if (id !== "" && !isNaN(sheet.getRange(k + 2, 1).getValue())) {
            existingIds.push(id);
            currentStudentCount++;
          } else break;
        }
      }
      
      var newEntries = csvRoster.filter(function(row) {
        return existingIds.indexOf(row[1].toString().trim()) === -1;
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
    }
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
  //props.deleteProperty('FLOW3_TOKEN');
  Logger.log("✅ All execution continuation tokens cleared. Next run will start from scratch.");
}


function manual_validateSetupConfig() {
  Logger.log("=========================================================");
  Logger.log("🔍 SETUP VALIDATION - Checking configuration");
  Logger.log("=========================================================\n");

  var errors = [];
  var warnings = [];

  // Check folders
  try {
    var inputFolder = getFolderByLink(INPUT_FOLDER_LINK);
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
    errors.push("Cannot access INPUT_FOLDER_LINK: " + e.message);
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
