// =========================================================================
// Utils.gs - ALL UTILITY & SYSTEM FUNCTIONS
// =========================================================================

function getFolderByLink(url) {
  if (!url || url === "YOUR_CONFIG_FOLDER_LINK_HERE" || url === "YOUR_CLASS_ROSTER_FOLDER_LINK_HERE" || url === "YOUR_ATTENDANCE_SHEETS_FOLDER_LINK_HERE") {
    throw new Error("❌ Folder link not configured. Please set the folder links in Config.gs and run manual_validateConfig()");
  }

  var id = url.match(/[-\w]{25,}/) ? url.match(/[-\w]{25,}/)[0] : url;

  if (!id || id.length < 25) {
    throw new Error("❌ Invalid folder link or ID: " + url);
  }

  try {
    return DriveApp.getFolderById(id);
  } catch(e) {
    throw new Error("❌ Cannot access folder (ID: " + id + "). Check permissions or verify the link is correct. Original error: " + e.message);
  }
}

// =========================================================================
// COMPREHENSIVE SEARCH & UTILITIES
// =========================================================================
function comprehensiveFileSearch(keyword, folder) {
  var matchingFiles = [];
  var cleanKeyword = keyword.toLowerCase().replace(/[\s_]+/g, "");
  // Guard: scanning the whole Drive (DriveApp.getFiles()) can iterate tens of
  // thousands of files and blow the 6-min limit. Require a folder scope.
  if (!folder) {
    Logger.log("    [WARNING] comprehensiveFileSearch called without a folder for '" +
               keyword + "'. Skipping whole-Drive scan to avoid timeout.");
    return matchingFiles;
  }
  var filesIterator = folder.getFiles();

  while (filesIterator.hasNext()) {
    var file = filesIterator.next();
    var cleanFileName = file.getName().toLowerCase().replace(/[\s_]+/g, "");

    if (cleanFileName.indexOf(cleanKeyword) > -1) {
      matchingFiles.push(file);
    }
  }

  matchingFiles.sort(function(a, b) {
    return getFilePriority(a.getMimeType()) - getFilePriority(b.getMimeType());
  });

  return matchingFiles;
}

function getFilePriority(mimeType) {
  if (mimeType === MimeType.MICROSOFT_EXCEL || mimeType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") return 1;
  if (mimeType === MimeType.GOOGLE_SHEETS) return 2;
  if (mimeType === MimeType.CSV) return 3;
  return 4;
}

function getExcelDataFromFolder(folder, fileNameMatch) {
  var matches = comprehensiveFileSearch(fileNameMatch, folder);
  if (matches.length > 0) {
    var bestFile = matches[0];
    Logger.log("    [MATCH FOUND] Keyword '" + fileNameMatch + "' matched to: " + bestFile.getName());
    return parseAndNormalizeData(bestFile, true);
  }
  return null;
}

function parseAndNormalizeData(file, returnRaw) {
  var mimeType = file.getMimeType();
  var rawRows = [];

  if (mimeType === MimeType.CSV) {
    var fileString = file.getBlob().getDataAsString();
    rawRows = Utilities.parseCsv(fileString);
  } else if (mimeType === MimeType.MICROSOFT_EXCEL || mimeType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
    var tempFile = Drive.Files.copy(
      {title: "Temp_" + file.getName(), mimeType: MimeType.GOOGLE_SHEETS},
      file.getId()
    );
    var tempSs = SpreadsheetApp.openById(tempFile.id);
    rawRows = tempSs.getSheets()[0].getDataRange().getValues();
    DriveApp.getFileById(tempFile.id).setTrashed(true);
  } else if (mimeType === MimeType.GOOGLE_SHEETS) {
    var tempSs = SpreadsheetApp.openById(file.getId());
    rawRows = tempSs.getSheets()[0].getDataRange().getValues();
  }

  if (rawRows.length === 0) return [];
  if (returnRaw) return rawRows;

  var hasHeader = false, rollNoIdx = -1, childIdIdx = -1, nameIdx = -1;
  var firstRowStr = rawRows[0].join(" ").toLowerCase();

  if (firstRowStr.includes("child") || firstRowStr.includes("name") || firstRowStr.includes("roll")) {
    hasHeader = true;
    for (var c = 0; c < rawRows[0].length; c++) {
      var cellLower = rawRows[0][c].toString().trim().toLowerCase();
      if (cellLower.includes("roll")) rollNoIdx = c;
      else if (cellLower.includes("child")) childIdIdx = c;
      else if (cellLower.includes("student") || cellLower.includes("name")) nameIdx = c;
    }
    rawRows.shift();
  }

  var normalizedData = [];
  for (var i = 0; i < rawRows.length; i++) {
    var row = rawRows[i];
    if (row.length === 0 || (row.length === 1 && row[0].toString().trim() === "")) continue;
    var childId = hasHeader && childIdIdx !== -1 ? row[childIdIdx].toString().trim() : row[0].toString().trim();
    var studentName = hasHeader && nameIdx !== -1 ? row[nameIdx].toString().trim() : row[1].toString().trim();
    var rollNo = hasHeader && rollNoIdx !== -1 && row[rollNoIdx].toString().trim() ? row[rollNoIdx].toString().trim() : (i + 1).toString();
    normalizedData.push([rollNo, childId, studentName]);
  }
  return normalizedData;
}

function parseClassAndSectionFromText(text) {
  var clean = text.replace(/\.(csv|xlsx)$/i, "").replace(/_/g, " ");
  var classMatch = clean.match(/(?:class|grade)\s+([a-zA-Z0-9]+)/i);
  var secMatch = clean.replace(/(?:class|grade)\s+[a-zA-Z0-9]+/i, "").match(/\b[a-zA-Z]\b/);
  return { classNum: classMatch ? classMatch[1] : "1", section: secMatch ? secMatch[0].toUpperCase() : "A" };
}

// SINGLE SOURCE OF TRUTH for a class workbook's file name. Both the Setup flow
// (manual_generateSheets / manual_updateSheets) and the Daily flow
// (automated_sendDailyForms / manual_sendOnDemandForm) must resolve the SAME file,
// so they all build the name here. Keyed by class, section, and academic year.
function getWorkbookName(classNum, section) {
  return "Class_" + classNum + "_" + section.toString().toUpperCase() + "_" + ACADEMIC_YEAR;
}

function getAcademicMonthsList(academicYear, startMonth, endMonth) {
  var years = academicYear.split("-");
  var startYear = parseInt(years[0].trim(), 10), endYear = parseInt(years[1].trim(), 10);
  if (startYear < 100) startYear += 2000;
  if (endYear < 100) endYear += 2000;

  var list = [];
  var curr = new Date(startYear, startMonth - 1, 1);
  var end = new Date(endYear, endMonth - 1, 1);

  while (curr <= end) {
    list.push({ name: curr.toLocaleString('en-US', { month: 'long' }), monthIndex: curr.getMonth(), year: curr.getFullYear() });
    curr.setMonth(curr.getMonth() + 1);
  }
  return list;
}

function getColLetter(col) {
  var letter = "";
  while (col > 0) {
    var temp = (col - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}

/**
 * Parse a date string in DD-MM-YYYY format (e.g., "05-07-2026" = July 5, 2026).
 * If the input is empty or doesn't match DD-MM-YYYY, returns null (caller should use default).
 * If format matches but date is invalid, throws an error.
 */
function parseDDMMYYYY(dateString) {
  if (!dateString || dateString.trim() === "") {
    return null;
  }

  var trimmed = dateString.trim();

  // Check if format matches DD-MM-YYYY (e.g., "05-07-2026" or "30-12-2026")
  var match = trimmed.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);

  if (!match) {
    // Not DD-MM-YYYY format, return null so caller can try other formats
    return null;
  }

  var day = parseInt(match[1], 10);
  var month = parseInt(match[2], 10);
  var year = parseInt(match[3], 10);

  // Validate ranges
  if (month < 1 || month > 12) {
    throw new Error("Invalid month in date '" + dateString + "'. Month must be 01-12.");
  }

  if (day < 1 || day > 31) {
    throw new Error("Invalid day in date '" + dateString + "'. Day must be 01-31.");
  }

  // Create date (months are 0-indexed in JavaScript)
  var date = new Date(year, month - 1, day);

  // Verify the date is valid (handles cases like Feb 30)
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    throw new Error("Invalid date '" + dateString + "'. This date does not exist (e.g., Feb 30).");
  }

  return date;
}

// =========================================================================
// CONFIGURATION & PERMISSIONS LOGIC
// =========================================================================
function buildMasterConfig(configFolder) {
  var masterConfig = { roles: {}, stakeholders: { editors: [], viewers: [] }, classMap: {} };

  var ssScopes = null;
  // Search ONLY inside the config folder. Passing null here would make
  // comprehensiveFileSearch fall back to DriveApp.getFiles() and scan the
  // ENTIRE Drive — on a large Drive that runs until the 6-min limit and hangs.
  var scopeMatches = comprehensiveFileSearch("default sharing scopes", configFolder);
  if (scopeMatches.length > 0) {
    var bestScopeFile = scopeMatches[0];
    Logger.log("    [MATCH FOUND] Scopes config matched to: " + bestScopeFile.getName());
    ssScopes = SpreadsheetApp.openById(bestScopeFile.getId());
  } else {
    Logger.log("    [WARNING] 'Default Sharing Scopes' file not found. Using fallbacks.");
  }

  if (ssScopes) {
    var rolesSheet = ssScopes.getSheetByName("Roles") || ssScopes.getSheetByName("roles");
    if (rolesSheet) {
      var rData = rolesSheet.getDataRange().getValues();
      for (var i = 1; i < rData.length; i++) {
        if (rData[i][0]) {
          masterConfig.roles[rData[i][0].toString().trim().toLowerCase()] = rData[i][1].toString().trim().toLowerCase();
        }
      }
    }

    var shSheet = ssScopes.getSheetByName("Stakeholders") || ssScopes.getSheetByName("stakeholders");
    if (shSheet) {
      var sData = shSheet.getDataRange().getValues();
      if (sData.length > 1) {
        var header = sData[0].map(function(h) { return h.toString().toLowerCase().trim(); });
        var emailIdx = header.indexOf("email id") > -1 ? header.indexOf("email id") : header.indexOf("email");
        var scopeIdx = header.indexOf("scope");

        if (emailIdx > -1 && scopeIdx > -1) {
          for (var j = 1; j < sData.length; j++) {
            var rawEmailCell = sData[j][emailIdx] ? sData[j][emailIdx].toString() : "";
            var scope = sData[j][scopeIdx] ? sData[j][scopeIdx].toString().trim().toLowerCase() : "";
            
            // Clean and split comma-separated emails into individual trimmed emails
            var cleanEmails = parseCleanList(rawEmailCell);
            cleanEmails.forEach(function(email) {
              if (scope === "edit" || scope === "editor") {
                masterConfig.stakeholders.editors.push(email);
              } else {
                masterConfig.stakeholders.viewers.push(email);
              }
            });
          }
        }
      }
    }
  }

  var managersData = getExcelDataFromFolder(configFolder, "programManagers");
  var leadsData = getExcelDataFromFolder(configFolder, "teacherLeads");
  var teachersData = getExcelDataFromFolder(configFolder, "teacherclassmapping");

  // 1. PROCESS PROGRAM MANAGERS
  var managersMap = {};
  if (managersData && managersData.length > 1) {
    var mgrHeader = managersData[0];
    var mgrIdIdx = findColIdx(mgrHeader, ["id", "manager id", "manager_id", "pm id, manager_id"], 0);
    var mgrEmailIdx = findColIdx(mgrHeader, ["email", "email id", "email_id", "manager email", "manageremail"], 3);

    for (var i = 1; i < managersData.length; i++) {
      var row = managersData[i];
      if (row[mgrIdIdx]) {
        var managerId = row[mgrIdIdx].toString().trim();
        var managerEmail = parseCleanList(row[mgrEmailIdx]).join(",");
        managersMap[managerId] = managerEmail;
      }
    }
  }

  // 2. PROCESS TEACHER LEADS
  var leadsMap = {};
  if (leadsData && leadsData.length > 1) {
    var leadHeader = leadsData[0];
    var leadIdIdx = findColIdx(leadHeader, ["id", "lead id", "lead_id", "tl id", "leadid"], 0);
    var leadEmailIdx = findColIdx(leadHeader, ["email", "email id", "email_id", "lead email", "leademail"], 2);
    var leadMgrIdIdx = findColIdx(leadHeader, ["manager id", "manager_id", "manager", "pm id", "managerid"], 3);

    for (var i = 1; i < leadsData.length; i++) {
      var row = leadsData[i];
      if (row[leadIdIdx]) {
        var leadId = row[leadIdIdx].toString().trim();
        var leadEmail = parseCleanList(row[leadEmailIdx]).join(",");
        var managerIdRef = row[leadMgrIdIdx] ? row[leadMgrIdIdx].toString().trim() : "";
        leadsMap[leadId] = { email: leadEmail, managerId: managerIdRef };
      }
    }
  }

  if (teachersData && teachersData.length > 1) {
    var tkrHeader = teachersData[0];
    var tkrIdIdx = findColIdx(tkrHeader, ["id", "teacher id", "teacher_id", "teacherid"], 0);
    var tkrEmailIdx = findColIdx(tkrHeader, ["email", "email id", "email_id", "teacher email", "teacheremail"], 2);
    var tkrLeadIdx = findColIdx(tkrHeader, ["lead id", "lead_id", "tl id", "lead", "leadid"], 3);
    var tkrClassIdx = findColIdx(tkrHeader, ["class", "class num", "class_num", "grade"], 4);
    var tkrSecIdx = findColIdx(tkrHeader, ["section", "sec", "division"], 5);

    for (var i = 1; i < teachersData.length; i++) {
      var row = teachersData[i];
      if (row[tkrIdIdx] || row[tkrClassIdx]) {
        var teacherEmail = parseCleanList(row[tkrEmailIdx]).join(",");
        var leadId = row[tkrLeadIdx] ? row[tkrLeadIdx].toString().trim() : "";
        var classNum = row[tkrClassIdx] ? row[tkrClassIdx].toString().trim() : "";
        var section = row[tkrSecIdx] ? row[tkrSecIdx].toString().trim().toUpperCase() : "";

        var leadEmail = leadsMap[leadId] ? leadsMap[leadId].email : null;
        var managerId = leadsMap[leadId] ? leadsMap[leadId].managerId : null;
        var managerEmail = managerId && managersMap[managerId] ? managersMap[managerId] : null;

        if (classNum) {
          masterConfig.classMap[classNum + "_" + section] = { 
            teacher: teacherEmail, 
            lead: leadEmail, 
            manager: managerEmail 
          };
        }
      }
    }
  }
  return masterConfig;
}

/**
 * Helper to dynamically locate a column index by checking header row variations.
 */
function findColIdx(headerRow, candidates, defaultIdx) {
  if (!headerRow || !headerRow.length) return defaultIdx;
  var normalized = headerRow.map(function(h) { 
    return h ? h.toString().toLowerCase().trim() : ""; 
  });
  
  for (var i = 0; i < candidates.length; i++) {
    var idx = normalized.indexOf(candidates[i]);
    if (idx > -1) return idx;
  }
  return defaultIdx;
}

/**
 * Safely converts a cell value or comma-separated string into a clean, trimmed array of items.
 */
function parseCleanList(input) {
  if (!input) return [];
  
  // Convert to string and split by comma
  return String(input)
    .split(',')
    .map(function(item) { 
      return item.trim(); // Removes leading & trailing whitespace/spaces
    })
    .filter(function(item) { 
      return item.length > 0; // Excludes empty strings
    });
}

function applyPermissionsToSpreadsheet(ssFile, classNum, section, masterConfig) {
  var key = classNum + "_" + section;
  var classUsers = masterConfig.classMap[key];

  // Helper to split, clean, and validate email lists
  function cleanAndFlattenEmails(inputList) {
    var cleanList = [];
    if (!inputList) return cleanList;
    
    var rawList = Array.isArray(inputList) ? inputList : [inputList];
    
    rawList.forEach(function(item) {
      if (!item) return;
      // Split by comma in case an item is a joined string like "email1@x.com,email2@x.com"
      String(item).split(',').forEach(function(email) {
        var trimmed = email.trim();
        // Check for valid email pattern
        if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(trimmed)) {
          if (cleanList.indexOf(trimmed) === -1) {
            cleanList.push(trimmed);
          }
        } else if (trimmed.length > 0) {
          Logger.log("      [WARNING] Skipping invalid email format: " + trimmed);
        }
      });
    });
    return cleanList;
  }

  var rawEditors = [].concat(masterConfig.stakeholders.editors || []);
  var rawViewers = [].concat(masterConfig.stakeholders.viewers || []);

  if (classUsers) {
    function assignScope(email, roleName) {
      if (!email) return;
      var scope = masterConfig.roles[roleName.toLowerCase()] || "view";
      if (scope === "edit" || scope === "editor") {
        rawEditors.push(email);
      } else {
        rawViewers.push(email);
      }
    }
    assignScope(classUsers.teacher, "Teacher");
    assignScope(classUsers.lead, "Teacher Lead");
    assignScope(classUsers.manager, "Program Manager");
  }

  var editorsToGrant = cleanAndFlattenEmails(rawEditors);
  var viewersToGrant = cleanAndFlattenEmails(rawViewers);

  // Prevent users from being added as viewers if they already have edit access
  viewersToGrant = viewersToGrant.filter(function(email) {
    return editorsToGrant.indexOf(email) === -1;
  });

  // Grant Edit permissions safely
  if (editorsToGrant.length > 0) {
    Logger.log("      [GRANTING EDIT] -> " + editorsToGrant.join(", "));
    editorsToGrant.forEach(function(email) {
      try {
        ssFile.addEditor(email);
      } catch (e) {
        Logger.log("      [ERROR] Could not add Editor (" + email + "): " + e.message);
      }
    });
  }

  // Grant View permissions safely
  if (viewersToGrant.length > 0) {
    Logger.log("      [GRANTING VIEW] -> " + viewersToGrant.join(", "));
    viewersToGrant.forEach(function(email) {
      try {
        ssFile.addViewer(email);
      } catch (e) {
        Logger.log("      [ERROR] Could not add Viewer (" + email + "): " + e.message);
      }
    });
  }
}

/**
 * Updates permissions for an existing spreadsheet by checking current access
 * against the master config and adding any new emails that should have access.
 * Does NOT revoke access from removed users (additive only).
 */
function updatePermissionsIfNeeded(ssFile, classNum, section, masterConfig) {
  var key = classNum + "_" + section;
  var classUsers = masterConfig.classMap[key];

  // Helper to split, clean, and validate email lists
  function cleanAndFlattenEmails(inputList) {
    var cleanList = [];
    if (!inputList) return cleanList;

    var rawList = Array.isArray(inputList) ? inputList : [inputList];

    rawList.forEach(function(item) {
      if (!item) return;
      String(item).split(',').forEach(function(email) {
        var trimmed = email.trim();
        if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(trimmed)) {
          if (cleanList.indexOf(trimmed) === -1) {
            cleanList.push(trimmed);
          }
        } else if (trimmed.length > 0) {
          Logger.log("      [WARNING] Skipping invalid email format: " + trimmed);
        }
      });
    });
    return cleanList;
  }

  // Build list of emails that should have access
  var rawEditors = [].concat(masterConfig.stakeholders.editors || []);
  var rawViewers = [].concat(masterConfig.stakeholders.viewers || []);

  if (classUsers) {
    function assignScope(email, roleName) {
      if (!email) return;
      var scope = masterConfig.roles[roleName.toLowerCase()] || "view";
      if (scope === "edit" || scope === "editor") {
        rawEditors.push(email);
      } else {
        rawViewers.push(email);
      }
    }
    assignScope(classUsers.teacher, "Teacher");
    assignScope(classUsers.lead, "Teacher Lead");
    assignScope(classUsers.manager, "Program Manager");
  }

  var editorsToGrant = cleanAndFlattenEmails(rawEditors);
  var viewersToGrant = cleanAndFlattenEmails(rawViewers);

  // Prevent users from being added as viewers if they should have edit access
  viewersToGrant = viewersToGrant.filter(function(email) {
    return editorsToGrant.indexOf(email) === -1;
  });

  // Get file owner (cannot add owner as editor/viewer)
  var ownerEmail = "";
  try {
    ownerEmail = ssFile.getOwner().getEmail().toLowerCase();
  } catch (e) {
    Logger.log("      [WARNING] Could not fetch file owner: " + e.message);
  }

  // Get current permissions
  var currentEditors = [];
  var currentViewers = [];

  try {
    var editors = ssFile.getEditors();
    editors.forEach(function(editor) {
      currentEditors.push(editor.getEmail().toLowerCase());
    });
  } catch (e) {
    Logger.log("      [WARNING] Could not fetch current editors: " + e.message);
  }

  try {
    var viewers = ssFile.getViewers();
    viewers.forEach(function(viewer) {
      currentViewers.push(viewer.getEmail().toLowerCase());
    });
  } catch (e) {
    Logger.log("      [WARNING] Could not fetch current viewers: " + e.message);
  }

  // Find new editors (not currently editors, and not the owner)
  var newEditors = editorsToGrant.filter(function(email) {
    var emailLower = email.toLowerCase();
    return emailLower !== ownerEmail &&
           currentEditors.indexOf(emailLower) === -1;
  });

  // Find new viewers (not currently viewers or editors, and not the owner)
  var newViewers = viewersToGrant.filter(function(email) {
    var emailLower = email.toLowerCase();
    return emailLower !== ownerEmail &&
           currentViewers.indexOf(emailLower) === -1 &&
           currentEditors.indexOf(emailLower) === -1;
  });

  // Grant new Edit permissions
  if (newEditors.length > 0) {
    Logger.log("      [ADDING EDIT ACCESS] -> " + newEditors.join(", "));
    newEditors.forEach(function(email) {
      try {
        ssFile.addEditor(email);
      } catch (e) {
        Logger.log("      [ERROR] Could not add Editor (" + email + "): " + e.message);
      }
    });
  }

  // Grant new View permissions
  if (newViewers.length > 0) {
    Logger.log("      [ADDING VIEW ACCESS] -> " + newViewers.join(", "));
    newViewers.forEach(function(email) {
      try {
        ssFile.addViewer(email);
      } catch (e) {
        Logger.log("      [ERROR] Could not add Viewer (" + email + "): " + e.message);
      }
    });
  }

  // Log if no changes needed
  if (newEditors.length === 0 && newViewers.length === 0) {
    Logger.log("      [PERMISSIONS UP TO DATE] No new access to grant.");
  }
}

function getPublicHolidays(configFolder) {
  var holidays = [];
  var data = getExcelDataFromFolder(configFolder, HOLIDAY_FILE_NAME);
  if (!data || data.length < 2) return holidays;

  var header = data[0].map(function(h) { return h.toString().toLowerCase().trim(); });
  
  // Column matching
  var dateIdx = header.indexOf("date");
  if (dateIdx === -1) dateIdx = header.indexOf("dates");
  
  var startIdx = header.indexOf("start date");
  if (startIdx === -1) startIdx = header.indexOf("from");
  
  var endIdx = header.indexOf("end date");
  if (endIdx === -1) endIdx = header.indexOf("to");

  var timeZone = Session.getScriptTimeZone();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var startDate = null;
    var endDate = null;

    // 1. Check for separate Start Date and End Date columns
    if (startIdx > -1 && row[startIdx]) {
      startDate = parseDateValue(row[startIdx]);
      endDate = (endIdx > -1 && row[endIdx]) ? parseDateValue(row[endIdx]) : startDate;
    } 
    // 2. Check single Date column (handles both single dates & range strings like "2026-10-10 - 2026-10-15")
    else if (dateIdx > -1 && row[dateIdx]) {
      var dVal = row[dateIdx];

      if (typeof dVal === "string" && (dVal.includes(" - ") || dVal.toLowerCase().includes(" to "))) {
        var parts = dVal.split(/\s+-\s+|\s+to\s+/i);
        if (parts.length === 2) {
          startDate = parseDateValue(parts[0]);
          endDate = parseDateValue(parts[1]);
        }
      }

      if (!startDate) {
        startDate = parseDateValue(dVal);
        endDate = startDate;
      }
    }

    // 3. Loop through every day in the date range and collect YYYY-MM-DD
    if (startDate && endDate && !isNaN(startDate.getTime()) && !isNaN(endDate.getTime())) {
      var curr = new Date(startDate.getTime());
      while (curr <= endDate) {
        var formatted = Utilities.formatDate(curr, timeZone, "yyyy-MM-dd");
        if (holidays.indexOf(formatted) === -1) {
          holidays.push(formatted);
        }
        curr.setDate(curr.getDate() + 1);
      }
    }
  }

  return holidays;
}

// Helper to safely convert cell inputs to JavaScript Date objects
function parseDateValue(val) {
  if (!val) return null;
  if (val instanceof Date) return val;

  var str = val.toString().trim();

  // 1. Check for DD/MM/YYYY or DD-MM-YYYY (e.g. 19/10/2026)
  var ddmmyyyyPattern = /^(\d{1,2})[\/\.-](\d{1,2})[\/\.-](\d{4})$/;
  var match = str.match(ddmmyyyyPattern);

  if (match) {
    var day = parseInt(match[1], 10);
    var month = parseInt(match[2], 10) - 1; // JavaScript months are 0-indexed (0 = Jan, 9 = Oct)
    var year = parseInt(match[3], 10);
    var parsedDate = new Date(year, month, day);
    return isNaN(parsedDate.getTime()) ? null : parsedDate;
  }

  // 2. Fallback to standard JS parsing (for YYYY-MM-DD or standard text dates)
  var parsed = new Date(str);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function loadTeacherMapping() {
  var configFolder = getFolderByLink(CONFIG_FOLDER_LINK);
  var rawRows = getExcelDataFromFolder(configFolder, MAPPING_FILE_NAME);

  if (!rawRows) throw new Error("Critical Error: Mapping file missing.");

  var mapping = {};
  var header = rawRows[0].map(function(h) { return h.toString().toLowerCase().trim(); });
  var nameIdx = -1, emailIdx = -1, classIdx = -1, sectionIdx = -1;

  // Find column indices
  for (var c = 0; c < header.length; c++) {
    if (header[c].includes("name")) nameIdx = c;
    if (header[c].includes("email")) emailIdx = c;
    if (header[c].includes("class") || header[c].includes("grade")) classIdx = c;
    if (header[c].includes("section") || header[c].includes("sec")) sectionIdx = c;
  }

  for (var i = 1; i < rawRows.length; i++) {
    var row = rawRows[i];
    if (row.length < 3) continue;

    var teacherName = nameIdx > -1 ? row[nameIdx].toString().trim() : row[1].toString().trim();
    var teacherEmail = emailIdx > -1 ? row[emailIdx].toString().trim() : row[2].toString().trim();
    var classNum = (classIdx > -1 ? row[classIdx].toString().trim() : row[4].toString().trim()).toLowerCase().replace(/^(class|grade)[\s-.]*/, "").trim();
    var secLetter = (sectionIdx > -1 ? row[sectionIdx].toString().trim() : row[5].toString().trim()).toLowerCase().replace(/^(sec|section)[\s-.]*/, "").trim();

    if (classNum !== "" && secLetter !== "") {
      mapping[classNum + "_" + secLetter] = {
        name: teacherName,
        email: teacherEmail
      };
    }
  }
  return mapping;
}

function isGlobalHolidayOrWeekend(checkDate) {
  var dayOfWeek = checkDate.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) return true;

  try {
    var checkDateStr = Utilities.formatDate(checkDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var configFolder = getFolderByLink(CONFIG_FOLDER_LINK);
    var holidaysArray = getPublicHolidays(configFolder);

    if (holidaysArray.indexOf(checkDateStr) > -1) {
      Logger.log("Global Holiday detected: " + checkDateStr);
      return true;
    }
  } catch(e) {
    Logger.log("Warning: Holiday check failed, falling back to red column check. " + e.message);
  }
  return false;
}

function buildAttendanceWorkbook(ss, rosterData, monthsToCreate, holidays) {
  var numStudents = rosterData.length;
  for (var m = 0; m < monthsToCreate.length; m++) {
    var monthInfo = monthsToCreate[m];
    Logger.log("    -> Generating Sheet Tab: " + monthInfo.name);
    var sheet = ss.insertSheet(monthInfo.name);
    var daysInMonth = new Date(monthInfo.year, monthInfo.monthIndex + 1, 0).getDate();

    var headers = ["Roll No.", "Child ID", "Name of the Student"];
    for (var d = 1; d <= daysInMonth; d++) headers.push(d.toString());
    headers.push("No of days Present", "No of days Absent", "No of days Late", "Total Attendance", "Total Present", "Percentage");

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(2, 1, numStudents, 3).setValues(rosterData);

    refreshFormulasAndStyles(sheet, numStudents, daysInMonth, monthInfo.year, monthInfo.monthIndex, monthInfo.name, holidays);
  }
  if (ss.getSheetByName("Sheet1")) ss.deleteSheet(ss.getSheetByName("Sheet1"));
}

function refreshFormulasAndStyles(sheet, totalStudents, daysInMonth, targetYear, monthIndex, monthName, holidays) {
  var startDayLetter = "D", endDayLetter = getColLetter(3 + daysInMonth), formulas = [];
  for (var i = 0; i < totalStudents; i++) {
    var r = i + 2;
    formulas.push([
      '=COUNTIF(' + startDayLetter + r + ':' + endDayLetter + r + ', "Present")',
      '=COUNTIF(' + startDayLetter + r + ':' + endDayLetter + r + ', "Absent")',
      '=COUNTIF(' + startDayLetter + r + ':' + endDayLetter + r + ', "Late")',
      '=SUM(' + getColLetter(4 + daysInMonth) + r + ':' + getColLetter(6 + daysInMonth) + r + ')',
      '=' + getColLetter(4 + daysInMonth) + r + '+' + getColLetter(6 + daysInMonth) + r,
      '=IF(' + getColLetter(7 + daysInMonth) + r + '>0, ' + getColLetter(8 + daysInMonth) + r + '/' + getColLetter(7 + daysInMonth) + r + ', 0)'
    ]);
  }
  sheet.getRange(2, 4 + daysInMonth, totalStudents, 6).setFormulas(formulas);

  var footerRow = totalStudents + 2;
  sheet.getRange(footerRow, 1, 3, sheet.getLastColumn()).clearContent();

  for (var col = 4; col <= 3 + daysInMonth; col++) {
    var cLetter = getColLetter(col);
    sheet.getRange(footerRow, col).setFormula('=COUNTIF(' + cLetter + '2:' + cLetter + (totalStudents + 1) + ', "Present") + COUNTIF(' + cLetter + '2:' + cLetter + (totalStudents + 1) + ', "Late")');
  }

  sheet.getRange(footerRow + 2, 4 + daysInMonth).setValue("CLASS AVERAGE:");
  sheet.getRange(footerRow + 2, 5 + daysInMonth).setFormula('=AVERAGE(' + getColLetter(9 + daysInMonth) + '2:' + getColLetter(9 + daysInMonth) + (totalStudents + 1) + ')');

  sheet.getRange(2, 9 + daysInMonth, totalStudents, 1).setNumberFormat("0.00%");
  sheet.getRange(footerRow + 2, 5 + daysInMonth).setNumberFormat("0.00%");

  var validationRule = SpreadsheetApp.newDataValidation().requireValueInList(['Present', 'Absent', 'Late']).build();
  for (var d = 1; d <= daysInMonth; d++) {
    var dayCol = 3 + d;
    var currentCellDate = new Date(targetYear, monthIndex, d);
    var dayOfWeek = currentCellDate.getDay();
    var formattedCellDate = Utilities.formatDate(currentCellDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

    var isHoliday = holidays && holidays.indexOf(formattedCellDate) > -1;
    var dayRange = sheet.getRange(2, dayCol, totalStudents, 1);

    if (dayOfWeek === 0 || dayOfWeek === 6 || isHoliday) {
      dayRange.clearDataValidations();
      dayRange.setBackground("#FF0000");
      var footerCell = sheet.getRange(footerRow, dayCol);
      footerCell.clearContent();
    } else {
      dayRange.setDataValidation(validationRule);
      dayRange.setBackground(null);
    }
  }

  var dynamicDayGridRange = sheet.getRange(2, 4, totalStudents, daysInMonth);
  var pRule = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Present").setBackground("#D4EDDA").setFontColor("#155724").setRanges([dynamicDayGridRange]).build();
  var aRule = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Absent").setBackground("#F8D7DA").setFontColor("#721C24").setRanges([dynamicDayGridRange]).build();
  var lRule = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Late").setBackground("#FFF3CD").setFontColor("#856404").setRanges([dynamicDayGridRange]).build();
  sheet.setConditionalFormatRules([pRule, aRule, lRule]);

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);
  sheet.getRange(1, 1, 1, 9 + daysInMonth).setFontWeight("bold").setBackground("#F1F4F9").setHorizontalAlignment("center");
  sheet.getRange(footerRow + 2, 4 + daysInMonth).setFontWeight("bold");

  sheet.autoResizeColumns(1, 3);
  sheet.setColumnWidths(4, daysInMonth, 90);
  sheet.setColumnWidths(4 + daysInMonth, 6, 115);

  // if (ADD_VISUALISATIONS) { //old visualisation demo (derecated)
  //   buildVisualizations(sheet, totalStudents, daysInMonth);
  //   if (monthName) Logger.log("      -> Generated Visualizations for " + monthName);
  // }
}

function fillWithPlaceholders(range, rows, cols) {
  var output = [];
  for (var r = 0; r < rows; r++) {
    var rowArr = [];
    for (var c = 0; c < cols; c++) rowArr.push("-");
    output.push(rowArr);
  }
  range.setValues(output).setHorizontalAlignment("center");
}

function buildVisualizations(sheet, totalStudents, daysInMonth) {
  var footerRow = totalStudents + 2;

  var existingCharts = sheet.getCharts();
  for (var c = 0; c < existingCharts.length; c++) sheet.removeChart(existingCharts[c]);

  var chartsStartRow = footerRow + 5;
  sheet.getRange(chartsStartRow, 1, 40, 6).clear();

  sheet.getRange(chartsStartRow, 1).setValue("Status").setFontWeight("bold").setBackground("#F1F4F9");
  sheet.getRange(chartsStartRow, 2).setValue("Total Logs").setFontWeight("bold").setBackground("#F1F4F9");
  sheet.getRange(chartsStartRow + 1, 1).setValue("Present");
  sheet.getRange(chartsStartRow + 1, 2).setFormula("=SUM(" + getColLetter(4 + daysInMonth) + "2:" + getColLetter(4 + daysInMonth) + (totalStudents + 1) + ")");
  sheet.getRange(chartsStartRow + 2, 1).setValue("Absent");
  sheet.getRange(chartsStartRow + 2, 2).setFormula("=SUM(" + getColLetter(5 + daysInMonth) + "2:" + getColLetter(5 + daysInMonth) + (totalStudents + 1) + ")");
  sheet.getRange(chartsStartRow + 3, 1).setValue("Late");
  sheet.getRange(chartsStartRow + 3, 2).setFormula("=SUM(" + getColLetter(6 + daysInMonth) + "2:" + getColLetter(6 + daysInMonth) + (totalStudents + 1) + ")");

  sheet.getRange(chartsStartRow + 1, 2, 3, 1).setNumberFormat("0");
  sheet.getRange(chartsStartRow, 1, 4, 2).setBorder(true, true, true, true, null, null).setHorizontalAlignment("center");

  sheet.getRange(chartsStartRow, 4).setValue("Day Index").setFontWeight("bold").setBackground("#F1F4F9");
  sheet.getRange(chartsStartRow, 5).setValue("Present Count").setFontWeight("bold").setBackground("#F1F4F9");

  var trendFormulas = [];
  for (var d = 1; d <= daysInMonth; d++) {
    var dayColLetter = getColLetter(3 + d);
    trendFormulas.push([d.toString(), "=" + dayColLetter + footerRow]);
  }
  sheet.getRange(chartsStartRow + 1, 4, daysInMonth, 2).setValues(trendFormulas).setNumberFormat("0").setHorizontalAlignment("center");
  sheet.getRange(chartsStartRow, 4, daysInMonth + 1, 2).setBorder(true, true, true, true, null, null);
  sheet.setColumnWidth(4, 90);
  sheet.setColumnWidth(5, 110);

  var pieChart = sheet.newChart()
    .asPieChart()
    .addRange(sheet.getRange(chartsStartRow, 1, 4, 2))
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setOption('title', 'Monthly Attendance Status Ratio')
    .setOption('pieHole', 0.4)
    .setOption('colors', ['#28a745', '#dc3545', '#ffc107'])
    .setOption('chartArea', {left: '5%', top: '15%', width: '90%', height: '75%'})
    .setPosition(chartsStartRow + 5, 1, 0, 0)
    .build();
  sheet.insertChart(pieChart);

  var trendChart = sheet.newChart()
    .asColumnChart()
    .addRange(sheet.getRange(chartsStartRow, 4, daysInMonth + 1, 2))
    .setOption('title', 'Daily Active Attendance Profile Trend')
    .setOption('legend', {position: 'none'})
    .setOption('colors', ['#4285F4'])
    .setOption('hAxis', {title: 'Day of the Calendar Month', textStyle: {fontSize: 9}})
    .setOption('vAxis', {title: 'Total Tracked Headcount', minValue: 0})
    .setOption('chartArea', {left: '10%', top: '15%', width: '85%', height: '70%'})
    .setPosition(chartsStartRow + 23, 1, 0, 0)
    .build();
  sheet.insertChart(trendChart);
}

function generateAlertBlocks(ss, sheet, studentNames, today) {
  var MAX_LOOKBACK = Math.max(
    CONSECUTIVE_ABSENT_THRESHOLD_DAYS + ALLOWED_PRESENT_SKIPS,
    CONSECUTIVE_LATE_THRESHOLD_DAYS + ALLOWED_PRESENT_SKIPS,
    CONSECUTIVE_MIXED_THRESHOLD_DAYS + ALLOWED_PRESENT_SKIPS,
    ABSENT_LOOKBACK_DAYS,
    LATE_LOOKBACK_DAYS,
    20
  );

  var validColumnsToCheck = [];
  var checkDayOffset = today.getDate() - 1;
  var scanSheet = sheet, scanMonthObj = new Date(today.getTime());

  while (validColumnsToCheck.length < MAX_LOOKBACK) {
    if (checkDayOffset < 1) {
      scanMonthObj = new Date(scanMonthObj.getFullYear(), scanMonthObj.getMonth() - 1, 1);
      var prevMonthName = scanMonthObj.toLocaleString('en-US', { month: 'long' });
      scanSheet = ss.getSheetByName(prevMonthName);
      if (!scanSheet) break;
      checkDayOffset = new Date(scanMonthObj.getFullYear(), scanMonthObj.getMonth() + 1, 0).getDate();
    }
    var currentHeaders = scanSheet.getRange(1, 1, 1, scanSheet.getLastColumn()).getValues()[0];
    var colIdx = -1;
    for (var h = 0; h < currentHeaders.length; h++) {
      if (parseInt(currentHeaders[h], 10) === checkDayOffset) {
        colIdx = h + 1;
        break;
      }
    }
    if (colIdx !== -1) {
      var bgColor = scanSheet.getRange(2, colIdx).getBackground().toLowerCase();
      if (bgColor !== '#ff0000') {
        var sampleData = scanSheet.getRange(2, colIdx, Math.min(5, studentNames.length), 1).getValues();
        var isInstructionalDay = sampleData.some(function(cell) {
          var s = cell[0].toString().trim();
          return s === "Present" || s === "Absent" || s === "Late";
        });
        if (isInstructionalDay) { validColumnsToCheck.push({ sheet: scanSheet, colIndex: colIdx }); }
      }
    }
    checkDayOffset--;
  }

  var teacherChartData = [];
  var stakeholderChartData = [];

  function getForgivingStreak(statusArray, validTypes, maxSkips) {
    var validCount = 0; var skips = 0; var windowSize = 0;
    for (var i = 0; i < statusArray.length; i++) {
      if (validTypes.indexOf(statusArray[i]) !== -1) {
        validCount++; windowSize++;
      } else {
        if (skips < maxSkips) { skips++; windowSize++; }
        else { break; }
      }
    }
    return { validCount: validCount, windowSize: windowSize };
  }

  // Student IDs (Child ID, col B) aligned to the same rows the status loop
  // scans (row = sIdx + 2), so every flagged entry can carry its ID for the
  // inline email roster.
  var studentIds = [];
  try {
    var idVals = sheet.getRange(2, 2, studentNames.length, 1).getValues();
    for (var ii = 0; ii < idVals.length; ii++) studentIds.push(idVals[ii][0].toString().trim());
  } catch (e) { studentIds = []; }

  for (var sIdx = 0; sIdx < studentNames.length; sIdx++) {
    var rowNum = sIdx + 2; var statuses = [];
    for (var c = 0; c < validColumnsToCheck.length; c++) {
      statuses.push(validColumnsToCheck[c].sheet.getRange(rowNum, validColumnsToCheck[c].colIndex).getValue().toString().trim());
    }

    var absentStreak = getForgivingStreak(statuses, ["Absent"], ALLOWED_PRESENT_SKIPS);
    var lateStreak = getForgivingStreak(statuses, ["Late"], ALLOWED_PRESENT_SKIPS);
    var mixedStreak = getForgivingStreak(statuses, ["Absent", "Late"], ALLOWED_PRESENT_SKIPS);

    if (absentStreak.validCount >= CONSECUTIVE_ABSENT_THRESHOLD_DAYS ||
        lateStreak.validCount >= CONSECUTIVE_LATE_THRESHOLD_DAYS ||
        mixedStreak.validCount >= CONSECUTIVE_MIXED_THRESHOLD_DAYS) {
      var maxWindow = Math.max(absentStreak.windowSize, lateStreak.windowSize, mixedStreak.windowSize);
      var aCount = 0, lCount = 0;
      for (var w = 0; w < maxWindow; w++) {
        if (statuses[w] === "Absent") aCount++;
        if (statuses[w] === "Late") lCount++;
      }
      teacherChartData.push({ id: studentIds[sIdx] || "", name: studentNames[sIdx], absences: aCount, lates: lCount, total: aCount + lCount });
    }

    var daysToCheckAbsent = Math.min(statuses.length, ABSENT_LOOKBACK_DAYS);
    var absentCount = 0;
    for (var i = 0; i < daysToCheckAbsent; i++) { if (statuses[i] === "Absent") absentCount++; }

    var daysToCheckLate = Math.min(statuses.length, LATE_LOOKBACK_DAYS);
    var lateCount = 0;
    for (var i = 0; i < daysToCheckLate; i++) { if (statuses[i] === "Late") lateCount++; }

    if (absentCount >= ABSENT_THRESHOLD_DAYS || lateCount >= LATE_THRESHOLD_DAYS) {
      stakeholderChartData.push({ id: studentIds[sIdx] || "", name: studentNames[sIdx], absences: absentCount, lates: lateCount, total: absentCount + lateCount });
    }
  }

  teacherChartData.sort(function(a, b) { return b.total - a.total; });
  stakeholderChartData.sort(function(a, b) { return b.total - a.total; });

  var MAX_CHART_STUDENTS = 20;

  // Builds a collapsible <details> block listing every flagged student in a
  // category with their Student ID, Name and day-count. <details>/<summary> is
  // natively collapsible in Apple Mail & Outlook; Gmail strips the toggle but
  // still renders the list expanded — so it degrades gracefully everywhere.
  // metricKey = 'absences' | 'lates'; only students with a positive count for
  // that metric are listed. accent = header/border color.
  function buildCategoryRoster(dataArray, metricKey, label, accent) {
    var rows = dataArray.filter(function(d) { return (d[metricKey] || 0) > 0; });
    if (rows.length === 0) return "";
    var body = "";
    for (var i = 0; i < rows.length; i++) {
      var rowBg = (i % 2 === 0) ? "#ffffff" : "#f7fafc";
      body += '<tr style="background-color:' + rowBg + ';">' +
              '<td style="padding:6px 10px; font-size:13px; color:#4a5568; border-bottom:1px solid #edf2f7;">' + (rows[i].id || "—") + '</td>' +
              '<td style="padding:6px 10px; font-size:13px; color:#2d3748; border-bottom:1px solid #edf2f7;">' + rows[i].name + '</td>' +
              '<td style="padding:6px 10px; font-size:13px; color:#2d3748; text-align:center; font-weight:bold; border-bottom:1px solid #edf2f7;">' + rows[i][metricKey] + '</td>' +
              '</tr>';
    }
    return '<details style="margin-top:12px; border:1px solid #e2e8f0; border-radius:6px; overflow:hidden;">' +
           '<summary style="cursor:pointer; padding:10px 14px; background-color:' + accent + '; color:#fff; font-size:14px; font-weight:bold; list-style:none;">' +
           label + ' (' + rows.length + ' student' + (rows.length === 1 ? '' : 's') + ') ▾</summary>' +
           '<table style="width:100%; border-collapse:collapse;">' +
           '<tr style="background-color:#edf2f7;">' +
           '<th style="padding:6px 10px; font-size:12px; color:#4a5568; text-align:left;">Student ID</th>' +
           '<th style="padding:6px 10px; font-size:12px; color:#4a5568; text-align:left;">Name</th>' +
           '<th style="padding:6px 10px; font-size:12px; color:#4a5568; text-align:center;">Days</th>' +
           '</tr>' + body + '</table></details>';
  }

  // Two collapsible categories (Absences, Lates) covering ALL flagged students
  // — not just the top-20 shown in the chart image.
  function buildCollapsibleRosters(dataArray) {
    return buildCategoryRoster(dataArray, 'absences', '🔴 Absences', '#c53030') +
           buildCategoryRoster(dataArray, 'lates', '🟡 Lates', '#b7791f');
  }

  var teacherChartDisplay = teacherChartData.slice(0, MAX_CHART_STUDENTS);
  var stakeholderChartDisplay = stakeholderChartData.slice(0, MAX_CHART_STUDENTS);
  // teacherData / stakeholderData expose the raw flagged-student arrays (full,
  // not sliced to the top 20) so other consumers — e.g. the Analysis_Dashboard
  // Section 4 — can render the same chronic-risk data without re-deriving the
  // lookback logic. Each entry: { name, absences, lates, total }.
  var result = {
    teacherHtml: "", teacherBlob: null, stakeholderHtml: "", stakeholderBlob: null,
    teacherData: teacherChartData, stakeholderData: stakeholderChartData
  };

  if (teacherChartDisplay.length > 0) {
    var chartTitle = teacherChartData.length > MAX_CHART_STUDENTS ? "Consecutive Absences & Lates (Top " + MAX_CHART_STUDENTS + ")" : "Consecutive Absences & Lates";
    result.teacherBlob = createStackedBarChart(teacherChartDisplay, chartTitle);
    result.teacherHtml = '<div style="background-color: #fff; padding: 15px; border-radius: 8px; border: 1px solid #e2e8f0; margin-bottom: 20px;">' +
                         '<h3 style="color: #2d3748; margin-top: 0;">⚠️ Immediate Attention Required</h3>' +
                         '<p style="color: #718096; font-size: 14px;">The chart below highlights students currently on a consecutive streak (including mixed and skip-adjusted streaks) of absences or late arrivals.</p>' +
                         '<img src="cid:teacher_chart" style="max-width: 100%; height: auto; border-radius: 4px;">' +
                         buildCollapsibleRosters(teacherChartData) + '</div>';
  }

  if (stakeholderChartDisplay.length > 0) {
    var chartTitle2 = stakeholderChartData.length > MAX_CHART_STUDENTS ? "Chronic Absences & Lates (Top " + MAX_CHART_STUDENTS + ")" : "Chronic Absences & Lates";
    result.stakeholderBlob = createStackedBarChart(stakeholderChartDisplay, chartTitle2);
    result.stakeholderHtml = '<div style="background-color: #fff; padding: 15px; border-radius: 8px; border: 1px solid #e2e8f0; margin-bottom: 20px;">' +
                             '<h3 style="color: #2d3748; margin-top: 0;">📉 Chronic Attendance Risks</h3>' +
                             '<p style="color: #718096; font-size: 14px;">Students with frequent absences or late arrivals over the lookback period.</p>' +
                             '<img src="cid:stakeholder_chart_cid" style="max-width: 100%; height: auto; border-radius: 4px;">' +
                             buildCollapsibleRosters(stakeholderChartData) + '</div>';
  }

  return result;
}

function createStackedBarChart(data, title) {
  var dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Student')
    .addColumn(Charts.ColumnType.NUMBER, 'Absences')
    .addColumn(Charts.ColumnType.NUMBER, 'Lates');

  for (var i = 0; i < data.length; i++) {
    var metrics = [];
    if (data[i].absences > 0) metrics.push(data[i].absences + "A");
    if (data[i].lates > 0) metrics.push(data[i].lates + "L");
    var studentLabel = data[i].name + " (" + metrics.join(", ") + ")";
    dataTable.addRow([studentLabel, data[i].absences, data[i].lates]);
  }

  var chartHeight = Math.max(250, data.length * 45 + 100);
  var chart = Charts.newBarChart()
    .setDataTable(dataTable)
    .setTitle(title)
    .setStacked()
    .setColors(['#e53e3e', '#d69e2e'])
    .setDimensions(550, chartHeight)
    .setXAxisTitle('Total Days (Combined)')
    .setOption('chartArea', {left: '25%', width: '70%'})
    .build();

  return chart.getAs('image/png');
}

// =========================================================================
// SIMPLIFIED HTML EMAIL BUILDER
// =========================================================================
function buildSimpleHtmlEmail(dateStr, classNum, section, teacherName, formUrl, alertsHtml) {
  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>';
  html += '<body style="margin: 0; padding: 0; font-family: Arial, sans-serif; background-color: #f5f5f5;">';
  html += '<div style="max-width: 600px; margin: 20px auto; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">';

  // Header
  html += '<div style="background-color: #4CAF50; padding: 20px; text-align: center;">';
  html += '<h1 style="margin: 0; color: white; font-size: 24px;">📋 Daily Attendance</h1>';
  html += '</div>';

  // Body
  html += '<div style="padding: 30px;">';
  html += '<p style="font-size: 16px; color: #333; margin-top: 0;">Hello ' + teacherName + ',</p>';
  html += '<p style="font-size: 14px; color: #666; line-height: 1.6;">Please mark attendance for your class:</p>';

  // Class Info Box
  html += '<div style="background-color: #f9f9f9; padding: 15px; border-radius: 5px; margin: 20px 0;">';
  html += '<p style="margin: 5px 0; color: #333;"><strong>Date:</strong> ' + dateStr + '</p>';
  html += '<p style="margin: 5px 0; color: #333;"><strong>Class:</strong> ' + classNum + '-' + section.toUpperCase() + '</p>';
  html += '</div>';

  // Alerts (if any)
  if (alertsHtml && alertsHtml !== "") {
    html += alertsHtml;
  }

  // Button
  html += '<div style="text-align: center; margin: 30px 0;">';
  html += '<a href="' + formUrl + '" style="display: inline-block; background-color: #4CAF50; color: white; padding: 15px 40px; text-decoration: none; border-radius: 5px; font-size: 16px; font-weight: bold;">Open Attendance Form</a>';
  html += '</div>';

  html += '<p style="font-size: 12px; color: #999; margin-top: 30px; border-top: 1px solid #eee; padding-top: 15px;">You will receive a confirmation email once you submit the form.</p>';

  html += '</div>'; // End body padding
  html += '</div>'; // End container
  html += '</body></html>';

  return html;
}

// =========================================================================
// HTML EMAIL BUILDERS (OLD - KEEP FOR STAKEHOLDER REPORTS)
// =========================================================================
function buildHtmlEmail(todayStr, classNum, section, teacherName, liveUrl, combinedAlerts) {
  return '<div style="background-color: #f4f7f9; padding: 40px 20px; font-family: \'Segoe UI\',Helvetica,Arial,sans-serif; color: #2d3748; line-height: 1.6; max-width: 600px; margin: 0 auto; border-radius: 12px; border: 1px solid #e2e8f0;">' +
  '<div style="text-align: center; margin-bottom: 25px;"><span style="font-size: 55px; margin: 0;">📋</span></div>' +
  '<h2 style="color: #1a365d; text-align: center; font-size: 24px; margin-top: 0; margin-bottom: 8px;">Good Morning, ' + teacherName + '! 😊</h2>' +
  '<p style="font-size: 16px; text-align: center; color: #4a5568; margin-top: 0; margin-bottom: 30px;">It is time to log classroom registration variables. Let\'s make today a spectacular day of learning!</p>' +
  combinedAlerts +
  '<div style="background-color: #ffffff; padding: 20px 25px; border-radius: 10px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); margin: 20px 0; border-left: 6px solid #3182ce;">' +
  '<table style="width: 100%; font-size: 16px; border-collapse: collapse;">' +
  '<tr style="border-bottom: 1px solid #edf2f7;"><td style="padding: 8px 0; color: #718096;"><strong>📅 Attendance Date:</strong></td><td style="padding: 8px 0; text-align: right; color: #2d3748; font-weight: 600;">' + todayStr + '</td></tr>' +
  '<tr style="border-bottom: 1px solid #edf2f7;"><td style="padding: 8px 0; color: #718096;"><strong>⩃ Class Grade:</strong></td><td style="padding: 8px 0; text-align: right; color: #2d3748; font-weight: 600;">Class ' + classNum + '</td></tr>' +
  '<tr><td style="padding: 8px 0; color: #718096;"><strong> 𑗕 Section Code:</strong></td><td style="padding: 8px 0; text-align: right; color: #2d3748; font-weight: 600;">Section ' + section.toUpperCase() + '</td></tr>' +
  '</table>' +
  '</div>' +
  '<div style="text-align: center; margin: 35px 0;">' +
  '<a href="' + liveUrl + '" target="_blank" style="background-color: #319795; color: #ffffff; padding: 14px 35px; text-decoration: none; font-size: 16px; font-weight: bold; border-radius: 8px; display: inline-block; box-shadow: 0 4px 14px rgba(49, 151, 149, 0.4); transition: background-color 0.2s ease;">📝 Open Attendance Grid</a>' +
  '</div>' +
  '<hr style="border: 0; border-top: 1px solid #e2e8f0; margin: 35px 0;">' +
  '<p style="font-size: 11px; text-align: center; color: #a0aec0; margin-bottom: 0; letter-spacing: 0.5px;">Automated Attendance Distribution Matrix &bull; Core Classroom Operations</p>' +
  '</div>';
}

function buildStakeholderDigestHtml(todayStr, digestContent) {
  return '<div style="background-color: #f4f7f9; padding: 40px 20px; font-family: \'Segoe UI\',Helvetica,Arial,sans-serif; color: #2d3748; line-height: 1.6; max-width: 600px; margin: 0 auto; border-radius: 12px; border: 1px solid #e2e8f0;">' +
  '<h2 style="color: #1a365d; text-align: center; font-size: 24px; margin-top: 0; margin-bottom: 8px;">🚨 Comprehensive Attendance & Lateness Report</h2>' +
  '<p style="font-size: 16px; text-align: center; color: #4a5568; margin-top: 0; margin-bottom: 30px;">Date: ' + todayStr + '</p>' +
  digestContent +
  '</div>';
}

// =========================================================================
// CORE SYNC SUBROUTINE
// =========================================================================
function executeSheetSyncProcessing(sheet, activeFormId, dayOfMonthDigit, ssId) {
  try {
    var form = FormApp.openById(activeFormId);
    var responses = form.getResponses();
    if (responses.length === 0) {
      Logger.log("      [SYNC] No responses yet for form on '" + sheet.getName() + "'. Nothing to write.");
      return;
    }

    // Locate the column whose header equals today's day-of-month digit.
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var dateColIndex = -1;
    for (var c = 0; c < headers.length; c++) {
      if (parseInt(headers[c], 10) === dayOfMonthDigit) {
        dateColIndex = c + 1;
        break;
      }
    }
    if (dateColIndex === -1) {
      Logger.log("      [SYNC][WARNING] No day column for day " + dayOfMonthDigit +
                 " on sheet '" + sheet.getName() + "'. Possible weekend/holiday or month mismatch. Skipping.");
      return;
    }

    // Build a normalized lookup: collapsed-whitespace + lowercase name -> sheet row index.
    // This tolerates trailing/double spaces and case differences between the form's
    // grid labels and the sheet's Name column (a common cause of silently-blank cells).
    var studentIndexByName = {};
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var nameRange = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
      for (var r = 0; r < nameRange.length; r++) {
        var normName = normalizeNameForMatch(nameRange[r][0]);
        if (normName !== "") studentIndexByName[normName] = r; // 0-based within student block
      }
    }

    // AUTHORIZATION: only the class's assigned teacher may set attendance.
    // We enforce this here (in the hourly sync) instead of via a per-form
    // onFormSubmit trigger, because one form-submit trigger per class would
    // exceed Apps Script's 20-trigger-per-user cap. Trade-off: an unauthorized
    // submission is ignored at the next sync rather than the instant of submit.
    var authorizedEmail = PropertiesService.getScriptProperties()
      .getProperty('AUTHORIZED_TEACHER_' + ssId);
    authorizedEmail = authorizedEmail ? authorizedEmail.toLowerCase().trim() : null;

    // Pick the LATEST response submitted BY the authorized teacher. Scanning
    // newest-first means an unauthorized submission can neither overwrite nor
    // block the legitimate one.
    var chosenResponse = null;
    for (var ri = responses.length - 1; ri >= 0; ri--) {
      if (!authorizedEmail) { chosenResponse = responses[ri]; break; } // no guard configured -> latest wins
      var respEmail = responses[ri].getRespondentEmail();
      respEmail = respEmail ? respEmail.toLowerCase().trim() : "";
      if (respEmail === authorizedEmail) { chosenResponse = responses[ri]; break; }
    }

    // Email each submitter a single accept/reject confirmation per day. Deduped
    // by submitter email (persisted in NOTIFIED_<ssId>) so the hourly re-sync
    // over accumulated responses emails each teacher at most once per day; the
    // record is cleared nightly by automated_closeForms.
    notifySubmitters(ssId, responses, authorizedEmail, sheet.getName(), dayOfMonthDigit);

    if (!chosenResponse) {
      Logger.log("      [SYNC][REJECTED] '" + sheet.getName() + "' day " + dayOfMonthDigit +
                 ": " + responses.length + " response(s) present but none from the authorized teacher (" +
                 (authorizedEmail || "none set") + "). Nothing written.");
      return;
    }

    var itemResponses = chosenResponse.getItemResponses();
    var written = 0, skipped = 0;
    var skippedNames = [];

    for (var j = 0; j < itemResponses.length; j++) {
      var itemResponse = itemResponses[j];
      if (itemResponse.getItem().getType() === FormApp.ItemType.GRID) {
        var rows = itemResponse.getItem().asGridItem().getRows();
        var studentStatuses = itemResponse.getResponse();

        for (var s = 0; s < rows.length; s++) {
          var status = studentStatuses[s];
          if (!status) { continue; } // teacher left this student blank in the grid

          var key = normalizeNameForMatch(rows[s]);
          if (studentIndexByName.hasOwnProperty(key)) {
            var sheetIndex = studentIndexByName[key];
            sheet.getRange(sheetIndex + 2, dateColIndex).setValue(status);
            written++;
          } else {
            skipped++;
            skippedNames.push(rows[s]);
          }
        }
      }
    }

    Logger.log("      [SYNC] '" + sheet.getName() + "' day " + dayOfMonthDigit +
               ": wrote " + written + " status(es), skipped " + skipped +
               " (accepted from " + (chosenResponse.getRespondentEmail() || "unknown") + ").");
    if (skipped > 0) {
      Logger.log("      [SYNC][WARNING] " + skipped + " form row(s) had no matching student in the sheet: " +
                 skippedNames.join(" | "));
    }
  } catch(err) {
    Logger.log("Error inside sync block: " + err.message);
  }
}

// Normalizes a name for tolerant matching between form grid rows and the sheet's
// Name column: trims, collapses internal whitespace to a single space, lowercases.
function normalizeNameForMatch(value) {
  if (value === null || value === undefined) return "";
  return value.toString().replace(/\s+/g, " ").trim().toLowerCase();
}

// Emails each distinct submitter a single accept/reject confirmation per day.
// The authorized teacher gets an "accepted" note; anyone else gets a "rejected"
// note (their submission is not written to the sheet). Because the hourly sync
// re-reads all accumulated responses, we persist the set of already-notified
// emails in Script Property NOTIFIED_<ssId> and skip them on later runs, so each
// teacher is emailed at most once per day. automated_closeForms clears this key
// nightly, resetting the set for the next day.
function notifySubmitters(ssId, responses, authorizedEmail, sheetName, dayOfMonthDigit) {
  var props = PropertiesService.getScriptProperties();
  var notifiedKey = 'NOTIFIED_' + ssId;

  // Derive a readable class label from the workbook name (e.g.
  // "Class_5_A_2026-2027" -> "Class 5-A") so the confirmation names the class.
  var classLabel = "your class";
  try {
    var wbName = SpreadsheetApp.openById(ssId).getName();
    var parts = wbName.split("_");            // [Class, 5, A, 2026-2027]
    if (parts.length >= 3) {
      classLabel = parts[0] + " " + parts[1] + "-" + parts[2];
    }
  } catch (e) { /* fall back to generic label */ }

  var notified = [];
  try {
    var stored = props.getProperty(notifiedKey);
    if (stored) notified = JSON.parse(stored);
  } catch (e) {
    notified = [];
  }

  var newlyNotified = false;

  for (var i = 0; i < responses.length; i++) {
    var email = responses[i].getRespondentEmail();
    email = email ? email.toLowerCase().trim() : "";
    if (email === "") continue;                       // no collected email -> can't notify
    if (notified.indexOf(email) !== -1) continue;      // already emailed today

    var isAccepted = (!authorizedEmail) || (email === authorizedEmail);
    var subject = isAccepted ? "✅ Attendance Submitted Successfully — " + classLabel
                             : "❌ Attendance Submission Rejected — " + classLabel;
    var htmlBody = buildSyncConfirmationHtml(isAccepted, authorizedEmail, sheetName, dayOfMonthDigit, classLabel);

    try {
      MailApp.sendEmail({ to: email, subject: subject, htmlBody: htmlBody });
      notified.push(email);
      newlyNotified = true;
      Logger.log("      [SYNC][EMAIL] " + (isAccepted ? "accepted" : "rejected") + " notice sent to " + email);
    } catch (err) {
      Logger.log("      [SYNC][EMAIL][ERROR] Could not email " + email + ": " + err.message);
    }
  }

  if (newlyNotified) {
    props.setProperty(notifiedKey, JSON.stringify(notified));
  }
}

// Builds the accept/reject confirmation email body sent from the hourly sync.
function buildSyncConfirmationHtml(isAccepted, authorizedEmail, sheetName, dayOfMonthDigit, classLabel) {
  classLabel = classLabel || "your class";
  var headerColor = isAccepted ? "#4CAF50" : "#f44336";
  var headerText = isAccepted ? "✅ Submission Accepted" : "❌ Submission Rejected";
  var message = isAccepted
    ? "Your attendance for <strong>" + classLabel + "</strong> on " + sheetName + " " + dayOfMonthDigit +
      " has been recorded successfully."
    : "You are not authorized to submit attendance for <strong>" + classLabel + "</strong>, so your submission was not recorded." +
      (authorizedEmail ? " Only " + authorizedEmail + " can submit for this class." : "");

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>';
  html += '<body style="margin: 0; padding: 0; font-family: Arial, sans-serif; background-color: #f5f5f5;">';
  html += '<div style="max-width: 600px; margin: 20px auto; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">';
  html += '<div style="background-color: ' + headerColor + '; padding: 20px; text-align: center;">';
  html += '<h1 style="margin: 0; color: white; font-size: 24px;">' + headerText + '</h1>';
  html += '</div>';
  html += '<div style="padding: 30px;">';
  html += '<p style="font-size: 16px; color: #333;">' + message + '</p>';
  if (!isAccepted) {
    html += '<p style="margin-top: 20px; padding: 15px; background-color: #fff3cd; border-left: 4px solid #f44336; color: #856404;">⚠️ This data was NOT saved. Contact your administrator if you believe this is an error.</p>';
  }
  html += '</div></div></body></html>';
  return html;
}

/**
 * Visualisation Function 1: Creates the 'Analysis_Dashboard' tab if it does NOT already exist.
 * Automatically populates it upon creation.
 * 
 * @param {Spreadsheet} ss - The Target Spreadsheet object.
 */
function createDashboardIfNotExists(ss) {
  if (!ss) return;

  const dashboardName = 'Analysis_Dashboard';
  let dashboardSheet = ss.getSheetByName(dashboardName);

  if (!dashboardSheet) {
    dashboardSheet = ss.insertSheet(dashboardName, 0);  // create as the FIRST tab
    Logger.log("Created missing 'Analysis_Dashboard' for: " + ss.getName());
    // Populate the newly created sheet
    updateDashboard(ss);
  } else {
    ss.setActiveSheet(dashboardSheet);
    ss.moveActiveSheet(1);  // ensure it stays pinned as the first tab
    Logger.log("'Analysis_Dashboard' already exists for: " + ss.getName());
  }
}

/**
 *  Visualisation Function 2: Re-calculates metrics across all sheets and updates the existing 'Analysis_Dashboard'.
 * 
 * @param {Spreadsheet} ss - The Target Spreadsheet object.
 */
function updateDashboard(ss) {
  if (!ss) return;

  const dashboardName = 'Analysis_Dashboard';
  let dashboardSheet = ss.getSheetByName(dashboardName);

  // Fallback check to ensure tab exists before updating
  if (!dashboardSheet) {
    dashboardSheet = ss.insertSheet(dashboardName, 0);  // create as the FIRST tab
  } else {
    dashboardSheet.clear();
  }

  // Keep the dashboard pinned as the first tab so it's the landing view.
  ss.setActiveSheet(dashboardSheet);
  ss.moveActiveSheet(1);

  // Clear existing charts
  dashboardSheet.getCharts().forEach(c => dashboardSheet.removeChart(c));

  const currentClassName = ss.getName().split('_').slice(0, 3).join(' ') || "Class 1A";

  // 1. GET ALL MONTH TABS ONLY — use a WHITELIST of academic month names rather
  //    than a blacklist. A blacklist ("exclude tabs named Form Responses*") is
  //    fragile: any unexpected tab (e.g. "Form Responses 7", a renamed response
  //    tab, or a stray sheet) slips through and gets processed as if it were a
  //    month, injecting garbage rows into the metrics. Only real month tabs
  //    (June, July, ...) are ever valid, so match against that set exactly.
  const validMonthNames = {};
  getAcademicMonthsList(ACADEMIC_YEAR, START_MONTH, END_MONTH).forEach(function(m) {
    validMonthNames[m.name] = true;
  });
  const monthSheets = ss.getSheets().filter(function(s) {
    return validMonthNames[s.getName()] === true;
  });

  const studentMap = {};
  const monthClassMap = {};
  const momMap = {};

  // 2. PROCESS EACH MONTH TAB
  monthSheets.forEach(sheet => {
    const monthName = sheet.getName();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    let monthPresent = 0, monthAbsent = 0, monthLate = 0, monthTotal = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const studentId = row[1];
      const studentName = row[2];

      if (!studentId || studentId.toString().trim() === '' || 
          studentName.toString().toUpperCase().includes('CLASS AVERAGE') || 
          studentName.toString().toUpperCase().includes('STATUS')) continue;

      if (!studentMap[studentId]) {
        studentMap[studentId] = { name: studentName, className: currentClassName, present: 0, late: 0, absent: 0, total: 0 };
      }

      // Columns 3..N-6 are the daily status cells; the LAST 6 columns are
      // computed summaries (No of days Present/Absent/Late, Total Attendance,
      // Total Present, Percentage) holding NUMBERS. We must NOT let those
      // numbers count as sessions. So a session is only ever tallied when the
      // cell is a recognized status word — this simultaneously excludes the
      // summary numbers, empty weekend/holiday cells, and unfilled future days.
      for (let j = 3; j < row.length; j++) {
        const status = row[j];
        if (!status || status.toString().trim() === '') continue;
        const s = status.toString().trim().toLowerCase();

        if (s === 'present') { monthPresent++; studentMap[studentId].present++; monthTotal++; studentMap[studentId].total++; }
        else if (s === 'late') { monthLate++; studentMap[studentId].late++; monthTotal++; studentMap[studentId].total++; }
        else if (s === 'absent') { monthAbsent++; studentMap[studentId].absent++; monthTotal++; studentMap[studentId].total++; }
        // any other value (summary counts, percentage) is intentionally ignored
      }
    }

    if (monthTotal > 0) {
      momMap[monthName] = {
        present: monthPresent,
        absent: monthAbsent,
        late: monthLate,
        total: monthTotal,
        attRate: ((monthPresent + monthLate) / monthTotal) * 100,
        absRate: (monthAbsent / monthTotal) * 100,
        lateRate: (monthLate / monthTotal) * 100
      };

      if (!monthClassMap[monthName]) monthClassMap[monthName] = {};
      monthClassMap[monthName][currentClassName] = ((monthPresent + monthLate) / monthTotal) * 100;
    }
  });

  // 3. WRITE DATA TO DASHBOARD
  let currentRow = 1;

  // --- Section 1: Month on Month Analysis ---
  dashboardSheet.getRange(currentRow, 1).setValue("1. Month on Month Analysis").setFontWeight("bold").setFontSize(12);
  currentRow++;

  // Columns 9-11 (Present %, Late %, Absent %) are kept CONTIGUOUS and in this
  // exact order so the stacked chart below renders bottom->top = Present, Late,
  // Absent. These three are % of total sessions and sum to 100%.
  const momHeaders = ['Month', 'Present', 'Absent', 'Late', 'Total Sessions', 'Attendance Rate (%)', 'Absenteeism Rate (%)', 'Late Rate (%)', 'Present %', 'Late %', 'Absent %'];
  dashboardSheet.getRange(currentRow, 1, 1, momHeaders.length).setValues([momHeaders]).setFontWeight("bold");
  currentRow++;

  const momStartRow = currentRow;
  const momData = [];

  monthSheets.forEach(sheet => {
    const month = sheet.getName();
    if (momMap[month]) {
      const d = momMap[month];
      const presentRate = d.total > 0 ? (d.present / d.total) * 100 : 0;
      momData.push([
        month, d.present, d.absent, d.late, d.total, d.attRate.toFixed(2), d.absRate.toFixed(2), d.lateRate.toFixed(2),
        presentRate.toFixed(2), d.lateRate.toFixed(2), d.absRate.toFixed(2)
      ]);
    }
  });

  if (momData.length > 0) {
    dashboardSheet.getRange(currentRow, 1, momData.length, momHeaders.length).setValues(momData);
    currentRow += momData.length;
  }
  const momEndRow = currentRow - 1;
  currentRow += 2;

  // --- Section 2: Class Average Attendance Rate ---
  dashboardSheet.getRange(currentRow, 1).setValue("2. Class Average Attendance Rate (%)").setFontWeight("bold").setFontSize(12);
  currentRow++;

  const allClasses = new Set();
  Object.values(monthClassMap).forEach(classObj => Object.keys(classObj).forEach(c => allClasses.add(c)));
  const classArray = Array.from(allClasses);

  const classHeaders = ['Month', ...classArray];
  dashboardSheet.getRange(currentRow, 1, 1, classHeaders.length).setValues([classHeaders]).setFontWeight("bold");
  currentRow++;

  const classStartRow = currentRow;
  const classData = [];

  monthSheets.forEach(sheet => {
    const month = sheet.getName();
    if (monthClassMap[month]) {
      const rowData = [month];
      classArray.forEach(cls => {
        rowData.push(monthClassMap[month][cls] ? monthClassMap[month][cls].toFixed(2) : "");
      });
      classData.push(rowData);
    }
  });

  if (classData.length > 0) {
    dashboardSheet.getRange(currentRow, 1, classData.length, classHeaders.length).setValues(classData);
    currentRow += classData.length;
  }
  const classEndRow = currentRow - 1;
  currentRow += 2;

  // --- Section 3: Student Level Analysis ---
  dashboardSheet.getRange(currentRow, 1).setValue("3. Student Level Analysis (Yearly)").setFontWeight("bold").setFontSize(12);
  currentRow++;

  const studentHeaders = ['Student ID', 'Name', 'Class', 'Total Present', 'Total Late', 'Total Absent', 'Total Tracked', 'Attendance Rate (%)', '<40% (Red)', '40-75% (Yellow)', '>75% (Green)'];
  dashboardSheet.getRange(currentRow, 1, 1, studentHeaders.length).setValues([studentHeaders]).setFontWeight("bold");
  currentRow++;

  // Sort LOWEST attendance % first — students most at risk surface at the top.
  const sortedStudents = Object.keys(studentMap).sort((a, b) => {
    const rateA = studentMap[a].total > 0 ? (studentMap[a].present + studentMap[a].late) / studentMap[a].total : 0;
    const rateB = studentMap[b].total > 0 ? (studentMap[b].present + studentMap[b].late) / studentMap[b].total : 0;
    return rateA - rateB;
  });

  const studentStartRow = currentRow;
  const studentData = [];

  sortedStudents.forEach(id => {
    const s = studentMap[id];
    const rateVal = s.total > 0 ? ((s.present + s.late) / s.total * 100) : 0;
    const rateStr = rateVal.toFixed(2);

    let red = "", yellow = "", green = "";
    if (rateVal < 40) red = rateVal;
    else if (rateVal <= 75) yellow = rateVal;
    else green = rateVal;

    studentData.push([
      id, s.name, s.className, s.present, s.late, s.absent, s.total, rateStr, red, yellow, green
    ]);
  });

  if (studentData.length > 0) {
    dashboardSheet.getRange(currentRow, 1, studentData.length, studentHeaders.length).setValues(studentData);
    currentRow += studentData.length;
  }
  const studentEndRow = currentRow - 1;

  // --- Section 4: Chronic Attendance Risks (same signal as Weekly Report) ---
  // Reuses generateAlertBlocks so this stays identical to the Friday stakeholder
  // email: students with >= ABSENT_THRESHOLD_DAYS absences in the last
  // ABSENT_LOOKBACK_DAYS days, OR >= LATE_THRESHOLD_DAYS lates in the last
  // LATE_LOOKBACK_DAYS days. Evaluated against the CURRENT month tab (the alert
  // scan walks back into prior months on its own when the lookback needs it).
  currentRow += 2;
  dashboardSheet.getRange(currentRow, 1).setValue("4. Chronic Attendance Risks (Last "
    + Math.max(ABSENT_LOOKBACK_DAYS, LATE_LOOKBACK_DAYS) + " Days)").setFontWeight("bold").setFontSize(12);
  currentRow++;

  const riskHeaders = ['Student', 'Absences', 'Lates', 'Total Flags'];
  dashboardSheet.getRange(currentRow, 1, 1, riskHeaders.length).setValues([riskHeaders]).setFontWeight("bold");
  currentRow++;

  const riskStartRow = currentRow;
  const riskData = [];

  const currentMonthName = new Date().toLocaleString('en-US', { month: 'long' });
  const currentMonthSheet = ss.getSheetByName(currentMonthName);
  if (currentMonthSheet && currentMonthSheet.getLastRow() > 1) {
    const nameRange = currentMonthSheet.getRange(2, 3, currentMonthSheet.getLastRow() - 1, 1).getValues();
    const riskStudentNames = [];
    nameRange.forEach(function(r) {
      const nm = r[0].toString().trim();
      if (nm !== '' && nm.toUpperCase().indexOf('CLASS AVERAGE') === -1 && nm.toUpperCase().indexOf('STATUS') === -1) {
        riskStudentNames.push(nm);
      }
    });

    if (riskStudentNames.length > 0) {
      const alerts = generateAlertBlocks(ss, currentMonthSheet, riskStudentNames, new Date());
      // Already sorted most-flags-first inside generateAlertBlocks.
      (alerts.stakeholderData || []).forEach(function(d) {
        riskData.push([d.name, d.absences, d.lates, d.total]);
      });
    }
  }

  if (riskData.length > 0) {
    dashboardSheet.getRange(currentRow, 1, riskData.length, riskHeaders.length).setValues(riskData);
    currentRow += riskData.length;
  } else {
    dashboardSheet.getRange(currentRow, 1).setValue("No chronic risks flagged.").setFontColor("#718096");
    currentRow++;
  }
  const riskEndRow = currentRow - 1;

  dashboardSheet.hideColumns(9, 3);

  // 4. GENERATE CHARTS
  const standardWidth = 600;
  const standardHeight = 400;
  const colLeft = 13;
  const rowTop = 2;
  const rowBottom = 23;

  if (momEndRow >= momStartRow) {
    // Stacked % column chart. Series order (bottom->top) follows the column
    // order: Present % (col 9), Late % (col 10), Absent % (col 11). Colors map
    // to that same order: green (Present), orange (Late), red (Absent).
    const momRows = momEndRow - momStartRow + 2;  // include header row
    const momChart = dashboardSheet.newChart()
      .asColumnChart()
      .addRange(dashboardSheet.getRange(momStartRow - 1, 1, momRows, 1))   // Month (x-axis)
      .addRange(dashboardSheet.getRange(momStartRow - 1, 9, momRows, 3))   // Present %, Late %, Absent %
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setNumHeaders(1)
      .setOption('title', 'Month-over-Month Attendance Analysis (%)')
      .setOption('isStacked', true)
      .setOption('colors', ['#2CA02C', '#FF7F0E', '#D62728'])
      .setOption('vAxis.title', 'Percentage (%)')
      .setOption('hAxis.title', 'Month')
      .setOption('width', standardWidth)
      .setOption('height', standardHeight)
      .setPosition(rowTop, colLeft, 0, 0)
      .build();
    dashboardSheet.insertChart(momChart);
  }

  if (studentEndRow >= studentStartRow) {
    const numStudents = studentEndRow - studentStartRow + 1;
    const dynamicHeight = Math.max(400, (numStudents * 25) + 100);

    const studentChart = dashboardSheet.newChart()
      .asBarChart()
      .addRange(dashboardSheet.getRange(studentStartRow - 1, 2, studentEndRow - studentStartRow + 2, 1))
      .addRange(dashboardSheet.getRange(studentStartRow - 1, 9, studentEndRow - studentStartRow + 2, 3))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setNumHeaders(1)
      .setOption('title', 'Yearly Student Attendance %')
      .setOption('hAxis.title', 'Attendance %')
      .setOption('isStacked', true)
      .setOption('colors', ['#D62728', '#FBBC04', '#2CA02C'])
      .setOption('height', dynamicHeight)
      .setOption('width', 700)
      .setOption('chartArea', {left: '20%', top: '10%', width: '75%', height: '80%'})
      .setPosition(rowBottom, colLeft, 0, 0)
      .build();
    dashboardSheet.insertChart(studentChart);
  }

  if (riskData.length > 0 && riskEndRow >= riskStartRow) {
    // Stacked bar: Absences (col 2) + Lates (col 3) per flagged student.
    const riskRows = riskEndRow - riskStartRow + 2;  // include header row
    const riskChartHeight = Math.max(300, (riskData.length * 30) + 100);
    const riskChart = dashboardSheet.newChart()
      .asBarChart()
      .addRange(dashboardSheet.getRange(riskStartRow - 1, 1, riskRows, 1))   // Student (y-axis)
      .addRange(dashboardSheet.getRange(riskStartRow - 1, 2, riskRows, 2))   // Absences, Lates
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setNumHeaders(1)
      .setOption('title', 'Chronic Attendance Risks (Absences & Lates)')
      .setOption('isStacked', true)
      .setOption('colors', ['#D62728', '#FBBC04'])
      .setOption('hAxis.title', 'Total Days Flagged')
      .setOption('width', 700)
      .setOption('height', riskChartHeight)
      .setOption('chartArea', {left: '25%', top: '10%', width: '70%', height: '80%'})
      .setPosition(rowBottom + 22, colLeft, 0, 0)
      .build();
    dashboardSheet.insertChart(riskChart);
  }

  dashboardSheet.autoResizeColumns(1, 8);
  Logger.log("Successfully updated 'Analysis_Dashboard' for: " + ss.getName());
}

