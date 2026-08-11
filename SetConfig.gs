// Folder Links — REPLACE these placeholders with your own Google Drive folder URLs.
// (Get a folder's link: open it in Drive → copy the address bar URL.)
var STUDENT_CLASS_LIST_FOLDER_LINK = "https://drive.google.com/drive/folders/YOUR_STUDENT_CLASS_LIST_FOLDER_ID";
var ATTENDANCE_SHEETS_FOLDER_LINK = "https://drive.google.com/drive/folders/YOUR_OUTPUT_FOLDER_ID";
var CONFIG_FOLDER_LINK = "https://drive.google.com/drive/folders/YOUR_CONFIG_FOLDER_ID";

// File Names
var MAPPING_FILE_NAME = "TeacherClassMapping"; 
var HOLIDAY_FILE_NAME = "publicHoliday";  

// Academic Settings
var ACADEMIC_YEAR = "2026-2027"; 
var START_MONTH = 6; // 6 = June
var END_MONTH = 4;   // 4 = April
// In-sheet demo charts (buildVisualizations) are deprecated in favor of the
// Analysis_Dashboard tab. Kept defined (false) because the config validators
// still log this flag — leaving it undefined throws a ReferenceError.
var ADD_VISUALISATIONS = false;
var MAX_EXECUTION_TIME = 6 * 60 * 1000; 

// Alert Settings
var CONSECUTIVE_ABSENT_THRESHOLD_DAYS = 3; 
var CONSECUTIVE_LATE_THRESHOLD_DAYS = 3; 
var CONSECUTIVE_MIXED_THRESHOLD_DAYS = 4; 
var ALLOWED_PRESENT_SKIPS = 2; 

// Chronic attendance risk (weekly stakeholder report + dashboard): flag a
// student whose combined Absent + Late count reaches the threshold within
// the last N working days (weekends/holidays/non-instructional days are
// already excluded from the lookback before this count runs).
var CHRONIC_MISHAP_THRESHOLD_DAYS = 3;
var CHRONIC_MISHAP_LOOKBACK_DAYS = 5;

// Comma-separated recipients of the consolidated weekly report (automated_sendWeeklyReport).
// Every class's Program Manager is added automatically on top of this list — REPLACE with
// real emails for any additional stakeholders who aren't already a Program Manager.
var STAKEHOLDER_EMAILS = "stakeholder1@example.com,stakeholder2@example.com";

// Restrict manual_updateSheets to a single class-section instead of scanning every
// roster file. Set both to run just that class (e.g. classNum "5", section "A").
// Leave either blank ("") to update every class as before.
var UPDATE_SHEETS_CLASS_FILTER = "";
var UPDATE_SHEETS_SECTION_FILTER = "";

// On-Demand Form Defaults — defaults for manual_sendOnDemandForm. REPLACE as needed.
var ONDEMAND_CLASS_NUM = "5";
var ONDEMAND_SECTION = "A";
var ONDEMAND_DATE = ""; // Leave empty for today, or use DD-MM-YYYY format: "05-07-2026" for July 5, 2026
var ONDEMAND_TEACHER_NAME = "Teacher Name";
var ONDEMAND_TEACHER_EMAIL = "teacher@example.com";