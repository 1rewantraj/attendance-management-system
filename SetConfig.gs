// =========================================================================
// FILE 1: Config.gs (Global Variables)
// =========================================================================

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

var ABSENT_THRESHOLD_DAYS = 3; 
var ABSENT_LOOKBACK_DAYS = 5; 
var LATE_THRESHOLD_DAYS = 3; 
var LATE_LOOKBACK_DAYS = 5; 

// Comma-separated recipients of the weekly stakeholder report. REPLACE with real emails.
var STAKEHOLDER_EMAILS = "stakeholder1@example.com,stakeholder2@example.com";

// Comma-separated Program Managers CC'd on every attendance-submission
// confirmation email (accept AND reject) so they have live visibility into
// what was recorded per class. Leave empty ("") to disable the CC. REPLACE.
var PROGRAM_MANAGER_EMAILS = "";

// On-Demand Form Defaults — defaults for manual_sendOnDemandForm. REPLACE as needed.
var ONDEMAND_CLASS_NUM = "5";
var ONDEMAND_SECTION = "A";
var ONDEMAND_DATE = ""; // Leave empty for today, or use DD-MM-YYYY format: "05-07-2026" for July 5, 2026
var ONDEMAND_TEACHER_NAME = "Teacher Name";
var ONDEMAND_TEACHER_EMAIL = "teacher@example.com";