// =========================================================================
// FILE 1: Config.gs (Global Variables)
// =========================================================================

// Folder Links — REPLACE these placeholders with your own Google Drive folder URLs.
// (Get a folder's link: open it in Drive → copy the address bar URL.)
var INPUT_FOLDER_LINK = "https://drive.google.com/drive/folders/YOUR_INPUT_FOLDER_ID";
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

// Adhoc (makeup) Form Overrides — defaults for manual_runAdhocForm. REPLACE as needed.
var ADHOC_CLASS_NUM = "5";
var ADHOC_SECTION = "A";
var ADHOC_DATE = "";
var ADHOC_TEACHER_NAME = "Teacher Name";
var ADHOC_TEACHER_EMAIL = "teacher@example.com";