/**
 * Config.gs
 * Central configuration. Sensitive values stored in Script Properties.
 */

var FUELIX_CONFIG = (function () {
  var props = PropertiesService.getScriptProperties();
  return {
    baseUrl: props.getProperty('ANTHROPIC_BASE_URL')   || 'https://api.fuelix.ai',
    model:   props.getProperty('ANTHROPIC_MODEL')      || 'claude-sonnet-4',
    apiKey:  props.getProperty('ANTHROPIC_AUTH_TOKEN') || ''
  };
})();

var SPREADSHEET_NAME      = 'NH FCR, Transfer or Sales Call Analyzer';
var MAIN_SPREADSHEET_ID   = '1wy65jF6bz15tIQs5uozxksA2z7yDwxWOszaOrVCGuaM';
var ADMIN_DOMAIN          = 'telus.com';

// ── Audit reference prefix, chosen by WHICH script is running ─────────────────────────────────
// The NH and TI Analyzers are separate Apps Script projects running this same
// code against ONE spreadsheet, and each keeps its own daily counter that
// restarts at 1. With the same prefix both hand out NHA-YYYYMMDD-0001, -0002,
// ... and two Audit_Log rows end up sharing one reference, so an email link
// can open the other script's agent. A different prefix per script makes the
// references unique. Keyed by script ID so this file stays identical in both
// projects. Any script not listed here (a copy, a dev project) uses NHA.
var AUDIT_REF_PREFIX_DEFAULT = 'NHA';
var AUDIT_REF_PREFIX_BY_SCRIPT_ID = {
  '1P_dbehoDRZsqG22se2cYcaD1aKBxTuX3g4G2-moBzKLg7Sc5GLtPnAca': 'NHA',  // NH Analyzer
  '1KfR8crJ3ep5SqQsrUEFuksplZeacsMFIdW6yO56fxRLPxTppiOtr4Yww': 'TIA'   // TI Analyzer
};
function getAuditRefPrefix_() {
  try {
    return AUDIT_REF_PREFIX_BY_SCRIPT_ID[ScriptApp.getScriptId()] || AUDIT_REF_PREFIX_DEFAULT;
  } catch (e) {
    return AUDIT_REF_PREFIX_DEFAULT;
  }
}
var ROSTER_SHEET_ID       = '1-UIO2kTsIdx5oru_uvg0q9rtNM-1Hb5e_R_YgISoSh4';
var GLOBAL_ROSTER_SS_ID   = '18bB1ZIb_Qkn8Kt1UwVtswYhFqQ5MdWXZolei2MoT6Yg';
var REPEATS_FOLDER_ID     = '13gcW62ubbfZ-Zk-zxTozjcIvs5_55bD1';
var SALES_FOLDER_ID       = '1zvFsLLse7X1UtfLtZ-ggSc5CNLd0ZrE0';
var PDF_CACHE_KEY_REPEATS = 'pdf_knowledge_repeats';
var PDF_CACHE_KEY_SALES   = 'pdf_knowledge_sales';
var PDF_CACHE_SECONDS     = 6 * 60 * 60;

// Roster column indices (0-based) — confirmed by user
// A=0:sap_ID  C=2:Agent_Name  F=5:Agent Role (Scorecard)→LOB
// I=8:Team_Mgr_Name  K=10:Ops_Mgr_Name  S=18:Location→Locale
var ROSTER_COL_SAP_ID      = 0;   // A - sap_ID
var ROSTER_COL_AGENT_NAME  = 2;   // C - Agent_Name
var ROSTER_COL_DEPT_CD     = 3;   // D - Dept_Cd
var ROSTER_COL_DOMAIN_NAME = 5;   // F - Agent Role (Scorecard) → Line of Business
var ROSTER_COL_TEAM_MGR    = 8;   // I - Team_Mgr_Name → Team Leader
var ROSTER_COL_OPS_MGR     = 10;  // K - Ops_Mgr_Name → Operations Manager
var ROSTER_COL_LOCALE      = 18;  // S - Location → Locale (Site)

// AT Data GCP — Roster tab, Col H=Agent name, Col F=VTID
// A=0 … F=5:VTID … H=7:Agent

var TRANSCRIPTS_SHEET    = 'Transcripts';
var ANALYSIS_SHEET       = 'Analysis';
var CACHE_SHEET          = 'Cache';
var DASHBOARD_DATA_SHEET = 'Dashboard_Data';
var AUDIT_LOG_SHEET      = 'Audit_Log';
var AI_ANALYTICS_SHEET   = 'AI_Analytics';
var ERROR_LOG_SHEET      = 'Error_Log';

// Error_Log headers — every submitTranscript() failure, written directly into
// the main spreadsheet so the admin can see raw error text without needing
// Stackdriver/GCP log access (which requires a non-default GCP project).
var ERROR_LOG_HEADERS = [
  'Timestamp', 'Function', 'Observer', 'SAP ID', 'Interaction ID',
  'Error Message', 'Stack (truncated)'
];

// External audit tracking spreadsheet (shared by user)
var AUDIT_TRACKING_SS_ID = '1wy65jF6bz15tIQs5uozxksA2z7yDwxWOszaOrVCGuaM';

// Trainee roster — maps SAP ID → Facilitator (Trainer) name + username
var TRAINEE_ROSTER_SS_ID = '1EQ1QDTgXukygNlt6mOJ92jFWP5sc-X0AWq8wreVy3ls';
// Trainer email lookup — maps Trainer Name → Trainer Email + Supervisor Email
var TRAINER_LOOKUP_SS_ID = '1nqHtScRcW7xdMOi8OPcpgxEmm4DBWuopoqp030DtBXk';

// Role names in the Audit Tracking SS → Roster tab (Col B).
// Update these if the exact spelling in your sheet differs.
var QA_ROLE      = 'QA';
var QA_TL_ROLE   = 'QA Team Leader';
var TRAINER_ROLE = 'Trainer';

// Admin/Dev usernames (part before @) — controls Live/Test toggle visibility
var ADMIN_USERNAMES = [
  'danzen.palisoc',
  'lobell.bantaya',
  'carlos.maldonado',
  'robin.joseph'
];

// Domains an admin may sign in from. The owner's account is on
// telusinternational.com, so listing only ADMIN_DOMAIN would lock them out.
var ADMIN_EMAIL_DOMAINS = ['telus.com', 'telusinternational.com'];

// True only when the person actually running this call is on ADMIN_USERNAMES.
//
// Any public function can be called by any signed-in user from the browser
// console, so the repair / backfill / diagnose functions must check this first.
// Deliberately NOT a fallback to Session.getEffectiveUser(): the web app runs as
// "Execute as: Me", so the effective user is the OWNER for every caller, and a
// fallback would make everyone an admin the moment their own email is hidden. A
// blank or unreadable identity is therefore refused.
function isAdminCaller_() {
  var email = '';
  try { email = (Session.getActiveUser().getEmail() || '').toString().toLowerCase().trim(); } catch (e) { return false; }
  var at = email.lastIndexOf('@');
  if (at < 1) return false;
  return ADMIN_USERNAMES.indexOf(email.slice(0, at)) !== -1 &&
         ADMIN_EMAIL_DOMAINS.indexOf(email.slice(at + 1)) !== -1;
}

// First line of every maintenance function. Throws so a google.script.run
// failure handler (or the editor's execution log) shows why nothing happened.
function requireAdmin_() {
  if (!isAdminCaller_()) {
    throw new Error('This function is for the QA and LS admin team only.');
  }
}

// AT DATA GCP file — for VTID lookup
var AT_DATA_GCP_SS_ID    = '1EQ1QDTgXukygNlt6mOJ92jFWP5sc-X0AWq8wreVy3ls';
var AT_DATA_AGENT_COL    = 2;   // Column C (0-based) — Agent_Name
// AT_DATA_VTID_COL removed — lookupVTID_() auto-detects column from headers

// FCR Dashboard Data file — third lookup source for Team Leader / Ops Manager
var FCR_DASHBOARD_SS_ID  = '1wy65jF6bz15tIQs5uozxksA2z7yDwxWOszaOrVCGuaM';

// Users sheet — canonical user directory (Email Address, Name, Role, Status…)
// Used by resolveObserver_() to identify QA analysts by their logged-in email.
var USERS_SS_ID = '1nqHtScRcW7xdMOi8OPcpgxEmm4DBWuopoqp030DtBXk';
var USERS_TAB   = 'Users';

var TRANSCRIPTS_HEADERS = [
  'Audit Ref', 'Timestamp', 'SAP ID', 'Team Member', 'Team Leader',
  'Operations Manager', 'Line of Business', 'Locale', 'Observer Name',
  'Transcript Start Time', 'Customer BAN', 'Interaction ID',
  'Direction', 'Transcript Duration', 'Analysis Type', 'Transcript',
  'Agent Email'
];

var ANALYSIS_HEADERS = [
  'Audit Ref', 'Timestamp', 'SAP ID', 'Team Member', 'Observer Name',
  'Analysis Type', 'AI Analysis Result'
];

var CACHE_HEADERS = [
  'Interaction ID', 'Analysis Type', 'Timestamp', 'HTML Result', 'SAP ID'
];

// Dashboard_Data headers — structured for charting and filtering
var DASHBOARD_HEADERS = [
  'Audit Ref', 'Timestamp', 'Interaction ID', 'SAP ID', 'Team Member',
  'Team Leader', 'Operations Manager', 'Line of Business', 'Locale', 'VTID',
  'Observer Name', 'Department', 'Direction', 'Transcript Duration',
  'Analysis Type',
  // ── Raw extracted fields ──
  'Call Reason', 'Call Summary', 'Overall Opportunities',
  'SMART Recommendation', 'Critical Flags', 'Repeat Projection %',
  'Issue Resolved', 'Transfer Occurred',
  // ── Structured RCA fields (AI-enriched) ──
  'Call Driver',              // Short label: e.g. "Service Move Request"
  'RCA Category',             // Agent Controllable | Process/Policy | Customer Driven | Transfer Issue
  'RCA Sub-Parameter',        // Specific missed step or gap
  'Top SMART Opportunity',    // #1 coaching point in SMART format (short)
  'Product Opportunity',      // Missed upsell/product opportunity (Sales only)
  'Sales Attempted',          // Yes | No | N/A — did agent make any offer/upsell (Sales only)
  // ── SMART Coaching Breakdown (from AI) ──
  'SMART S (Specific)',       // Exact behavior to change
  'SMART M (Measurable)',     // How success will be measured
  'SMART A (Attainable)',     // Achievable target
  'SMART R (Realistic)',      // Why this is realistic for this agent now
  'SMART T (Time-bound)',     // Timeline for achieving the goal
  // ── Status ──
  'PDF Email Link', 'Email Status', 'Agent Email',
  'Auditor Email'             // Who submitted the audit (server-side session email, else the Users sheet)
];

// Audit_Log headers — permanent archive of every submitted evaluation
// Targets on the agent scorecard slides (ScorecardMetrics.gs). direction 'max' = lower is better (at or below the
// target passes), 'min' = higher is better. Chosen by the admin on 2026-10-13. The number is compared as it is
// shown on the slide (percentages whole, flags to one decimal).
// A short, stable key for a cache entry (an address or an id), so a long or look-alike value cannot collide after being cut.
function _hashKey_(s) {
  var h = 5381, str = String(s);
  for (var i = 0; i < str.length; i++) h = ((h * 33) ^ str.charCodeAt(i)) >>> 0;
  var h2 = 52711;
  for (var j = str.length - 1; j >= 0; j--) h2 = ((h2 * 31) ^ str.charCodeAt(j)) >>> 0;
  return h.toString(36) + h2.toString(36) + str.length.toString(36);
}

// True when this person may do something again this hour (and counts it). Fails OPEN when the cache is unavailable, with a log line.
function _hourlyTake_(prefix, who, max) {
  try {
    var cache = CacheService.getScriptCache(), key = prefix + _hashKey_(String(who || 'unknown').toLowerCase());
    var n = parseInt(cache.get(key) || '0', 10) || 0;
    if (n >= max) return false;
    cache.put(key, String(n + 1), 3600);
  } catch (e) {
    Logger.log('_hourlyTake_(' + prefix + '): cache unavailable, limit not applied: ' + e);
  }
  return true;
}

// The folder with this name that the app OWNER owns. getFoldersByName also returns folders other people shared with the owner,
// so anyone could plant a folder with the same name and have decks, exports or PDFs filed in (or cleaned out of) their folder.
// When the owner cannot be read (the tests) the first match is used, as before.
function _ownedFolder_(name) {
  var me = '';
  try { me = String(Session.getEffectiveUser().getEmail() || '').toLowerCase(); } catch (e) {}
  var it = DriveApp.getFoldersByName(name), first = null;
  while (it.hasNext()) {
    var f = it.next();
    if (!first) first = f;
    if (!me) return f;
    var owner = '';
    try { owner = String((f.getOwner() && f.getOwner().getEmail()) || '').toLowerCase(); } catch (e2) {}
    if (owner === me) return f;
  }
  return DriveApp.createFolder(name);
}

var SCORECARD_THRESHOLDS = {
  repeatRisk:   { label: 'Avg Repeat Risk',     target: 30, direction: 'max' },   // percent
  resolution:   { label: 'Resolution rate',     target: 80, direction: 'min' },   // percent of audits with a verdict
  flags:        { label: 'Critical flags',      target: 0,  direction: 'max' },   // flags per audit
  transfer:     { label: 'Transfer rate',       target: 20, direction: 'max' },   // percent of Repeats audits
  salesAttempt: { label: 'Sales attempted',     target: 80, direction: 'min' }    // percent of Sales audits
};

var AUDIT_LOG_HEADERS = [
  'Audit Ref', 'Submitted At', 'Interaction ID', 'SAP ID', 'Team Member',
  'Team Leader', 'Operations Manager', 'Line of Business', 'Locale', 'VTID',
  'Observer Name', 'Analysis Type', 'Direction', 'Duration',
  'Repeat Projection %', 'Issue Resolved', 'Transfer Occurred',
  'Email Status', 'Recipients', 'Agent Email'
];

