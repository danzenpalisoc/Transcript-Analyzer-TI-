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
var ROSTER_SHEET_ID       = '1-UIO2kTsIdx5oru_uvg0q9rtNM-1Hb5e_R_YgISoSh4';
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

// External audit tracking spreadsheet (shared by user)
var AUDIT_TRACKING_SS_ID = '1wy65jF6bz15tIQs5uozxksA2z7yDwxWOszaOrVCGuaM';

// Role names in the Audit Tracking SS → Roster tab (Col B).
// Update these if the exact spelling in your sheet differs.
var QA_ROLE    = 'QA';
var QA_TL_ROLE = 'QA Team Leader';

// AT DATA GCP file — for VTID lookup
var AT_DATA_GCP_SS_ID    = '1EQ1QDTgXukygNlt6mOJ92jFWP5sc-X0AWq8wreVy3ls';
var AT_DATA_AGENT_COL    = 2;   // Column C (0-based) — Agent_Name
// AT_DATA_VTID_COL removed — lookupVTID() auto-detects column from headers

// FCR Dashboard Data file — third lookup source for Team Leader / Ops Manager
var FCR_DASHBOARD_SS_ID  = '1wy65jF6bz15tIQs5uozxksA2z7yDwxWOszaOrVCGuaM';

var TRANSCRIPTS_HEADERS = [
  'Audit Ref', 'Timestamp', 'SAP ID', 'Team Member', 'Team Leader',
  'Operations Manager', 'Line of Business', 'Locale', 'Observer Name',
  'Transcript Start Time', 'Customer BAN', 'Interaction ID',
  'Direction', 'Transcript Duration', 'Analysis Type', 'Transcript'
];

var ANALYSIS_HEADERS = [
  'Audit Ref', 'Timestamp', 'SAP ID', 'Team Member', 'Observer Name',
  'Analysis Type', 'AI Analysis Result'
];

var CACHE_HEADERS = [
  'Interaction ID', 'Analysis Type', 'Timestamp', 'HTML Result'
];

// Dashboard_Data headers — structured for charting and filtering
var DASHBOARD_HEADERS = [
  'Audit Ref', 'Timestamp', 'Interaction ID', 'SAP ID', 'Team Member',
  'Team Leader', 'Operations Manager', 'Line of Business', 'Locale', 'VTID',
  'Observer Name', 'Department', 'Direction', 'Transcript Duration',
  'Analysis Type',
  // ── Raw extracted fields ──
  'Call Reason', 'Call Summary', 'Overall Opportunities',
  'Recommendations', 'Critical Flags', 'Repeat Projection %',
  'Issue Resolved', 'Transfer Occurred',
  // ── Structured RCA fields (AI-enriched) ──
  'Call Driver',          // Short label: e.g. "Service Move Request"
  'RCA Category',         // Agent Controllable | Process/Policy | Customer Driven | Transfer Issue
  'RCA Sub-Parameter',    // Specific missed step or gap
  'Top Opportunity',      // #1 actionable coaching point (short)
  'Product Opportunity',  // Missed upsell/product opportunity (Sales only)
  // ── Status ──
  'PDF Email Link', 'Email Status'
];

// Audit_Log headers — permanent archive of every submitted evaluation
var AUDIT_LOG_HEADERS = [
  'Audit Ref', 'Submitted At', 'Interaction ID', 'SAP ID', 'Team Member',
  'Team Leader', 'Operations Manager', 'Line of Business', 'Locale', 'VTID',
  'Observer Name', 'Analysis Type', 'Direction', 'Duration',
  'Repeat Projection %', 'Issue Resolved', 'Transfer Occurred',
  'Email Status', 'Recipients'
];

