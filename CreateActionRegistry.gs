/**
 * CreateActionRegistry.gs
 *
 * Run once from your standalone GAS project.
 * Creates "NH Analyzer — Action Registry" Google Sheet
 * with full commit history from Day 1 (2026-08-04) to present.
 *
 * After running, check Execution Log for the sheet URL.
 */
function createActionRegistry() {

  // ── 1. Create spreadsheet ────────────────────────────────────────────────
  var ss = SpreadsheetApp.create('NH Analyzer — Action Registry');
  var sh = ss.getActiveSheet();
  sh.setName('Registry');

  // ── 2. Headers ───────────────────────────────────────────────────────────
  var headers = [
    '#',
    'Date',
    'Type',
    'Scope',
    'Issue / Request',
    'Resolution / Changes Made',
    'Files Affected',
    'Author',
    'Status'
  ];
  var numCols = headers.length;
  sh.getRange(1, 1, 1, numCols).setValues([headers]);

  var hdrRange = sh.getRange(1, 1, 1, numCols);
  hdrRange.setBackground('#3B0764');
  hdrRange.setFontColor('#FFFFFF');
  hdrRange.setFontWeight('bold');
  hdrRange.setFontSize(11);
  hdrRange.setVerticalAlignment('middle');
  hdrRange.setHorizontalAlignment('center');
  sh.setRowHeight(1, 36);
  sh.setFrozenRows(1);

  // ── 3. Column widths ─────────────────────────────────────────────────────
  sh.setColumnWidth(1, 45);   // #
  sh.setColumnWidth(2, 105);  // Date
  sh.setColumnWidth(3, 125);  // Type
  sh.setColumnWidth(4, 80);   // Scope
  sh.setColumnWidth(5, 300);  // Issue / Request
  sh.setColumnWidth(6, 420);  // Resolution
  sh.setColumnWidth(7, 200);  // Files Affected
  sh.setColumnWidth(8, 130);  // Author
  sh.setColumnWidth(9, 85);   // Status

  // ── 4. Full history data ─────────────────────────────────────────────────
  // Col order: #, Date, Type, Scope, Issue, Resolution, Files, Author, Status
  var rows = [
    [
      '001',
      '2026-08-04',
      'Feature',
      'NH',
      'Project creation requested — NH FCR, Transfer or Sales Call Analyzer',
      'Initial commit: GAS project scaffolded with index.html, Code.gs, Config.gs, AIService.gs, SheetService.gs.',
      'All files',
      'Danzen / Lobell',
      'Done'
    ],
    [
      '002',
      '2026-08-04',
      'Feature',
      'NH',
      'Full project structure needed — transcript paste and AI analysis flow',
      'Structured NH FCR, Transfer or Sales Call Analyzer with complete initial codebase; transcript paste and AI analysis flow established.',
      'All files',
      'Danzen / Lobell',
      'Done'
    ],
    [
      '003',
      '2026-08-04',
      'Maintenance',
      'NH',
      'Remote branch sync required after initial commits from both contributors',
      'Merged main branch from GitHub to align local and remote state after initial dual commits.',
      'All files',
      'Danzen / Lobell',
      'Done'
    ],
    [
      '004',
      '2026-08-05',
      'Feature',
      'NH',
      'No caching, no AI analytics tab, no admin notifications, no RCA enrichment',
      'Major update: added CacheService caching for AI results, AI Analytics tab, admin notification emails, and RCA Category enrichment on Dashboard_Data.',
      'Code.gs, AIService.gs, Config.gs',
      'Danzen',
      'Done'
    ],
    [
      '005',
      '2026-08-10',
      'Bug Fix',
      'NH',
      '15 bugs found during code review — race conditions, incorrect column writes, missing error guards',
      'Resolved 15 bugs: cache invalidation issues, incorrect column writes, race conditions in sheet operations, missing null guards.',
      'Code.gs, AIService.gs',
      'Danzen',
      'Done'
    ],
    [
      '006',
      '2026-08-12',
      'Feature',
      'NH',
      'No auto-send of audit email after form submission; GAS changes not synced to GitHub',
      'Built auto-send audit email on form submission; synced GAS changes to GitHub repository.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '007',
      '2026-08-12',
      'Feature',
      'NH',
      'SMART coaching output not structured; error rate tracking missing; caching insufficient',
      'Added SMART coaching format (Specific, Measurable, Attainable, Realistic, Time-bound). Fixed error rate calculation. Extended CacheService caching.',
      'Code.gs, AIService.gs',
      'Danzen',
      'Done'
    ],
    [
      '008',
      '2026-08-12',
      'Maintenance',
      'NH',
      'BOM characters and trailing newlines in GAS files after clasp pull causing sync noise',
      'Removed BOM markers and trailing newlines from all .gs and .html files after clasp pull.',
      'All files',
      'Danzen',
      'Done'
    ],
    [
      '009',
      '2026-08-13',
      'Feature',
      'NH',
      'SMART table missing R (Realistic) column in coaching output',
      'Added R (Realistic) column to all SMART coaching tables in AI output.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '010',
      '2026-08-13',
      'Feature',
      'NH',
      'Spreadsheet missing SMART columns; sheet headers outdated after SMART format change',
      'Added SMART coaching columns to spreadsheet; renamed headers to match new SMART structure.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '011',
      '2026-08-13',
      'Feature',
      'NH',
      'Audit email design was plain and did not match New Hire Evaluation branding',
      'Redesigned agent evaluation email to match New Hire Evaluation style and branding.',
      'Code.gs, EvalView.html',
      'Danzen',
      'Done'
    ],
    [
      '012',
      '2026-08-13',
      'Bug Fix',
      'NH',
      'Email subject line incorrect — needed to reflect feedback format, not analyzer tool name',
      'Changed email subject line to "Real time Analyzer" as interim fix.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '013',
      '2026-08-13',
      'Bug Fix',
      'NH',
      'Email subject still not matching standard feedback format after first fix',
      'Updated email subject to "Real Time Feedback" to match standard format.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '014',
      '2026-08-13',
      'Feature',
      'NH',
      'No hosted evaluation view — audit email button linked nowhere useful',
      'Added hosted EvalView page; email button now links to web-hosted evaluation page.',
      'Code.gs, EvalView.html',
      'Danzen',
      'Done'
    ],
    [
      '015',
      '2026-08-13',
      'Feature',
      'NH',
      'Admin notification email was plain; Interaction ID overflowing its layout cell',
      'Redesigned admin notification email; fixed Interaction ID overflow with truncation.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '016',
      '2026-08-13',
      'Bug Fix',
      'NH',
      'SMART table columns squished in PDF — wasting horizontal space',
      'Maximized PDF space by fixing squished SMART table column widths.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '017',
      '2026-08-13',
      'Feature',
      'NH',
      'Full email flow needed overhaul; EvalView needed redesign to match Observation layout',
      'Complete email flow overhaul; EvalView redesigned to match New Hire Evaluation (Observation) layout.',
      'Code.gs, EvalView.html',
      'Danzen',
      'Done'
    ],
    [
      '018',
      '2026-08-13',
      'Performance',
      'NH',
      'PDF wasting space with narrow tables and excessive margins',
      'Maximized PDF page usage — full-width tables, tighter margins, reduced padding.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '019',
      '2026-08-13',
      'Performance',
      'NH',
      'Overall app performance slow — too many redundant calls, slow sheet reads',
      'Comprehensive performance overhaul: faster sheet reads, smarter caching, reduced redundant API calls.',
      'Code.gs, AIService.gs',
      'Danzen',
      'Done'
    ],
    [
      '020',
      '2026-08-13',
      'Maintenance',
      'NH',
      'RCA Category chart name was ambiguous — not descriptive of what it shows',
      'Renamed RCA Category chart to "Flags/Violation Breakdown" for clarity.',
      'Code.gs, Dashboard.html',
      'Danzen',
      'Done'
    ],
    [
      '021',
      '2026-08-14',
      'Feature',
      'NH',
      'SMART coaching columns missing from spreadsheet; backfill and enrich utilities needed',
      'Added SMART coaching columns to Dashboard_Data; added backfill and enrichment utilities for existing data.',
      'Code.gs, AIService.gs',
      'Danzen',
      'Done'
    ],
    [
      '022',
      '2026-08-14',
      'Maintenance',
      'NH',
      'Local and remote branches diverged after SMART columns work',
      'Merged remote and local commits: SMART columns, backfill/enrich, EvalView layout.',
      'Code.gs, AIService.gs',
      'Danzen',
      'Done'
    ],
    [
      '023',
      '2026-08-14',
      'Feature',
      'NH',
      'No way to safely test email sends without hitting live recipients',
      'Added Live/Test email mode toggle — visible only to Admin/Dev users.',
      'Code.gs, index.html',
      'Danzen',
      'Done'
    ],
    [
      '024',
      '2026-08-14',
      'Maintenance',
      'NH',
      'Remote branch had commits not in local after Live/Test toggle work',
      'Merged remote branch into local after Live/Test toggle implementation.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '025',
      '2026-08-14',
      'Bug Fix',
      'NH',
      'Live/Test toggle placed in wrong UI location — not visible in audit flow',
      'Moved Live/Test toggle to audit meta bar beside Save as Draft button.',
      'Code.gs, index.html',
      'Danzen',
      'Done'
    ],
    [
      '026',
      '2026-08-14',
      'Bug Fix',
      'NH',
      'isAdmin check failing when getEffectiveUser() returned null user object',
      'Used getEffectiveUser() with null fallback in getObserverInfo() for isAdmin check.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '027',
      '2026-08-14',
      'Bug Fix',
      'NH',
      'isAdmin check failing for @telusinternational.com users — only matched @telus.com domain',
      'Changed isAdmin to match by username part only, ignoring domain suffix.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '028',
      '2026-08-14',
      'Bug Fix',
      'NH',
      'userinfo.email OAuth scope missing; getObserverInfo() crashing with no error context',
      'Re-added userinfo.email scope; wrapped getObserverInfo() in per-step try-catch with Logger.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '029',
      '2026-08-14',
      'Bug Fix',
      'NH',
      'isAdmin check hitting external sheet on every call — slow and fragile under load',
      'Replaced dynamic sheet-based isAdmin check with hardcoded ADMIN_USERNAMES array in Config.gs.',
      'Config.gs, Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '030',
      '2026-08-14',
      'Maintenance',
      'NH',
      'Second remote-local merge needed after isAdmin hardcoding committed',
      'Merged branch "main" to sync hardcoded ADMIN_USERNAMES across local and remote.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '031',
      '2026-08-19',
      'Bug Fix',
      'NH',
      'Team member auto-fill from transcript not working after previous refactor',
      'Implemented lookupSapFromName() to restore auto-fill of Team Member field from transcript.',
      'Code.gs, AIService.gs',
      'Danzen',
      'Done'
    ],
    [
      '032',
      '2026-08-19',
      'Bug Fix',
      'NH',
      'Autofill triggered on participant name field instead of sapId field — wrong trigger condition',
      'Changed autofill trigger condition to check sapId value, not participant name.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '033',
      '2026-08-19',
      'Bug Fix',
      'NH',
      'Team Member field corrupting with wrong data due to column-shift in H&L overview',
      'Added guard against column-shift data corruption in Team Member field of H&L overview section.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '034',
      '2026-08-19',
      'Feature',
      'NH',
      'No AI-generated Highlights & Lowlights section on Overview tab',
      'Added AI H&L section on Overview tab via getOverviewHighlightsLowlights() function.',
      'Code.gs, AIService.gs',
      'Danzen',
      'Done'
    ],
    [
      '035',
      '2026-08-19',
      'Bug Fix',
      'NH',
      'Team Member column-shift guard missing from tmData loop — only patched H&L overview',
      'Added column-shift guard in tmData loop to cover additional corruption path.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '036',
      '2026-08-19',
      'Bug Fix',
      'NH',
      'userinfo.email scope and executeAs misaligned after merge — Live/Test toggle broken',
      'Restored userinfo.email scope and aligned executeAs to USER_ACCESSING to restore toggle.',
      'Code.gs, Config.gs',
      'Danzen',
      'Done'
    ],
    [
      '037',
      '2026-08-20',
      'Bug Fix',
      'NH',
      'Unhelpful error message shown when user has Viewer-only access to main spreadsheet',
      'Cleared cryptic error and displayed a friendly notice when Viewer-only access detected.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '038',
      '2026-08-20',
      'Enhancement',
      'NH',
      'Spreadsheet lookup vulnerable to wrong-file access; formula injection, XSS, and email spoofing risks',
      'Security hardening: guarded against wrong-file lookup, sanitized formula injection, escaped XSS output, blocked email spoofing via server-side resolution.',
      'Code.gs, AIService.gs',
      'Danzen',
      'Done'
    ],
    [
      '039',
      '2026-08-20',
      'Feature',
      'NH',
      'SAP ID lookup only checked internal roster — agents in Global Roster were missed',
      'Added Global Roster as Step 1.5 in SAP ID lookup chain; added Agent Email field from roster data.',
      'Code.gs, AIService.gs',
      'Danzen',
      'Done'
    ],
    [
      '040',
      '2026-08-22',
      'Bug Fix',
      'NH',
      'LOB and Locale not populating from Global Roster; multiple email-sending bugs',
      'Fixed LOB and Locale population from Global Roster lookup; resolved multiple email-sending bugs.',
      'Code.gs, AIService.gs',
      'Danzen',
      'Done'
    ],
    [
      '041',
      '2026-08-24',
      'Bug Fix',
      'NH',
      'RCA Category not populating on new submissions; enrichDashboardData backfill incomplete',
      'Fixed RCA Category population on every new submission; corrected backfill call sequence in enrichDashboardData.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '042',
      '2026-08-25',
      'Feature',
      'NH',
      'AI Analytics modal missing filters that main dashboard had; filters not applied on open',
      'Added full filter parity to AI Analytics modal; filters auto-apply when modal opens.',
      'Code.gs, Dashboard.html',
      'Danzen',
      'Done'
    ],
    [
      '043',
      '2026-08-25',
      'Bug Fix',
      'NH',
      'New anFilter keys accessed before anInitFilters runs — undefined crash',
      'Added undefined guard for new anFilter keys before anInitFilters initialization.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '044',
      '2026-08-25',
      'Feature',
      'NH',
      'No Sales Attempted tracking field; Sales tab donut chart missing data source',
      'Added Sales Attempted field (Yes/No/N-A options) as data source for Sales tab donut chart.',
      'Code.gs, index.html',
      'Danzen',
      'Done'
    ],
    [
      '045',
      '2026-08-25',
      'Feature',
      'NH',
      'Existing sheet rows had no Sales Attempted column — needed one-time upgrade path',
      'Added addSalesAttemptedColumn() one-time upgrade function to backfill existing rows.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '046',
      '2026-08-25',
      'Feature',
      'NH',
      'Repeats and Sales tabs had no analytics sections — only raw data tables',
      'Added analytics sections (summaries and charts) to both Repeats and Sales tabs.',
      'Code.gs, Dashboard.html',
      'Danzen',
      'Done'
    ],
    [
      '047',
      '2026-08-25',
      'Bug Fix',
      'NH',
      'AI H&L not populating on dashboard data load — only triggered on manual action',
      'Fixed renderTabAIHighlights call to run from dbApplyAll so H&L populates on every data load.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '048',
      '2026-08-25',
      'Bug Fix',
      'NH',
      'Sales Attempted donut chart showing blank with no message when data is missing',
      'Added fallback message to Sales Attempted donut when no data is available.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '049',
      '2026-08-25',
      'Maintenance',
      'NH',
      'Debug: needed to verify Sales Attempted yes/no/other/total counts during development',
      'Temporary debug: added yes/no/other/total counts in Sales Attempted section title.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '050',
      '2026-08-25',
      'Maintenance',
      'NH',
      'Debug: needed to see raw Sales Attempted value from first row to confirm field reading',
      'Temporary debug: showed first Sales row raw Sales Attempted value in section title.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '051',
      '2026-08-25',
      'Bug Fix',
      'NH',
      'readDashboardSheet using fixed DASHBOARD_HEADERS index positions — broke when columns added',
      'readDashboardSheet now reads actual sheet headers instead of fixed index — permanently fixes column mapping for any future columns.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '052',
      '2026-08-25',
      'Maintenance',
      'NH',
      'Debug: needed to inspect all keys near Sales Attempted in first row to trace undefined',
      'Temporary debug: logged all keys near Sales Attempted in first row object.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '053',
      '2026-08-25',
      'Bug Fix',
      'NH',
      'Stale 28-column data from old cache causing Sales Attempted undefined after column addition',
      'Bumped GAS cache key v1→v2 and localStorage key v3→v4 to force full cache invalidation.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '054',
      '2026-08-25',
      'Maintenance',
      'NH',
      'Debug title removed; testSalesAttempted() needed for future Sales Attempted diagnosis',
      'Removed debug title output; added testSalesAttempted() diagnostic helper function.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '055',
      '2026-08-25',
      'Bug Fix',
      'NH',
      'Sales Attempted donut and bar chart had incorrect or missing labels',
      'Fixed donut and bar chart labels for Sales Attempted section.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '056',
      '2026-08-25',
      'Bug Fix',
      'NH',
      'Dashboard filters not applying to Overview tab, AI H&L, or Sales Attempted donut',
      'Fixed filter application to cover Overview tab, AI H&L section, and Sales Attempted donut.',
      'Code.gs, Dashboard.html',
      'Danzen',
      'Done'
    ],
    [
      '057',
      '2026-08-25',
      'Bug Fix',
      'NH',
      'Content-hash cache key not used in getOverviewHighlightsLowlights — stale H&L returned',
      'Switched to content-hash cache key in getOverviewHighlightsLowlights to prevent stale results.',
      'Code.gs, AIService.gs',
      'Danzen',
      'Done'
    ],
    [
      '058',
      '2026-08-26',
      'Bug Fix',
      'NH',
      'EvalView link in audit emails loading wrong content — auditRef not embedded server-side',
      'Embedded auditRef server-side in EvalView link so email button always loads the correct evaluation.',
      'Code.gs, EvalView.html',
      'Danzen',
      'Done'
    ],
    [
      '059',
      '2026-08-26',
      'Bug Fix',
      'NH',
      'Eval link loading slowly — re-rendering on every open instead of reading from cache',
      'Fixed eval link to load instantly by reading pre-cached HTML from Cache sheet in getEvalViewData.',
      'Code.gs, EvalView.html',
      'Danzen',
      'Done'
    ],
    [
      '060',
      '2026-08-26',
      'Bug Fix',
      'NH',
      'getEvalViewData not reading full HTML directly from Cache sheet — partial content returned',
      'Fixed getEvalViewData to read complete HTML result directly from Cache sheet.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '061',
      '2026-08-26',
      'Bug Fix',
      'NH',
      'max_tokens set too low — AI evaluations truncating on long transcripts',
      'Raised max_tokens to 16000 to prevent AI evaluation truncation.',
      'AIService.gs',
      'Danzen',
      'Done'
    ],
    [
      '062',
      '2026-08-26',
      'Feature',
      'NH',
      'No way to diagnose or repair truncated Sales evaluations already written to the sheet',
      'Added diagnoseSalesEvaluations() and repairTruncatedSalesEvaluations() admin utility functions.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '063',
      '2026-08-26',
      'Bug Fix',
      'NH',
      'max_tokens 16000 still truncating some evaluations with many scoring parameters',
      'Increased max_tokens to 20000; added repairNextTruncated() for one-at-a-time repair.',
      'AIService.gs',
      'Danzen',
      'Done'
    ],
    [
      '064',
      '2026-08-26',
      'Maintenance',
      'NH',
      'Sales and Repeats AI prompts too long — approaching token context ceiling',
      'Shortened Sales and Repeats AI prompts to prevent hitting context ceiling.',
      'AIService.gs',
      'Danzen',
      'Done'
    ],
    [
      '065',
      '2026-08-26',
      'Bug Fix',
      'NH',
      'CacheService 48KB limit crashing when AI evaluation HTML exceeded it',
      'Implemented chunked Cache sheet storage for HTML results over 48KB.',
      'Code.gs, AIService.gs',
      'Danzen',
      'Done'
    ],
    [
      '066',
      '2026-08-26',
      'Bug Fix',
      'NH',
      'Old cache rows accumulating in Cache sheet; empty rows causing diagnosis script errors',
      'Fixed cache cleanup to delete old rows instead of clearing; skip empty rows in diagnosis.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '067',
      '2026-08-26',
      'Bug Fix',
      'NH',
      'executeAs set incorrectly — scripts running as submitting user instead of owner',
      'Set executeAs to USER_DEPLOYING so script runs as spreadsheet owner on all calls.',
      'Config.gs',
      'Danzen',
      'Done'
    ],
    [
      '068',
      '2026-08-27',
      'Bug Fix',
      'NH',
      'Edit preservation broken after result load; loading timer not showing; locale filter and RCA enrichment not running',
      'Fixed edit preservation after AI result load; fixed loading timer display; fixed locale filter logic; fixed RCA enrichment call sequence.',
      'Code.gs, index.html',
      'Danzen',
      'Done'
    ],
    [
      '069',
      '2026-08-28',
      'Feature',
      'NH',
      'No way to fix truncated Locale values already stored in Dashboard_Data',
      'Added fixTruncatedLocale() admin function to repair existing truncated Locale values.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '070',
      '2026-08-28',
      'Bug Fix',
      'NH',
      'fixTruncatedLocale only patching Dashboard_Data — Audit_Log also had truncated Locale values',
      'Extended fixTruncatedLocale to also patch truncated Locale values in Audit_Log.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '071',
      '2026-08-28',
      'Feature',
      'NH',
      'No translation support — non-English transcripts analyzed with incorrect context',
      'Added Translate to English toggle; transcripts auto-translated before AI analysis.',
      'Code.gs, index.html',
      'Danzen',
      'Done'
    ],
    [
      '072',
      '2026-08-28',
      'Feature',
      'NH',
      'No TTS (text-to-speech) for SMART positioning sample statements',
      'Added TTS play buttons next to SMART sample statements for audio playback.',
      'Code.gs, index.html',
      'Danzen',
      'Done'
    ],
    [
      '073',
      '2026-08-28',
      'Bug Fix',
      'NH',
      'TTS reading wrong text — play button bound to incorrect source element',
      'Corrected TTS text source — play button now reads from the correct statement element.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '074',
      '2026-08-28',
      'Bug Fix',
      'NH',
      'TTS text embedded as DOM attribute getting sanitized and removed before playback',
      'Embedded TTS text as data-tts-stmt attribute to prevent sanitizer removal.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '075',
      '2026-08-28',
      'Bug Fix',
      'NH',
      'TTS play button outside text box — data-attr being stripped by sanitizer',
      'Placed play button inside statement text box to avoid data-attr sanitizer issue.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '076',
      '2026-08-29',
      'Bug Fix',
      'NH',
      'TTS play buttons blocked by CSP and client-side sanitizer when rendered client-side',
      'Moved TTS play button rendering to server-side to bypass CSP and sanitizer restrictions.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '077',
      '2026-08-29',
      'Bug Fix',
      'NH',
      '9 filter section bugs found — mix of Critical, High, and Minor severity',
      'Resolved 9 filter section bugs ranging from Critical to Minor across filter bar and pill panel.',
      'Code.gs, Dashboard.html',
      'Danzen',
      'Done'
    ],
    [
      '078',
      '2026-08-29',
      'Bug Fix',
      'NH',
      '4 post-deploy filter regressions introduced after the 9-bug filter fix',
      'Resolved 4 post-deploy filter regressions (Critical + Minor).',
      'Code.gs, Dashboard.html',
      'Danzen',
      'Done'
    ],
    [
      '079',
      '2026-08-29',
      'Feature',
      'NH',
      'No QA filter pill in dashboard — observers could not filter by QA analyst',
      'Added QA (Observer Name) filter pill to shared filter bar.',
      'Code.gs, Dashboard.html',
      'Danzen',
      'Done'
    ],
    [
      '080',
      '2026-08-29',
      'Bug Fix',
      'NH',
      'readAuditLog reading by column position — broke when new columns were added or reordered',
      'readAuditLog now reads by actual header names instead of fixed position index.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '081',
      '2026-08-29',
      'Bug Fix',
      'NH',
      'observerName not included in handleSubmit formData — QA identity lost on submission',
      'Added observerName field to formData object passed through handleSubmit.',
      'Code.gs, index.html',
      'Danzen',
      'Done'
    ],
    [
      '082',
      '2026-08-29',
      'Feature',
      'NH',
      'QA observers identified by name input only — no verification of actual logged-in identity',
      'Added QA observer identification by logged-in Google account via Users sheet (GID 653814081).',
      'Code.gs, ObserverHelper.gs',
      'Danzen',
      'Done'
    ],
    [
      '083',
      '2026-08-29',
      'Feature',
      'NH',
      'QA filter pills populated from roster data — missed QA analysts not in roster',
      'Populated QA filter directly from Users sheet so all QA analysts appear regardless of roster.',
      'Code.gs, ObserverHelper.gs',
      'Danzen',
      'Done'
    ],
    [
      '084',
      '2026-08-29',
      'Bug Fix',
      'NH',
      'Observer/QA functions mixed into Code.gs — no separation of concerns',
      'Moved all observer and QA helper functions to separate ObserverHelper.gs file.',
      'ObserverHelper.gs, Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '085',
      '2026-08-29',
      'Bug Fix',
      'NH',
      'Users sheet found by tab name — lookup broke when sheet tab was renamed',
      'Changed Users sheet lookup to use GID (653814081) instead of tab name.',
      'Code.gs, ObserverHelper.gs',
      'Danzen',
      'Done'
    ],
    [
      '086',
      '2026-08-29',
      'Bug Fix',
      'NH',
      'dbBuildPanelHTML not including QA names from Users sheet in filter pill panel',
      'Fixed dbBuildPanelHTML to merge QA names from Users sheet when building pill panel.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '087',
      '2026-08-31',
      'Bug Fix',
      'NH',
      'Dashboard Overview KPIs, charts, and H&L showing stale data when filter returns 0 records',
      'Fixed dbBuildPanelHTML to clear stale Overview KPIs, charts, and H&L when filter yields zero records.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '088',
      '2026-09-01',
      'Feature',
      'NH',
      'Audit email only sent to Agent and TL — QA, QA TL, and Admin/Dev not receiving',
      'Expanded audit email recipients to Agent + Agent TL + QA + QA TL + Admin/Dev.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '089',
      '2026-09-01',
      'Bug Fix',
      'NH',
      'TTS play button in initPlayButtons targeting wrong CSS class — not triggering on click',
      'Fixed initPlayButtons to target .ai-flag-stmt class for correct TTS play button binding.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '090',
      '2026-09-01',
      'Bug Fix',
      'NH',
      'TTS event listener not reaching elements inside applied-output container',
      'Extended TTS event delegation to document level to cover applied-output container.',
      'index.html',
      'Danzen',
      'Done'
    ],
    [
      '091',
      '2026-09-01',
      'Bug Fix',
      'NH',
      'TTS handler not injected into new-window result views — silent in popup windows',
      'Injected Web Speech TTS handler into all new-window result views.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '092',
      '2026-09-01',
      'Bug Fix',
      'NH',
      'EvalView.html (email link page) had no TTS play buttons — read-only text experience',
      'Added TTS play buttons to EvalView.html using Web Speech API.',
      'EvalView.html',
      'Danzen',
      'Done'
    ],
    [
      '093',
      '2026-09-01',
      'Feature',
      'NH',
      'AI evaluation prompts were generic — not grounded in TELUS CX standards or LOB context',
      'Baked TELUS CX Blueprint framework into all AI evaluation prompts for LOB-specific scoring.',
      'AIService.gs',
      'Danzen',
      'Done'
    ],
    [
      '094',
      '2026-09-04',
      'Bug Fix',
      'NH',
      'getRecipientsFromRoster returning duplicate entries — Admin/Dev receiving multiple emails',
      'Deduplicated recipients in getRecipientsFromRoster to prevent duplicate admin emails.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '095',
      '2026-09-04',
      'Bug Fix',
      'NH',
      'Admin/Dev receiving a separate notification email per recipient — inbox flooding',
      'Consolidated all admin notifications into one email instead of sending per-recipient.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '096',
      '2026-09-04',
      'Bug Fix',
      'NH',
      'QA and Team Member not receiving audit emails — observerEmail not passed from client; agentEmail from untrusted client field',
      'Fixed observerEmail passthrough from client to server; fixed agentEmail to resolve from roster not client input.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '097',
      '2026-09-04',
      'Bug Fix',
      'NH',
      'Email subject format inconsistent — missing Audit Ref, QA, and LOB fields',
      'Updated email subject to "Name | Audit Ref | LOB | Locale" format.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '098',
      '2026-09-05',
      'Bug Fix',
      'NH',
      'QA name not included in email subject — harder to filter and track in inbox',
      'Added QA name to email subject format.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '099',
      '2026-09-05',
      'Bug Fix',
      'NH',
      'AI analysis scoring entire transcript including other agents — not focused on audited agent only',
      'Added transcript filtering to isolate target agent turns before AI analysis; implemented partial name matching.',
      'Code.gs, AIService.gs',
      'Danzen',
      'Done'
    ],
    [
      '100',
      '2026-09-05',
      'Bug Fix',
      'NH',
      'Transcript still including other-agent turns after initial filter implementation',
      'Fixed transcript filter to properly isolate only target agent turns before AI analysis.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '101',
      '2026-09-05',
      'Bug Fix',
      'NH',
      'False skipBlock reset triggered by inline colons in transcript turns — filter corrupted',
      'Fixed skip block logic to not reset on colons appearing mid-line in transcript content.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '102',
      '2026-09-05',
      'Bug Fix',
      'NH',
      'Partial name matching not excluding target agent from otherAgents list correctly',
      'Fixed partial name matching logic for excluding target agent from otherAgents array.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '103',
      '2026-09-05',
      'Feature',
      'NH',
      'Transcript filtering had 5 accuracy gaps (P1–P5) for edge-case multi-agent transcripts',
      'Implemented 5 accuracy improvements (P1–P5) covering edge cases in transcript filtering.',
      'Code.gs, AIService.gs',
      'Danzen',
      'Done'
    ],
    [
      '104',
      '2026-09-07',
      'Bug Fix',
      'NH',
      '4 bugs identified in post-implementation review of transcript filtering',
      'Resolved 4 bugs found post-implementation: null handling, edge case guards, output consistency.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '105',
      '2026-09-07',
      'Bug Fix',
      'NH',
      'Modal showing contradictory UI when target agent not found in transcript',
      'Fixed modal to hide excluded-agents block when target agent is not found.',
      'index.html, Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '106',
      '2026-09-07',
      'Bug Fix',
      'NH',
      'Surgical review found 12 bugs (2 Critical, 5 High, 5 Medium/Low) in transcript filtering implementation',
      'Resolved 12 bugs from surgical review: null references, incorrect conditionals, UI state inconsistencies.',
      'Code.gs, index.html',
      'Danzen',
      'Done'
    ],
    [
      '107',
      '2026-09-08',
      'Bug Fix',
      'NH',
      'Audited agent (Team Member) not receiving audit emails — email resolution failing',
      'Fixed agent email lookup chain; ensured agentEmail always resolved server-side before sending.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '108',
      '2026-09-08',
      'Performance',
      'NH',
      'agentEmailMap rebuilt on every call even when email column is absent from roster',
      'Added cache guard for empty agentEmailMap when email column is missing — skips rebuild on absence.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '109',
      '2026-09-10',
      'Bug Fix',
      'NH',
      'Stale agent email cache returning old data after roster changes; no debug tool available',
      'Invalidated stale agent email cache; added email lookup debug function for future diagnosis.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '110',
      '2026-09-10',
      'Feature',
      'NH',
      'All audits using same generic AI prompt regardless of Line of Business',
      'Built LOB selection modal — pops on transcript paste, forces LOB before analysis. getLOBKnowledge() injects LOB-specific knowledge into AI prompt. Supports PureFibre, SHS, Mobility, Koodo.',
      'Code.gs, AIService.gs, index.html',
      'Danzen',
      'Done'
    ],
    [
      '111',
      '2026-09-10',
      'Feature',
      'NH',
      'LOB modal not re-appearing after Clear All — stale LOB selection carried forward',
      'LOB modal now re-shows after Clear All to force fresh LOB selection per audit.',
      'index.html, Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '112',
      '2026-09-10',
      'Bug Fix',
      'NH',
      'Listener accumulation after multiple LOB selections; email subject LOB field order inconsistent',
      'Fixed listener accumulation bug in LOB modal; standardized LOB field order in email subject.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '113',
      '2026-09-11',
      'Bug Fix',
      'NH',
      'LOB modal close timer not canceling on Clear All — timer firing and auto-selecting LOB',
      'Fixed LOB close timer cancellation when Clear All is triggered.',
      'Code.gs, index.html',
      'Danzen',
      'Done'
    ],
    [
      '114',
      '2026-09-11',
      'Bug Fix',
      'NH',
      '_pendingFormData not cleared on Clear All; Process Script button text not restored',
      'Nulled _pendingFormData on Clear All; restored Process Script button text to original.',
      'index.html',
      'Danzen',
      'Done'
    ],
    [
      '115',
      '2026-09-11',
      'Feature',
      'NH',
      'Auto-populate and BAN validation running sequentially; BAN field not enforcing numeric input',
      'Built simultaneous auto-populate + numeric BAN validation to run in parallel.',
      'Code.gs, index.html',
      'Danzen',
      'Done'
    ],
    [
      '116',
      '2026-09-11',
      'Bug Fix',
      'NH',
      'Missing try/catch blocks; dead code paths not removed; _sapResolvedParticipant used inconsistently',
      'Surgical sweep: added try/catch, removed dead code, made _sapResolvedParticipant explicit throughout.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '117',
      '2026-09-11',
      'Bug Fix',
      'NH',
      '_sapResolvedParticipant flag set inside conditional — not guaranteed to run on all autofill paths',
      'Moved _sapResolvedParticipant=true to unconditional position after participant field write.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '118',
      '2026-09-11',
      'Feature',
      'NH',
      'No way to remove individual Critical Flag cards once added to audit output',
      'Added delete button to Critical Flag cards in AI result output.',
      'Code.gs, index.html',
      'Danzen',
      'Done'
    ],
    [
      '119',
      '2026-09-11',
      'Bug Fix',
      'NH',
      'Dead event listeners in applied-output; invalid CSS properties; lock sweep cursor missing',
      'Removed dead listeners from applied-output; fixed invalid CSS; added lock sweep cursor.',
      'Code.gs, index.html',
      'Danzen',
      'Done'
    ],
    [
      '120',
      '2026-09-11',
      'Bug Fix',
      'NH',
      'Button elements included in getEffectiveHTML output — button markup written to spreadsheet',
      'Added inline styles for button positioning; stripped button elements from getEffectiveHTML output.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '121',
      '2026-09-11',
      'Bug Fix',
      'NH',
      'position:relative from ai-flag being serialized into getEffectiveHTML — polluting stored HTML',
      'Stripped position:relative from ai-flag elements during getEffectiveHTML serialization.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '122',
      '2026-09-11',
      'Bug Fix',
      'NH',
      'position:relative persisting in _lastHTML before serialization — stored HTML polluted',
      'Stripped position:relative before DOM serialization to keep _lastHTML clean.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '123',
      '2026-09-11',
      'Feature',
      'NH',
      'Trainers not receiving audit emails — Trainer role not included in sendAuditEmail/sendSubmissionEmail',
      'Added Trainer role to email recipients in both sendAuditEmail and sendSubmissionEmail.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '124',
      '2026-09-11',
      'Bug Fix',
      'NH',
      'Stale roster cache returning old Trainer data after Trainer role was added to recipients',
      'Bumped roster cache key v1→v2 to bust stale cache after Trainer role addition.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '125',
      '2026-09-12',
      'Bug Fix',
      'NH',
      'Email sending could hit GAS 50-recipient-per-message limit — no batching in place',
      'Added sendEmailInBatches_() to split large recipient lists into batches of 50.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '126',
      '2026-09-12',
      'Bug Fix',
      'NH',
      'sendEmailInBatches_() not catching partial batch failures — silent errors on partial sends',
      'Added per-batch try-catch and partial failure handling to sendEmailInBatches_().',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '127',
      '2026-09-12',
      'Bug Fix',
      'NH',
      'Surgical hardening pass 1-2: 5 bugs across Code.gs, AIService.gs, Config.gs',
      'Fixed 5 bugs: null guards, error handling gaps, incorrect conditional branches.',
      'Code.gs, AIService.gs, Config.gs',
      'Danzen',
      'Done'
    ],
    [
      '128',
      '2026-09-12',
      'Bug Fix',
      'NH',
      'Third-pass hardening: 3 more bugs found after first two hardening passes',
      'Fixed 3 additional bugs in third-pass surgical review.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '129',
      '2026-09-12',
      'Bug Fix',
      'NH',
      'Fourth-pass hardening: 4 more bugs — async call guards and missing null checks',
      'Fixed 4 bugs in fourth hardening pass: async call guards, missing null checks.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '130',
      '2026-09-12',
      'Bug Fix',
      'NH',
      'Pass-5: execution timeout risk on long operations; email redirect lacking security check',
      'Fixed execution timeout risk in long sheet operations; added email redirect security validation.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '131',
      '2026-09-12',
      'Bug Fix',
      'NH',
      'Pass-6: double-deletion dead code persisting; ensureHeaders race condition',
      'Removed double-deletion dead code; fixed ensureHeaders race condition.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '132',
      '2026-09-12',
      'Bug Fix',
      'NH',
      'Pass-7: Admin/Dev double-send bug; ev2_ cache going stale after repair operations',
      'Fixed Admin/Dev double-send; fixed ev2_ cache invalidation after repair operations.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '133',
      '2026-09-12',
      'Bug Fix',
      'NH',
      'Pass-8: agent auto-email not firing; empty cache guard missing; resolveEmail cache stale',
      'Fixed agent auto-email trigger; added empty cache guard; fixed resolveEmail cache staleness.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '134',
      '2026-09-12',
      'Bug Fix',
      'NH',
      'Pass-9: sendSubmissionEmail Admin/Dev double-send',
      'Fixed Admin/Dev double-send in sendSubmissionEmail.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '135',
      '2026-09-12',
      'Bug Fix',
      'NH',
      'Pass-10: sendSubmissionEmail agent email fallback chain incomplete — some agents missed',
      'Completed agent email fallback chain in sendSubmissionEmail.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '136',
      '2026-09-12',
      'Bug Fix',
      'NH',
      'Pass-11: esc() function in dead-code block — never called; href attributes not escaped',
      'Rescued esc() from dead-code block; added href attribute escaping.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '137',
      '2026-09-12',
      'Bug Fix',
      'NH',
      'Pass-12: two withFailureHandler calls missing in autofill debounce paths',
      'Added missing withFailureHandler to final two autofill debounce call sites.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '138',
      '2026-09-14',
      'Feature',
      'NH + TI',
      'Audit result cards cramped — no visual hierarchy, no manual flag UI, no compact layout',
      'Overhauled audit result output: compact card layout, severity flags, manual flag add/remove UI.',
      'Code.gs, AIService.gs, index.html',
      'Danzen',
      'Done'
    ],
    [
      '139',
      '2026-09-14',
      'Bug Fix',
      'NH + TI',
      'sharedCSS_client() not updated to match new compact card layout — visual breakage',
      'Updated sharedCSS_client() to match new compact card layout.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '140',
      '2026-09-14',
      'Bug Fix',
      'NH + TI',
      'extractTextBlock chip regex failing; repeatPct keyword not matching; PDF CSS broken in buildEvalFormHTML',
      'Fixed chip extraction regex; fixed repeatPct keyword match; corrected PDF CSS in buildEvalFormHTML.',
      'Code.gs, AIService.gs',
      'Danzen',
      'Done'
    ],
    [
      '141',
      '2026-09-14',
      'Performance',
      'NH + TI',
      'Autofill waiting for full GAS round-trip on every field; no lookup cache pre-warming on load',
      'Added instant autofill for basic fields; added warmLookupCaches() pre-warm on page load.',
      'Code.gs, AIService.gs',
      'Danzen',
      'Done'
    ],
    [
      '142',
      '2026-09-14',
      'Enhancement',
      'NH + TI',
      'No loading indicator during GAS autofill; repeat agents triggering full GAS call every time',
      'Added _setAutofillLoading() loading indicator; added localStorage name→SAP cache for instant repeat-agent autofill (LS_NAME_SAP_KEY).',
      'Code.gs, index.html',
      'Danzen',
      'Done'
    ],
    [
      '143',
      '2026-09-15',
      'Bug Fix',
      'NH + TI',
      'All trainers receiving audit emails for agents not on their team (Matt O\'Brien complaint)',
      'Built buildAuditRecipients(): Trainee → specific Trainer Email + Supervisor Email (from trainee roster); Tenured → Team Leader + OM (from Global Roster). Removed TRAINER_ROLE blast. Added getTraineeInfo() and getTrainerInfo() with CacheService + header detection.',
      'Config.gs, AIService.gs, Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '144',
      '2026-09-17',
      'Bug Fix',
      'NH + TI',
      'Pressing Process Script immediately sends email to the audited agent before QA analyst reviews the result (regression)',
      'Removed auto-call of sendSubmissionEmail() from inside submitTranscript(). Email must only be triggered by the explicit "Submit & Email Audit" button (handleSubmitEmail → sendAuditEmail). The function was previously dead code that got incorrectly wired in.',
      'Code.gs',
      'Danzen',
      'Done'
    ],
    [
      '145',
      '2026-09-17',
      'Feature',
      'NH + TI',
      'Dashboard not accessible on load — QA analysts blocked by LOB and Audit Type modals before reaching dashboard',
      'Swapped landing page: Dashboard now opens automatically on DOMContentLoaded (openDashboard()). Renamed "← Back to Analyzer" button in dashboard topbar to "📋 Analyzer Form". Top-right "Dashboard" button on form still navigates back to dashboard.',
      'index.html',
      'Danzen',
      'Done'
    ],
    [
      '146',
      '2026-09-25',
      'Bug Fix',
      'NH + TI',
      'Observers viewing their own submitted evaluation saw a DIFFERENT team member and observer than the one they audited (reported: Danzen\'s Kenji Cagusangco audit, ref NHA-20260921-0003, resolved to Keiry Marisol Lazo Zometa / Francisco Rivera Campos). Root cause: generateAuditRef() sized its CacheService TTL as seconds-until-midnight, exceeding CacheService\'s 21600s (6-hour) hard cap on nearly every daytime submission, silently falling into a catch block that returned an uncollision-checked random 4-digit suffix — unrelated same-day submissions could land on the same NHA-YYYYMMDD-#### ref, and getEvalViewData()/getEvaluationByAuditRef() return the first matching row, permanently shadowing the later one.',
      'Replaced the CacheService-backed daily counter with PropertiesService, which has no expiry, so the sequence survives the full day without needing a TTL. Also stripped a pre-existing UTF-8 BOM from appsscript.json that was silently rejected by the Apps Script manifest validator and blocked clasp push, and added .claspignore so the Node-only tests/ harness is never pushed to the live script.',
      'Code.gs, appsscript.json, .claspignore',
      'Danzen',
      'Done'
    ],
    [
      '147',
      '2026-09-25',
      'Bug Fix',
      'NH + TI',
      'Code review of #146 found the same unguarded random-ref fallback still reachable via a LockService timeout under concurrent submissions (e.g. several observers submitting near end of shift), reproducing the identical collision risk through a different trigger than the one just fixed.',
      'Fallback now checks Audit_Log for the candidate ref before returning it (retries up to 10x via new _uniqueFallbackAuditRef()), and only drops to a millisecond-timestamp suffix if every guess collides. Added tests/audit-ref-test.js covering both the TTL-cap trigger and the lock-timeout trigger.',
      'Code.gs, tests/audit-ref-test.js',
      'Danzen',
      'Done'
    ]
  ];

  var dataRange = sh.getRange(2, 1, rows.length, numCols);
  dataRange.setValues(rows);

  // ── 5. Row formatting ────────────────────────────────────────────────────
  for (var r = 0; r < rows.length; r++) {
    var bg = (r % 2 === 0) ? '#FFFFFF' : '#F5F0FF';
    sh.getRange(r + 2, 1, 1, numCols).setBackground(bg);
    sh.setRowHeight(r + 2, 80);
  }

  sh.getRange(2, 5, rows.length, 3).setWrap(true);
  sh.getRange(2, 1, rows.length, 4).setHorizontalAlignment('center');
  sh.getRange(2, 8, rows.length, 2).setHorizontalAlignment('center');
  sh.getRange(2, 1, rows.length, numCols).setVerticalAlignment('top');

  // ── 6. Type badge colors (col 3) ─────────────────────────────────────────
  var typeColors = {
    'Bug Fix':     { bg: '#FEE2E2', text: '#991B1B' },
    'Feature':     { bg: '#DCFCE7', text: '#166534' },
    'Enhancement': { bg: '#DBEAFE', text: '#1D4ED8' },
    'Performance': { bg: '#FEF9C3', text: '#854D0E' },
    'Maintenance': { bg: '#F1F5F9', text: '#334155' }
  };
  for (var r = 0; r < rows.length; r++) {
    var type = rows[r][2];
    if (typeColors[type]) {
      var typeCell = sh.getRange(r + 2, 3);
      typeCell.setBackground(typeColors[type].bg);
      typeCell.setFontColor(typeColors[type].text);
      typeCell.setFontWeight('bold');
    }
  }

  // ── 7. Status badge colors (col 9) ───────────────────────────────────────
  var statusColors = {
    'Done':        { bg: '#DCFCE7', text: '#166534' },
    'In Progress': { bg: '#FEF9C3', text: '#854D0E' },
    'Pending':     { bg: '#F1F5F9', text: '#334155' }
  };
  for (var r = 0; r < rows.length; r++) {
    var status = rows[r][8];
    if (statusColors[status]) {
      var statusCell = sh.getRange(r + 2, 9);
      statusCell.setBackground(statusColors[status].bg);
      statusCell.setFontColor(statusColors[status].text);
      statusCell.setFontWeight('bold');
    }
  }

  // ── 8. Data validation ───────────────────────────────────────────────────
  sh.getRange(2, 3, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Bug Fix', 'Feature', 'Enhancement', 'Performance', 'Maintenance'], true)
      .setAllowInvalid(false).build()
  );
  sh.getRange(2, 4, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['NH', 'TI', 'NH + TI'], true)
      .setAllowInvalid(false).build()
  );
  sh.getRange(2, 9, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Done', 'In Progress', 'Pending'], true)
      .setAllowInvalid(false).build()
  );

  // ── 9. Filter ────────────────────────────────────────────────────────────
  sh.getRange(1, 1, rows.length + 1, numCols).createFilter();

  // ── 10. Done ─────────────────────────────────────────────────────────────
  Logger.log('Action Registry created: ' + ss.getUrl());
  SpreadsheetApp.flush();
}

/**
 * appendRegistryRow
 * Called via Apps Script Execution API to append a change entry to the live Action Registry.
 * @param {Object} data  { date, type, scope, issue, resolution, files, author, status }
 * @returns {Object}     { success, rowNum, sheetRow }
 */
function appendRegistryRow(data) {
  var REGISTRY_SS_ID = '1a58nCQPv9B0C1E0m30x5fibDrgnaJL6A7Pyz265Fjac';
  var ss = SpreadsheetApp.openById(REGISTRY_SS_ID);
  var sh = ss.getSheetByName('Registry');
  if (!sh) throw new Error('Registry sheet not found');

  // Determine next sequential #
  var lastRow = sh.getLastRow();
  var nextNum = '001';
  if (lastRow > 1) {
    var lastVal = sh.getRange(lastRow, 1).getValue();
    var n = parseInt(String(lastVal), 10);
    nextNum = isNaN(n) ? String(lastRow).padStart ? String(lastRow) : String(lastRow) : String(n + 1);
    while (nextNum.length < 3) nextNum = '0' + nextNum;
  }

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var row = [
    nextNum,
    data.date       || today,
    data.type       || 'Bug Fix',
    data.scope      || 'TI',
    data.issue      || '',
    data.resolution || '',
    data.files      || '',
    data.author     || 'Danzen',
    data.status     || 'Done'
  ];

  sh.appendRow(row);
  SpreadsheetApp.flush();

  var newRow = sh.getLastRow();
  var numCols = 9;

  // Alternating row background
  var bg = (newRow % 2 === 0) ? '#FFFFFF' : '#F5F0FF';
  sh.getRange(newRow, 1, 1, numCols).setBackground(bg);
  sh.setRowHeight(newRow, 80);

  // Alignment + wrap
  sh.getRange(newRow, 1, 1, 4).setHorizontalAlignment('center');
  sh.getRange(newRow, 5, 1, 3).setWrap(true);
  sh.getRange(newRow, 8, 1, 2).setHorizontalAlignment('center');
  sh.getRange(newRow, 1, 1, numCols).setVerticalAlignment('top');

  // Type badge
  var typeColors = {
    'Bug Fix':     { bg: '#FEE2E2', text: '#991B1B' },
    'Feature':     { bg: '#DCFCE7', text: '#166534' },
    'Enhancement': { bg: '#DBEAFE', text: '#1D4ED8' },
    'Performance': { bg: '#FEF9C3', text: '#854D0E' },
    'Maintenance': { bg: '#F1F5F9', text: '#334155' }
  };
  var tc = typeColors[data.type || 'Bug Fix'];
  if (tc) {
    var typeCell = sh.getRange(newRow, 3);
    typeCell.setBackground(tc.bg);
    typeCell.setFontColor(tc.text);
    typeCell.setFontWeight('bold');
  }

  // Status badge
  var statusColors = {
    'Done':        { bg: '#DCFCE7', text: '#166534' },
    'In Progress': { bg: '#FEF9C3', text: '#854D0E' },
    'Pending':     { bg: '#F1F5F9', text: '#334155' }
  };
  var sc = statusColors[data.status || 'Done'];
  if (sc) {
    var statusCell = sh.getRange(newRow, 9);
    statusCell.setBackground(sc.bg);
    statusCell.setFontColor(sc.text);
    statusCell.setFontWeight('bold');
  }

  SpreadsheetApp.flush();
  return { success: true, rowNum: nextNum, sheetRow: newRow };
}
