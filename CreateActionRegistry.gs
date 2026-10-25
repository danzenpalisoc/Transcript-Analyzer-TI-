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
  requireAdmin_();

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
    ],
    [
      '148',
      '2026-09-25',
      'Bug Fix',
      'NH + TI',
      'Analysts intermittently got "A lookup source could not be opened... the file is shared with a different account than the one the Analyzer runs as" and the audit was not saved (reported for Mohamed Rafiki, SAP 2008800, recurring). Investigation via Executions tab showed many unrelated functions (getRosterBySapId, getTranscriptMetadataAndRoster, detectTranscriptSpeakers, warmLookupCaches, sendAuditEmail, submitTranscript) failing together in short bursts then recovering — and the deployment actually runs as "Execute as: Me" (owner has full backend access), so the existing message\'s per-user file-sharing explanation was wrong for this deployment. Root cause could not be confirmed from Stackdriver logs: this script uses the GCP "Default" project, which has no accessible project ID for `clasp logs`, and the Executions UI detail view was not expandable in this session.',
      'Corrected the analyst-facing message to stop asserting an unconfirmed cause. Added logSubmissionError() (SheetService.gs), which writes every submitTranscript() failure — timestamp, observer, SAP ID, interaction ID, full error + stack — straight into a new Error_Log tab in the main spreadsheet, so future occurrences are diagnosable without needing GCP Cloud Logging access at all.',
      'Code.gs, Config.gs, SheetService.gs',
      'Danzen',
      'Done'
    ],
    [
      '149',
      '2026-09-25',
      'Bug Fix',
      'NH + TI',
      'Follow-up to #148: the "lookup source could not be opened" error was still unresolved and recurring, so a real fix (not just better logging) was needed. SpreadsheetApp.openById() is a documented source of a false "You do not have permission to access the requested document" error under heavy concurrent access, even for a file\'s own owner — a good fit for this deployment (runs as "Execute as: Me"; every analyst\'s requests share the owner\'s access) and for the observed pattern (many unrelated lookup functions failing together in short bursts, then recovering on their own).',
      'Added openSheetWithRetry() (SheetService.gs): opens a Spreadsheet by ID, retrying up to 3x with backoff only when the error looks permission-shaped, then throws naming the specific file if every attempt fails. Wired into every external lookup source the analyzer depends on: Main Spreadsheet, Roster, AT Data GCP / Trainee Roster, Global Roster, Users / Trainer Lookup (getOrCreateSpreadsheet, Code.gs, AIService.gs, ObserverHelper.gs). Covered by tests/sheet-retry-test.js.',
      'SheetService.gs, Code.gs, AIService.gs, ObserverHelper.gs, tests/sheet-retry-test.js',
      'Danzen',
      'Done'
    ],
    [
      '150',
      '2026-09-25',
      'Bug Fix',
      'NH + TI',
      'Deeper investigation into #148/#149 found the ACTUAL root cause: submitTranscript() (Code.gs) calls analyzeTranscript() -> callFuelIX() completely unguarded. A FuelIX auth/outage error\'s response body often contains wording like "permission" or "access" (the exposed FuelIX key was a live suspect), so submitTranscript()\'s generic substring check misclassified it as "a lookup source could not be opened" — a Sheets/Drive diagnosis that was wrong the whole time and sent every prior investigation (this session\'s and the original "TI report") looking at Roster/Drive sharing instead of the AI call. Fit the evidence better than #148/#149\'s theories: explains the 60+ second failing submitTranscript durations (AI call timeout, not a Sheets permission check) and is unrelated to whether Roster/Global-Roster/AT-Data sharing is fine (confirmed fine by the admin).',
      'callFuelIX() (AIService.gs) now tags every error it throws with AI_SERVICE_ERROR, and submitTranscript()\'s catch checks that tag BEFORE the generic permission substring check, returning an accurate "AI analysis service did not respond correctly" message instead of blaming a lookup source. Also retries up to 3x (1s/2s backoff) for transient-looking failures only — 429, 5xx, or a network-level exception — never for a bad/expired key (401/403) or malformed request (400), which fail identically every attempt. Found and fixed a bug in the first draft of the retry loop during test-writing: re-declaring `var` with no initializer does not reset it between loop iterations, so a stale network-error flag from one failed attempt was blocking the success path on a later attempt that actually returned 200. Covered by tests/fuelix-retry-test.js (18 checks).',
      'AIService.gs, Code.gs, tests/fuelix-retry-test.js',
      'Danzen',
      'Done'
    ],
    [
      '151',
      '2026-09-25',
      'Enhancement',
      'NH + TI',
      'Admin-proposed resilience improvement raised during the #148-150 investigation: if a team member has been audited before, reuse their known Team Leader / Ops Manager / LOB / Locale / VTID from that history as a fallback whenever the live roster autofill leaves those fields blank (e.g. an external lookup failure), instead of saving the evaluation incomplete or blocking the submission.',
      'Added getTeamDetailsFromAuditHistory(sapId) (SheetService.gs): reads Audit_Log — the analyzer\'s own spreadsheet, not an external file — for this SAP ID\'s most recent PRIOR entry, scanning newest-first. Wired into submitTranscript() (Code.gs): only backfills fields still blank after the normal autofill/locale-fallback steps, never overrides a value the live lookup already resolved, and pushes a warning ("please verify they are still current") whenever a fallback value is actually used, since the team member may have been reassigned since their last audit. Covered by tests/team-history-fallback-test.js (12 checks, including newest-entry-wins-over-oldest and graceful null on any read failure).',
      'SheetService.gs, Code.gs, tests/team-history-fallback-test.js',
      'Danzen',
      'Done'
    ],
    [
      '152',
      '2026-10-01',
      'Bug Fix',
      'NH + TI',
      'Evaluation emails linked to a DIFFERENT agent\'s evaluation, often with a different observer (reported: Rose Anne Ocana ref NHA-20260929-0004 showing Fatima Gabrielle Guadron Paredes; Rafael Rivas showing Heaven Esperilla; Brayan Torress showing another SHS TS agent with observer Emjay De Los Santos). The #146 audit-ref fix was already live (version 187 identical to HEAD) and the symptom continued, so it was not the cause. Root cause: Interaction ID was treated as unique everywhere, but a Repeats & Transfer call has several agents on ONE Interaction ID and each is audited separately. The result cache was keyed by Interaction ID + analysis type only, so the second agent got the first agent\'s cached result AND audit ref and no Audit_Log/Dashboard_Data row was ever written for them; sendAuditEmail only checked that the ref existed (subject/recipients come from the form, the evaluation link from the saved row); updateDashboardPDFLink wrote email status, recipients and the Issue Resolved override to the FIRST row with that Interaction ID; updateCachedResult overwrote the first agent\'s cached evaluation; and the proc_ lock blocked the second agent. getEvaluationFromCache also read Audit_Log column B (Submitted At) as the analysis type.',
      'Cache rows now carry the agent\'s SAP ID (Cache sheet column E, header labelled on first save) and findCachedResult/saveCachedResult/updateCachedResult match on Interaction ID + analysis type + SAP ID. Rows from before the column existed are served only to the agent who already has an audit for that Interaction ID in Audit_Log, never to a different agent. findAuditRefForInteraction (now backed by _findAuditRefForAgent in SheetService.gs) matches SAP ID and analysis type. A second agent on the same call now misses the cache, gets their own AI run, ref and rows. sendAuditEmail refuses to send, with a message to Clear All and re-process, when the form\'s SAP ID or Interaction ID differs from the saved audit (blank fields are not compared). updateDashboardPDFLink matches rows by Audit Ref; the processing lock is per Interaction ID + SAP ID; getEvalViewData and getEvaluationFromCache read the agent\'s own cache entry, and getEvaluationFromCache resolves Audit_Log columns by header name. Covered by tests/multi-agent-audit-test.js (30 checks, 18 failed before the fix). Audits already lost to this bug have no sheet row and must be re-submitted; emails already sent keep their wrong links.',
      'Code.gs, SheetService.gs, tests/multi-agent-audit-test.js',
      'Danzen',
      'Done'
    ],
    [
      '153',
      '2026-10-01',
      'Bug Fix',
      'NH + TI',
      'Follow-up to #152. The NH and TI Analyzers are two separate Apps Script projects running the same code against ONE spreadsheet (identical MAIN_SPREADSHEET_ID in Config). Each keeps its daily audit-ref counter in its own Script Properties and both restart at 1 every morning, so both issue NHA-YYYYMMDD-0001, -0002, ... and Audit_Log can hold two rows with the same reference. The evaluation link opens the first row for a reference, so it shows the other script\'s agent and observer. This fits the Rose Anne Ocana / Rafael Rivas / Brayan Torress reports (Brayan\'s email even named another observer) and is also a risk for the #152 identity check, which would block the legitimate owner of the second duplicate row.',
      'The reference prefix is now chosen by the running script\'s ID: NH keeps NHA-, TI becomes TIA-, any other script uses NHA- (getAuditRefPrefix_() and AUDIT_REF_PREFIX_BY_SCRIPT_ID in Config.gs), so Code.gs and Config.gs stay identical in both projects. generateAuditRef, _uniqueFallbackAuditRef and _generateTempRef use it. References already duplicated in the sheet are made harmless: sendAuditEmail uses the row that matches the form\'s SAP ID and Interaction ID instead of always the first, and updateDashboardPDFLink picks the row matching both reference and Interaction ID via the new _findAuditRowIndex(). Old references and old links are unchanged. Covered by tests/shared-sheet-audit-test.js (18 checks, 10 failed before the change). Known limit: duplicate references already in the sheet stay duplicated, so a link to one of them can still open the other row until the duplicate rows are cleaned up by hand.',
      'Config.gs, Code.gs, tests/shared-sheet-audit-test.js, tests/audit-ref-test.js',
      'Danzen',
      'Done'
    ],
    [
      '154',
      '2026-10-01',
      'Bug Fix',
      'NH + TI',
      'A multi-angle code review of #152 and #153 found problems in the fixes themselves. (1) The blank-owner rule for Cache rows written before the SAP ID column existed treated any agent with an Audit_Log row as the owner, so a second agent\'s edit-and-email deleted the first agent\'s only evaluation and findCachedResult served it to the second agent. (2) saveCachedResult read columns A:E of every Cache row, loading up to 48,000 characters of HTML per cached evaluation on the step that protects the paid AI result, which risks the 6-minute limit and a duplicate email when sendAuditEmail times out after sending. (3) #153 made sendAuditEmail send for an already-shared reference, but the link carries only the reference and opens the first row, so the second agent would be shown the first agent\'s evaluation. (4) sendAuditEmail wrote the edited evaluation under the form\'s Repeats/Sales radio instead of the saved audit\'s type. (5) A failed CacheService put() left the previous HTML to be served for up to 6 hours. (6) The processing-lock key was spelled out three times.',
      'A blank-owner Cache row now belongs to the agent whose audit CREATED it, the earliest Audit_Log row for the Interaction ID and type (_legacyCacheOwnerSap), and is served to and replaced by that agent only; the lookup uses TextFinder on Audit_Log and runs only when a blank-owner row is met. saveCachedResult reads columns A, B and E only and drops the old CacheService copy before writing; findCachedResult reads each matching row once and builds chunks from memory. sendAuditEmail refuses a reference that two Audit_Log rows share (with a message to contact the admin team, and a log entry) and updates the cache under the saved audit\'s type. _procKey_() builds the lock key in one place and CACHE_HEADERS lists the SAP ID column. New RefMaintenance.gs gives admins reportDuplicateAuditRefs() (read-only list) and applyDuplicateAuditRefRenames() (keeps the earliest row\'s reference, gives each later row a new one such as NHA-20260929-0004-D2 in Audit_Log, Dashboard_Data, Transcripts and Analysis, logs every change in Ref_Rename_Log, safe to run twice). Evaluation links already emailed for a renamed row stop working instead of showing another agent\'s evaluation. Covered by tests/cache-owner-test.js, tests/duplicate-ref-rename-test.js and a shared tests/lib/gas-fakes.js. Supersedes the duplicate-reference handling described in #153. NOT done: repairNextTruncated, repairTruncatedSalesEvaluations, diagnoseSalesEvaluations, backfillDashboardData and enrichDashboardData still address Cache rows by Interaction ID alone and are not owner-aware, so do not run them on Interaction IDs that have several agents; the client still shows an NHA- placeholder in the audit reference box before Process; generateAndServePDF still trusts the identity fields sent by the browser.',
      'Code.gs, SheetService.gs, Config.gs, RefMaintenance.gs, tests/cache-owner-test.js, tests/duplicate-ref-rename-test.js, tests/lib/gas-fakes.js',
      'Danzen',
      'Done'
    ],
    [
      '155',
      '2026-10-13',
      'Bug Fix',
      'NH + TI',
      'Typing a valid SAP ID sometimes left the form blank with no message (reported for SAP 243141, across hard refreshes and a fresh deploy; the deployed version was identical to main). Every failure in the lookup ended in a bare return: getRosterBySapId turned any exception into null, each roster loader swallowed its own failure into an empty result, and the page treated null as "not on the roster". The page also cached a row with no participant in localStorage and read that cache BEFORE the server, so one empty row was replayed for two hours and survived hard refreshes and redeploys.',
      'The roster loaders now record the source they could not read (_noteSourceError_). getRosterBySapId returns { _error: source_unavailable, sources } when it found nothing AND a source failed, or { _error: lookup_failed } on an exception; a genuine miss is still null and a hit is still the row. The page shows a status line under the SAP ID field (looking up, no agent found, which source could not be read, try again), ignores replies for a SAP ID that is no longer in the field, and only caches or replays a row that names a person. Commit 737df64. Covered by tests/roster-lookup-diagnostics-test.js and tests/sap-lookup-feedback-test.js. The exact cause of the original blank form was never confirmed (it worked after the redeploy); the page now reports the cause instead of hiding it.',
      'AIService.gs, Code.gs, index.html, tests/roster-lookup-diagnostics-test.js, tests/sap-lookup-feedback-test.js',
      'Danzen',
      'Done'
    ],
    [
      '156',
      '2026-10-13',
      'Bug Fix',
      'NH + TI',
      'App-wide review: one agent details, and one agent email address, could end up on another agent audit. _applyRosterRow wrote only the fields that had a value, so a blank Team Leader, VTID or LOB on the new agent left the previous agent value in the form. Editing Team Member replaced the SAP ID but kept the old Team Leader, VTID and the hidden agentEmail. The name lookup and the SAP lookup shared one timer and cancelled each other. Clear All did not cancel a transcript autofill still in flight, so it could write the old agent into the cleared form. On the server, buildAuditRecipients resolved the tenured agent email by NAME first, from a name map where the last row wins, so two roster agents with the same name meant one received the other evaluation.',
      'Every roster field is assigned (blank when missing) and the previous agent roster fields are cleared the moment a new lookup starts or the SAP ID drops below four digits; a name the analyst typed is never wiped. Editing Team Member now resolves the new agent in full. The two lookups have separate timers. Clear All cancels pending lookups and marks in-flight replies stale (a form epoch). The agent address is now taken from the roster row for the SAP ID, then the address the analyst was shown, then the name only when exactly one person has it; a name shared by people with different addresses resolves to nobody. Commit e058b38. Covered by tests/identity-reset-test.js, tests/audit-recipient-test.js and tests/lib/page-harness.js. NOT done: lookupSapId with no SAP ID on the form still returns the first roster row for a shared name, and Team Leader and Ops Manager addresses are still found by name (a shared name now resolves to nobody instead of the last row).',
      'index.html, Code.gs, tests/identity-reset-test.js, tests/audit-recipient-test.js, tests/lib/page-harness.js',
      'Danzen',
      'Done'
    ],
    [
      '157',
      '2026-10-13',
      'Security',
      'NH + TI',
      'App-wide review: the HTML sanitiser was a few regexes that removed quoted on*= attributes, so img src=x onerror=... and svg/onload=... passed straight through. The model output contains whatever was said on the call, so a transcript could ask for such markup; it ran in the analyst session, was stored, and was shown to agents and leaders. The evaluation page printed the ?ref= query value into a script block with JSON.stringify (which does not escape the less-than sign), a reflected XSS in a link that is emailed to agents and leaders. sendAnalyticsBriefing sent any HTML to any address from the owner mailbox for any signed-in user, and renderAnalyticsHTML printed the caller filter values unescaped into HTML that is cached and emailed.',
      'New HtmlSanitizer.gs (sanitizeEvalHtml_), copied identically into index.html and EvalView.html: an allow-list that rebuilds the markup from a token stream, so only listed tags and attributes come out with every value re-escaped, hrefs are limited to http, https, mailto and tel after entity decoding, and unclosed tags are closed. The Play button (a button holding a small svg) is kept inert: always type=button, svg limited to plain shapes. It runs on the server before anything is stored, mailed or turned into a PDF (saveCachedResult_, sendAuditEmail, generateAndServePDF, sendAnalyticsBriefing and the model output in submitTranscript). doGet refuses a ref that is not letters, digits, dot, underscore or hyphen and escapes every angle bracket, ampersand and line separator it prints. sendAnalyticsBriefing takes exactly one address on telus.com or telusinternational.com, and the filter values are escaped. A first version stripped the report CSS because the server sends sharedCSS() as a style block in front of the report; the page now keeps ONE leading style block if it holds only plain rules (commit cdd032b), and generateAndServePDF puts the server own stylesheet back. Commits 68dd847 and cdd032b. Covered by tests/html-sanitizer-test.js (61 payloads, a seeded 4000-input fuzz, real-report fidelity, copy-identity check) and tests/server-input-hardening-test.js. NOT done: contenteditable is still stripped on purpose so the evaluation page stays read-only; the page sets it again itself (see #162).',
      'HtmlSanitizer.gs, Code.gs, SheetService.gs, index.html, EvalView.html, tests/html-sanitizer-test.js, tests/server-input-hardening-test.js',
      'Danzen',
      'Done'
    ],
    [
      '158',
      '2026-10-13',
      'Security',
      'NH + TI',
      'App-wide review: every public function in a .gs file can be called by any signed-in user from the browser console with google.script.run. Maintenance functions that rewrite or delete sheet data had no check, callFuelIX was a free proxy to the AI key, getAllRosterData dumped the roster, saveCachedResult and updateCachedResult let anyone replace an agent cached evaluation, and the registry secret was committed in source. sendSubmissionEmail was dead code that sends mail and was public.',
      'New requireAdmin_() and isAdminCaller_() in Config.gs: the caller must be on ADMIN_USERNAMES on telus.com or telusinternational.com, with NO fallback to the effective user (in Execute as Me that is the owner for every caller). 35 maintenance functions start with it (repair, backfill, diagnose, test, audit-ref renames, createActionRegistry, appendRegistryRow, verifyLatestAudit, verifyMyLatestAudit and the others). A scan showed none of them has a caller in the browser, in other server code, or in a trigger (the project has none). Nine internal helpers are renamed with a trailing underscore, which Apps Script does not expose to the browser: callFuelIX_, getAllRosterData_, lookupBySapId_, lookupSapId_, lookupVTID_, saveCachedResult_, updateCachedResult_, findCachedResult_, getTeamDetailsFromAuditHistory_. The registry doPost now reads the secret from Script Property REGISTRY_SECRET, there is one appendRegistryRow (it was defined twice) that is admin-only and a private writer that goes through the formula-neutralising appendRow, and sendSubmissionEmail is removed. Commit 7094282. tests/access-control-test.js lists every function the pages call and fails if one is gated or private, or if a page starts calling an unlisted function. NOT done: REGISTRY_SECRET still falls back to the old constant, which is in git history and must be treated as public, until the Script Property is set and the external poster is updated. exportToSpreadsheet, updateDashboardPDFLink, getEvalViewData and getEvaluationByAuditRef are still callable by any signed-in user, and audit references are guessable.',
      'Config.gs, Code.gs, AIService.gs, SheetService.gs, RefMaintenance.gs, ObserverHelper.gs, CreateActionRegistry.gs, tests/access-control-test.js',
      'Danzen',
      'Done'
    ],
    [
      '159',
      '2026-10-13',
      'Bug Fix',
      'NH + TI',
      'Two problems found while testing sendAuditEmail. (1) In TEST mode the send failed with Could not resolve recipient email addresses although the Roster tab has Admin/Dev rows; the only log line was getRecipientsFromRoster(Admin/Dev): 0 found, which did not say which tab was read, whether the rows came from the 2-hour script cache, or whether the Role and Email columns were found. It cleared after the cache key changed, so a stale cache was the cause. (2) The screen showed the TEST toggle to anyone on ADMIN_USERNAMES (and used the effective user, the OWNER, when the caller own email was hidden), while the server honoured TEST only for the telus.com domain and otherwise QUIETLY SENT LIVE, so an admin on telusinternational.com could email the real agent, team leader and ops manager.',
      'The roster recipients cache key is now audit_roster_rows_v3. getRecipientsFromRoster logs the tab it read, the columns it found (and a WARNING if Role or Email is missing), whether the rows came from sheet, cache or memory, and which roles exist; sendAuditEmail says whether the TEST list (Admin/Dev) or the LIVE lists were empty. The screen and the server now use the same rule, isAdminCaller_(). A TEST request from someone who is not an admin is REFUSED with a clear message and nobody real is looked up; with no mode given it is still LIVE. Commits 16a03ff and a59101b. Covered by tests/roster-recipients-test.js and tests/audit-email-resilience-test.js (admin on either domain, non-admin, hidden email, foreign domain). Confirmed by the admin: the email arrived in TEST mode.',
      'Code.gs, tests/roster-recipients-test.js, tests/audit-email-resilience-test.js',
      'Danzen',
      'Done'
    ],
    [
      '160',
      '2026-10-13',
      'Bug Fix',
      'NH + TI',
      'App-wide review, data loss and corruption. (1) submitTranscript cached the report (step 5) and then appended the Analysis sheet with nothing around it; if that threw, the Dashboard_Data and Audit_Log rows were never written and the retry hit the cache, answered success with an EMPTY audit reference and wrote nothing, so the analyst saw a report that was in no sheet and the email failed with Audit reference not found. (2) The being-processed flag was released on only two of its four exits (a cache hit and Access denied left it for ten minutes) and with a blank Interaction ID on the form there was no guard. (3) A report over 48,000 characters is stored in chunks; one that began with = + - or @ became a formula and one that began with an apostrophe lost it, so the agent was served a damaged evaluation. (4) Replacing cached rows (read, delete by index, append) ran without a lock. (5) The dashboard cache chunk was 90,000 characters but the CacheService limit is 100 KB in bytes, so accented text overflowed it and the exception made getDashboardData return an EMPTY dashboard although the sheet had been read.',
      'submitTranscript is split into produceEvaluation_ (translate, filter, AI, sanitise, cache) and writeAuditRecords_ (Analysis, Dashboard_Data, Audit_Log; each write independent, and the Analysis write is now non-fatal with a warning naming the ref). A cache hit that has no Audit_Log row now completes the records from the cached report: no new AI call, the original ref is reused from the Transcripts row (findOrphanTranscriptRef_), and only missing rows are written (_sheetHasRef_). The flag guard keys on the resolved Interaction ID and is released in a finally block, only the flag this call took. Every cache chunk is written with the Sheets text marker (a leading apostrophe) and the other cells go through appendRow; the replace runs under the script lock (saved without it if the lock cannot be had). Dashboard cache chunks are 30,000 characters, the write cannot throw out of the function, and the meta entry goes in last so a half-written cache is never read as complete (keys now v3). Commit 982cc2a. Covered by tests/submit-recovery-test.js, tests/data-integrity-test.js, and a FakeSheet that behaves like Sheets for apostrophes and formula-looking text. Confirmed on a real audit by verifyMyLatestAudit (10 of 10 checks). NOT done: a report over 48,000 characters has only been tested against the fake, not on the real sheet.',
      'Code.gs, SheetService.gs, tests/submit-recovery-test.js, tests/data-integrity-test.js, tests/lib/gas-fakes.js',
      'Danzen',
      'Done'
    ],
    [
      '161',
      '2026-10-13',
      'Feature',
      'NH + TI',
      'The admin needed a cheap way to prove that one audit landed everywhere it should, instead of testing the sheets, the cache, the flag and the email by hand after every change.',
      'New AuditCheck.gs. verifyLatestAudit(ref) and verifyMyLatestAudit() are admin-only and read-only (they write no row and no cache entry). Run one from the Apps Script editor and read the Execution log: Audit_Log has exactly one row for the ref; Transcripts, Analysis and Dashboard_Data have theirs; the sheets agree on SAP ID and Interaction ID; the saved evaluation read from the Cache SHEET is present, not an error value and not cut off; the email status and recipients; no being-processed flag is left; and no error was logged for the interaction since the audit. It reads the Cache sheet directly because the 6-hour script-cache copy would hide a damaged sheet copy. verifyLatestAudit takes the newest audit in the whole sheet, which on a busy day is another analyst; verifyMyLatestAudit picks the newest audit whose Observer Name is the caller. Commits 1ebe2ab and dc519c3. Covered by tests/audit-check-test.js and ten mutation checks. Limit: it checks the data, not the screen (whether the email reached an inbox, how the report looks, whether the Play button works).',
      'AuditCheck.gs, tests/audit-check-test.js',
      'Danzen',
      'Done'
    ],
    [
      '162',
      '2026-10-13',
      'Bug Fix',
      'NH + TI',
      'The three report columns (What Working, What Needs to Change, How to Change It) were no longer editable in the Process Call Script modal. The model report carries contenteditable on the bullets and statements, but the HTML sanitiser from #157 removes the attribute on purpose (the evaluation page agents open must stay read-only) and the page only set it again after Apply Suggestions and Close. Separately, a Play button spoke its data-tts, which is fixed when the button is made, so an edited sample statement still played the original AI words.',
      '_makeReportEditable() runs when the modal opens and the applied view uses the same list. The How-to-Change list item itself is deliberately NOT editable: only the statement inside it is, because the outermost contenteditable is the editing host and an editable list item made Select-All plus typing replace the label, the statement and its Play button with plain text (found by trying it in Chrome). Play now reads the statement as it reads at click time (_ttsTextOf) and an emptied statement plays nothing. data-tts is copied from the edited text before the HTML is applied, emailed or saved (_syncTtsText), so the Play button on the saved evaluation page says what the analyst wrote. A Play button removed by backspacing is put back when the field loses focus, and all double quotation marks are removed from the text sent to the voice. Commit 66b6877. Verified in Chrome against the real helper code and CSS (typing at the end and in the middle, Select-All, blur, an emptied statement) and by the admin in the live app; covered by tests/report-editing-test.js and five mutation checks. NOT done: the audio itself was not heard in the browser test, only the text passed to the voice; the web-speech fallback is unchanged.',
      'index.html, tests/report-editing-test.js',
      'Danzen',
      'Done'
    ],
    [
      '163',
      '2026-10-13',
      'Audit',
      'NH + TI',
      'The admin asked for the dashboard to be audited: do all the numbers match the data behind them, and are they right. The dashboard is built FROM Audit_Log and joined to Dashboard_Data by Audit Ref; index.html and Dashboard.html are two separate implementations.',
      'Audit only, nothing in the dashboard was changed yet. tests/dashboard-audit.js runs the REAL dashboard code against small datasets with obvious answers (not part of the regular suite; it always exits 0). It confirmed 18 defects, from 10 root causes. (1) Avg Repeat Risk drops audits whose repeat risk is 0% (index.html dbRenderStats, renderOverviewTab and the per-agent table) while Dashboard.html keeps them, so the two dashboards disagree: 0, 0 and 40 shows 40 and 20 where 13 is right. (2) Unresolved is total minus resolved, so blanks and Unknown count as unresolved, the Resolution donut uses a different rule (the letters no anywhere), and Resolved is any text containing yes. (3) Flags are stored joined by commas and split on commas, so a flag title that contains a comma is counted as two. (4) The Top Agent column of the flag table is always the last agent seen. (5) A top performer is named even when everyone has 0 flags, and an agent whose name has a stray space is credited with 0. (6) The Sales Attempted chart can never show data because the merged rows do not carry that column (nor the five SMART columns or Agent Email), and renderSalesAttemptedDonut is never called. (7) A Dashboard_Data row with no Audit_Log row is hidden, and an Audit_Log row with no Dashboard_Data row is counted with blank analysis. (8) An audit with a blank Analysis Type is in no tab, so Repeats plus Sales is less than the Audit Log. (9) Call Driver is free text, so differently capitalised or spaced wording makes separate bars. (10) RCA Sub-Parameter is really the first two flag titles joined, so the Top Opportunities chart counts pairs of flags. Also noted, by design: Dashboard_Data keeps what the AI wrote, so flags the analyst deletes or adds, and edited bullets and statements, never reach the dashboard. New verifyDashboardData() in AuditCheck.gs is the admin-only read-only report for the REAL data: the join, headers, tab totals, exact versus dashboard Resolved and Unresolved, true versus dashboard Avg Repeat Risk, bad timestamps, shifted columns, email statuses and Call Driver variants. Covered by tests/dashboard-data-check-test.js and eight mutation checks. NOT done: none of the 18 defects is fixed; the admin has to choose which, and whether analyst edits should update Dashboard_Data. The real sheets have not been read yet: verifyDashboardData() has to be run by the admin.',
      'AuditCheck.gs, tests/dashboard-audit.js, tests/dashboard-data-check-test.js, tests/access-control-test.js',
      'Danzen',
      'Done'
    ],
    [
      '164',
      '2026-10-13',
      'Fix',
      'NH + TI',
      'The dashboard code was fixed for the defects confirmed in #163, after the admin ran verifyDashboardData on the real data: 540 Issue Resolved values are long sentences that start with a verdict, so the old rule (the letters yes anywhere, everything else unresolved) was wrong for most rows. The admin chose Partially resolved as its own count.',
      'One block of helper functions, IDENTICAL in index.html and Dashboard.html (a test compares the two). Issue Resolved is read by the START of the text: Resolved, Partially resolved (own KPI card), Unresolved, and Not Recorded for a blank or text with no verdict (never guessed); the KPI cards and the donut use the same rule. Avg Repeat Risk counts a real 0% and ignores blank and non-numeric values, on every tab and in both dashboards. Critical Flags: new audits are stored joined by " | " (Code.gs, so a title with a comma stays one flag); older rows are still read correctly, because a piece that starts in lower case belongs to the flag before it; "None" is no flag; one audit counts a flag once. Top Agent per flag is the agent with the most of it; a top performer or most-flags agent is named only when ONE agent is strictly lowest or highest (no ties, not a lone agent); names with stray spaces are one agent; Call Driver and flag buckets ignore capitals and spacing; Top Opportunities counts each RCA item, not pairs. The merge now carries EVERY Dashboard_Data column (Sales Attempted, the SMART columns, Agent Email), keeps an audit that exists only in Dashboard_Data, and the status line reports audits with no analysis row or no Analysis Type instead of leaving them out silently. verifyDashboardData uses the same Issue Resolved rule (Resolved / Partially / Unresolved / Not recorded, with a warning for text that has no verdict) and no longer reports a repeat-average gap. tests/dashboard-audit.js became the strict tests/dashboard-correctness-test.js and the full suite passes (26 files); 15 mutation checks were all caught. Numbers on the dashboard will change: Resolved, Unresolved, Avg Repeat Risk, flag counts and Top Agent were wrong before. NOT done: the real data repairs (the hidden audit with no Audit_Log row, the duplicate ref, the SAP mismatch, the shifted rows, the missing Agent Email header, non-numeric repeat values), analyst edits still do not reach Dashboard_Data, and the unused renderSalesAttemptedDonut is still in index.html.',
      'index.html, Dashboard.html, Code.gs, AuditCheck.gs, tests/dashboard-correctness-test.js, tests/dashboard-data-check-test.js',
      'Danzen',
      'Done'
    ],
    [
      '165',
      '2026-10-13',
      'Tool',
      'NH + TI',
      'After the dashboard fix, verifyDashboardData on the real data still reported 8 data problems: Agent Email header missing in Audit_Log and Dashboard_Data, duplicate ref NHA-20261001-0001, NHA-20260916-0016 in Dashboard_Data only, 2 SAP mismatches, 23 non-numeric Repeat %, 7 shifted Team Member rows. The admin asked for all of them to be fixed. The code cannot read the live sheets, so which value is right in a mismatch or a shifted row cannot be decided without seeing the rows.',
      'New DataRepair.gs (admin-only). addMissingHeaders() writes a missing header at the END of a sheet header row (Agent Email is the last column in the code for Audit_Log, Dashboard_Data and Transcripts, and rows are appended by position, so only the header cell is missing). It moves and clears nothing, and does nothing when the existing headers are not exactly the first part of the expected list, so it replaces the risky repairSheetHeaders (which clears and rewrites the whole sheet) for this case. diagnoseDataIssues() is read-only: it prints the exact rows behind the duplicate ref, the Dashboard_Data-only audit, the SAP mismatches, the non-numeric Repeat % values and the shifted rows, across Audit_Log, Dashboard_Data, Transcripts and Analysis, with long cells cut short. Covered by tests/data-repair-test.js (23 checks) and 5 mutation checks. NOT done yet: the duplicate ref, hidden audit, SAP mismatch, Repeat % values and shifted rows are not changed; the exact fix will be written after the admin runs diagnoseDataIssues and the rows can be seen.',
      'DataRepair.gs, tests/data-repair-test.js, tests/access-control-test.js',
      'Danzen',
      'Partly done'
    ],
    [
      '166',
      '2026-10-13',
      'Fix',
      'NH + TI',
      'diagnoseDataIssues on the real data showed three things. (1) A regression from #164: _dbNum read a Repeat Projection % cell strictly, but hundreds of older audits store a whole sentence there ("95% repeat risk. No issue was identified...", "~55% repeat risk...", "High - approximately 80%..."). The old code used parseFloat and read the leading number, so after #164 the Avg Repeat Risk silently left those audits out. The tests only used clean values like 40%, so they did not catch it. (2) NHA-20260916-0016 exists in Dashboard_Data, Transcripts and Analysis but has no Audit_Log row. (3) NHA-20260911-0023 has SAP 2008652 in Dashboard_Data and 2008666 in the other three sheets. NHA-20261001-0001 is two different audits (different SAP and interaction) that share a ref, which the existing rename tools handle.',
      '_dbNum now reads a plain number, or the first NN% inside a sentence (a range such as 40-50% gives its middle, a percentage above 100 inside a sentence is not trusted, text with no percentage is left out, not counted as 0). The same reading is in index.html, Dashboard.html (still one identical helper block) and the verifyDashboardData checker, which now warns how many cells hold a sentence. DataRepair.gs gets previewDataRepairs() (read-only) and applyDataRepairs(): rebuilds an Audit_Log row from Dashboard_Data when the audit is missing there (only when SAP ID, Interaction ID, Team Member (80 characters or less) and Timestamp are present and the ref is not listed twice; Recipients stays blank); corrects a Dashboard_Data SAP ID only when the other three sheets all agree and Audit_Log is not the odd one out; turns a Repeat % sentence into the number and keeps the original sentence in a cell note. Anything unclear is skipped and listed. It stops before the time limit and is safe to run again. diagnoseDataIssues output is now short (summary plus a few examples) because the full list was cut off at 50,000 characters. Covered by tests/data-repair-test.js (all repair checks) plus the sentence checks in tests/dashboard-correctness-test.js and tests/dashboard-data-check-test.js; 12 mutation checks all caught. NOT done: the 7 shifted Team Member rows (the log was cut off before that section), and the duplicate ref NHA-20261001-0001 still has to be renamed by running reportDuplicateAuditRefs then applyDuplicateAuditRefRenames.',
      'index.html, Dashboard.html, AuditCheck.gs, DataRepair.gs, tests/data-repair-test.js, tests/dashboard-correctness-test.js, tests/dashboard-data-check-test.js, tests/lib/gas-fakes.js',
      'Danzen',
      'Done'
    ],
    [
      '167',
      '2026-10-13',
      'Fix',
      'NH + TI',
      'After #166 was applied, verifyDashboardData left one problem: 7 audits with a Team Member over 80 characters. diagnoseDataIssues showed two kinds. (1) Two audits (NHA-20260824-0006, NHA-20260825-0004) had transcript header text glued onto Team Member, Interaction ID, Direction and Duration ("Nouamane BadraneExternal Participant(s): HARV..."). (2) Five audits had a list of every participant in Team Member, observer included. The same output also showed that Repeat Projection % uses two scales: about 1,826 cells hold a fraction (0.78 for 78%, because Sheets turns the text "78%" into a percent number) and the rest hold whole numbers (65, or the number written by #166). The dashboard read both as they were, so Avg Repeat Risk mixed 0.78 with 78 and the Audit Log table showed "0.78%". This was a standing problem, not caused by #164 or #166.',
      '_dbNum (one identical helper block in index.html and Dashboard.html), the checker and the server briefing now read a bare number above 0 and up to 1 as a fraction (0.78 -> 78); a written "1%" or "0.5%" stays as written, and a bare 1 means 100%. The Audit Log table in both dashboards shows N% from the same reading. DataRepair.gs previewDataRepairs / applyDataRepairs also clean messy cells in Audit_Log, Dashboard_Data, Transcripts and Analysis: a Team Member that holds several names or glued header text becomes the clean Team Member that the SAME SAP ID has on its other Audit_Log audits, and only when that name is inside the messy cell (otherwise the cut before "External Participant(s)" / "Direction:" when that is a single clean name, otherwise the row is skipped and listed); an Interaction ID longer than a UUID is cut to its first 36 characters; Direction becomes Inbound or Outbound; a Duration with text glued on becomes h:mm or h:mm:ss and is stored as text so Sheets does not turn it into a time. Every cleaned cell keeps its original value in a cell note. Covered by tests (scale, both tables, cleaning, skip cases, notes, text format) and 10 mutation checks, all caught. The first preview of the cleaning would have shortened the real Direction value Inbound / Outbound (held by about 1,250 audits) to Inbound; found by reading the preview before applying, fixed so Direction is only cleaned when text is glued straight onto the word, with a test and a mutation check. Two audits (NHA-20260903-0003 and NHA-20260929-0024) had no clean audit for their SAP ID, so the admin chose the agent (Ali Guennoun and Jonathan Cardoza; Jonathan Cardoza is also the observer on that audit, and the evaluation text names Krishna Singh); the choice is used only when the name is inside the messy cell. NOT done: Issue Resolved and Transfer Occurred cells in some of these rows hold a sentence or HTML text from the same shift; the dashboard counts them as Not Recorded and they were not rewritten. A Duration that Sheets stored as a 1899 date-time (for example NHA-20260903-0003) is not converted. The Not Sent count (503) was not looked at.',
      'index.html, Dashboard.html, AuditCheck.gs, Code.gs, DataRepair.gs, tests/data-repair-test.js, tests/dashboard-correctness-test.js, tests/dashboard-data-check-test.js, tests/lib/gas-fakes.js',
      'Danzen',
      'Done'
    ],
    [
      '168',
      '2026-10-13',
      'Feature',
      'NH + TI',
      'The admin asked for a business-review scorecard for each agent, like a slide shown in team reviews: name, SAP ID, number of audited interactions, a status badge, four large KPI tiles with Target or Below threshold, supporting numbers, Audit feedback (Strengths, Gaps, Also noted) and Coaching advice, one slide per agent, in Google Slides. Built from analyzer data only, from the current dashboard filters, open to every analyst, in TELUS style and easy to read.',
      'New Generate Slides button in the header of the Team Member Trends and Repeat Offenders section of the Overview tab, beside the High, Moderate and Low legend (first put in the top bar next to Refresh, moved because users could not find it; the filters apply to every tab, so it works from there). It is a filled purple button, 44 px tall, with a plain SVG icon. It opens a dialog (agents and audits that will be used, optional deck title) and builds a Google Slides deck: a cover slide with the period, the filters and a legend of the targets, then one slide per agent. The page sends only the Audit Refs of the audits that pass the filters; createScorecardDeck (ScorecardSlides.gs) reads the sheets itself, so no number comes from the browser. Tiles: Avg Repeat Risk (target 30 percent or lower), Resolution rate (80 percent or higher, of audits with a clear verdict), Critical flags per audit (target 0, kept by the admin although the AI template writes 2 to 5 flags per audit, so most agents show red on it), Transfer rate (20 percent or lower, Repeats audits) or Sales attempted (80 percent or higher, Sales-only agents). The badge counts tiles below target: none = On track, one = Needs attention, two or more = Priority support. Strengths come from the What is Working items of the newest 3 reports of each agent (read cell by cell from the Analysis sheet, never the whole column), Gaps from the most common flags, Coaching from the most common Top SMART Opportunity. The rules are copies of the dashboard rules (a parity test runs both on the same examples). The look is one layout and theme (ScorecardLayout.gs) that the approved HTML preview and the Slides builder both read: 16:9, Arial, TELUS purple, text at least 11 pt (footer 9 pt), contrast at least 4.5 to 1 for every colour pair, status always in words, text cut at a word end so nothing overflows (tests check contrast, font size, overflow and bounds, including the worst-case agent). How the deck is made: the first design called the Google Slides API, but the Cloud project cannot have that API switched on (the first run of testSlidesAccess returned 403 SERVICE_DISABLED and the admin cannot enable it). So ScorecardPptx.gs now writes the slides as a PowerPoint (.pptx) package of plain XML, ScorecardSlides.gs zips it, and Drive converts it to Google Slides with Drive.Files.insert (the advanced Drive service the PDF export already uses): no Slides API, no Cloud setting, no new permission, one call for any number of agents. In a PowerPoint file the layout is written exactly: position, size, fill, outline, corner radius (so the approved rounded cards are back), inner margin, vertical alignment, font, size, colour, line spacing and bullets. Drive has two conversion styles; the one the PDF export uses is tried first, and the other if it is refused. Guards: refs validated and capped at 2,000, at most 40 agents per deck, 5 decks per person per hour, one deck at a time, the deck is filed in the folder NH Analyzer Scorecards of the app owner and shared as editor only with the person who asked, an analyst sees a plain message on failure and an admin also sees the detail. Admin tools: previewScorecards (read-only, logs the numbers to compare with the dashboard) and testSlidesAccess (builds a one-slide deck to prove Slides works). No new permission is asked for (a test checks the manifest). Covered by tests/scorecard-metrics-test.js, tests/scorecard-pptx-test.js (strict XML, package integrity, exact geometry, escaping), tests/scorecard-slides-test.js, tests/scorecard-button-test.js, tests/gs-syntax-test.js and 46 mutation checks, all caught. The package was also opened with python-pptx, an independent reader, which found every shape, position, text, font and colour as in the layout. The dialog was tried in Chrome in each state, including at phone width. Tried by the admin on 2026-10-13: testSlidesAccess made a sample deck through the Drive conversion, and the Generate Slides button works in the live dashboard. NOT done: the exact look of a real deck in Google Slides (fonts, rounded corners, bullets, spacing) was reported as working but was not seen by the developer; a run with more than 40 agents is refused by design. The standalone Dashboard.html page has no button. Batch name and batch VTID of the sample slide are not in analyzer data, so the footer shows Team Leader, LOB, Locale and VTID.',
      'ScorecardMetrics.gs, ScorecardLayout.gs, ScorecardPptx.gs, ScorecardSlides.gs, Config.gs, index.html, tests/scorecard-metrics-test.js, tests/scorecard-pptx-test.js, tests/scorecard-slides-test.js, tests/scorecard-button-test.js, tests/gs-syntax-test.js, tests/access-control-test.js',
      'Danzen',
      'Done'
    ],
    [
      '169',
      '2026-10-13',
      'Fix',
      'NH + TI',
      'After the dashboard data and the scorecard were done, the admin asked for everything that was left to be done: the Duration cells that Sheets turned into a time of day, the 505 Not Sent audits, the Issue Resolved answers with no verdict, the registry secret, the unprotected exportToSpreadsheet and updateDashboardPDFLink, and the AI reliability items found in the app-wide review.',
      'SECURITY. updateDashboardPDFLink was public although only sendAuditEmail calls it, so any signed-in user could change the Email Status, Recipients or Issue Resolved of any audit from the browser console; it is now private (updateDashboardPDFLink_). exportToSpreadsheet took the rows from the browser, wrote them to a spreadsheet in the owner Drive shared with the whole domain, turned any cell starting with = into a formula and had no limit; it is replaced by exportAudits (ExportService.gs), which takes only a view name (repeats, sales, auditlog) and Audit Refs and builds the rows itself from the sheets with the same merge as the dashboard (parity test on every column), stores cells as text (= + - @ protected, time-like text kept as text, cells cut at 49,000 characters), allows up to 5,000 audits, 10 exports per person per hour and one at a time, files the sheet in the folder NH Analyzer Exports and shares it as viewer with the requester only (never with the domain). The page sends only refs and opens only a Google Sheets link. AI RELIABILITY. A cut-off AI answer (finish_reason length or max_tokens) used to be logged as a warning and returned as a finished report, saved and shown; it is now an error tagged AI_SERVICE_ERROR[truncated], nothing is saved, no retry (the same answer would be cut off again), and the analyst is told the report was cut off and to submit again. The three attempts shared no time limit (3 x 270 s against the 360 s Apps Script limit); they now share one 300 s budget, a retry gets only the time that is left, and there is no retry when under 60 s is left. A saved draft lost its Audit Ref because the restore ran before the variable was declared further down the script and was then reset to blank; the restore now runs after the script has loaded (the test harness can start with a saved draft). DATA. applyDataRepairs also turns a Duration that Sheets stored as a time of day (a Date on 1899-12-30, for example NHA-20260903-0003) back into text such as 2:15:05, in Audit_Log, Dashboard_Data and Transcripts, keeps the original in a cell note, and skips and lists anything that is not a time within one day. New read-only diagnoseDashboardGaps reports how many Duration cells are affected, where the Not Sent audits are (by month, type and observer, and whether any has recipients), and how the Issue Resolved answers with no verdict look and start, so that nothing is guessed. Covered by tests/export-test.js, tests/ai-reliability-test.js, new cases in tests/submit-recovery-test.js, tests/data-repair-test.js and tests/access-control-test.js, and 28 mutation checks, all caught. NOT done: (1) the registry secret: the old secret is still in the code and an admin must set the Script Property REGISTRY_SECRET to a new value (and tell whoever posts to the registry); (2) Not Sent and Issue Resolved are only reported, a repair waits for the admin to run diagnoseDashboardGaps; (3) evaluation links carry only the audit ref, which is guessable, so anyone in the domain who guesses a ref can open that evaluation; signing the links would stop the links already emailed, so it needs a decision; (4) accessibility of the main page was not started.',
      'ExportService.gs, Code.gs, AIService.gs, DataRepair.gs, index.html, tests/export-test.js, tests/ai-reliability-test.js, tests/submit-recovery-test.js, tests/data-repair-test.js, tests/access-control-test.js, tests/lib/page-harness.js',
      'Danzen',
      'Partly done'
    ],
    [
      '170',
      '2026-10-13',
      'Fix',
      'NH + TI',
      'App-wide code review, batch 1 (security). Three independent reviewers and a manual check found that any signed-in analyst could use the app to send mail as the owner to anyone, run code in another user browser through a crafted agent name, plant formulas in the sheets, and call dozens of internal functions from the browser console.',
      'EMAIL RELAY. sendAuditEmail accepted any recipient list, any audit and any observer from the browser, and sent it from the owner account. New EmailGuard.gs and changes in sendAuditEmail: (1) only the analyst who saved the audit, or an admin, can email it (matched by observer name, e-mail or the part before the at sign); anyone else gets Only the analyst who saved this audit, or an admin, can email it, nothing is sent, and the attempt is logged. (2) Recipients must be on a company domain (telus.com, telusinternational.com, plus any in the Script Property EMAIL_EXTRA_DOMAINS); outside addresses are dropped, logged and reported back as droppedRecipients, and if none is left nothing is sent. (3) A rate limit per person per hour (30, admins 100) using a hashed cache key. (4) Agent name, SAP ID, interaction id and audit ref are taken from the saved row first, the observer e-mail comes from the signed-in caller, and text typed in the form is cut to one line so it cannot add mail headers. FORMULA GUARD. updateDashboardPDFLink_ stored the e-mail status, recipients and the analyst Issue Resolved override with setValue, so a value starting with = ran as a formula with the owner rights; every string is now stored as text (apostrophe guard), the status is cut to 40 characters and the override can only be Yes, No or Partially resolved. XSS IN THE DASHBOARD. The filter pills, the member lists and the Audit Ref links built their onclick text from the name or reference; the browser decodes an attribute before running it, so an agent named with a quote and a closing bracket ran code for whoever opened the filter (an admin included), and a name like OBrien with a quote broke its own checkbox. The handlers are now fixed text and the value travels in the value or data attribute and is read with this.value or this.getAttribute. The report window escapes the reference in its title and heading and tells the person when the browser blocks the pop-up instead of failing. PRIVATE FUNCTIONS. 32 public functions that send mail, write the shared caches, spend AI quota or return rosters, e-mail addresses and the whole audit log are now private (trailing underscore): analyzeTranscript, callAIForAnalytics, translateTranscriptText, extractStructuredRCA, notifyAdmins, writeDashboardCache, writeAnalyticsCache, invalidateDashboardCache, invalidateAuditLogCache, readAuditLog, readDashboardSheet, readDashboardCache, readAnalyticsCache, getAdminRecipients, getRecipientsFromRoster, lookupAgentEmail, resolveEmail, resolveObserver, logSubmissionError and the others listed in tests/access-control-test.js. A new test fails when any public function is not in exactly one group (Browser, Maintenance, Pure or Accepted), when a Pure function touches a sheet, mail, cache, Drive or the network, or when a group lists a function that is no longer public. Covered by tests/email-relay-test.js, tests/dashboard-handlers-test.js (it runs each handler the way the browser does, with hostile names), tests/access-control-test.js, updated older tests, and 14 mutation checks, all caught. NOT done: the registry secret still needs the Script Property REGISTRY_SECRET; the evaluation links are still guessable; the mail guard cannot stop an analyst from emailing their own audit to a colleague, which is allowed by design.',
      'EmailGuard.gs, Code.gs, index.html, AIService.gs, SheetService.gs, ObserverHelper.gs, RefMaintenance.gs, DataRepair.gs, AuditCheck.gs, ExportService.gs, ScorecardSlides.gs, Config.gs, tests/email-relay-test.js, tests/dashboard-handlers-test.js, tests/access-control-test.js, tests/cache-owner-test.js, tests/multi-agent-audit-test.js, tests/shared-sheet-audit-test.js, tests/audit-email-resilience-test.js',
      'Danzen',
      'Partly done'
    ],
    [
      '171',
      '2026-10-13',
      'Fix',
      'NH + TI',
      'App-wide code review, batch 2 (features that were silently broken): the Strengths on the scorecard, tab switching on the dashboard, the Flag Records button, the Previous Evaluations table, a stale Audit Log cache and the flags and dates of the scorecard.',
      'STRENGTHS. The scorecard read the What is Working items from report HTML, but the Analysis sheet holds the report as plain text, so every deck fell back to Meets the target. extractWhatsWorking_ now also reads the plain text (the heading line and the bullets under it, blank lines between bullets allowed, stops at the next heading, at most 3), and the new test builds the text with the real htmlToPlainText instead of hand-made HTML (the first run of that test caught a blank-line mistake in the new reader). TAB SWITCH. dbSwitchTab read _dbCurrentFilters, which was never declared, so every tab change threw a ReferenceError after the tab was shown and the chart redraw, the analytics button and the last-tab memory never ran; it is declared and set in dbApplyAll. FLAG RECORDS. The button called dbDownloadCSV, which exists nowhere; it now exports only the audits that have a flag through the same safe server export (exportAudits), one row per audit. PREVIOUS EVALUATIONS. getEvalViewData returned ref, date and ban but EvalView.html reads interactionId and analysisType, so every row showed a dash and Repeats; both are returned now. An evaluation still being processed is no longer cached for 6 hours. AUDIT LOG CACHE. updateDashboardPDFLink_ wrote the Email Status to Audit_Log but only dropped the dashboard cache, so the Audit Log kept showing Not Sent for up to 20 minutes and an export built from it could send twice; it drops the Audit Log cache too. SCORECARD NUMBERS. 1 flag in 21 audits rounded to 0.0 and showed On target; the tile is now judged on the real average and shows <0.1. Dates on the slide used UTC, so an audit at 8 PM Eastern moved to the next day (and 10 PM on Dec 31 to the next year); they use the script time zone. Covered by tests/dashboard-broken-features-test.js (new, it first failed with the real ReferenceError), new cases in tests/scorecard-metrics-test.js, tests/multi-agent-audit-test.js and tests/shared-sheet-audit-test.js, and 14 mutation checks, all caught. NOT done: the standalone Dashboard.html was checked and does not have these bugs; the remaining items of the review (batch 3 mediums, batch 4 accessibility) are open.',
      'ScorecardMetrics.gs, Code.gs, index.html, tests/dashboard-broken-features-test.js, tests/scorecard-metrics-test.js, tests/multi-agent-audit-test.js, tests/shared-sheet-audit-test.js',
      'Danzen',
      'Partly done'
    ],
    [
      '172',
      '2026-10-14',
      'Fix',
      'NH + TI',
      'App-wide code review, batch 3 (medium findings): limits and folders of the scorecard and export, AI replies, admin repair tools, dashboard loading, the standalone dashboard and the entry form.',
      'LIMITS AND FOLDERS. createScorecardDeck and exportAudits counted their hourly quota only after the heavy sheet reads and only for a finished build, so one person repeating a try that ended empty could keep the lock and the sheets busy for everybody. Every try is now counted before the lock (30 per hour for decks, 60 for exports, helper _hourlyTake_ in Config.gs, hashed key), next to the existing limits of 5 decks and 10 exports. The folders for decks, exports and PDFs were found by name, which also returns folders other people shared with the owner, so anyone could plant a folder with that name and have files put in (or cleaned out of) it; the new _ownedFolder_ takes only a folder the app owner owns and makes a new one otherwise. AI. A 200 reply that was not JSON, or had empty or null text, was saved as a finished report or crashed on null; both are now errors (AI_SERVICE_ERROR[invalid] and [empty]) and are retried while the time budget allows. The 300 second budget counted from the first fetch, but the roster and sheet lookups before it are part of the same 360 second limit; it now counts from the start of the execution. The pasted transcript is untrusted text: it is cut at 150,000 characters, a forged END OF TRANSCRIPT marker inside it is defused, and the prompt says nothing inside the transcript is an instruction. ADMIN TOOLS. repairSheetHeaders deleted every column the code did not know (for example Recommendations) and ran without a lock; it now keeps extra columns after the standard ones with their data (a column with no header only when it holds data), and takes the script lock. repairNextTruncated deleted the Sales evaluations of every agent who shared the Interaction ID; it now skips an interaction that more than one agent shares (logged, repair by hand), deletes only rows with the same SAP ID, and saves for that SAP ID. DASHBOARD. A Refresh or background reload put every filter pill back on All; the chosen pills are kept (a value no longer in the data drops out). A slow older reply could overwrite a newer one; each load is numbered and an older reply is dropped. A failed Refresh left the spinner on. One failed call of three was counted twice, so the page could render before the others answered. Repeats and Sales shared one counter for their AI insights, so starting one cancelled the other for good; each tab has its own. Every tick of a filter box sent its own AI requests (three per tick); the request now waits 800 ms after the last tick. STANDALONE DASHBOARD. Dashboard.html used midnight UTC for the From date (cutting off the first hours of the day east of UTC) and let rows with no date through; it now follows index.html (local midnight, undated rows left out while a date filter is on). Its CSV export neutralises cells that start with = + - @. ENTRY FORM. A second pasted transcript kept the first Interaction ID, start time and duration; a field that still holds what an earlier transcript put there is now replaced (what the person typed is never touched, and a transcript without the line clears the old value). sendAuditEmail looks the recipients up with the SAP ID of the saved audit even if the form SAP ID was blanked. Covered by tests/admin-tools-test.js, tests/dashboard-refresh-test.js, tests/dashboard-standalone-test.js and tests/form-autofill-test.js (new), new cases in tests/export-test.js, tests/scorecard-slides-test.js, tests/ai-reliability-test.js and tests/email-relay-test.js, and 22 mutation checks, all caught (two were missed at first and led to two more tests). NOT done: batch 4 (accessibility); the other medium items of the review that need the real Google environment to confirm.',
      'Config.gs, EmailGuard.gs, ScorecardSlides.gs, ExportService.gs, AIService.gs, Code.gs, index.html, Dashboard.html, tests/admin-tools-test.js, tests/dashboard-refresh-test.js, tests/dashboard-standalone-test.js, tests/form-autofill-test.js, tests/export-test.js, tests/scorecard-slides-test.js, tests/ai-reliability-test.js, tests/email-relay-test.js, tests/lib/gas-fakes.js',
      'Danzen',
      'Partly done'
    ],
    [
      '173',
      '2026-10-14',
      'Fix',
      'NH + TI',
      'App-wide code review, batch 4 (accessibility): the main page and the dashboard could not be used well with a keyboard or a screen reader, and much of the text was too small or too pale.',
      'FORM AND NAMES. The 12 labels of the entry form (SAP ID, Team Member, Team Leader, Operations Manager, Line of Business, Locale, VTID, Transcript Start Time, Customer BAN, Interaction ID, Direction, Duration) had no for, so no field was read with its label and a click on the label did nothing; each is tied to its field now. The speaker select, the transcript box, the date boxes, the SAP and pill search boxes, the chat box, the analytics email box and the icon-only buttons (close, chat, send) have an accessible name. A scan of the live page found no unnamed control. MESSAGES. An error is announced as an alert, a warning and success as a status, and the progress overlays (processing, email, export) as live status. DIALOGS. The Process Call Script dialog, the Audit Type and Line of Business choices, the speaker confirmation and the analytics panel have a role, a name, and for the modal ones aria-modal. A small layer in index.html moves the focus into a dialog when it opens, keeps Tab inside it, closes the ones that may be closed with Escape (the audit type and line of business choices need an answer, so Escape does nothing there) and puts the focus back where it was. FILTERS AND TABS. The filter pills are keyboard buttons (Tab, Enter, Space, aria-expanded, Escape closes and returns the focus); the Show only link of each name is a focusable button with a name; the four dashboard tabs are a tab list with aria-selected and arrow, Home and End keys. The small clear mark in a pill stays a mouse shortcut (hidden from screen readers) because a button cannot hold another button; the keyboard clears from the list with Select All and Deselect All. REPORT AND EVALVIEW. The Yes or No Issue Resolved badge is a focusable button whose name says its state and works with Enter and Space; the EvalView stars are a radio group with arrow, Home, End, Enter and Space keys and aria-checked, and the page has a language. LOOK. Every visible keyboard focus gets a 3 pixel purple ring with a white halo (inputs had the outline removed and nothing in its place). Text under 12 pixels (88 places) is 12 pixels; the grey text colours 888, 999, aaa and bbb (under 4.5 to 1 on white) are 6B6B6B. Checked in Chrome with Lighthouse: Accessibility 90 before, 100 after on desktop, on mobile and with the dashboard open; the layouts were looked at in screenshots (form and dashboard, wide and narrow) and no horizontal scroll was added. Covered by tests/accessibility-test.js (new: static rules, the keyboard layer, the badge, the stars) and 21 mutation checks, all caught. NOT done: the standalone Dashboard.html was not given this treatment; the report text generated by the AI (sharedCSS in AIService.gs) keeps its own sizes; Lighthouse and the browser checks ran on the page with a fake server, not on the deployed web app, and a real screen reader (NVDA or VoiceOver) was not used.',
      'index.html, EvalView.html, tests/accessibility-test.js, tests/dashboard-broken-features-test.js',
      'Danzen',
      'Partly done'
    ],
    [
      '174',
      '2026-10-14',
      'Fix',
      'NH + TI',
      'The admin ran diagnoseDashboardGaps on the live data and pasted the log; this entry acts on what it showed.',
      'WHAT THE RUN SHOWED. Duration: 480 cells in Audit_Log, 480 in Dashboard_Data and 490 in Transcripts are times of day that applyDataRepairs can turn back into text (the repair is already built; it has to be run), and 1196 more hold a Date that is not on day zero, which the repair leaves alone and the log could not explain. Email Status: 507 audits are Not Sent and none of them has recipients written, so they were saved but never emailed (the biggest observers are Danzen Palisoc with 100, which are test audits, then Wiam Aherchach 86), not a data fault, nothing to repair. Issue Resolved: 1473 clear verdicts, 105 blank, and 125 with no verdict (1 leaked report HTML, 124 sentences). WHAT CHANGED. (1) The Issue Resolved rule now also reads a sentence that begins with its verdict: Issue not resolved, Issue was ultimately not resolved and Neither issue was resolved count as No, and Issue resolved, Primary issue resolved and Issue was ultimately resolved count as Yes; a Yes is not taken from a sentence that goes on to qualify it (but, however, although, except, remains, secondary issue), and unclear openings such as Technically resolved, Functionally resolved or Service was restored are still left for a person. The same rule is in the four places that hold it (ScorecardMetrics.gs, AuditCheck.gs, index.html and Dashboard.html, one shared helper block) and a parity test runs the same 28 examples through all of them. Nothing in the sheets is rewritten: the dashboard, the scorecard and verifyDashboardData read the sentence differently, so the No verdict count goes down by the clear ones. (2) diagnoseDashboardGaps now says where the 1196 Date cells that are not on day zero are (by sheet), on which day they fall (top 6) and shows 8 samples, so the cause can be seen on the next run. NOT done: the Not Sent audits (nothing to repair); the long tail of unclear sentences (about 100); the 1196 Date cells until the new breakdown has been read; the 1 leaked HTML cell.',
      'ScorecardMetrics.gs, AuditCheck.gs, index.html, Dashboard.html, DataRepair.gs, tests/scorecard-metrics-test.js, tests/data-repair-test.js',
      'Danzen',
      'Partly done'
    ],
    [
      '175',
      '2026-10-14',
      'Fix',
      'NH + TI',
      'The second run of diagnoseDashboardGaps showed that the Duration repair built earlier would have written wrong call lengths, and it is corrected before anyone runs it.',
      'WHAT THE RUN SHOWED. 1196 Duration cells hold a Date on 1899-12-31 or 1900-01-01 (for example 07:53, 03:41, 01:18), and the 480 cells per sheet that were on day zero all had 00 seconds (11:39:00, 13:12:00, 18:43:00). Together this means the call lengths were typed as MM:SS (11:39 is 11 minutes 39 seconds, from the Transcript Duration line of the transcript) and Sheets read them as hours and minutes: 11:39 became 11 h 39 min, and 31:53 became a day and 7 h 53 min. The repair built earlier turned a day-zero Date back to H:mm:ss, so 11:39:00 would have been written as 11 hours 39 minutes, which is wrong. It was not run. WHAT CHANGED. The text that is written back is now the one the saved transcript says: the Transcripts sheet keeps the whole transcript of each audit, and its Transcript Duration line (MM:SS or H:MM:SS) is read and used, only when it explains the Date exactly (so a length that the analyst typed over by hand is never overwritten). A Date that has seconds (for example 2:15:05) cannot be a mistyped MM:SS and is still written as H:mm:ss. A Date with no seconds and no length in the saved transcript is NOT guessed: it is left as it is and listed with the reason. The preview and the diagnosis use the same rule and say how many cells come from the transcript, how many are three-part lengths, how many are left and why (no length in the transcript, or the transcript says something different), by sheet and by day with samples. NOT done: applyDataRepairs has not been run on the real data; the first step is previewDataRepairs and a look at its numbers. Cells that stay Dates after that need a decision. Covered by new cases in tests/data-repair-test.js (the MM:SS case end to end in three sheets, a hand-typed length, no transcript length, the diagnosis) and 7 mutation checks, all caught.',
      'DataRepair.gs, tests/data-repair-test.js',
      'Danzen',
      'Partly done'
    ],
    [
      '176',
      '2026-10-14',
      'Fix',
      'NH + TI',
      'The third run of diagnoseDashboardGaps confirmed the Duration plan and left 42 cells that the saved transcript contradicts; a read-only check is added to find out why.',
      'WHAT THE RUN SHOWED. 2604 Duration cells can be written back (862 in Audit_Log, 862 in Dashboard_Data, 880 in Transcripts): 2601 from the length written in the saved transcript and 3 three-part lengths. The samples agree with the MM:SS explanation: 11:39 on day zero becomes 11:39, 07:53 on 1899-12-31 becomes 31:53, 03:41 becomes 27:41, 01:18 on 1900-01-01 becomes 49:18. 42 cells (14 per sheet) are left alone: 3 have no length in the saved transcript and 39 hold a length that differs from the transcript (for example the cell says 4:09 and the transcript says 1:24:37). The suspicion is the form bug fixed in batch 3: pasting a second transcript kept the first transcript duration in the form, so the saved cell holds the length of an earlier paste. WHAT CHANGED. diagnoseDashboardGaps now also says, for each contradicted cell, whether it equals the length of a DIFFERENT audit transcript (and which audit), and counts them. Nothing is written by this entry. NOT done: no repair of the 42 cells until that count is read: if nearly all equal another audit length the transcript is the truth for them and the repair can be extended; otherwise they stay as they are. applyDataRepairs has still not been run: next is previewDataRepairs, then apply after the preview is read. Covered by new cases in tests/data-repair-test.js and 2 mutation checks (1 caught, 1 an equivalent case that cannot be reached).',
      'DataRepair.gs, tests/data-repair-test.js',
      'Danzen',
      'Partly done'
    ],
    [
      '177',
      '2026-10-14',
      'Feature',
      'NH + TI',
      'The admin is about to run applyDataRepairs, which writes 2604 Duration cells, and asked for the backup copy to be made first.',
      'A new admin function, backupSpreadsheet in DataRepair.gs, makes a plain Drive copy of the data spreadsheet in the same folder as the original, named with the sheet name, BACKUP, and the date and time (for example NH Call Analyzer Data - BACKUP 2026-10-14 02-05), and returns and logs its name and link. It holds every sheet, note and format as they are at that moment. The original is not renamed, shared or trashed. It is admin only (requireAdmin_ first, listed in the maintenance group of the access test). If Drive refuses the copy the error is raised on purpose, so nobody goes on to repair without a backup. I cannot open Drive from the developer side, so the admin runs it once in the Apps Script editor. Covered by tests/backup-test.js (new) and 5 mutation checks, all caught. NOT done: the copy itself, which only the admin can make by running backupSpreadsheet; applyDataRepairs waits for it.',
      'DataRepair.gs, tests/backup-test.js, tests/access-control-test.js',
      'Danzen',
      'Partly done'
    ],
    [
      '178',
      '2026-10-14',
      'Fix',
      'NH + TI',
      'The admin sent a screenshot of the live dashboard after the Duration repair; reading it showed that the AI summary quotes different numbers from the tiles.',
      'WHAT THE SCREENSHOT SHOWED. The dashboard loads with 1709 evaluations, 1480 Repeats, 229 Sales, 71 percent average repeat risk and 5149 critical flags, and the Highlights and Lowlights say only 23 percent resolution rate (392 of 1708). The tiles and verifyDashboardData say Resolved 430 of 1709. The cause is buildAnalyticsContext_ in Code.gs, the text given to the AI: it counted an audit as resolved when the cell contained the letters yes anywhere (so Partially resolved - yes the credit posted counted as resolved) and missed a sentence that begins with its verdict (Issue resolved - ...), and it read Repeat Projection % with parseFloat. WHAT CHANGED. The AI context now uses the dashboard rules, the same functions as the scorecard (_scResolution_ and _scNum_): Issues Resolved counts only a real Yes, an extra line gives all four answers (Resolved, Partially resolved, Unresolved, Not recorded) as the tiles show them, and the average Repeat Risk reads a fraction, a sentence that starts with a percentage and a real 0 the way the dashboard does. The AI summaries are cached, so the old wording stays until the cache expires or a filter changes. Covered by tests/analytics-context-test.js (new, the first test of this function) and 4 mutation checks, all caught (the first run missed one because two counts happened to be equal, so the fixture was changed). NOT done: the Flags / Violation Breakdown doughnut on the screenshot is almost all grey Other and was not investigated (the screenshot is too small to read); 599 unique agents for 1709 audits was not checked.',
      'Code.gs, tests/analytics-context-test.js',
      'Danzen',
      'Partly done'
    ],
    [
      '179',
      '2026-10-14',
      'Fix',
      'NH + TI',
      'After the new version was deployed the dashboard summary still said 23 percent resolution (392 of 1709) while the tiles and verifyDashboardData say 430 resolved; a read-only check is added to find out whether the web app runs old code.',
      'WHAT THE SCREENSHOT SHOWED. After the fix of entry 178 and a new version and a hard refresh, the Lowlights still say a 23 percent resolution rate (392 of 1709). 392 is the number the old counting rule gave (any cell that contains the letters yes); the new rule should give 430 (verifyDashboardData). The text changed only in its denominator (1708 became 1709), which fits a new summary written by the AI from the old numbers, so the web app is probably still running an older version of the code than the one in the editor. This is not proven. WHAT CHANGED. New admin function diagnoseAnalyticsContext in DataRepair.gs, read only: from the real sheets, with the code that is in the editor now, it logs the Issue Resolved counts from Audit_Log alone, the counts as the dashboard page builds its rows (the Audit_Log answer, else the Dashboard_Data one), the count the old way, and the lines that the AI is given (TOTAL INTERACTIONS, Issues Resolved, all four answers, Average Repeat Risk). How to read it: if the new count is 430 and the old way is 392 then the editor code is right and the web app deployment is not on the newest version (Deploy, Manage deployments, edit, New version, Deploy; the summaries are cached, so one filter change forces a fresh one); if the new count is already 392 the numbers themselves differ and the cause is elsewhere. Covered by tests/diagnose-context-test.js (new). NOT done: nothing is changed in the data; the cause is not known until the function has been run.',
      'DataRepair.gs, tests/diagnose-context-test.js, tests/access-control-test.js',
      'Danzen',
      'Partly done'
    ],
    [
      '180',
      '2026-10-14',
      'Feature',
      'NH + TI',
      'The admin asked for the look of a collaborator project (Agent Trainer Scores) for the Analyzer: a Home page, tabs in the order Home, Overview and Analyzer Form (with Repeats, Sales and Audit Log as top-level tabs too), the purple and neon green colour scheme, and the loading screen. Every function and every number had to stay the same.',
      'LOOK. The colours were measured from the three reference screenshots (nothing guessed): navigation purple 6F008E, open tab 811F9C with neon green 00EB00, deep purple 57006E for titles, labels and buttons, hero gradient from 440058 through 6D008B to a green glow 47C50F, page F5F3F7, chip tint F8E0FF, tier tints rose, yellow and mint with their dark texts, and the Inter font with the system fonts as fallback. They are tokens in :root, the TELUS token names were kept, and the hard-coded purple, dark purple, green and lilac in the page were moved to the new values; the report documents (print window, emailed report style) keep their own TELUS look. A test checks 26 pairs of colours for 4.5 to 1 (the reference green 008F2D on white is only 4.2 to 1, so text uses 00802F and 007332, and the muted grey was darkened a little). NAVIGATION. One top bar with six tabs, Home, Overview, Repeats and Transfer, Sales, Audit Log and Analyzer Form, as a real tab list (roving tabindex, arrows, Home and End, aria-selected, SVG icons). A showView router shows one view at a time and reuses openDashboard, closeDashboard and dbSwitchTab; the form page and the dashboard are hidden and never removed, so nothing typed and no filter is lost when a tab is left. The dashboard sits under the bar; its old back button and tab bar are gone from sight (the tab bar stays in the page, hidden, because dbSwitchTab and the tests use it); Process Script lives on the form view. The page opens on Home and still fetches the dashboard data at once, so Overview opens ready. Charts that were drawn while the panel was hidden have no size, so coming to the dashboard from Home or the form clears the render memo. HOME. A hero with live counts (evaluations, agents, Repeats and Sales, latest audit, when the data was produced) taken from the rows the dashboard already holds and counted the way the Overview tiles count them, five Choose a Tab cards, and four How the Analyzer Works cards. What the page says is checked against the code: the targets against SCORECARD_THRESHOLDS, the browser copy ages against the LS constants, the deck limit against SC_MAX_AGENTS. On narrower screens the green glow gives way to plain purple, because white text on the glow reads only 2.4 to 1. LOADING SCREEN. Floating bubbles, a green curve and three dots, painted before any script runs, taken away when the page is ready (at least 600 ms, never a flash) and always after 8 seconds, static when the visitor asks for reduced motion, and ignored without scripts; the dashboard loading line uses the same dots. EXISTING SCREENS. Rounder cards, KPI values in tinted boxes (purple, mint, yellow, rose by status), purple label pills on the form cards. Both the form and the dashboard have a main landmark. CHECKED. In Chrome with the real page and a fake backend: every view at 1442 and 502 pixels wide with no page-level horizontal scroll, the loading screen, keyboard arrows on the bar, form text kept while moving between tabs, and Lighthouse accessibility 100 on Home (desktop and mobile), the form, Overview and the Audit Log. Covered by tests/navigation-test.js (new), the contrast test added to tests/accessibility-test.js, and 39 mutation checks on the router, Home and loading screen, all caught. NOT done: Dashboard.html (the standalone page) and EvalView.html keep their old look; the deployed web app was not seen, only the real page with a fake backend; Repeats, Sales and Audit Log were restyled by the shared card and tile rules, not redrawn like a reference picture (there was none); no real screen reader was used.',
      'index.html, tests/navigation-test.js, tests/accessibility-test.js, tests/scorecard-button-test.js, tests/dashboard-broken-features-test.js',
      'Danzen',
      'Partly done'
    ],
    [
      '181',
      '2026-10-14',
      'Fix',
      'NH + TI',
      'The admin approved the Home page and asked for every other tab to be reviewed from screenshots of the live dashboard (1716 audits). Overview, Repeats, Sales and Audit Log each had something to fix; the Analyzer Form had nothing.',
      'FOUND. (1) KPI tiles: the row sat in a wrapper limited to 900 px, so the six tiles of Overview wrapped five and one, and the eight of Repeats and Sales wrapped five and three, however wide the screen. (2) The Flags / Violation Breakdown doughnut gave every flag text its own slice, but nearly all of the 5182 flag texts are different, so the top 8 were tiny and Other was 99 percent grey (98.5 to 99.3 percent on the three tabs). (3) The wide detail tables (Repeats, Sales, Audit Log) were cut at the right with no sign that they scroll, and the first column was lost when scrolling. (4) Chart legends, tooltips and axes were 10 and 11 px. (5) Found while checking: the charts keep the pixel width they were drawn at, so after a zoom or a resize the dashboard stayed wider than the screen (the admin screenshots show different zoom levels). FIXED. (1) The wrapper is as wide as the page content (full width, up to 1280 px, because auto margins in a flex column had shrunk it to its content) and the tiles share the row: flex 1 1 140 px, at most 230 px, so Overview shows 6 in one row, Repeats 8, Sales 7 and Audit Log 5 at 1442 px (measured in Chrome; a first try with an auto-fit grid still fitted only 5, because a grid with a fixed maximum counts its columns by the maximum). (2) The doughnut is now Critical Flags per Audit (the admin chose this): how many audits have no flag, 1, 2, 3 or 4 or more, counted with the same split as the tiles (the helper is outside the shared block that must stay identical to Dashboard.html), with colours from green to dark red and a 12 px legend, on all three tabs; nothing is grouped as Other any more. (3) The table scrollbar is thin, purple and visible, and the first column stays in view with a solid background and a soft edge. (4) Chart text is 12 px. (5) When the window size settles (400 ms) the open dashboard view is drawn again: the old drawings are removed first (they would keep the columns wide) and the chart columns may shrink; measured in Chrome from 1442 to 1000 px: panel width 1255 over 986 before, 986 over 986 after. CHECKED. In Chrome with 320 audits with all-different flag texts: tile rows, the new doughnut, the table at a scrolled position, and Lighthouse accessibility 100 on Repeats and Overview. Covered by tests/dashboard-layout-test.js (new) and 26 mutation checks, all caught (the static test had not seen the grid fault and the shrunk wrapper; the browser did, so the real page was measured). NOT done: the emoji icons on the KPI tiles stay; Audit Log rows with no Interaction ID or Team Member are data, not design; Dashboard.html and EvalView.html keep the old look; the deployed app was not seen.',
      'index.html, tests/dashboard-layout-test.js',
      'Danzen',
      'Partly done'
    ],
    [
      '182',
      '2026-10-14',
      'Fix',
      'NH + TI',
      'A collaborator said the purple in the reference screenshot is a TELUS colour and sent the TELUS palette table. The admin asked to check this and to change colours only, keeping the design.',
      'FOUND. The palette table is the real TELUS palette (purples DFD5E5 to 4B286D, greens DBF9BD to 248700, greys 54595F, 6A6E74, 9DA5AA). The colours used in entry 180 were read as pixels from screenshots taken on a wide-gamut screen, which records the same colours more saturated, so they were not the CSS values (a fit against the palette gave 4B286D, 6E3D8F and 66CC00). FIXED. Only colours changed, not layout. Navigation and deep purple are 4B286D, the open tab is 6E3D8F with white text (the green on 6E3D8F is only 3.7 to 1) and a green underline, the accent and glow are 66CC00, text green is 248700, body grey 54595F, muted grey 6A6E74, tints DFD5E5 and DBF9BD; the dark hero start 3A1F57, the tier tints and the status reds and yellows are not in the table and were kept as derived shades. Found by Lighthouse after the change: the Highlights title 248700 on the mint box F0FFF4 was 4.47 to 1, so it uses the dark green 094218 (Overview and Analyzer). CHECKED. In Chrome with the real page and a fake backend: Lighthouse accessibility 100 on Home, Overview, Repeats, Sales, Audit Log and the form; the whole suite passes. Covered by the palette guard (Test 7) in tests/accessibility-test.js, which fails if a brand token leaves the table or a sampled colour returns, and by 10 mutation checks, all caught. NOT done: it looks darker and less vivid on a normal screen than the reference did on a wide-gamut one; Dashboard.html and EvalView.html keep the old look; the deployed app was not seen.',
      'index.html, tests/accessibility-test.js, tests/dashboard-layout-test.js, tests/navigation-test.js, tests/scorecard-button-test.js',
      'Danzen',
      'Partly done'
    ],
    [
      '183',
      '2026-10-14',
      'Fix',
      'NH + TI',
      'On the deployed web app the title row of Overview, Repeats and Transfer, Sales and Audit Log (the title, the Live line and the Refresh button) sat tight under the navigation bar, with the top of the row hidden behind the bar. The admin sent full screenshots after deploying the version with more padding and said it was not the zoom.',
      'FOUND. In the screenshots the page sits about 16 px lower than the top of the window (white space above the navigation bar), and the hidden part of the title row is also about 16 px. The navigation bar is in the page flow, but the dashboard panel was position fixed at 56 px from the top of the window, so when the page sits lower the bar covers the top of the panel. It was reproduced in Chrome by giving the page a 16 px top margin: bar 16 to 72, title row 56 to 121. The earlier fix of entry 182 (more padding) could not help, and an attempt to measure the bar height was dropped because the bar height is the variable itself. The 16 px come from outside this page (the frame the web app runs in); this is not proven. FIXED. The dashboard panel is now in the page flow right under the bar (position relative, no inset, the page scrolls instead of the panel), and the title row sticks under the bar (top is the bar height). Measured in Chrome for all four tabs with a 0 and a 16 px offset: the title row always starts at the bottom of the bar, the bar and the row stay put while scrolling, no sideways scroll, Home and the form still hide the panel. Found by Lighthouse in the same pass: the green text 248700 on the mint chips EDF7E6 (Sent, Sales) was 4.19 to 1, now 1A5E00; accessibility 100 on Sales and Audit Log, Overview and Repeats unchanged by this. Covered by tests/navigation-test.js (panel in the flow, sticky row), tests/accessibility-test.js (chip colour), and mutation checks on both rules, all caught. NOT done: the deployed app was not seen after this change; the source of the 16 px was not found; Dashboard.html and EvalView.html keep the old look.',
      'index.html, tests/navigation-test.js, tests/accessibility-test.js',
      'Danzen',
      'Partly done'
    ],
    [
      '184',
      '2026-10-14',
      'Fix',
      'NH + TI',
      'App-wide code review of the dashboard data, asked by the admin: every number must match across the tabs, the exports and the AI text, also while filters are applied.',
      'METHOD. A new test, tests/filter-consistency-test.js, loads 150 generated audits through the real path (Audit_Log merged with Dashboard_Data) with the awkward cases (blank Audit_Log cells that Dashboard_Data fills, names with stray spaces, a flag written twice, None, the old comma format, dates on the edge of a day, audits with no date or no Dashboard_Data row), applies 23 filter combinations with the real filter code, and compares each tab, tile, table, chart, export, the scorecard selection, Home and the text the AI is given with an answer worked out in the test from the raw data (about 1450 checks). FOUND AND FIXED. (1) The AI Highlights and Lowlights told the AI that every tab was Repeats and Transfer, also Sales and the Overview, and always said the date range was All time. The page now sends the filters that are applied and the tab, and the server names the analyzer (Sales, Repeats and Transfer, or both together). (2) The AI text counted flags with a plain comma split, so its counts differed from the tiles and charts; it now uses the same splitter, spellings that differ only in capitals or spaces are one flag, and agent names are cleaned the same way. (3) The Audit Log tab filtered the bare Audit_Log row, while Repeats and Sales use the merged row, so an audit with a blank Team Member, LOB or date in Audit_Log was counted for the agent on Repeats and Sales but missed by the same filter on the Audit Log; blank cells are now filled from Dashboard_Data the same way. (4) Unique Agents and Observers on the Audit Log counted a name with stray spaces as another agent (13 against 8 in the test data); counted now like the Overview. (5) Flagged only export on the page, and the flagged count of the standalone Dashboard.html, counted a cell that says None as flagged; now the same rule as the tiles. (6) A flag written twice in one audit counted twice in some places and once in others; the splitter now drops the repeat in all three copies (index.html, Dashboard.html, ScorecardMetrics.gs, kept identical by the parity test). (7) The AI Analytics panel cut the first day at midnight UTC, 8 hours off from the dashboard in Manila; it now uses local midnight and leaves out an audit with no readable date once a date is chosen, as the tabs do. (8) The Repeats and Sales tile that counts audits with a flag was called Critical Flags, the same name as the Overview tile that counts flags (1211 against 5223 in the live screenshots); it is now called Audits with Flags. (9) Dead code that could mislead: a second copy of the Audit Log filter (dbApplyAuditLog, with other rules, now calls the shared filter) and an older duplicate of dbRenderAuditStats. CHECKED. The whole suite passes; 11 mutation checks on the fixes, all caught (one was missed at first, the choice of the analyzer inside the real server function, and a test was added); in Chrome with 320 audits the Overview 320 equals Repeats 267 plus Sales 53, Emailed 214 on every tab, 10 agents on the Overview and the Audit Log. NOT done and by design: Home counts all audits and ignores filters (it matches the Overview with no filter); the Audit Log can hold audits with no type that Repeats and Sales cannot show (the page says how many); the real sheets were not seen, run verifyDashboardData for those; Dashboard.html and EvalView.html keep the old look.',
      'index.html, Code.gs, Dashboard.html, ScorecardMetrics.gs, tests/filter-consistency-test.js, tests/dashboard-correctness-test.js',
      'Danzen',
      'Partly done'
    ],
    [
      '185',
      '2026-10-14',
      'Fix',
      'NH + TI',
      'Second app-wide code review, asked by the admin, on three new areas: input safety, the write path of an audit, and scale. Three read-only reviews traced the code; every finding below was checked in the code before it was fixed.',
      'FIXED. (1) A call longer than 50000 characters made the save of the transcript row throw (a Sheets cell holds 50000, the AI takes 150000), so the whole submit failed with a raw Google error and a retry failed the same way; the saved copy is now cut at 49000 with a note, the AI still gets the whole text, and nothing reads the saved text except its first line (the Duration) and whether it is empty. (2) getAIAnalytics wrote the analysis type it was sent into the AI_Analytics sheet with a plain appendRow, so any signed-in user could put a formula there (for example IMPORTDATA) that runs in the owner spreadsheet; the type is now reduced to sales or repeats and the write goes through the safe appendRow wrapper. (3) sendAnalyticsBriefing sends from the owner mailbox and had no hourly limit (sendAuditEmail has one); it now counts against the same limit. (4) The AI Highlights, Lowlights and analytics were sent whole rows (about 35 fields with long texts, for 1700 audits, three calls per settled filter) when the text for the AI uses ten fields; the page now sends those ten (the Call Summary cut at 200 characters) and a test proves the text for the AI is identical with whole rows and with ten fields. (5) The SAP ID box redrew every table and chart on every key; it now waits 250 ms after the last key. (6) A click in a filter list searched the whole store once per item (the Audit Ref list has one item per audit, about 1725 times 1725 searches); it now uses one lookup table. CHECKED. The whole suite passes; tests/review-round2-test.js (new, 28 checks) and 11 mutation checks, all caught (one was missed at first because the test used its own copy of the limit; it now reads the real one). FOUND, NOT FIXED (the admin should decide, each is a larger change). Write path: a Cache sheet write that fails is swallowed, so the evaluation link can show an empty report after 6 hours; an edit on send deletes the old Cache rows before it writes the new ones; the three audit rows (Analysis, Dashboard_Data, Audit_Log) are written one by one and a failed middle write is never repaired on resubmit (the dashboard under-counts that audit); audit saves do not take the script lock that the sheet repair functions take; the NH and TI scripts share one spreadsheet but the lock is per script; sendAuditEmail has no check that the audit is already Sent; a fallback audit ref can collide with a later numbered one after a lock timeout; the audit ref date uses the script time zone (New York), so a Manila morning audit gets the day before. Scale: the dashboard cache is stored in 30 KB chunks, one call each (hundreds per open); the Audit_Log and roster caches never hold because they only cache under 99000 characters; the data is fetched again on each dbHandleData; every filter change redraws all three tables; a pill click draws twice; the AI_Analytics sheet is never pruned and is read whole; localStorage quota errors are swallowed; the evaluation link reads whole sheets; the export refuses more than 5000 rows (about 165 days away). Security, lower: sendAnalyticsBriefing can mail any telus address as the owner (now rate limited, not tied to the caller); the analytics cache key does not include the row content, so fabricated rows of the right count can fill a cached report for the same filters; the cache key is cut at 250 characters. Checked and safe: doGet parameters, every server HTML builder, the sanitiser, the client innerHTML paths, window.open, the exports, the column order of every sheet write, the double submit flag, the audit ref counter, the email quota checks.',
      'Code.gs, index.html, tests/review-round2-test.js, tests/server-input-hardening-test.js, tests/submit-recovery-test.js, tests/audit-check-test.js',
      'Danzen',
      'Partly done'
    ],
    [
      '186',
      '2026-10-14',
      'Fix',
      'NH + TI',
      'The admin asked that the NH and TI Analyzers always stay in step, and for an app-wide review of the TI version after everything was mirrored to it.',
      'FOUND. The TI project (separate Apps Script, separate GitHub repo) had stopped receiving the NH changes after 29 September: 12 files were missing and index.html differed by about 13000 lines, but the code is meant to be identical (Config.gs picks the TIA audit prefix by script id), so every difference was old code, not a TI feature. FIXED. All .gs, .html, tests and docs of NH were copied to TI, the whole suite passes there, and it was pushed with clasp. After the copy every file compares identical, except appsscript.json: the TI manifest still said USER_ACCESSING while the code assumes the web app runs as the owner (isAdminCaller_, the owned folders, exports, mail), as the NH manifest says since entry of 30 September. The admin confirmed the TI deployment is Execute as Me, so the TI manifest now says USER_DEPLOYING and the two manifests are identical. RULE. Every change pushed to NH is mirrored to TI and pushed with clasp (the admin makes the versions). STILL TRUE FOR BOTH, because they share one spreadsheet: the lock is per script, so an NH save and a TI save can run at the same time; the Script Properties are per project (ANTHROPIC_AUTH_TOKEN, REGISTRY_SECRET, EMAIL_EXTRA_DOMAINS) and are not copied by clasp. NOT done: the TI Script Properties were not seen; the deployed TI app was not seen.',
      'appsscript.json (TI), every file mirrored to TI',
      'Danzen',
      'Partly done'
    ],
    [
      '187',
      '2026-10-26',
      'Feature',
      'NH + TI',
      'The admin asked for a new column, Auditor Email, in Dashboard_Data holding the email of whoever submitted each audit (it was first asked as a new tab, then clarified as a column in the same tab).',
      'DONE IN THE CODE. Auditor Email is the LAST header of Dashboard_Data, after Agent Email, because the sheet is written by position and addMissingHeaders only adds at the end. Each new audit stores the signed-in caller as the server reads it (never the observerEmail the browser sends); when Google hides the caller, the Users sheet address for the observer name; a name with several different addresses gets all of them separated by a comma (the admin chose this on 2026-10-26, so one cell can hold two addresses, and it may be two people with one name); blank when the name is not on the Users sheet. A failed lookup never stops the save. backfillAuditorEmail() (DataRepair.gs, admin, run from the editor) fills blank cells of existing rows from the Users sheet by Observer Name, never overwrites an email, writes only that column, and lists the names with no match and those that got several addresses. FIRST RUN (2026-10-26): 1645 of 2003 rows filled; 358 were left blank because the name was on the Users sheet more than once (Danzen Palisoc 231, Erma Jomolo 49, Grace Landaverde 29, Sheila Danipog 23, Dulce Lorena Garcia Ramirez 8) or not on it (yasmin.hernandez02 10, alizon.avila 5, rabab.loussahi 2, Leidy Mariela Araujo Colocho 1); the rule for several addresses was changed afterwards and the run is repeated. The column also reaches the dashboard in the browser (readDashboardSheet_ sends every column); the admin chose to show it. DONE BY THE ADMIN IN THE SHEET. addMissingHeaders() added the header (column 38); backfillAuditorEmail() ran twice (1645, then 340 rows); the 18 rows of the four names that are not on the Users sheet were typed by hand, and the admin checked that none is blank. TI. NH and TI share one spreadsheet, and until the TI script had this code the audits saved by TI left Auditor Email blank, so backfillAuditorEmail() has to be run once more after the TI Version that has it (it only fills blank cells). The code was mirrored to TI (rule of entry 186). NOT done: the page was not opened in a browser; no audit was submitted after the change to see the column fill; the Users sheet was not seen by the code author (the addresses of the names with several were not checked).',
      'Config.gs, Code.gs, ObserverHelper.gs, DataRepair.gs, tests/auditor-email-test.js, tests/data-repair-test.js, DOCUMENTATION.md',
      'Danzen',
      'Partly done'
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

