/**
 * Code.gs
 * Web app entry point and server-side handlers.
 */

// ── Run from GAS editor: find + repair truncated Sales evaluation HTML ────────
// Step 1 — diagnose: shows which Sales cache entries look truncated
function diagnoseSalesEvaluations() {
  var ss         = getOrCreateSpreadsheet();
  var cacheSheet = getOrCreateSheet(ss, CACHE_SHEET);
  var lastRow    = cacheSheet.getLastRow();
  if (lastRow < 2) { Logger.log('Cache sheet is empty.'); return; }

  var data      = cacheSheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var truncated = [];
  var complete  = [];

  // Build a lookup of chunk rows: key = interactionId + '|' + chunkType → html
  var chunkMap = {};
  data.forEach(function(row) {
    var id        = (row[0] || '').toString().trim();
    var atype     = (row[1] || '').toString().trim().toLowerCase();
    var chunkHtml = (row[3] || '').toString();
    if (/^sales_\d+$/.test(atype) && chunkHtml) {
      chunkMap[id + '|' + atype] = chunkHtml;
    }
  });

  data.forEach(function(row, idx) {
    var interactionId = (row[0] || '').toString().trim();
    var atype         = (row[1] || '').toString().trim().toLowerCase();
    var html          = (row[3] || '').toString();
    if (atype !== 'sales') return;
    if (!html) return; // skip cleared/deleted ghost rows

    // Assemble chunks
    var ckNum = 2;
    while (true) {
      var ckKey = interactionId + '|sales_' + ckNum;
      if (chunkMap[ckKey] === undefined) break;
      html += chunkMap[ckKey];
      ckNum++;
    }

    var trimmed       = html.trim();
    var tail          = trimmed.slice(-120);  // last 120 chars
    var hasCallSum    = html.indexOf('Call Summary') !== -1 || html.indexOf('call-summary') !== -1;
    var endsClean     = /(<\/div>|<\/section>|<\/html>|<\/p>|<\/ul>|<\/ol>|<\/table>|<\/span>)\s*$/.test(trimmed);
    var looksComplete = html.length > 500 && hasCallSum && endsClean;

    var entry = {
      row: idx + 2,
      interactionId: interactionId,
      htmlLength: html.length,
      hasCallSummary: hasCallSum,
      endsClean: endsClean,
      tail: tail
    };

    if (looksComplete) {
      complete.push(entry);
    } else {
      truncated.push(entry);
    }
  });

  Logger.log('=== Sales Evaluation Diagnosis ===');
  Logger.log('Complete : ' + complete.length);
  Logger.log('Truncated: ' + truncated.length);
  truncated.forEach(function(e) {
    Logger.log('ROW=' + e.row + ' id=' + e.interactionId.substring(0,8) + '...' +
               ' len=' + e.htmlLength +
               ' hasCallSum=' + e.hasCallSummary +
               ' endsClean=' + e.endsClean);
    Logger.log('  ...tail: [' + e.tail.replace(/\n/g,' ') + ']');
  });
  return { complete: complete.length, truncated: truncated.length, truncatedList: truncated };
}

// Step 2a — repair ONE entry at a time (safe for 6-min GAS timeout).
// Run this function repeatedly until diagnoseSalesEvaluations() shows Truncated: 0.
function repairNextTruncated() {
  var diagnosis = diagnoseSalesEvaluations();
  if (!diagnosis || !diagnosis.truncatedList || !diagnosis.truncatedList.length) {
    Logger.log('✅ All Sales evaluations complete — nothing left to repair.');
    return;
  }

  var entry = diagnosis.truncatedList[0];
  Logger.log('Repairing: row=' + entry.row + ' id=' + entry.interactionId.substring(0,8) + '... ('+
             diagnosis.truncatedList.length + ' remaining)');

  var ss = getOrCreateSpreadsheet();
  var cacheSheet  = getOrCreateSheet(ss, CACHE_SHEET);
  var transcripts = getOrCreateSheet(ss, TRANSCRIPTS_SHEET);

  var tLast = transcripts.getLastRow();
  if (tLast < 2) { Logger.log('No transcripts found.'); return; }
  var tHeaders = transcripts.getRange(1, 1, 1, transcripts.getLastColumn()).getValues()[0];
  var tData    = transcripts.getRange(2, 1, tLast - 1, tHeaders.length).getValues();
  var tIdIdx   = tHeaders.indexOf('Interaction ID');
  var tTxtIdx  = tHeaders.indexOf('Transcript');

  var transcript = '';
  for (var i = 0; i < tData.length; i++) {
    var tid = (tData[i][tIdIdx] || '').toString().trim();
    if (tid === entry.interactionId.trim()) {
      transcript = tData[i][tTxtIdx] ? tData[i][tTxtIdx].toString() : '';
      break;
    }
  }

  if (!transcript) {
    Logger.log('SKIP — transcript not found for row=' + entry.row);
    return;
  }

  var newHtml = analyzeTranscript(transcript, 'sales');
  newHtml = fixBadgeClasses(
    newHtml.replace(/^```html\s*/i,'').replace(/^```\s*/,'').replace(/\s*```$/,'').trim()
  );

  // Delete all old rows for this interaction ID (main + chunks) bottom-to-top
  // so row numbers stay valid while deleting, and no ghost empty rows remain
  var deleteFinder = cacheSheet.getRange('A:A').createTextFinder(entry.interactionId.trim()).matchEntireCell(true);
  var rowsToDelete = [];
  deleteFinder.findAll().forEach(function(match) {
    var rowType = (cacheSheet.getRange(match.getRow(), 2).getValue() || '').toString().toLowerCase();
    if (rowType === 'sales' || /^sales_\d+$/.test(rowType)) {
      rowsToDelete.push(match.getRow());
    }
  });
  rowsToDelete.sort(function(a, b) { return b - a; }); // delete bottom-to-top
  rowsToDelete.forEach(function(rowNum) { cacheSheet.deleteRow(rowNum); });
  SpreadsheetApp.flush();

  // Save with chunking — handles HTML > 48k chars transparently
  saveCachedResult(entry.interactionId.trim(), 'sales', newHtml);
  SpreadsheetApp.flush();

  var chunks = Math.ceil(newHtml.length / 48000);
  Logger.log('✅ Repaired id=' + entry.interactionId.substring(0,8) + '...' +
             ' len=' + newHtml.length + ' chunks=' + chunks +
             ' (' + (diagnosis.truncatedList.length - 1) + ' remaining)');
}

// Step 2b — bulk repair (legacy — use repairNextTruncated() to avoid timeout).
function repairTruncatedSalesEvaluations() {
  var diagnosis = diagnoseSalesEvaluations();
  if (!diagnosis || !diagnosis.truncatedList || !diagnosis.truncatedList.length) {
    Logger.log('Nothing to repair — all Sales evaluations look complete.');
    return;
  }

  var ss          = getOrCreateSpreadsheet();
  var cacheSheet  = getOrCreateSheet(ss, CACHE_SHEET);
  var transcripts = getOrCreateSheet(ss, TRANSCRIPTS_SHEET);
  var analysisSheet = getOrCreateSheet(ss, ANALYSIS_SHEET);

  // Build Interaction ID → transcript row map
  var tLast = transcripts.getLastRow();
  var tMap  = {};
  if (tLast >= 2) {
    var tHeaders = transcripts.getRange(1, 1, 1, transcripts.getLastColumn()).getValues()[0];
    var tData    = transcripts.getRange(2, 1, tLast - 1, tHeaders.length).getValues();
    var tIdIdx   = tHeaders.indexOf('Interaction ID');
    var tTxtIdx  = tHeaders.indexOf('Transcript');
    tData.forEach(function(row) {
      var id = (row[tIdIdx] || '').toString().trim();
      if (id) tMap[id] = row[tTxtIdx] ? row[tTxtIdx].toString() : '';
    });
  }

  var repaired = 0, failed = 0;
  diagnosis.truncatedList.forEach(function(entry) {
    var transcript = tMap[entry.interactionId] || '';
    if (!transcript) {
      Logger.log('SKIP ' + entry.interactionId + ' — transcript not found');
      failed++;
      return;
    }
    try {
      Logger.log('Re-running AI for: ' + entry.interactionId);
      var newHtml = analyzeTranscript(transcript, 'sales');
      newHtml = fixBadgeClasses(newHtml
        .replace(/^```html\s*/i, '').replace(/^```\s*/, '').replace(/\s*```$/, '').trim());

      // Update Cache sheet cell
      cacheSheet.getRange(entry.row, 4).setValue(newHtml);

      // Invalidate ev2_ CacheService entry for all audits referencing this interaction
      var sc = CacheService.getScriptCache();
      Logger.log('Repaired: ' + entry.interactionId + ' (new len=' + newHtml.length + ')');
      repaired++;
      Utilities.sleep(2000); // avoid rate limiting between AI calls
    } catch(e) {
      Logger.log('FAILED ' + entry.interactionId + ': ' + e);
      failed++;
    }
  });

  Logger.log('=== Repair complete: ' + repaired + ' repaired, ' + failed + ' failed ===');
}

// Run once from editor to confirm exact Locale column in roster
function diagnoseRosterColumns() {
  var ss    = SpreadsheetApp.openById(ROSTER_SHEET_ID);
  var sheet = ss.getSheetByName('roster');
  if (!sheet) { Logger.log('roster sheet not found'); return; }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  headers.forEach(function(h, i) {
    var col = String.fromCharCode(65 + i);
    Logger.log('Col ' + col + ' (idx ' + i + '): ' + h);
  });
  // Also log first data row for spot-check
  var row1 = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  row1.forEach(function(v, i) {
    var col = String.fromCharCode(65 + i);
    if (v) Logger.log('  Row2 Col ' + col + ': ' + v);
  });
}

// ── Run this from the editor to see exact column values for a given SAP ID ────
// Change the SAP_ID value below before running
// ── Run this directly from the editor — no deployment needed ─────────────────
// Run this to see ALL columns in AT Data GCP for SAP ID 2007888
// Run to see FCR Dashboard Data sheets and columns
// ── One-time: backfill blank Dashboard_Data columns from Analysis sheet HTML ───
// Run from the editor after pushing the extractTextBlock fix.
function backfillDashboardData() {
  var ss       = getOrCreateSpreadsheet();
  var aSheet   = getOrCreateSheet(ss, ANALYSIS_SHEET);
  var dSheet   = getOrCreateSheet(ss, DASHBOARD_DATA_SHEET);

  var aLast = aSheet.getLastRow();
  var dLast = dSheet.getLastRow();
  if (aLast < 2 || dLast < 2) { Logger.log('No data to backfill'); return; }

  // Build map: Audit Ref → HTML
  // Primary: Cache sheet (Interaction ID → HTML), linked via Audit_Log
  // Fallback: Analysis sheet (if it still has raw HTML)
  var htmlMap = {};

  // Step 1: read Audit_Log to map Audit Ref → Interaction ID
  var logSheet = getOrCreateSheet(ss, AUDIT_LOG_SHEET);
  var logLast  = logSheet.getLastRow();
  var refToId  = {};  // Audit Ref → Interaction ID
  if (logLast >= 2) {
    var logData = logSheet.getRange(2, 1, logLast - 1, 3).getValues();
    logData.forEach(function(row) {
      var ref = (row[0] || '').toString().trim();
      var id  = (row[2] || '').toString().trim();
      if (ref && id) refToId[ref] = id;
    });
  }

  // Step 2: load Cache sheet into a map: Interaction ID → HTML
  var cacheSheet = getOrCreateSheet(ss, CACHE_SHEET);
  var cacheLast  = cacheSheet.getLastRow();
  var idToHtml   = {};
  if (cacheLast >= 2) {
    var cacheData = cacheSheet.getRange(2, 1, cacheLast - 1, 4).getValues();
    cacheData.forEach(function(row) {
      var id   = (row[0] || '').toString().trim();
      var html = (row[3] || '').toString();
      if (id && html.indexOf('<') !== -1) idToHtml[id] = html;
    });
  }

  // Step 3: build Audit Ref → HTML by joining the two maps
  Object.keys(refToId).forEach(function(ref) {
    var id = refToId[ref];
    if (idToHtml[id]) htmlMap[ref] = idToHtml[id];
  });

  // Step 4: also pull any remaining HTML from Analysis sheet (rows not yet cleaned)
  var aData  = aSheet.getRange(2, 1, aLast - 1, ANALYSIS_HEADERS.length).getValues();
  var refIdx  = ANALYSIS_HEADERS.indexOf('Audit Ref');
  var htmlIdx = ANALYSIS_HEADERS.indexOf('AI Analysis Result');
  aData.forEach(function(row) {
    var ref  = (row[refIdx]  || '').toString().trim();
    var html = (row[htmlIdx] || '').toString();
    if (ref && html.indexOf('<') !== -1 && !htmlMap[ref]) htmlMap[ref] = html;
  });

  Logger.log('HTML source map: ' + Object.keys(htmlMap).length + ' entries from cache/analysis');

  // Read Dashboard_Data and fill blank columns
  var dHeaders = dSheet.getRange(1, 1, 1, dSheet.getLastColumn()).getValues()[0];
  var dData    = dSheet.getRange(2, 1, dLast - 1, dHeaders.length).getValues();
  var dRefIdx  = dHeaders.indexOf('Audit Ref');
  var colMap   = {
    'Call Reason':          dHeaders.indexOf('Call Reason'),
    'Call Summary':         dHeaders.indexOf('Call Summary'),
    'Overall Opportunities': dHeaders.indexOf('Overall Opportunities'),
    'SMART Recommendation':  dHeaders.indexOf('SMART Recommendation') > -1
                              ? dHeaders.indexOf('SMART Recommendation')
                              : dHeaders.indexOf('Recommendations'),   // fallback for existing sheets
    'Critical Flags':        dHeaders.indexOf('Critical Flags'),
    'Repeat Projection %':  dHeaders.indexOf('Repeat Projection %'),
    'Issue Resolved':       dHeaders.indexOf('Issue Resolved'),
    'Transfer Occurred':    dHeaders.indexOf('Transfer Occurred')
  };
  var updated = 0;
  dData.forEach(function(row, i) {
    var ref  = (row[dRefIdx] || '').toString().trim();
    var html = htmlMap[ref];
    if (!html) return;
    var changed = false;
    if (colMap['Call Reason'] > -1 && !row[colMap['Call Reason']]) {
      row[colMap['Call Reason']] = extractTextBlock(html, 'Call Reason');
      changed = true;
    }
    if (colMap['Call Summary'] > -1 && !row[colMap['Call Summary']]) {
      row[colMap['Call Summary']] = extractTextBlock(html, 'Call Summary');
      changed = true;
    }
    if (colMap['Overall Opportunities'] > -1 && !row[colMap['Overall Opportunities']]) {
      row[colMap['Overall Opportunities']] = extractTextBlock(html, 'Opportunit');
      changed = true;
    }
    if (colMap['SMART Recommendation'] > -1 && !row[colMap['SMART Recommendation']]) {
      row[colMap['SMART Recommendation']] = extractTextBlock(html, 'Recommendation');
      changed = true;
    }
    if (colMap['Critical Flags'] > -1 && !row[colMap['Critical Flags']]) {
      row[colMap['Critical Flags']] = extractTextBlock(html, 'Critical Flag');
      changed = true;
    }
    if (colMap['Repeat Projection %'] > -1 && !row[colMap['Repeat Projection %']]) {
      row[colMap['Repeat Projection %']] = extractTextBlock(html, 'Repeat');
      changed = true;
    }
    if (colMap['Issue Resolved'] > -1 && !row[colMap['Issue Resolved']]) {
      row[colMap['Issue Resolved']] = extractTextBlock(html, 'Issue Resolv');
      changed = true;
    }
    if (changed) updated++;
  });

  if (updated > 0) {
    dSheet.getRange(2, 1, dData.length, dHeaders.length).setValues(dData);
    Logger.log('Backfilled ' + updated + ' Dashboard_Data rows');
    invalidateDashboardCache();
  } else {
    Logger.log('Nothing to backfill — all rows already have data');
  }
}

// ── RCA enrichment prompt ─────────────────────────────────────────────────────
// Reads cached HTML for each audit, calls AI to extract structured RCA fields,
// writes them back to Dashboard_Data. Run once from the editor after submission.
function enrichDashboardData() {
  var ss      = getOrCreateSpreadsheet();
  var dSheet  = getOrCreateSheet(ss, DASHBOARD_DATA_SHEET);
  var logSheet= getOrCreateSheet(ss, AUDIT_LOG_SHEET);
  var cache   = getOrCreateSheet(ss, CACHE_SHEET);

  // Build Audit Ref → Interaction ID from Audit_Log
  var logLast = logSheet.getLastRow();
  if (logLast < 2) { Logger.log('No Audit_Log rows'); return; }
  var logData = logSheet.getRange(2, 1, logLast - 1, 3).getValues();
  var refToId = {};
  logData.forEach(function(r) {
    var ref = (r[0]||'').toString().trim();
    var id  = (r[2]||'').toString().trim();
    if (ref && id) refToId[ref] = id;
  });

  // Build Interaction ID → HTML from Cache
  var cacheLast = cache.getLastRow();
  if (cacheLast < 2) { Logger.log('No Cache rows'); return; }
  var cacheData = cache.getRange(2, 1, cacheLast - 1, 4).getValues();
  var idToHtml  = {};
  cacheData.forEach(function(r) {
    var id   = (r[0]||'').toString().trim();
    var html = (r[3]||'').toString();
    if (id && html.indexOf('<') !== -1) idToHtml[id] = html;
  });

  // Read Dashboard_Data headers to find column positions
  var dLast   = dSheet.getLastRow();
  if (dLast < 2) { Logger.log('No Dashboard_Data rows'); return; }
  var dHeaders = dSheet.getRange(1, 1, 1, dSheet.getLastColumn()).getValues()[0];
  var colOf    = {};
  dHeaders.forEach(function(h, i) { if (h) colOf[h] = i + 1; }); // 1-based

  var dData    = dSheet.getRange(2, 1, dLast - 1, dHeaders.length).getValues();
  var enriched = 0;

  for (var i = 0; i < dData.length; i++) {
    var row    = dData[i];
    var ref    = (row[colOf['Audit Ref'] - 1] || '').toString().trim();
    var atype  = (row[colOf['Analysis Type'] - 1] || '').toString().trim();

    // Skip only if ALL structured fields are already filled (including Sales Attempted for Sales rows)
    var hasDriver       = colOf['Call Driver']         && (row[colOf['Call Driver']         - 1] || '').toString().trim();
    var hasFlags        = colOf['Critical Flags']      && (row[colOf['Critical Flags']      - 1] || '').toString().trim();
    var hasRepeat       = colOf['Repeat Projection %'] && (row[colOf['Repeat Projection %'] - 1] || '').toString().trim();
    var hasRcaCat       = colOf['RCA Category']        && (row[colOf['RCA Category']        - 1] || '').toString().trim();
    var isSalesRow      = (atype || '').indexOf('Sales') !== -1;
    var hasSalesAttempt = !isSalesRow || (colOf['Sales Attempted'] && (row[colOf['Sales Attempted'] - 1] || '').toString().trim());
    if (hasDriver && hasFlags && hasRepeat && hasRcaCat && hasSalesAttempt) continue;

    var id   = refToId[ref];
    var html = id ? idToHtml[id] : '';
    if (!html) {
      Logger.log('No cached HTML for ' + ref);
      continue;
    }

    Logger.log('Enriching ' + ref + ' (' + atype + ')...');

    // Extract repeat % from HTML using multiple strategies
    var repeatFromHTML = '';

    // Strategy 1: ai-chip-val after ai-chip-label containing "Repeat"
    var chipMatch = html.match(/ai-chip-label[^>]*>[^<]*Repeat[^<]*<\/span>\s*<span[^>]*ai-chip-val[^>]*>(\d+)/i);
    if (chipMatch) repeatFromHTML = chipMatch[1];

    // Strategy 2: any "N% repeat risk" pattern in raw HTML
    if (!repeatFromHTML) {
      var pctMatch = html.match(/(\d{1,3})%\s*(?:repeat|Repeat)\s*(?:risk|Risk|projection|Projection)/i);
      if (pctMatch) repeatFromHTML = pctMatch[1];
    }

    // Strategy 3: "repeatProjectionPct": N in JSON remnants
    if (!repeatFromHTML) {
      var jsonMatch = html.match(/repeatProjectionPct["'\s:]+(\d{1,3})/i);
      if (jsonMatch) repeatFromHTML = jsonMatch[1];
    }

    // Strategy 4: search in plain text version
    if (!repeatFromHTML) {
      var pt = htmlToPlainText(html);
      var ptMatch = pt.match(/(\d{1,3})\s*%\s*(?:repeat|Repeat)/i)
                 || pt.match(/Repeat\s*(?:Risk|Projection)[^\d]*(\d{1,3})/i)
                 || pt.match(/(\d{1,3})%\s*(?:—|repeat|risk|projection)/i);
      if (ptMatch) repeatFromHTML = ptMatch[1];
    }

    if (repeatFromHTML) Logger.log('  Repeat % from HTML: ' + repeatFromHTML + '%');

    try {
      var structured = extractStructuredRCA(html, atype, ref);
      // Prefer HTML-extracted repeat % over AI-extracted (AI gets confused by plain text)
      if (repeatFromHTML && structured) structured.repeatPct = repeatFromHTML;
      if (structured) {
        // Write structured fields back to the sheet
        if (colOf['Call Driver'])             dSheet.getRange(i+2, colOf['Call Driver']).setValue(structured.callDriver);
        if (colOf['RCA Category'])            dSheet.getRange(i+2, colOf['RCA Category']).setValue(structured.rcaCategory);
        if (colOf['RCA Sub-Parameter'])       dSheet.getRange(i+2, colOf['RCA Sub-Parameter']).setValue(structured.rcaSubParameter);
        var topOppCol = colOf['Top SMART Opportunity'] || colOf['Top Opportunity'];
        if (topOppCol)                        dSheet.getRange(i+2, topOppCol).setValue(structured.topOpportunity);
        if (colOf['Product Opportunity'])     dSheet.getRange(i+2, colOf['Product Opportunity']).setValue(structured.productOpportunity);
        if (colOf['Sales Attempted'])         dSheet.getRange(i+2, colOf['Sales Attempted']).setValue(structured.salesAttempted);
        if (colOf['SMART S (Specific)']   && structured.smartS) dSheet.getRange(i+2, colOf['SMART S (Specific)']).setValue(structured.smartS);
        if (colOf['SMART M (Measurable)'] && structured.smartM) dSheet.getRange(i+2, colOf['SMART M (Measurable)']).setValue(structured.smartM);
        if (colOf['SMART A (Attainable)'] && structured.smartA) dSheet.getRange(i+2, colOf['SMART A (Attainable)']).setValue(structured.smartA);
        if (colOf['SMART R (Realistic)']  && structured.smartR) dSheet.getRange(i+2, colOf['SMART R (Realistic)']).setValue(structured.smartR);
        if (colOf['SMART T (Time-bound)'] && structured.smartT) dSheet.getRange(i+2, colOf['SMART T (Time-bound)']).setValue(structured.smartT);
        if (colOf['Call Summary'] && !row[colOf['Call Summary']-1])
          dSheet.getRange(i+2, colOf['Call Summary']).setValue(structured.callSummaryShort);
        // Write critical flags (always overwrite)
        if (colOf['Critical Flags'] && structured.criticalFlags)
          dSheet.getRange(i+2, colOf['Critical Flags']).setValue(structured.criticalFlags);
        // Write repeat % only if current value is blank or 0
        var existingRepeat = (row[colOf['Repeat Projection %'] - 1] || '').toString().trim().replace(/[^0-9]/g,'');
        if (colOf['Repeat Projection %'] && (!existingRepeat || existingRepeat === '0') && structured.repeatPct && structured.repeatPct !== '0')
          dSheet.getRange(i+2, colOf['Repeat Projection %']).setValue(structured.repeatPct);
        enriched++;
        Logger.log('  ✓ ' + structured.callDriver + ' | ' + structured.rcaCategory + ' | Flags: ' + (structured.criticalFlags||'none') + ' | Repeat: ' + (structured.repeatPct||'n/a'));
        Utilities.sleep(500); // avoid rate limits
      }
    } catch(e) {
      Logger.log('Error enriching ' + ref + ': ' + e);
    }
  }

  invalidateDashboardCache();
  Logger.log('=== Enriched ' + enriched + ' rows ===');
}

// ── Backfill blank Locale values in Dashboard_Data using roster lookup ────────
// Run ONCE from the Apps Script editor: open Code.gs → Run → backfillLocaleInDashboard
// Safe to re-run — only patches rows where Locale is currently blank.
function backfillLocaleInDashboard() {
  var ss    = getOrCreateSpreadsheet();
  var sheet = getOrCreateSheet(ss, DASHBOARD_DATA_SHEET);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('backfillLocale: sheet is empty'); return; }

  var headers  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var localeCol = headers.indexOf('Locale') + 1;   // 1-based
  var sapCol    = headers.indexOf('SAP ID') + 1;
  if (localeCol < 1 || sapCol < 1) {
    Logger.log('backfillLocale: could not find Locale or SAP ID column'); return;
  }

  var data    = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var patched = 0;
  var skipped = 0;

  for (var i = 0; i < data.length; i++) {
    var locale = (data[i][localeCol - 1] || '').toString().trim();
    if (locale) { skipped++; continue; }   // already has a locale

    var sapId = (data[i][sapCol - 1] || '').toString().trim();
    if (!sapId) { skipped++; continue; }   // no SAP ID to look up

    try {
      var rosterRow = lookupBySapId(sapId);
      if (rosterRow && rosterRow.locale) {
        sheet.getRange(i + 2, localeCol).setValue(rosterRow.locale);
        patched++;
        Logger.log('backfillLocale: row ' + (i + 2) + ' SAP ' + sapId + ' → ' + rosterRow.locale);
      } else {
        skipped++;
        Logger.log('backfillLocale: no locale found for SAP ' + sapId);
      }
    } catch(e) {
      skipped++;
      Logger.log('backfillLocale: lookup failed for SAP ' + sapId + ': ' + e);
    }
  }

  invalidateDashboardCache();
  Logger.log('=== backfillLocale done: ' + patched + ' patched, ' + skipped + ' skipped ===');
}

// ── Fast HTML-only RCA extraction (no AI call — used in submitTranscript) ────────
// Parses the AI-generated HTML directly instead of making a second API call.
// Saves 5–30 seconds on every submission.
function extractStructuredRCAFromHTML(html) {
  if (!html) return null;
  try {
    var stripTags = function(s) {
      return (s || '').replace(/<[^>]+>/g, ' ').replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&nbsp;/g,' ').replace(/\s+/g,' ').trim();
    };

    // ── Critical flags (from ai-flag-title divs) ──────────────────────────────
    var flags = [];
    var flagRx = /<div[^>]*class="ai-flag-title"[^>]*>([\s\S]*?)<\/div>/gi;
    var fm;
    while ((fm = flagRx.exec(html)) !== null) {
      var ft = stripTags(fm[1]).replace(/^[^a-zA-Z]+/, '').trim();
      if (ft && ft.length < 200 && ft.toLowerCase() !== 'none') flags.push(ft);
    }

    // ── Repeat projection % ───────────────────────────────────────────────────
    var repeatPct = '0';
    var rpMatch = html.match(/Repeat(?:.*?)(\d{1,3})%/i) ||
                  html.match(/(\d{1,3})%(?:.*?)repeat/i);
    if (rpMatch) repeatPct = rpMatch[1];

    // ── SMART table rows (amber header = FDE8B0 or SMART table) ──────────────
    var smartS='', smartM='', smartA='', smartR='', smartT='';
    var tbodyMatch = html.match(/FDE8B0[\s\S]*?<tbody>([\s\S]*?)<\/tbody>/i);
    if (!tbodyMatch) tbodyMatch = html.match(/SMART[\s\S]*?<tbody>([\s\S]*?)<\/tbody>/i);
    if (tbodyMatch) {
      var rowRx = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
      var rm = rowRx.exec(tbodyMatch[1]);
      if (rm) {
        var cells = [];
        var cellRx = /<td[^>]*>([\s\S]*?)<\/td>/gi;
        var cm;
        while ((cm = cellRx.exec(rm[1])) !== null) cells.push(stripTags(cm[1]));
        // cells[0] = #, [1]=S, [2]=M, [3]=A, [4]=R or T, [5]=T
        if (cells.length >= 3) {
          smartS = cells[1] || '';
          smartM = cells[2] || '';
          smartA = cells[3] || '';
          smartR = cells.length >= 5 ? (cells[4] || '') : '';
          smartT = cells.length >= 6 ? (cells[5] || '') : (cells[4] || '');
        }
      }
    }

    // ── Call summary (from ai-summary div or Summary: paragraph) ─────────────
    var callSummary = '';
    var sumDiv = html.match(/<div[^>]*class="ai-summary"[^>]*>([\s\S]*?)<\/div>/i);
    if (sumDiv) { callSummary = stripTags(sumDiv[1]).substring(0, 250); }
    if (!callSummary) {
      var sumPara = html.match(/Summary[:\s]+([\s\S]*?)(?:<br|<p|<\/p|\n\n)/i);
      if (sumPara) callSummary = stripTags(sumPara[1]).substring(0, 250);
    }

    // ── Top opportunity = first SMART S, or first flag ────────────────────────
    var topOpp = (smartS || (flags[0] || '')).substring(0, 150);

    // ── Call driver from Issue Resolution or first flag ───────────────────────
    var callDriver = '';
    var cdMatch = html.match(/Call (?:Reason|Driver)[^:]*[:\s]+([\s\S]{5,80}?)(?:<\/td>|<br|<\/p)/i);
    if (cdMatch) callDriver = stripTags(cdMatch[1]).substring(0, 100);
    if (!callDriver && flags[0]) callDriver = flags[0].substring(0, 100);

    return {
      callDriver:        callDriver,
      rcaCategory:       '',   // requires AI — populated later by enrichDashboardData
      rcaSubParameter:   flags.slice(0,2).join('; ').substring(0, 100),
      topOpportunity:    topOpp,
      productOpportunity:'',   // requires AI
      callSummaryShort:  callSummary,
      criticalFlags:     flags.slice(0, 5).join(', ').substring(0, 500),
      repeatPct:         repeatPct,
      smartS:            smartS.substring(0, 300),
      smartM:            smartM.substring(0, 300),
      smartA:            smartA.substring(0, 300),
      smartR:            smartR.substring(0, 300),
      smartT:            smartT.substring(0, 200)
    };
  } catch(e) {
    Logger.log('extractStructuredRCAFromHTML error: ' + e);
    return null;
  }
}

// ── Call AI to extract structured RCA fields from cached HTML ─────────────────
function extractStructuredRCA(html, analysisType, auditRef) {
  var isSales = (analysisType || '').indexOf('Sales') !== -1;

  // Strip HTML to plain text first
  var plainText = htmlToPlainText(html).substring(0, 6000);

  var prompt =
    'You are a QA analyst reading a call evaluation report for a TELUS contact center agent.\n\n' +
    'Based on the evaluation text below, extract the following fields.\n' +
    'Return ONLY a JSON object with exactly these keys — no other text:\n\n' +
    '{\n' +
    '  "callDriver": "Short label for the main reason the customer called (5-8 words max). Examples: Service Move Request, Billing Inquiry, Internet Troubleshooting, Loyalty Retention, New Service Inquiry",\n' +
    '  "rcaCategory": "Exactly one of: Agent Controllable | Process/Policy | Customer Driven | Transfer Issue",\n' +
    '  "rcaSubParameter": "The specific missed step or policy gap in 10 words or less. Examples: Missed callback policy, Incomplete authentication, No retention attempt, Invalid transfer",\n' +
    '  "topOpportunity": "The single most impactful coaching opportunity in 12 words or less. Be specific to this call.",\n' +
    (isSales
      ? '  "productOpportunity": "Name the specific product or service the agent could have offered but did not. If a sale was made, write what was sold.",\n' +
        '  "salesAttempted": "Yes or No — did the agent make any sales offer, upsell attempt, or product recommendation during this call?",\n'
      : '  "productOpportunity": "N/A",\n' +
        '  "salesAttempted": "N/A",\n'
    ) +
    '  "callSummaryShort": "One sentence (max 20 words) describing what happened on this call.",\n' +
    '  "criticalFlags": "Comma-separated list of specific critical behaviors flagged in this evaluation. Use the exact flag names from the text (e.g. Missed callback policy, Incomplete authentication, No retention attempt). Max 5 flags. If none, write None.",\n' +
    '  "repeatPct": "Extract ONLY the repeat call risk percentage number from the evaluation (e.g. 85). Just the number, no % sign, no other text. If not found, write 0.",\n' +
    '  "smartS": "The single most important SPECIFIC behavior for this agent to change. One clear sentence.",\n' +
    '  "smartM": "How success will be MEASURED — what metric or observable outcome proves it worked. One sentence.",\n' +
    '  "smartA": "The ATTAINABLE target — what the agent can realistically achieve with this change.",\n' +
    '  "smartR": "Why this goal is REALISTIC for this agent right now — e.g. they already have the knowledge, it is a small habit shift.",\n' +
    '  "smartT": "TIME-BOUND — specific deadline or timeline, e.g. Starting next call, Within 2 weeks, By next QA review."\n' +
    '}\n\n' +
    'RCA Category definitions:\n' +
    '- Agent Controllable: Agent had the knowledge/tools to resolve but did not (wrong steps, missed policy, lack of ownership)\n' +
    '- Process/Policy: Company policy prevented resolution or created the issue (known limitation, required escalation, system issue)\n' +
    '- Customer Driven: Customer behavior/request drove the outcome (refused resolution, requested callback, comparing competitors)\n' +
    '- Transfer Issue: Call was transferred invalidly, unnecessarily, or without proper protocol\n\n' +
    'EVALUATION TEXT:\n\n' + plainText;

  var payload = {
    model:    FUELIX_CONFIG.model,
    messages: [{ role: 'user', content: prompt }],
    stream:   false
  };
  var options = {
    method: 'post', contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + FUELIX_CONFIG.apiKey, 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload), muteHttpExceptions: true, deadline: 120
  };

  var response = UrlFetchApp.fetch(FUELIX_CONFIG.baseUrl + '/v1/chat/completions', options);
  if (response.getResponseCode() !== 200) {
    Logger.log('AI error: ' + response.getResponseCode());
    return null;
  }

  var content = JSON.parse(response.getContentText()).choices[0].message.content.trim();

  // Parse JSON from response
  var jsonMatch = content.match(/\{[\s\S]*\}/);
  if (!jsonMatch) { Logger.log('No JSON in response for ' + auditRef); return null; }

  try {
    var obj = JSON.parse(jsonMatch[0]);
    return {
      callDriver:        (obj.callDriver         || '').substring(0, 100),
      rcaCategory:       (obj.rcaCategory        || '').substring(0, 50),
      rcaSubParameter:   (obj.rcaSubParameter    || '').substring(0, 100),
      topOpportunity:    (obj.topOpportunity      || '').substring(0, 150),
      productOpportunity:(obj.productOpportunity  || '').substring(0, 100),
      salesAttempted:    (obj.salesAttempted      || '').substring(0, 10),
      callSummaryShort:  (obj.callSummaryShort    || '').substring(0, 200),
      criticalFlags:     (obj.criticalFlags       || 'None').replace(/^None$/i, '').substring(0, 500),
      repeatPct:         (obj.repeatPct           || '0').toString().replace(/[^0-9]/g,'').substring(0, 5),
      smartS:            (obj.smartS              || '').substring(0, 300),
      smartM:            (obj.smartM              || '').substring(0, 300),
      smartA:            (obj.smartA              || '').substring(0, 300),
      smartR:            (obj.smartR              || '').substring(0, 300),
      smartT:            (obj.smartT              || '').substring(0, 200)
    };
  } catch(e) {
    Logger.log('JSON parse error for ' + auditRef + ': ' + e);
    return null;
  }
}

// ── Deferred RCA enrichment — called by client after submitTranscript returns ─
// Runs the second AI call in the background and patches the 3 AI-only cells
// in the Dashboard_Data row for this auditRef. Errors are non-fatal.
function enrichDashboardRCA(auditRef, analysisType, html) {
  try {
    if (!auditRef || !html) return;

    var aiRca = extractStructuredRCA(html, analysisType, auditRef);
    if (!aiRca) {
      Logger.log('enrichDashboardRCA: AI returned null for ' + auditRef);
      return;
    }

    var ss    = getOrCreateSpreadsheet();
    var sheet = getOrCreateSheet(ss, DASHBOARD_DATA_SHEET);
    if (sheet.getLastRow() < 2) return;

    // Find the row by auditRef using TextFinder (same pattern as findCachedResult)
    var finder  = sheet.getRange('A:A').createTextFinder(auditRef.trim()).matchEntireCell(true);
    var matches = finder.findAll();
    if (!matches || !matches.length) {
      Logger.log('enrichDashboardRCA: auditRef not found: ' + auditRef);
      return;
    }
    var targetRow = matches[matches.length - 1].getRow();

    // Resolve column numbers by header name — safe against column-order changes
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var colOf   = function(name) {
      var idx = headers.indexOf(name);
      return idx >= 0 ? idx + 1 : -1;
    };

    var rcaCatCol  = colOf('RCA Category');
    var prodOppCol = colOf('Product Opportunity');
    var salesAttCol= colOf('Sales Attempted');

    if (rcaCatCol   > 0 && aiRca.rcaCategory)        sheet.getRange(targetRow, rcaCatCol).setValue(aiRca.rcaCategory);
    if (prodOppCol  > 0 && aiRca.productOpportunity) sheet.getRange(targetRow, prodOppCol).setValue(aiRca.productOpportunity);
    if (salesAttCol > 0 && aiRca.salesAttempted)     sheet.getRange(targetRow, salesAttCol).setValue(aiRca.salesAttempted);

    invalidateDashboardCache();
    Logger.log('enrichDashboardRCA: patched row ' + targetRow + ' for ' + auditRef);
  } catch(e) {
    Logger.log('enrichDashboardRCA error (non-fatal): ' + e);
  }
}

// ── Admin: run once from GAS editor to fix manually-typed locale values ──
// Finds Dashboard_Data rows where Locale exactly matches a known bad value
// and replaces it with the correct full locale name.
function fixTruncatedLocale() {
  var FIXES = [
    { bad: 'TI M',     good: 'TI Morocco (CAS)' },
    { bad: 'TI Mo',    good: 'TI Morocco (CAS)' },
    { bad: 'TI Mor',   good: 'TI Morocco (CAS)' }
  ];

  var ss      = getOrCreateSpreadsheet();
  var sheet   = getOrCreateSheet(ss, DASHBOARD_DATA_SHEET);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('fixTruncatedLocale: no data rows'); return; }

  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var localeCol = headers.indexOf('Locale');
  if (localeCol < 0) { Logger.log('fixTruncatedLocale: Locale column not found'); return; }
  localeCol++; // convert to 1-based

  var data    = sheet.getRange(2, localeCol, lastRow - 1, 1).getValues();
  var patched = 0;

  for (var i = 0; i < data.length; i++) {
    var cell = (data[i][0] || '').toString().trim();
    for (var j = 0; j < FIXES.length; j++) {
      if (cell === FIXES[j].bad) {
        sheet.getRange(i + 2, localeCol).setValue(FIXES[j].good);
        Logger.log('fixTruncatedLocale: row ' + (i + 2) + ' "' + cell + '" → "' + FIXES[j].good + '"');
        patched++;
        break;
      }
    }
  }

  if (patched > 0) invalidateDashboardCache();
  Logger.log('=== fixTruncatedLocale done: ' + patched + ' patched ===');
}

// ── Run ONCE from the Apps Script editor to upgrade the Dashboard_Data sheet ──
// Renames 'Recommendations'→'SMART Recommendation', 'Top Opportunity'→'Top SMART Opportunity'
// and appends 5 new SMART columns (S/M/A/R/T). Safe to run on a live sheet —
// it only adds/renames header cells and does not touch data rows.
function updateSpreadsheetSMARTFormat() {
  var ss     = getOrCreateSpreadsheet();
  var dSheet = getOrCreateSheet(ss, DASHBOARD_DATA_SHEET);
  if (dSheet.getLastRow() < 1) { Logger.log('Sheet is empty — nothing to update'); return; }

  var headerRange = dSheet.getRange(1, 1, 1, dSheet.getLastColumn());
  var headers     = headerRange.getValues()[0];

  // ── 1. Rename existing headers ─────────────────────────────────────────────
  var renames = {
    'Recommendations': 'SMART Recommendation',
    'Top Opportunity': 'Top SMART Opportunity'
  };
  headers.forEach(function(h, i) {
    if (renames[h]) {
      dSheet.getRange(1, i + 1).setValue(renames[h]);
      Logger.log('Renamed col ' + (i+1) + ': "' + h + '" → "' + renames[h] + '"');
    }
  });

  // ── 2. Add missing SMART columns (insert before PDF Email Link) ────────────
  var newCols = [
    'SMART S (Specific)',
    'SMART M (Measurable)',
    'SMART A (Attainable)',
    'SMART R (Realistic)',
    'SMART T (Time-bound)'
  ];

  // Find insertion point: before "PDF Email Link" or at end
  headers = dSheet.getRange(1, 1, 1, dSheet.getLastColumn()).getValues()[0];
  var insertBefore = headers.indexOf('PDF Email Link');
  if (insertBefore === -1) insertBefore = headers.length; // append at end

  // Only add columns that don't already exist
  var toAdd = newCols.filter(function(c){ return headers.indexOf(c) === -1; });

  if (toAdd.length) {
    // Insert columns at insertBefore position (1-based)
    dSheet.insertColumnsBefore(insertBefore + 1, toAdd.length);
    var newRange = dSheet.getRange(1, insertBefore + 1, 1, toAdd.length);
    newRange.setValues([toAdd])
      .setFontWeight('bold')
      .setBackground('#4B286D')
      .setFontColor('#ffffff');
    Logger.log('Added ' + toAdd.length + ' SMART columns at position ' + (insertBefore + 1));
  } else {
    Logger.log('All SMART columns already exist — no columns added');
  }

  Logger.log('updateSpreadsheetSMARTFormat complete.');
}

// ── Run ONCE: insert 'Sales Attempted' column after 'Product Opportunity' ────────
function addSalesAttemptedColumn() {
  var ss     = getOrCreateSpreadsheet();
  var dSheet = getOrCreateSheet(ss, DASHBOARD_DATA_SHEET);
  if (dSheet.getLastRow() < 1) { Logger.log('Sheet is empty — nothing to update'); return; }

  var headers = dSheet.getRange(1, 1, 1, dSheet.getLastColumn()).getValues()[0];

  // Already exists — skip
  if (headers.indexOf('Sales Attempted') !== -1) {
    Logger.log('Sales Attempted column already exists — nothing to do');
    return;
  }

  // Find insertion point: right after 'Product Opportunity'
  var afterIdx = headers.indexOf('Product Opportunity');
  if (afterIdx === -1) {
    Logger.log('Product Opportunity column not found — appending Sales Attempted at end');
    afterIdx = headers.length - 1;
  }

  var insertCol = afterIdx + 2; // 1-based, insert AFTER Product Opportunity
  dSheet.insertColumnBefore(insertCol);
  var cell = dSheet.getRange(1, insertCol);
  cell.setValue('Sales Attempted')
    .setFontWeight('bold')
    .setBackground('#4B286D')
    .setFontColor('#ffffff');

  Logger.log('Sales Attempted column inserted at column ' + insertCol);
  Logger.log('Run enrichDashboardData() next to backfill values.');
}

// ── Direct read of Sales Attempted counts — bypasses all caching ─────────────
function getSalesAttemptedCounts() {
  try {
    var ss     = getOrCreateSpreadsheet();
    var dSheet = getOrCreateSheet(ss, DASHBOARD_DATA_SHEET);
    var lastRow = dSheet.getLastRow();
    var lastCol = dSheet.getLastColumn();
    if (lastRow < 2) return { yes: 0, no: 0 };
    var headers = dSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var saCol   = -1, atCol = -1;
    headers.forEach(function(h, i) {
      if (h === 'Sales Attempted') saCol = i + 1;
      if (h === 'Analysis Type')  atCol = i + 1;
    });
    if (saCol < 1 || atCol < 1) return { yes: 0, no: 0 };
    var data = dSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var yes = 0, no = 0;
    data.forEach(function(row) {
      var atype = (row[atCol - 1] || '').toString().toLowerCase();
      if (atype.indexOf('sales') === -1) return;
      var v = (row[saCol - 1] || '').toString().trim().toLowerCase();
      if (v === 'yes') yes++;
      else if (v === 'no') no++;
    });
    return { yes: yes, no: no };
  } catch(e) {
    Logger.log('getSalesAttemptedCounts error: ' + e);
    return { yes: 0, no: 0 };
  }
}

// ── Run once from GAS editor to verify Sales Attempted column mapping ─────────
function testSalesAttempted() {
  var ss     = getOrCreateSpreadsheet();
  var dSheet = getOrCreateSheet(ss, DASHBOARD_DATA_SHEET);
  var lastCol = dSheet.getLastColumn();
  var headers = dSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var saCol = -1;
  headers.forEach(function(h, i) { if (h === 'Sales Attempted') saCol = i + 1; });
  Logger.log('Total columns: ' + lastCol);
  Logger.log('Sales Attempted column (1-based): ' + saCol);
  Logger.log('Header at index 28 (0-based): "' + headers[28] + '"');
  Logger.log('Header at index 27 (0-based): "' + headers[27] + '"');
  Logger.log('Header at index 29 (0-based): "' + headers[29] + '"');
  if (saCol > 0) {
    var data = dSheet.getRange(2, saCol, Math.min(10, dSheet.getLastRow()-1), 1).getValues();
    Logger.log('First 10 Sales Attempted values: ' + JSON.stringify(data.map(function(r){return r[0];})));
  } else {
    Logger.log('ERROR: Sales Attempted column NOT FOUND in sheet header row!');
  }
}

// ── Run once: add header notes/comments to Dashboard_Data columns ─────────────
// ── One-time fix: extract Repeat % from Issue Resolved text and Audit_Log ──────
// ── Backfill VTID into Audit_Log from Dashboard_Data ─────────────────────────
// ── Backfill VTID in Dashboard_Data from AT Data GCP using SAP ID ────────────
function backfillVTIDFromATData() {
  var ss     = getOrCreateSpreadsheet();
  var dSheet = getOrCreateSheet(ss, DASHBOARD_DATA_SHEET);

  var dHeaders = dSheet.getRange(1,1,1,dSheet.getLastColumn()).getValues()[0];
  var colOf = {};
  dHeaders.forEach(function(h,i){ if(h) colOf[h.toString().trim()] = i+1; });

  var sapCol  = colOf['SAP ID'];
  var vtidCol = colOf['VTID'];
  if (!sapCol || !vtidCol) { Logger.log('SAP ID or VTID column not found'); return; }

  var dLast = dSheet.getLastRow();
  if (dLast < 2) return;
  var dData = dSheet.getRange(2, 1, dLast-1, dHeaders.length).getValues();

  // Get unique SAP IDs that have no VTID
  var sapIds = dbUniqServer(dData
    .filter(function(r){ return r[sapCol-1] && !r[vtidCol-1]; })
    .map(function(r){ return r[sapCol-1].toString().trim(); }));

  Logger.log('Looking up VTID for ' + sapIds.length + ' unique SAP IDs');

  // Fetch AT Data GCP sheet
  var atSs    = SpreadsheetApp.openById(AT_DATA_GCP_SS_ID);
  var atSheet = atSs.getSheetByName('Roster') || atSs.getSheetByName('roster') || atSs.getSheets()[0];
  var atData  = atSheet.getDataRange().getValues();
  var atHeaders = atData[0];
  var prodCol = 6, vtidAtCol = 5;
  atHeaders.forEach(function(h,i){
    var hl = h.toString().toLowerCase().trim();
    if (hl === 'production id' || hl === 'production_id') prodCol  = i;
    if (hl === 'reference no.' || hl === 'vtid')           vtidAtCol = i;
  });

  // Build SAP → VTID map from AT Data
  var sapToVtid = {};
  atData.slice(1).forEach(function(row) {
    var sap  = (row[prodCol]  || '').toString().trim();
    var vtid = (row[vtidAtCol]|| '').toString().trim();
    if (sap && vtid) sapToVtid[sap] = vtid;
  });
  Logger.log('AT Data map: ' + Object.keys(sapToVtid).length + ' entries');

  // Write VTIDs back to Dashboard_Data
  var updated = 0;
  dData.forEach(function(row, i) {
    var sap      = (row[sapCol-1]  || '').toString().trim();
    var existing = (row[vtidCol-1] || '').toString().trim();
    if (!existing && sapToVtid[sap]) {
      dSheet.getRange(i+2, vtidCol).setValue(sapToVtid[sap]);
      Logger.log('Row ' + (i+2) + ' SAP ' + sap + ' → VTID ' + sapToVtid[sap]);
      updated++;
    }
  });
  invalidateDashboardCache();
  Logger.log('=== Updated VTID for ' + updated + ' Dashboard_Data rows ===');
}

function dbUniqServer(arr) {
  var seen = {};
  return arr.filter(function(v){ return v && !seen[v] && (seen[v]=1); });
}

function backfillVTIDToAuditLog() {
  var ss       = getOrCreateSpreadsheet();
  var logSheet = getOrCreateSheet(ss, AUDIT_LOG_SHEET);
  var dSheet   = getOrCreateSheet(ss, DASHBOARD_DATA_SHEET);

  var logHeaders = logSheet.getRange(1,1,1,logSheet.getLastColumn()).getValues()[0];
  var dHeaders   = dSheet.getRange(1,1,1,dSheet.getLastColumn()).getValues()[0];
  var logColOf = {}, dColOf = {};
  logHeaders.forEach(function(h,i){ if(h) logColOf[h.toString().trim()] = i+1; });
  dHeaders.forEach(function(h,i){ if(h) dColOf[h.toString().trim()] = i+1; });

  var logVtidCol = logColOf['VTID'];
  var logRefCol  = logColOf['Audit Ref'];
  var dRefCol    = dColOf['Audit Ref'];
  var dVtidCol   = dColOf['VTID'];

  if (!logVtidCol || !dVtidCol) { Logger.log('VTID column missing'); return; }

  // Build Audit Ref → VTID from Dashboard_Data
  var dLast = dSheet.getLastRow();
  var vtidMap = {};
  if (dLast >= 2) {
    var dData = dSheet.getRange(2, 1, dLast-1, dHeaders.length).getValues();
    dData.forEach(function(row) {
      var ref  = (row[dRefCol-1]||'').toString().trim();
      var vtid = (row[dVtidCol-1]||'').toString().trim();
      if (ref && vtid) vtidMap[ref] = vtid;
    });
  }

  var logLast = logSheet.getLastRow();
  if (logLast < 2) return;
  var logData = logSheet.getRange(2,1,logLast-1,logHeaders.length).getValues();
  var updated = 0;
  logData.forEach(function(row, i) {
    var ref      = (row[logRefCol-1]||'').toString().trim();
    var existing = (row[logVtidCol-1]||'').toString().trim();
    if (!existing && vtidMap[ref]) {
      logSheet.getRange(i+2, logVtidCol).setValue(vtidMap[ref]);
      Logger.log('Audit_Log row ' + (i+2) + ': VTID = ' + vtidMap[ref]);
      updated++;
    }
  });
  invalidateDashboardCache();
  Logger.log('=== Backfilled VTID for ' + updated + ' Audit_Log rows ===');
}

function fixRepeatProjectionPct() {
  var ss      = getOrCreateSpreadsheet();
  var dSheet  = getOrCreateSheet(ss, DASHBOARD_DATA_SHEET);
  var logSheet= getOrCreateSheet(ss, AUDIT_LOG_SHEET);

  var dHeaders = dSheet.getRange(1, 1, 1, dSheet.getLastColumn()).getValues()[0];
  var colOf = {};
  dHeaders.forEach(function(h, i) { if (h) colOf[h.toString().trim()] = i + 1; });

  var repeatCol  = colOf['Repeat Projection %'];
  var issueCol   = colOf['Issue Resolved'];
  var refCol     = colOf['Audit Ref'];
  if (!repeatCol) { Logger.log('Repeat Projection % column not found'); return; }

  // Build Audit Ref → Repeat % from Audit_Log (most reliable source)
  var logLast = logSheet.getLastRow();
  var logRefToRepeat = {};
  if (logLast >= 2) {
    var logHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
    var logColOf = {};
    logHeaders.forEach(function(h, i){ if (h) logColOf[h.toString().trim()] = i + 1; });
    var logRepeatCol = logColOf['Repeat Projection %'];
    var logRefCol    = logColOf['Audit Ref'];
    if (logRepeatCol && logRefCol) {
      var logData = logSheet.getRange(2, 1, logLast - 1, logSheet.getLastColumn()).getValues();
      logData.forEach(function(row) {
        var ref = (row[logRefCol - 1] || '').toString().trim();
        var pct = (row[logRepeatCol - 1] || '').toString().trim();
        if (ref && pct) logRefToRepeat[ref] = pct;
      });
      Logger.log('Audit_Log repeat values: ' + Object.keys(logRefToRepeat).length + ' entries');
    }
  }

  var dLast = dSheet.getLastRow();
  if (dLast < 2) { Logger.log('No data rows'); return; }
  var dData = dSheet.getRange(2, 1, dLast - 1, dHeaders.length).getValues();
  var fixed = 0;

  dData.forEach(function(row, i) {
    var existing = (row[repeatCol - 1] || '').toString().trim();
    // Skip only if already has a clean integer between 1-100
    var existingInt = parseInt(existing, 10);
    if (!isNaN(existingInt) && existingInt > 0 && existingInt <= 100 && existing === String(existingInt)) return;

    var ref      = refCol ? (row[refCol - 1] || '').toString().trim() : '';
    var foundPct = '';

    // Source 1: Audit_Log direct value
    if (ref && logRefToRepeat[ref]) {
      foundPct = logRefToRepeat[ref].toString().replace(/[^0-9.]/g, '');
    }

    // Source 2: Parse from Issue Resolved column text ("85% repeat risk...")
    if (!foundPct && issueCol) {
      var issueText = (row[issueCol - 1] || '').toString();
      var m = issueText.match(/(\d{1,3})\s*%\s*repeat/i)
           || issueText.match(/repeat[^\d]*(\d{1,3})\s*%/i);
      if (m) foundPct = m[1];
    }

    // Source 3: Parse from any column that contains "% repeat risk"
    if (!foundPct) {
      for (var c = 0; c < row.length; c++) {
        var cell = (row[c] || '').toString();
        var pm   = cell.match(/(\d{1,3})\s*%\s*repeat\s*risk/i)
                || cell.match(/repeat\s*risk[^0-9]*(\d{1,3})\s*%/i);
        if (pm) { foundPct = pm[1]; break; }
      }
    }

    // Extract only the leading integer (e.g. "55.164" → 55, "92...." → 92)
    var cleanPct = parseInt(foundPct, 10);
    if (!isNaN(cleanPct) && cleanPct > 0 && cleanPct <= 100) {
      dSheet.getRange(i + 2, repeatCol).setValue(cleanPct);
      Logger.log('Row ' + (i+2) + ' (' + ref + '): Repeat % = ' + cleanPct);
      fixed++;
    }
  });

  invalidateDashboardCache();
  Logger.log('=== Fixed Repeat Projection % for ' + fixed + ' rows ===');
}

function addDashboardColumnNotes() {
  var ss     = getOrCreateSpreadsheet();
  var sheet  = getOrCreateSheet(ss, DASHBOARD_DATA_SHEET);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var notes = {
    'Critical Flags':       'Lists specific critical behaviors or policy violations observed during the interaction. Each flag represents a coaching priority that requires immediate attention. Agents with repeated flags may need a Performance Improvement Plan.',
    'Repeat Projection %':  'AI-generated repeat call risk score (0–100%). Indicates the likelihood the customer will call back without the issue being fully resolved.\n\n0–25% = Low risk (issue likely resolved)\n26–50% = Moderate risk (partial resolution)\n51–75% = High risk (unresolved or incomplete)\n76–100% = Critical risk (guaranteed repeat call)\n\nTarget: Keep below 30% on average.',
    'Call Driver':           'Short label for the main reason the customer called. Used for trend analysis and call volume charts. Examples: Service Move Request, Billing Dispute, Seasonal Hold.',
    'RCA Category':          'Root Cause Analysis category:\n• Agent Controllable — agent had tools/knowledge but did not resolve\n• Process/Policy — company policy prevented resolution\n• Customer Driven — customer behavior drove the outcome\n• Transfer Issue — invalid or unnecessary transfer occurred',
    'RCA Sub-Parameter':     'Specific missed step or policy gap that caused the repeat risk. Used for Top 5 Opportunities chart. Examples: Missed callback policy, Incomplete authentication, No retention attempt.',
    'Top SMART Opportunity':  'The single most impactful SMART coaching action for the agent on this specific call — written in Specific/Measurable/Attainable/Realistic/Time-bound format. Should be referenced during coaching sessions.',
    'SMART S (Specific)':    'The exact specific behavior the agent needs to change or develop.',
    'SMART M (Measurable)':  'How success will be measured — the observable metric that proves the coaching worked.',
    'SMART A (Attainable)':  'The attainable target — what the agent can realistically achieve with this change.',
    'SMART R (Realistic)':   'Why this goal is realistic for this agent right now — confirms the goal is achievable given their current skill level.',
    'SMART T (Time-bound)':  'The deadline or timeline for achieving the SMART coaching goal.',
    'Product Opportunity':   'For Sales Analyzer: the specific product or service the agent could have offered. For Repeats: N/A.'
  };

  var updated = 0;
  headers.forEach(function(h, i) {
    var note = notes[h.toString().trim()];
    if (note) {
      sheet.getRange(1, i + 1).setNote(note);
      Logger.log('Added note to column ' + String.fromCharCode(65 + i) + ' (' + h + ')');
      updated++;
    }
  });
  Logger.log('=== Added notes to ' + updated + ' columns ===');
}

// ── Run from editor — shows recent errors from Stackdriver logs ───────────────
function showRecentErrors() {
  try {
    // Test getRosterBySapId directly to reproduce the error
    Logger.log('=== Testing getRosterBySapId ===');
    var testSap = '2007888';
    var start = new Date();
    var result = lookupBySapId(testSap);
    var elapsed = new Date() - start;
    Logger.log('Result: ' + JSON.stringify(result));
    Logger.log('Time taken: ' + elapsed + 'ms');

    // Also test roster data load
    Logger.log('\n=== Testing roster cache ===');
    var start2 = new Date();
    var rosterData = _getRosterSheetData();
    var elapsed2 = new Date() - start2;
    Logger.log('Roster rows loaded: ' + rosterData.length);
    Logger.log('Roster load time: ' + elapsed2 + 'ms');

    // Check cache state
    var cache = CacheService.getScriptCache();
    var hasDash = cache.get('dashboard_data_meta_v1') ? 'YES' : 'NO';
    var hasRoster = cache.get('roster_sheet_data_v2') ? 'YES' : 'NO';
    Logger.log('\n=== Cache state ===');
    Logger.log('Dashboard cache: ' + hasDash);
    Logger.log('Roster cache: ' + hasRoster);
  } catch(e) {
    Logger.log('ERROR: ' + e.toString());
    Logger.log('Stack: ' + e.stack);
  }
}

function diagnoseFCRDashboard() {
  var SAP_ID = '2007888';
  var ss     = SpreadsheetApp.openById(FCR_DASHBOARD_SS_ID);
  var sheets = ss.getSheets();

  Logger.log('=== FCR Dashboard — all tabs ===');
  sheets.forEach(function(sh) {
    var rows = sh.getLastRow();
    var cols = sh.getLastColumn();
    Logger.log('Tab: "' + sh.getName() + '" rows=' + rows + ' cols=' + cols);
    if (rows > 0 && cols > 0) {
      try {
        var hdr = sh.getRange(1, 1, 1, cols).getValues()[0];
        hdr.forEach(function(h, i) {
          if (h) Logger.log('  Col ' + String.fromCharCode(65+i) + ' [idx ' + i + ']: ' + h);
        });
        // Show first data row to spot check values
        if (rows > 1) {
          var row1 = sh.getRange(2, 1, 1, Math.min(cols, 15)).getValues()[0];
          Logger.log('  Sample row 2: ' + row1.join(' | '));
        }
      } catch(e) { Logger.log('  Error reading: ' + e); }
    }
  });

  Logger.log('\n=== FCR lookup for SAP: ' + SAP_ID + ' ===');
  var result = lookupFromFCRDashboard(SAP_ID, Number(SAP_ID));
  Logger.log('Result: ' + JSON.stringify(result));
}

function diagnoseATDataRow() {
  var SAP_ID = '2007888';
  var ss     = SpreadsheetApp.openById(AT_DATA_GCP_SS_ID);
  var sheet  = ss.getSheetByName('Roster') || ss.getSheetByName('roster') || ss.getSheets()[0];
  var data   = sheet.getDataRange().getValues();
  var headers = data[0];

  Logger.log('=== AT DATA GCP HEADERS ===');
  headers.forEach(function(h, i) {
    Logger.log('Col ' + String.fromCharCode(65 + i) + ' [idx ' + i + ']: ' + h);
  });

  Logger.log('\n=== ROW FOR SAP ID: ' + SAP_ID + ' (row 4364) ===');
  var target = Number(SAP_ID);
  for (var i = 1; i < data.length; i++) {
    var cell = data[i][6]; // Col G = Production ID
    if (Number(cell) === target) {
      Logger.log('FOUND at row ' + (i+1));
      data[i].forEach(function(v, j) {
        if (v || v === 0) Logger.log('  Col ' + String.fromCharCode(65+j) + ' [idx ' + j + '] "' + headers[j] + '": ' + v);
      });
      return;
    }
  }
  Logger.log('Not found');
}

function testLookupDirect() {
  var SAP_ID     = '2007888';
  var AGENT_NAME = 'James Collado';

  Logger.log('=== DIRECT LOOKUP TEST ===');

  // 1. Search roster by AGENT NAME to find actual stored SAP ID
  var ss    = SpreadsheetApp.openById(ROSTER_SHEET_ID);
  var sheet = ss.getSheetByName('roster') || ss.getSheets()[0];
  var data  = sheet.getDataRange().getValues();
  var nameLower = AGENT_NAME.toLowerCase().trim();

  Logger.log('--- Searching roster for agent: "' + AGENT_NAME + '" ---');
  var found = false;
  for (var i = 1; i < data.length; i++) {
    var cellName = (data[i][ROSTER_COL_AGENT_NAME] || '').toString().toLowerCase().trim();
    if (cellName === nameLower) {
      Logger.log('FOUND at row ' + (i+1));
      Logger.log('  Col A (sap_ID):               [' + data[i][0]  + '] type=' + typeof data[i][0]);
      Logger.log('  Col C (Agent_Name):            [' + data[i][2]  + ']');
      Logger.log('  Col F (Agent Role/LOB):        [' + data[i][5]  + ']');
      Logger.log('  Col I (Team_Mgr_Name):         [' + data[i][8]  + ']');
      Logger.log('  Col K (Ops_Mgr_Name):          [' + data[i][10] + ']');
      Logger.log('  Col S (Location/Locale idx18): [' + data[i][18] + ']');
      found = true;
      break;
    }
  }
  if (!found) Logger.log('Agent "' + AGENT_NAME + '" NOT FOUND in roster');

  // 2. Test VTID lookup with the typed SAP ID
  Logger.log('--- VTID lookup for SAP ID: ' + SAP_ID + ' ---');
  var vtid = lookupVTID(SAP_ID);
  Logger.log('VTID result: [' + vtid + ']');

  // 3. Full lookup
  Logger.log('--- Full lookupBySapId for: ' + SAP_ID + ' ---');
  var result = lookupBySapId(SAP_ID);
  Logger.log('Result: ' + JSON.stringify(result));

  if (!result) {
    Logger.log('RESULT IS NULL — checking raw sheet data...');
    var ss    = SpreadsheetApp.openById(ROSTER_SHEET_ID);
    var sheet = ss.getSheetByName('roster') || ss.getSheets()[0];
    Logger.log('Sheet name: ' + sheet.getName());
    var data = sheet.getDataRange().getValues();
    Logger.log('Total rows: ' + data.length);
    Logger.log('Headers: ' + data[0].slice(0, 12).join(' | '));
    Logger.log('Searching for SAP ID [' + SAP_ID + '] in col A...');
    for (var i = 1; i < data.length; i++) {
      var cell = data[i][0];
      if (cell !== null && cell !== undefined && cell !== '') {
        var cellStr = cell.toString().trim();
        if (cellStr === SAP_ID || cellStr === SAP_ID.toString()) {
          Logger.log('FOUND at row ' + (i+1) + ':');
          data[i].forEach(function(v, j) {
            Logger.log('  [' + j + '] ' + data[0][j] + ' = [' + v + ']');
          });
          return;
        }
      }
    }
    // Show first 10 SAP IDs to compare format
    Logger.log('NOT FOUND. First 10 SAP values in Col A:');
    for (var k = 1; k <= Math.min(10, data.length-1); k++) {
      Logger.log('  Row ' + (k+1) + ': type=' + typeof data[k][0] + ' value=[' + data[k][0] + ']');
    }
  }
}

function diagnoseSapIdRow() {
  var SAP_ID = '2007888';   // <-- change this to test any SAP ID

  var ss    = SpreadsheetApp.openById(ROSTER_SHEET_ID);
  var sheet = ss.getSheetByName('roster');
  var data  = sheet.getDataRange().getValues();
  var headers = data[0];

  Logger.log('=== ROSTER HEADERS ===');
  headers.forEach(function(h, i) {
    Logger.log('Col ' + String.fromCharCode(65 + i) + ' [idx ' + i + ']: ' + h);
  });

  Logger.log('\n=== ROW FOR SAP ID: ' + SAP_ID + ' ===');
  var targetNum = parseFloat(SAP_ID);
  for (var i = 1; i < data.length; i++) {
    var cell = data[i][0];
    if (!cell && cell !== 0) continue;
    var cellStr = cell.toString().trim();
    var cellNum = parseFloat(cellStr);
    if (cellStr === SAP_ID || (!isNaN(targetNum) && !isNaN(cellNum) && cellNum === targetNum)) {
      data[i].forEach(function(v, j) {
        if (v || v === 0) Logger.log('Col ' + String.fromCharCode(65 + j) + ' [idx ' + j + '] "' + headers[j] + '": ' + v);
      });
      return;
    }
  }
  Logger.log('SAP ID not found: ' + SAP_ID);
}

// ── Run this to see AT Data GCP columns ──────────────────────────────────────
function diagnoseATDataForAgent() {
  var AGENT_NAME = 'James Collado';  // <-- change to test any agent

  var ss    = SpreadsheetApp.openById(AT_DATA_GCP_SS_ID);
  var sheet = ss.getSheets()[0];
  var data  = sheet.getDataRange().getValues();
  var headers = data[0];

  Logger.log('=== AT DATA GCP HEADERS ===');
  headers.forEach(function(h, i) {
    Logger.log('Col ' + String.fromCharCode(65 + i) + ' [idx ' + i + ']: ' + h);
  });

  Logger.log('\n=== ROW FOR AGENT: ' + AGENT_NAME + ' ===');
  var nameLower = AGENT_NAME.toLowerCase().trim();
  for (var i = 1; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] && data[i][j].toString().toLowerCase().trim() === nameLower) {
        data[i].forEach(function(v, k) {
          if (v) Logger.log('Col ' + String.fromCharCode(65 + k) + ' [idx ' + k + '] "' + headers[k] + '": ' + v);
        });
        return;
      }
    }
  }
  Logger.log('Agent not found: ' + AGENT_NAME);
}

// ── Run to clean ALL HTML from every cell in the Analysis sheet ───────────────
// Scans every column in every row — catches any analysis type, any column order
function cleanAnalysisSheet() {
  var ss      = getOrCreateSpreadsheet();
  var sheet   = getOrCreateSheet(ss, ANALYSIS_SHEET);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) { Logger.log('No data rows to clean'); return; }

  var data    = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var fixed   = 0;

  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      var val = data[i][j];
      if (!val) continue;
      var str = val.toString();
      // Clean if it contains HTML tags or CSS code
      if (str.indexOf('<') !== -1 || str.indexOf('{') !== -1 || str.indexOf('```') !== -1) {
        data[i][j] = htmlToPlainText(str);
        fixed++;
        Logger.log('Cleaned row ' + (i+2) + ' col ' + (j+1));
      }
    }
  }

  if (fixed > 0) {
    sheet.getRange(2, 1, data.length, lastCol).setValues(data);
    Logger.log('Done — cleaned ' + fixed + ' cell(s) in Analysis sheet');
  } else {
    Logger.log('No HTML/code found — all cells are already clean');
  }
}

// setupScriptProperties was removed — never store credentials in source code.
// Set ANTHROPIC_AUTH_TOKEN, ANTHROPIC_BASE_URL, ANTHROPIC_MODEL manually in
// Apps Script editor → Project Settings → Script Properties.

// ─────────────────────────────────────────────────────────────────────────────
// ADMIN NOTIFICATION SYSTEM
// Reads the Roster sheet for Admin/Dev recipients and sends them an email
// summary whenever a new evaluation is submitted.
//
// Roster sheet structure (in the main NH FCR spreadsheet):
//   Column A: Email Address
//   Column B: Name
//   Column C: Role  ← filter for exactly "Admin/Dev"
//
// Usage: call notifyAdmins(formData, auditRef) after each submitTranscript
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Reads the Audit Tracking SS → Roster tab and returns all rows matching
 * the given role (case-insensitive). Columns detected by header name:
 * Name | Role | Email Address | Supervisor | Supervisor Email
 * Results are cached for 2 hours.
 */
var _rosterRecipientsInMemory = null;

function getRecipientsFromRoster(roleFilter) {
  try {
    var allRows = null;

    if (_rosterRecipientsInMemory) {
      allRows = _rosterRecipientsInMemory;
    } else {
      var cache    = CacheService.getScriptCache();
      var cacheKey = 'audit_roster_rows_v1';
      var cached   = cache.get(cacheKey);
      if (cached) {
        try { allRows = JSON.parse(cached); } catch(e) {}
      }
      if (!allRows) {
        var ss    = SpreadsheetApp.openById(AUDIT_TRACKING_SS_ID);
        var sheet = ss.getSheetByName('Roster') || ss.getSheetByName('roster');
        if (!sheet) { Logger.log('getRecipientsFromRoster: Roster sheet not found'); return []; }
        var data = sheet.getDataRange().getValues();
        if (data.length < 2) return [];
        var hdrs = data[0];
        var nameCol = -1, roleCol = -1, emailCol = -1, supCol = -1, supEmailCol = -1;
        hdrs.forEach(function(h, i) {
          var hl = h.toString().toLowerCase().trim();
          if (hl === 'name')                                           nameCol     = i;
          if (hl === 'role')                                           roleCol     = i;
          if (hl === 'email address' || hl === 'email')                emailCol    = i;
          if (hl === 'supervisor')                                     supCol      = i;
          if (hl === 'supervisor email' || hl === 'supervisoremail')   supEmailCol = i;
        });
        allRows = data.slice(1).map(function(row) {
          return {
            name:            nameCol     > -1 ? (row[nameCol]     || '').toString().trim() : '',
            role:            roleCol     > -1 ? (row[roleCol]     || '').toString().trim() : '',
            email:           emailCol    > -1 ? (row[emailCol]    || '').toString().trim() : '',
            supervisor:      supCol      > -1 ? (row[supCol]      || '').toString().trim() : '',
            supervisorEmail: supEmailCol > -1 ? (row[supEmailCol] || '').toString().trim() : ''
          };
        });
        try { cache.put(cacheKey, JSON.stringify(allRows).substring(0, 95000), 2 * 60 * 60); } catch(e) {}
        _rosterRecipientsInMemory = allRows;
      }
    }

    var roleLower = roleFilter.toLowerCase().trim();
    var validRe   = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    var results   = allRows.filter(function(r) {
      return r.role.toLowerCase() === roleLower && validRe.test(r.email);
    });
    Logger.log('getRecipientsFromRoster("' + roleFilter + '"): ' + results.length + ' found');
    return results;
  } catch(e) {
    Logger.log('getRecipientsFromRoster error: ' + e);
    return [];
  }
}

/**
 * Returns Admin/Dev recipients from the Audit Tracking SS Roster tab.
 */
function getAdminRecipients() {
  return getRecipientsFromRoster('Admin/Dev');
}

// ── Agent email lookup from main roster file ──────────────────────────────────
// Builds a name → email map from the "Team Member Email" column once per execution,
// cached in CacheService for 2 hours to avoid repeated sheet reads.
var _agentEmailMapInMemory = null;

function _getAgentEmailMap() {
  if (_agentEmailMapInMemory) return _agentEmailMapInMemory;
  try {
    var cache    = CacheService.getScriptCache();
    var cacheKey = 'agent_email_map_v1';
    var cached   = cache.get(cacheKey);
    if (cached) {
      try { _agentEmailMapInMemory = JSON.parse(cached); return _agentEmailMapInMemory; } catch(e) {}
    }
    var ss    = SpreadsheetApp.openById(ROSTER_SHEET_ID);
    var sheet = ss.getSheetByName('roster') || ss.getSheetByName('Roster');
    if (!sheet) { Logger.log('_getAgentEmailMap: roster sheet not found'); return {}; }
    var data    = sheet.getDataRange().getValues();
    var headers = data[0];
    var nameCol = -1, emailCol = -1;
    headers.forEach(function(h, i) {
      var hl = h.toString().toLowerCase().trim().replace(/_/g, ' ');
      if (hl === 'agent name' || hl === 'agent_name')                     nameCol  = i;
      if (hl === 'team member email' || hl === 'member email' ||
          hl === 'agent email')                                            emailCol = i;
    });
    if (nameCol === -1 || emailCol === -1) {
      Logger.log('_getAgentEmailMap: missing columns — nameCol=' + nameCol + ' emailCol=' + emailCol);
      return {};
    }
    var map = {};
    data.slice(1).forEach(function(row) {
      var name  = (row[nameCol]  || '').toString().trim().toLowerCase();
      var email = (row[emailCol] || '').toString().trim();
      if (name && email) map[name] = email;
    });
    try { cache.put(cacheKey, JSON.stringify(map).substring(0, 95000), 2 * 60 * 60); } catch(e) {}
    _agentEmailMapInMemory = map;
    Logger.log('_getAgentEmailMap: loaded ' + Object.keys(map).length + ' entries');
    return map;
  } catch(e) {
    Logger.log('_getAgentEmailMap error: ' + e);
    return {};
  }
}

function lookupAgentEmail(participantName) {
  if (!participantName) return '';
  var map = _getAgentEmailMap();
  return map[participantName.trim().toLowerCase()] || '';
}

/**
 * Sends an admin notification email with the evaluation summary.
 * Called automatically after each new submitTranscript.
 *
 * @param {Object} formData - The submitted form data
 * @param {string} auditRef - The generated audit reference number
 */
function notifyAdmins(formData, auditRef) {
  try {
    var recipients = getAdminRecipients();
    if (!recipients.length) {
      Logger.log('notifyAdmins: No Admin/Dev recipients found — skipping');
      return;
    }

    var agentName     = formData.participant  || 'Unknown Agent';
    var sapId         = formData.sapId        || 'N/A';
    var analysisType  = formData.analysisType === 'sales'
                        ? 'Sales Analyzer' : 'Repeats & Transfer Analyzer';
    var interactionId = formData.interactionId || 'N/A';
    var observer      = formData.observerName || 'N/A';
    var direction     = formData.direction    || 'N/A';
    var duration      = formData.duration     || 'N/A';
    var teamLeader    = formData.teamLeader   || 'N/A';
    var opsManager    = formData.opsManager   || 'N/A';
    var lob           = formData.lineOfBusiness || 'N/A';
    var locale        = formData.locale       || 'N/A';
    var vtid          = formData.vtid         || 'N/A';
    var submittedAt   = new Date().toLocaleString();

    // Fix: truncate interactionId to prevent transcript blob leaking into email
    var cleanIntId = interactionId.replace(/[\r\n\t]/g, ' ').substring(0, 120);

    var evalUrl  = ScriptApp.getService().getUrl() + '?page=eval&ref=' + encodeURIComponent(auditRef);
    var subject  = '[Admin] New Audit Submitted — ' + auditRef + ' | ' + agentName;

    // ── Plain text body ───────────────────────────────────────────────────────
    var plainBody =
      'A new QA audit has been submitted.\n\n' +
      'AUDIT REFERENCE: ' + auditRef + '\n' +
      '─────────────────────────────────────\n' +
      'Submitted At:     ' + submittedAt + '\n' +
      'Observer:         ' + observer + '\n' +
      'Analysis Type:    ' + analysisType + '\n' +
      '\nAGENT DETAILS\n' +
      'Agent Name:       ' + agentName + '\n' +
      'SAP ID:           ' + sapId + '\n' +
      'VTID:             ' + vtid + '\n' +
      'Team Leader:      ' + teamLeader + '\n' +
      'Operations Mgr:   ' + opsManager + '\n' +
      'Line of Business: ' + lob + '\n' +
      'Locale / Site:    ' + locale + '\n' +
      '\nCALL DETAILS\n' +
      'Interaction ID:   ' + cleanIntId + '\n' +
      'Direction:        ' + direction + '\n' +
      'Duration:         ' + duration + '\n' +
      '─────────────────────────────────────\n\n' +
      'View full evaluation: ' + evalUrl + '\n\n' +
      'This is an automated notification from the NH Real Time Analyzer.\n';

    // ── HTML email body ───────────────────────────────────────────────────────
    var htmlBody = buildAdminEmailHTML(
      auditRef, submittedAt, observer, analysisType,
      agentName, sapId, vtid, teamLeader, opsManager, lob, locale,
      cleanIntId, direction, duration, evalUrl
    );

    // ── Send individually — MailApp works without extra OAuth grant ───────────
    var sent = [], failed = [];
    recipients.forEach(function(r) {
      try {
        MailApp.sendEmail({
          to:       r.email,
          subject:  subject,
          body:     plainBody,
          htmlBody: htmlBody,
          name:     'NH Call Analyzer — Admin Notifications'
        });
        sent.push(r.email);
        Logger.log('Sent to: ' + r.email);
      } catch(emailErr) {
        Logger.log('Failed to send to ' + r.email + ': ' + emailErr);
        failed.push(r.email);
      }
    });

    Logger.log('Admin notification sent to: ' + sent.join(', '));
    if (failed.length) Logger.log('Failed recipients: ' + failed.join(', '));

  } catch(e) {
    Logger.log('notifyAdmins error: ' + e.toString());
    // Non-fatal — don't break the main submission flow
  }
}

/** HTML table row helper for admin email */
function adminRow(label, value) {
  return '<tr>' +
    '<td style="padding:6px 10px 6px 0;font-weight:600;color:#54565A;' +
         'font-size:12px;width:38%;vertical-align:top">' + escEmail(label) + '</td>' +
    '<td style="padding:6px 0;font-size:12px;color:#1A1A2E;border-bottom:1px solid #F4F4F7">' +
         escEmail(value) + '</td>' +
    '</tr>';
}

/** Run from editor to see exactly what's in the Roster tab of the main spreadsheet */
function diagnoseRosterTabInMainSS() {
  var ss    = getOrCreateSpreadsheet();
  var sheet = ss.getSheetByName('Roster') || ss.getSheetByName('roster');
  if (!sheet) { Logger.log('No Roster tab found in main spreadsheet'); return; }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  Logger.log('Roster tab: ' + lastRow + ' rows, ' + lastCol + ' cols');

  // Print all headers
  if (lastRow >= 1) {
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    headers.forEach(function(h, i) {
      Logger.log('Col ' + String.fromCharCode(65+i) + ' [idx ' + i + ']: ' + h);
    });
  }

  // Print first 5 data rows
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, Math.min(5, lastRow-1), lastCol).getValues();
    data.forEach(function(row, i) {
      Logger.log('Row ' + (i+2) + ': ' + row.slice(0, Math.min(5, lastCol)).join(' | '));
    });
  }
}

/** One-time test: run from editor to verify recipients and email content */
function testAdminNotification() {
  var recipients = getAdminRecipients();
  Logger.log('=== Admin/Dev recipients ===');
  if (!recipients.length) {
    Logger.log('NONE FOUND — check that the Roster sheet has a "Role" column with value "Admin/Dev"');
    return;
  }
  recipients.forEach(function(r){
    Logger.log('  ' + r.name + ' <' + r.email + '>');
  });

  // Send a test email using sample data
  notifyAdmins({
    participant:    'Test Agent',
    sapId:          '9999999',
    vtid:           'VTI-TEST-001',
    analysisType:   'repeats',
    interactionId:  'TEST-INTERACTION-ID',
    observer:       'Observer Name',
    observerName:   'Observer Name',
    direction:      'Inbound',
    duration:       '15:00',
    teamLeader:     'Test Team Leader',
    opsManager:     'Test Ops Manager',
    lineOfBusiness: 'FFH',
    locale:         'TI El Salvador'
  }, 'NHA-TEST-0000');

  Logger.log('=== Test notification sent ===');
}

// ─────────────────────────────────────────────────────────────────────────────
// ONE-TIME REPAIR: Run this from the Apps Script editor (Run → repairSheetHeaders)
// It re-writes every sheet so columns match the current header definitions.
// Safe to run multiple times — it never deletes data, only reorders columns.
// ─────────────────────────────────────────────────────────────────────────────
function repairSheetHeaders() {
  var ss = getOrCreateSpreadsheet();
  Logger.log('=== Starting sheet header repair ===');

  _repairSheet(ss, TRANSCRIPTS_SHEET,    TRANSCRIPTS_HEADERS);
  _repairSheet(ss, ANALYSIS_SHEET,       ANALYSIS_HEADERS);
  _repairSheet(ss, DASHBOARD_DATA_SHEET, DASHBOARD_HEADERS);
  _repairSheet(ss, AUDIT_LOG_SHEET,      AUDIT_LOG_HEADERS);

  // Invalidate dashboard cache so next open reflects repaired data
  invalidateDashboardCache();

  Logger.log('=== Repair complete ===');
}

function _repairSheet(ss, sheetName, correctHeaders) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(sheetName + ': sheet not found, creating fresh');
    sheet = ss.insertSheet(sheetName);
    _writeHeaderRow(sheet, correctHeaders);
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    Logger.log(sheetName + ': empty, writing headers');
    _writeHeaderRow(sheet, correctHeaders);
    return;
  }

  // Read everything currently in the sheet
  var lastCol    = sheet.getLastColumn();
  var allData    = sheet.getRange(1, 1, lastRow, Math.max(lastCol, correctHeaders.length)).getValues();
  var currentHdr = allData[0].map(function(h) { return h.toString().trim(); });

  // Check if headers already match
  var alreadyCorrect = correctHeaders.every(function(h, i) { return currentHdr[i] === h; });
  if (alreadyCorrect && currentHdr.length === correctHeaders.length) {
    Logger.log(sheetName + ': headers already correct, skipping');
    return;
  }

  Logger.log(sheetName + ': repairing. Current headers: ' + currentHdr.join(' | '));
  Logger.log(sheetName + ': target headers:  ' + correctHeaders.join(' | '));

  // Build a map from current header name → column index
  var currentIndexOf = {};
  currentHdr.forEach(function(h, i) { if (h) currentIndexOf[h] = i; });

  // Reorder each data row to match correct headers
  var dataRows   = allData.slice(1); // skip existing header row
  var reordered  = dataRows.map(function(row) {
    return correctHeaders.map(function(h) {
      var idx = currentIndexOf[h];
      return (idx !== undefined && idx < row.length) ? row[idx] : '';
    });
  });

  // Clear the sheet and rewrite cleanly
  sheet.clearContents();

  // Write header row with styling
  _writeHeaderRow(sheet, correctHeaders);

  // Write data rows
  if (reordered.length > 0) {
    sheet.getRange(2, 1, reordered.length, correctHeaders.length).setValues(reordered);
  }

  Logger.log(sheetName + ': repaired — ' + reordered.length + ' data rows rewritten with ' + correctHeaders.length + ' columns');
}

function _writeHeaderRow(sheet, headers) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#4B286D')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  // Auto-resize columns for readability
  try { sheet.autoResizeColumns(1, headers.length); } catch(e) {}
}

function doGet(e) {
  var page = e && e.parameter && e.parameter.page;
  if (page === 'dashboard') return doGetDashboard();
  if (page === 'eval') {
    var ref = (e && e.parameter && e.parameter.ref) || '';
    var tmpl = HtmlService.createTemplateFromFile('EvalView');
    tmpl.auditRef = ref;
    tmpl.preloadedData = 'null';
    if (ref) {
      try {
        var evalData = getEvalViewData(ref);
        if (evalData && evalData.success) {
          // Escape </script> to prevent HTML parser from closing the JS block early
          tmpl.preloadedData = JSON.stringify(evalData).replace(/<\/script/gi, '<\\/script');
        }
      } catch(er) { Logger.log('EvalView preload failed: ' + er); }
    }
    return tmpl.evaluate()
      .setTitle('Real Time Feedback — Evaluation View')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('NH FCR, Transfer or Sales Call Analyzer')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Evaluation view — fetch all data for a single audit ref ──────────────────
function getEvalViewData(auditRef) {
  try {
    // ── Fast path: CacheService (avoids all sheet reads on repeat visits) ────────
    var sc       = CacheService.getScriptCache();
    var evKey    = 'ev2_' + Utilities.base64Encode(auditRef.trim()).substring(0, 200);
    var evCached = sc.get(evKey);
    if (evCached) { try { return JSON.parse(evCached); } catch(pe) {} }

    var ss = getOrCreateSpreadsheet();

    // 1. Audit_Log → metadata
    var logSheet = getOrCreateSheet(ss, AUDIT_LOG_SHEET);
    var logLast  = logSheet.getLastRow();
    var meta     = {};
    if (logLast >= 2) {
      var logHeaders = logSheet.getRange(1,1,1,logSheet.getLastColumn()).getValues()[0];
      var logData    = logSheet.getRange(2,1,logLast-1,logHeaders.length).getValues();
      for (var i = 0; i < logData.length; i++) {
        var row = logData[i];
        var ref = (row[logHeaders.indexOf('Audit Ref')] || '').toString().trim();
        if (ref === auditRef.trim()) {
          logHeaders.forEach(function(h, idx) { meta[h] = (row[idx] || '').toString(); });
          break;
        }
      }
    }

    // 2. Dashboard_Data → SMART fields + extra structured data
    var dSheet  = getOrCreateSheet(ss, DASHBOARD_DATA_SHEET);
    var dLast   = dSheet.getLastRow();
    var dash    = {};
    if (dLast >= 2) {
      var dHeaders = dSheet.getRange(1,1,1,dSheet.getLastColumn()).getValues()[0];
      var dData    = dSheet.getRange(2,1,dLast-1,dHeaders.length).getValues();
      for (var j = 0; j < dData.length; j++) {
        var dRow = dData[j];
        var dRef = (dRow[dHeaders.indexOf('Audit Ref')] || '').toString().trim();
        if (dRef === auditRef.trim()) {
          dHeaders.forEach(function(h, idx) { dash[h] = (dRow[idx] || '').toString(); });
          break;
        }
      }
    }

    // 3. Cache sheet → AI HTML (keyed by Interaction ID)
    // Read directly from the sheet — bypasses CacheService which truncates at 95KB
    var interactionId = meta['Interaction ID'] || dash['Interaction ID'] || '';
    var evalHtml = '';
    if (interactionId) {
      var rawType = (meta['Analysis Type'] || dash['Analysis Type'] || '').toLowerCase();
      var normalizedType = rawType.indexOf('sales') !== -1 ? 'sales' : 'repeats';
      try {
        var cSheet   = getOrCreateSheet(ss, CACHE_SHEET);
        var cfinder  = cSheet.getRange('A:A').createTextFinder(interactionId.trim()).matchEntireCell(true);
        var cmatches = cfinder.findAll();
        for (var ci = 0; ci < cmatches.length; ci++) {
          var crow  = cmatches[ci].getRow();
          if (crow < 2) continue;
          var cdata = cSheet.getRange(crow, 1, 1, 4).getValues()[0];
          var ctype = (cdata[1] || '').toString().trim().toLowerCase();
          if (ctype === normalizedType) {
            var chtml = cdata[3] ? cdata[3].toString() : '';
            if (!chtml) break;
            // Assemble continuation chunks (sales_2, repeats_2, etc.)
            var ckNum = 2;
            while (true) {
              var ckType = normalizedType + '_' + ckNum;
              var ckFound = false;
              for (var ck = 0; ck < cmatches.length; ck++) {
                var ckRow  = cmatches[ck].getRow();
                if (ckRow < 2) continue;
                var ckData = cSheet.getRange(ckRow, 1, 1, 4).getValues()[0];
                if ((ckData[1] || '').toString().trim().toLowerCase() === ckType) {
                  chtml += (ckData[3] || '').toString();
                  ckFound = true;
                  break;
                }
              }
              if (!ckFound) break;
              ckNum++;
            }
            evalHtml = sharedCSS() + fixBadgeClasses(chtml);
            break;
          }
        }
      } catch(ce) { Logger.log('Cache sheet read error: ' + ce); }
    }

    // 4. Recent evaluations by same SAP ID (last 5, excluding this one)
    var sapId   = meta['SAP ID'] || dash['SAP ID'] || '';
    var recent  = [];
    if (sapId && logLast >= 2) {
      var logHeaders2 = logSheet.getRange(1,1,1,logSheet.getLastColumn()).getValues()[0];
      var logData2    = logSheet.getRange(2,1,logLast-1,logHeaders2.length).getValues();
      logData2.forEach(function(r) {
        var rRef = (r[logHeaders2.indexOf('Audit Ref')] || '').toString().trim();
        var rSap = (r[logHeaders2.indexOf('SAP ID')]    || '').toString().trim();
        if (rSap === sapId && rRef !== auditRef.trim()) {
          recent.push({
            ref:  rRef,
            date: (r[logHeaders2.indexOf('Submitted At')] || '').toString().substring(0,10),
            ban:  (r[logHeaders2.indexOf('Interaction ID')] || '').toString()
          });
        }
      });
      recent = recent.slice(-5).reverse();
    }

    // 5. Web app base URL for links
    var baseUrl = ScriptApp.getService().getUrl();

    var result = {
      success: true,
      auditRef: auditRef,
      meta: meta,
      dash: dash,
      evalHtml: evalHtml,
      recent: recent,
      baseUrl: baseUrl
    };

    // Cache for 6 hours so repeat visits (and doGet injection) skip sheet reads
    try {
      var resultJson = JSON.stringify(result);
      if (resultJson.length < 90000) {
        sc.put(evKey, resultJson, 21600);
      }
    } catch(ce) {}

    return result;
  } catch(e) {
    Logger.log('getEvalViewData error: ' + e);
    return { success: false, error: e.toString() };
  }
}

// ── Fix badge colours deterministically based on score value ─────────────────
// The AI is unreliable at assigning correct badge classes.
// This scans every ai-badge span, reads the number, and forces the right class.
// Rules: 0 = zero(grey), 1–1.9 = bad(red), 2–2.9 = mid(amber), 3–5 = good(green)
function fixBadgeClasses(html) {
  if (!html) return html;
  return html.replace(
    /<span\s+class="ai-badge(?:\s+ai-badge-\w+)?"\s*>([\s\S]*?)<\/span>/gi,
    function(match, inner) {
      // Extract the leading number from text like "2.5/5" or "<strong>2.7/5</strong>"
      var stripped = inner.replace(/<[^>]+>/g, '').trim();
      var numMatch = stripped.match(/^(\d+(?:\.\d+)?)/);
      if (!numMatch) return match;
      var score = parseFloat(numMatch[1]);
      var cls   = score === 0 ? 'ai-badge-zero'
                : score >= 3  ? 'ai-badge-good'
                : score >= 2  ? 'ai-badge-mid'
                :               'ai-badge-bad';
      return '<span class="ai-badge ' + cls + '">' + inner + '</span>';
    }
  );
}

// ── Convert HTML to clean readable plain text for spreadsheet storage ─────────
function htmlToPlainText(html) {
  if (!html) return '';
  return html
    // Strip <style>...</style> blocks entirely (CSS text, not just tags)
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    // Strip <script>...</script> blocks
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    // Strip code fences
    .replace(/```html/gi, '').replace(/```/g, '')
    // Convert block elements to line breaks before stripping
    .replace(/<br\s*\/?>/gi,          '\n')
    .replace(/<\/p>/gi,               '\n')
    .replace(/<\/div>/gi,             '\n')
    .replace(/<\/li>/gi,              '\n')
    .replace(/<\/tr>/gi,              '\n')
    .replace(/<\/th>/gi,              '\t')
    .replace(/<\/td>/gi,              '\t')
    .replace(/<li[^>]*>/gi,           '• ')
    .replace(/<th[^>]*>/gi,           '')
    .replace(/<td[^>]*>/gi,           '')
    // Decode common HTML entities
    .replace(/&amp;/g,  '&')
    .replace(/&lt;/g,   '<')
    .replace(/&gt;/g,   '>')
    .replace(/&quot;/g, '"')
    .replace(/&#10003;/g, '✓')
    .replace(/&#10007;/g, '✗')
    .replace(/&#128681;/g, '🚩')
    .replace(/&#127908;/g, '🎤')
    .replace(/&#128161;/g, '💡')
    .replace(/&#127979;/g, '🎓')
    .replace(/&#9733;/g,   '★')
    .replace(/&#9888;/g,   '⚠')
    .replace(/&#128200;/g, '📈')
    .replace(/&#128222;/g, '📞')
    .replace(/&#128100;/g, '👤')
    .replace(/&#128203;/g, '📋')
    .replace(/&#128313;/g, '🔹')
    .replace(/&#127919;/g, '🎯')
    .replace(/&#128221;/g, '📝')
    .replace(/&#127775;/g, '🌟')
    .replace(/&[a-zA-Z0-9#]+;/g, ' ')
    // Strip all remaining HTML tags
    .replace(/<[^>]+>/g, '')
    // Clean up whitespace — collapse 3+ newlines to 2, trim
    .replace(/\t+/g, '  ')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

// Extract plain text from a named section in the AI HTML output.
// Tries multiple patterns to handle both Repeats and Sales AI output formats.
function extractTextBlock(html, keyword) {
  if (!html || !keyword) return '';
  try {
    var decode = function(s) {
      return s.replace(/<[^>]+>/g,'')
              .replace(/&amp;/g,'&').replace(/&lt;/g,'<')
              .replace(/&gt;/g,'>').replace(/&quot;/g,'"')
              .replace(/&#\d+;/g,' ').replace(/\s+/g,' ').trim();
    };

    // 1. ai-label-col table cell (repeats analysis table: Parameter | Finding | Recommendation)
    var re1 = new RegExp('<td[^>]*ai-label-col[^>]*>[^<]*' + keyword + '[^<]*<\\/td>\\s*<td[^>]*>(.*?)<\\/td>', 'is');
    var m1   = html.match(re1);
    if (m1) return decode(m1[1]).substring(0, 500);

    // 2. Plain td label then next td (generic table row)
    var re2 = new RegExp('<td[^>]*>\\s*' + keyword + '[^<]{0,30}<\\/td>\\s*<td[^>]*>(.*?)<\\/td>', 'is');
    var m2   = html.match(re2);
    if (m2) return decode(m2[1]).substring(0, 500);

    // 3. ai-title section header → ai-summary div (Call Summary)
    var re3 = new RegExp('<div[^>]*ai-title[^>]*>[^<]*' + keyword + '[^<]*<\\/div>\\s*<div[^>]*ai-summary[^>]*>(.*?)<\\/div>', 'is');
    var m3   = html.match(re3);
    if (m3) return decode(m3[1]).substring(0, 500);

    // 4. ai-title section header → next content div (general sections)
    var re4 = new RegExp('<div[^>]*ai-title[^>]*>[^<]*' + keyword + '[^<]*<\\/div>\\s*(?:<[^>]+>)?(.*?)(?=<div[^>]*ai-(?:title|section))', 'is');
    var m4   = html.match(re4);
    if (m4) return decode(m4[1]).substring(0, 500);

    // 5. ai-chip label → ai-chip-val span (info bar chips like Call Reason, Issue Resolved)
    var re5 = new RegExp('<span[^>]*ai-chip-label[^>]*>[^<]*' + keyword + '[^<]*<\\/span>\\s*<span[^>]*ai-chip-val[^>]*>([^<]*)<\\/span>', 'i');
    var m5   = html.match(re5);
    if (m5) return m5[1].trim().substring(0, 200);

    // 6. Broad keyword anywhere in text — last resort, grab surrounding sentence
    var reB = new RegExp('[^.]{0,100}' + keyword + '[^.]{0,200}\\.', 'i');
    var mB  = html.match(reB);
    if (mB) return decode(mB[0]).substring(0, 300);

    return '';
  } catch(e) { Logger.log('extractTextBlock error: ' + e); return ''; }
}

function getSapIdForParticipant(name) { return lookupSapId(name); }

// ─────────────────────────────────────────────────────────────────────────────
// AI ANALYTICS ENGINE
// Generates executive briefing from filtered dashboard data using FuelIX AI.
// Caches results in AI_Analytics sheet keyed by filter hash.
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Main entry: called from client with filtered rows + active filters.
 * Returns HTML briefing. Checks cache first, calls AI if miss.
 */
function getAIAnalytics(rows, filters, analysisType) {
  try {
    if (!rows || !rows.length) {
      return { success: false, error: 'No data to analyze. Apply filters that return at least 1 record.' };
    }

    // Build a cache key from filters + row count + analysis type
    var cacheKey = Utilities.base64Encode(
      analysisType + '|' + rows.length + '|' + JSON.stringify(filters)
    ).substring(0, 250);

    // Check sheet cache (lasts 24h)
    var cached = readAnalyticsCache(cacheKey);
    if (cached) {
      Logger.log('AI Analytics cache HIT');
      return { success: true, html: cached.html, fromCache: true, generatedAt: cached.generatedAt };
    }

    Logger.log('AI Analytics cache MISS — calling AI for ' + rows.length + ' rows');

    // Build context summary from rows
    var context = buildAnalyticsContext(rows, filters, analysisType);

    // Call AI
    var briefingJson = callAIForAnalytics(context, analysisType);
    if (!briefingJson) return { success: false, error: 'AI did not return a valid response.' };

    // Render to HTML
    var html = renderAnalyticsHTML(briefingJson, filters, analysisType, rows.length);

    // Cache result
    var generatedAt = new Date().toISOString();
    writeAnalyticsCache(cacheKey, html, generatedAt, rows.length, analysisType);

    return { success: true, html: html, fromCache: false, generatedAt: generatedAt };
  } catch(e) {
    Logger.log('getAIAnalytics error: ' + e);
    return { success: false, error: e.toString() };
  }
}

function getOverviewHighlightsLowlights(rows, filters) {
  try {
    if (!rows || !rows.length) return { success: false, error: 'No data.' };

    // Include content fingerprint so different locale/member filters with equal row counts
    // don't collide on the same cache key
    var rowFingerprint = rows.map(function(r) {
      return r['Interaction ID'] || r['Audit Ref'] || '';
    }).sort().join('|');
    var digestBytes = Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5,
      rowFingerprint + '|' + JSON.stringify(filters || {})
    );
    var digestHex = digestBytes.map(function(b) {
      return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0');
    }).join('');
    var cacheKey = 'ov_hl|' + rows.length + '|' + digestHex;

    var cached = readAnalyticsCache(cacheKey);
    if (cached) {
      try {
        var parsed = JSON.parse(cached.html);
        return { success: true, highlights: parsed.highlights, lowlights: parsed.lowlights,
                 fromCache: true, generatedAt: cached.generatedAt };
      } catch(e) {}
    }

    var context = buildAnalyticsContext(rows, filters || {}, 'repeats');
    var prompt =
      'You are a QA analytics expert for TELUS contact center operations. ' +
      'Based on the data summary below, return ONLY a valid JSON object — no extra text:\n' +
      '{\n' +
      '  "highlights": ["Specific positive finding with supporting data (max 20 words)", "highlight 2", "highlight 3"],\n' +
      '  "lowlights":  ["Specific gap or risk with supporting data (max 20 words)", "lowlight 2", "lowlight 3"]\n' +
      '}\n\nDATA SUMMARY:\n' + context;

    var payload = { model: FUELIX_CONFIG.model,
                    messages: [{ role: 'user', content: prompt }], stream: false };
    var options = {
      method: 'post', contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + FUELIX_CONFIG.apiKey,
                 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload), muteHttpExceptions: true, deadline: 120
    };
    var resp = UrlFetchApp.fetch(FUELIX_CONFIG.baseUrl + '/v1/chat/completions', options);
    if (resp.getResponseCode() !== 200)
      return { success: false, error: 'AI error ' + resp.getResponseCode() };

    var content = JSON.parse(resp.getContentText()).choices[0].message.content.trim();
    var match   = content.match(/\{[\s\S]*\}/);
    if (!match) return { success: false, error: 'No JSON in response' };
    var d       = JSON.parse(match[0]);

    var genAt = new Date().toISOString();
    writeAnalyticsCache(cacheKey,
      JSON.stringify({ highlights: d.highlights || [], lowlights: d.lowlights || [] }),
      genAt, rows.length, 'overview');

    return { success: true, highlights: d.highlights || [], lowlights: d.lowlights || [],
             fromCache: false, generatedAt: genAt };
  } catch(e) {
    Logger.log('getOverviewHighlightsLowlights error: ' + e);
    return { success: false, error: e.toString() };
  }
}

/** Build a rich text context from the filtered rows for the AI prompt */
function buildAnalyticsContext(rows, filters, analysisType) {
  var isSales = analysisType === 'sales';
  var total   = rows.length;

  // Aggregate stats
  var members    = dbUniqServer(rows.map(function(r){ return r['Team Member']||''; })).filter(Boolean);
  var rcaCats    = {};
  var callDrivers= {};
  var topOpps    = {};
  var critFlags  = {};
  var repeatPcts = [];
  var resolved   = 0;

  rows.forEach(function(r) {
    var cat = (r['RCA Category']||'').trim();
    if (cat) rcaCats[cat] = (rcaCats[cat]||0) + 1;

    var drv = (r['Call Driver']||'').trim();
    if (drv) callDrivers[drv] = (callDrivers[drv]||0) + 1;

    var opp = (r['RCA Sub-Parameter']||'').trim();
    if (opp) topOpps[opp] = (topOpps[opp]||0) + 1;

    var flags = (r['Critical Flags']||'').split(',');
    flags.forEach(function(f){ f=f.trim(); if(f && f!=='None') critFlags[f]=(critFlags[f]||0)+1; });

    var pct = parseFloat(r['Repeat Projection %']);
    if (!isNaN(pct) && pct > 0) repeatPcts.push(pct);

    if ((r['Issue Resolved']||'').toLowerCase().indexOf('yes') !== -1) resolved++;
  });

  var avgRepeat = repeatPcts.length
    ? Math.round(repeatPcts.reduce(function(s,v){return s+v;},0)/repeatPcts.length) : 0;

  // Find top performer (least critical flags)
  var memberFlagCount = {};
  rows.forEach(function(r){
    var m = r['Team Member']||''; var f = (r['Critical Flags']||'').split(',').filter(function(x){return x.trim()&&x.trim()!=='None';}).length;
    memberFlagCount[m] = (memberFlagCount[m]||0) + f;
  });
  var topPerformer = members.reduce(function(best, m){
    return (!best || (memberFlagCount[m]||0) < (memberFlagCount[best]||0)) ? m : best;
  }, '');

  // Sort helpers
  var sortDesc = function(obj) {
    return Object.keys(obj).sort(function(a,b){return obj[b]-obj[a];}).map(function(k){return k+' ('+obj[k]+')';});
  };

  var ctx = 'ANALYSIS TYPE: ' + (isSales ? 'Sales Analyzer' : 'Repeats & Transfer Analyzer') + '\n';
  ctx += 'TOTAL INTERACTIONS: ' + total + '\n';
  ctx += 'AGENTS EVALUATED: ' + members.join(', ') + '\n';
  ctx += 'DATE RANGE: ' + (filters.dateFrom||'All time') + ' to ' + (filters.dateTo||'present') + '\n';
  if (filters.members && filters.members.length) ctx += 'FILTERED TEAM MEMBERS: ' + filters.members.join(', ') + '\n';
  if (filters.leaders && filters.leaders.length) ctx += 'TEAM LEADERS: ' + filters.leaders.join(', ') + '\n';
  if (filters.lobs && filters.lobs.length) ctx += 'LINE OF BUSINESS: ' + filters.lobs.join(', ') + '\n';
  if (filters.locales && filters.locales.length) ctx += 'LOCALE/SITE: ' + filters.locales.join(', ') + '\n';
  ctx += '\nKEY METRICS:\n';
  ctx += '- Issues Resolved: ' + resolved + ' of ' + total + ' (' + Math.round(resolved/total*100) + '%)\n';
  ctx += '- Average Repeat Risk: ' + avgRepeat + '%\n';
  ctx += '\nTOP CALL DRIVERS:\n' + sortDesc(callDrivers).slice(0,5).map(function(x){return '- '+x;}).join('\n') + '\n';
  ctx += '\nRCA CATEGORIES:\n' + sortDesc(rcaCats).map(function(x){return '- '+x;}).join('\n') + '\n';
  ctx += '\nTOP OPPORTUNITIES (RCA Sub-Parameters):\n' + sortDesc(topOpps).slice(0,5).map(function(x){return '- '+x;}).join('\n') + '\n';
  ctx += '\nCRITICAL FLAGS (most frequent):\n' + sortDesc(critFlags).slice(0,8).map(function(x){return '- '+x;}).join('\n') + '\n';
  ctx += '\nTOP PERFORMER: ' + topPerformer + ' (fewest critical flags)\n';

  // Include call summaries for qualitative insight
  var summaries = rows.slice(0,5).map(function(r,i){
    return (i+1) + '. [' + (r['Team Member']||'?') + '] ' + (r['Call Summary']||'').substring(0,150);
  }).join('\n');
  ctx += '\nSAMPLE CALL SUMMARIES:\n' + summaries + '\n';

  return ctx;
}

/** Call FuelIX to generate the executive briefing JSON */
function callAIForAnalytics(context, analysisType) {
  var isSales = analysisType === 'sales';
  var prompt =
    'You are a senior QA analytics expert for TELUS contact center operations. ' +
    'Based on the following data summary, generate an executive intelligence briefing.\n\n' +
    'Return ONLY a valid JSON object with NO other text:\n' +
    '{\n' +
    '  "executiveSummary": "2-3 sentence overall assessment of performance and key finding",\n' +
    '  "highlights": ["Specific positive finding with data to back it up", "highlight 2", "highlight 3"],\n' +
    '  "lowlights": ["Specific gap or risk with data to back it up", "lowlight 2", "lowlight 3"],\n' +
    '  "topPerformer": { "name": "agent name", "insight": "why they stand out — be specific" },\n' +
    '  "criticalFlags": [\n' +
    '    { "label": "flag name", "count": 0, "type": "Agent Controllable or Process/Policy or Customer Driven", "impact": "one sentence impact on FCR or sales" }\n' +
    '  ],\n' +
    '  "recommendations": [\n' +
    '    {\n' +
    '      "category": "PROCESS or AGENT or TECHNOLOGY",\n' +
    '      "priority": "HIGH or MEDIUM or LOW",\n' +
    '      "smart_s": "Specific behavior or process to change — exact, not vague",\n' +
    '      "smart_m": "How success will be measured (e.g. FCR rate below 20%, offer rate above 80%)",\n' +
    '      "smart_a": "Attainable target — what the team can realistically hit",\n' +
    '      "smart_r": "Why this is realistic right now — e.g. the team already has the tools/knowledge, it is a small habit shift",\n' +
    '      "smart_t": "Timeline (e.g. within 2 weeks, by next QA cycle, before end of month)",\n' +
    '      "positioningStatement": "A sample statement a TL can use verbatim when coaching, e.g. \'[Agent], when a customer calls about X, your goal is to Y. Let me show you how: [exact phrase].\'"\n' +
    '    }\n' +
    '  ],\n' +
    '  "pillarInsights": [\n' +
    '    { "pillar": "pillar name", "score": "score or qualitative", "insight": "one sentence" }\n' +
    '  ],\n' +
    (isSales ? '  "salesInsights": "key sales performance observation in 2 sentences",\n' : '') +
    '  "coachingPriority": "The single most important SMART coaching action for this period. Format: S:[specific] | M:[metric] | A:[attainable] | R:[realistic] | T:[timeline]"\n' +
    '}\n\n' +
    'DATA SUMMARY:\n' + context;

  var payload = { model: FUELIX_CONFIG.model, messages: [{ role: 'user', content: prompt }], stream: false };
  var options = {
    method: 'post', contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + FUELIX_CONFIG.apiKey, 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload), muteHttpExceptions: true, deadline: 270
  };
  var resp = UrlFetchApp.fetch(FUELIX_CONFIG.baseUrl + '/v1/chat/completions', options);
  if (resp.getResponseCode() !== 200) { Logger.log('AI Analytics API error: ' + resp.getResponseCode()); return null; }

  var content = JSON.parse(resp.getContentText()).choices[0].message.content.trim();
  var match = content.match(/\{[\s\S]*\}/);
  if (!match) return null;
  try { return JSON.parse(match[0]); } catch(e) { Logger.log('Analytics JSON parse error: ' + e); return null; }
}

/** Render the AI briefing JSON into styled HTML matching the screenshot */
function renderAnalyticsHTML(d, filters, analysisType, rowCount) {
  var isSales  = analysisType === 'sales';
  var atype    = isSales ? 'Sales Analyzer — Agent' : 'Repeats & Transfer Analyzer — Agent';
  var genDate  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy');
  var e        = function(s){ return (s||'').toString().replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); };

  // Active filters summary
  var filterSummary = [];
  if (filters.dateFrom || filters.dateTo) filterSummary.push((filters.dateFrom||'*') + ' → ' + (filters.dateTo||'*'));
  if (filters.members  && filters.members.length)  filterSummary.push('Agent: ' + filters.members.join(', '));
  if (filters.leaders  && filters.leaders.length)  filterSummary.push('TL: ' + filters.leaders.join(', '));
  if (filters.lobs     && filters.lobs.length)     filterSummary.push('LOB: ' + filters.lobs.join(', '));
  if (filters.locales  && filters.locales.length)  filterSummary.push('Locale: ' + filters.locales.join(', '));
  if (filters.vtids    && filters.vtids.length)    filterSummary.push('VTID: ' + filters.vtids.join(', '));
  var filterLine = filterSummary.length ? filterSummary.join(' &nbsp;·&nbsp; ') : 'All records — no filters applied';

  var html = '<div class="analytics-body">';

  // Executive Briefing header
  html += '<div class="an-exec-header">' +
    '<div>' +
      '<span class="an-exec-label">EXECUTIVE BRIEFING</span>' +
      '<div class="an-exec-title">' + (isSales ? 'AI Sales Analytics' : 'AI Repeats and Transfer Analytics') + '</div>' +
      '<div class="an-exec-sub">Generated: ' + e(genDate) + ' &nbsp;·&nbsp; ' + e(atype) + '</div>' +
      '<div class="an-exec-filter">Filters: ' + filterLine + '</div>' +
    '</div>' +
  '</div>';

  // KPI strip
  html += '<div class="an-kpi-row">' +
    '<div class="an-kpi"><div class="an-kpi-lbl">INTERACTIONS ANALYZED</div><div class="an-kpi-val">' + rowCount + '</div></div>' +
    '<div class="an-kpi"><div class="an-kpi-lbl">ANALYSIS TYPE</div><div class="an-kpi-val" style="font-size:16px">' + e(isSales?'Sales':'Repeats') + '</div></div>' +
    (d.topPerformer ? '<div class="an-kpi"><div class="an-kpi-lbl">TOP PERFORMER</div><div class="an-kpi-val" style="font-size:16px">' + e(d.topPerformer.name||'') + '</div></div>' : '') +
  '</div>';

  // Highlights & Lowlights
  html += '<div class="an-hl-grid">';
  html += '<div class="an-hl-box an-hl-high"><div class="an-hl-title">&#9728; HIGHLIGHTS</div><ul>';
  (d.highlights||[]).forEach(function(h){ html += '<li>' + e(h) + '</li>'; });
  html += '</ul></div>';
  html += '<div class="an-hl-box an-hl-low"><div class="an-hl-title">&#9651; LOWLIGHTS</div><ul>';
  (d.lowlights||[]).forEach(function(l){ html += '<li>' + e(l) + '</li>'; });
  html += '</ul></div>';
  html += '</div>';

  // SMART Recommendations
  if (d.recommendations && d.recommendations.length) {
    html += '<div class="an-section"><div class="an-section-title">&#128161; SMART RECOMMENDATIONS</div>';
    d.recommendations.forEach(function(r, idx) {
      var pri = (r.priority || 'MEDIUM').toUpperCase();
      var priColor = pri === 'HIGH' ? '#C12335' : pri === 'LOW' ? '#2B8000' : '#8C4A00';
      html += '<div style="background:#FAFAFA;border:1px solid #E0D8F0;border-left:4px solid ' + priColor + ';border-radius:4px;padding:12px 14px;margin-bottom:10px">';
      html += '<div style="display:flex;align-items:center;gap:8px;margin-bottom:6px">' +
        '<span style="background:' + priColor + ';color:#fff;border-radius:3px;padding:2px 8px;font-size:11px;font-weight:700">' + e(pri) + '</span>' +
        '<span style="font-size:12px;font-weight:700;color:#4B286D">' + e(r.category || 'GENERAL') + '</span>' +
        '</div>';
      if (r.smart_s || r.smart_m || r.smart_a || r.smart_r || r.smart_t) {
        html += '<table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:6px">' +
          '<tr style="background:#F0EBF8">' +
          '<td style="padding:3px 8px;font-weight:700;color:#4B286D;width:5%">S</td>' +
          '<td style="padding:3px 8px">' + e(r.smart_s || '') + '</td></tr>' +
          '<tr><td style="padding:3px 8px;font-weight:700;color:#4B286D">M</td>' +
          '<td style="padding:3px 8px">' + e(r.smart_m || '') + '</td></tr>' +
          '<tr style="background:#F0EBF8"><td style="padding:3px 8px;font-weight:700;color:#4B286D">A</td>' +
          '<td style="padding:3px 8px">' + e(r.smart_a || '') + '</td></tr>' +
          '<tr><td style="padding:3px 8px;font-weight:700;color:#4B286D">R</td>' +
          '<td style="padding:3px 8px">' + e(r.smart_r || '') + '</td></tr>' +
          '<tr style="background:#F0EBF8"><td style="padding:3px 8px;font-weight:700;color:#4B286D">T</td>' +
          '<td style="padding:3px 8px">' + e(r.smart_t || '') + '</td></tr>' +
          '</table>';
      } else if (r.action) {
        html += '<p style="font-size:13px;margin:0 0 6px">' + e(r.action) + '</p>';
      }
      if (r.positioningStatement) {
        html += '<div style="background:#EDF7E6;border:1px solid #B3DFA0;border-radius:4px;padding:8px 10px;font-size:12px;color:#1A5E00;font-style:italic">' +
          '&#127908; ' + e(r.positioningStatement) + '</div>';
      }
      html += '</div>';
    });
    html += '</div>';
  }

  // Coaching Priority (SMART)
  if (d.coachingPriority) {
    html += '<div class="an-section" style="background:#FFF8E1;border:2px solid #F9A825;border-radius:6px;padding:12px 16px">' +
      '<div class="an-section-title" style="color:#E65100">&#127919; TOP COACHING PRIORITY (SMART)</div>' +
      '<p style="font-size:13px;line-height:1.7;margin:0">' + e(d.coachingPriority) + '</p>' +
      '</div>';
  }

  // Bottom 3-col: Pillar Analytics | Critical Flags | Top Performer
  html += '<div class="an-bottom-grid">';

  // Pillar Analytics
  if (d.pillarInsights && d.pillarInsights.length) {
    html += '<div class="an-bottom-card">';
    html += '<div class="an-bottom-title">&#128202; Pillar Analytics</div>';
    d.pillarInsights.forEach(function(p){
      html += '<div class="an-pillar-row"><span>' + e(p.pillar||'') + '</span><span class="an-pillar-score">' + e(p.score||'') + '</span></div>';
      if (p.insight) html += '<div class="an-pillar-insight">' + e(p.insight) + '</div>';
    });
    html += '</div>';
  }

  // Critical Flags
  if (d.criticalFlags && d.criticalFlags.length) {
    html += '<div class="an-bottom-card">';
    html += '<div class="an-bottom-title">&#128681; Critical Flags</div>';
    d.criticalFlags.forEach(function(f){
      var cls = (f.type||'').indexOf('Agent') !== -1 ? 'color:#C12335' : (f.type||'').indexOf('Process') !== -1 ? 'color:#8C4A00' : '';
      html += '<div class="an-flag-row"><span style="' + cls + '">' + e(f.label||'') + '</span>' +
        '<span class="an-flag-count" style="' + cls + '">' + e(f.count||'') + '</span></div>';
      if (f.impact) html += '<div class="an-pillar-insight">' + e(f.impact) + '</div>';
    });
    html += '</div>';
  }

  // Top Performer
  if (d.topPerformer) {
    html += '<div class="an-bottom-card an-top-performer">';
    html += '<div class="an-bottom-title">&#127775; Top Performer</div>';
    html += '<div class="an-tp-name">' + e(d.topPerformer.name||'') + '</div>';
    if (d.topPerformer.insight) html += '<div class="an-tp-insight">' + e(d.topPerformer.insight) + '</div>';
    html += '</div>';
  }

  html += '</div>'; // end bottom-grid

  // Sales-specific
  if (isSales && d.salesInsights) {
    html += '<div class="an-section"><div class="an-section-title">&#128200; Sales Insights</div>' +
      '<p style="font-size:13px;line-height:1.7">' + e(d.salesInsights) + '</p></div>';
  }

  html += '</div>'; // analytics-body
  return html;
}

/** Cache read — checks AI_Analytics sheet for matching key within 24h */
function readAnalyticsCache(key) {
  try {
    var ss    = getOrCreateSpreadsheet();
    var sheet = ss.getSheetByName(AI_ANALYTICS_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return null;
    var data  = sheet.getRange(2, 1, sheet.getLastRow()-1, 4).getValues();
    var cutoff = Date.now() - 24*60*60*1000; // 24h
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === key) {
        var ts = new Date(data[i][1]).getTime();
        if (ts > cutoff) return { html: data[i][2].toString(), generatedAt: data[i][1].toString() };
      }
    }
    return null;
  } catch(e) { return null; }
}

/** Cache write — stores result in AI_Analytics sheet */
function writeAnalyticsCache(key, html, generatedAt, rowCount, atype) {
  try {
    var ss    = getOrCreateSpreadsheet();
    var sheet = ss.getSheetByName(AI_ANALYTICS_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(AI_ANALYTICS_SHEET);
      sheet.getRange(1,1,1,5).setValues([['Cache Key','Generated At','HTML Result','Row Count','Analysis Type']]);
      sheet.getRange(1,1,1,5).setFontWeight('bold').setBackground('#4B286D').setFontColor('#fff');
      sheet.setFrozenRows(1);
    }
    // Remove old entry for same key
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var keys = sheet.getRange(2, 1, lastRow-1, 1).getValues();
      for (var i = keys.length-1; i >= 0; i--) {
        if (keys[i][0] === key) sheet.deleteRow(i+2);
      }
    }
    sheet.appendRow([key, generatedAt, html, rowCount, atype]);
    Logger.log('AI Analytics cached: ' + atype + ' (' + rowCount + ' rows)');
  } catch(e) { Logger.log('Analytics cache write error: ' + e); }
}

/** Send analytics briefing email */
function sendAnalyticsBriefing(toEmail, analyticsHTML, filters, analysisType, rowCount) {
  try {
    var isSales = analysisType === 'sales';
    var atype   = isSales ? 'Sales Analyzer' : 'Repeats & Transfer Analyzer';
    var subject = 'AI Performance Intelligence Report — ' + atype + ' (' + rowCount + ' records)';
    var genDate = new Date().toLocaleDateString('en-US',{year:'numeric',month:'long',day:'numeric'});

    var htmlBody =
      '<div style="font-family:Helvetica Neue,Helvetica,Arial,sans-serif;max-width:700px;color:#1A1A2E">' +
      '<div style="background:#4B286D;padding:18px 24px;border-radius:8px 8px 0 0">' +
        '<h2 style="color:#fff;margin:0;font-size:16px">AI Performance Intelligence Report</h2>' +
        '<p style="color:rgba(255,255,255,.7);margin:5px 0 0;font-size:12px">' + atype + ' &nbsp;·&nbsp; ' + genDate + '</p>' +
      '</div>' +
      '<div style="background:#fff;border:1px solid #E0E0E0;border-top:none;padding:24px;border-radius:0 0 8px 8px">' +
      analyticsHTML +
      '</div></div>';

    MailApp.sendEmail({
      to:      toEmail,
      subject: subject,
      body:    'AI Performance Intelligence Report — ' + atype + ' — ' + genDate,
      htmlBody: htmlBody,
      name:    'NH Call Analyzer — AI Analytics'
    });
    Logger.log('Analytics briefing sent to: ' + toEmail);
    return { success: true };
  } catch(e) {
    Logger.log('sendAnalyticsBriefing error: ' + e);
    return { success: false, error: e.toString() };
  }
}

// ── Generate PDF and return a temporary Drive view URL ────────────────────────
function generateAndServePDF(formData, htmlResult) {
  try {
    var auditRef  = formData.auditRef || _generateTempRef();
    var agentName = formData.participant || 'Agent';
    var fullHtml  = buildEvalFormHTML(formData, auditRef, htmlResult);
    var blob      = Utilities.newBlob(fullHtml, 'text/html', 'eval.html');
    var driveFile = Drive.Files.insert(
      { title: 'Eval_' + auditRef + '_tmp', mimeType: 'application/vnd.google-apps.document' },
      blob
    );
    var pdf = null;
    try {
      pdf = DriveApp.getFileById(driveFile.id).getAs('application/pdf');
      pdf.setName('Evaluation_' + agentName.replace(/\s+/g,'_') + '_' + auditRef + '.pdf');
    } finally {
      try { DriveApp.getFileById(driveFile.id).setTrashed(true); } catch(e) {}
    }

    // Save PDF to Drive and return link
    var folder   = DriveApp.getRootFolder();
    var pdfFile  = folder.createFile(pdf);
    var sharingOk = true;
    try { pdfFile.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW); }
    catch(se) { sharingOk = false; Logger.log('PDF setSharing failed: ' + se); }

    // Note: Drive API v2 has no TTL/scheduled-trash. File persists until manually deleted.

    return { success: true, url: pdfFile.getDownloadUrl() || pdfFile.getUrl(), name: pdfFile.getName(), sharingFailed: !sharingOk };
  } catch(e) {
    Logger.log('generateAndServePDF error: ' + e);
    return { success: false, error: e.toString() };
  }
}

function _generateTempRef() {
  var now = new Date();
  return 'NHA-' + now.getFullYear() + String(now.getMonth()+1).padStart(2,'0') + String(now.getDate()).padStart(2,'0') + '-TEMP';
}

// ── Retrieve cached evaluation HTML by Audit Ref (for dashboard link) ─────────
function getEvaluationByAuditRef(auditRef) {
  try {
    var ss     = getOrCreateSpreadsheet();
    var sheet  = getOrCreateSheet(ss, ANALYSIS_SHEET);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;
    var data = sheet.getRange(2, 1, lastRow - 1, ANALYSIS_HEADERS.length).getValues();
    // Col A = Audit Ref (index 0), Col G = AI Analysis Result (index 6)
    var refIdx  = ANALYSIS_HEADERS.indexOf('Audit Ref');
    var htmlIdx = ANALYSIS_HEADERS.indexOf('AI Analysis Result');
    for (var i = 0; i < data.length; i++) {
      if ((data[i][refIdx] || '').toString().trim() === auditRef.trim()) {
        var stored = (data[i][htmlIdx] || '').toString();
        // If plain text (already cleaned), return null — can't show as HTML
        if (stored.indexOf('<') === -1) {
          // Try Cache sheet for the original HTML
          return getEvaluationFromCache(auditRef);
        }
        return { html: sharedCSS() + fixBadgeClasses(stored), css: '' };
      }
    }
    // Not in Analysis sheet — try cache
    return getEvaluationFromCache(auditRef);
  } catch(e) {
    Logger.log('getEvaluationByAuditRef error: ' + e);
    return null;
  }
}

function getEvaluationFromCache(auditRef) {
  try {
    // Find Interaction ID for this Audit Ref from Audit_Log
    var ss       = getOrCreateSpreadsheet();
    var logSheet = getOrCreateSheet(ss, AUDIT_LOG_SHEET);
    var logLast  = logSheet.getLastRow();
    if (logLast < 2) return null;
    var logData  = logSheet.getRange(2, 1, logLast - 1, 3).getValues();
    var interactionId = '';
    var analysisType  = '';
    for (var i = 0; i < logData.length; i++) {
      if ((logData[i][0] || '').toString().trim() === auditRef.trim()) {
        interactionId = (logData[i][2] || '').toString().trim();
        break;
      }
    }
    if (!interactionId) return null;
    // Now look up cache by interaction ID
    var cacheSheet = getOrCreateSheet(ss, CACHE_SHEET);
    var cacheFind  = cacheSheet.getRange('A:A').createTextFinder(interactionId).matchEntireCell(true).findAll();
    if (!cacheFind.length) return null;
    var row     = cacheFind[0].getRow();
    var cached  = cacheSheet.getRange(row, 4).getValue().toString();
    if (!cached) return null;
    return { html: sharedCSS() + fixBadgeClasses(cached), css: '' };
  } catch(e) {
    Logger.log('getEvaluationFromCache error: ' + e);
    return null;
  }
}

// ── Chat assistant — general QA inquiries, positioning, grammar ───────────────
function chatAssistant(userMessage, recentHistory) {
  try {
    var systemPrompt =
      'You are a helpful QA coaching assistant for a TELUS contact center quality team. ' +
      'You specialise in:\n' +
      '1. Quality assurance and FCR (First Call Resolution) coaching\n' +
      '2. Writing sample positioning statements for coaches to use when giving feedback to agents\n' +
      '3. Grammar checking and correcting text\n' +
      '4. Improving sentence construction for professional coaching notes\n' +
      '5. Answering general questions about call center quality metrics\n\n' +
      'Keep responses concise and practical. For positioning statements, use a coaching tone. ' +
      'For grammar/sentence help, show the corrected version clearly.';

    // Build proper message array: system role + alternating user/assistant turns
    var messages = [{ role: 'system', content: systemPrompt }];
    if (recentHistory && recentHistory.length > 1) {
      recentHistory.slice(0, -1).forEach(function(h) {
        messages.push({ role: h.role === 'user' ? 'user' : 'assistant', content: h.text || '' });
      });
    }
    messages.push({ role: 'user', content: userMessage });

    var payload = {
      model:    FUELIX_CONFIG.model,
      messages: messages,
      stream:   false
    };
    var options = {
      method:      'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + FUELIX_CONFIG.apiKey,
        'Content-Type':  'application/json'
      },
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    };
    var response = UrlFetchApp.fetch(FUELIX_CONFIG.baseUrl + '/v1/chat/completions', options);
    if (response.getResponseCode() !== 200) {
      return 'AI service error (' + response.getResponseCode() + '). Please try again.';
    }
    var data = JSON.parse(response.getContentText());
    return data.choices[0].message.content.trim();
  } catch(e) {
    Logger.log('chatAssistant error: ' + e);
    return 'Sorry, something went wrong: ' + e.toString();
  }
}
function getTranscriptMetadata(txt)   { return parseTranscriptMetadata(txt); }
function getRosterBySapId(sapId)      { return lookupBySapId(sapId); }
function getRosterAll()               { return getAllRosterData(); }
function getObserverInfo() {
  var obs = resolveObserver() || {};

  // Get email — try both Session methods independently
  var userEmail = '';
  try { userEmail = Session.getActiveUser().getEmail()  || ''; } catch(e) {}
  try { if (!userEmail) userEmail = Session.getEffectiveUser().getEmail() || ''; } catch(e) {}
  userEmail = userEmail.toLowerCase().trim();

  // Always populate email/name so Observer display works
  if (!obs.email && userEmail) obs.email = userEmail;
  if (!obs.name  && userEmail) obs.name  = userEmail.split('@')[0];

  // Admin check — compare username against hardcoded ADMIN_USERNAMES list (no external sheet needed)
  var userLocal = userEmail.split('@')[0];
  obs.isAdmin = !!userLocal && ADMIN_USERNAMES.indexOf(userLocal) !== -1;

  return obs;
}
function getAuditLogData()            { return readAuditLog(); }

// ── Generate unique Audit Reference Number: NHA-YYYYMMDD-XXXX ────────────────
// Uses CacheService to persist today's counter — avoids scanning the full sheet.
function generateAuditRef() {
  var now     = new Date();
  var y       = now.getFullYear();
  var m       = String(now.getMonth() + 1).padStart(2, '0');
  var d       = String(now.getDate()).padStart(2, '0');
  var dateStr = '' + y + m + d;
  var cacheKey = 'audit_ref_seq_' + dateStr;

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    var cache   = CacheService.getScriptCache();
    var current = parseInt(cache.get(cacheKey) || '0', 10);
    var seq     = current + 1;
    // TTL: expire at midnight (seconds remaining in the day)
    var msLeft  = new Date(y, now.getMonth(), now.getDate() + 1).getTime() - now.getTime();
    cache.put(cacheKey, String(seq), Math.floor(msLeft / 1000));
    return 'NHA-' + dateStr + '-' + String(seq).padStart(4, '0');
  } catch(e) {
    Logger.log('generateAuditRef error: ' + e);
    return 'NHA-' + dateStr + '-' + String(Math.floor(Math.random() * 9999)).padStart(4, '0');
  } finally {
    try { lock.releaseLock(); } catch(re) {}
  }
}

// ── Resolve observer from logged-in user email via roster ─────────────────────
// Returns { name, sapId, email }
function resolveObserver() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (!email) return { name: '', sapId: '', email: '' };

    // Cache observer result per email for 8 hours
    var cache    = CacheService.getUserCache();
    var cacheKey = 'observer_' + email.replace(/[^a-zA-Z0-9]/g, '_');
    var cached   = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch(e) {}
    }

    var data       = _getRosterSheetData();  // uses cached roster
    var emailLower = email.toLowerCase();

    var result = null;
    for (var i = 0; i < data.length; i++) {
      // Search email columns
      for (var j = 0; j < data[i].length; j++) {
        if (data[i][j] && data[i][j].toString().toLowerCase() === emailLower) {
          result = {
            name:  (data[i][ROSTER_COL_AGENT_NAME] || '').toString().trim(),
            sapId: (data[i][ROSTER_COL_SAP_ID]     || '').toString().trim(),
            email: email
          };
          break;
        }
      }
      if (result) break;
      // Fallback: construct email match
      var agentName = (data[i][ROSTER_COL_AGENT_NAME] || '').toString().trim();
      if (agentName) {
        var parts = agentName.split(/\s+/);
        if (parts.length >= 2) {
          var constructed = (parts[0] + '.' + parts[parts.length - 1] + '@telus.com').toLowerCase();
          if (constructed === emailLower) {
            result = { name: agentName, sapId: (data[i][ROSTER_COL_SAP_ID] || '').toString().trim(), email: email };
            break;
          }
        }
      }
    }
    if (!result) result = { name: email.split('@')[0], sapId: '', email: email };

    // Cache for 8 hours
    try { cache.put(cacheKey, JSON.stringify(result), 8 * 60 * 60); } catch(e) {}
    return result;
  } catch(e) {
    Logger.log('resolveObserver error: ' + e);
    return { name: '', sapId: '', email: '' };
  }
}

// ── Read Audit_Log for dashboard tab 3 (cached 10 min) ───────────────────────
var AUDIT_LOG_CACHE_KEY = 'audit_log_data_v2';
var AUDIT_LOG_CACHE_TTL = 20 * 60; // 20 minutes — localStorage handles the faster layer

function readAuditLog() {
  try {
    var cache  = CacheService.getScriptCache();
    var cached = cache.get(AUDIT_LOG_CACHE_KEY);
    if (cached) {
      try { return JSON.parse(cached); } catch(e) {}
    }

    var ss    = getOrCreateSpreadsheet();
    var sheet = getOrCreateSheet(ss, AUDIT_LOG_SHEET);
    ensureHeaders(sheet, AUDIT_LOG_HEADERS);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var data   = sheet.getRange(2, 1, lastRow - 1, AUDIT_LOG_HEADERS.length).getValues();
    var result = data.map(function(row) {
      var obj = {};
      AUDIT_LOG_HEADERS.forEach(function(h, i) {
        obj[h] = row[i] !== undefined ? row[i].toString() : '';
      });
      return obj;
    });

    try { cache.put(AUDIT_LOG_CACHE_KEY, JSON.stringify(result).substring(0, 95000), AUDIT_LOG_CACHE_TTL); } catch(e) {}
    return result;
  } catch(e) { Logger.log('readAuditLog: ' + e); return []; }
}

function invalidateAuditLogCache() {
  try { CacheService.getScriptCache().remove(AUDIT_LOG_CACHE_KEY); } catch(e) {}
}

// Dashboard entry point — served at ?page=dashboard
function doGetDashboard() {
  return HtmlService
    .createHtmlOutputFromFile('Dashboard')
    .setTitle('NH Call Analyzer — Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Dashboard data cache keys ─────────────────────────────────────────────────
var DB_CACHE_KEY      = 'dashboard_data_v2';
var DB_CACHE_META_KEY = 'dashboard_data_meta_v2';
var DB_CACHE_TTL      = 60 * 60;   // 60 minutes — extended since localStorage is the fast layer
// CacheService limit is 100 KB per key — split into chunks if needed
var DB_CACHE_CHUNK    = 90000;      // bytes per chunk (safe margin under 100 KB)

// ── Read from sheet and return raw rows ───────────────────────────────────────
function readDashboardSheet() {
  var ss    = getOrCreateSpreadsheet();
  var sheet = getOrCreateSheet(ss, DASHBOARD_DATA_SHEET);
  ensureHeaders(sheet, DASHBOARD_HEADERS);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  // Read actual headers from the sheet so any added columns (e.g. Sales Attempted)
  // are always mapped correctly regardless of DASHBOARD_HEADERS index order.
  var actualHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return data.map(function(row) {
    var obj = {};
    actualHeaders.forEach(function(h, i) {
      if (h) obj[h] = row[i] !== undefined ? row[i].toString() : '';
    });
    return obj;
  });
}

// ── Write rows to CacheService in chunks ─────────────────────────────────────
function writeDashboardCache(rows) {
  var cache   = CacheService.getScriptCache();
  var json    = JSON.stringify(rows);
  var chunks  = [];
  for (var i = 0; i < json.length; i += DB_CACHE_CHUNK) {
    chunks.push(json.slice(i, i + DB_CACHE_CHUNK));
  }
  var meta = { chunks: chunks.length, ts: new Date().toISOString(), rowCount: rows.length };
  cache.put(DB_CACHE_META_KEY, JSON.stringify(meta), DB_CACHE_TTL);
  chunks.forEach(function(chunk, idx) {
    cache.put(DB_CACHE_KEY + '_' + idx, chunk, DB_CACHE_TTL);
  });
  Logger.log('Dashboard cache written: ' + rows.length + ' rows, ' + chunks.length + ' chunk(s)');
}

// ── Read rows from CacheService ───────────────────────────────────────────────
function readDashboardCache() {
  var cache = CacheService.getScriptCache();
  var metaStr = cache.get(DB_CACHE_META_KEY);
  if (!metaStr) return null;
  try {
    var meta   = JSON.parse(metaStr);
    var json   = '';
    for (var i = 0; i < meta.chunks; i++) {
      var chunk = cache.get(DB_CACHE_KEY + '_' + i);
      if (!chunk) return null;   // a chunk expired — treat as cache miss
      json += chunk;
    }
    return { rows: JSON.parse(json), ts: meta.ts, rowCount: meta.rowCount };
  } catch(e) {
    Logger.log('readDashboardCache parse error: ' + e);
    return null;
  }
}

// ── Invalidate dashboard cache (call after new data is written) ───────────────
function invalidateDashboardCache() {
  var cache = CacheService.getScriptCache();
  var metaStr = cache.get(DB_CACHE_META_KEY);
  if (!metaStr) return;
  try {
    var meta = JSON.parse(metaStr);
    var keys = [DB_CACHE_META_KEY];
    for (var i = 0; i < meta.chunks; i++) keys.push(DB_CACHE_KEY + '_' + i);
    cache.removeAll(keys);
    Logger.log('Dashboard cache invalidated');
  } catch(e) {}
}

// ── Public: cached load (used on dashboard open) ──────────────────────────────
// Returns { rows, ts, rowCount, fromCache }
function getDashboardData() {
  try {
    var cached = readDashboardCache();
    if (cached) {
      Logger.log('Dashboard cache HIT — ' + cached.rowCount + ' rows from ' + cached.ts);
      return { rows: cached.rows, ts: cached.ts, rowCount: cached.rowCount, fromCache: true };
    }
    Logger.log('Dashboard cache MISS — reading sheet');
    var rows = readDashboardSheet();
    writeDashboardCache(rows);
    var ts = new Date().toISOString();
    return { rows: rows, ts: ts, rowCount: rows.length, fromCache: false };
  } catch(e) {
    Logger.log('getDashboardData error: ' + e);
    return { rows: [], ts: '', rowCount: 0, fromCache: false };
  }
}

// ── Public: force refresh (used by Refresh button) ────────────────────────────
function getDashboardDataFresh() {
  try {
    invalidateDashboardCache();
    var rows = readDashboardSheet();
    writeDashboardCache(rows);
    var ts = new Date().toISOString();
    Logger.log('Dashboard force-refreshed: ' + rows.length + ' rows');
    return { rows: rows, ts: ts, rowCount: rows.length, fromCache: false };
  } catch(e) {
    Logger.log('getDashboardDataFresh error: ' + e);
    return { rows: [], ts: '', rowCount: 0, fromCache: false };
  }
}

// ── Look up audit ref for a cached interaction ────────────────────────────────
function findAuditRefForInteraction(interactionId, analysisType) {
  try {
    var ss    = getOrCreateSpreadsheet();
    var sheet = getOrCreateSheet(ss, AUDIT_LOG_SHEET);
    ensureHeaders(sheet, AUDIT_LOG_HEADERS);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return '';
    var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); // col A=AuditRef, C=InteractionId
    for (var i = 0; i < data.length; i++) {
      if (data[i][2] && data[i][2].toString().trim() === interactionId.trim()) {
        return data[i][0].toString().trim();
      }
    }
    return '';
  } catch(e) { return ''; }
}

// Called after email sent — update Dashboard_Data and Audit_Log
function updateDashboardPDFLink(interactionId, pdfLink, recipients, auditRef) {
  try {
    var ss = getOrCreateSpreadsheet();

    // Update Dashboard_Data — col C = Interaction ID (index 3), col W=PDF, col X=Status
    var dSheet  = getOrCreateSheet(ss, DASHBOARD_DATA_SHEET);
    var lastRow = dSheet.getLastRow();
    if (lastRow >= 2) {
      // Resolve column positions by header name — never hardcode
      var dHeaderRow = dSheet.getRange(1, 1, 1, dSheet.getLastColumn()).getValues()[0];
      var pdfCol    = dHeaderRow.indexOf('PDF Email Link') + 1;
      var dStatusCol= dHeaderRow.indexOf('Email Status')   + 1;
      var ids = dSheet.getRange(2, 3, lastRow - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (ids[i][0].toString().trim() === interactionId.trim()) {
          if (pdfCol    > 0) dSheet.getRange(i + 2, pdfCol).setValue(pdfLink || '');
          if (dStatusCol > 0) dSheet.getRange(i + 2, dStatusCol).setValue('Sent');
          break;
        }
      }
    }

    // Update Audit_Log — resolve columns by header name
    var logSheet  = getOrCreateSheet(ss, AUDIT_LOG_SHEET);
    var logLast   = logSheet.getLastRow();
    if (logLast >= 2) {
      var logHeaderRow  = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
      var logStatusCol  = logHeaderRow.indexOf('Email Status') + 1;
      var logRecipCol   = logHeaderRow.indexOf('Recipients')   + 1;
      var logIds = logSheet.getRange(2, 3, logLast - 1, 1).getValues();
      for (var j = 0; j < logIds.length; j++) {
        if (logIds[j][0].toString().trim() === interactionId.trim()) {
          if (logStatusCol > 0) logSheet.getRange(j + 2, logStatusCol).setValue('Sent');
          if (logRecipCol  > 0) logSheet.getRange(j + 2, logRecipCol).setValue((recipients || []).join(', '));
          break;
        }
      }
    }

    invalidateDashboardCache();
  } catch(e) { Logger.log('updateDashboardPDFLink: ' + e); }
}

// ─────────────────────────────────────────────────────────────────────────────
// Main submission — cache check first, then AI if needed
// ─────────────────────────────────────────────────────────────────────────────
function submitTranscript(formData) {
  try {
    if (!formData.sapId || !formData.transcript) {
      return { success: false, error: 'SAP ID and transcript are required.' };
    }

    var analysisType  = formData.analysisType || 'repeats';
    var spreadsheet   = getOrCreateSpreadsheet();
    var now           = new Date();

    // ── Server-side metadata extraction — reliable fallback if autofill missed ─
    var transcriptMeta = parseTranscriptMetadata(formData.transcript);

    var interactionId = (formData.interactionId || '').trim()
                     || (transcriptMeta.interactionId || '').trim();
    var startTime     = (formData.startTime || '').trim()
                     || (transcriptMeta.startTime || '').trim();
    var duration      = (formData.duration || '').trim()
                     || (transcriptMeta.duration || '').trim();
    var direction     = (formData.direction || '').trim()
                     || (transcriptMeta.direction || '').trim();
    var participant   = (formData.participant || '').trim()
                     || (transcriptMeta.participant || '').trim();
    var customerBAN   = (formData.customerBAN || '').trim()
                     || (transcriptMeta.customerBAN || '').trim()
                     || (transcriptMeta.customerName || '').trim();

    Logger.log('Interaction ID resolved: "' + interactionId + '" (form: "' + (formData.interactionId||'') + '", transcript: "' + (transcriptMeta.interactionId||'') + '")');

    // ── 1. Cache check ────────────────────────────────────────────────────────
    if (interactionId) {
      var cached = findCachedResult(interactionId, analysisType);
      if (cached) {
        Logger.log('Returning cached result for: ' + interactionId);
        // Still return the auditRef from the cache row if available
        var cachedRef = findAuditRefForInteraction(interactionId, analysisType);
        return {
          success:      true,
          html:         sharedCSS() + fixBadgeClasses(cached),
          sheetUrl:     spreadsheet.getUrl(),
          fromCache:    true,
          auditRef:     cachedRef,
          observerName: (formData.observerName || '')
        };
      }
    }

    // ── 2. Generate audit reference + resolve observer ────────────────────────
    var auditRef     = generateAuditRef();
    var observerName = (formData.observerName || '').trim();
    var analysisLabel = analysisType === 'sales' ? 'Sales Analyzer' : 'Repeats & Transfer Analyzer';

    // Resolve locale — fall back to roster lookup if the form field was blank
    var resolvedLocale = (formData.locale || '').trim();
    if (!resolvedLocale && formData.sapId) {
      try {
        var rosterRow = lookupBySapId(formData.sapId.trim());
        if (rosterRow && rosterRow.locale) resolvedLocale = rosterRow.locale;
      } catch(le) { Logger.log('Locale fallback lookup failed (non-fatal): ' + le); }
    }

    // ── 3. Save transcript row ────────────────────────────────────────────────
    try {
      var tSheet = getOrCreateSheet(spreadsheet, TRANSCRIPTS_SHEET);
      ensureHeaders(tSheet, TRANSCRIPTS_HEADERS);
      appendRow(tSheet, [
        auditRef, now,
        formData.sapId || '', participant,
        formData.teamLeader || '', formData.opsManager || '',
        formData.lineOfBusiness || '', resolvedLocale,
        observerName,
        startTime, customerBAN, interactionId,
        direction, duration,
        analysisLabel, formData.transcript,
        formData.agentEmail || ''
      ]);
    } catch(we) {
      var msg = we.toString();
      if (msg.indexOf('permission') !== -1 || msg.indexOf('document') !== -1 || msg.indexOf('access') !== -1) {
        return {
          success: false,
          error: 'Access denied: You need Editor (not Viewer) access to the "' + SPREADSHEET_NAME + '" spreadsheet. ' +
                 'Please ask your admin (danzen.palisoc@telus.com) to update your sharing permissions.'
        };
      }
      throw we;
    }

    // ── 4. Run AI ─────────────────────────────────────────────────────────────
    var rawAI = analyzeTranscript(formData.transcript, analysisType);
    Logger.log('AI response length: ' + rawAI.length);

    var html = fixBadgeClasses(
      rawAI
        .replace(/^```html\s*/i, '')
        .replace(/^```\s*/,      '')
        .replace(/\s*```$/,      '')
        .trim()
    );

    // ── 5. Save to Cache sheet FIRST — so a retry after any downstream failure ──
    //      hits the cache instead of re-running the paid AI call and duplicating rows.
    if (interactionId) {
      saveCachedResult(interactionId, analysisType, html);
    }

    // ── 6. Save to Analysis sheet (plain text only — no HTML) ─────────────────
    var aSheet = getOrCreateSheet(spreadsheet, ANALYSIS_SHEET);
    ensureHeaders(aSheet, ANALYSIS_HEADERS);
    appendRow(aSheet, [
      auditRef, now,
      formData.sapId || '', formData.participant || '',
      observerName, analysisLabel, htmlToPlainText(html)
    ]);

    var warnings = [];

    // ── 7. Write to Dashboard_Data ────────────────────────────────────────────
    try {
      var dSheet = getOrCreateSheet(spreadsheet, DASHBOARD_DATA_SHEET);
      ensureHeaders(dSheet, DASHBOARD_HEADERS);
      var callReason       = extractTextBlock(html, 'Call Reason')      || '';
      var callSummary      = extractTextBlock(html, 'Call Summary')      || '';
      var opportunities    = extractTextBlock(html, 'Opportunities')     || '';
      var recommendations  = extractTextBlock(html, 'Recommendation')    || '';
      var criticalFlags    = extractTextBlock(html, 'Critical Flag')     || '';
      var repeatPct        = extractTextBlock(html, 'Repeat Projection') || '';
      var issueResolved    = extractTextBlock(html, 'Issue Resolution')  || '';
      var transferOccurred = extractTextBlock(html, 'Transfer')          || '';

      // Structured RCA fields — fast HTML parse for flags/SMART, then AI call for RCA category
      var structured = null;
      try { structured = extractStructuredRCAFromHTML(html); } catch(re) {}
      var callDriver        = structured ? structured.callDriver        : '';
      var rcaCategory       = structured ? structured.rcaCategory       : '';
      var rcaSubParameter   = structured ? structured.rcaSubParameter   : '';
      var topOpportunity    = structured ? structured.topOpportunity    : '';
      var productOpportunity= structured ? structured.productOpportunity: '';
      var salesAttempted    = '';


      var smartS            = structured ? structured.smartS            : '';
      var smartM            = structured ? structured.smartM            : '';
      var smartA            = structured ? structured.smartA            : '';
      var smartR            = structured ? structured.smartR            : '';
      var smartT            = structured ? structured.smartT            : '';
      if (structured && structured.callSummaryShort && !callSummary)
        callSummary = structured.callSummaryShort;

      appendRow(dSheet, [
        auditRef, now, interactionId,
        formData.sapId || '', participant,
        formData.teamLeader || '', formData.opsManager || '',
        formData.lineOfBusiness || '', resolvedLocale,
        formData.vtid || '',
        observerName,
        '',               // Department
        direction, duration,
        analysisLabel,
        callReason, callSummary, opportunities, recommendations,
        criticalFlags, repeatPct, issueResolved, transferOccurred,
        callDriver, rcaCategory, rcaSubParameter, topOpportunity, productOpportunity, salesAttempted,
        smartS, smartM, smartA, smartR, smartT,
        'Pending', 'Not Sent', formData.agentEmail || ''
      ]);
      invalidateDashboardCache();
    } catch(de) {
      Logger.log('Dashboard_Data write error (non-fatal): ' + de);
      warnings.push('Dashboard entry failed to save — contact admin with ref ' + auditRef);
    }

    // ── 8. Write to Audit_Log ─────────────────────────────────────────────────
    try {
      var logSheet = getOrCreateSheet(spreadsheet, AUDIT_LOG_SHEET);
      ensureHeaders(logSheet, AUDIT_LOG_HEADERS);
      appendRow(logSheet, [
        auditRef, now, interactionId,
        formData.sapId || '', participant,
        formData.teamLeader || '', formData.opsManager || '',
        formData.lineOfBusiness || '', resolvedLocale,
        formData.vtid || '',
        observerName, analysisLabel,
        direction, duration,
        repeatPct || '', issueResolved || '', transferOccurred || '',
        'Not Sent', '', formData.agentEmail || ''
      ]);
      invalidateAuditLogCache(); // bust cache so next read is fresh
    } catch(le) {
      Logger.log('Audit_Log write error (non-fatal): ' + le);
      warnings.push('Audit log entry failed to save — contact admin with ref ' + auditRef);
    }

    return {
      success:      true,
      html:         sharedCSS() + html,
      sheetUrl:     spreadsheet.getUrl(),
      fromCache:    false,
      auditRef:     auditRef,
      observerName: observerName,
      warnings:     warnings,
      // Return resolved fields so client can backfill any blanks
      resolvedFields: {
        interactionId: interactionId,
        startTime:     startTime,
        duration:      duration,
        direction:     direction,
        participant:   participant,
        customerBAN:   customerBAN
      }
    };

  } catch (e) {
    Logger.log('submitTranscript error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Auto-send email on every submission — Team Member, QA, QA TL, Admin/Dev
// ─────────────────────────────────────────────────────────────────────────────
function sendSubmissionEmail(formData, htmlResult, auditRef) {
  try {
    var agentName     = formData.participant    || 'Agent';
    var sapId         = formData.sapId         || 'N/A';
    var interactionId = formData.interactionId || 'N/A';
    var observerName  = (formData.observerName || '').trim();
    var analysisLabel = formData.analysisType === 'sales'
                        ? 'Sales Analyzer' : 'Repeats & Transfer Analyzer';

    // ── 1. Team Member (agent being audited) ──────────────────────────────────
    var teamMemberEmail = lookupAgentEmail(agentName);

    // ── 2. QA / Observer ─────────────────────────────────────────────────────
    var qaEmail = '';
    if (observerName) {
      var qaRows      = getRecipientsFromRoster(QA_ROLE);
      var nameLower   = observerName.toLowerCase();
      var qaMatch     = qaRows.filter(function(r) { return r.name.toLowerCase() === nameLower; });
      qaEmail = qaMatch.length ? qaMatch[0].email : resolveEmail(observerName);
    }

    // ── 3. QA Team Leader(s) ─────────────────────────────────────────────────
    var qaTLEmails = getRecipientsFromRoster(QA_TL_ROLE).map(function(r) { return r.email; });

    // ── 4. Admin/Dev ──────────────────────────────────────────────────────────
    var adminEmails = getRecipientsFromRoster('Admin/Dev').map(function(r) { return r.email; });

    // ── Deduplicate and validate ───────────────────────────────────────────────
    var allEmails = [teamMemberEmail, qaEmail].concat(qaTLEmails).concat(adminEmails);
    var validRe   = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    var seen      = {};
    var recipients = allEmails.filter(function(e) {
      if (!e || !validRe.test(e) || seen[e.toLowerCase()]) return false;
      seen[e.toLowerCase()] = true;
      return true;
    });

    if (!recipients.length) {
      Logger.log('sendSubmissionEmail: no valid recipients — skipping');
      return;
    }

    Logger.log('sendSubmissionEmail: sending to ' + recipients.join(', '));

    var firstName  = agentName.split(' ')[0];
    var evalTitle  = formData.analysisType === 'sales' ? 'Sales Performance Evaluation' : 'New Hire Evaluation';
    var subject    = 'Real Time Feedback — ' + agentName + ' (' + sapId + ') | BAN: ' + (formData.customerBAN || 'N/A');

    var body =
      'Hi ' + firstName + ',\n\n' +
      'We\'re excited to share feedback from your recent customer interaction!\n\n' +
      'Your call has been reviewed to highlight your strengths and provide insights that will help you continue to grow and excel in your role.\n\n' +
      'Remember: This evaluation is a tool for your development, and we\'re here to support your success every step of the way.\n\n' +
      'EVALUATION DETAILS\n' +
      '─────────────────────────────────────\n' +
      'Audit Reference:  ' + auditRef + '\n' +
      'Interaction Date: ' + (formData.startTime || 'N/A') + '\n' +
      'Customer BAN:     ' + (formData.customerBAN || 'N/A') + '\n' +
      'Interaction ID:   ' + interactionId + '\n' +
      'Listening Type:   ' + (formData.direction || 'N/A') + '\n' +
      'Analysis Type:    ' + analysisLabel + '\n' +
      '─────────────────────────────────────\n\n' +
      'What\'s included in your evaluation:\n' +
      '  • Call summary and key points\n' +
      '  • Highlights of your performance\n' +
      '  • SMART coaching recommendations\n' +
      '  • Skill ratings and detailed feedback\n\n' +
      'Please review your evaluation and discuss with your Team Leader for coaching and development opportunities.\n\n' +
      'QA & LS Team\n' +
      'This evaluation is for development purposes.\n';

    var evalUrl  = ScriptApp.getService().getUrl() + '?page=eval&ref=' + encodeURIComponent(auditRef);
    var htmlBody = buildAgentEmailHTML(
      evalTitle, firstName, agentName, sapId,
      auditRef, interactionId,
      formData.startTime || 'N/A',
      formData.customerBAN || 'N/A',
      formData.direction || 'N/A',
      analysisLabel, evalUrl
    );

    MailApp.sendEmail({
      to:       recipients.join(','),
      subject:  subject,
      body:     body,
      htmlBody: htmlBody,
      name:     'NH Call Analyzer'
    });

    Logger.log('Submission email sent to: ' + recipients.join(', '));
    updateDashboardPDFLink(interactionId, 'Auto-sent', recipients, auditRef);
    try { notifyAdmins(formData, auditRef); } catch(ne) { Logger.log('notifyAdmins failed: ' + ne); }

  } catch(e) {
    Logger.log('sendSubmissionEmail error: ' + e.toString());
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Email audit — called from Submit button
// ─────────────────────────────────────────────────────────────────────────────
function sendAuditEmail(formData, htmlResult) {
  try {
    // ── Verify the audit actually exists before sending anything ──────────────
    var auditRefCheck = (formData.auditRef || '').trim();
    if (!auditRefCheck) {
      return { success: false, error: 'Missing audit reference.' };
    }
    var auditLog = readAuditLog();
    var auditRow = null;
    for (var ai = 0; ai < auditLog.length; ai++) {
      if ((auditLog[ai]['Audit Ref'] || '').toString().trim() === auditRefCheck) {
        auditRow = auditLog[ai];
        break;
      }
    }
    if (!auditRow) {
      return { success: false, error: 'Audit reference not found.' };
    }

    // ── Re-check admin status server-side — never trust the client's emailMode claim ──
    var callerEmail = '';
    try { callerEmail = Session.getActiveUser().getEmail() || ''; } catch(ee) {}
    callerEmail = callerEmail.toLowerCase().trim();
    var callerLocal  = callerEmail.split('@')[0];
    var callerDomain = callerEmail.split('@')[1] || '';
    var callerIsAdmin = !!callerLocal &&
                        ADMIN_USERNAMES.indexOf(callerLocal) !== -1 &&
                        callerDomain === ADMIN_DOMAIN.toLowerCase();
    var effectiveMode = (formData.emailMode === 'test' && callerIsAdmin) ? 'test' : 'live';

    var recipients;
    if (effectiveMode === 'test') {
      var adminList = getRecipientsFromRoster('Admin/Dev');
      recipients = adminList.map(function(r) { return r.email; }).filter(Boolean);
    } else {
      var teamLeaderEmail = resolveEmail(formData.teamLeader);
      var opsMgrEmail     = resolveEmail(formData.opsManager);
      recipients = [teamLeaderEmail, opsMgrEmail].filter(Boolean);
    }

    if (!recipients.length) {
      return { success: false, error: 'Could not resolve recipient email addresses.' };
    }

    var agentName     = formData.participant   || 'Agent';
    var sapId         = formData.sapId         || 'N/A';
    var interactionId = formData.interactionId || 'N/A';
    var auditRef      = formData.auditRef      || 'N/A';
    var analysisLabel = formData.analysisType === 'sales' ? 'Sales Analyzer' : 'Repeats & Transfer Analyzer';
    var firstName     = agentName.split(' ')[0];
    var evalTitle     = formData.analysisType === 'sales' ? 'Sales Performance Evaluation' : 'New Hire Evaluation';
    var subject       = (effectiveMode === 'test' ? '[TEST] ' : '') +
                        'Real Time Feedback — ' + agentName + ' (' + sapId + ') | BAN: ' + (formData.customerBAN || 'N/A');

    // ── Plain-text body ───────────────────────────────────────────────────────
    var body =
      'Hi ' + firstName + ',\n\n' +
      'We\'re excited to share feedback from your recent customer interaction!\n\n' +
      'Your call has been reviewed to highlight your strengths and provide insights that will help you continue to grow and excel in your role.\n\n' +
      'EVALUATION DETAILS\n' +
      '─────────────────────────────────────\n' +
      'Audit Reference:  ' + auditRef + '\n' +
      'Interaction Date: ' + (formData.startTime || 'N/A') + '\n' +
      'Customer BAN:     ' + (formData.customerBAN || 'N/A') + '\n' +
      'Interaction ID:   ' + interactionId + '\n' +
      'Listening Type:   ' + (formData.direction || 'N/A') + '\n' +
      'Analysis Type:    ' + analysisLabel + '\n' +
      '─────────────────────────────────────\n\n' +
      'What\'s included in your evaluation:\n' +
      '  • Call summary and key points\n' +
      '  • Highlights of your performance\n' +
      '  • SMART coaching recommendations\n' +
      '  • Skill ratings and detailed feedback\n\n' +
      'Please review your evaluation and discuss with your Team Leader for coaching and development opportunities.\n\n' +
      'QA & LS Team\n' +
      'This evaluation is for development purposes.\n';

    // ── HTML email body ───────────────────────────────────────────────────────
    var evalUrl  = ScriptApp.getService().getUrl() + '?page=eval&ref=' + encodeURIComponent(auditRef);
    var htmlBody = buildAgentEmailHTML(
      evalTitle, firstName, agentName, sapId,
      auditRef, interactionId,
      formData.startTime || 'N/A',
      formData.customerBAN || 'N/A',
      formData.direction || 'N/A',
      analysisLabel, evalUrl
    );

    // ── Send email ────────────────────────────────────────────────────────────
    MailApp.sendEmail({
      to:      recipients.join(','),
      subject: subject,
      body:    body,
      htmlBody: htmlBody,
      name:        'NH Call Analyzer'
    });

    // ── Update sheets ─────────────────────────────────────────────────────────
    updateDashboardPDFLink(interactionId, 'Sent via email', recipients, auditRef);

    Logger.log('Audit email sent to: ' + recipients.join(', '));
    try { notifyAdmins(formData, auditRefCheck); } catch(ne) { Logger.log('notifyAdmins failed: ' + ne); }

    // Update Cache Sheet with user-edited HTML so EvalView shows the edited version
    try {
      var intId = (formData.interactionId || '').trim();
      var aType = (formData.analysisType  || 'repeats').trim();
      if (intId && htmlResult) {
        updateCachedResult(intId, aType, htmlResult);
        // Clear the EvalView CacheService entry so the next visit re-reads the updated sheet
        var evKey = 'ev2_' + Utilities.base64Encode(auditRefCheck).substring(0, 200);
        CacheService.getScriptCache().remove(evKey);
      }
    } catch(ue) { Logger.log('updateCachedResult in sendAuditEmail failed: ' + ue); }

    return { success: true, recipients: recipients, auditRef: auditRef };

  } catch(e) {
    Logger.log('sendAuditEmail error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ── Resolve email from name — checks Primary Roster then Global Roster ────────
function resolveEmail(name) {
  if (!name) return '';
  try {
    var nameLower = name.trim().toLowerCase();

    // 1. Primary Roster — match on Agent_Name column only (col C, index 2)
    var ss    = SpreadsheetApp.openById(ROSTER_SHEET_ID);
    var sheet = ss.getSheetByName('roster');
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (data[i][2] && data[i][2].toString().trim().toLowerCase() === nameLower) {
          for (var k = 0; k < data[0].length; k++) {
            var h = data[0][k].toString().toLowerCase();
            if (h.indexOf('email') !== -1 || h.indexOf('mail') !== -1) {
              if (data[i][k]) return data[i][k].toString().trim();
            }
          }
        }
      }
    }

    // 2. Global Roster — covers TL/OM who are also tracked as agents there
    var gData = _getGlobalRosterData();
    var gHdr  = gData.header;
    var gRows = gData.rows;
    var gNameCol = -1, gEmailCol = -1;
    gHdr.forEach(function(h, idx) {
      var hl = h.toString().toLowerCase().trim();
      if (hl === 'member full name')                gNameCol  = idx;
      if (hl === 'email address' || hl === 'email') gEmailCol = idx;
    });
    if (gNameCol !== -1 && gEmailCol !== -1) {
      for (var g = 0; g < gRows.length; g++) {
        if (gRows[g][gNameCol] && gRows[g][gNameCol].toString().trim().toLowerCase() === nameLower) {
          var gEmail = gRows[g][gEmailCol] ? gRows[g][gEmailCol].toString().trim() : '';
          if (gEmail) return gEmail;
        }
      }
    }

    return '';
  } catch(e) {
    Logger.log('resolveEmail error: ' + e);
    return '';
  }
}

// ── Build the structured evaluation form HTML (matches screenshot layout) ─────
function buildEvalFormHTML(formData, auditRef, analysisResult) {
  var f          = formData || {};
  var agentName  = f.participant    || 'N/A';
  var sapId      = f.sapId          || 'N/A';
  var teamLeader = f.teamLeader     || 'N/A';
  var opsManager = f.opsManager     || 'N/A';
  var lob        = f.lineOfBusiness || 'N/A';
  var locale     = f.locale         || 'N/A';
  var vtid       = f.vtid           || 'N/A';
  var observer   = f.observerName   || 'N/A';
  var intId      = f.interactionId  || 'N/A';
  var direction  = f.direction      || 'N/A';
  var duration   = f.duration       || 'N/A';
  var startTime  = f.startTime      || 'N/A';
  var customerBAN= f.customerBAN    || 'N/A';
  var atype      = f.analysisType === 'sales' ? 'Sales Analyzer — Agent' : 'Repeats & Transfer Analyzer — Agent';
  var genDate    = new Date().toLocaleDateString('en-US', { year:'numeric', month:'long', day:'numeric' });
  var ref        = auditRef || 'N/A';

  // Extract key insights from AI analysis plain text
  var plainAI    = typeof analysisResult === 'string' ? htmlToPlainText(analysisResult) : '';
  var callSummary = '';
  var highlights  = '';
  var lowlights   = '';
  var coachingPts = '';

  // Extract call summary (first 300 chars after "Summary:")
  var sumMatch = plainAI.match(/Summary[:\s]+([^\n]{10,})/i);
  if (sumMatch) callSummary = sumMatch[1].substring(0, 400);

  // Extract highlights
  var hlMatch = plainAI.match(/Highlight[s]?[:\s]+([\s\S]{20,200}?)(?=Lowlight|Coaching|Recommendation|\n\n)/i);
  if (hlMatch) highlights = hlMatch[1].trim();

  // Extract lowlights
  var llMatch = plainAI.match(/Lowlight[s]?[^:]*[:\s]+([\s\S]{20,200}?)(?=Coaching|Recommendation|\n\n)/i);
  if (llMatch) lowlights = llMatch[1].trim();

  // Extract coaching points
  var coachMatch = plainAI.match(/Coaching[^:]*[:\s]+([\s\S]{20,300}?)(?=\n\n|$)/i);
  if (coachMatch) coachingPts = coachMatch[1].trim();

  var e = function(s){ return (s||'').toString().replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); };

  return '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
  '<style>' +
  /* ── Page: force narrow margins so Google Docs PDF uses full width ── */
  '@page{margin:7mm 6mm!important;size:A4}' +
  /* ── Body ── */
  'html,body{margin:0!important;padding:2px 3px!important;background:#fff;width:100%!important;' +
    'font-family:"Helvetica Neue",Helvetica,Arial,sans-serif;font-size:10.5px;color:#222}' +
  /* ── Header ── */
  '.hdr{display:flex;justify-content:space-between;align-items:flex-start;' +
    'margin-bottom:10px;border-bottom:3px solid #4B286D;padding-bottom:8px}' +
  '.hdr-left h1{font-size:15px;font-weight:800;color:#4B286D;margin:0 0 2px}' +
  '.hdr-left p{font-size:9.5px;color:#555;margin:0}' +
  '.hdr-logo{font-size:9.5px;font-weight:700;color:#4B286D;text-align:right;border:2px solid #4B286D;padding:4px 8px;border-radius:4px}' +
  '.ref-bar{background:#4B286D;color:#fff;padding:4px 10px;border-radius:3px;font-size:9.5px;margin-bottom:9px;display:flex;justify-content:space-between}' +
  /* ── Sections ── */
  '.section{margin-bottom:9px}' +
  '.section-title{font-size:10px;font-weight:700;color:#4B286D;border-bottom:2px solid #4B286D;padding-bottom:2px;margin-bottom:6px}' +
  /* ── Detail grids ── */
  '.detail-grid{display:table;width:100%;border-collapse:collapse}' +
  '.detail-row{display:table-row}' +
  '.detail-label{display:table-cell;font-weight:700;color:#555;padding:2px 10px 2px 0;width:110px;font-size:8.5px;text-transform:uppercase;letter-spacing:.3px;vertical-align:top;white-space:nowrap}' +
  '.detail-val{display:table-cell;padding:2px 0;font-size:10px;vertical-align:top;border-bottom:1px solid #F0F0F0;word-break:break-word}' +
  '.two-col{display:table;width:100%;table-layout:fixed}' +
  '.col-left{display:table-cell;width:50%;padding-right:6px;vertical-align:top}' +
  '.col-right{display:table-cell;width:50%;padding-left:6px;vertical-align:top}' +
  /* ── Remarks boxes ── */
  '.remarks-box{border:1px solid #D8D8D8;border-radius:3px;padding:6px 8px;background:#FAFAFA;font-size:10px;line-height:1.55;white-space:pre-wrap;min-height:50px;word-break:break-word}' +
  '.remarks-label{font-size:8.5px;font-weight:700;color:#4B286D;text-transform:uppercase;margin-bottom:4px;letter-spacing:.3px}' +
  '.highlight{color:#2B8000;font-weight:600}' +
  '.lowlight{color:#C12335;font-weight:600}' +
  '.footer{margin-top:12px;border-top:1px solid #D8D8D8;padding-top:5px;font-size:8px;color:#888;text-align:center}' +
  '.badge{display:inline-block;padding:1px 6px;border-radius:8px;font-size:9px;font-weight:700;background:#4B286D;color:#fff}' +
  /* ── AI content — full-width, tight tables ── */
  '.ai-section{margin-bottom:8px!important;padding:0!important}' +
  '.ai-title{font-size:10px!important;padding:4px 8px!important;margin-bottom:5px!important}' +
  /* all tables: full width, fixed layout, small font, tight cells */
  'table{width:100%!important;table-layout:fixed!important;border-collapse:collapse!important;font-size:8.5px!important}' +
  'th,td{padding:2px 4px!important;word-break:break-word!important;overflow-wrap:break-word!important;vertical-align:top!important;font-size:8.5px!important}' +
  'th{white-space:normal!important;font-size:8px!important}' +
  '.ai-table{width:100%!important;table-layout:fixed!important;font-size:8.5px!important;border-collapse:collapse!important}' +
  '.ai-table th,.ai-table td{padding:2px 4px!important;word-break:break-word!important;vertical-align:top!important;font-size:8.5px!important}' +
  '.ai-table th{font-size:8px!important}' +
  '.ai-label-col{width:12%!important;font-size:8px!important;white-space:normal!important}' +
  '.ai-flag{padding:5px 8px!important;margin-bottom:5px!important}' +
  '.ai-flag-title{font-size:9.5px!important}' +
  '.ai-flag-detail{font-size:8.5px!important}' +
  '.ai-flag-stmt{font-size:8.5px!important;padding:4px 6px!important}' +
  '.ai-summary,.ai-perfect{font-size:9.5px!important;line-height:1.45!important}' +
  '.ai-hl-grid{display:table!important;width:100%!important}' +
  '.ai-hl-box{display:table-cell!important;width:50%!important;padding:5px 8px!important}' +
  '.ai-hl-box:first-child{padding-right:4px!important}' +
  '.ai-hl-box:last-child{padding-left:4px!important}' +
  '.ai-hl-box ul{font-size:8.5px!important;line-height:1.45!important}' +
  '.ai-coaching{font-size:8.5px!important;line-height:1.5!important}' +
  '.ai-score-panel{padding:5px 8px!important}' +
  '.ai-score-row{font-size:9.5px!important;padding:3px 0!important}' +
  '.ai-badge{font-size:8.5px!important;padding:1px 5px!important;min-width:36px!important}' +
  '.ai-info{font-size:8.5px!important;gap:3px!important;flex-wrap:wrap!important}' +
  '.ai-chip{font-size:8px!important;padding:1px 5px!important}' +
  '.ai-call-badge{font-size:10px!important;padding:3px 10px!important;margin-bottom:6px!important}' +
  /* ── SMART tables (amber header) — even tighter ── */
  'table[style*="FDE8B0"] th,table[style*="FDE8B0"] td,' +
  'table[style*="F5F0FF"] th,table[style*="F5F0FF"] td{font-size:8px!important;padding:2px 3px!important}' +
  '@media print{html,body{margin:0!important;padding:1px 2px!important}}' +
  '</style></head><body>' +

  // Header
  '<div class="hdr">' +
    '<div class="hdr-left">' +
      '<h1>FCR Live Call Listening Evaluation</h1>' +
      '<p>Quality Analyst Team &nbsp;·&nbsp; NH FCR, Transfer or Sales Call Analyzer</p>' +
    '</div>' +
    '<div class="hdr-logo">GLE+QA<br>TEAM</div>' +
  '</div>' +

  // Audit ref bar
  '<div class="ref-bar">' +
    '<span>Audit Reference: <strong>' + e(ref) + '</strong></span>' +
    '<span>Analysis Type: <strong>' + e(atype) + '</strong></span>' +
    '<span>Generated: <strong>' + e(genDate) + '</strong></span>' +
  '</div>' +

  // Evaluation Details
  '<div class="section">' +
    '<div class="section-title">&#128196; Evaluation Details</div>' +
    '<div class="two-col">' +
      '<div class="col-left">' +
        '<div class="detail-grid">' +
          '<div class="detail-row"><span class="detail-label">Team Member</span><span class="detail-val">' + e(agentName) + '</span></div>' +
          '<div class="detail-row"><span class="detail-label">SAP ID</span><span class="detail-val">' + e(sapId) + '</span></div>' +
          '<div class="detail-row"><span class="detail-label">VTID</span><span class="detail-val">' + e(vtid) + '</span></div>' +
          '<div class="detail-row"><span class="detail-label">Team Leader</span><span class="detail-val">' + e(teamLeader) + '</span></div>' +
          '<div class="detail-row"><span class="detail-label">Operations Manager</span><span class="detail-val">' + e(opsManager) + '</span></div>' +
          '<div class="detail-row"><span class="detail-label">Line of Business</span><span class="detail-val">' + e(lob) + '</span></div>' +
          '<div class="detail-row"><span class="detail-label">Locale / Site</span><span class="detail-val">' + e(locale) + '</span></div>' +
          '<div class="detail-row"><span class="detail-label">Observer</span><span class="detail-val">' + e(observer) + '</span></div>' +
        '</div>' +
      '</div>' +
      '<div class="col-right">' +
        '<div class="detail-grid">' +
          '<div class="detail-row"><span class="detail-label">BAN / Customer</span><span class="detail-val">' + e(customerBAN) + '</span></div>' +
          '<div class="detail-row"><span class="detail-label">Interaction Date</span><span class="detail-val">' + e(startTime) + '</span></div>' +
          '<div class="detail-row"><span class="detail-label">Interaction ID</span><span class="detail-val" style="font-size:9px">' + e(intId) + '</span></div>' +
          '<div class="detail-row"><span class="detail-label">Direction</span><span class="detail-val">' + e(direction) + '</span></div>' +
          '<div class="detail-row"><span class="detail-label">Call Duration</span><span class="detail-val">' + e(duration) + '</span></div>' +
        '</div>' +
      '</div>' +
    '</div>' +
  '</div>' +

  // Call Summary
  (callSummary ? (
    '<div class="section">' +
      '<div class="section-title">&#128222; Call Summary</div>' +
      '<div class="remarks-box">' + e(callSummary) + '</div>' +
    '</div>'
  ) : '') +

  // QA Remarks & Feedback
  '<div class="section">' +
    '<div class="section-title">&#128172; QA Remarks &amp; Feedback</div>' +
    '<div class="two-col">' +
      '<div class="col-left">' +
        '<div class="remarks-label highlight">&#10003; Highlights</div>' +
        '<div class="remarks-box" style="border-color:#B3DFA0;background:#F0FFF4">' + e(highlights || 'See full analysis for highlights.') + '</div>' +
      '</div>' +
      '<div class="col-right">' +
        '<div class="remarks-label lowlight">&#10007; Lowlights / Recommendations</div>' +
        '<div class="remarks-box" style="border-color:#F5AAAA;background:#FFF5F5">' + e(lowlights || 'See full analysis for recommendations.') + '</div>' +
      '</div>' +
    '</div>' +
    (coachingPts ? (
      '<div style="margin-top:10px">' +
        '<div class="remarks-label" style="color:#8C4A00">&#127979; Coaching Takeaways</div>' +
        '<div class="remarks-box" style="border-color:#FFE082;background:#FFFDE7">' + e(coachingPts) + '</div>' +
      '</div>'
    ) : '') +
  '</div>' +

  // Full AI Analysis
  '<div class="section">' +
    '<div class="section-title">&#128203; Full AI Analysis</div>' +
    sanitiseHTMLForPDF(typeof analysisResult === 'string' ? analysisResult : '') +
  '</div>' +

  '<div class="footer">Audit Ref: ' + e(ref) + ' &nbsp;·&nbsp; Generated by NH FCR, Transfer or Sales Call Analyzer &nbsp;·&nbsp; Quality Analyst Team</div>' +
  '</body></html>';
}

// ── Generate PDF blob from structured form HTML ───────────────────────────────
function generateAuditPDF(formData, htmlResult) {
  try {
    var agentName     = formData.participant    || 'Agent';
    var interactionId = formData.interactionId  || 'N/A';
    var auditRef      = formData.auditRef       || 'NHA';
    var fullHtml = buildEvalFormHTML(formData, auditRef, htmlResult);

    var blob      = Utilities.newBlob(fullHtml, 'text/html', 'audit.html');
    var driveFile = Drive.Files.insert(
      { title: 'audit_tmp', mimeType: 'application/vnd.google-apps.document' },
      blob
    );
    var pdf = null;
    try {
      pdf = DriveApp.getFileById(driveFile.id).getAs('application/pdf');
      pdf.setName('Call_Audit_' + agentName.replace(/\s+/g,'_') + '_' + interactionId + '.pdf');
    } finally {
      // Always trash the temp doc even if PDF conversion throws
      try { DriveApp.getFileById(driveFile.id).setTrashed(true); } catch(te) {}
    }
    return pdf;
  } catch(e) {
    Logger.log('generateAuditPDF error: ' + e);
    return null;
  }
}

// ── Strip dangerous tags from AI HTML before embedding in PDF ─────────────────
function sanitiseHTMLForPDF(html) {
  if (!html) return '';
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<iframe[\s\S]*?>/gi, '')
    .replace(/<object[\s\S]*?<\/object>/gi, '')
    .replace(/<form[\s\S]*?<\/form>/gi, '')
    .replace(/<meta[^>]*http-equiv[^>]*>/gi, '')
    .replace(/\s+on\w+="[^"]*"/gi, '')   // remove inline event handlers
    .replace(/\s+on\w+='[^']*'/gi, '');
}

// ── Email helper ───────────────────────────────────────────────────────────────
function emailRow(label, value) {
  return '<tr>' +
    '<td style="padding:7px 12px;border-bottom:1px solid #eee;font-weight:700;' +
    'color:#54565A;font-size:12px;width:35%;background:#FAFAFA">' + escEmail(label) + '</td>' +
    '<td style="padding:7px 12px;border-bottom:1px solid #eee;font-size:12px">' + escEmail(value) + '</td>' +
    '</tr>';
}

function escEmail(s) {
  if (!s) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// ─────────────────────────────────────────────────────────────────────────────
// Agent-facing evaluation email — friendly intro matching the New Hire design
// ─────────────────────────────────────────────────────────────────────────────
function buildAgentEmailHTML(evalTitle, firstName, agentName, sapId,
                              auditRef, interactionId,
                              interactionDate, customerBAN,
                              listeningType, analysisLabel, evalUrl) {
  var e = escEmail;
  var isSales = (analysisLabel || '').indexOf('Sales') !== -1;

  // Detail row helper (label + value, lavender card style)
  function detailRow(label, value) {
    return '<tr>' +
      '<td style="padding:10px 14px 2px;font-size:10px;font-weight:700;' +
           'text-transform:uppercase;letter-spacing:.8px;color:#6A3D99">' + e(label) + '</td>' +
      '</tr><tr>' +
      '<td style="padding:0 14px 10px;font-size:14px;color:#1A1A2E;' +
           'border-bottom:1px solid #E8DEF5">' + e(value) + '</td>' +
      '</tr>';
  }

  var bullets = isSales
    ? ['Call summary and key observations',
       'Sales highlights and missed opportunities',
       'SMART coaching recommendations (S·M·A·R·T)',
       'Skill ratings across 5 sales dimensions',
       'Sample positioning statements for practice']
    : ['Call summary and key points',
       'Highlights of your performance',
       'SMART coaching recommendations (S·M·A·R·T)',
       'Repeat risk analysis and FCR assessment',
       'Critical flags and positioning statements'];

  var bulletHTML = bullets.map(function(b) {
    return '<li style="margin-bottom:7px">' + e(b) + '</li>';
  }).join('');

  return (
    '<div style="font-family:Helvetica Neue,Helvetica,Arial,sans-serif;' +
         'max-width:600px;margin:0 auto;background:#F4F4F8;padding:24px 16px">' +

    // ── Header ───────────────────────────────────────────────────────────────
    '<div style="background:#4B286D;padding:28px 32px;border-radius:10px 10px 0 0;text-align:center">' +
      '<h1 style="color:#fff;margin:0 0 6px;font-size:22px;font-weight:700;letter-spacing:-.3px">' +
        e(evalTitle) +
      '</h1>' +
      '<p style="color:rgba(255,255,255,.75);margin:0;font-size:13px;font-weight:500">' +
        'QA &amp; LS Team' +
      '</p>' +
    '</div>' +

    // ── Body card ─────────────────────────────────────────────────────────────
    '<div style="background:#fff;border-radius:0 0 10px 10px;padding:30px 32px">' +

      // Greeting
      '<p style="font-size:15px;font-weight:600;margin:0 0 14px;color:#1A1A2E">' +
        'Hi ' + e(firstName) + ',' +
      '</p>' +
      '<p style="font-size:13px;line-height:1.7;margin:0 0 10px;color:#3A3A4E">' +
        'We\'re excited to share feedback from your recent customer interaction!' +
      '</p>' +
      '<p style="font-size:13px;line-height:1.7;margin:0 0 10px;color:#3A3A4E">' +
        'Your call has been reviewed to highlight your strengths and provide insights ' +
        'that will help you continue to grow and excel in your role.' +
      '</p>' +
      '<p style="font-size:13px;line-height:1.7;margin:0 0 24px;color:#3A3A4E">' +
        '<strong>Remember:</strong> This evaluation is a tool for your development, and we\'re here ' +
        'to support your success every step of the way.' +
      '</p>' +

      // Evaluation Details card
      '<div style="background:#F5F0FF;border:1px solid #D5BFF0;border-left:4px solid #4B286D;' +
           'border-radius:8px;margin-bottom:24px;overflow:hidden">' +
        '<div style="padding:10px 14px 6px;font-size:11px;font-weight:700;' +
             'text-transform:uppercase;letter-spacing:1px;color:#4B286D;' +
             'border-bottom:1px solid #E0D0F7;display:flex;align-items:center;gap:6px">' +
          '&#128203; EVALUATION DETAILS' +
        '</div>' +
        '<table style="width:100%;border-collapse:collapse">' +
          detailRow('INTERACTION DATE', interactionDate) +
          detailRow('CUSTOMER BAN', customerBAN) +
          detailRow('FLP CONVERSATION ID', interactionId) +
          detailRow('LISTENING TYPE', listeningType) +
          detailRow('ANALYSIS TYPE', analysisLabel) +
          detailRow('AUDIT REFERENCE', auditRef) +
        '</table>' +
      '</div>' +

      // CTA Button
      '<div style="text-align:center;margin-bottom:28px">' +
        '<a href="' + (evalUrl || '#') + '" ' +
           'style="display:inline-block;background:#4B286D;color:#fff;text-decoration:none;' +
                  'border-radius:6px;padding:13px 32px;font-size:14px;font-weight:700;' +
                  'letter-spacing:.3px;text-align:center">' +
          'VIEW YOUR EVALUATION' +
        '</a>' +
        '<p style="font-size:11px;color:#888;margin:8px 0 0">' +
          'Click the button above to open your full evaluation. You must be signed in with your work account.' +
        '</p>' +
      '</div>' +

      // What's included
      '<p style="font-size:13px;font-weight:600;color:#1A1A2E;margin:0 0 10px">' +
        'What\'s included in your evaluation:' +
      '</p>' +
      '<ul style="padding-left:20px;font-size:13px;line-height:1.8;color:#3A3A4E;margin:0 0 24px">' +
        bulletHTML +
      '</ul>' +

      // Closing
      '<p style="font-size:13px;line-height:1.7;color:#3A3A4E;margin:0 0 0">' +
        'Please review your evaluation and discuss with your Team Leader for coaching ' +
        'and development opportunities. Remember, every call is a chance to grow — ' +
        'keep up the great work! &#127775;' +
      '</p>' +

    '</div>' + // end body card

    // ── Footer ───────────────────────────────────────────────────────────────
    '<div style="padding:18px 8px;text-align:center">' +
      '<p style="font-size:13px;font-weight:600;color:#4B286D;margin:0 0 4px">' +
        'QA &amp; LS Team' +
      '</p>' +
      '<p style="font-size:11px;color:#999;margin:0">' +
        'This evaluation is for development purposes.' +
      '</p>' +
    '</div>' +

    '</div>' // end outer wrapper
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// Admin notification email — TELUS-branded, matches New Hire Evaluation style
// ─────────────────────────────────────────────────────────────────────────────
function buildAdminEmailHTML(
  auditRef, submittedAt, observer, analysisType,
  agentName, sapId, vtid, teamLeader, opsManager, lob, locale,
  interactionId, direction, duration, evalUrl
) {
  var e = escEmail;

  function infoCard(label, value) {
    return '<tr>' +
      '<td style="padding:10px 0 2px;font-size:10px;font-weight:700;text-transform:uppercase;' +
           'letter-spacing:.8px;color:#6A3D99;border-bottom:none">' + e(label) + '</td>' +
    '</tr><tr>' +
      '<td style="padding:0 0 10px;font-size:13px;color:#1A1A2E;font-weight:500;' +
           'border-bottom:1px solid #EDE8F7">' + e(value || '—') + '</td>' +
    '</tr>';
  }

  function sectionHeader(icon, title) {
    return '<tr><td style="padding:16px 0 6px;font-size:10px;font-weight:700;' +
      'text-transform:uppercase;letter-spacing:1px;color:#4B286D;' +
      'border-bottom:2px solid #4B286D">' + icon + ' ' + e(title) + '</td></tr>';
  }

  return (
    '<div style="font-family:Helvetica Neue,Helvetica,Arial,sans-serif;' +
         'max-width:600px;margin:0 auto;background:#F0EBF8;padding:24px 16px">' +

    // ── Header ───────────────────────────────────────────────────────────────
    '<div style="background:#4B286D;padding:26px 30px;border-radius:10px 10px 0 0">' +
      '<div style="display:flex;align-items:flex-start;justify-content:space-between">' +
        '<div>' +
          '<div style="font-size:10px;font-weight:700;text-transform:uppercase;' +
               'letter-spacing:1.2px;color:rgba(255,255,255,.6);margin-bottom:6px">' +
            'ADMIN NOTIFICATION' +
          '</div>' +
          '<h1 style="color:#fff;margin:0 0 8px;font-size:20px;font-weight:700;letter-spacing:-.2px">' +
            '&#128196; New Audit Submitted' +
          '</h1>' +
          '<div style="background:rgba(255,255,255,.15);border-radius:20px;display:inline-block;' +
               'padding:4px 14px;font-size:12px;font-weight:700;color:#fff;letter-spacing:.2px">' +
            e(auditRef) +
          '</div>' +
        '</div>' +
        '<div style="background:#fff;border-radius:6px;padding:6px 12px;text-align:center;' +
             'font-size:10px;font-weight:700;color:#4B286D;letter-spacing:.5px;line-height:1.4;' +
             'flex-shrink:0;margin-left:16px">' +
          'GLE·QA<br>TEAM' +
        '</div>' +
      '</div>' +
      '<div style="margin-top:10px;font-size:11px;color:rgba(255,255,255,.65)">' +
        '&#128337; ' + e(submittedAt) +
      '</div>' +
    '</div>' +

    // ── Body card ─────────────────────────────────────────────────────────────
    '<div style="background:#fff;border-radius:0 0 10px 10px;padding:26px 30px">' +

      // Observer / Type badge
      '<div style="background:#F5F0FF;border-left:4px solid #4B286D;border-radius:6px;' +
           'padding:11px 16px;margin-bottom:22px;font-size:13px;display:flex;gap:20px">' +
        '<span><strong style="color:#4B286D">Observer:</strong> ' + e(observer) + '</span>' +
        '<span><strong style="color:#4B286D">Type:</strong> ' + e(analysisType) + '</span>' +
      '</div>' +

      // Agent Details
      '<table style="width:100%;border-collapse:collapse">' +
        sectionHeader('&#128100;', 'Agent Details') +
        // 2-col agent grid
        '<tr><td>' +
          '<table style="width:100%;border-collapse:collapse">' +
            '<tr>' +
              '<td style="width:50%;vertical-align:top"><table style="width:100%;border-collapse:collapse">' +
                infoCard('Agent Name',    agentName) +
                infoCard('SAP ID',        sapId) +
                infoCard('VTID',          vtid) +
                infoCard('Team Leader',   teamLeader) +
              '</table></td>' +
              '<td style="width:50%;vertical-align:top;padding-left:20px"><table style="width:100%;border-collapse:collapse">' +
                infoCard('Operations Mgr',   opsManager) +
                infoCard('Line of Business', lob) +
                infoCard('Locale / Site',    locale) +
              '</table></td>' +
            '</tr>' +
          '</table>' +
        '</td></tr>' +

        // Call Details
        sectionHeader('&#128222;', 'Call Details') +
        '<tr><td>' +
          '<table style="width:100%;border-collapse:collapse">' +
            '<tr>' +
              '<td style="width:50%;vertical-align:top"><table style="width:100%;border-collapse:collapse">' +
                infoCard('Interaction ID', interactionId) +
                infoCard('Direction',      direction) +
              '</table></td>' +
              '<td style="width:50%;vertical-align:top;padding-left:20px"><table style="width:100%;border-collapse:collapse">' +
                infoCard('Duration', duration) +
              '</table></td>' +
            '</tr>' +
          '</table>' +
        '</td></tr>' +
      '</table>' +

      // CTA Button
      '<div style="text-align:center;margin:24px 0 20px">' +
        '<a href="' + e(evalUrl) + '" ' +
           'style="display:inline-block;background:#4B286D;color:#fff;text-decoration:none;' +
                  'border-radius:6px;padding:12px 30px;font-size:14px;font-weight:700;letter-spacing:.3px">' +
          'VIEW FULL EVALUATION' +
        '</a>' +
      '</div>' +

      // Footer note
      '<div style="border-top:1px solid #EDE8F7;padding-top:14px;font-size:11px;color:#999;text-align:center">' +
        'You received this because your role is listed as <strong>Admin/Dev</strong> in the Roster.<br>' +
        'NH FCR, Transfer or Sales Call Analyzer — Real Time Feedback' +
      '</div>' +

    '</div>' + // end body card

    // Outer footer
    '<div style="padding:14px 8px;text-align:center">' +
      '<p style="font-size:12px;font-weight:600;color:#4B286D;margin:0 0 3px">QA &amp; LS Team</p>' +
      '<p style="font-size:11px;color:#999;margin:0">This is an automated admin notification.</p>' +
    '</div>' +

    '</div>' // end wrapper
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// Export rows to a new Google Spreadsheet and return its URL
// Called from dashboard Export All / Export Filtered buttons
// ─────────────────────────────────────────────────────────────────────────────
function exportToSpreadsheet(rows, title) {
  try {
    if (!rows || !rows.length) return { success: false, error: 'No records to export.' };

    var ss      = SpreadsheetApp.create(title + ' — ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
    var sheet   = ss.getActiveSheet();
    var headers = Object.keys(rows[0]);

    // Write header row
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold')
      .setBackground('#4B286D')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);

    // Write data rows
    var data = rows.map(function(row) {
      return headers.map(function(h) {
        var v = row[h];
        return (v !== null && v !== undefined) ? v.toString() : '';
      });
    });
    if (data.length) sheet.getRange(2, 1, data.length, headers.length).setValues(data);

    // Auto-resize all columns
    try { sheet.autoResizeColumns(1, headers.length); } catch(e) {}

    // Restrict to org domain only — never public
    var file = DriveApp.getFileById(ss.getId());
    try {
      file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
    } catch(se) {
      // Fallback if domain sharing not available (personal accounts)
      Logger.log('Domain sharing unavailable, export URL requires login: ' + se);
    }

    Logger.log('Export spreadsheet created: ' + ss.getUrl());
    return { success: true, url: ss.getUrl(), title: ss.getName(), rowCount: rows.length };
  } catch(e) {
    Logger.log('exportToSpreadsheet error: ' + e);
    return { success: false, error: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// serverParseJSON removed — project now uses AI-returned HTML directly.
// The function below is intentionally unreachable and kept only as a tombstone.
// ─────────────────────────────────────────────────────────────────────────────
function _removedServerParseJSON_doNotCall(text) {
  if (!text) return null;

  // Extract the raw JSON candidate: first { to last }
  var fi = text.indexOf('{');
  var li = text.lastIndexOf('}');
  if (fi === -1 || li <= fi) {
    Logger.log('serverParseJSON: no { } block found');
    return null;
  }
  var candidate = text.slice(fi, li + 1);

  // Attempt 1: parse as-is
  try {
    var obj = JSON.parse(candidate);
    if (obj && typeof obj === 'object') { Logger.log('serverParseJSON: parsed on attempt 1'); return obj; }
  } catch(e1) {
    Logger.log('serverParseJSON attempt 1 failed: ' + e1.message);
  }

  // Attempt 2: fix unescaped newlines/tabs inside JSON string values
  // Replace bare \n and \r inside quoted strings with \\n / \\r
  try {
    var fixed = candidate.replace(/"((?:[^"\\]|\\.)*)"/g, function(match) {
      return match
        .replace(/\n/g, '\\n')
        .replace(/\r/g, '\\r')
        .replace(/\t/g, '\\t');
    });
    var obj2 = JSON.parse(fixed);
    if (obj2 && typeof obj2 === 'object') { Logger.log('serverParseJSON: parsed on attempt 2 (newline fix)'); return obj2; }
  } catch(e2) {
    Logger.log('serverParseJSON attempt 2 failed: ' + e2.message);
  }

  // Attempt 3: strip any markdown fences then retry
  try {
    var stripped = candidate.replace(/```(?:json)?/g,'').replace(/```/g,'').trim();
    var obj3 = JSON.parse(stripped);
    if (obj3 && typeof obj3 === 'object') { Logger.log('serverParseJSON: parsed on attempt 3 (stripped fences)'); return obj3; }
  } catch(e3) {
    Logger.log('serverParseJSON attempt 3 failed: ' + e3.message);
  }

  // Attempt 4: aggressive — replace ALL literal newlines in the full block then re-parse
  try {
    var aggressive = candidate
      .replace(/[\r\n]+/g, ' ')   // collapse all newlines to space
      .replace(/\t/g, ' ');       // collapse tabs
    var obj4 = JSON.parse(aggressive);
    if (obj4 && typeof obj4 === 'object') { Logger.log('serverParseJSON: parsed on attempt 4 (aggressive collapse)'); return obj4; }
  } catch(e4) {
    Logger.log('serverParseJSON attempt 4 failed: ' + e4.message);
  }

  Logger.log('serverParseJSON: ALL attempts failed');
  return null;
}

// ─────────────────────────────────────────────────────────────────────────────
// DEAD CODE — HTML builders below are no longer called.
// The project switched to AI returning HTML directly (prompt-based approach).
// Kept commented out to avoid breaking any lingering references during cleanup.
// Safe to delete entirely in a future cleanup pass.
// ─────────────────────────────────────────────────────────────────────────────
function esc(s) {
  if (s === null || s === undefined) return '';
  return String(s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function buildResultHTML(d, analysisType) {
  return analysisType === 'sales'
    ? buildSalesHTML(d)
    : buildRepeatsHTML(d);
}

// ── Shared section wrapper ────────────────────────────────────────────────────
function section(title, icon, body) {
  return '<div class="r-section">' +
    '<div class="r-section-title">' + (icon ? '<span>' + icon + '</span> ' : '') + esc(title) + '</div>' +
    body +
    '</div>';
}

// ── Editable cell helper ──────────────────────────────────────────────────────
function cell(text) {
  return '<td contenteditable="true">' + esc(text || '') + '</td>';
}
function cellHtml(html) {
  return '<td contenteditable="true">' + (html || '') + '</td>';
}

// ── Info bar (key-value chips) ────────────────────────────────────────────────
function infoBar(pairs) {
  var html = '<div class="r-info-bar">';
  pairs.forEach(function(p) {
    if (p[1]) html += '<div class="r-info-chip"><span class="r-info-label">' + esc(p[0]) + '</span><span class="r-info-val">' + esc(p[1]) + '</span></div>';
  });
  html += '</div>';
  return html;
}

// ── Analysis table ────────────────────────────────────────────────────────────
function analysisTable(headers, rows) {
  var html = '<table class="r-table"><thead><tr>';
  headers.forEach(function(h) { html += '<th>' + esc(h) + '</th>'; });
  html += '</tr></thead><tbody>';
  rows.forEach(function(r) {
    html += '<tr>';
    r.forEach(function(c, i) {
      html += i === 0
        ? '<td class="r-table-label">' + esc(c) + '</td>'
        : cell(c);
    });
    html += '</tr>';
  });
  html += '</tbody></table>';
  return html;
}

// ── Positioning statement card ────────────────────────────────────────────────
function posCard(flag) {
  var param = typeof flag === 'string' ? flag : (flag.parameter || '');
  var expl  = typeof flag === 'object' ? (flag.explanation || '') : '';
  var stmt  = typeof flag === 'object' ? (flag.positioningStatement || flag.statement || '') : '';
  var html  = '<div class="r-pos-card">';
  html += '<div class="r-pos-param">&#9888; ' + esc(param) + '</div>';
  if (expl) html += '<div class="r-pos-detail">' + esc(expl) + '</div>';
  if (stmt) html += '<div class="r-pos-stmt-label">&#127908; Sample Positioning Statement — Roleplay &amp; Practice</div>' +
                    '<div class="r-pos-stmt" contenteditable="true">"' + esc(stmt) + '"</div>';
  html += '</div>';
  return html;
}

// ── Highlights / Lowlights ────────────────────────────────────────────────────
function highlightBox(highs, lows) {
  var html = '<div class="r-hl-grid">';
  html += '<div class="r-hl-box r-hl-high"><div class="r-hl-title">&#10003; Highlights</div><ul>';
  (highs || []).forEach(function(h) { html += '<li contenteditable="true">' + esc(h) + '</li>'; });
  html += '</ul></div>';
  html += '<div class="r-hl-box r-hl-low"><div class="r-hl-title">&#10007; Lowlights / Recommendations</div><ul>';
  (lows || []).forEach(function(l) { html += '<li contenteditable="true">' + esc(l) + '</li>'; });
  html += '</ul></div>';
  html += '</div>';
  return html;
}

// ── Coaching takeaways ────────────────────────────────────────────────────────
function coachingList(items) {
  if (!items || !items.length) return '';
  var html = '<ol class="r-coaching-list">';
  (Array.isArray(items) ? items : [items]).forEach(function(item) {
    html += '<li contenteditable="true">' + esc(item) + '</li>';
  });
  html += '</ol>';
  return html;
}

// ── Score badge ───────────────────────────────────────────────────────────────
// 3-5 = green, 2-2.9 = amber, 1-1.9 = red, 0 = grey
function scoreBadge(score) {
  var s   = parseFloat(score) || 0;
  var cls = s === 0        ? 'r-badge-zero'
          : s >= 3         ? 'r-badge-good'
          : s >= 2         ? 'r-badge-mid'
          :                  'r-badge-bad';
  return '<span class="r-badge ' + cls + '">' + score + '/5</span>';
}

// ─────────────────────────────────────────────────────────────────────────────
// REPEATS HTML
// ─────────────────────────────────────────────────────────────────────────────
function buildRepeatsHTML(d) {
  var calls = d.calls || [];
  var html  = '';

  // Per-call sections
  calls.forEach(function(c, idx) {
    html += '<div class="r-call-block">';
    html += '<div class="r-call-badge">Call ' + (idx + 1) + '</div>';

    // ── Call Summary ──
    html += section('Call Summary', '&#128222;',
      '<div class="r-summary-box" contenteditable="true">' + esc(c.callSummary || '') + '</div>'
    );

    // ── Call Information bar ──
    html += section('Call Information', '&#128100;',
      infoBar([
        ['Agent',       c.agentName],
        ['Date',        c.interactionDate],
        ['Phone',       c.phoneNumber],
        ['Country',     c.agentCountry],
        ['Department',  c.department],
        ['Call Reason', c.callReason],
        ['Issue Resolved', c.issueResolved],
        ['Repeat Risk', c.repeatProjectionPct !== undefined ? c.repeatProjectionPct + '%' : '']
      ])
    );

    // ── Analysis Table ──
    var drivers = Array.isArray(c.repeatProjectionDrivers)
      ? c.repeatProjectionDrivers.join('; ')
      : (c.repeatProjectionDrivers || '');

    html += section('Analysis — Opportunities & Recommendations', '&#128203;',
      analysisTable(
        ['Parameter', 'Finding / Detail', 'Recommendation'],
        [
          ['Repeat Projection',        c.repeatProjectionPct !== undefined ? c.repeatProjectionPct + '%' : '',  c.repeatProjectionDetail || ''],
          ['Repeat Drivers',           drivers,                                                                  c.coachingOpportunities || ''],
          ['Issue Resolution',         c.issueResolved || '',                                                    c.unresolvedDetails || ''],
          ['Process / Policy Gaps',    c.processGaps || '',                                                      c.missingSteps || ''],
          ['Callback Policy',          c.callbackPolicyFollowed || '',                                           ''],
          ['Transfer Occurred',        c.transferOccurred || '',                                                 c.transferValid || ''],
          ['Transfer Validity',        c.transferValid || '',                                                    c.transferBehaviors || ''],
          ['FCR Assessment',           '',                                                                       c.fcrAssessment || ''],
          ['Agent Strengths',          '',                                                                       c.agentStrengths || ''],
          ['Performance Improvement',  '',                                                                       c.performanceImprovements || ''],
          ['3 Actions to Resolve',     '',
            Array.isArray(c.threeActionsToResolve)
              ? c.threeActionsToResolve.map(function(a,i){ return (i+1)+'. '+a; }).join('\n')
              : (c.threeActionsToResolve || '')
          ]
        ]
      )
    );

    // ── Probing Questions ──
    if (c.probingQuestions) {
      var pq = c.probingQuestions;
      html += section('Probing Questions Analysis', '&#10067;',
        analysisTable(
          ['Item', 'Detail'],
          [
            ['Questions Asked',     Array.isArray(pq.examples) ? pq.examples.join('\n') : ''],
            ['Were They Effective', pq.effective || ''],
            ['Better Questions',    Array.isArray(pq.betterQuestions) ? pq.betterQuestions.join('\n') : '']
          ]
        )
      );
    }

    // ── Positioning Statements ──
    if (c.criticalFlags && c.criticalFlags.length) {
      var posHtml = '';
      c.criticalFlags.forEach(function(f) { posHtml += posCard(f); });
      html += section('Critical Flags &amp; Positioning Statements', '&#128681;', posHtml);
    }

    html += '</div>'; // end r-call-block
  });

  // ── Highlights & Lowlights ──
  if ((d.highlights && d.highlights.length) || (d.lowlights && d.lowlights.length)) {
    html += section('Highlights &amp; Lowlights', '&#9733;', highlightBox(d.highlights, d.lowlights));
  }

  // ── Coaching Takeaways ──
  var takeaways = [];
  (d.calls || []).forEach(function(c) {
    if (c.coachingOpportunities) takeaways.push(c.coachingOpportunities);
  });
  if (d.overallSummary) takeaways.unshift(d.overallSummary);
  if (takeaways.length) {
    html += section('Coaching Takeaways', '&#127979;', coachingList(takeaways));
  }

  // ── Callback flag warning ──
  if (d.callbackFlaggedAgents && d.callbackFlaggedAgents.length) {
    html += '<div class="r-callback-warning">&#9888; <strong>Callback Flagged:</strong> ' +
      esc(d.callbackFlaggedAgents.join(', ')) + '</div>';
  }

  return html;
}

// ─────────────────────────────────────────────────────────────────────────────
// SALES HTML
// ─────────────────────────────────────────────────────────────────────────────
function buildSalesHTML(d) {
  var html   = '';
  var scores = d.scores || {};

  // ── Score panel ──
  var scoreRows = [
    ['Building Rapport',     scores.buildingRapport],
    ['Needs Identification', scores.needsIdentification],
    ['Product Presentation', scores.productPresentation],
    ['Objection Handling',   scores.objectionHandling],
    ['Closing Techniques',   scores.closingTechniques]
  ];
  var scorePanel = '<div class="r-score-panel">';
  scoreRows.forEach(function(r) {
    var sc = r[1] ? (r[1].total || 0) : 0;
    scorePanel += '<div class="r-score-row"><span>' + esc(r[0]) + '</span>' + scoreBadge(sc) + '</div>';
  });
  scorePanel += '<div class="r-score-row r-score-total"><span>Total Score</span>' + scoreBadge(d.overallScore || 0) + '</div>';
  scorePanel += '</div>';
  html += section('Overall Performance Score', '&#128200;', scorePanel);

  // ── Call Summary ──
  html += section('Call Summary', '&#128222;',
    '<div class="r-summary-box" contenteditable="true">' + esc(d.callSummary || '') + '</div>'
  );

  // ── Overall info ──
  html += section('Interaction Overview', '&#128100;',
    infoBar([
      ['Sale Occurred',   d.saleOccurred],
      ['Product Sold',    d.productSold],
      ['Sale Initiator',  d.saleInitiator]
    ]) +
    (d.perfectMoment ? '<div class="r-perfect-moment"><strong>Perfect Sales Moment:</strong> <span contenteditable="true">' + esc(d.perfectMoment) + '</span></div>' : '')
  );

  // ── Framework tables ──
  var frameworks = [
    ['Building Rapport',     scores.buildingRapport],
    ['Needs Identification', scores.needsIdentification],
    ['Product Presentation', scores.productPresentation],
    ['Objection Handling',   scores.objectionHandling],
    ['Closing Techniques',   scores.closingTechniques]
  ];
  frameworks.forEach(function(fw) {
    if (!fw[1] || !fw[1].subelements) return;
    var rows = fw[1].subelements.map(function(s) {
      return [s.name || '', s.score + '/5', s.analysis || '', s.strengths || '', s.opportunities || ''];
    });
    html += section(fw[0], '',
      '<table class="r-table"><thead><tr>' +
      '<th>Skill Name</th><th style="width:70px;text-align:center">Score</th>' +
      '<th>Analysis</th><th>Strengths</th><th>Opportunities</th>' +
      '</tr></thead><tbody>' +
      rows.map(function(r) {
        return '<tr>' +
          '<td class="r-table-label">' + esc(r[0]) + '</td>' +
          '<td style="text-align:center">' + scoreBadge(parseFloat(r[1]) || 0) + '</td>' +
          cell(r[2]) + cell(r[3]) + cell(r[4]) +
          '</tr>';
      }).join('') +
      '</tbody></table>'
    );
  });

  // ── Critical Flags / Positioning Statements ──
  if (d.criticalFlags && d.criticalFlags.length) {
    var posHtml = '';
    d.criticalFlags.forEach(function(f) { posHtml += posCard(f); });
    html += section('Critical Flags &amp; Positioning Statements', '&#128681;', posHtml);
  }

  // ── Highlights & Lowlights ──
  if ((d.highlights && d.highlights.length) || (d.lowlights && d.lowlights.length)) {
    html += section('Highlights &amp; Lowlights', '&#9733;', highlightBox(d.highlights, d.lowlights));
  }

  // ── Coaching Takeaways ──
  var takeaways = (d.top3Opportunities || []).concat(d.keyStrengths || []);
  if (takeaways.length) {
    html += section('Coaching Takeaways', '&#127979;', coachingList(takeaways));
  }

  if (d.systemicIssues) {
    html += section('Systemic Issues', '&#9888;',
      '<p contenteditable="true" class="r-systemic">' + esc(d.systemicIssues) + '</p>'
    );
  }

  return html;
}

// ─────────────────────────────────────────────────────────────────────────────
// Fallback when JSON cannot be parsed — clean up raw AI text
// ─────────────────────────────────────────────────────────────────────────────
function buildFallbackHTML(rawText) {
  var clean = rawText
    .replace(/```json[\s\S]*?```/g, '')
    .replace(/```[\s\S]*?```/g, '')
    .replace(/<JSON_ANALYSIS>[\s\S]*?<\/JSON_ANALYSIS>/g, '')
    .trim();

  // Convert markdown headings, bullets, tables to basic HTML
  var lines = clean.split('\n');
  var html  = '<div class="r-fallback">';
  lines.forEach(function(line) {
    var hm = line.match(/^(#{1,4})\s+(.+)/);
    if (hm) {
      html += '<div class="r-section-title" style="margin-top:16px">' + esc(hm[2]) + '</div>';
    } else if (/^[\*\-]\s/.test(line)) {
      html += '<li contenteditable="true" style="margin-left:20px;font-size:13px;line-height:1.8">' + esc(line.replace(/^[\*\-]\s/,'')) + '</li>';
    } else if (/^\d+\.\s/.test(line)) {
      html += '<li contenteditable="true" style="margin-left:20px;font-size:13px;line-height:1.8">' + esc(line.replace(/^\d+\.\s/,'')) + '</li>';
    } else if (line.trim() === '') {
      html += '<div style="height:8px"></div>';
    } else {
      html += '<p contenteditable="true" style="font-size:13px;line-height:1.7;margin:4px 0">' + esc(line) + '</p>';
    }
  });
  html += '</div>';
  return html;
}

