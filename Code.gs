/**
 * Code.gs
 * Web app entry point and server-side handlers.
 */

// ── Run from GAS editor: find + repair truncated Sales evaluation HTML ────────
// Step 1 — diagnose: shows which Sales cache entries look truncated
function diagnoseSalesEvaluations() {
  requireAdmin_();
  var ss         = getOrCreateSpreadsheet();
  var cacheSheet = getOrCreateSheet(ss, CACHE_SHEET);
  var lastRow    = cacheSheet.getLastRow();
  if (lastRow < 2) { Logger.log('Cache sheet is empty.'); return; }

  var data      = cacheSheet.getRange(2, 1, lastRow - 1, 5).getValues();
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
      sapId: (row[4] || '').toString().trim(),
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
// ── Clears ev2_ CacheService entries for all audit refs tied to an interaction ID ──
// Called after repair functions so EvalView pages serve the repaired HTML immediately
// instead of serving stale broken HTML until the 6-hour CacheService TTL expires.
function _clearEv2CacheForInteraction_(interactionId, ss) {
  if (!interactionId) return;
  try {
    var logSheet = getOrCreateSheet(ss || getOrCreateSpreadsheet(), AUDIT_LOG_SHEET);
    var lastRow  = logSheet.getLastRow();
    if (lastRow < 2) return;
    var headers  = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
    var idCol    = headers.indexOf('Interaction ID');
    var refCol   = headers.indexOf('Audit Ref');
    if (idCol < 0 || refCol < 0) return;
    var data = logSheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    var sc   = CacheService.getScriptCache();
    data.forEach(function(row) {
      if ((row[idCol] || '').toString().trim() === interactionId) {
        var auditRef = (row[refCol] || '').toString().trim();
        if (auditRef) {
          var evKey = 'ev2_' + Utilities.base64Encode(auditRef).substring(0, 200);
          try { sc.remove(evKey); } catch(ce) {}
        }
      }
    });
    Logger.log('_clearEv2CacheForInteraction_: cleared ev2_ entries for ' + interactionId);
  } catch(e) {
    Logger.log('_clearEv2CacheForInteraction_ error: ' + e);
  }
}

function repairNextTruncated() {
  requireAdmin_();
  var diagnosis = diagnoseSalesEvaluations();
  if (!diagnosis || !diagnosis.truncatedList || !diagnosis.truncatedList.length) {
    Logger.log('✅ All Sales evaluations complete — nothing left to repair.');
    return;
  }

  // Several agents can share one Interaction ID, and each has its own cached evaluation (told apart by the SAP ID in column E). This
  // tool re-runs the AI on the whole transcript for ONE evaluation, so it only works on an interaction whose sales evaluation belongs
  // to a single agent; deleting by Interaction ID alone used to wipe the other agents' evaluations as well.
  var cacheForIds = getOrCreateSheet(getOrCreateSpreadsheet(), CACHE_SHEET);
  var cacheRows = cacheForIds.getLastRow() >= 2 ? cacheForIds.getRange(2, 1, cacheForIds.getLastRow() - 1, 5).getValues() : [];
  var sapsOf = function(id) {
    var seen = {}, n = 0;
    cacheRows.forEach(function(r) {
      var t = (r[1] || '').toString().trim().toLowerCase();
      if ((r[0] || '').toString().trim() === id && t === 'sales') { var s = (r[4] || '').toString().trim(); if (!seen[s]) { seen[s] = true; n++; } }
    });
    return n;
  };
  var entry = null;
  for (var ti = 0; ti < diagnosis.truncatedList.length; ti++) {
    var cand = diagnosis.truncatedList[ti];
    if (sapsOf(cand.interactionId.trim()) > 1) {
      Logger.log('SKIP row=' + cand.row + ' id=' + cand.interactionId.substring(0, 8) + '... — several agents share this interaction, repair it by hand (re-submit that audit)');
      continue;
    }
    entry = cand; break;
  }
  if (!entry) { Logger.log('Nothing left that is safe to repair automatically.'); return; }
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

  var newHtml = analyzeTranscript_(transcript, 'sales');
  newHtml = fixBadgeClasses(
    newHtml.replace(/^```html\s*/i,'').replace(/^```\s*/,'').replace(/\s*```$/,'').trim()
  );

  // Guard: if the AI returned empty, skip the repair entirely — deleting the old
  // rows before this check would permanently destroy the previous (truncated but
  // at least partial) cached result with nothing to replace it.
  if (!newHtml) {
    Logger.log('repairNextTruncated: AI returned empty for ' + entry.interactionId + ' — skipping');
    return;
  }

  // Delete all old rows for this interaction ID (main + chunks) bottom-to-top
  // so row numbers stay valid while deleting, and no ghost empty rows remain
  var deleteFinder = cacheSheet.getRange('A:A').createTextFinder(entry.interactionId.trim()).matchEntireCell(true);
  var rowsToDelete = [];
  deleteFinder.findAll().forEach(function(match) {
    var rowType = (cacheSheet.getRange(match.getRow(), 2).getValue() || '').toString().toLowerCase();
    var rowSap  = (cacheSheet.getRange(match.getRow(), 5).getValue() || '').toString().trim();
    if ((rowType === 'sales' || /^sales_\d+$/.test(rowType)) && rowSap === entry.sapId) {
      rowsToDelete.push(match.getRow());
    }
  });
  rowsToDelete.sort(function(a, b) { return b - a; }); // delete bottom-to-top
  rowsToDelete.forEach(function(rowNum) { cacheSheet.deleteRow(rowNum); });
  SpreadsheetApp.flush();

  // Save with chunking — handles HTML > 48k chars transparently
  saveCachedResult_(entry.interactionId.trim(), 'sales', newHtml, entry.sapId);
  SpreadsheetApp.flush();

  // Clear ev2_ CacheService entries for all audit refs tied to this interaction.
  // Without this, EvalView pages loaded before the repair continue serving the old
  // broken HTML until the 6-hour CacheService TTL expires on its own.
  _clearEv2CacheForInteraction_(entry.interactionId.trim(), ss);

  var chunks = Math.ceil(newHtml.length / 48000);
  Logger.log('✅ Repaired id=' + entry.interactionId.substring(0,8) + '...' +
             ' len=' + newHtml.length + ' chunks=' + chunks +
             ' (' + (diagnosis.truncatedList.length - 1) + ' remaining)');
}

// Step 2b — bulk repair (legacy — use repairNextTruncated() to avoid timeout).
function repairTruncatedSalesEvaluations() {
  requireAdmin_();
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
      var newHtml = analyzeTranscript_(transcript, 'sales');
      newHtml = fixBadgeClasses(newHtml
        .replace(/^```html\s*/i, '').replace(/^```\s*/, '').replace(/\s*```$/, '').trim());

      // Use saveCachedResult_ so chunk rows are properly deleted before writing.
      // The old approach (direct setValue on entry.row) only updated the primary
      // chunk but left any sales_2/sales_3 rows intact — causing findCachedResult_
      // to concatenate new HTML + stale old chunks, producing worse corruption
      // than the truncation it was trying to fix.
      saveCachedResult_(entry.interactionId, 'sales', newHtml);

      // Clear ev2_ CacheService entries so EvalView pages serve the repaired HTML
      // immediately instead of waiting up to 6 hours for the TTL to expire.
      _clearEv2CacheForInteraction_(entry.interactionId.trim(), ss);
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
  requireAdmin_();
  var ss    = openSheetWithRetry(ROSTER_SHEET_ID, 'Roster Sheet');
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
  requireAdmin_();
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
    invalidateDashboardCache_();
  } else {
    Logger.log('Nothing to backfill — all rows already have data');
  }
}

// ── RCA enrichment prompt ─────────────────────────────────────────────────────
// Reads cached HTML for each audit, calls AI to extract structured RCA fields,
// writes them back to Dashboard_Data. Run once from the editor after submission.
function enrichDashboardData() {
  requireAdmin_();
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
      var structured = extractStructuredRCA_(html, atype, ref);
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

  invalidateDashboardCache_();
  Logger.log('=== Enriched ' + enriched + ' rows ===');
}

// ── Backfill blank Locale values in Dashboard_Data using roster lookup ────────
// Run ONCE from the Apps Script editor: open Code.gs → Run → backfillLocaleInDashboard
// Safe to re-run — only patches rows where Locale is currently blank.
function backfillLocaleInDashboard() {
  requireAdmin_();
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
      var rosterRow = lookupBySapId_(sapId);
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

  invalidateDashboardCache_();
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

    // ── Call driver ───────────────────────────────────────────────────────────
    // The evaluation now states this outright in div.ai-meta. The regex below is
    // kept only for reports produced before that existed; its old last resort was
    // flags[0], which put a fault FOUND on the call into a column meaning why the
    // customer CALLED — e.g. "Agent Used a False Name During Introduction".
    var callDriver = '';
    var mCallReason = html.match(/data-call-reason="([^"]*)"/i);
    if (mCallReason) {
      var cr = mCallReason[1].trim();
      if (cr && cr.charAt(0) !== '[' && cr.toLowerCase() !== 'n/a') callDriver = cr.substring(0, 100);
    }
    if (!callDriver) {
      var cdMatch = html.match(/Call (?:Reason|Driver)[^:]*[:\s]+([\s\S]{5,80}?)(?:<\/td>|<br|<\/p)/i);
      if (cdMatch) callDriver = stripTags(cdMatch[1]).substring(0, 100);
    }

    // These three used to require a SECOND AI call (enrichDashboardRCA), which
    // asked for 13 fields, wrote only 3, and read nothing this HTML does not
    // already contain. The evaluation prompt now emits them in a hidden
    // div.ai-meta and they are read straight off it.
    function metaAttr(name) {
      var m = html.match(new RegExp('data-' + name + '="([^"]*)"', 'i'));
      if (!m) return '';
      var v = m[1].trim();
      // A model that echoed the placeholder instead of answering leaves the
      // square brackets behind; treat that as no answer rather than as data.
      if (v.charAt(0) === '[' || v.toLowerCase() === 'n/a') return '';
      return v;
    }

    return {
      callDriver:        callDriver,
      rcaCategory:       metaAttr('rca-category').substring(0, 50),
      rcaSubParameter:   flags.slice(0,2).join('; ').substring(0, 100),
      topOpportunity:    topOpp,
      productOpportunity:metaAttr('product-opportunity').substring(0, 100),
      salesAttempted:    metaAttr('sales-attempted').substring(0, 10),
      callSummaryShort:  callSummary,
      criticalFlags:     flags.slice(0, 5).join(' | ').substring(0, 500),
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
function extractStructuredRCA_(html, analysisType, auditRef) {
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
  requireAdmin_();
  try {
    if (!auditRef || !html) return;

    var aiRca = extractStructuredRCA_(html, analysisType, auditRef);
    if (!aiRca) {
      Logger.log('enrichDashboardRCA: AI returned null for ' + auditRef);
      return;
    }

    var ss    = getOrCreateSpreadsheet();
    var sheet = getOrCreateSheet(ss, DASHBOARD_DATA_SHEET);
    if (sheet.getLastRow() < 2) return;

    // Find the row by auditRef using TextFinder (same pattern as findCachedResult_)
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

    invalidateDashboardCache_();
    Logger.log('enrichDashboardRCA: patched row ' + targetRow + ' for ' + auditRef);
  } catch(e) {
    Logger.log('enrichDashboardRCA error (non-fatal): ' + e);
  }
}

// ── Translate a transcript to English using FuelIX ───────────────────────────
function translateTranscriptText_(text) {
  if (!text || !text.trim()) return null;
  var prompt =
    'Translate the following call center transcript to English.\n' +
    'Rules:\n' +
    '- Preserve all speaker labels, timestamps, and metadata tags exactly as they appear.\n' +
    '- Translate ONLY the spoken dialogue — do not translate system metadata, IDs, or timestamps.\n' +
    '- Keep the same formatting and line structure.\n' +
    '- Return ONLY the translated transcript with no commentary, preamble, or explanation.\n\n' +
    'TRANSCRIPT:\n\n' + text;

  var payload = {
    model:      FUELIX_CONFIG.model,
    messages:   [{ role: 'user', content: prompt }],
    stream:     false,
    max_tokens: 8000
  };
  // deadline capped at 55s (not 120s) because translateTranscriptText_ runs inside
  // submitTranscript BEFORE the main AI analysis call (deadline 270s). The GAS hard
  // execution limit is 360s. Budget: 6s lock + 55s translate + 270s AI + ~20s I/O = 351s.
  // A 120s translate deadline would leave only 360-120-6-20 = 214s for AI analysis,
  // which is not enough for large Sales transcripts requiring max_tokens:20000.
  var options = {
    method: 'post', contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + FUELIX_CONFIG.apiKey, 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true, deadline: 55
  };

  try {
    var response = UrlFetchApp.fetch(FUELIX_CONFIG.baseUrl + '/v1/chat/completions', options);
    if (response.getResponseCode() !== 200) {
      Logger.log('translateTranscriptText_ HTTP error: ' + response.getResponseCode());
      return null;
    }
    var content = JSON.parse(response.getContentText()).choices[0].message.content.trim();
    Logger.log('Translation complete: ' + content.length + ' chars (original: ' + text.length + ')');
    return content || null;
  } catch(e) {
    Logger.log('translateTranscriptText_ error: ' + e);
    return null;
  }
}

// ── Text-to-Speech via FuelIX — returns base64 MP3 or null on failure ────────
function synthesizeSpeech(text) {
  if (!text || !text.trim()) return null;
  var payload = {
    model: 'tts-1',
    input: text.trim().substring(0, 4096),
    voice: 'nova',
    response_format: 'mp3'
  };
  var options = {
    method: 'post', contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + FUELIX_CONFIG.apiKey, 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true, deadline: 30
  };
  try {
    var response = UrlFetchApp.fetch(FUELIX_CONFIG.baseUrl + '/v1/audio/speech', options);
    var code = response.getResponseCode();
    if (code !== 200) {
      Logger.log('synthesizeSpeech HTTP ' + code + ': ' + response.getContentText().substring(0, 300));
      return null;
    }
    return Utilities.base64Encode(response.getContent());
  } catch(e) {
    Logger.log('synthesizeSpeech error: ' + e);
    return null;
  }
}

// ── Admin: run once from GAS editor to fix manually-typed locale values ──
// Fixes bad Locale values in both Dashboard_Data and Audit_Log sheets.
function fixTruncatedLocale() {
  requireAdmin_();
  var FIXES = [
    { bad: 'TI M',   good: 'TI Morocco (CAS)' },
    { bad: 'TI Mo',  good: 'TI Morocco (CAS)' },
    { bad: 'TI Mor', good: 'TI Morocco (CAS)' },
    { bad: 'TIM',    good: 'TI Morocco (CAS)' }
  ];

  var ss      = getOrCreateSpreadsheet();
  var total   = 0;

  total += _fixLocaleInSheet(ss, DASHBOARD_DATA_SHEET, FIXES);
  total += _fixLocaleInSheet(ss, AUDIT_LOG_SHEET,      FIXES);

  if (total > 0) invalidateDashboardCache_();
  Logger.log('=== fixTruncatedLocale done: ' + total + ' patched across both sheets ===');
}

function _fixLocaleInSheet(ss, sheetName, fixes) {
  var sheet   = getOrCreateSheet(ss, sheetName);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  var headers   = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var localeCol = headers.indexOf('Locale');
  if (localeCol < 0) { Logger.log('_fixLocaleInSheet: no Locale column in ' + sheetName); return 0; }
  localeCol++;

  var data    = sheet.getRange(2, localeCol, lastRow - 1, 1).getValues();
  var patched = 0;

  for (var i = 0; i < data.length; i++) {
    var cell = (data[i][0] || '').toString().trim();
    for (var j = 0; j < fixes.length; j++) {
      if (cell === fixes[j].bad) {
        sheet.getRange(i + 2, localeCol).setValue(fixes[j].good);
        Logger.log('fixLocale [' + sheetName + '] row ' + (i + 2) + ': "' + cell + '" → "' + fixes[j].good + '"');
        patched++;
        break;
      }
    }
  }

  Logger.log('fixLocale [' + sheetName + ']: ' + patched + ' patched');
  return patched;
}

// ── Run ONCE from the Apps Script editor to upgrade the Dashboard_Data sheet ──
// Renames 'Recommendations'→'SMART Recommendation', 'Top Opportunity'→'Top SMART Opportunity'
// and appends 5 new SMART columns (S/M/A/R/T). Safe to run on a live sheet —
// it only adds/renames header cells and does not touch data rows.
function updateSpreadsheetSMARTFormat() {
  requireAdmin_();
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
  requireAdmin_();
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
  requireAdmin_();
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
  requireAdmin_();
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
  var atSs    = openSheetWithRetry(AT_DATA_GCP_SS_ID, 'AT Data GCP / Trainee Roster Sheet');
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
  invalidateDashboardCache_();
  Logger.log('=== Updated VTID for ' + updated + ' Dashboard_Data rows ===');
}

function dbUniqServer(arr) {
  var seen = {};
  return arr.filter(function(v){ return v && !seen[v] && (seen[v]=1); });
}

function backfillVTIDToAuditLog() {
  requireAdmin_();
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
  invalidateDashboardCache_();
  Logger.log('=== Backfilled VTID for ' + updated + ' Audit_Log rows ===');
}

function fixRepeatProjectionPct() {
  requireAdmin_();
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

  invalidateDashboardCache_();
  Logger.log('=== Fixed Repeat Projection % for ' + fixed + ' rows ===');
}

function addDashboardColumnNotes() {
  requireAdmin_();
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
  requireAdmin_();
  try {
    // Test getRosterBySapId directly to reproduce the error
    Logger.log('=== Testing getRosterBySapId ===');
    var testSap = '2007888';
    var start = new Date();
    var result = lookupBySapId_(testSap);
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
  requireAdmin_();
  var SAP_ID = '2007888';
  var ss     = openSheetWithRetry(FCR_DASHBOARD_SS_ID, 'Main Analyzer Spreadsheet');
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
  var result = lookupFromFCRDashboard_(SAP_ID, Number(SAP_ID));
  Logger.log('Result: ' + JSON.stringify(result));
}

function diagnoseATDataRow() {
  requireAdmin_();
  var SAP_ID = '2007888';
  var ss     = openSheetWithRetry(AT_DATA_GCP_SS_ID, 'AT Data GCP / Trainee Roster Sheet');
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
  requireAdmin_();
  var SAP_ID     = '2007888';
  var AGENT_NAME = 'James Collado';

  Logger.log('=== DIRECT LOOKUP TEST ===');

  // 1. Search roster by AGENT NAME to find actual stored SAP ID
  var ss    = openSheetWithRetry(ROSTER_SHEET_ID, 'Roster Sheet');
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
  var vtid = lookupVTID_(SAP_ID);
  Logger.log('VTID result: [' + vtid + ']');

  // 3. Full lookup
  Logger.log('--- Full lookupBySapId_ for: ' + SAP_ID + ' ---');
  var result = lookupBySapId_(SAP_ID);
  Logger.log('Result: ' + JSON.stringify(result));

  if (!result) {
    Logger.log('RESULT IS NULL — checking raw sheet data...');
    var ss    = openSheetWithRetry(ROSTER_SHEET_ID, 'Roster Sheet');
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
  requireAdmin_();
  var SAP_ID = '2007888';   // <-- change this to test any SAP ID

  var ss    = openSheetWithRetry(ROSTER_SHEET_ID, 'Roster Sheet');
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
  requireAdmin_();
  var AGENT_NAME = 'James Collado';  // <-- change to test any agent

  var ss    = openSheetWithRetry(AT_DATA_GCP_SS_ID, 'AT Data GCP / Trainee Roster Sheet');
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
  requireAdmin_();
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
// Usage: call notifyAdmins_(formData, auditRef) after each submitTranscript
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Reads the Audit Tracking SS → Roster tab and returns all rows matching
 * the given role (case-insensitive). Columns detected by header name:
 * Name | Role | Email Address | Supervisor | Supervisor Email
 * Results are cached for 2 hours.
 */
var _rosterRecipientsInMemory = null;

function getRecipientsFromRoster_(roleFilter) {
  try {
    var allRows = null;
    var source  = 'sheet';   // where this call's rows came from, for the log line below

    if (_rosterRecipientsInMemory) {
      allRows = _rosterRecipientsInMemory;
      source  = 'memory';
    } else {
      var cache    = CacheService.getScriptCache();
      var cacheKey = 'audit_roster_rows_v3'; // v3: v2 could hold a stale list that matched nobody (sendAuditEmail "could not resolve recipients")
      var cached   = cache.get(cacheKey);
      if (cached) {
        try { allRows = JSON.parse(cached); } catch(e) {}
        if (allRows) source = 'cache';
      }
      if (!allRows) {
        var ss    = openSheetWithRetry(AUDIT_TRACKING_SS_ID, 'Main Analyzer Spreadsheet');
        var sheet = ss.getSheetByName('Roster') || ss.getSheetByName('roster');
        if (!sheet) { Logger.log('getRecipientsFromRoster_: Roster sheet not found'); return []; }
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
        // Say WHICH tab was read and whether the columns were found. A tab of the
        // same name with other columns reads as "nobody has this role", silently.
        Logger.log('getRecipientsFromRoster_: read tab "' + sheet.getName() + '" (' + (data.length - 1) +
                   ' rows); name col=' + nameCol + ', role col=' + roleCol + ', email col=' + emailCol +
                   '; headers=' + JSON.stringify(hdrs.slice(0, 8)));
        if (roleCol === -1 || emailCol === -1) {
          Logger.log('getRecipientsFromRoster_ WARNING: this tab has no "Role" and/or "Email Address" column, so no recipient can be found in it.');
        }
        allRows = data.slice(1).map(function(row) {
          return {
            name:            nameCol     > -1 ? (row[nameCol]     || '').toString().trim() : '',
            role:            roleCol     > -1 ? (row[roleCol]     || '').toString().trim() : '',
            email:           emailCol    > -1 ? (row[emailCol]    || '').toString().trim() : '',
            supervisor:      supCol      > -1 ? (row[supCol]      || '').toString().trim() : '',
            supervisorEmail: supEmailCol > -1 ? (row[supEmailCol] || '').toString().trim() : ''
          };
        });
        var _ar = JSON.stringify(allRows);
        if (_ar.length <= 99000) { try { cache.put(cacheKey, _ar, 2 * 60 * 60); } catch(e) {} }
        _rosterRecipientsInMemory = allRows;
      }
    }

    var roleLower = roleFilter.toLowerCase().trim();
    var validRe   = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    var results   = allRows.filter(function(r) {
      return r.role.toLowerCase() === roleLower && validRe.test(r.email);
    });
    // Deduplicate by email — Roster sheet may have multiple rows for same address
    var seenEmails = {};
    results = results.filter(function(r) {
      var key = r.email.toLowerCase();
      if (seenEmails[key]) return false;
      seenEmails[key] = true;
      return true;
    });
    var roleCounts = {};
    allRows.forEach(function(r) { var k = r.role || '(blank)'; roleCounts[k] = (roleCounts[k] || 0) + 1; });
    Logger.log('getRecipientsFromRoster_("' + roleFilter + '"): ' + results.length + ' found (after dedup); source=' + source +
               '; rows=' + allRows.length + '; roles present=' + JSON.stringify(roleCounts));
    return results;
  } catch(e) {
    Logger.log('getRecipientsFromRoster_ error: ' + e);
    return [];
  }
}

/**
 * Returns Admin/Dev recipients from the Audit Tracking SS Roster tab.
 */
function getAdminRecipients_() {
  return getRecipientsFromRoster_('Admin/Dev');
}

/**
 * Builds the deduplicated recipient list for an audit email.
 *
 * Trainee  (SAP ID found in trainee roster):
 *   Team Member + Trainer Email + Trainer Supervisor Email + QA Observer
 * Tenured  (SAP ID not in trainee roster):
 *   Team Member + Team Leader Email + OM Email + QA Observer
 *
 * Admin/Dev is intentionally excluded — notifyAdmins_() handles that separately.
 */
function buildAuditRecipients_(formData) {
  var agentName = (formData.participant || '').trim();
  var sapId     = (formData.sapId       || '').trim();

  // ── 1. Check trainee status FIRST ────────────────────────────────────────
  // Trainees are not in the regular agent roster — checking it for them wastes
  // a lookup and returns nothing useful. Branch immediately on trainee status.
  //
  // Guard: if sapId is blank the trainee check returns null immediately.
  // Attempt a name→SAP reverse lookup so a transcript-parsed name still routes correctly.
  var resolvedSapId = sapId || (agentName ? lookupSapId_(agentName) : '');
  var traineeInfo = getTraineeInfo_(resolvedSapId);
  var teamMemberEmail, recipientB, recipientC;

  if (traineeInfo) {
    // Trainee — all info comes from trainee roster + trainer lookup.
    // No need to touch the regular agent roster at all.
    teamMemberEmail = traineeInfo.username ? traineeInfo.username + '@telus.com'
                    : (formData.agentEmail || '').trim();
    if (!teamMemberEmail) {
      Logger.log('buildAuditRecipients_ [trainee WARNING]: no email for ' + agentName +
                 ' (SAP ' + resolvedSapId + ') — username field is blank in trainee roster.' +
                 ' Trainee will NOT receive the audit email.');
    }

    var trainerInfo = getTrainerInfo_(traineeInfo.facilitator);
    if (!trainerInfo) {
      Logger.log('buildAuditRecipients_ [trainee WARNING]: facilitator "' + traineeInfo.facilitator +
                 '" not found in Trainer Lookup. Trainer and Supervisor will NOT receive the audit email.' +
                 ' Check spelling in trainee roster vs trainer lookup sheet.');
    }
    recipientB = trainerInfo ? trainerInfo.trainerEmail    : '';
    recipientC = trainerInfo ? trainerInfo.supervisorEmail : '';
    Logger.log('buildAuditRecipients_ [trainee]: ' + agentName +
               ' | facilitator=' + traineeInfo.facilitator +
               ' | trainer='    + recipientB +
               ' | supervisor=' + recipientC);
  } else {
    // Tenured — TL + OM from formData. The agent's address is resolved by who
    // the SAP ID IDENTIFIES, not by what they are called: a name can be shared by
    // two roster agents, and the old name-first order sent one of them the
    // other's evaluation. Order: the roster row for the SAP ID, then the address
    // the SAP lookup showed the analyst, then the name — which only resolves
    // when exactly one person has it (see _getAgentEmailMap).
    teamMemberEmail = _agentEmailForSap_(sapId)
                    || (formData.agentEmail || '').trim()
                    || lookupAgentEmail_(agentName)
                    || resolveEmail_(agentName);
    recipientB = resolveEmail_(formData.teamLeader || '');
    recipientC = resolveEmail_(formData.opsManager || '');
    Logger.log('buildAuditRecipients_ [tenured]: ' + agentName +
               ' | TL=' + recipientB + ' | OM=' + recipientC);
  }

  // ── 3. QA / Observer ─────────────────────────────────────────────────────
  var observerName = (formData.observerName  || '').trim();
  var qaEmail      = (formData.observerEmail || '').trim();
  if (!qaEmail && observerName) {
    var qaRows  = getRecipientsFromRoster_(QA_ROLE);
    var qaMatch = qaRows.filter(function(r) { return r.name.toLowerCase() === observerName.toLowerCase(); });
    qaEmail = qaMatch.length ? qaMatch[0].email : resolveEmail_(observerName);
  }

  // ── Deduplicate + validate ────────────────────────────────────────────────
  var validRe = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  var seen = {};
  return [teamMemberEmail, recipientB, recipientC, qaEmail].filter(function(e) {
    if (!e || !validRe.test(e) || seen[e.toLowerCase()]) return false;
    seen[e.toLowerCase()] = true;
    return true;
  });
}

// ── Agent email lookup from main roster file ──────────────────────────────────
// Builds a name → email map from the "Team Member Email" column once per execution,
// cached in CacheService for 2 hours to avoid repeated sheet reads.
var _agentEmailMapInMemory = null;

function _getAgentEmailMap() {
  if (_agentEmailMapInMemory) return _agentEmailMapInMemory;
  try {
    var cache    = CacheService.getScriptCache();
    var cacheKey = 'agent_email_map_v3';   // v3: shared names are null, not last-row-wins
    var cached   = cache.get(cacheKey);
    if (cached) {
      try { _agentEmailMapInMemory = JSON.parse(cached); return _agentEmailMapInMemory; } catch(e) {}
    }
    var ss    = openSheetWithRetry(ROSTER_SHEET_ID, 'Roster Sheet');
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
      _agentEmailMapInMemory = {};
      return {};
    }
    // A name shared by people with DIFFERENT addresses is stored as null: it
    // identifies nobody, and guessing sends one person the other's evaluation.
    // The same address listed twice is one person and is kept.
    var map = {};
    data.slice(1).forEach(function(row) {
      var name  = (row[nameCol]  || '').toString().trim().toLowerCase();
      var email = (row[emailCol] || '').toString().trim();
      if (!name || !email) return;
      if (!Object.prototype.hasOwnProperty.call(map, name)) { map[name] = email; return; }
      if (map[name] !== null && map[name].toLowerCase() !== email.toLowerCase()) map[name] = null;
    });
    var _em = JSON.stringify(map);
    if (_em.length <= 99000) { try { cache.put(cacheKey, _em, 2 * 60 * 60); } catch(e) {} }
    _agentEmailMapInMemory = map;
    Logger.log('_getAgentEmailMap: loaded ' + Object.keys(map).length + ' entries');
    return map;
  } catch(e) {
    Logger.log('_getAgentEmailMap error: ' + e);
    return {};
  }
}

function lookupAgentEmail_(participantName) {
  if (!participantName) return '';
  var map = _getAgentEmailMap();
  return map[participantName.trim().toLowerCase()] || '';   // null (shared name) -> ''
}

// The address on the roster row for this exact SAP ID, or ''. Unlike the name
// map this cannot land on a namesake.
function _agentEmailForSap_(sapId) {
  if (!sapId) return '';
  var hit = _findRosterRow_(ROSTER_COL_SAP_ID, sapId);
  return hit ? (hit.email || '') : '';
}

/**
 * Sends an admin notification email with the evaluation summary.
 * Called automatically after each new submitTranscript.
 *
 * @param {Object} formData - The submitted form data
 * @param {string} auditRef - The generated audit reference number
 */
function notifyAdmins_(formData, auditRef) {
  try {
    var recipients = getAdminRecipients_();
    if (!recipients.length) {
      Logger.log('notifyAdmins_: No Admin/Dev recipients found — skipping');
      return;
    }

    // Quota guard: reserve at least 50 units for the critical main audit emails.
    // Admin notification is secondary — skip it rather than starve the main email.
    try {
      var quotaLeft = MailApp.getRemainingDailyQuota();
      var needed    = recipients.length + 50; // admin list + safety buffer
      if (quotaLeft < needed) {
        Logger.log('notifyAdmins_: quota too low (' + quotaLeft + ' remaining, need ' + needed + ') — skipping to protect main email quota');
        return;
      }
    } catch(qe) {
      Logger.log('notifyAdmins_: quota check failed — ' + qe + ' — proceeding anyway');
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
    var auditType     = formData.auditType    || 'N/A';
    var submittedAt   = new Date().toLocaleString();

    // Fix: truncate interactionId to prevent transcript blob leaking into email
    var cleanIntId = interactionId.replace(/[\r\n\t]/g, ' ').substring(0, 120);

    var evalUrl  = ScriptApp.getService().getUrl() + '?page=eval&ref=' + encodeURIComponent(auditRef);
    var subject  = '[Admin] ' + agentName + ' | ' + auditRef + ' | ' + observer + ' | ' + auditType + ' | ' + lob + ' | ' + locale;

    // ── Plain text body ───────────────────────────────────────────────────────
    var plainBody =
      'A new QA audit has been submitted.\n\n' +
      'AUDIT REFERENCE: ' + auditRef + '\n' +
      '─────────────────────────────────────\n' +
      'Submitted At:     ' + submittedAt + '\n' +
      'Observer:         ' + observer + '\n' +
      'Analysis Type:    ' + analysisType + '\n' +
      'Audit Type:       ' + auditType + '\n' +
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
      auditRef, submittedAt, observer, analysisType, auditType,
      agentName, sapId, vtid, teamLeader, opsManager, lob, locale,
      cleanIntId, direction, duration, evalUrl
    );

    // ── Send via batch helper — handles Admin/Dev rosters > 50 addresses ────────
    var adminEmails = recipients.map(function(r) { return r.email; });
    try {
      sendEmailInBatches_(adminEmails, subject, plainBody, htmlBody,
                          'NH Call Analyzer — Admin Notifications');
      Logger.log('Admin notification sent to: ' + adminEmails.join(', '));
    } catch(emailErr) {
      Logger.log('Admin notification failed: ' + emailErr);
    }

  } catch(e) {
    Logger.log('notifyAdmins_ error: ' + e.toString());
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
  requireAdmin_();
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
  requireAdmin_();
  var recipients = getAdminRecipients_();
  Logger.log('=== Admin/Dev recipients ===');
  if (!recipients.length) {
    Logger.log('NONE FOUND — check that the Roster sheet has a "Role" column with value "Admin/Dev"');
    return;
  }
  recipients.forEach(function(r){
    Logger.log('  ' + r.name + ' <' + r.email + '>');
  });

  // Send a test email using sample data
  notifyAdmins_({
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

/** Run from GAS editor to verify agent email lookup is working */
function testEmailLookup() {
  requireAdmin_();
  // ── 1. Clear stale cache so we force a fresh sheet read ───────────────────
  CacheService.getScriptCache().remove('agent_email_map_v1');
  CacheService.getScriptCache().remove('agent_email_map_v2');
  _agentEmailMapInMemory = null;
  Logger.log('Cache cleared.');

  // ── 2. Check what _getAgentEmailMap builds ────────────────────────────────
  var map = _getAgentEmailMap();
  var entries = Object.keys(map).length;
  Logger.log('_getAgentEmailMap: ' + entries + ' entries loaded.');
  if (entries === 0) {
    Logger.log('ERROR: Map is empty — check Primary Roster for Agent_Name and Team Member Email columns.');
    return;
  }

  // ── 3. Show first 5 entries as sample ─────────────────────────────────────
  Logger.log('Sample entries:');
  var count = 0;
  for (var name in map) {
    if (count++ >= 5) break;
    Logger.log('  "' + name + '" → ' + map[name]);
  }

  // ── 4. Test a specific name (change this to a real agent name) ────────────
  var testName = 'Harpreet Singh'; // ← change to actual agent name to test
  var result   = lookupAgentEmail_(testName);
  Logger.log('lookupAgentEmail_("' + testName + '"): "' + result + '"');
  var result2  = resolveEmail_(testName);
  Logger.log('resolveEmail_("' + testName + '"): "' + result2 + '"');
}

// ─────────────────────────────────────────────────────────────────────────────
// ONE-TIME REPAIR: Run this from the Apps Script editor (Run → repairSheetHeaders)
// It re-writes every sheet so columns match the current header definitions.
// Safe to run multiple times — it never deletes data, only reorders columns.
// ─────────────────────────────────────────────────────────────────────────────
function repairSheetHeaders() {
  requireAdmin_();
  var ss = getOrCreateSpreadsheet();
  Logger.log('=== Starting sheet header repair ===');

  // The repair clears each sheet and writes it again, so no audit may be saved in between (it would be wiped): take the same
  // script lock the audit save takes, and refuse to run without it.
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    _repairSheet(ss, TRANSCRIPTS_SHEET,    TRANSCRIPTS_HEADERS);
    _repairSheet(ss, ANALYSIS_SHEET,       ANALYSIS_HEADERS);
    _repairSheet(ss, DASHBOARD_DATA_SHEET, DASHBOARD_HEADERS);
    _repairSheet(ss, AUDIT_LOG_SHEET,      AUDIT_LOG_HEADERS);
  } finally {
    try { lock.releaseLock(); } catch (le) {}
  }

  // Invalidate dashboard cache so next open reflects repaired data
  invalidateDashboardCache_();

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
  var extraAfter = currentHdr.slice(correctHeaders.length).some(function(h) { return h; });
  if (alreadyCorrect && (currentHdr.length === correctHeaders.length || extraAfter)) {   // known columns first: extra ones after are kept as they are
    Logger.log(sheetName + ': headers already correct, skipping');
    return;
  }

  Logger.log(sheetName + ': repairing. Current headers: ' + currentHdr.join(' | '));
  Logger.log(sheetName + ': target headers:  ' + correctHeaders.join(' | '));

  // Build a map from current header name → column index
  var currentIndexOf = {};
  currentHdr.forEach(function(h, i) { if (h) currentIndexOf[h] = i; });

  // A column the code does not know (an old one such as Recommendations, or one somebody added) is KEPT after the known ones, with
  // its header and data. The repair used to drop it for good. A column with no header is kept only when it holds something.
  var dataRows   = allData.slice(1); // skip existing header row
  var known = {}; correctHeaders.forEach(function(h) { known[h] = true; });
  var extraIdx = [];
  currentHdr.forEach(function(h, i) {
    if (h && known[h] && currentIndexOf[h] === i) return;              // a known column, used in place
    if (h && known[h]) return;                                          // a repeated known header: the later copy is not kept twice
    var holdsData = dataRows.some(function(row) { return i < row.length && String(row[i] === null || row[i] === undefined ? '' : row[i]).trim() !== ''; });
    if (h || holdsData) extraIdx.push(i);
  });
  var allHeaders = correctHeaders.concat(extraIdx.map(function(i) { return currentHdr[i]; }));
  var reordered  = dataRows.map(function(row) {
    var known = correctHeaders.map(function(h) {
      var idx = currentIndexOf[h];
      return (idx !== undefined && idx < row.length) ? row[idx] : '';
    });
    return known.concat(extraIdx.map(function(i) { return i < row.length ? row[i] : ''; }));
  });
  if (extraIdx.length) Logger.log(sheetName + ': kept ' + extraIdx.length + ' extra column(s) after the standard ones: ' + extraIdx.map(function(i) { return currentHdr[i] || '(no header)'; }).join(', '));

  // Clear the sheet and rewrite cleanly
  sheet.clearContents();

  // Write header row with styling
  _writeHeaderRow(sheet, allHeaders);

  // Write data rows
  if (reordered.length > 0) {
    sheet.getRange(2, 1, reordered.length, allHeaders.length).setValues(reordered);
  }

  Logger.log(sheetName + ': repaired — ' + reordered.length + ' data rows rewritten with ' + allHeaders.length + ' columns');
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

function isValidAuditRef_(ref) {
  return typeof ref === 'string' && /^[A-Za-z0-9._-]{1,64}$/.test(ref);
}

function doGet(e) {
  var page = e && e.parameter && e.parameter.page;
  if (page === 'dashboard') return doGetDashboard_();
  if (page === 'eval') {
    var ref = (e && e.parameter && e.parameter.ref) || '';
    // An audit ref is letters, digits, dot, underscore and hyphen (NHA-20261012-0001,
    // NHA-20261012-0001-D2). Anything else is not a ref: it is printed into the page
    // and sent to the lookup, so it is refused rather than escaped and trusted.
    if (!isValidAuditRef_(ref)) ref = '';
    var tmpl = HtmlService.createTemplateFromFile('EvalView');
    tmpl.auditRef = ref;
    tmpl.preloadedData = 'null';
    if (ref) {
      try {
        var evalData = getEvalViewData(ref);
        if (evalData && evalData.success) {
          // Escape every "<" (not just "</script"): "<!--<script>" also confuses a
          // script block. \u003c is a valid JSON string escape, so the data is unchanged.
          tmpl.preloadedData = JSON.stringify(evalData)
            .replace(/</g, '\\u003c').replace(/>/g, '\\u003e').replace(/&/g, '\\u0026')
            .replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
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

    // 3. Cache sheet → AI HTML, keyed by Interaction ID + analysis type + the SAP
    // ID of THIS audit's agent (several agents can share one Interaction ID).
    // findCachedResult_ assembles continuation chunks and never returns another
    // agent's evaluation.
    var interactionId = meta['Interaction ID'] || dash['Interaction ID'] || '';
    var evalHtml = '';
    if (interactionId) {
      var rawType = (meta['Analysis Type'] || dash['Analysis Type'] || '').toLowerCase();
      var normalizedType = rawType.indexOf('sales') !== -1 ? 'sales' : 'repeats';
      var evalSapId = meta['SAP ID'] || dash['SAP ID'] || '';
      try {
        var chtml = findCachedResult_(interactionId, normalizedType, evalSapId);
        if (chtml) evalHtml = sharedCSS() + fixBadgeClasses(chtml);
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
            ban:  (r[logHeaders2.indexOf('Interaction ID')] || '').toString(),
            // EvalView.html reads these two names (it showed "—" and "Repeats" for every row before)
            interactionId: (r[logHeaders2.indexOf('Interaction ID')] || '').toString(),
            analysisType:  (r[logHeaders2.indexOf('Analysis Type')] || '').toString()
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
      if (evalHtml && resultJson.length < 90000) {   // not while the report is still being processed
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
              .replace(/&nbsp;/gi,' ')          // the report's meta line separates fields
                                                // with &nbsp; — without this the entity
                                                // survives into values ("Inbound &nbsp;")
              .replace(/&amp;/g,'&').replace(/&lt;/g,'<')
              .replace(/&gt;/g,'>').replace(/&quot;/g,'"')
              .replace(/&#\d+;/g,' ').replace(/\s+/g,' ').trim();
    };

    // 0. Known fields read from where they ACTUALLY live in the report.
    //
    // Four Dashboard columns had no strategy pointing at them, so they fell
    // through to the last-resort guess below and were filled with whatever text
    // happened to sit near the keyword — typically the info chips run together
    // ("DateWednesday, September 9, 2026 Issue ResolutionNo Repeat Risk92%...").
    // Every one of them has a real home in the report; nothing needed guessing.
    var kw = keyword.toLowerCase();

    function joinMatches(scope, inner, sep) {
      var hits = scope.match(inner) || [];
      var out  = [];
      for (var i = 0; i < hits.length; i++) {
        var t = decode(hits[i]);
        if (t) out.push(t);
      }
      return out.join(sep);
    }

    if (kw.indexOf('call summary') !== -1) {
      var mSum = html.match(/<div[^>]*ai-summary[^>]*>([\s\S]*?)<\/div>/i);
      if (mSum) return decode(mSum[1]).substring(0, 500);
    }

    // 'Call Reason' and 'Call Driver' are the same question — why the customer
    // called — and the evaluation now answers it once in div.ai-meta. Both
    // columns read it from there so they can never disagree.
    if (kw.indexOf('call reason') !== -1 || kw.indexOf('call driver') !== -1) {
      var mReason = html.match(/data-call-reason="([^"]*)"/i);
      if (mReason) {
        var reason = mReason[1].trim();
        // A placeholder the model failed to fill is not an answer.
        if (reason && reason.charAt(0) !== '[' && reason.toLowerCase() !== 'n/a') {
          return reason.substring(0, 200);
        }
      }
      return '';   // never guess this one — a wrong call reason reads as fact
    }

    if (kw.indexOf('opportunit') !== -1) {
      // "What Needs to Change" — the report's own name for the opportunities.
      var mChange = html.match(/<div[^>]*report-col-change[^>]*>([\s\S]*?)<\/ul>/i);
      if (mChange) {
        var lis = joinMatches(mChange[1], /<li[^>]*>[\s\S]*?<\/li>/gi, '; ');
        if (lis) return lis.substring(0, 500);
      }
    }

    if (kw.indexOf('recommendation') !== -1) {
      // The roleplay statements are the actionable recommendation.
      var mRp = html.match(/<ul[^>]*report-col-roleplays[^>]*>([\s\S]*?)<\/ul>/i);
      if (mRp) {
        var stmts = joinMatches(mRp[1], /<div[^>]*ai-flag-stmt[^>]*>[\s\S]*?<\/div>/gi, ' | ');
        if (stmts) return stmts.substring(0, 500);
      }
    }

    if (kw.indexOf('flag') !== -1) {
      var flagTitles = joinMatches(html, /<div[^>]*ai-flag-title[^>]*>[\s\S]*?<\/div>/gi, ' | ');
      if (flagTitles) return flagTitles.substring(0, 500);
    }

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
    // Uses [\s\S]*? so nested badge spans (<span class="report-badge-yes">Yes</span>) are captured
    // and stripped by decode(), returning the clean text value (e.g. "Yes", "72%").
    var re5 = new RegExp('<span[^>]*ai-chip-label[^>]*>[^<]*' + keyword + '[^<]*<\\/span>\\s*<span[^>]*ai-chip-val[^>]*>([\\s\\S]*?)<\\/span>', 'i');
    var m5   = html.match(re5);
    if (m5) return decode(m5[1]).substring(0, 200);

    // 5b. report-summary-meta line: "Duration: X | Direction: Y | Transfer: Z".
    // This is where Transfer, Duration and Direction actually live in the current
    // report — they are not chips, so strategy 5 never matched them and they fell
    // through to strategy 6 below.
    var mMeta = html.match(/<div[^>]*report-summary-meta[^>]*>([\s\S]*?)<\/div>/i);
    if (mMeta) {
      var segs = decode(mMeta[1]).split('|');
      for (var s = 0; s < segs.length; s++) {
        var seg = segs[s].trim();
        var ci  = seg.indexOf(':');
        if (ci > 0 &&
            seg.substring(0, ci).toLowerCase().indexOf(keyword.toLowerCase()) !== -1) {
          return seg.substring(ci + 1).trim().substring(0, 200);
        }
      }
    }

    // 6. Broad keyword search — genuine last resort.
    //
    // Two defects fixed here. It used to run against RAW HTML, so the match
    // window could open or close inside a tag and write fragments like
    // '/div> Karen (daugh' or 'port-header-title">' straight into the sheet.
    // It now runs on decoded text, which cannot contain markup at all.
    //
    // And the report header is dropped first: the first "Transfer" in the
    // document is the title "Repeats & Transfer Audit Report", so every
    // Transfer lookup was matching the heading rather than the data.
    var body  = html.replace(/<div[^>]*class="report-header-[^"]*"[^>]*>[\s\S]*?<\/div>/gi, '');
    var plain = decode(body);
    var reB   = new RegExp('[^.]{0,100}' + keyword + '[^.]{0,200}\\.', 'i');
    var mB    = plain.match(reB);
    if (mB) return mB[0].trim().substring(0, 300);

    return '';
  } catch(e) { Logger.log('extractTextBlock error: ' + e); return ''; }
}

function getSapIdForParticipant(name) { return lookupSapId_(name); }

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
    // The page sends 'sales' or 'repeats'. Anything else would be written to the AI_Analytics sheet as it came (a text that starts with =
    // would run as a formula there), so it is reduced to one of the two.
    analysisType = analysisType === 'sales' ? 'sales' : 'repeats';
    if (!rows || !rows.length) {
      return { success: false, error: 'No data to analyze. Apply filters that return at least 1 record.' };
    }

    // Build a cache key from filters + row count + analysis type
    var cacheKey = Utilities.base64Encode(
      analysisType + '|' + rows.length + '|' + JSON.stringify(filters)
    ).substring(0, 250);

    // Check sheet cache (lasts 24h)
    var cached = readAnalyticsCache_(cacheKey);
    if (cached) {
      Logger.log('AI Analytics cache HIT');
      return { success: true, html: cached.html, fromCache: true, generatedAt: cached.generatedAt };
    }

    Logger.log('AI Analytics cache MISS — calling AI for ' + rows.length + ' rows');

    // Build context summary from rows
    var context = buildAnalyticsContext_(rows, filters, analysisType);

    // Call AI
    var briefingJson = callAIForAnalytics_(context, analysisType);
    if (!briefingJson) return { success: false, error: 'AI did not return a valid response.' };

    // Render to HTML
    var html = renderAnalyticsHTML(briefingJson, filters, analysisType, rows.length);

    // Cache result
    var generatedAt = new Date().toISOString();
    writeAnalyticsCache_(cacheKey, html, generatedAt, rows.length, analysisType);

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

    var cached = readAnalyticsCache_(cacheKey);
    if (cached) {
      try {
        var parsed = JSON.parse(cached.html);
        return { success: true, highlights: parsed.highlights, lowlights: parsed.lowlights,
                 fromCache: true, generatedAt: cached.generatedAt };
      } catch(e) {}
    }

    // The page says which tab it is (filters.analysisType: 'repeats' or 'sales'); the Overview sends none and gets both analyzers together.
    var kind = filters && (filters.analysisType === 'sales' || filters.analysisType === 'repeats') ? filters.analysisType : 'overview';
    var context = buildAnalyticsContext_(rows, filters || {}, kind);
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
    writeAnalyticsCache_(cacheKey,
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
function buildAnalyticsContext_(rows, filters, analysisType) {
  var isSales = analysisType === 'sales';
  var total   = rows.length;

  // Aggregate stats
  // An agent name with stray spaces is the same agent (the dashboard's _dbName), and a cell that holds several people is not one agent.
  var members    = dbUniqServer(rows.map(function(r){ var m = _scName_(r['Team Member']); return (m.length > 0 && m.length <= 80 && String(r['Team Member']).indexOf('\n') === -1) ? m : ''; })).filter(Boolean);
  var flagLabel  = {};
  var rcaCats    = {};
  var callDrivers= {};
  var topOpps    = {};
  var critFlags  = {};
  var repeatPcts = [];
  var resolved   = 0;
  var verdicts   = { yes: 0, partial: 0, no: 0, unknown: 0 };   // the dashboard's rule (_scResolution_), so the AI quotes the same numbers the tiles show

  rows.forEach(function(r) {
    var cat = (r['RCA Category']||'').trim();
    if (cat) rcaCats[cat] = (rcaCats[cat]||0) + 1;

    var drv = (r['Call Driver']||'').trim();
    if (drv) callDrivers[drv] = (callDrivers[drv]||0) + 1;

    var opp = (r['RCA Sub-Parameter']||'').trim();
    if (opp) topOpps[opp] = (topOpps[opp]||0) + 1;

    // The dashboard's rule (_scSplitFlags_): " | " or ", " between flags, a flag title may hold a comma, "None" is no flag, a flag written twice in
    // one audit counts once. The same flag in other capitals or spacing is the same flag (the first spelling is shown).
    _scSplitFlags_(r['Critical Flags']).forEach(function(f){
      var fk = _scKey_(f);
      if (!flagLabel[fk]) flagLabel[fk] = f;
      critFlags[flagLabel[fk]] = (critFlags[flagLabel[fk]]||0) + 1;
    });

    // The dashboard's rule (_scNum_): 0.78 is 78 percent, a sentence that starts with a percentage is read, a blank is not a number, and a real 0 counts.
    var pct = _scNum_(r['Repeat Projection %']);
    if (pct !== null) repeatPcts.push(pct);

    // This used to count any cell that contained the letters yes anywhere ("Partially resolved - yes the credit..."), and missed a sentence
    // that begins with its verdict, so the AI summary said 23 percent while the dashboard tile said 25.
    var verdict = _scResolution_(r['Issue Resolved']);
    verdicts[verdict]++;
    if (verdict === 'yes') resolved++;
  });

  var avgRepeat = repeatPcts.length
    ? Math.round(repeatPcts.reduce(function(s,v){return s+v;},0)/repeatPcts.length) : 0;

  // Find top performer (least critical flags)
  var memberFlagCount = {};
  rows.forEach(function(r){
    var m = _scName_(r['Team Member']);
    memberFlagCount[m] = (memberFlagCount[m]||0) + _scSplitFlags_(r['Critical Flags']).length;
  });
  var topPerformer = members.reduce(function(best, m){
    return (!best || (memberFlagCount[m]||0) < (memberFlagCount[best]||0)) ? m : best;
  }, '');

  // Sort helpers
  var sortDesc = function(obj) {
    return Object.keys(obj).sort(function(a,b){return obj[b]-obj[a];}).map(function(k){return k+' ('+obj[k]+')';});
  };

  // 'sales' and 'repeats' are one tab each; anything else (the Overview) is both analyzers together. It used to say Repeats & Transfer for all.
  var typeLabel = isSales ? 'Sales Analyzer' : analysisType === 'repeats' ? 'Repeats & Transfer Analyzer' : 'All analyzers (Repeats & Transfer and Sales together)';
  var ctx = 'ANALYSIS TYPE: ' + typeLabel + '\n';
  ctx += 'TOTAL INTERACTIONS: ' + total + '\n';
  ctx += 'AGENTS EVALUATED: ' + members.join(', ') + '\n';
  ctx += 'DATE RANGE: ' + (filters.dateFrom||'All time') + ' to ' + (filters.dateTo||'present') + '\n';
  if (filters.members && filters.members.length) ctx += 'FILTERED TEAM MEMBERS: ' + filters.members.join(', ') + '\n';
  if (filters.leaders && filters.leaders.length) ctx += 'TEAM LEADERS: ' + filters.leaders.join(', ') + '\n';
  if (filters.ops && filters.ops.length) ctx += 'OPERATIONS MANAGERS: ' + filters.ops.join(', ') + '\n';
  if (filters.lobs && filters.lobs.length) ctx += 'LINE OF BUSINESS: ' + filters.lobs.join(', ') + '\n';
  if (filters.locales && filters.locales.length) ctx += 'LOCALE/SITE: ' + filters.locales.join(', ') + '\n';
  if (filters.vtids && filters.vtids.length) ctx += 'VTID: ' + filters.vtids.join(', ') + '\n';
  if (filters.qas && filters.qas.length) ctx += 'QA OBSERVERS: ' + filters.qas.join(', ') + '\n';
  if (filters.sapQ) ctx += 'SAP ID SEARCH: ' + filters.sapQ + '\n';
  ctx += '\nKEY METRICS:\n';
  ctx += '- Issues Resolved: ' + resolved + ' of ' + total + ' (' + Math.round(resolved/total*100) + '%)\n';
  ctx += '- Issue Resolved, all answers: Resolved ' + verdicts.yes + ', Partially resolved ' + verdicts.partial + ', Unresolved ' + verdicts.no + ', Not recorded ' + verdicts.unknown + '\n';
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
function callAIForAnalytics_(context, analysisType) {
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
  // `filters` comes from the caller's browser and this HTML is saved to the
  // analytics cache and emailed, so every value is escaped with e() — an
  // unescaped member name was a stored-HTML injection.
  filters = filters || {};
  var fList = function(a) { return (Array.isArray(a) ? a : []).map(function(x) { return e(x); }).join(', '); };
  var filterSummary = [];
  if (filters.dateFrom || filters.dateTo) filterSummary.push(e(filters.dateFrom||'*') + ' → ' + e(filters.dateTo||'*'));
  if (Array.isArray(filters.members) && filters.members.length) filterSummary.push('Agent: '  + fList(filters.members));
  if (Array.isArray(filters.leaders) && filters.leaders.length) filterSummary.push('TL: '     + fList(filters.leaders));
  if (Array.isArray(filters.lobs)    && filters.lobs.length)    filterSummary.push('LOB: '    + fList(filters.lobs));
  if (Array.isArray(filters.locales) && filters.locales.length) filterSummary.push('Locale: ' + fList(filters.locales));
  if (Array.isArray(filters.vtids)   && filters.vtids.length)   filterSummary.push('VTID: '   + fList(filters.vtids));
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
        html += '<div class="r-tts-stmt-text" style="background:#EDF7E6;border:1px solid #B3DFA0;border-radius:4px;padding:8px 10px;font-size:12px;color:#1A5E00;font-style:italic">' +
          '<button class="tts-play-btn" data-tts="' + esc(r.positioningStatement) + '" contenteditable="false" title="Play sample statement" style="background:#22c55e;border:none;border-radius:50%;width:20px;height:20px;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;padding:0;vertical-align:middle;margin-right:5px;flex-shrink:0"><svg width="9" height="10" viewBox="0 0 9 10" fill="white"><polygon points="1,1 8,5 1,9"/></svg></button>' +
          e(r.positioningStatement) + '</div>';
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
function readAnalyticsCache_(key) {
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
function writeAnalyticsCache_(key, html, generatedAt, rowCount, atype) {
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
    appendRow(sheet, [key, generatedAt, html, rowCount, atype]);   // the wrapper makes a text that starts with = + - @ plain text, not a formula
    Logger.log('AI Analytics cached: ' + atype + ' (' + rowCount + ' rows)');
  } catch(e) { Logger.log('Analytics cache write error: ' + e); }
}

// A Sheets cell holds at most 50,000 characters and the AI accepts 150,000 (TRANSCRIPT_MAX_CHARS), so a long call made the save of the
// transcript row throw and the whole submit fail with a raw Google error. The saved copy is cut with a note; the AI still gets the whole text.
function _cellSafeTranscript_(t) {
  var max = 49000;   // under the 50,000 of a cell, with room for the note
  var s = String(t === null || t === undefined ? '' : t);
  if (s.length <= max) return s;
  var note = '\n[saved copy cut here: the call text was ' + s.length + ' characters and a sheet cell holds 50000]';
  return s.substring(0, max - note.length) + note;
}

var BRIEFING_EMAIL_DOMAINS = ['telus.com', 'telusinternational.com'];

// Exactly one address (no list separators, whitespace or angle brackets) whose
// domain is on BRIEFING_EMAIL_DOMAINS. Subdomains do not count.
function isAllowedBriefingRecipient_(addr) {
  if (typeof addr !== 'string' || addr.length > 254) return false;
  var m = /^[A-Za-z0-9._%+'-]+@([A-Za-z0-9.-]+)$/.exec(addr);
  return !!m && BRIEFING_EMAIL_DOMAINS.indexOf(m[1].toLowerCase()) !== -1;
}

/** Send analytics briefing email */
function sendAnalyticsBriefing(toEmail, analyticsHTML, filters, analysisType, rowCount) {
  try {
    // This function is callable from any signed-in user's browser and sends from
    // the script owner's mailbox, so it must not be an open relay: one address,
    // on a TELUS domain, and the body reduced to the allow-list.
    var to = (toEmail || '').toString().trim();
    if (!isAllowedBriefingRecipient_(to)) {
      Logger.log('sendAnalyticsBriefing refused recipient: ' + to);
      return { success: false, error: 'Briefings can only be sent to one @telus.com or @telusinternational.com address.' };
    }
    toEmail = to;
    // Sent from the owner's mailbox, so it counts against the same hourly limit as the audit emails (sendAuditEmail).
    var _briefCaller = ''; try { _briefCaller = (Session.getActiveUser().getEmail() || '').toString().toLowerCase().trim(); } catch (ce) {}
    if (!_emailRateTake_(_briefCaller, isAdminCaller_())) {
      return { success: false, error: 'You have already sent a lot of emails in the last hour, so nothing was sent. Please try again later.' };
    }
    analyticsHTML = sanitizeEvalHtml_(analyticsHTML);
    rowCount = Number(rowCount) || 0;
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
    // Client-supplied, so sanitised (see sendAuditEmail). The sanitiser drops every
    // <style>, and the PDF used to get its styling from the one the browser sent
    // along, so the server puts its OWN stylesheet back rather than trusting theirs.
    htmlResult = sharedCSS() + sanitizeEvalHtml_(htmlResult);
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

    // Save PDF to a named folder and lazy-clean files older than 2 hours
    var folderName = 'NH Analyzer PDFs (auto-cleanup)';
    var folder     = _ownedFolder_(folderName);

    // Lazy cleanup: trash PDF files in this folder older than 2 hours
    try {
      var cutoff = new Date(Date.now() - 2 * 60 * 60 * 1000);
      var oldFiles = folder.getFiles();
      while (oldFiles.hasNext()) {
        var f = oldFiles.next();
        if (f.getDateCreated() < cutoff) { try { f.setTrashed(true); } catch(te) {} }
      }
    } catch(ce) { Logger.log('PDF cleanup error (non-fatal): ' + ce); }

    var pdfFile   = folder.createFile(pdf);
    var sharingOk = true;
    try { pdfFile.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW); }
    catch(se) { sharingOk = false; Logger.log('PDF setSharing failed: ' + se); }

    return { success: true, url: pdfFile.getDownloadUrl() || pdfFile.getUrl(), name: pdfFile.getName(), sharingFailed: !sharingOk };
  } catch(e) {
    Logger.log('generateAndServePDF error: ' + e);
    return { success: false, error: e.toString() };
  }
}

function _generateTempRef() {
  var now = new Date();
  return getAuditRefPrefix_() + '-' + now.getFullYear() + String(now.getMonth()+1).padStart(2,'0') + String(now.getDate()).padStart(2,'0') + '-TEMP';
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
          return getEvaluationFromCache_(auditRef);
        }
        return { html: sharedCSS() + fixBadgeClasses(stored), css: '' };
      }
    }
    // Not in Analysis sheet — try cache
    return getEvaluationFromCache_(auditRef);
  } catch(e) {
    Logger.log('getEvaluationByAuditRef error: ' + e);
    return null;
  }
}

function getEvaluationFromCache_(auditRef) {
  try {
    var ss       = getOrCreateSpreadsheet();
    var logSheet = getOrCreateSheet(ss, AUDIT_LOG_SHEET);
    var logLast  = logSheet.getLastRow();
    if (logLast < 2) return null;

    // Resolve columns by header name. Column B is 'Submitted At', not the
    // analysis type, so reading it positionally never matched a cache row type.
    var headers = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
    var refCol  = headers.indexOf('Audit Ref');
    var idCol   = headers.indexOf('Interaction ID');
    var sapCol  = headers.indexOf('SAP ID');
    var typeCol = headers.indexOf('Analysis Type');
    if (refCol < 0 || idCol < 0) return null;

    var logData = logSheet.getRange(2, 1, logLast - 1, headers.length).getValues();
    var interactionId = '';
    var analysisType  = '';
    var sapId         = '';
    for (var i = 0; i < logData.length; i++) {
      if ((logData[i][refCol] || '').toString().trim() === auditRef.trim()) {
        interactionId = (logData[i][idCol] || '').toString().trim();
        analysisType  = typeCol >= 0 ? _normalizeAnalysisType(logData[i][typeCol]) : 'repeats';
        sapId         = sapCol  >= 0 ? (logData[i][sapCol] || '').toString().trim() : '';
        break;
      }
    }
    if (!interactionId) return null;

    // Keyed by Interaction ID + analysis type + THIS audit's agent, so two
    // agents on the same call each get their own evaluation back.
    var html = findCachedResult_(interactionId, analysisType, sapId);
    if (!html) return null;

    return { html: sharedCSS() + fixBadgeClasses(html), css: '' };
  } catch(e) {
    Logger.log('getEvaluationFromCache_ error: ' + e);
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
function getTranscriptMetadata_(txt)   { return parseTranscriptMetadata(txt); }
// A lookup that cannot complete must not reach the analyst as a red banner.
// One of the roster sheets is shared with the telus.com domain only, while the
// TI deployment runs as a @telusinternational.com account, so opening it throws
// "You do not have permission to access the requested document." That surfaced
// to users mid-typing — and often for a partial SAP ID they had not finished
// entering. Returning null leaves the fields blank, which is what already
// happens for an agent who genuinely is not on the roster.
//
// But null alone cannot tell "this agent is not on the roster" from "the roster
// could not be read", so the page showed nothing at all in both cases. A miss
// that coincides with an unreadable source now returns { _error, ... } so the
// page can say which. A genuine miss is still null, and a hit is still the row.
function getRosterBySapId(sapId) {
  _resetSourceErrors_();
  try {
    var row = lookupBySapId_(sapId);
    if (row) return row;
    if (_lookupSourceErrors.length) {
      return { _error: 'source_unavailable', sources: _lookupSourceErrors.slice() };
    }
    return null;
  } catch(e) {
    Logger.log('getRosterBySapId failed for "' + sapId + '": ' + e);
    return { _error: 'lookup_failed' };
  }
}

// ── Combined single-call autofill: parses transcript + resolves SAP ID + returns full roster ──
// Eliminates 3 sequential client→server round-trips by doing all lookups in one call.
function getTranscriptMetadataAndRoster(txt) {
  try {
    var meta = parseTranscriptMetadata(txt);
    var sapId = '';
    var roster = null;
    if (meta && meta.participant) {
      sapId = lookupSapId_(meta.participant) || '';
      if (sapId) roster = lookupBySapId_(sapId);
    }
    return { meta: meta || {}, sapId: sapId, roster: roster };
  } catch(e) {
    Logger.log('getTranscriptMetadataAndRoster error: ' + e);
    return { meta: {}, sapId: '', roster: null };
  }
}
function getRosterAll()               { return getAllRosterData_(); }
function getObserverInfo() {
  var obs = resolveObserver_() || {};

  // Get email — try both Session methods independently
  var userEmail = '';
  try { userEmail = Session.getActiveUser().getEmail()  || ''; } catch(e) {}
  try { if (!userEmail) userEmail = Session.getEffectiveUser().getEmail() || ''; } catch(e) {}
  userEmail = userEmail.toLowerCase().trim();

  // Always populate email/name so Observer display works
  if (!obs.email && userEmail) obs.email = userEmail;
  if (!obs.name  && userEmail) obs.name  = userEmail.split('@')[0];

  // Admin check — the SAME rule sendAuditEmail enforces (isAdminCaller_): the person
  // actually running this call, on ADMIN_USERNAMES, on a TELUS domain. The display
  // fallback above to the effective user is only for showing a name; it must not
  // decide who sees the TEST toggle. In "Execute as: Me" the effective user is the
  // OWNER, so a caller whose own email is hidden would otherwise look like an admin.
  obs.isAdmin = isAdminCaller_();

  return obs;
}
function getAuditLogData()            { return readAuditLog_(); }

// ── Generate unique Audit Reference Number: <PREFIX>-YYYYMMDD-XXXX ──────────────
// Uses PropertiesService to persist today's counter — avoids scanning the full
// sheet. PropertiesService has no expiry, unlike CacheService (hard 6-hour cap
// on put()'s TTL) — a per-day counter needs to survive up to 24 hours, so
// CacheService can't hold it without throwing and silently corrupting refs.
function generateAuditRef_() {
  var now      = new Date();
  var y        = now.getFullYear();
  var m        = String(now.getMonth() + 1).padStart(2, '0');
  var d        = String(now.getDate()).padStart(2, '0');
  var dateStr  = '' + y + m + d;
  var propKey  = 'audit_ref_seq_' + dateStr;

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    var props   = PropertiesService.getScriptProperties();
    var current = parseInt(props.getProperty(propKey) || '0', 10);
    var seq     = current + 1;
    props.setProperty(propKey, String(seq));
    return getAuditRefPrefix_() + '-' + dateStr + '-' + String(seq).padStart(4, '0');
  } catch(e) {
    Logger.log('generateAuditRef_ error: ' + e);
    return _uniqueFallbackAuditRef(dateStr);
  } finally {
    try { lock.releaseLock(); } catch(re) {}
  }
}

// Last-resort ref for when the primary counter path above throws (e.g. a
// LockService timeout under heavy concurrent submissions). A bare random
// guess with no collision check is exactly what caused the original
// same-day audit-ref collision bug, so this checks Audit_Log before handing
// one out, and only falls back to a millisecond-based suffix (which cannot
// collide with the padded 4-digit sequence/random refs) if every guess
// collides or the sheet check itself fails.
function _uniqueFallbackAuditRef(dateStr) {
  try {
    var ss       = getOrCreateSpreadsheet();
    var logSheet = getOrCreateSheet(ss, AUDIT_LOG_SHEET);
    for (var attempt = 0; attempt < 10; attempt++) {
      var candidate = getAuditRefPrefix_() + '-' + dateStr + '-' + String(Math.floor(Math.random() * 9999)).padStart(4, '0');
      var finder    = logSheet.getRange('A:A').createTextFinder(candidate).matchEntireCell(true);
      if (!finder.findNext()) return candidate;
    }
  } catch(e2) {
    Logger.log('_uniqueFallbackAuditRef error: ' + e2);
  }
  return getAuditRefPrefix_() + '-' + dateStr + '-T' + (Date.now() % 100000);
}

// ── Resolve observer from logged-in user email ────────────────────────────────
// Priority: 1) Users sheet  2) Roster sheet  3) email prefix fallback
// Returns { name, role, sapId, email }
function resolveObserver_() {
  try {
    var email = '';
    try { email = Session.getActiveUser().getEmail() || ''; } catch(e) {}
    if (!email) return { name: '', role: '', sapId: '', email: '' };

    // Cache observer result per email for 8 hours
    var cache    = CacheService.getUserCache();
    var cacheKey = 'observer_' + email.replace(/[^a-zA-Z0-9]/g, '_');
    var cached   = cache.get(cacheKey);
    if (cached) { try { return JSON.parse(cached); } catch(e) {} }

    // ── 1. Users sheet (canonical source — Email Address, Name, Role) ─────────
    var userRow = lookupUserFromUsersSheet_(email);
    if (userRow && userRow.name) {
      var result = { name: userRow.name, role: userRow.role, sapId: '', email: email };
      try { cache.put(cacheKey, JSON.stringify(result), 8 * 60 * 60); } catch(e) {}
      return result;
    }

    // ── 2. Roster sheet fallback ──────────────────────────────────────────────
    var data       = _getRosterSheetData();  // uses cached roster
    var emailLower = email.toLowerCase();

    var result = null;
    for (var i = 0; i < data.length; i++) {
      // Search email columns
      for (var j = 0; j < data[i].length; j++) {
        if (data[i][j] && data[i][j].toString().toLowerCase() === emailLower) {
          result = {
            name:  (data[i][ROSTER_COL_AGENT_NAME] || '').toString().trim(),
            role:  '',
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
            result = { name: agentName, role: '', sapId: (data[i][ROSTER_COL_SAP_ID] || '').toString().trim(), email: email };
            break;
          }
        }
      }
    }
    if (!result) result = { name: email.split('@')[0], role: '', sapId: '', email: email };

    // Cache for 8 hours
    try { cache.put(cacheKey, JSON.stringify(result), 8 * 60 * 60); } catch(e) {}
    return result;
  } catch(e) {
    Logger.log('resolveObserver_ error: ' + e);
    return { name: '', role: '', sapId: '', email: '' };
  }
}

// ── Dynamic flag dropdown options from Google Sheet ───────────────────────────
// Reads "Dropdowns" column from "Add a Flag" tab — never hardcoded so future
// additions to the sheet appear automatically without any code change.
function getFlagDropdownOptions() {
  try {
    var ss    = SpreadsheetApp.openById('1baZbWTLNw7_ij04oH3aO_xT4eOTffcUwRRL836cHzJo');
    var sheet = ss.getSheetByName('Add a Flag');
    if (!sheet) return { success: false, options: [], error: 'Tab "Add a Flag" not found' };
    var data    = sheet.getDataRange().getValues();
    if (!data || !data.length) return { success: true, options: [] };
    var headers = data[0];
    var colIdx  = -1;
    for (var i = 0; i < headers.length; i++) {
      if ((headers[i] + '').trim().toLowerCase() === 'dropdowns') { colIdx = i; break; }
    }
    if (colIdx === -1) return { success: false, options: [], error: 'Column "Dropdowns" not found' };
    var options = [];
    for (var r = 1; r < data.length; r++) {
      var val = (data[r][colIdx] + '').trim();
      if (val) options.push(val);
    }
    return { success: true, options: options };
  } catch(e) {
    Logger.log('getFlagDropdownOptions error: ' + e);
    return { success: false, options: [], error: e.toString() };
  }
}

// ── Read Audit_Log for dashboard tab 3 (cached 10 min) ───────────────────────
var AUDIT_LOG_CACHE_KEY = 'audit_log_data_v2';
var AUDIT_LOG_CACHE_TTL = 20 * 60; // 20 minutes — localStorage handles the faster layer

function readAuditLog_() {
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

    // Read by actual header names so column order in the sheet never matters.
    // Mirrors readDashboardSheet_ — fixes silent misreads when columns were added.
    var lastCol      = sheet.getLastColumn();
    var actualHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var data   = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var result = data.map(function(row) {
      var obj = {};
      actualHeaders.forEach(function(h, i) {
        if (h) obj[h] = row[i] !== undefined ? row[i].toString() : '';
      });
      return obj;
    });

    var _al = JSON.stringify(result);
    if (_al.length <= 99000) { try { cache.put(AUDIT_LOG_CACHE_KEY, _al, AUDIT_LOG_CACHE_TTL); } catch(e) {} }
    return result;
  } catch(e) { Logger.log('readAuditLog_: ' + e); return []; }
}

function invalidateAuditLogCache_() {
  try { CacheService.getScriptCache().remove(AUDIT_LOG_CACHE_KEY); } catch(e) {}
}

// Dashboard entry point — served at ?page=dashboard
function doGetDashboard_() {
  return HtmlService
    .createHtmlOutputFromFile('Dashboard')
    .setTitle('NH Call Analyzer — Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Dashboard data cache keys ─────────────────────────────────────────────────
var DB_CACHE_KEY      = 'dashboard_data_v3';       // v3: v2 could be left half-written
var DB_CACHE_META_KEY = 'dashboard_data_meta_v3';
var DB_CACHE_TTL      = 60 * 60;   // 60 minutes — extended since localStorage is the fast layer
// CacheService limit is 100 KB per key, counted in BYTES. Accented and non-Latin
// text is up to 3 bytes a character, so the chunk is 30,000 CHARACTERS (at most
// ~90 KB). It used to be 90,000 characters, which overflowed as soon as the
// dashboard held enough such text, and the failure blanked the whole dashboard.
var DB_CACHE_CHUNK    = 30000;

// ── Read from sheet and return raw rows ───────────────────────────────────────
function readDashboardSheet_() {
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
function writeDashboardCache_(rows) {
  var cache   = CacheService.getScriptCache();
  var json    = JSON.stringify(rows);
  var chunks  = [];
  for (var i = 0; i < json.length; i += DB_CACHE_CHUNK) {
    chunks.push(json.slice(i, i + DB_CACHE_CHUNK));
  }
  var meta = { chunks: chunks.length, ts: new Date().toISOString(), rowCount: rows.length };
  // The cache is an optimisation. A failure to fill it must NEVER take the dashboard
  // down with it: the callers have already read the sheet successfully, and an
  // exception here used to make them return an empty dashboard. The meta entry goes
  // in LAST, so its presence means every chunk was stored; on any failure what was
  // written is removed and the next load simply reads the sheet again.
  try {
    chunks.forEach(function(chunk, idx) {
      cache.put(DB_CACHE_KEY + '_' + idx, chunk, DB_CACHE_TTL);
    });
    cache.put(DB_CACHE_META_KEY, JSON.stringify(meta), DB_CACHE_TTL);
    Logger.log('Dashboard cache written: ' + rows.length + ' rows, ' + chunks.length + ' chunk(s)');
  } catch (e) {
    Logger.log('Dashboard cache write failed (served from the sheet instead): ' + e);
    try {
      var keys = [DB_CACHE_META_KEY];
      for (var k = 0; k < chunks.length; k++) keys.push(DB_CACHE_KEY + '_' + k);
      cache.removeAll(keys);
    } catch (e2) {}
  }
}

// ── Read rows from CacheService ───────────────────────────────────────────────
function readDashboardCache_() {
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
    Logger.log('readDashboardCache_ parse error: ' + e);
    return null;
  }
}

// ── Invalidate dashboard cache (call after new data is written) ───────────────
function invalidateDashboardCache_() {
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
    var cached = readDashboardCache_();
    if (cached) {
      Logger.log('Dashboard cache HIT — ' + cached.rowCount + ' rows from ' + cached.ts);
      return { rows: cached.rows, ts: cached.ts, rowCount: cached.rowCount, fromCache: true };
    }
    Logger.log('Dashboard cache MISS — reading sheet');
    var rows = readDashboardSheet_();
    writeDashboardCache_(rows);
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
    invalidateDashboardCache_();
    var rows = readDashboardSheet_();
    writeDashboardCache_(rows);
    var ts = new Date().toISOString();
    Logger.log('Dashboard force-refreshed: ' + rows.length + ' rows');
    return { rows: rows, ts: ts, rowCount: rows.length, fromCache: false };
  } catch(e) {
    Logger.log('getDashboardDataFresh error: ' + e);
    return { rows: [], ts: '', rowCount: 0, fromCache: false };
  }
}

// ── Look up the audit ref of THIS agent's audit of a cached interaction ────────
// Several agents can share one Interaction ID, so the agent (SAP ID) and the
// analysis type are part of the match. Returning the first row for the
// Interaction ID handed the second agent the first agent's ref.
function findAuditRefForInteraction_(interactionId, analysisType, sapId) {
  return _findAuditRefForAgent(interactionId, analysisType, sapId);
}

// Index of the sheet row that belongs to THIS audit, or -1.
// A reference should be unique, but rows written before the NH and TI scripts
// got separate prefixes can share one. Prefer the row whose reference AND
// Interaction ID both match; otherwise fall back to the first row with the
// reference (the old behaviour). A blank or 'N/A' Interaction ID is not a claim.
// refVals / idVals are the sheet's Audit Ref and Interaction ID columns, as
// returned by getRange(...).getValues().
function _findAuditRowIndex(refVals, idVals, wantRef, wantInteractionId) {
  var ref = (wantRef || '').toString().trim();
  var iid = (wantInteractionId || '').toString().trim();
  if (iid.toUpperCase() === 'N/A') iid = '';
  var firstByRef = -1;
  for (var i = 0; i < refVals.length; i++) {
    if ((refVals[i][0] || '').toString().trim() !== ref) continue;
    if (firstByRef < 0) firstByRef = i;
    if (iid && idVals && (idVals[i][0] || '').toString().trim() === iid) return i;
  }
  return firstByRef;
}

// Called after email sent — update Dashboard_Data and Audit_Log
// emailStatus goes into the 'Email Status' column verbatim. Callers MUST pass the
// literal 'Sent' on success — nine places in index.html and Dashboard.html gate on
// `=== 'Sent'` — and a descriptive reason on failure.
//
// This used to take the outcome in a `pdfLink` argument and write it to the
// 'PDF Email Link' column, while forcing 'Email Status' to the literal 'Sent'
// unconditionally. So a send that failed or hit quota still reported as Sent, and
// the real outcome sat in the wrong column. No PDF link is produced on this path
// (generateAndServePDF is a separate, on-demand flow that writes nothing here), so
// the PDF column is left alone rather than filled with a status string.
//
// issueResolved (optional): the analyst's own Yes/No from the report badge, which
// they can now toggle. The sheet row is written during submitTranscript, BEFORE
// the analyst ever sees the report, so an override made afterwards has to be
// written back — and this is the one call that already reaches both sheets on
// send, so it costs no extra execution. Omitted or blank leaves the value alone.
function updateDashboardPDFLink_(interactionId, emailStatus, recipients, auditRef, issueResolved) {
  try {
    var ss = getOrCreateSpreadsheet();
    // Everything written below goes in with setValue, which turns a string that starts with = into a FORMULA that runs with the
    // owner's rights. The analyst's override comes from the report HTML the browser sent, so it is reduced to the three answers the
    // badge can show, and every other string is stored as text.
    var safe = function(v) { var s = String(v === null || v === undefined ? '' : v); return /^[=+\-@\t\r]/.test(s) ? "'" + s : s; };
    var status           = safe((emailStatus || 'Sent').toString().trim().substring(0, 40));
    var resolvedOverride = (function(v) {
      var m = /^(yes|no|partial|partially|partially resolved)$/i.exec(String(v || '').trim());
      return m ? (/^y/i.test(m[1]) ? 'Yes' : (/^n/i.test(m[1]) ? 'No' : 'Partially resolved')) : '';
    })(issueResolved);
    // Match this audit's OWN row by Audit Ref. Several agents can share one
    // Interaction ID, and matching on it alone wrote the status, recipients and
    // Issue Resolved override onto the first agent's row.
    var wantRef          = (auditRef || '').toString().trim();

    // Update Dashboard_Data — col C = Interaction ID (index 3), col W=PDF, col X=Status
    var dSheet  = getOrCreateSheet(ss, DASHBOARD_DATA_SHEET);
    var lastRow = dSheet.getLastRow();
    if (lastRow >= 2) {
      // Resolve column positions by header name — never hardcode
      var dHeaderRow = dSheet.getRange(1, 1, 1, dSheet.getLastColumn()).getValues()[0];
      var dStatusCol= dHeaderRow.indexOf('Email Status')   + 1;
      var dIdCol    = dHeaderRow.indexOf('Interaction ID') + 1;
      var dRefCol   = dHeaderRow.indexOf('Audit Ref')      + 1;
      var dResCol   = dHeaderRow.indexOf('Issue Resolved') + 1;
      if (dIdCol < 1) dIdCol = 3; // safe fallback to col C
      var dByRef    = !!wantRef && dRefCol >= 1;
      var ids = dSheet.getRange(2, dByRef ? dRefCol : dIdCol, lastRow - 1, 1).getValues();
      var dPick = dByRef
        ? _findAuditRowIndex(ids, dSheet.getRange(2, dIdCol, lastRow - 1, 1).getValues(), wantRef, interactionId)
        : -1;
      for (var i = 0; i < ids.length; i++) {
        if (dByRef ? i === dPick : ids[i][0].toString().trim() === interactionId.trim()) {
          if (dStatusCol > 0) dSheet.getRange(i + 2, dStatusCol).setValue(status);
          if (dResCol > 0 && resolvedOverride)
            dSheet.getRange(i + 2, dResCol).setValue(resolvedOverride);
          break;
        }
      }
    }

    // Update Audit_Log — resolve columns by header name
    var logSheet  = getOrCreateSheet(ss, AUDIT_LOG_SHEET);
    var logLast   = logSheet.getLastRow();
    if (logLast >= 2) {
      var logHeaderRow  = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
      var logStatusCol  = logHeaderRow.indexOf('Email Status')   + 1;
      var logRecipCol   = logHeaderRow.indexOf('Recipients')     + 1;
      var logIdCol      = logHeaderRow.indexOf('Interaction ID') + 1;
      var logRefCol     = logHeaderRow.indexOf('Audit Ref')      + 1;
      var logResCol     = logHeaderRow.indexOf('Issue Resolved') + 1;
      if (logIdCol < 1) logIdCol = 3; // safe fallback
      var logByRef      = !!wantRef && logRefCol >= 1;
      var logIds = logSheet.getRange(2, logByRef ? logRefCol : logIdCol, logLast - 1, 1).getValues();
      var logPick = logByRef
        ? _findAuditRowIndex(logIds, logSheet.getRange(2, logIdCol, logLast - 1, 1).getValues(), wantRef, interactionId)
        : -1;
      for (var j = 0; j < logIds.length; j++) {
        if (logByRef ? j === logPick : logIds[j][0].toString().trim() === interactionId.trim()) {
          if (logStatusCol > 0) logSheet.getRange(j + 2, logStatusCol).setValue(status);
          if (logRecipCol  > 0) logSheet.getRange(j + 2, logRecipCol).setValue(safe((recipients || []).join(', ')));
          if (logResCol > 0 && resolvedOverride)
            logSheet.getRange(j + 2, logResCol).setValue(resolvedOverride);
          break;
        }
      }
    }

    invalidateDashboardCache_();
    invalidateAuditLogCache_();   // Audit_Log holds the Email Status and Recipients too; its cache showed Not Sent for 20 minutes
  } catch(e) { Logger.log('updateDashboardPDFLink_: ' + e); }
}

// ─────────────────────────────────────────────────────────────────────────────
// Pre-analysis: detect speakers in a transcript before submitting for AI.
// Called by the client to power the speaker-confirmation modal (P1/P4).
// Returns:
//   { hasHeader, headerAgents, detectedSpeakers, targetName, targetMatch,
//     targetFound, otherAgents }
// ─────────────────────────────────────────────────────────────────────────────
function detectTranscriptSpeakers(transcriptText, sapId) {
  try {
    if (!transcriptText) return { error: 'No transcript provided.' };

    // Resolve target agent name from SAP ID (or use raw sapId as fallback label)
    var targetName = '';
    if (sapId) {
      try {
        var looked = lookupBySapId_(sapId);
        targetName = (looked && looked.participant) ? looked.participant : sapId;
      } catch(e) { targetName = sapId; }
    }
    var targetLower = targetName.toLowerCase().trim();

    // ── Parse Internal Participant(s) header ──────────────────────────────────
    var hasHeader    = false;
    var headerAgents = [];
    var pMatch = transcriptText.match(/Internal Participant\(s\)[:\s]+([^\n\r]+)/i);
    if (pMatch) {
      hasHeader    = true;
      headerAgents = pMatch[1].split(/[,;]+/)
        .map(function(n) { return n.trim(); })
        .filter(function(n) { return n.length > 0; });
    }

    // ── Pre-scan: detect all real speakers (≥2 turns) ────────────────────────
    // Count case-insensitively to handle inconsistent transcript capitalisation;
    // preserve first-seen original casing for display.
    var lines     = transcriptText.split('\n');
    var counts    = {};
    var firstSeen = {};
    for (var i = 0; i < lines.length; i++) {
      var m = lines[i].match(/^([^:\n]{2,60}):\s/);
      if (m) {
        var spOrig  = m[1].trim();
        var spLower = spOrig.toLowerCase();
        counts[spLower]    = (counts[spLower] || 0) + 1;
        if (!firstSeen[spLower]) firstSeen[spLower] = spOrig;
      }
    }
    var detectedSpeakers = Object.keys(counts).filter(function(k) {
      return counts[k] >= 2;
    }).map(function(k) { return firstSeen[k]; });
    // Also include any header agent names not caught by pre-scan
    headerAgents.forEach(function(h) {
      var found = detectedSpeakers.some(function(sp) {
        return sp.toLowerCase() === h.toLowerCase() ||
               sp.toLowerCase().indexOf(h.toLowerCase()) !== -1 ||
               h.toLowerCase().indexOf(sp.toLowerCase()) !== -1;
      });
      if (!found) detectedSpeakers.push(h);
    });

    // ── Match target against detected speakers ────────────────────────────────
    var targetMatch = null;
    if (targetLower) {
      for (var j = 0; j < detectedSpeakers.length; j++) {
        var spLower = detectedSpeakers[j].toLowerCase();
        if (spLower === targetLower ||
            (targetLower.length > 4 && spLower.indexOf(targetLower) !== -1) ||
            (spLower.length > 4 && targetLower.indexOf(spLower) !== -1)) {
          targetMatch = detectedSpeakers[j];
          break;
        }
      }
    }

    // ── Other agents = header agents minus the target ─────────────────────────
    // Only identify "other agents" when a header is present — without it we
    // cannot reliably distinguish customer from another internal agent, so we
    // return an empty list. filterTranscriptByAgent also returns NO_FILTER when
    // there is no header, keeping both the UI and the server in sync.
    var otherAgents = [];
    if (hasHeader) {
      otherAgents = detectedSpeakers.filter(function(sp) {
        var spLower = sp.toLowerCase();
        var isTarget = targetMatch && (spLower === targetMatch.toLowerCase() ||
                        spLower.indexOf(targetMatch.toLowerCase()) !== -1 ||
                        targetMatch.toLowerCase().indexOf(spLower) !== -1);
        if (isTarget) return false;
        return headerAgents.some(function(h) {
          var hLow = h.toLowerCase();
          return spLower === hLow ||
                 (hLow.length > 4 && spLower.indexOf(hLow) !== -1) ||
                 (spLow.length > 4 && hLow.indexOf(spLow) !== -1);
        });
      });
    }

    return {
      hasHeader:        hasHeader,
      headerAgents:     headerAgents,
      detectedSpeakers: detectedSpeakers,
      targetName:       targetName,
      targetMatch:      targetMatch,
      targetFound:      !!targetMatch,
      otherAgents:      otherAgents
    };
  } catch(e) {
    Logger.log('detectTranscriptSpeakers error: ' + e);
    return { error: e.toString() };
  }
}

// Key of the "this audit is being processed" flag. One per Interaction ID AND
// agent, so a second agent on the same call is not blocked by the first. The
// acquire and both release paths in submitTranscript must build it the same way.
function _procKey_(interactionId, sapId) {
  return 'proc_' + (interactionId || '').toString().trim() + '_' + (sapId || '').toString().trim();
}

// Translate (if asked), narrow to the agent, run the AI, tidy and sanitise the output and
// put it in the cache. Throws on an AI failure; submitTranscript turns that into its
// friendly error. Returns { html, filterStats }.
function produceEvaluation_(formData, p) {
  var interactionId = p.interactionId, participant = p.participant, analysisType = p.analysisType,
      auditRef = p.auditRef, observerName = p.observerName;
  // ── 3.5 Translate transcript to English if requested ─────────────────────
  var transcriptForAnalysis = formData.transcript;
  if (formData.translateToEnglish && transcriptForAnalysis.trim()) {
    try {
      Logger.log('translateToEnglish=true — translating transcript for ' + (interactionId || 'unknown'));
      var translated = translateTranscriptText_(transcriptForAnalysis);
      if (translated) {
        transcriptForAnalysis = translated;
      } else {
        Logger.log('Translation returned null — proceeding with original transcript');
      }
    } catch(te) {
      Logger.log('Translation step failed (non-fatal, using original): ' + te);
    }
  }

  // ── 4. Run AI ─────────────────────────────────────────────────────────────
  // Filter transcript to only include the target agent's turns + customer lines
  var filterStats = { applied: false, targetFound: false };
  if (participant) {
    var filterResult = filterTranscriptByAgent(transcriptForAnalysis, participant);
    transcriptForAnalysis = filterResult.filtered;
    filterStats = filterResult.stats;
  }
  var rawAI = analyzeTranscript_(transcriptForAnalysis, analysisType, participant, formData.selectedLOB || '');
  Logger.log('AI response length: ' + rawAI.length);

  var html = fixBadgeClasses(
    rawAI
      .replace(/^```html\s*/i, '')
      .replace(/^```\s*/,      '')
      .replace(/\s*```$/,      '')
      .trim()
  );

  // Replace report placeholders with actual values so cache, email, and
  // Dashboard all receive the final content (no deferred substitution needed).
  html = html
    .replace(/__AUDIT_REF__/g, auditRef     || '')
    .replace(/__OBSERVER__/g,  observerName || 'N/A');

  // The model's output includes whatever was said on the call, and a transcript
  // can ask it to emit markup that runs. Reduce it to the allow-list BEFORE it is
  // cached, written to a sheet, emailed or shown. (After the placeholders, so a
  // name substituted in is covered too.)
  html = sanitizeEvalHtml_(html);

  // ── 5. Save to Cache sheet FIRST — so a retry after any downstream failure ──
  //      hits the cache instead of re-running the paid AI call and duplicating rows.
  if (interactionId) {
    saveCachedResult_(interactionId, analysisType, html, formData.sapId);
  }
  return { html: html, filterStats: filterStats };
}

// Writes the Analysis, Dashboard_Data and Audit_Log rows of one audit and returns the
// warnings. Each write is independent: one failing never stops the others. When
// `f.recovering` is set (completing an audit whose report was already cached) a row that
// is already there is not written a second time.
function writeAuditRecords_(spreadsheet, formData, f) {
  var auditRef = f.auditRef, now = f.now, participant = f.participant, interactionId = f.interactionId,
      direction = f.direction, duration = f.duration, resolvedLocale = f.resolvedLocale,
      observerName = f.observerName, analysisLabel = f.analysisLabel, html = f.html,
      recovering = !!f.recovering, warnings = [];
  // ── 6. Save to Analysis sheet (plain text only — no HTML) ─────────────────
  try {
    var aSheet = getOrCreateSheet(spreadsheet, ANALYSIS_SHEET);
    ensureHeaders(aSheet, ANALYSIS_HEADERS);
    if (!recovering || !_sheetHasRef_(aSheet, auditRef)) appendRow(aSheet, [
      auditRef, now,
      formData.sapId || '', formData.participant || '',
      observerName, analysisLabel, htmlToPlainText(html)
    ]);
  } catch (ae) {
    // Like the Dashboard and Audit_Log writes below, this must not stop the others: it
    // used to throw, and the Dashboard and Audit_Log rows were then never written.
    Logger.log('Analysis write error (non-fatal): ' + ae);
    warnings.push('Analysis entry failed to save — contact admin with ref ' + auditRef);
  }

  // ── 7. Write to Dashboard_Data ────────────────────────────────────────────
  try {
    var dSheet = getOrCreateSheet(spreadsheet, DASHBOARD_DATA_SHEET);
    ensureHeaders(dSheet, DASHBOARD_HEADERS);
    var callReason       = extractTextBlock(html, 'Call Reason')      || '';
    var callSummary      = extractTextBlock(html, 'Call Summary')      || '';
    var opportunities    = extractTextBlock(html, 'Opportunities')     || '';
    var recommendations  = extractTextBlock(html, 'Recommendation')    || '';
    var criticalFlags    = extractTextBlock(html, 'Critical Flag')     || '';
    var repeatPct        = extractTextBlock(html, 'Repeat Risk')       || '';
    var issueResolved    = extractTextBlock(html, 'Issue Resolution')  || '';
    var transferOccurred = extractTextBlock(html, 'Transfer')          || '';

    // Structured RCA fields — all parsed from the evaluation HTML. RCA category,
    // product opportunity and sales attempted now arrive in the hidden
    // div.ai-meta the evaluation prompt emits, so the separate enrichment AI
    // call this used to need is gone.
    var structured = null;
    try {
      structured = extractStructuredRCAFromHTML(html);
    } catch(re) {
      // Do not swallow this silently: it blanks eleven Dashboard columns while
      // the submit still reports success.
      Logger.log('extractStructuredRCAFromHTML FAILED for ' + auditRef + ' — 11 Dashboard columns will be blank: ' + re);
      warnings.push('Some dashboard fields could not be extracted for ref ' + auditRef);
    }
    var callDriver        = structured ? structured.callDriver        : '';
    var rcaCategory       = structured ? structured.rcaCategory       : '';
    var rcaSubParameter   = structured ? structured.rcaSubParameter   : '';
    var topOpportunity    = structured ? structured.topOpportunity    : '';
    var productOpportunity= structured ? structured.productOpportunity: '';
    var salesAttempted    = structured ? structured.salesAttempted    : '';


    var smartS            = structured ? structured.smartS            : '';
    var smartM            = structured ? structured.smartM            : '';
    var smartA            = structured ? structured.smartA            : '';
    var smartR            = structured ? structured.smartR            : '';
    var smartT            = structured ? structured.smartT            : '';
    if (structured && structured.callSummaryShort && !callSummary)
      callSummary = structured.callSummaryShort;

    if (!recovering || !_sheetHasRef_(dSheet, auditRef)) appendRow(dSheet, [
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
      'Pending', 'Not Sent', formData.agentEmail || '',
      resolveAuditorEmail_(observerName)
    ]);
    invalidateDashboardCache_();
  } catch(de) {
    Logger.log('Dashboard_Data write error (non-fatal): ' + de);
    warnings.push('Dashboard entry failed to save — contact admin with ref ' + auditRef);
  }

  // ── 8. Write to Audit_Log ─────────────────────────────────────────────────
  try {
    var logSheet = getOrCreateSheet(spreadsheet, AUDIT_LOG_SHEET);
    ensureHeaders(logSheet, AUDIT_LOG_HEADERS);
    if (!recovering || !_sheetHasRef_(logSheet, auditRef)) appendRow(logSheet, [
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
    invalidateAuditLogCache_(); // bust cache so next read is fresh
  } catch(le) {
    Logger.log('Audit_Log write error (non-fatal): ' + le);
    warnings.push('Audit log entry failed to save — contact admin with ref ' + auditRef);
  }
  return warnings;
}

// ─────────────────────────────────────────────────────────────────────────────
// Main submission — cache check first, then AI if needed
// ─────────────────────────────────────────────────────────────────────────────
function submitTranscript(formData) {
  var _procFlagKey = '';   // set only while THIS call holds the "being processed" flag
  try {
    if (!formData.sapId || !formData.transcript) {
      return { success: false, error: 'SAP ID and transcript are required.' };
    }

    // If QA analyst manually selected a LOB from the modal, use it as the primary LOB
    if (formData.selectedLOB) {
      formData.lineOfBusiness = formData.selectedLOB;
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

    // The guard keys on the RESOLVED Interaction ID (the form's, else the one read from the
    // transcript). It used to read only the form field, so with that field blank (it is
    // filled from the transcript) there was no guard and a double click ran the AI twice.
    // It is released in the finally block at the bottom: a cache hit and "Access denied"
    // used to return without releasing it, which told the analyst their own finished
    // submit was "already being processed" for the next ten minutes.
    // ── Per-Interaction-ID duplicate-submission guard ──────────────────────────
    // Two QA analysts submitting the same Interaction ID concurrently would each
    // pass the cache-miss check, both run the AI (~2-5 min each), and both
    // append duplicate rows to every sheet.  We use a short LockService window
    // to atomically check/set a CacheService "processing" flag, then release
    // the lock immediately so the slow AI work does not block other submissions.
    // CacheService (not PropertiesService) is used because it supports TTL:
    // if a GAS execution is killed by the 6-minute hard timeout, the flag
    // auto-expires and the interaction becomes submittable again — PropertiesService
    // has no expiry and would deadlock that interaction ID permanently.
    if (interactionId) {
      var _lock = LockService.getScriptLock();
      var _alreadyProcessing = false;
      try {
        _lock.waitLock(6000); // only hold long enough to read+write one cache entry
        var _procCache = CacheService.getScriptCache();
        var _flagKey   = _procKey_(interactionId, formData.sapId);
        if (_procCache.get(_flagKey)) {
          _alreadyProcessing = true;
        } else {
          _procCache.put(_flagKey, String(Date.now()), 10 * 60); // auto-expires in 10 min
          _procFlagKey = _flagKey;   // ours: released in the finally block, on EVERY way out
        }
      } catch(le) {
        Logger.log('submitTranscript: lock acquisition failed — ' + le);
        // Non-fatal: proceed without the guard rather than blocking the user
      } finally {
        try { _lock.releaseLock(); } catch(le2) {}
      }
      if (_alreadyProcessing) {
        return { success: false, error: 'This interaction is already being processed. Please wait a moment and refresh.' };
      }
    }

    // ── 1. Cache check ────────────────────────────────────────────────────────
    var reusedHtml = null;   // set when the evaluation is cached but this audit has no record yet
    if (interactionId) {
      var cached = findCachedResult_(interactionId, analysisType, formData.sapId);
      if (cached) {
        var cachedRef = findAuditRefForInteraction_(interactionId, analysisType, formData.sapId);
        if (cachedRef) {
          Logger.log('Returning cached result for: ' + interactionId);
          return {
            success:      true,
            html:         sharedCSS() + fixBadgeClasses(cached),
            sheetUrl:     spreadsheet.getUrl(),
            fromCache:    true,
            auditRef:     cachedRef,
            observerName: (formData.observerName || '')
          };
        }
        // The report is cached but no Audit_Log row exists for this agent: an earlier attempt
        // failed (or was killed) after the cache write. Reporting success here used to hand
        // back an EMPTY audit reference and write nothing, so the analyst saw a report that
        // was in no sheet and the email then failed. Finish the records from the cached
        // report instead; the AI is not called again.
        Logger.log('Cached evaluation for ' + interactionId + ' has no audit record - completing the records from the cache');
        reusedHtml = cached;
      }
    }

    // ── 2. Generate audit reference + resolve observer ────────────────────────
    var observerName  = (formData.observerName || '').trim();
    var analysisLabel = analysisType === 'sales' ? 'Sales Analyzer' : 'Repeats & Transfer Analyzer';
    var warnings      = [];
    // Completing a cached audit: reuse the reference the failed attempt already stamped into
    // the report (its Transcripts row has it). Otherwise a new one.
    var reusedRef = (reusedHtml !== null) ? findOrphanTranscriptRef_(interactionId, formData.sapId, analysisLabel) : '';
    var auditRef  = reusedRef || generateAuditRef_();

    // Resolve locale — fall back to roster lookup if the form field was blank
    var resolvedLocale = (formData.locale || '').trim();
    if (!resolvedLocale && formData.sapId) {
      try {
        var rosterRow = lookupBySapId_(formData.sapId.trim());
        if (rosterRow && rosterRow.locale) resolvedLocale = rosterRow.locale;
      } catch(le) { Logger.log('Locale fallback lookup failed (non-fatal): ' + le); }
    }

    // Backfill any still-blank team-member fields from this SAP ID's most
    // recent PRIOR audit in Audit_Log, so a failed/blank client-side roster
    // lookup (e.g. the external Roster/AT-Data lookup failing) doesn't block
    // the submission entirely. Data may be stale if the team member was
    // reassigned since their last audit — flagged as a warning, not applied
    // silently.
    if (formData.sapId && (!formData.teamLeader || !formData.opsManager || !formData.lineOfBusiness || !resolvedLocale || !formData.vtid)) {
      var teamHistory = getTeamDetailsFromAuditHistory_(formData.sapId.trim());
      if (teamHistory) {
        var usedHistoryFallback = false;
        if (!formData.teamLeader     && teamHistory.teamLeader)     { formData.teamLeader     = teamHistory.teamLeader;     usedHistoryFallback = true; }
        if (!formData.opsManager     && teamHistory.opsManager)     { formData.opsManager     = teamHistory.opsManager;     usedHistoryFallback = true; }
        if (!formData.lineOfBusiness && teamHistory.lineOfBusiness) { formData.lineOfBusiness = teamHistory.lineOfBusiness; usedHistoryFallback = true; }
        if (!resolvedLocale          && teamHistory.locale)         { resolvedLocale          = teamHistory.locale;         usedHistoryFallback = true; }
        if (!formData.vtid           && teamHistory.vtid)           { formData.vtid           = teamHistory.vtid;           usedHistoryFallback = true; }
        if (usedHistoryFallback) {
          warnings.push('Some team member details (Team Leader / Ops Manager / LOB / Locale / VTID) could not be freshly looked up and were filled in from a previous audit for this SAP ID — please verify they are still current.');
        }
      }
    }

    // ── 3. Save transcript row (not again when completing an audit whose row exists) ─
    if (!reusedRef) {
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
          analysisLabel, _cellSafeTranscript_(formData.transcript),
          formData.agentEmail || ''
        ]);
      } catch(we) {
        var msg = we.toString();
        if (msg.indexOf('permission') !== -1 || msg.indexOf('document') !== -1 || msg.indexOf('access') !== -1) {
          return {
            success: false,
            error: 'Access denied: You need Editor (not Viewer) access to the "' + SPREADSHEET_NAME + '" spreadsheet. ' +
                   'Please ask your admin to update your sharing permissions to Editor access.'
          };
        }
        throw we;
      }
    }

    var filterStats = { applied: false, targetFound: false };
    var html;
    if (reusedHtml !== null) {
      html = reusedHtml;
      warnings.push('This evaluation had already been generated, but its audit record was missing, so the records were completed from the saved report (ref ' + auditRef + ').');
    } else {
      var made = produceEvaluation_(formData, {
        interactionId: interactionId, participant: participant, analysisType: analysisType,
        auditRef: auditRef, observerName: observerName
      });
      html = made.html;
      filterStats = made.filterStats;
    }

    // ── 6-8. Analysis, Dashboard_Data and Audit_Log rows ──────────────────────
    warnings = warnings.concat(writeAuditRecords_(spreadsheet, formData, {
      auditRef: auditRef, now: now, participant: participant, interactionId: interactionId,
      direction: direction, duration: duration, resolvedLocale: resolvedLocale,
      observerName: observerName, analysisLabel: analysisLabel, html: html,
      recovering: reusedHtml !== null
    }));

    return {
      success:      true,
      html:         sharedCSS() + html,
      sheetUrl:     spreadsheet.getUrl(),
      fromCache:    false,
      auditRef:     auditRef,
      observerName: observerName,
      warnings:     warnings,
      filterStats:  filterStats,
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
    // Log the stack, not just the message. A bare "You do not have permission to
    // access the requested document" names neither the file nor the line, which
    // is why the TI report took so long to trace. The stack names both.
    Logger.log('submitTranscript error: ' + e.toString() + '\n' + (e.stack || '(no stack)'));
    // Also write it straight into the spreadsheet — reading Stackdriver/GCP
    // logs needs a non-default GCP project linked to this script, which this
    // project does not have, so Logger.log alone is not actually reachable.
    logSubmissionError_('submitTranscript', formData, e);
    // Google's raw error was going straight to the analyst's screen. It tells
    // them nothing they can act on. Check the AI-service tag FIRST: callFuelIX_()
    // (AIService.gs) is called completely unguarded above, and a FuelIX auth/
    // outage error's response body often contains wording like "permission" or
    // "access" — the generic substring check below used to catch that and
    // misreport it as a Sheets/Drive lookup problem, sending everyone hunting
    // through Roster sharing for a bug that was actually the AI call.
    var _msg = e.toString();
    if (_msg.indexOf('AI_SERVICE_ERROR[truncated]') !== -1) {
      return {
        success: false,
        error: 'The AI report came out too long and was cut off, so this audit was not saved. ' +
               'Please submit again; a new attempt is usually shorter. If it keeps happening ' +
               'with the same transcript, report it to the Analyzer admin.'
      };
    }
    if (_msg.indexOf('AI_SERVICE_ERROR') !== -1) {
      return {
        success: false,
        error: 'The AI analysis service did not respond correctly, so this audit ' +
               'was not saved. This is usually temporary — please wait a moment ' +
               'and try submitting again. If it keeps happening, report it to the ' +
               'Analyzer admin; the full error has been recorded in the Error_Log tab.'
      };
    }
    // This deployment runs as the script owner (Execute as: Me), not as the
    // accessing user, so a per-user file-sharing gap is not the likely cause
    // of a genuine Sheets/Drive permission error here — this is more often a
    // shared quota/concurrency limit (every user's requests count against the
    // same owner account) or a transient Google-side hiccup. Either way the
    // analyst can't fix it, so point them at the admin instead of guessing why.
    if (_msg.indexOf('permission') !== -1 || _msg.indexOf('do not have access') !== -1) {
      return {
        success: false,
        error: 'A backend lookup failed, so this audit was not saved. This is not ' +
               'a problem with your own access. Please report this to the Analyzer ' +
               'admin — the full error has been recorded in the Error_Log tab.'
      };
    }
    return { success: false, error: _msg };
  } finally {
    // Every way out (success, cache hit, access denied, AI failure, an unexpected error)
    // releases the flag THIS call took, and only that one.
    if (_procFlagKey) {
      try { CacheService.getScriptCache().remove(_procFlagKey); } catch (pe) {}
    }
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Email audit — called from Submit button
// ─────────────────────────────────────────────────────────────────────────────
function sendAuditEmail(formData, htmlResult) {
  try {
    // htmlResult is the analyst's EDITED report as the browser serialised it. It
    // goes into the email and, via updateCachedResult_, into the page the agent
    // opens, so it is reduced to the allow-list first.
    htmlResult = sanitizeEvalHtml_(htmlResult);

    // ── Verify the audit actually exists before sending anything ──────────────
    var auditRefCheck = (formData.auditRef || '').trim();
    if (!auditRefCheck) {
      return { success: false, error: 'Missing audit reference.' };
    }
    var auditLog = readAuditLog_();
    var formSapId         = (formData.sapId         || '').toString().trim();
    var formInteractionId = (formData.interactionId || '').toString().trim();
    var rowsWithRef = auditLog.filter(function(r) {
      return (r['Audit Ref'] || '').toString().trim() === auditRefCheck;
    });
    if (!rowsWithRef.length) {
      return { success: false, error: 'Audit reference not found.' };
    }
    // A reference that two audits share cannot be emailed safely: the link
    // carries only the reference and always opens the first of the two rows, so
    // the second agent would be shown the first agent's evaluation. This only
    // affects references issued before the NH and TI scripts got separate
    // prefixes; reportDuplicateAuditRefs() lists them and
    // applyDuplicateAuditRefRenames() gives the later row of each pair a new one.
    if (rowsWithRef.length > 1) {
      var sharedMsg = 'The audit reference ' + auditRefCheck + ' is shared by more than one saved evaluation ' +
        '(an old reference from before the NH and TI references were separated), so nothing was emailed. ' +
        'Please contact the QA and LS admin team to clean it up.';
      Logger.log('sendAuditEmail: reference ' + auditRefCheck + ' is shared by ' + rowsWithRef.length + ' rows');
      logSubmissionError_('sendAuditEmail:sharedReference', formData, new Error(sharedMsg));
      return { success: false, error: sharedMsg };
    }
    var auditRow = rowsWithRef[0];

    // ── The form must still be the audit that was saved under this ref ────────────────
    // The subject and recipients below come from the form, but the evaluation
    // link opens whatever the Audit_Log row for this ref holds. If the two are
    // about different agents the email would tell one person and link to
    // another's evaluation, so send nothing. A blank field on either side is
    // not a claim and is not compared.
    var rowSapId         = (auditRow['SAP ID']         || '').toString().trim();
    var rowInteractionId = (auditRow['Interaction ID'] || '').toString().trim();
    // (the check below makes the two equal whenever both are filled in; when the form's is blank the saved one is used, so the
    // recipients are always looked up for the agent who was audited, never for whatever is left in the form)
    if ((formSapId && rowSapId && formSapId !== rowSapId) ||
        (formInteractionId && rowInteractionId && formInteractionId !== rowInteractionId)) {
      var mismatchMsg = 'This evaluation (' + auditRefCheck + ') was saved for a different team member or ' +
        'interaction than the one now on the form, so nothing was emailed. ' +
        'Please click Clear All, then Process Script again for the correct team member.';
      Logger.log('sendAuditEmail: identity mismatch for ' + auditRefCheck +
                 ' — form SAP ' + formSapId + ' / interaction ' + formInteractionId +
                 ' vs saved SAP ' + rowSapId + ' / interaction ' + rowInteractionId);
      logSubmissionError_('sendAuditEmail:identityMismatch', formData, new Error(mismatchMsg));
      return { success: false, error: mismatchMsg };
    }
    if (rowSapId) formData.sapId = rowSapId;   // the recipients are looked up for the audited agent, whatever the form still holds

    // ── Re-check admin status server-side — never trust the client's emailMode claim ──
    // Same rule as the screen's TEST toggle (getObserverInfo): isAdminCaller_().
    // They used to differ — the screen accepted any domain, this accepted only
    // telus.com — so an admin signed in on telusinternational.com saw TEST but the
    // server quietly made it LIVE and emailed the real agent and leaders.
    var callerIsAdmin = isAdminCaller_();

    // Only an admin, or the analyst who saved this audit, may email it. sendAuditEmail is open to every signed-in user (the page calls it)
    // and sends from the owner's mailbox, so without this anyone could send mail about any audit (the refs are easy to guess).
    var _obs = {}; try { _obs = resolveObserver_() || {}; } catch (oe) {}
    var _callerEmail = ''; try { _callerEmail = (Session.getActiveUser().getEmail() || '').toString().toLowerCase().trim(); } catch (ce) {}
    if (!callerIsAdmin && !_callerOwnsAudit_(auditRow, _obs, _callerEmail)) {
      var notYoursMsg = 'Only the analyst who saved this audit, or an admin, can email it. Nothing was sent.';
      Logger.log('sendAuditEmail: ' + (_callerEmail || 'unknown caller') + ' is not the observer of ' + auditRefCheck + ' (' + (auditRow['Observer Name'] || 'no observer') + '); nothing sent');
      logSubmissionError_('sendAuditEmail:notYourAudit', formData, new Error(notYoursMsg));
      return { success: false, error: notYoursMsg };
    }
    if (!_emailRateTake_(_callerEmail, callerIsAdmin)) {
      return { success: false, error: 'You have already sent a lot of audit emails in the last hour, so nothing was sent. Please try again later.' };
    }
    // The QA recipient is the person who is signed in, never an address typed into the form (an admin sending on someone's behalf is
    // resolved from the roster by the observer name on the saved audit).
    if (!callerIsAdmin) formData.observerEmail = _callerEmail;
    else formData.observerEmail = '';
    if (!formData.observerName && auditRow['Observer Name']) formData.observerName = auditRow['Observer Name'];

    // Asking for TEST and not being allowed to is REFUSED, never turned into a real
    // send: sending nothing is recoverable, emailing the wrong people is not.
    if (formData.emailMode === 'test' && !callerIsAdmin) {
      Logger.log('sendAuditEmail: TEST mode requested by a caller who is not an admin; nothing sent for ' + auditRefCheck);
      return { success: false, error: 'Test mode is only for the admin team, so nothing was emailed. ' +
        'Switch the email mode to LIVE if you mean to email the team, or ask an admin to run the test.' };
    }
    var effectiveMode = (formData.emailMode === 'test') ? 'test' : 'live';

    var recipients;
    if (effectiveMode === 'test') {
      var adminList = getRecipientsFromRoster_('Admin/Dev');
      recipients = adminList.map(function(r) { return r.email; }).filter(Boolean);
    } else {
      recipients = buildAuditRecipients_(formData);
    }
    // An address outside the company domains is never mailed. It is dropped, logged and reported back so it is not silent.
    var _split = _splitInternalRecipients_(recipients);
    var _dropped = _split.dropped;
    recipients = _split.ok;
    if (_dropped.length) {
      Logger.log('sendAuditEmail: not mailed (not a company address): ' + _dropped.join(', '));
      logSubmissionError_('sendAuditEmail:externalRecipient', formData, new Error('Not mailed, not a company address: ' + _dropped.join(', ')));
    }

    if (!recipients.length && _dropped.length) {
      return { success: false, error: 'None of the recipients is a company address, so nothing was emailed: ' + _dropped.join(', ') };
    }
    if (!recipients.length) {
      // Say which list came up empty: "nobody found" in TEST mode (the Admin/Dev
      // list) is a different problem from LIVE mode (agent, leader, ops manager, QA).
      Logger.log('sendAuditEmail: no recipients resolved in ' + effectiveMode + ' mode for ' + auditRefCheck);
      return { success: false, error: 'Could not resolve recipient email addresses' +
        (effectiveMode === 'test'
          ? ' (TEST mode: no Admin/Dev address was found in the Roster tab).'
          : ' (LIVE mode: none of the agent, team leader, ops manager or QA addresses were found).') };
    }

    var agentName     = _oneLine_(auditRow['Team Member'] || formData.participant || 'Agent', 80);
    var sapId         = _oneLine_(rowSapId || formData.sapId || 'N/A', 20);
    var interactionId = _oneLine_(rowInteractionId || formData.interactionId || 'N/A', 120);
    var auditRef      = _oneLine_(auditRefCheck, 64);
    var analysisLabel = formData.analysisType === 'sales' ? 'Sales Analyzer' : 'Repeats & Transfer Analyzer';
    var firstName     = agentName.split(' ')[0];
    var evalTitle     = formData.analysisType === 'sales' ? 'Sales Performance Evaluation' : 'New Hire Evaluation';
    var lobDisplay    = _oneLine_(formData.selectedLOB || formData.lineOfBusiness || 'N/A', 60);
    var auditType2    = _oneLine_(formData.auditType || 'N/A', 40);
    var subject       = (effectiveMode === 'test' ? '[TEST] ' : '') +
                        agentName + ' | ' + auditRef + ' | ' + _oneLine_(formData.observerName || 'N/A', 80) + ' | ' + auditType2 + ' | ' + lobDisplay + ' | ' + _oneLine_(formData.locale || 'N/A', 60);

    // ── Plain-text body ───────────────────────────────────────────────────────
    var body =
      'Hi ' + firstName + ',\n\n' +
      'We\'re excited to share feedback from your recent customer interaction!\n\n' +
      'Your call has been reviewed to highlight your strengths and provide insights that will help you continue to grow and excel in your role.\n\n' +
      'EVALUATION DETAILS\n' +
      '─────────────────────────────────────\n' +
      'Audit Reference:  ' + auditRef + '\n' +
      'Interaction Date: ' + _oneLine_(formData.startTime || 'N/A', 40) + '\n' +
      'Customer BAN:     ' + _oneLine_(formData.customerBAN || 'N/A', 40) + '\n' +
      'Interaction ID:   ' + interactionId + '\n' +
      'Listening Type:   ' + _oneLine_(formData.direction || 'N/A', 30) + '\n' +
      'Audit Type:       ' + auditType2 + '\n' +
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
      _oneLine_(formData.startTime || 'N/A', 40),
      _oneLine_(formData.customerBAN || 'N/A', 40),
      _oneLine_(formData.direction || 'N/A', 30),
      analysisLabel, evalUrl
    );

    // ── Send email (in batches — wrapped so partial failure doesn't skip sheets) ──
    var _auditEmailErr  = null;
    var _auditEmailWarn = null; // set when email failed due to quota (audit still saved)
    try {
      sendEmailInBatches_(recipients, subject, body, htmlBody, 'NH Call Analyzer');
    } catch(emailEx) {
      var _errStr = emailEx.toString();
      Logger.log('sendAuditEmail: email failure — ' + _errStr);
      if (_errStr.indexOf('QUOTA_EXCEEDED') !== -1) {
        // Quota exhausted: audit is saved — treat as a warning, not a hard failure
        _auditEmailWarn = 'Audit saved. Email could not be sent — Google\'s daily email quota was reached. Recipients will NOT receive the email today. Quota resets at midnight Pacific time.';
      } else {
        _auditEmailErr = _errStr;
      }
    }

    // ── Update sheets — always runs even on partial email failure ─────────────
    // 'Sent' exactly on success — nine call sites in index.html and Dashboard.html
    // count a row as delivered only when this reads `=== 'Sent'`.
    var sheetStatus = _auditEmailErr ? 'Partial send' : (_auditEmailWarn ? 'Quota exceeded' : 'Sent');

    // htmlResult is the analyst's EDITED report, so the Issue Resolution badge
    // here reflects any override they made after reading the evaluation. The
    // row was written during submitTranscript from the AI's original answer,
    // so read it back off the badge and correct the sheet.
    var resolvedByAnalyst = '';
    try { resolvedByAnalyst = extractTextBlock(htmlResult || '', 'Issue Resolution') || ''; } catch(rx) {}

    // The email above may have already been sent successfully by this point —
    // a failure here is a separate, non-critical step (recording delivery
    // status on the sheet) and must not make the whole request report as
    // failed when the email itself went out fine. Same reasoning as the
    // notifyAdmins_/updateCachedResult_ guards just below.
    try {
      updateDashboardPDFLink_(interactionId, sheetStatus, recipients, auditRef, resolvedByAnalyst);
    } catch(de) {
      Logger.log('updateDashboardPDFLink in sendAuditEmail failed: ' + de);
      logSubmissionError_('sendAuditEmail:updateDashboardPDFLink', formData, de);
    }

    Logger.log('Audit email ' + (_auditEmailErr ? 'PARTIALLY FAILED' : _auditEmailWarn ? 'QUOTA EXCEEDED' : '') + ' sent to: ' + recipients.join(', '));
    // notifyAdmins_ only runs in live mode. In test mode the main email already went
    // exclusively to Admin/Dev, so calling notifyAdmins_ would cause a double-send.
    if (effectiveMode === 'live') {
      try { notifyAdmins_(formData, auditRefCheck); } catch(ne) { Logger.log('notifyAdmins_ failed: ' + ne); }
    }

    // Update Cache Sheet with user-edited HTML so EvalView shows the edited version
    try {
      var intId = formInteractionId || rowInteractionId;
      // The saved audit decides the type. The Repeats/Sales radio can still be
      // flipped after Process, and an edit written under the other type would
      // never be shown and would be served to a later audit of that type.
      var aType = auditRow['Analysis Type']
        ? _normalizeAnalysisType(auditRow['Analysis Type'])
        : (formData.analysisType || 'repeats').trim();
      if (intId && htmlResult) {
        updateCachedResult_(intId, aType, htmlResult, rowSapId || formSapId);
        // Clear the EvalView CacheService entry so the next visit re-reads the updated sheet
        var evKey = 'ev2_' + Utilities.base64Encode(auditRefCheck).substring(0, 200);
        CacheService.getScriptCache().remove(evKey);
      }
    } catch(ue) { Logger.log('updateCachedResult_ in sendAuditEmail failed: ' + ue); }

    var _result;
    if (_auditEmailErr) {
      return { success: false, error: _auditEmailErr, recipients: recipients, auditRef: auditRef };
    }
    // Quota warning: audit was saved and form should lock, but show a warning banner
    if (_auditEmailWarn) {
      _result = { success: true, emailWarning: _auditEmailWarn, recipients: recipients, auditRef: auditRef };
    } else {
      _result = { success: true, recipients: recipients, auditRef: auditRef };
    }
    if (_dropped.length) _result.droppedRecipients = _dropped;
    return _result;

  } catch(e) {
    Logger.log('sendAuditEmail error: ' + e.toString() + '\n' + (e.stack || '(no stack)'));
    logSubmissionError_('sendAuditEmail', formData, e);
    // Same reasoning as submitTranscript's catch: this deployment runs as the
    // script owner (Execute as: Me), so a per-user permission gap is not the
    // likely cause here — showing the raw Google/Sheets message just sends
    // analysts hunting for an access problem that isn't theirs to fix.
    var _msg = e.toString();
    if (_msg.indexOf('permission') !== -1 || _msg.indexOf('do not have access') !== -1) {
      return {
        success: false,
        error: 'A temporary system error prevented this email from sending. This is not ' +
               'a problem with your own access — please try Submit Email again. If it keeps ' +
               'happening, report it to the Analyzer admin; the full error has been recorded ' +
               'in the Error_Log tab.'
      };
    }
    return { success: false, error: _msg };
  }
}

// ── Batch email sender — splits recipients into chunks of 50 to stay under ────
// ── Google Apps Script's "Email Recipients Per Message" limit ─────────────────
// Continues ALL batches even if one fails, then throws a summary error so the
// caller can detect partial failures without leaving later batches un-attempted.
function sendEmailInBatches_(recipients, subject, body, htmlBody, senderName) {
  if (!recipients || recipients.length === 0) {
    Logger.log('sendEmailInBatches_: recipients list is empty — no email sent');
    return;
  }
  var BATCH_SIZE = 50;
  var failedBatches = [];
  for (var i = 0; i < recipients.length; i += BATCH_SIZE) {
    var batchNum = Math.floor(i / BATCH_SIZE) + 1;
    var chunk = recipients.slice(i, i + BATCH_SIZE);

    // Check remaining quota before attempting each batch
    try {
      var remaining = MailApp.getRemainingDailyQuota();
      if (remaining <= 0) {
        Logger.log('sendEmailInBatches_: quota exhausted (' + remaining + ' remaining) — skipping batch ' + batchNum);
        failedBatches.push({ batch: batchNum, count: chunk.length, error: 'QUOTA_EXCEEDED' });
        continue;
      }
    } catch(qe) {
      Logger.log('sendEmailInBatches_: could not check quota — ' + qe);
    }

    try {
      MailApp.sendEmail({
        to:       chunk.join(','),
        subject:  subject,
        body:     body,
        htmlBody: htmlBody,
        name:     senderName || 'NH Call Analyzer'
      });
      Logger.log('sendEmailInBatches_: batch ' + batchNum + ' sent (' + chunk.length + ' recipients)');
    } catch(e) {
      var errStr = e.toString();
      Logger.log('sendEmailInBatches_: batch ' + batchNum + ' FAILED (' + chunk.length + ' recipients skipped): ' + errStr);
      // Normalize GAS quota errors to a consistent prefix so callers can detect them
      var isQuota = errStr.indexOf('too many times') !== -1 || errStr.indexOf('premium email') !== -1;
      failedBatches.push({ batch: batchNum, count: chunk.length, error: isQuota ? 'QUOTA_EXCEEDED' : errStr });
    }
  }
  // Throw summary only after ALL batches have been attempted so partial sends
  // don't block later batches. Caller decides how to surface this.
  if (failedBatches.length > 0) {
    var allQuota = failedBatches.every(function(b) { return b.error === 'QUOTA_EXCEEDED'; });
    throw new Error(
      (allQuota ? 'QUOTA_EXCEEDED: ' : '') +
      'sendEmailInBatches_: ' + failedBatches.length + ' of ' +
      Math.ceil(recipients.length / BATCH_SIZE) + ' batch(es) failed. ' +
      'First error: ' + failedBatches[0].error
    );
  }
}

// ── Resolve email from name — checks Primary Roster then Global Roster ────────
function resolveEmail_(name) {
  if (!name) return '';
  try {
    var nameLower = name.trim().toLowerCase();

    // 0. Fast path: check the in-memory/CacheService agent email map first.
    //    _getAgentEmailMap() reads the sheet once per execution then caches in
    //    memory, so this avoids an additional uncached sheet read on every call.
    var _emailMap = _getAgentEmailMap();
    if (_emailMap && Object.prototype.hasOwnProperty.call(_emailMap, nameLower)) {
      // null means two people share this name: stop, do not scan for a first match.
      return _emailMap[nameLower] || '';
    }

    // 1. Primary Roster — match on Agent_Name column only (col C, index 2)
    var ss    = openSheetWithRetry(ROSTER_SHEET_ID, 'Roster Sheet');
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
    Logger.log('resolveEmail_ error: ' + e);
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
  /* ── Compact card layout — PDF-safe overrides (CSS grid → table model) ── */
  '.report-wrap *{box-sizing:border-box!important}' +
  '.report-header{background:#4B286D!important;border-radius:6px!important;padding:10px 14px!important;margin-bottom:10px!important;color:#fff!important}' +
  '.report-header-title{font-size:13px!important;font-weight:800!important;color:#fff!important}' +
  '.report-header-sub{font-size:9px!important;color:#ddd!important}' +
  '.ai-info{display:block!important;margin-bottom:10px!important}' +
  '.ai-chip{display:inline-block!important;background:#F4F4F7!important;border:1px solid #D8D8D8!important;border-radius:4px!important;padding:3px 7px!important;margin:2px!important;font-size:8px!important;vertical-align:top!important}' +
  '.ai-chip-label{font-size:7px!important;font-weight:700!important;color:#54565A!important;text-transform:uppercase!important;display:block!important}' +
  '.ai-chip-val{font-size:9px!important;font-weight:600!important;color:#1A1A2E!important}' +
  '.report-badge-yes{display:inline-block!important;background:#2B8000!important;color:#fff!important;border-radius:3px!important;padding:1px 5px!important;font-size:8px!important;font-weight:700!important}' +
  '.report-badge-no{display:inline-block!important;background:#C12335!important;color:#fff!important;border-radius:3px!important;padding:1px 5px!important;font-size:8px!important;font-weight:700!important}' +
  '.report-summary-wrap{background:#F9F9F9!important;border:1px solid #E0E0E0!important;border-radius:5px!important;padding:8px 12px!important;margin-bottom:10px!important}' +
  '.report-summary-label{font-size:8px!important;font-weight:700!important;color:#4B286D!important;text-transform:uppercase!important;margin-bottom:5px!important}' +
  '.ai-summary{font-size:9px!important;line-height:1.5!important;color:#1A1A2E!important}' +
  '.report-summary-meta{font-size:8px!important;color:#767676!important;margin-top:6px!important;border-top:1px solid #E8E8E8!important;padding-top:5px!important}' +
  '.report-3col{display:table!important;width:100%!important;table-layout:fixed!important;border-collapse:separate!important;border-spacing:4px!important;margin-bottom:12px!important}' +
  '.report-col{display:table-cell!important;width:33.33%!important;border-radius:5px!important;padding:8px 10px!important;vertical-align:top!important}' +
  '.report-col-working{background:#EDF7E6!important;border:1px solid #B3DFA0!important}' +
  '.report-col-change{background:#FFF5F5!important;border:1px solid #F5AAAA!important}' +
  '.report-col-howto{background:#F5F0FF!important;border:1px solid #D1B8E8!important}' +
  '.report-col-head{font-size:8.5px!important;font-weight:700!important;margin-bottom:6px!important;padding-bottom:4px!important;border-bottom:1px solid rgba(0,0,0,.1)!important}' +
  '.report-col-working .report-col-head{color:#2B8000!important}' +
  '.report-col-change .report-col-head{color:#C12335!important}' +
  '.report-col-howto .report-col-head{color:#4B286D!important}' +
  '.report-col-list{padding-left:12px!important;font-size:8.5px!important;line-height:1.6!important;margin:0!important}' +
  '.report-col-list li{margin-bottom:3px!important}' +
  '.report-col-roleplays{list-style:none!important;padding-left:0!important}' +
  '.report-col-roleplays li{margin-bottom:8px!important}' +
  '.report-flags{margin-bottom:10px!important}' +
  '.report-flags-head{font-size:9.5px!important;font-weight:700!important;color:#C12335!important;padding:5px 8px!important;background:#FFF0F0!important;border:1px solid #F5AAAA!important;border-radius:4px 4px 0 0!important}' +
  '.report-flags-body{border:1px solid #F5AAAA!important;border-top:none!important;border-radius:0 0 4px 4px!important;padding:8px!important;background:#fff!important}' +
  '.report-manual-flags{margin-bottom:10px!important}' +
  '.report-manual-flags-head{font-size:9.5px!important;font-weight:700!important;color:#4B286D!important;padding:5px 8px!important;background:#F5F0FF!important;border:1px solid #D1B8E8!important;border-radius:4px 4px 0 0!important}' +
  '.report-manual-flags-body{border:1px solid #D1B8E8!important;border-top:none!important;border-radius:0 0 4px 4px!important;padding:8px!important;background:#fff!important}' +
  '.report-footer{background:#F4F4F7!important;border:1px solid #E0E0E0!important;border-radius:5px!important;padding:6px 12px!important;display:block!important;font-size:8px!important;color:#54565A!important;margin-top:6px!important}' +
  '.report-footer-ref{font-weight:700!important;color:#4B286D!important}' +
  '.report-footer-obs{color:#767676!important}' +
  '.ai-flag-rl-label{font-size:8px!important;font-weight:700!important;color:#4B286D!important;text-transform:uppercase!important}' +
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
  requireAdmin_();
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
    .replace(/\s+on\w+="[^"]*"/gi, '')
    .replace(/\s+on\w+='[^']*'/gi, '')
    .replace(/(href|src|action)\s*=\s*["']javascript:[^"']*["']/gi, '$1="#"')
    .replace(/(href|src|action)\s*=\s*javascript:[^\s>]*/gi, '$1="#"');
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

// Full HTML escaping (including double-quotes for href/attribute contexts).
// NOTE: esc() is used by renderAnalyticsHTML — do NOT remove it when cleaning up
// the dead-code HTML builders block below.
function esc(s) {
  if (s === null || s === undefined) return '';
  return String(s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
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
        '<a href="' + e(evalUrl || '#') + '" ' +
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
  auditRef, submittedAt, observer, analysisType, auditType,
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
        '<span><strong style="color:#4B286D">Audit:</strong> ' + e(auditType) + '</span>' +
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

// Export to Google Sheets: see exportAudits in ExportService.gs (the browser sends Audit Refs only).

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
// Safe to delete in a future cleanup pass — EXCEPT: esc() has been moved above
// this block (near escEmail) because renderAnalyticsHTML() is live and uses it.
// ─────────────────────────────────────────────────────────────────────────────

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
  if (stmt) html += '<div class="r-pos-stmt-label">Sample Positioning Statement &#9654; Roleplay &amp; Practice</div>' +
                    '<div class="r-pos-stmt" contenteditable="true">' +
                    '<button class="tts-play-btn" data-tts="' + esc(stmt) + '" contenteditable="false" title="Play sample statement" style="background:#22c55e;border:none;border-radius:50%;width:20px;height:20px;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;padding:0;vertical-align:middle;margin-right:5px;flex-shrink:0"><svg width="9" height="10" viewBox="0 0 9 10" fill="white"><polygon points="1,1 8,5 1,9"/></svg></button>' +
                    '"' + esc(stmt) + '"</div>';
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

// ── Action Registry logging — called from append_registry.py via Web App POST ──
//
// The secret belongs in Script Properties (key: REGISTRY_SECRET), not in source
// control. This constant WAS the secret and is in the git history, so treat it as
// public. It is kept only as a TRANSITIONAL fallback so the external poster keeps
// working until the property is set; once the property exists the constant is
// ignored, and it can then be deleted.
var REGISTRY_SECRET_LEGACY = 'NH_REGISTRY_2026_DPZ';

function getRegistrySecret_() {
  var v = PropertiesService.getScriptProperties().getProperty('REGISTRY_SECRET');
  if (v) return v;
  Logger.log('WARNING: Script Property REGISTRY_SECRET is not set; using the legacy secret that is in source control. Set the property and rotate it.');
  return REGISTRY_SECRET_LEGACY;
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    if (!payload.key || payload.key !== getRegistrySecret_()) {
      return ContentService.createTextOutput(JSON.stringify({success: false, error: 'Unauthorized'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var result = appendRegistryRow_(payload.data || {});
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({success: false, error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Public entry for the Apps Script Execution API and manual runs. It must stay
// public (private functions cannot be called that way), so it is admin-only; the
// secret-guarded doPost() calls the private writer below instead.
function appendRegistryRow(data) {
  requireAdmin_();
  return appendRegistryRow_(data);
}

function appendRegistryRow_(data) {
  var REGISTRY_SS_ID = '1a58nCQPv9B0C1E0m30x5fibDrgnaJL6A7Pyz265Fjac';
  var ss = SpreadsheetApp.openById(REGISTRY_SS_ID);
  var sh = ss.getSheetByName('Registry');
  if (!sh) throw new Error('Registry sheet not found');

  var lastRow = sh.getLastRow();
  var nextNum = '001';
  if (lastRow > 1) {
    var lastVal = sh.getRange(lastRow, 1).getValue();
    var n = parseInt(String(lastVal), 10);
    nextNum = isNaN(n) ? String(lastRow) : String(n + 1);
    while (nextNum.length < 3) nextNum = '0' + nextNum;
  }

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var row = [
    nextNum,
    data.date || today,
    data.type || 'Bug Fix',
    data.scope || 'NH',
    data.issue || '',
    data.resolution || '',
    data.files || '',
    data.author || 'Danzen',
    data.status || 'Done'
  ];
  appendRow(sh, row);   // the sanitising wrapper: a value starting with = + - @ must not become a formula
  SpreadsheetApp.flush();

  var newRow = sh.getLastRow();
  var cols   = sh.getLastColumn();

  var bg = (newRow % 2 === 0) ? '#F8F9FA' : '#FFFFFF';
  sh.getRange(newRow, 1, 1, cols).setBackground(bg);
  sh.setRowHeight(newRow, 80);
  sh.getRange(newRow, 1, 1, cols).setWrap(true).setVerticalAlignment('middle');

  var typeColors = {
    'Bug Fix':     {bg: '#FEE2E2', fg: '#991B1B'},
    'Feature':     {bg: '#DCFCE7', fg: '#166534'},
    'Enhancement': {bg: '#DBEAFE', fg: '#1E40AF'},
    'Performance': {bg: '#FEF3C7', fg: '#92400E'},
    'Maintenance': {bg: '#F3E8FF', fg: '#6B21A8'}
  };
  var tc = typeColors[data.type] || {bg: '#F3F4F6', fg: '#374151'};
  sh.getRange(newRow, 3).setBackground(tc.bg).setFontColor(tc.fg).setFontWeight('bold');

  var statusColors = {
    'Done':        {bg: '#D1FAE5', fg: '#065F46'},
    'In Progress': {bg: '#FEF3C7', fg: '#92400E'},
    'Pending':     {bg: '#E5E7EB', fg: '#374151'}
  };
  var sc = statusColors[data.status] || {bg: '#E5E7EB', fg: '#374151'};
  sh.getRange(newRow, 9).setBackground(sc.bg).setFontColor(sc.fg).setFontWeight('bold');

  return {success: true, rowNum: nextNum, sheetRow: newRow};
}

