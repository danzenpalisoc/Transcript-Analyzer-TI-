/**
 * AIService.gs
 * FuelIX API + prompt builders that return HTML directly (no JSON parsing).
 */

function callFuelIX(prompt) {
  var endpoint = FUELIX_CONFIG.baseUrl + '/v1/chat/completions';
  var payload  = {
    model:      FUELIX_CONFIG.model,
    messages:   [{ role: 'user', content: prompt }],
    stream:     false,
    max_tokens: 20000   // raise ceiling — some Sales evals need ~17-18k tokens to complete
  };
  var options = {
    method:         'post',
    contentType:    'application/json',
    headers: {
      'Authorization': 'Bearer ' + FUELIX_CONFIG.apiKey,
      'Content-Type':  'application/json'
    },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true,
    deadline:           270   // 270s — leaves headroom before the GAS 360s execution limit
  };
  var response = UrlFetchApp.fetch(endpoint, options);
  var code     = response.getResponseCode();
  var text     = response.getContentText();
  if (code !== 200) throw new Error('AI service error ' + code + ': ' + text);
  var data = JSON.parse(text);
  if (!data.choices || !data.choices[0] || !data.choices[0].message)
    throw new Error('Unexpected AI response structure');
  var finishReason = (data.choices[0].finish_reason || '').toString();
  if (finishReason === 'length' || finishReason === 'max_tokens') {
    Logger.log('WARNING: AI response truncated by token limit (finish_reason=' + finishReason + '). Consider increasing max_tokens further.');
  }
  return data.choices[0].message.content.trim();
}

// ── PDF knowledge ─────────────────────────────────────────────────────────────
function getPDFKnowledgeFromFolder(folderId, cacheKey) {
  var cache  = CacheService.getScriptCache();
  var cached = cache.get(cacheKey);
  if (cached) return cached;
  try {
    var folder  = DriveApp.getFolderById(folderId);
    var files   = folder.getFilesByType(MimeType.PDF);
    var allText = '';
    while (files.hasNext()) {
      var file = files.next();
      try {
        var docFile = Drive.Files.copy(
          { title: file.getName() + '_tmp', mimeType: MimeType.GOOGLE_DOCS },
          file.getId()
        );
        allText += '\n\n--- ' + file.getName() + ' ---\n';
        allText += DocumentApp.openById(docFile.id).getBody().getText();
        DriveApp.getFileById(docFile.id).setTrashed(true);
      } catch(e) { Logger.log('PDF read error: ' + e); }
    }
    if (allText) cache.put(cacheKey, allText.substring(0, 100000), PDF_CACHE_SECONDS);
    return allText;
  } catch(e) { Logger.log('getPDFKnowledgeFromFolder: ' + e); return ''; }
}

// lookupBySapId is now defined lower in this file (includes VTID lookup)

// ── Shared: load full roster sheet data (cached 2h) ─────────────────────────
// In-call cache only — GAS spawns a fresh process per google.script.run call,
// so this prevents duplicate sheet reads within ONE server execution, not across calls.
// Cross-call caching is handled by CacheService below.
var _rosterDataInMemory = null;

function _getRosterSheetData() {
  if (_rosterDataInMemory) return _rosterDataInMemory;
  try {
    var cache    = CacheService.getScriptCache();
    var cacheKey = 'roster_sheet_data_v2';
    var cached   = cache.get(cacheKey);
    if (cached) {
      try {
        _rosterDataInMemory = JSON.parse(cached);
        return _rosterDataInMemory;
      } catch(e) {}
    }
    var ss    = SpreadsheetApp.openById(ROSTER_SHEET_ID);
    var sheet = ss.getSheetByName('roster') || ss.getSheetByName('Roster');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    // Store only rows 1+ (skip header)
    var serializable = data.slice(1);
    try { cache.put(cacheKey, JSON.stringify(serializable).substring(0, 95000), 4 * 60 * 60); } catch(e) {}
    _rosterDataInMemory = serializable;
    return serializable;
  } catch(e) { Logger.log('_getRosterSheetData: ' + e); return []; }
}

// ── Roster lookup by participant name (reverse — name → SAP ID, cached) ──────
function lookupSapId(participantName) {
  try {
    var data      = _getRosterSheetData();
    var nameLower = participantName.trim().toLowerCase();
    for (var i = 0; i < data.length; i++) {
      if (data[i][ROSTER_COL_AGENT_NAME] &&
          data[i][ROSTER_COL_AGENT_NAME].toString().trim().toLowerCase() === nameLower) {
        var sap = data[i][ROSTER_COL_SAP_ID];
        return sap !== '' && sap !== undefined ? sap.toString().trim() : '';
      }
    }
    return '';
  } catch(e) { Logger.log('lookupSapId: ' + e); return ''; }
}

// ── Get all roster data for dashboard filter prefill (cached 2h) ─────────────
function getAllRosterData() {
  try {
    var cache    = CacheService.getScriptCache();
    var cacheKey = 'roster_all_data_v2';
    var cached   = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch(e) {}
    }
    var data   = _getRosterSheetData();
    var result = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row[ROSTER_COL_SAP_ID]) continue;
      var agentName = (row[ROSTER_COL_AGENT_NAME] || '').toString().trim();
      var sapIdStr  = (row[ROSTER_COL_SAP_ID]     || '').toString().trim();
      result.push({
        sapId:          sapIdStr,
        participant:    agentName,
        lineOfBusiness: (row[ROSTER_COL_DOMAIN_NAME] || '').toString().trim(),
        teamLeader:     (row[ROSTER_COL_TEAM_MGR]    || '').toString().trim(),
        opsManager:     (row[ROSTER_COL_OPS_MGR]     || '').toString().trim(),
        locale:         (row[ROSTER_COL_LOCALE]       || '').toString().trim()
      });
    }
    try { cache.put(cacheKey, JSON.stringify(result).substring(0, 95000), 4 * 60 * 60); } catch(e) {}
    return result;
  } catch(e) { Logger.log('getAllRosterData: ' + e); return []; }
}

// ── Diagnostic: run once from editor to find VTID column in AT Data GCP ───────
function diagnoseATDataColumns() {
  try {
    var ss    = SpreadsheetApp.openById(AT_DATA_GCP_SS_ID);
    var sheet = ss.getSheets()[0];
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log('=== AT Data GCP columns ===');
    headers.forEach(function(h, i) {
      Logger.log('Col ' + String.fromCharCode(65 + i) + ' (idx ' + i + '): ' + h);
    });
    // Log first 3 data rows for spot-check
    var sample = sheet.getRange(2, 1, Math.min(3, sheet.getLastRow()-1), sheet.getLastColumn()).getValues();
    sample.forEach(function(row, ri) {
      Logger.log('Row ' + (ri+2) + ': ' + row.slice(0,8).join(' | '));
    });
  } catch(e) { Logger.log('diagnoseATDataColumns error: ' + e); }
}

// ── VTID lookup ───────────────────────────────────────────────────────────────
// ── Shared: load AT Data GCP sheet once, cache for 2 hours ──────────────────
var _atDataInMemory = null;
var _globalRosterDataInMemory = null;

function _getATDataGCPSheetData() {
  if (_atDataInMemory) return _atDataInMemory;
  try {
    var cache    = CacheService.getScriptCache();
    var cacheKey = 'at_data_gcp_v2';
    var cached   = cache.get(cacheKey);
    if (cached) {
      try {
        _atDataInMemory = JSON.parse(cached);
        Logger.log('AT Data GCP loaded from cache: ' + _atDataInMemory.length + ' rows');
        return _atDataInMemory;
      } catch(e) {}
    }
    Logger.log('AT Data GCP cache miss — reading sheet...');
    var ss    = SpreadsheetApp.openById(AT_DATA_GCP_SS_ID);
    var sheet = ss.getSheetByName('Roster') || ss.getSheetByName('roster') || ss.getSheets()[0];
    var data  = sheet.getDataRange().getValues();
    // Cache all rows (slice(1) to skip header — store header separately)
    var header = data[0];
    var rows   = data.slice(1);
    var payload = JSON.stringify({ header: header, rows: rows });
    try { cache.put(cacheKey, payload.substring(0, 95000), 4 * 60 * 60); } catch(e) {}
    _atDataInMemory = data; // full data including header row
    Logger.log('AT Data GCP loaded from sheet: ' + rows.length + ' rows');
    return _atDataInMemory;
  } catch(e) { Logger.log('_getATDataGCPSheetData: ' + e); return [[]]; }
}

// ── Shared: load Global Roster sheet data (cached 4h) ────────────────────────
function _getGlobalRosterData() {
  if (_globalRosterDataInMemory) return _globalRosterDataInMemory;
  try {
    var cache    = CacheService.getScriptCache();
    var cacheKey = 'global_roster_v1';
    var cached   = cache.get(cacheKey);
    if (cached) {
      try {
        _globalRosterDataInMemory = JSON.parse(cached);
        return _globalRosterDataInMemory;
      } catch(e) {}
    }
    var ss    = SpreadsheetApp.openById(GLOBAL_ROSTER_SS_ID);
    var sheet = ss.getSheetByName('Global Roster');
    if (!sheet) return { header: [], rows: [] };
    var data    = sheet.getDataRange().getValues();
    var payload = { header: data[0] || [], rows: data.slice(1) };
    try { cache.put(cacheKey, JSON.stringify(payload).substring(0, 95000), 4 * 60 * 60); } catch(e) {}
    _globalRosterDataInMemory = payload;
    return payload;
  } catch(e) { Logger.log('_getGlobalRosterData: ' + e); return { header: [], rows: [] }; }
}

// AT Data GCP file → "Roster" tab
// Col G (idx 6) = SAP ID to match
// Col F (idx 5) = VTID to return
function lookupVTID(sapId) {
  if (!sapId) return '';
  try {
    var cache    = CacheService.getScriptCache();
    var cacheKey = 'vtid3_' + sapId.toString().trim();
    var cached   = cache.get(cacheKey);
    if (cached !== null) {
      Logger.log('VTID cache hit for SAP ' + sapId + ': ' + cached);
      return cached;
    }

    var data    = _getATDataGCPSheetData();
    var headers = data[0] || [];

    // Default columns: G=6 for SAP ID, F=5 for VTID
    var sapCol  = 6;  // G
    var vtidCol = 5;  // F

    // Auto-detect from headers — Col G = "Production ID", Col F = VTID
    headers.forEach(function(h, i) {
      var hl = h.toString().toLowerCase().trim();
      if (hl === 'production id' || hl === 'production_id' ||
          hl === 'sap_id'        || hl === 'sapid' || hl === 'sap id') sapCol  = i;
      if (hl === 'vtid'          || hl === 'vt_id' || hl === 'vt id')  vtidCol = i;
    });
    Logger.log('VTID: sapCol=' + sapCol + ' ("' + headers[sapCol] + '"), vtidCol=' + vtidCol + ' ("' + headers[vtidCol] + '")');

    var targetNum = Number(sapId.toString().trim());
    var targetStr = sapId.toString().trim();

    for (var i = 1; i < data.length; i++) {
      var cell    = data[i][sapCol];
      if (cell === null || cell === undefined || cell === '') continue;
      var cellNum = Number(cell);
      var cellStr = cell.toString().trim();
      if (cellStr === targetStr || (!isNaN(targetNum) && !isNaN(cellNum) && cellNum === targetNum)) {
        var vtid = (data[i][vtidCol] || '').toString().trim();
        Logger.log('VTID found for SAP ' + targetStr + ': ' + vtid);
        cache.put(cacheKey, vtid, 6 * 60 * 60);
        return vtid;
      }
    }

    Logger.log('VTID not found for SAP: ' + targetStr);
    cache.put(cacheKey, '', 30 * 60);
    return '';
  } catch(e) {
    Logger.log('lookupVTID error: ' + e);
    return '';
  }
}

// ── SAP ID lookup: tries roster first, then AT Data GCP ──────────────────────
function lookupBySapId(sapId) {
  var targetStr = sapId.toString().trim();
  var targetNum = Number(targetStr);

  // ── 1. Try roster (uses cached data — no sheet read) ─────────────────────
  try {
    var rosterRows = _getRosterSheetData();  // cached 2h, in-memory within execution
    for (var i = 0; i < rosterRows.length; i++) {
      var cell = rosterRows[i][ROSTER_COL_SAP_ID];
      if (cell === null || cell === undefined || cell === '') continue;
      if (cell.toString().trim() === targetStr ||
          (!isNaN(targetNum) && Number(cell) === targetNum)) {
        var agentName = (rosterRows[i][ROSTER_COL_AGENT_NAME] || '').toString().trim();
        Logger.log('Roster match for SAP ' + targetStr + ' → ' + agentName);
        return {
          participant:    agentName,
          lineOfBusiness: (rosterRows[i][ROSTER_COL_DOMAIN_NAME] || '').toString().trim(),
          teamLeader:     (rosterRows[i][ROSTER_COL_TEAM_MGR]    || '').toString().trim(),
          opsManager:     (rosterRows[i][ROSTER_COL_OPS_MGR]     || '').toString().trim(),
          locale:         (rosterRows[i][ROSTER_COL_LOCALE]       || '').toString().trim(),
          vtid:           lookupVTID(targetStr),
          agentEmail:     ''
        };
      }
    }
    Logger.log('SAP ' + targetStr + ' not in roster — trying Global Roster');
  } catch(e) { Logger.log('Roster lookup error: ' + e); }

  // ── 1.5. Try Global Roster ────────────────────────────────────────────────
  try {
    var gData  = _getGlobalRosterData();
    var gHdr   = gData.header;
    var gRows  = gData.rows;
    var gSapCol = -1, gNameCol = -1, gTLCol = -1, gOMCol = -1, gEmailCol = -1, gLOBCol = -1, gLocaleCol = -1;
    gHdr.forEach(function(h, i) {
      var hl = h.toString().toLowerCase().trim();
      if (hl === 'sap id' || hl === 'sap_id')          gSapCol    = i;
      if (hl === 'member full name')                    gNameCol   = i;
      if (hl === 'tl full name fixed')                  gTLCol     = i;
      if (hl === 'om full name fixed')                  gOMCol     = i;
      if (hl === 'email address' || hl === 'email')     gEmailCol  = i;
      if (hl === 'domain name')                         gLOBCol    = i;
      if (hl === 'campus name')                         gLocaleCol = i;
    });
    if (gSapCol !== -1) {
      for (var g = 0; g < gRows.length; g++) {
        var gCell = gRows[g][gSapCol];
        if (gCell === null || gCell === undefined || gCell === '') continue;
        if (gCell.toString().trim() === targetStr ||
            (!isNaN(targetNum) && Number(gCell) === targetNum)) {
          var gName   = gNameCol   > -1 ? (gRows[g][gNameCol]   || '').toString().trim() : '';
          var gTL     = gTLCol     > -1 ? (gRows[g][gTLCol]     || '').toString().trim() : '';
          var gOM     = gOMCol     > -1 ? (gRows[g][gOMCol]     || '').toString().trim() : '';
          var gEmail  = gEmailCol  > -1 ? (gRows[g][gEmailCol]  || '').toString().trim() : '';
          var gLOB    = gLOBCol    > -1 ? (gRows[g][gLOBCol]    || '').toString().trim() : '';
          var gLocale = gLocaleCol > -1 ? (gRows[g][gLocaleCol] || '').toString().trim() : '';
          Logger.log('Global Roster match for SAP ' + targetStr + ' → ' + gName);
          return {
            participant:    gName,
            lineOfBusiness: gLOB,
            teamLeader:     gTL,
            opsManager:     gOM,
            locale:         gLocale,
            vtid:           lookupVTID(targetStr),
            agentEmail:     gEmail
          };
        }
      }
    }
    Logger.log('SAP ' + targetStr + ' not in Global Roster — trying AT Data GCP');
  } catch(e) { Logger.log('Global Roster lookup error: ' + e); }

  // ── 2. Fall back to AT Data GCP (uses cached data — no sheet read) ────────
  try {
    var data2  = _getATDataGCPSheetData();  // cached 2h, in-memory within execution
    var hdr    = data2[0] || [];

    // Detect column positions from headers
    var prodCol   = 6;  // G - Production ID
    var agentCol  = 7;  // H - Agent
    var lobCol    = 2;  // C - LOB
    var locCol    = 4;  // E - Location
    var vtidCol   = 5;  // F - Reference No. (VTID)

    hdr.forEach(function(h, idx) {
      var hl = h.toString().toLowerCase().trim();
      if (hl === 'production id' || hl === 'production_id') prodCol  = idx;
      if (hl === 'agent')                                    agentCol = idx;
      if (hl === 'lob')                                      lobCol   = idx;
      if (hl === 'location')                                 locCol   = idx;
      if (hl === 'reference no.' || hl === 'vtid' || hl === 'vt_id') vtidCol = idx;
    });

    for (var j = 1; j < data2.length; j++) {
      var prod = data2[j][prodCol];
      if (prod === null || prod === undefined || prod === '') continue;
      if (prod.toString().trim() === targetStr ||
          (!isNaN(targetNum) && Number(prod) === targetNum)) {
        var name = (data2[j][agentCol] || '').toString().trim();
        var lob  = (data2[j][lobCol]   || '').toString().trim();
        var loc  = (data2[j][locCol]   || '').toString().trim();
        var vtid = (data2[j][vtidCol]  || '').toString().trim();
        // Col I (idx 8) = Facilitator → use as Team Leader
        var facilitator = (data2[j][8] || '').toString().trim();
        hdr.forEach(function(h, idx) {
          if (h.toString().toLowerCase().trim() === 'facilitator') facilitator = (data2[j][idx] || '').toString().trim();
        });

        Logger.log('AT Data GCP match for SAP ' + targetStr + ' → ' + name +
                   ' | LOB=' + lob + ' | Locale=' + loc + ' | VTID=' + vtid + ' | Facilitator=' + facilitator);

        // Build result — then enrich from FCR Dashboard Data if possible
        var result = {
          participant:    name,
          lineOfBusiness: lob,
          teamLeader:     facilitator,
          opsManager:     '',
          locale:         loc,
          vtid:           vtid,
          agentEmail:     ''
        };

        // Try to enrich Team Leader / Ops Manager from FCR Dashboard Data
        var enriched = lookupFromFCRDashboard(targetStr, targetNum);
        if (enriched) {
          if (enriched.teamLeader)     result.teamLeader     = enriched.teamLeader;
          if (enriched.opsManager)     result.opsManager     = enriched.opsManager;
          if (enriched.lineOfBusiness && !result.lineOfBusiness) result.lineOfBusiness = enriched.lineOfBusiness;
          if (enriched.locale && !result.locale)                 result.locale         = enriched.locale;
        }

        return result;
      }
    }
    Logger.log('SAP ' + targetStr + ' not found in AT Data GCP either');
  } catch(e) { Logger.log('AT Data GCP lookup error: ' + e); }

  return null;
}

// ── FCR Dashboard Data lookup — enriches Team Leader / Ops Manager ────────────
// Searches every sheet in the FCR Dashboard file for the SAP ID
// Returns { teamLeader, opsManager, lineOfBusiness, locale } or null
function lookupFromFCRDashboard(targetStr, targetNum) {
  try {
    var cache    = CacheService.getScriptCache();
    var cacheKey = 'fcr2_' + targetStr;  // bumped version to bust stale cache
    var cached   = cache.get(cacheKey);
    if (cached && cached !== 'null') {
      Logger.log('FCR cache hit for SAP ' + targetStr);
      try { return JSON.parse(cached); } catch(e) {}
    }

    var ss     = SpreadsheetApp.openById(FCR_DASHBOARD_SS_ID);
    var sheets = ss.getSheets();

    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s];
      var data  = sheet.getDataRange().getValues();
      if (data.length < 2) continue;
      var hdr = data[0];

      // Find SAP ID column
      var sapCol = -1;
      var tlCol  = -1;
      var omCol  = -1;
      var lobCol = -1;
      var locCol = -1;

      hdr.forEach(function(h, i) {
        var hl = h.toString().toLowerCase().trim();
        if (hl === 'sap id'            || hl === 'sap_id'     ||
            hl === 'sapid'             || hl === 'production id' ||
            hl === 'production_id')                                      sapCol = i;
        if (hl === 'team leader'       || hl === 'team_leader' ||
            hl === 'team_mgr_name'     || hl === 'teamleader')           tlCol  = i;
        if (hl === 'operations manager'|| hl === 'ops_mgr_name'||
            hl === 'ops_mgr'           || hl === 'opsmgr')               omCol  = i;
        if (hl === 'line of business'  || hl === 'lob'        ||
            hl === 'domain_name'       || hl === 'agent role (scorecard)') lobCol = i;
        if (hl === 'locale'            || hl === 'location'   ||
            hl === 'site')                                               locCol = i;
      });

      if (sapCol === -1) continue;  // no SAP ID column on this sheet

      for (var i = 1; i < data.length; i++) {
        var cell = data[i][sapCol];
        if (cell === null || cell === undefined || cell === '') continue;
        if (cell.toString().trim() === targetStr ||
            (!isNaN(targetNum) && Number(cell) === targetNum)) {
          var result = {
            teamLeader:     tlCol  > -1 ? (data[i][tlCol]  || '').toString().trim() : '',
            opsManager:     omCol  > -1 ? (data[i][omCol]  || '').toString().trim() : '',
            lineOfBusiness: lobCol > -1 ? (data[i][lobCol] || '').toString().trim() : '',
            locale:         locCol > -1 ? (data[i][locCol] || '').toString().trim() : ''
          };
          Logger.log('FCR Dashboard match SAP ' + targetStr + ' on sheet "' + sheet.getName() + '": ' + JSON.stringify(result));
          // Only cache if at least one field has a value
          var hasData = result.teamLeader || result.opsManager || result.lineOfBusiness || result.locale;
          if (hasData) cache.put(cacheKey, JSON.stringify(result), 6 * 60 * 60);
          return result;
        }
      }
    }

    Logger.log('FCR Dashboard: SAP ' + targetStr + ' not found in any sheet');
    // Don't cache misses — data may be added later
    return null;
  } catch(e) {
    Logger.log('lookupFromFCRDashboard error: ' + e);
    return null;
  }
}

// ── Transcript metadata autofill ──────────────────────────────────────────────
function parseTranscriptMetadata(transcriptText) {
  var meta = {};
  var patterns = {
    interactionId: /Interaction ID[:\s]+([^\n\r]+)/i,
    startTime:     /Transcript Start Time[:\s]+([^\n\r]+)/i,
    duration:      /Transcript Duration[:\s]+([^\n\r]+)/i,
    direction:     /Direction[:\s]+([^\n\r]+)/i,
    participant:   /Internal Participant\(s\)[:\s]+([^\n\r]+)/i
  };
  for (var key in patterns) {
    var m = transcriptText.match(patterns[key]);
    if (m) meta[key] = m[1].trim();
  }
  var banPatterns = [
    /\bBAN[:\s#]+([A-Z0-9]{5,15})\b/i,
    /\baccount\s*(?:number|#|no\.?)[:\s]+([A-Z0-9]{5,15})\b/i,
    /\bcallback?\s*(?:number|#|no\.?)[:\s]+([\d\-\(\)\s]{7,20})/i
  ];
  for (var i = 0; i < banPatterns.length; i++) {
    var bm = transcriptText.match(banPatterns[i]);
    if (bm) { meta.customerBAN = bm[1].trim(); break; }
  }
  if (!meta.customerBAN) {
    var nm = transcriptText.match(/External\s+[^\n]*?([A-Z][a-z]+\s+[A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)\s/);
    if (nm) meta.customerName = nm[1].trim();
  }
  return meta;
}

// ─────────────────────────────────────────────────────────────────────────────
// CSS shared by both prompts (injected into the AI output)
// ─────────────────────────────────────────────────────────────────────────────
function sharedCSS() {
  return '<style>' +
    '.ai-section{margin-bottom:22px}' +
    '.ai-title{font-size:13px;font-weight:700;color:#4B286D;padding:8px 12px;' +
      'background:#F5F0FF;border-left:4px solid #4B286D;border-radius:0 4px 4px 0;' +
      'margin-bottom:11px;font-family:Helvetica Neue,Helvetica,Arial,sans-serif}' +
    '.ai-summary{background:#F9F9F9;border:1px solid #D8D8D8;border-radius:4px;' +
      'padding:12px 14px;font-size:13px;line-height:1.75;min-height:64px;' +
      'font-family:Helvetica Neue,Helvetica,Arial,sans-serif}' +
    '.ai-info{display:flex;flex-wrap:wrap;gap:9px;margin-bottom:6px}' +
    '.ai-chip{background:#F4F4F7;border:1px solid #D8D8D8;border-radius:4px;' +
      'padding:7px 12px;min-width:100px;font-family:Helvetica Neue,Helvetica,Arial,sans-serif}' +
    '.ai-chip-label{font-size:10px;font-weight:700;color:#54565A;text-transform:uppercase;' +
      'letter-spacing:.4px;display:block;margin-bottom:2px}' +
    '.ai-chip-val{font-size:13px;font-weight:600;color:#1A1A2E}' +
    '.ai-table{width:100%;border-collapse:collapse;font-size:13px;' +
      'font-family:Helvetica Neue,Helvetica,Arial,sans-serif;margin-bottom:4px}' +
    '.ai-table thead tr{background:#4B286D}' +
    '.ai-table th{padding:10px 14px;text-align:left;font-size:12px;' +
      'font-weight:700;color:#fff;letter-spacing:.3px}' +
    '.ai-table td{padding:10px 14px;border-bottom:1px solid #EBEBEB;' +
      'vertical-align:top;line-height:1.65;font-size:13px}' +
    '.ai-table tbody tr:nth-child(even) td{background:#FAFAFA}' +
    '.ai-table tbody tr:last-child td{border-bottom:none}' +
    '.ai-label-col{font-weight:600;color:#1A1A2E;width:22%;white-space:nowrap}' +
    '.ai-flag{background:#FFF5F5;border:1px solid #F5AAAA;border-left:4px solid #C12335;' +
      'border-radius:4px;padding:14px 16px;margin-bottom:11px;' +
      'font-family:Helvetica Neue,Helvetica,Arial,sans-serif}' +
    '.ai-flag-title{font-weight:700;color:#C12335;font-size:13px;margin-bottom:5px}' +
    '.ai-flag-detail{font-size:12px;color:#444;line-height:1.65;margin-bottom:8px}' +
    '.ai-flag-rl-label{font-size:10px;font-weight:700;text-transform:uppercase;' +
      'color:#2B8000;letter-spacing:.5px;margin-bottom:4px}' +
    '.ai-flag-stmt{background:#EDF7E6;border:1px solid #B3DFA0;border-radius:4px;' +
      'padding:9px 12px;font-size:13px;color:#1A5E00;font-style:italic;line-height:1.6}' +
    '.ai-hl-grid{display:grid;grid-template-columns:1fr 1fr;gap:13px}' +
    '.ai-hl-box{border-radius:6px;padding:14px 16px;' +
      'font-family:Helvetica Neue,Helvetica,Arial,sans-serif}' +
    '.ai-hl-high{background:#EDF7E6;border:1px solid #B3DFA0}' +
    '.ai-hl-low{background:#FFF5F5;border:1px solid #F5AAAA}' +
    '.ai-hl-title{font-size:12px;font-weight:700;margin-bottom:8px}' +
    '.ai-hl-high .ai-hl-title{color:#2B8000}' +
    '.ai-hl-low  .ai-hl-title{color:#C12335}' +
    '.ai-hl-box ul{padding-left:16px;font-size:13px;line-height:1.9}' +
    '.ai-coaching{padding-left:20px;font-size:13px;line-height:1.9;' +
      'font-family:Helvetica Neue,Helvetica,Arial,sans-serif}' +
    '.ai-coaching li{margin-bottom:4px}' +
    '.ai-score-panel{background:#F5F0FF;border:1px solid #D1B8E8;border-radius:6px;' +
      'padding:14px 18px;font-family:Helvetica Neue,Helvetica,Arial,sans-serif}' +
    '.ai-score-row{display:flex;align-items:center;justify-content:space-between;' +
      'padding:6px 0;border-bottom:1px solid #E4D8F5;font-size:13px}' +
    '.ai-score-row:last-child{border-bottom:none;font-weight:700}' +
    '.ai-badge{display:inline-block;border-radius:4px;padding:3px 10px;' +
      'font-size:12px;font-weight:700;color:#fff;min-width:44px;text-align:center}' +
    '.ai-badge-good{background:#2B8000}' +
    '.ai-badge-mid{background:#8C4A00}' +
    '.ai-badge-bad{background:#C12335}' +
    '.ai-call-badge{display:inline-block;background:#4B286D;color:#fff;font-size:11px;' +
      'font-weight:700;padding:3px 12px;border-radius:12px;margin-bottom:13px;' +
      'letter-spacing:.5px;font-family:Helvetica Neue,Helvetica,Arial,sans-serif}' +
    '.ai-call-block{border:1px solid #D8D8D8;border-radius:6px;padding:18px 20px;' +
      'margin-bottom:20px;background:#fff}' +
    '.ai-perfect{background:#EDF7E6;border:1px solid #B3DFA0;border-radius:4px;' +
      'padding:10px 14px;font-size:13px;line-height:1.6;margin-top:10px;' +
      'font-family:Helvetica Neue,Helvetica,Arial,sans-serif}' +
    '.ai-warning{background:#FFF0F0;border:1px solid #F5AAAA;border-left:4px solid #C12335;' +
      'border-radius:4px;padding:12px 16px;font-size:13px;color:#C12335;margin-top:8px}' +
    '</style>';
}

// ─────────────────────────────────────────────────────────────────────────────
// REPEATS PROMPT — asks AI to return complete HTML
// ─────────────────────────────────────────────────────────────────────────────
function buildRepeatsPrompt(transcriptText, knowledgeText) {
  var kb = knowledgeText
    ? '\n\nCOMPANY POLICIES AND PROCEDURES:\n' + knowledgeText + '\n\n'
    : '';

  return 'You are a Quality Analyst. Analyze the call transcript(s) below.\n' +
    'PURPOSE: Identify FCR opportunities, reduce repeat call rate and transfer rate.\n\n' +
    'EVALUATION FRAMEWORK — TELUS CUSTOMER EXPERIENCE BLUEPRINT:\n' +
    'Use these 4 pillars as your evaluation lens for ALL insights, coaching tips, and sample positioning statements:\n' +
    '• ENGAGE: Empathy vs Acknowledging — Acknowledge the OCCURRENCE (the event), validate the EMOTION. Scripted openers like "I\'m sorry" or "I apologize" without context show a lack of personalization and imply a mistake before understanding the situation.\n' +
    '• UNDERSTAND: Confirm vs Asking Questions — Ask open-ended questions to EXPLORE before confirming. Jumping to confirm before exploring means solving the wrong problem fast. Asking questions surfaces what the customer has not said yet.\n' +
    '• SOLVE: Explain vs Bridging — Explaining tells the customer what something is (information). Bridging shows WHY it matters to THEM specifically (connection). The customer should walk away thinking "that\'s exactly what I need" not just "I understand that."\n' +
    '• IMPRESS: Checking Understanding vs Setting Expectations — Checking understanding is REACTIVE (asks if the customer got it). Setting expectations is PROACTIVE (tells what comes next before they have to ask). Goal: customer leaves feeling informed, not just answered.\n' +
    'Non-Negotiables that must never be missing: Qualification → Research → Solve → Explain → Change → Summarize.\n' +
    'When writing insights, coaching tips, and sample positioning statements — ground them naturally in these CX Blueprint concepts. Use pillar language in your output (e.g. "The agent explained the solution but did not bridge it to the customer\'s specific situation" or "The agent confirmed an assumption instead of asking an open question to explore further").\n\n' +
    'CRITICAL INSTRUCTION: Return ONLY valid HTML. No markdown, no explanations, no JSON.\n' +
    'Your entire response must be HTML that uses exactly these CSS classes:\n\n' +
    'For each call, wrap everything in: <div class="ai-call-block">\n' +
    'Start with: <div class="ai-call-badge">Call 1</div>\n\n' +
    'Use this exact structure. START with the Overall Recommendation summary card BEFORE anything else:\n\n' +
    '<!-- OVERALL RECOMMENDATION SUMMARY — APPEARS FIRST AT THE TOP -->\n' +
    '<div class="ai-section" style="background:#FFF8E1;border:2px solid #F9A825;border-radius:8px;padding:16px 20px;margin-bottom:20px">\n' +
    '<div class="ai-title" style="background:#F9A825;color:#fff;border-left:4px solid #E65100;">&#128161; Overall Recommendation &amp; Coaching Summary</div>\n' +
    '<p contenteditable="true" style="font-size:13px;font-weight:700;margin-bottom:8px">&#128313; Summary: [2-3 sentence overall assessment of the agent performance on this call]</p>\n' +
    '<p contenteditable="true" style="font-size:13px;margin-bottom:8px">&#127919; Top Priority for Next Call: [The single most impactful action the agent must take immediately — be specific]</p>\n' +
    '<p style="font-size:12px;font-weight:700;color:#7B3F00;margin:0 0 4px">&#127919; SMART Coaching Focus Areas <em style="font-weight:400;font-size:11px">(Specific · Measurable · Attainable · Realistic · Time-bound)</em></p>\n' +
    '<table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:6px">\n' +
    '<thead><tr style="background:#FDE8B0"><th style="padding:4px 8px;text-align:left;width:5%">#</th><th style="padding:4px 8px;text-align:left;width:20%">S — Specific Behavior</th><th style="padding:4px 8px;text-align:left;width:19%">M — How to Measure</th><th style="padding:4px 8px;text-align:left;width:16%">A — Attainable Target</th><th style="padding:4px 8px;text-align:left;width:16%">R — Realistic</th><th style="padding:4px 8px;text-align:left;width:24%">T — Timeline</th></tr></thead>\n' +
    '<tbody>\n' +
    '<tr><td style="padding:4px 8px;border-bottom:1px solid #fde">1</td><td contenteditable="true" style="padding:4px 8px;border-bottom:1px solid #fde">[Exact behavior to change]</td><td contenteditable="true" style="padding:4px 8px;border-bottom:1px solid #fde">[e.g. Repeat rate drops below 20%]</td><td contenteditable="true" style="padding:4px 8px;border-bottom:1px solid #fde">[Achievable for this agent\'s current skill level]</td><td contenteditable="true" style="padding:4px 8px;border-bottom:1px solid #fde">[Why this is realistic — e.g. low-effort change, already done on some calls]</td><td contenteditable="true" style="padding:4px 8px;border-bottom:1px solid #fde">[By next 1-on-1 / within 2 weeks]</td></tr>\n' +
    '<tr><td style="padding:4px 8px;border-bottom:1px solid #fde">2</td><td contenteditable="true" style="padding:4px 8px;border-bottom:1px solid #fde">[Behavior 2]</td><td contenteditable="true" style="padding:4px 8px;border-bottom:1px solid #fde">[Measurement 2]</td><td contenteditable="true" style="padding:4px 8px;border-bottom:1px solid #fde">[Target 2]</td><td contenteditable="true" style="padding:4px 8px;border-bottom:1px solid #fde">[Realistic 2]</td><td contenteditable="true" style="padding:4px 8px;border-bottom:1px solid #fde">[Timeline 2]</td></tr>\n' +
    '<tr><td style="padding:4px 8px">3</td><td contenteditable="true" style="padding:4px 8px">[Behavior 3]</td><td contenteditable="true" style="padding:4px 8px">[Measurement 3]</td><td contenteditable="true" style="padding:4px 8px">[Target 3]</td><td contenteditable="true" style="padding:4px 8px">[Realistic 3]</td><td contenteditable="true" style="padding:4px 8px">[Timeline 3]</td></tr>\n' +
    '</tbody></table>\n' +
    '<p contenteditable="true" style="font-size:13px;">&#127775; Manager Coaching Tip: [Specific tip for the Team Leader on how to coach this agent — what to reinforce and what to redirect]</p>\n' +
    '</div>\n\n' +
    '<!-- COACHING TAKEAWAYS SUMMARY — APPEARS SECOND -->\n' +
    '<div class="ai-section" style="background:#E8F5E9;border:2px solid #2B8000;border-radius:8px;padding:14px 18px;margin-bottom:20px">\n' +
    '<div class="ai-title" style="background:#2B8000;color:#fff;border-left:4px solid #1B5E20;">&#127979; Key Coaching Takeaways</div>\n' +
    '<ol class="ai-coaching" style="padding-left:20px;font-size:13px;line-height:2.1">\n' +
    '<li contenteditable="true">[Most impactful takeaway — cite a verbatim moment from this call and the SMART action to replace it]</li>\n' +
    '<li contenteditable="true">[Second takeaway — specific, actionable, and measurable]</li>\n' +
    '<li contenteditable="true">[Third takeaway]</li>\n' +
    '</ol>\n' +
    '</div>\n\n' +
    'Then for EACH call use this structure:\n\n' +
    '<div class="ai-call-block">\n' +
    '<div class="ai-call-badge">Call 1</div>\n\n' +
    '<!-- CALL SUMMARY -->\n' +
    '<div class="ai-section">\n' +
    '<div class="ai-title">&#128222; Call Summary</div>\n' +
    '<div class="ai-summary">[100-word summary of what happened in the call]</div>\n' +
    '</div>\n\n' +
    '<!-- CALL INFORMATION -->\n' +
    '<div class="ai-section">\n' +
    '<div class="ai-title">&#128100; Call Information</div>\n' +
    '<div class="ai-info">\n' +
    '  <div class="ai-chip"><span class="ai-chip-label">Agent</span><span class="ai-chip-val">[name]</span></div>\n' +
    '  <div class="ai-chip"><span class="ai-chip-label">Date</span><span class="ai-chip-val">[date]</span></div>\n' +
    '  <div class="ai-chip"><span class="ai-chip-label">Phone</span><span class="ai-chip-val">[phone or N/A]</span></div>\n' +
    '  <div class="ai-chip"><span class="ai-chip-label">Country</span><span class="ai-chip-val">[country]</span></div>\n' +
    '  <div class="ai-chip"><span class="ai-chip-label">Department</span><span class="ai-chip-val">[dept]</span></div>\n' +
    '  <div class="ai-chip"><span class="ai-chip-label">Issue Resolved</span><span class="ai-chip-val">[Yes/No]</span></div>\n' +
    '  <div class="ai-chip"><span class="ai-chip-label">Repeat Risk</span><span class="ai-chip-val">[0-100%]</span></div>\n' +
    '</div>\n' +
    '</div>\n\n' +
    '<!-- ANALYSIS TABLE -->\n' +
    '<div class="ai-section">\n' +
    '<div class="ai-title">&#128203; Analysis — Opportunities &amp; Recommendations</div>\n' +
    '<table class="ai-table">\n' +
    '<thead><tr><th style="width:20%">Parameter</th><th style="width:28%">Finding / Detail</th><th style="width:52%">SMART Recommendation <span style="font-size:10px;font-weight:400">(Specific · Measurable · Attainable · Realistic · Time-bound)</span></th></tr></thead>\n' +
    '<tbody>\n' +
    '<tr><td class="ai-label-col">Repeat Projection</td><td contenteditable="true">[% and drivers]</td><td contenteditable="true">[S: exact behavior to change · M: target repeat % · A: achievable step · R: realistic given agent\'s current skill · T: by next audit or 2 weeks]</td></tr>\n' +
    '<tr><td class="ai-label-col">Issue Resolution</td><td contenteditable="true">[what was/was not resolved]</td><td contenteditable="true">[S: specific resolution step missed · M: measure by FCR rate · A: achievable · R: realistic — low effort to fix · T: apply from next call]</td></tr>\n' +
    '<tr><td class="ai-label-col">Process / Policy Gaps</td><td contenteditable="true">[gaps found]</td><td contenteditable="true">[S: which policy step · M: zero policy misses on next 5 calls · A: achievable with coaching · R: realistic — agent has tools · T: within 1 week]</td></tr>\n' +
    '<tr><td class="ai-label-col">Missing Steps</td><td contenteditable="true">[steps agent skipped]</td><td contenteditable="true">[S: name the exact missed step · M: applied on every relevant call · A: yes · R: realistic — already knows the step · T: immediately]</td></tr>\n' +
    '<tr><td class="ai-label-col">Callback Policy</td><td contenteditable="true">[followed or not]</td><td contenteditable="true">[S: exact policy requirement · M: 100% compliance on next calls · A: yes · R: realistic — simple script change · T: next call]</td></tr>\n' +
    '<tr><td class="ai-label-col">Transfer Analysis</td><td contenteditable="true">[transfer occurred? valid?]</td><td contenteditable="true">[S: correct transfer criteria · M: 0 invalid transfers next month · A: achievable · R: realistic — criteria are clear · T: by next QA review]</td></tr>\n' +
    '<tr><td class="ai-label-col">Probing Questions</td><td contenteditable="true">[questions used by agent]</td><td contenteditable="true">[S: 2 specific open-ended questions · M: used on every call · A: easy to practice · R: realistic — short habit to build · T: next call]</td></tr>\n' +
    '<tr><td class="ai-label-col">Agent Strengths</td><td contenteditable="true">[what agent did well]</td><td contenteditable="true">[S: keep doing X · M: maintain on 90% of calls · A: already demonstrated · R: natural strength · T: ongoing]</td></tr>\n' +
    '<tr><td class="ai-label-col">FCR Assessment</td><td contenteditable="true">[could this be 1 call?]</td><td contenteditable="true">[S: what would make it 1-call · M: FCR rate target · A: achievable · R: realistic with coaching · T: within 2 weeks]</td></tr>\n' +
    '<tr><td class="ai-label-col">3 SMART Actions</td><td contenteditable="true" colspan="2">[1. S:[action] M:[metric] A:[target] R:[why realistic] T:[timeline]   2. S:[action] M:[metric] A:[target] R:[why realistic] T:[timeline]   3. S:[action] M:[metric] A:[target] R:[why realistic] T:[timeline]]</td></tr>\n' +
    '</tbody>\n' +
    '</table>\n' +
    '</div>\n\n' +
    '<!-- CRITICAL FLAGS -->\n' +
    '<div class="ai-section">\n' +
    '<div class="ai-title">&#128681; Critical Flags &amp; Positioning Statements</div>\n' +
    '[Repeat this block for EACH critical flag found:]\n' +
    '<div class="ai-flag">\n' +
    '  <div class="ai-flag-title">&#9888; [Name of missed parameter or behavior]</div>\n' +
    '  <div class="ai-flag-detail">[Detailed explanation of what was missed and why it matters to the customer and FCR]</div>\n' +
    '  <div style="background:#FFF3E0;border:1px solid #FFB74D;border-radius:4px;padding:8px 12px;margin:8px 0;font-size:12px">\n' +
    '    <strong style="color:#E65100">&#127919; SMART Coaching Goal:</strong><br/>\n' +
    '    <span contenteditable="true">S: [Specific behavior to change for this flag] | M: [How success is measured — e.g. 0 occurrences in next 5 calls] | A: [Achievable target] | R: [Why this is realistic — e.g. agent already has the knowledge, low-effort fix] | T: [Timeline — e.g. within 2 coaching sessions]</span>\n' +
    '  </div>\n' +
    '  <div class="ai-flag-rl-label">&#127908; Sample Positioning Statement — Roleplay &amp; Practice</div>\n' +
    '  <div class="ai-flag-stmt" contenteditable="true">"[A complete, roleplay-ready statement the coach can say verbatim — e.g. \'[Agent name], when a customer asks X, try saying: [exact words]. This will help them feel Y and reduce repeat calls by Z%.\'  ]"</div>\n' +
    '</div>\n' +
    '</div>\n\n' +
    '</div>\n\n' +
    '<!-- HIGHLIGHTS AND LOWLIGHTS (outside call blocks, once at the end) -->\n' +
    '<div class="ai-section">\n' +
    '<div class="ai-title">&#9733; Highlights &amp; Lowlights</div>\n' +
    '<div class="ai-hl-grid">\n' +
    '<div class="ai-hl-box ai-hl-high"><div class="ai-hl-title">&#10003; Highlights</div><ul>\n' +
    '<li contenteditable="true">[highlight 1]</li>\n' +
    '<li contenteditable="true">[highlight 2]</li>\n' +
    '</ul></div>\n' +
    '<div class="ai-hl-box ai-hl-low"><div class="ai-hl-title">&#10007; Lowlights / Recommendations</div><ul>\n' +
    '<li contenteditable="true">[lowlight 1]</li>\n' +
    '<li contenteditable="true">[lowlight 2]</li>\n' +
    '</ul></div>\n' +
    '</div>\n' +
    '</div>\n\n' +
    'REPLACE all [placeholder] text with actual SMART analysis from the transcript.\n' +
    'Do NOT include any text outside the HTML tags.\n' +
    'Do NOT use markdown.\n' +
    'Make all contenteditable="true" attributes present on td and li elements.\n\n' +
    kb +
    'TRANSCRIPT:\n\n' + transcriptText;
}

// ─────────────────────────────────────────────────────────────────────────────
// SALES PROMPT — asks AI to return complete HTML
// ─────────────────────────────────────────────────────────────────────────────
function buildSalesPrompt(transcriptText, knowledgeText) {
  var kb = knowledgeText
    ? '\n\nCOMPANY POLICIES AND PROCEDURES:\n' + knowledgeText + '\n\n'
    : '';

  return 'You are an expert sales performance analyst for TELUS. Evaluate the call transcript.\n' +
    'PURPOSE: Identify sales opportunities and coach agents to increase sales.\n\n' +
    'EVALUATION FRAMEWORK — TELUS CUSTOMER EXPERIENCE BLUEPRINT:\n' +
    'Use these 4 pillars as your evaluation lens for ALL insights, coaching tips, and sample positioning statements:\n' +
    '• ENGAGE: Empathy vs Acknowledging — Acknowledge the OCCURRENCE (the event), validate the EMOTION. Scripted openers like "I\'m sorry" or "I apologize" without context show a lack of personalization and imply a mistake before understanding the situation.\n' +
    '• UNDERSTAND: Confirm vs Asking Questions — Ask open-ended questions to EXPLORE before confirming. Jumping to confirm before exploring means solving the wrong problem fast. Asking questions surfaces what the customer has not said yet.\n' +
    '• SOLVE: Explain vs Bridging — Explaining tells the customer what something is (information). Bridging shows WHY it matters to THEM specifically (connection). The customer should walk away thinking "that\'s exactly what I need" not just "I understand that." This is the core of value-based selling.\n' +
    '• IMPRESS: Checking Understanding vs Setting Expectations — Checking understanding is REACTIVE (asks if the customer got it). Setting expectations is PROACTIVE (tells what comes next before they have to ask). Goal: customer leaves feeling informed, not just answered.\n' +
    'Non-Negotiables that must never be missing: Qualification → Research → Solve → Explain → Change → Summarize.\n' +
    'When writing insights, coaching tips, and sample positioning statements — ground them naturally in these CX Blueprint concepts. Use pillar language in your output (e.g. "The agent explained the product features but did not bridge to why it matters to this customer\'s specific situation" or "The agent confirmed the need without asking open questions to fully explore it first").\n\n' +
    'CRITICAL INSTRUCTION: Return ONLY valid HTML. No markdown, no explanations, no JSON.\n' +
    'Your entire response must be HTML using exactly these CSS classes:\n\n' +
    'Use this exact structure. START with the Overall Recommendation and Coaching Takeaways summary cards at the very TOP before the score panel:\n\n' +
    '<!-- OVERALL RECOMMENDATION SUMMARY — FIRST THING AT THE TOP -->\n' +
    '<div class="ai-section" style="background:#FFF8E1;border:2px solid #F9A825;border-radius:8px;padding:16px 20px;margin-bottom:20px">\n' +
    '<div class="ai-title" style="background:#F9A825;color:#fff;border-left:4px solid #E65100;">&#128161; Overall Recommendation &amp; Coaching Summary</div>\n' +
    '<p contenteditable="true" style="font-size:13px;font-weight:700;margin-bottom:8px">&#128313; Summary: [2-3 sentence overall assessment of the agent sales performance on this call]</p>\n' +
    '<p contenteditable="true" style="font-size:13px;margin-bottom:8px">&#127919; Top Priority for Next Call: [The single most impactful sales action the agent must apply immediately — be specific and actionable]</p>\n' +
    '<p style="font-size:12px;font-weight:700;color:#7B3F00;margin:0 0 4px">&#127919; SMART Coaching Focus Areas <em style="font-weight:400;font-size:11px">(Specific · Measurable · Attainable · Realistic · Time-bound)</em></p>\n' +
    '<table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:6px">\n' +
    '<thead><tr style="background:#FDE8B0"><th style="padding:4px 8px;text-align:left;width:5%">#</th><th style="padding:4px 8px;text-align:left;width:20%">S — Specific Sales Behavior</th><th style="padding:4px 8px;text-align:left;width:19%">M — How to Measure</th><th style="padding:4px 8px;text-align:left;width:16%">A — Attainable Target</th><th style="padding:4px 8px;text-align:left;width:16%">R — Realistic</th><th style="padding:4px 8px;text-align:left;width:24%">T — Timeline</th></tr></thead>\n' +
    '<tbody>\n' +
    '<tr><td style="padding:4px 8px;border-bottom:1px solid #fde">1</td><td contenteditable="true" style="padding:4px 8px;border-bottom:1px solid #fde">[Exact sales behavior — e.g. Always present the bundle after identifying the need]</td><td contenteditable="true" style="padding:4px 8px;border-bottom:1px solid #fde">[e.g. Conversion rate or offer-made rate on next 5 calls]</td><td contenteditable="true" style="padding:4px 8px;border-bottom:1px solid #fde">[Achievable step for this agent]</td><td contenteditable="true" style="padding:4px 8px;border-bottom:1px solid #fde">[Why realistic — e.g. agent knows the product, small habit shift]</td><td contenteditable="true" style="padding:4px 8px;border-bottom:1px solid #fde">[By next 1-on-1 / within 2 weeks]</td></tr>\n' +
    '<tr><td style="padding:4px 8px;border-bottom:1px solid #fde">2</td><td contenteditable="true" style="padding:4px 8px;border-bottom:1px solid #fde">[Sales behavior 2]</td><td contenteditable="true" style="padding:4px 8px;border-bottom:1px solid #fde">[Measurement 2]</td><td contenteditable="true" style="padding:4px 8px;border-bottom:1px solid #fde">[Target 2]</td><td contenteditable="true" style="padding:4px 8px;border-bottom:1px solid #fde">[Realistic 2]</td><td contenteditable="true" style="padding:4px 8px;border-bottom:1px solid #fde">[Timeline 2]</td></tr>\n' +
    '<tr><td style="padding:4px 8px">3</td><td contenteditable="true" style="padding:4px 8px">[Sales behavior 3]</td><td contenteditable="true" style="padding:4px 8px">[Measurement 3]</td><td contenteditable="true" style="padding:4px 8px">[Target 3]</td><td contenteditable="true" style="padding:4px 8px">[Realistic 3]</td><td contenteditable="true" style="padding:4px 8px">[Timeline 3]</td></tr>\n' +
    '</tbody></table>\n' +
    '<p contenteditable="true" style="font-size:13px;">&#127775; Manager Coaching Tip: [Specific tip for the Team Leader on how to coach this agent — what to reinforce and what to redirect]</p>\n' +
    '</div>\n\n' +
    '<!-- COACHING TAKEAWAYS SUMMARY — SECOND AT THE TOP -->\n' +
    '<div class="ai-section" style="background:#E8F5E9;border:2px solid #2B8000;border-radius:8px;padding:14px 18px;margin-bottom:20px">\n' +
    '<div class="ai-title" style="background:#2B8000;color:#fff;border-left:4px solid #1B5E20;">&#127979; Key Coaching Takeaways</div>\n' +
    '<ol class="ai-coaching" style="padding-left:20px;font-size:13px;line-height:2.1">\n' +
    '<li contenteditable="true">[Most impactful sales coaching takeaway — cite a verbatim moment from this call + the SMART action to replace it]</li>\n' +
    '<li contenteditable="true">[Second takeaway — specific, actionable, and measurable]</li>\n' +
    '<li contenteditable="true">[Third takeaway]</li>\n' +
    '</ol>\n' +
    '</div>\n\n' +
    '<!-- SCORE PANEL -->\n' +
    '<div class="ai-section">\n' +
    '<div class="ai-title">&#128200; Overall Performance Score</div>\n' +
    '<div class="ai-score-panel">\n' +
    '  <div class="ai-score-row"><span>Building Rapport</span><span class="ai-badge [ai-badge-good OR ai-badge-mid OR ai-badge-bad]">[0-5]/5</span></div>\n' +
    '  <div class="ai-score-row"><span>Needs Identification</span><span class="ai-badge [class]">[0-5]/5</span></div>\n' +
    '  <div class="ai-score-row"><span>Product Presentation</span><span class="ai-badge [class]">[0-5]/5</span></div>\n' +
    '  <div class="ai-score-row"><span>Objection Handling</span><span class="ai-badge [class]">[0-5]/5</span></div>\n' +
    '  <div class="ai-score-row"><span>Closing Techniques</span><span class="ai-badge [class]">[0-5]/5</span></div>\n' +
    '  <div class="ai-score-row"><span><strong>Total Score</strong></span><span class="ai-badge [class]"><strong>[avg]/5</strong></span></div>\n' +
    '</div>\n' +
    '</div>\n\n' +
    '<!-- CALL SUMMARY -->\n' +
    '<div class="ai-section">\n' +
    '<div class="ai-title">&#128222; Call Summary</div>\n' +
    '<div class="ai-summary">[max 200-word linear description of customer experience]</div>\n' +
    '</div>\n\n' +
    '<!-- INTERACTION OVERVIEW -->\n' +
    '<div class="ai-section">\n' +
    '<div class="ai-title">&#128100; Interaction Overview</div>\n' +
    '<div class="ai-info">\n' +
    '  <div class="ai-chip"><span class="ai-chip-label">Sale Occurred</span><span class="ai-chip-val">[Yes/No]</span></div>\n' +
    '  <div class="ai-chip"><span class="ai-chip-label">Product Sold</span><span class="ai-chip-val">[name or N/A]</span></div>\n' +
    '  <div class="ai-chip"><span class="ai-chip-label">Sale Initiator</span><span class="ai-chip-val">[Agent/Client]</span></div>\n' +
    '</div>\n' +
    '[If no sale: <div class="ai-perfect"><strong>Perfect Sales Moment:</strong> [where in transcript + suggested statement]</div>]\n' +
    '</div>\n\n' +
    '<!-- ONE TABLE PER FRAMEWORK ELEMENT -->\n' +
    'Repeat this block for EACH of the 5 elements: Building Rapport, Needs Identification, Product Presentation, Objection Handling, Closing Techniques.\n' +
    'Each table now has 5 columns — Skill Name, Score, Analysis & Strengths, Areas of Opportunity, AND a new Recommendation + Sample Positioning Statement column.\n\n' +
    '<div class="ai-section">\n' +
    '<div class="ai-title">[Element Name]</div>\n' +
    '<table class="ai-table">\n' +
    '<thead><tr>\n' +
    '  <th style="width:18%">Skill Name</th>\n' +
    '  <th style="width:6%">Score</th>\n' +
    '  <th style="width:20%">Analysis &amp; Strengths</th>\n' +
    '  <th style="width:20%">Areas of Opportunity</th>\n' +
    '  <th style="width:36%">SMART Recommendation &amp; Sample Positioning Statement</th>\n' +
    '</tr></thead>\n' +
    '<tbody>\n' +
    '[One row per subelement — fill all 5 cells with actual analysis:]\n' +
    '<tr>\n' +
    '  <td class="ai-label-col">[subelement name]</td>\n' +
    '  <td><span class="ai-badge [class]">[0-5]/5</span></td>\n' +
    '  <td contenteditable="true">[what agent did well + verbatim quote]</td>\n' +
    '  <td contenteditable="true">[specific gap + verbatim missed moment]</td>\n' +
    '  <td contenteditable="true"><strong>Coaching:</strong> [1-sentence SMART action — exact behavior, how to measure, by when]<br/>&#127908; <em>"[Roleplay — max 25 words]"</em></td>\n' +
    '</tr>\n' +
    '</tbody>\n' +
    '</table>\n' +
    '</div>\n\n' +
    '<!-- CRITICAL FLAGS -->\n' +
    '<div class="ai-section">\n' +
    '<div class="ai-title">&#128681; Critical Flags &amp; Positioning Statements</div>\n' +
    '[For each critical flag:]\n' +
    '<div class="ai-flag">\n' +
    '  <div class="ai-flag-title">&#9888; [Missed parameter or behavior]</div>\n' +
    '  <div class="ai-flag-detail">[Detailed explanation of what was missed and why it matters to the customer and the sale]</div>\n' +
    '  <div style="background:#FFF3E0;border:1px solid #FFB74D;border-radius:4px;padding:8px 12px;margin:8px 0;font-size:12px">\n' +
    '    <strong style="color:#E65100">&#127919; SMART Coaching Goal:</strong><br/>\n' +
    '    <span contenteditable="true">S: [Specific sales behavior to change] | M: [How success is measured — e.g. offer made on every eligible call] | A: [Achievable target] | R: [Why realistic — e.g. agent understands the product, simple technique to apply] | T: [Timeline — e.g. within 2 coaching sessions / by next QA review]</span>\n' +
    '  </div>\n' +
    '  <div class="ai-flag-rl-label">&#127908; Sample Positioning Statement — Roleplay &amp; Practice</div>\n' +
    '  <div class="ai-flag-stmt" contenteditable="true">"[Complete, roleplay-ready statement the coach can say verbatim — e.g. \'[Agent], when the customer says X, try: [exact sales phrase]. This makes the offer feel relevant and personal — and it closes more naturally.\']"</div>\n' +
    '</div>\n' +
    '</div>\n\n' +
    '<!-- HIGHLIGHTS AND LOWLIGHTS -->\n' +
    '<div class="ai-section">\n' +
    '<div class="ai-title">&#9733; Highlights &amp; Lowlights — This Call</div>\n' +
    '<div class="ai-hl-grid">\n' +
    '<div class="ai-hl-box ai-hl-high"><div class="ai-hl-title">&#10003; Highlights</div><ul>\n' +
    '<li contenteditable="true">[Specific strength from this call with verbatim example]</li>\n' +
    '<li contenteditable="true">[Another highlight — be specific about what worked well]</li>\n' +
    '<li contenteditable="true">[Third highlight]</li>\n' +
    '</ul></div>\n' +
    '<div class="ai-hl-box ai-hl-low"><div class="ai-hl-title">&#10007; Lowlights</div><ul>\n' +
    '<li contenteditable="true">[Specific missed opportunity with verbatim example of what was said vs what should have been said]</li>\n' +
    '<li contenteditable="true">[Another lowlight — be specific]</li>\n' +
    '<li contenteditable="true">[Third lowlight]</li>\n' +
    '</ul></div>\n' +
    '</div>\n' +
    '</div>\n\n' +
    'Badge class rules: score 3-5 = ai-badge-good (green), score 2-2.9 = ai-badge-mid (amber), score 1-1.9 = ai-badge-bad (red), score 0 = ai-badge-zero (grey)\n' +
    'CRITICAL: Every framework table MUST have 5 columns including the Recommendation & Sample Positioning Statement column.\n' +
    'REPLACE all [placeholder] text with actual analysis from the transcript.\n' +
    'Do NOT include any text outside the HTML tags.\n' +
    'Make all td and li elements contenteditable="true".\n\n' +
    kb +
    'TRANSCRIPT:\n\n' + transcriptText;
}

// ── Main orchestrator ─────────────────────────────────────────────────────────
function analyzeTranscript(transcriptText, analysisType) {
  var knowledgeText = '';
  try {
    knowledgeText = analysisType === 'sales'
      ? getPDFKnowledgeFromFolder(SALES_FOLDER_ID,   PDF_CACHE_KEY_SALES)
      : getPDFKnowledgeFromFolder(REPEATS_FOLDER_ID, PDF_CACHE_KEY_REPEATS);
  } catch(e) { Logger.log('PDF fetch failed (non-fatal): ' + e); }

  var prompt = analysisType === 'sales'
    ? buildSalesPrompt(transcriptText, knowledgeText)
    : buildRepeatsPrompt(transcriptText, knowledgeText);

  return callFuelIX(prompt);
}

