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
    // Skip if payload exceeds 100KB — truncated JSON silently breaks every
    // subsequent parse, causing an unbounded sheet read on every invocation.
    var _rs = JSON.stringify(serializable);
    if (_rs.length <= 99000) { try { cache.put(cacheKey, _rs, 4 * 60 * 60); } catch(e) {} }
    _rosterDataInMemory = serializable;
    return serializable;
  } catch(e) { Logger.log('_getRosterSheetData: ' + e); return []; }
}

// ── Pre-warm Global Roster + AT Data GCP caches on page load ─────────────────
// Called fire-and-forget from the client so autofill is fast when transcript is pasted.
function warmLookupCaches() {
  try { _getGlobalRosterData(); } catch(e) {}
  try { _getATDataGCPSheetData(); } catch(e) {}
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
    var _rr = JSON.stringify(result);
    if (_rr.length <= 99000) { try { cache.put(cacheKey, _rr, 4 * 60 * 60); } catch(e) {} }
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
    // Store the full 2D array — callers use data[0] for header and data[i] for rows
    var _ad = JSON.stringify(data);
    if (_ad.length <= 99000) { try { cache.put(cacheKey, _ad, 4 * 60 * 60); } catch(e) {} }
    _atDataInMemory = data;
    Logger.log('AT Data GCP loaded from sheet: ' + (data.length - 1) + ' rows');
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
    var _gp = JSON.stringify(payload);
    if (_gp.length <= 99000) { try { cache.put(cacheKey, _gp, 4 * 60 * 60); } catch(e) {} }
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
          agentEmail:     lookupAgentEmail(agentName) || resolveEmail(agentName)
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
          agentEmail:     lookupAgentEmail(name) || resolveEmail(name)
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

// ── Speaker name matching helper ──────────────────────────────────────────────
function _nameMatches(a, b) {
  if (!a || !b) return false;
  if (a === b) return true;
  if (a.length > 4 && b.indexOf(a) !== -1) return true;
  if (b.length > 4 && a.indexOf(b) !== -1) return true;
  return false;
}

// ── Build a set of confirmed real speakers from a pre-scan ────────────────────
// Real speakers appear 2+ times OR are named in the Internal Participant(s) header.
// This prevents header fields ("Interaction ID:") and inline colons from being
// treated as speaker turns.
function _buildKnownSpeakers(lines, allKnownAgentsLower) {
  var counts = {};
  for (var p = 0; p < lines.length; p++) {
    var m = lines[p].match(/^([^:\n]{2,60}):\s/);
    if (m) {
      var sp = m[1].trim().toLowerCase();
      counts[sp] = (counts[sp] || 0) + 1;
    }
  }
  var known = {};
  for (var sp in counts) {
    var isAgent = allKnownAgentsLower.some(function(a) { return _nameMatches(a, sp); });
    if (counts[sp] >= 2 || isAgent) known[sp] = true;
  }
  return known;
}

// ── Main transcript filter ────────────────────────────────────────────────────
// Section-based + turn-based hybrid:
//   1. Always keep the metadata header (before any known speaker turn).
//   2. Skip everything from the first known speaker until the TARGET agent
//      first appears (pre-handoff content — other agents + customer before transfer).
//   3. From the target's first turn onwards: keep target + customer lines,
//      remove other internal agents' turns.
// Returns { filtered: <string>, stats: { totalLines, keptLines, skippedPreHandoff,
//           excludedAgents: [], targetFirstLine: <string|null>, targetFound: <bool> } }
function filterTranscriptByAgent(transcriptText, targetAgentName) {
  var NO_FILTER = { filtered: transcriptText, stats: { targetFound: false, applied: false } };
  if (!targetAgentName || !transcriptText) return NO_FILTER;

  var targetLower = targetAgentName.toLowerCase().trim();

  // Parse other internal agents from header
  var otherAgents = [];
  var pMatch = transcriptText.match(/Internal Participant\(s\)[:\s]+([^\n\r]+)/i);
  if (pMatch) {
    otherAgents = pMatch[1].split(/[,;]+/)
      .map(function(n) { return n.trim().toLowerCase(); })
      .filter(function(n) {
        if (!n) return false;
        return !_nameMatches(n, targetLower);
      });
  }

  if (otherAgents.length === 0) return NO_FILTER;

  var lines        = transcriptText.split('\n');
  var allKnown     = otherAgents.concat([targetLower]);
  var knownSpkrs   = _buildKnownSpeakers(lines, allKnown);

  // ── Find boundary indices ─────────────────────────────────────────────────
  // convStart: first line of actual conversation (first known-speaker turn)
  // targetStart: first line where TARGET agent speaks
  var convStart   = -1;
  var targetStart = -1;

  for (var i = 0; i < lines.length; i++) {
    var m = lines[i].match(/^([^:\n]{2,60}):\s/);
    if (!m) continue;
    var sp = m[1].trim().toLowerCase();
    if (!knownSpkrs[sp]) continue;              // not a real speaker — skip

    if (convStart === -1) convStart = i;        // first known-speaker line

    if (_nameMatches(sp, targetLower)) {
      targetStart = i;
      break;
    }
  }

  // Target never speaks — can't filter meaningfully
  if (targetStart === -1) {
    Logger.log('filterTranscriptByAgent: target "' + targetAgentName + '" not found in transcript — no filter applied');
    return NO_FILTER;
  }

  // ── Apply filtering ───────────────────────────────────────────────────────
  var result            = [];
  var skipBlock         = false;
  var skippedPreHandoff = 0;
  var excludedPerAgent  = {};

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];

    // Phase 1: header (before any known speaker) — always keep
    if (convStart === -1 || i < convStart) {
      result.push(line);
      continue;
    }

    // Phase 2: pre-handoff (between convStart and targetStart) — always skip
    if (i < targetStart) {
      skippedPreHandoff++;
      continue;
    }

    // Phase 3: from target's first turn — turn-based filter
    var m = line.match(/^([^:\n]{2,60}):\s/);
    if (m) {
      var sp = m[1].trim().toLowerCase();
      if (knownSpkrs[sp]) {
        var isOther = otherAgents.some(function(a) { return _nameMatches(a, sp); });
        skipBlock = isOther;
        // Do NOT count here — the else branch below counts every excluded line
        // (including this speaker-start line), avoiding double-counting.
      }
    }

    if (!skipBlock) {
      result.push(line);
    } else {
      // Count every excluded line once: speaker-start lines + continuation lines.
      // Speaker-start: origName2 = agent name. Continuation: origName2 = '(cont.)'.
      var origName2 = (line.match(/^([^:\n]{2,60}):\s/) || [])[1] || '(cont.)';
      excludedPerAgent[origName2.trim()] = (excludedPerAgent[origName2.trim()] || 0) + 1;
    }
  }

  var keptLines = result.length;
  Logger.log('filterTranscriptByAgent: ' + keptLines + '/' + lines.length + ' lines kept for "' +
             targetAgentName + '" | pre-handoff skipped: ' + skippedPreHandoff +
             ' | excluded agents: ' + JSON.stringify(excludedPerAgent));

  var excludedNames = Object.keys(excludedPerAgent).filter(function(k) { return k !== '(cont.)'; });

  return {
    filtered: result.join('\n'),
    stats: {
      applied:           true,
      targetFound:       true,
      totalLines:        lines.length,
      keptLines:         keptLines,
      skippedPreHandoff: skippedPreHandoff,
      excludedAgents:    excludedNames,
      targetFirstLine:   lines[targetStart] || null
    }
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// CSS shared by both prompts (injected into the AI output)
// ─────────────────────────────────────────────────────────────────────────────
function sharedCSS() {
  return '<style>' +
    // ── Base reset & font ──────────────────────────────────────────────────────
    '.report-wrap *{box-sizing:border-box}' +
    // ── Header ────────────────────────────────────────────────────────────────
    '.report-header{background:linear-gradient(135deg,#4B286D 0%,#7B4FA0 100%);' +
      'border-radius:12px;padding:24px 28px;margin-bottom:16px;color:#fff;' +
      'font-family:Helvetica Neue,Helvetica,Arial,sans-serif}' +
    '.report-header-title{font-size:20px;font-weight:800;letter-spacing:.2px;margin-bottom:4px}' +
    '.report-header-sub{font-size:13px;opacity:.85;font-weight:400}' +
    // ── Meta chips (reuse old ai-chip classes for extraction compatibility) ────
    '.ai-info{display:flex;flex-wrap:wrap;gap:9px;margin-bottom:16px}' +
    '.ai-chip{background:#F4F4F7;border:1px solid #D8D8D8;border-radius:6px;' +
      'padding:8px 12px;min-width:100px;font-family:Helvetica Neue,Helvetica,Arial,sans-serif}' +
    '.ai-chip-label{font-size:10px;font-weight:700;color:#54565A;text-transform:uppercase;' +
      'letter-spacing:.4px;display:block;margin-bottom:3px}' +
    '.ai-chip-val{font-size:13px;font-weight:600;color:#1A1A2E}' +
    '.report-badge-yes{display:inline-block;background:#2B8000;color:#fff;border-radius:4px;' +
      'padding:2px 10px;font-size:12px;font-weight:700}' +
    '.report-badge-no{display:inline-block;background:#C12335;color:#fff;border-radius:4px;' +
      'padding:2px 10px;font-size:12px;font-weight:700}' +
    // ── Call summary section ───────────────────────────────────────────────────
    '.report-summary-wrap{background:#F9F9F9;border:1px solid #E0E0E0;border-radius:8px;' +
      'padding:16px 20px;margin-bottom:16px;font-family:Helvetica Neue,Helvetica,Arial,sans-serif}' +
    '.report-summary-label{font-size:11px;font-weight:700;color:#4B286D;' +
      'text-transform:uppercase;letter-spacing:.6px;margin-bottom:8px}' +
    // ai-summary kept for extractTextBlock() compatibility
    '.ai-summary{font-size:13px;line-height:1.75;color:#1A1A2E;' +
      'min-height:48px;outline:none;padding:2px 0}' +
    '.report-summary-meta{font-size:12px;color:#767676;margin-top:10px;' +
      'padding-top:8px;border-top:1px solid #E8E8E8}' +
    // ── 3-column coaching grid ─────────────────────────────────────────────────
    '.report-3col{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-bottom:20px}' +
    '.report-col{border-radius:8px;padding:16px;font-family:Helvetica Neue,Helvetica,Arial,sans-serif}' +
    '.report-col-working{background:#EDF7E6;border:1px solid #B3DFA0}' +
    '.report-col-change{background:#FFF5F5;border:1px solid #F5AAAA}' +
    '.report-col-howto{background:#F5F0FF;border:1px solid #D1B8E8}' +
    '.report-col-head{font-size:12px;font-weight:700;margin-bottom:10px;' +
      'padding-bottom:6px;border-bottom:2px solid rgba(0,0,0,.08)}' +
    '.report-col-working .report-col-head{color:#2B8000}' +
    '.report-col-change .report-col-head{color:#C12335}' +
    '.report-col-howto .report-col-head{color:#4B286D}' +
    '.report-col-list{padding-left:18px;font-size:13px;line-height:1.8;margin:0}' +
    '.report-col-list li{margin-bottom:6px}' +
    '.report-col-roleplays{list-style:none;padding-left:0}' +
    '.report-col-roleplays li{margin-bottom:12px}' +
    // ── AI Spotted Flags section ───────────────────────────────────────────────
    '.report-flags{margin-bottom:16px}' +
    '.report-flags-head{font-size:13px;font-weight:700;color:#C12335;' +
      'padding:10px 14px;background:#FFF0F0;border:1px solid #F5AAAA;' +
      'border-radius:6px 6px 0 0;border-bottom:none}' +
    '.report-flags-body{border:1px solid #F5AAAA;border-top:none;' +
      'border-radius:0 0 6px 6px;padding:12px;background:#fff}' +
    // ── Manually Added Flags section ───────────────────────────────────────────
    '.report-manual-flags{margin-bottom:16px}' +
    '.report-manual-flags-head{font-size:13px;font-weight:700;color:#4B286D;' +
      'padding:10px 14px;background:#F5F0FF;border:1px solid #D1B8E8;' +
      'border-radius:6px 6px 0 0;border-bottom:none}' +
    '.report-manual-flags-body{border:1px solid #D1B8E8;border-top:none;' +
      'border-radius:0 0 6px 6px;padding:12px;background:#fff;min-height:36px}' +
    '.report-manual-empty{color:#767676;font-size:12px;font-style:italic;padding:4px 0}' +
    // ── Footer ────────────────────────────────────────────────────────────────
    '.report-footer{background:#F4F4F7;border:1px solid #E0E0E0;border-radius:8px;' +
      'padding:12px 20px;display:flex;justify-content:space-between;align-items:center;' +
      'font-size:12px;color:#54565A;margin-top:8px;' +
      'font-family:Helvetica Neue,Helvetica,Arial,sans-serif}' +
    '.report-footer-ref{font-weight:700;color:#4B286D}' +
    '.report-footer-obs{color:#767676}' +
    // ── Flag cards (kept for _injectFlagDeleteBtns + Dashboard extraction) ──────
    '.ai-flag{background:#FFF5F5;border:1px solid #F5AAAA;border-left:4px solid #C12335;' +
      'border-radius:6px;padding:14px 16px;margin-bottom:11px;position:relative;' +
      'font-family:Helvetica Neue,Helvetica,Arial,sans-serif}' +
    '.ai-flag-title{font-weight:700;color:#C12335;font-size:13px;margin-bottom:5px}' +
    '.ai-flag-detail{font-size:13px;color:#444;line-height:1.65;outline:none}' +
    // ── Roleplay statements (kept for initPlayButtons TTS) ─────────────────────
    '.ai-flag-rl-label{font-size:10px;font-weight:700;text-transform:uppercase;' +
      'color:#4B286D;letter-spacing:.5px;margin-bottom:4px;margin-top:4px}' +
    '.ai-flag-stmt{background:#F0EAF8;border:1px solid #C5A8E8;border-radius:4px;' +
      'padding:9px 12px;font-size:13px;color:#3A1060;font-style:italic;' +
      'line-height:1.6;outline:none}' +
    // ── Legacy classes (kept for old cached reports) ───────────────────────────
    '.ai-section{margin-bottom:22px}' +
    '.ai-title{font-size:13px;font-weight:700;color:#4B286D;padding:8px 12px;' +
      'background:#F5F0FF;border-left:4px solid #4B286D;border-radius:0 4px 4px 0;' +
      'margin-bottom:11px;font-family:Helvetica Neue,Helvetica,Arial,sans-serif}' +
    '.ai-table{width:100%;border-collapse:collapse;font-size:13px;' +
      'font-family:Helvetica Neue,Helvetica,Arial,sans-serif;margin-bottom:4px}' +
    '.ai-table thead tr{background:#4B286D}' +
    '.ai-table th{padding:10px 14px;text-align:left;font-size:12px;font-weight:700;color:#fff;letter-spacing:.3px}' +
    '.ai-table td{padding:10px 14px;border-bottom:1px solid #EBEBEB;vertical-align:top;line-height:1.65;font-size:13px}' +
    '.ai-table tbody tr:nth-child(even) td{background:#FAFAFA}' +
    '.ai-table tbody tr:last-child td{border-bottom:none}' +
    '.ai-label-col{font-weight:600;color:#1A1A2E;width:22%;white-space:nowrap}' +
    '.ai-hl-grid{display:grid;grid-template-columns:1fr 1fr;gap:13px}' +
    '.ai-hl-box{border-radius:6px;padding:14px 16px;font-family:Helvetica Neue,Helvetica,Arial,sans-serif}' +
    '.ai-hl-high{background:#EDF7E6;border:1px solid #B3DFA0}' +
    '.ai-hl-low{background:#FFF5F5;border:1px solid #F5AAAA}' +
    '.ai-hl-title{font-size:12px;font-weight:700;margin-bottom:8px}' +
    '.ai-hl-high .ai-hl-title{color:#2B8000}' +
    '.ai-hl-low  .ai-hl-title{color:#C12335}' +
    '.ai-hl-box ul{padding-left:16px;font-size:13px;line-height:1.9}' +
    '.ai-coaching{padding-left:20px;font-size:13px;line-height:1.9;font-family:Helvetica Neue,Helvetica,Arial,sans-serif}' +
    '.ai-coaching li{margin-bottom:4px}' +
    '.ai-score-panel{background:#F5F0FF;border:1px solid #D1B8E8;border-radius:6px;padding:14px 18px;font-family:Helvetica Neue,Helvetica,Arial,sans-serif}' +
    '.ai-score-row{display:flex;align-items:center;justify-content:space-between;padding:6px 0;border-bottom:1px solid #E4D8F5;font-size:13px}' +
    '.ai-score-row:last-child{border-bottom:none;font-weight:700}' +
    '.ai-badge{display:inline-block;border-radius:4px;padding:3px 10px;font-size:12px;font-weight:700;color:#fff;min-width:44px;text-align:center}' +
    '.ai-badge-good{background:#2B8000}' +
    '.ai-badge-mid{background:#8C4A00}' +
    '.ai-badge-bad{background:#C12335}' +
    '.ai-call-badge{display:inline-block;background:#4B286D;color:#fff;font-size:11px;font-weight:700;padding:3px 12px;border-radius:12px;margin-bottom:13px;letter-spacing:.5px;font-family:Helvetica Neue,Helvetica,Arial,sans-serif}' +
    '.ai-call-block{border:1px solid #D8D8D8;border-radius:6px;padding:18px 20px;margin-bottom:20px;background:#fff}' +
    '.ai-perfect{background:#EDF7E6;border:1px solid #B3DFA0;border-radius:4px;padding:10px 14px;font-size:13px;line-height:1.6;margin-top:10px;font-family:Helvetica Neue,Helvetica,Arial,sans-serif}' +
    '.ai-warning{background:#FFF0F0;border:1px solid #F5AAAA;border-left:4px solid #C12335;border-radius:4px;padding:12px 16px;font-size:13px;color:#C12335;margin-top:8px}' +
    '</style>';
}

// ─────────────────────────────────────────────────────────────────────────────
// REPEATS PROMPT — asks AI to return complete HTML
// ─────────────────────────────────────────────────────────────────────────────
function buildRepeatsPrompt(transcriptText, knowledgeText, agentName, selectedLOB) {
  var kb = knowledgeText
    ? '\n\nCOMPANY POLICIES AND PROCEDURES:\n' + knowledgeText + '\n\n'
    : '';
  var lobKb = selectedLOB ? getLOBKnowledge(selectedLOB) : '';
  var lobBlock = lobKb
    ? '\n\nLOB & ROLE-SPECIFIC EVALUATION GUIDELINES:\nAgent role: ' + selectedLOB + '. Flag any process deviations prominently in the AI Spotted Flags section.\n\n' + lobKb + '\n\n'
    : '';
  var focusLine = agentName
    ? 'IMPORTANT: Evaluate ONLY the performance of ' + agentName + '. Other agents in the transcript are context only.\n\n'
    : '';

  return 'You are a Quality Analyst. Analyze the call transcript(s) below.\n' +
    focusLine +
    'PURPOSE: Identify FCR opportunities, reduce repeat call rate and transfer rate.\n\n' +
    'CX BLUEPRINT LENS:\n' +
    '• ENGAGE: Acknowledge the OCCURRENCE; validate the EMOTION — not scripted apologies.\n' +
    '• UNDERSTAND: Ask open-ended questions to EXPLORE before confirming.\n' +
    '• SOLVE: BRIDGE solutions to why they matter to THIS customer specifically.\n' +
    '• IMPRESS: Set expectations PROACTIVELY — customer leaves informed, not just answered.\n\n' +
    'CRITICAL INSTRUCTION: Return ONLY valid HTML. No markdown. No text outside the tags.\n' +
    'Use EXACTLY these CSS classes in this exact structure:\n\n' +
    '<div class="report-wrap">\n\n' +

    '<!-- HEADER -->\n' +
    '<div class="report-header">\n' +
    '  <div class="report-header-title">&#128204; TELUS NH Analyzer</div>\n' +
    '  <div class="report-header-sub">Repeats &amp; Transfer Audit Report</div>\n' +
    '</div>\n\n' +

    '<!-- META ROW: fill values from transcript -->\n' +
    '<div class="ai-info">\n' +
    '  <div class="ai-chip">' +
    '<span class="ai-chip-label">Agent Name</span>' +
    '<span class="ai-chip-val">[agent full name]</span></div>\n' +
    '  <div class="ai-chip">' +
    '<span class="ai-chip-label">Call Date</span>' +
    '<span class="ai-chip-val">[date from transcript]</span></div>\n' +
    '  <div class="ai-chip">' +
    '<span class="ai-chip-label">Interaction Date</span>' +
    '<span class="ai-chip-val">[interaction date or N/A]</span></div>\n' +
    '  <div class="ai-chip">' +
    '<span class="ai-chip-label">Issue Resolution</span>' +
    '<span class="ai-chip-val"><span class="report-badge-[yes|no]">[Yes or No]</span></span></div>\n' +
    '  <div class="ai-chip">' +
    '<span class="ai-chip-label">Repeat Risk</span>' +
    '<span class="ai-chip-val">[0-100]%</span></div>\n' +
    '  <div class="ai-chip">' +
    '<span class="ai-chip-label">Audit Reference</span>' +
    '<span class="ai-chip-val">__AUDIT_REF__</span></div>\n' +
    '</div>\n\n' +

    '<!-- CALL SUMMARY & KEY INTERACTION DETAILS -->\n' +
    '<div class="report-summary-wrap">\n' +
    '  <div class="report-summary-label">&#128222; Call Summary &amp; Key Interaction Details</div>\n' +
    '  <div class="ai-summary" contenteditable="true">' +
    '[100-word summary: customer reason for call, key moments, outcome, ' +
    'any transfers or callbacks, tone of interaction]' +
    '</div>\n' +
    '  <div class="report-summary-meta">' +
    'Duration: [X min] &nbsp;|&nbsp; Direction: [Inbound/Outbound] &nbsp;|&nbsp; ' +
    'Transfer: [Yes/No]' +
    '</div>\n' +
    '</div>\n\n' +

    '<!-- 3-COLUMN COACHING GRID -->\n' +
    '<div class="report-3col">\n\n' +

    '  <!-- COLUMN 1: What\'s Working — exactly 2 items -->\n' +
    '  <div class="report-col report-col-working">\n' +
    '    <div class="report-col-head">&#9989; What\'s Working</div>\n' +
    '    <ul class="report-col-list">\n' +
    '      <li contenteditable="true">' +
    '[Most impactful highlight — specific behavior with a brief example or verbatim quote]' +
    '</li>\n' +
    '      <li contenteditable="true">' +
    '[Second specific strength from this call]' +
    '</li>\n' +
    '    </ul>\n' +
    '  </div>\n\n' +

    '  <!-- COLUMN 2: What Needs to Change — exactly 2 items -->\n' +
    '  <div class="report-col report-col-change">\n' +
    '    <div class="report-col-head">&#128205; What Needs to Change</div>\n' +
    '    <ul class="report-col-list">\n' +
    '      <li contenteditable="true">' +
    '[Most critical behavior — specific and actionable; name why it increases repeat call risk]' +
    '</li>\n' +
    '      <li contenteditable="true">' +
    '[Second behavior — specific and tied to FCR or customer experience]' +
    '</li>\n' +
    '    </ul>\n' +
    '  </div>\n\n' +

    '  <!-- COLUMN 3: How to Change It: Roleplays and Samples -->\n' +
    '  <div class="report-col report-col-howto">\n' +
    '    <div class="report-col-head">&#127908; How to Change It: Roleplays and Samples</div>\n' +
    '    <ul class="report-col-list report-col-roleplays">\n' +
    '      <li>\n' +
    '        <div class="ai-flag-rl-label">Roleplay Scenario 1</div>\n' +
    '        <div class="ai-flag-stmt" contenteditable="true">' +
    '"[Verbatim statement addressing What Needs to Change item 1 — complete, natural, roleplay-ready]"' +
    '</div>\n' +
    '      </li>\n' +
    '      <li>\n' +
    '        <div class="ai-flag-rl-label">Roleplay Scenario 2</div>\n' +
    '        <div class="ai-flag-stmt" contenteditable="true">' +
    '"[Verbatim statement addressing What Needs to Change item 2]"' +
    '</div>\n' +
    '      </li>\n' +
    '    </ul>\n' +
    '  </div>\n\n' +

    '</div>\n\n' +

    '<!-- AI SPOTTED FLAGS: 2-5 flags; ai-flag class = X-button works; ai-flag-title = Dashboard extraction -->\n' +
    '<div class="report-flags">\n' +
    '  <div class="report-flags-head">&#128681; AI Spotted Flags</div>\n' +
    '  <div class="report-flags-body">\n' +
    '    [Repeat for EACH critical flag (min 2, max 5):]\n' +
    '    <div class="ai-flag">\n' +
    '      <div class="ai-flag-title">&#9888; [Specific missed behavior or policy deviation]</div>\n' +
    '      <div class="ai-flag-detail" contenteditable="true">' +
    '[2-3 sentences: what was missed, why it matters for FCR, verbatim example from transcript]' +
    '</div>\n' +
    '    </div>\n' +
    '  </div>\n' +
    '</div>\n\n' +

    '<!-- FOOTER -->\n' +
    '<div class="report-footer">\n' +
    '  <span class="report-footer-ref">Audit Ref: __AUDIT_REF__</span>\n' +
    '  <span class="report-footer-obs">Observer: __OBSERVER__</span>\n' +
    '</div>\n\n' +

    '</div>\n\n' +

    'RULES:\n' +
    '1. Replace ALL [placeholder] text with real analysis from the transcript.\n' +
    '2. Issue Resolution badge: class="report-badge-yes" for Yes, class="report-badge-no" for No.\n' +
    '3. Keep __AUDIT_REF__ and __OBSERVER__ as literal text — do NOT replace them.\n' +
    '4. Keep contenteditable="true" on all elements shown with it.\n' +
    '5. Exactly 2 li items in What\'s Working and What Needs to Change.\n' +
    '6. AI Spotted Flags: 2-5 flags. Each title MUST have class="ai-flag-title".\n' +
    '7. Do NOT include any text outside the HTML. Do NOT use markdown.\n\n' +
    kb + lobBlock +
    'TRANSCRIPT:\n\n' + transcriptText;
}

// ─────────────────────────────────────────────────────────────────────────────
// SALES PROMPT — asks AI to return complete HTML
// ─────────────────────────────────────────────────────────────────────────────
function buildSalesPrompt(transcriptText, knowledgeText, agentName, selectedLOB) {
  var kb = knowledgeText
    ? '\n\nCOMPANY POLICIES AND PROCEDURES:\n' + knowledgeText + '\n\n'
    : '';
  var lobKb = selectedLOB ? getLOBKnowledge(selectedLOB) : '';
  var lobBlock = lobKb
    ? '\n\nLOB & ROLE-SPECIFIC EVALUATION GUIDELINES:\nAgent role: ' + selectedLOB + '. Flag any deviations prominently in the AI Spotted Flags section.\n\n' + lobKb + '\n\n'
    : '';
  var focusLine = agentName
    ? 'IMPORTANT: Evaluate ONLY the sales performance of ' + agentName + '. Other agents in the transcript are context only.\n\n'
    : '';

  return 'You are an expert sales performance analyst for TELUS. Evaluate the call transcript.\n' +
    focusLine +
    'PURPOSE: Identify sales opportunities and coach agents to increase sales.\n\n' +
    'CX BLUEPRINT LENS:\n' +
    '• ENGAGE: Acknowledge the OCCURRENCE; validate the EMOTION — not scripted apologies.\n' +
    '• UNDERSTAND: Ask open-ended questions to EXPLORE before confirming.\n' +
    '• SOLVE: BRIDGE solutions — show WHY they matter to THIS customer specifically.\n' +
    '• IMPRESS: Set expectations PROACTIVELY — customer leaves informed and confident.\n\n' +
    'CRITICAL INSTRUCTION: Return ONLY valid HTML. No markdown. No text outside the tags.\n' +
    'Use EXACTLY these CSS classes in this exact structure:\n\n' +
    '<div class="report-wrap">\n\n' +

    '<!-- HEADER -->\n' +
    '<div class="report-header">\n' +
    '  <div class="report-header-title">&#128204; TELUS NH Analyzer</div>\n' +
    '  <div class="report-header-sub">Sales Performance Audit Report</div>\n' +
    '</div>\n\n' +

    '<!-- META ROW: fill values from transcript -->\n' +
    '<div class="ai-info">\n' +
    '  <div class="ai-chip">' +
    '<span class="ai-chip-label">Agent Name</span>' +
    '<span class="ai-chip-val">[agent full name]</span></div>\n' +
    '  <div class="ai-chip">' +
    '<span class="ai-chip-label">Call Date</span>' +
    '<span class="ai-chip-val">[date from transcript]</span></div>\n' +
    '  <div class="ai-chip">' +
    '<span class="ai-chip-label">Interaction Date</span>' +
    '<span class="ai-chip-val">[interaction date or N/A]</span></div>\n' +
    '  <div class="ai-chip">' +
    '<span class="ai-chip-label">Issue Resolution</span>' +
    '<span class="ai-chip-val"><span class="report-badge-[yes|no]">[Yes or No]</span></span></div>\n' +
    '  <div class="ai-chip">' +
    '<span class="ai-chip-label">Pitched a Sale?</span>' +
    '<span class="ai-chip-val"><span class="report-badge-[yes|no]">[Yes or No]</span></span></div>\n' +
    '  <div class="ai-chip">' +
    '<span class="ai-chip-label">Audit Reference</span>' +
    '<span class="ai-chip-val">__AUDIT_REF__</span></div>\n' +
    '</div>\n\n' +

    '<!-- CALL SUMMARY & KEY INTERACTION DETAILS -->\n' +
    '<div class="report-summary-wrap">\n' +
    '  <div class="report-summary-label">&#128222; Call Summary &amp; Key Interaction Details</div>\n' +
    '  <div class="ai-summary" contenteditable="true">' +
    '[100-word summary: customer reason for call, key sales moments, ' +
    'whether a sale was attempted or completed, outcome, tone of interaction]' +
    '</div>\n' +
    '  <div class="report-summary-meta">' +
    'Duration: [X min] &nbsp;|&nbsp; Direction: [Inbound/Outbound] &nbsp;|&nbsp; ' +
    'Sale Outcome: [Sold / Not Sold / Attempted]' +
    '</div>\n' +
    '</div>\n\n' +

    '<!-- 3-COLUMN COACHING GRID -->\n' +
    '<div class="report-3col">\n\n' +

    '  <!-- COLUMN 1: What\'s Working — exactly 2 items -->\n' +
    '  <div class="report-col report-col-working">\n' +
    '    <div class="report-col-head">&#9989; What\'s Working</div>\n' +
    '    <ul class="report-col-list">\n' +
    '      <li contenteditable="true">' +
    '[Most impactful sales strength — specific behavior with example or verbatim quote]' +
    '</li>\n' +
    '      <li contenteditable="true">' +
    '[Second specific strength — what the agent did well that drove value]' +
    '</li>\n' +
    '    </ul>\n' +
    '  </div>\n\n' +

    '  <!-- COLUMN 2: What Needs to Change — exactly 2 items -->\n' +
    '  <div class="report-col report-col-change">\n' +
    '    <div class="report-col-head">&#128205; What Needs to Change</div>\n' +
    '    <ul class="report-col-list">\n' +
    '      <li contenteditable="true">' +
    '[Most critical sales gap — specific missed opportunity or technique; ' +
    'name why it cost the sale or damaged the customer experience]' +
    '</li>\n' +
    '      <li contenteditable="true">' +
    '[Second gap — specific behavior change tied to conversion or CX]' +
    '</li>\n' +
    '    </ul>\n' +
    '  </div>\n\n' +

    '  <!-- COLUMN 3: How to Change It: Roleplays and Samples -->\n' +
    '  <div class="report-col report-col-howto">\n' +
    '    <div class="report-col-head">&#127908; How to Change It: Roleplays and Samples</div>\n' +
    '    <ul class="report-col-list report-col-roleplays">\n' +
    '      <li>\n' +
    '        <div class="ai-flag-rl-label">Roleplay Scenario 1</div>\n' +
    '        <div class="ai-flag-stmt" contenteditable="true">' +
    '"[Verbatim sales statement addressing What Needs to Change item 1 — ' +
    'complete, natural, roleplay-ready, bridges value to this customer]"' +
    '</div>\n' +
    '      </li>\n' +
    '      <li>\n' +
    '        <div class="ai-flag-rl-label">Roleplay Scenario 2</div>\n' +
    '        <div class="ai-flag-stmt" contenteditable="true">' +
    '"[Verbatim sales statement addressing What Needs to Change item 2]"' +
    '</div>\n' +
    '      </li>\n' +
    '    </ul>\n' +
    '  </div>\n\n' +

    '</div>\n\n' +

    '<!-- AI SPOTTED FLAGS: 2-5 flags; ai-flag class = X-button works; ai-flag-title = Dashboard extraction -->\n' +
    '<div class="report-flags">\n' +
    '  <div class="report-flags-head">&#128681; AI Spotted Flags</div>\n' +
    '  <div class="report-flags-body">\n' +
    '    [Repeat for EACH critical flag (min 2, max 5):]\n' +
    '    <div class="ai-flag">\n' +
    '      <div class="ai-flag-title">&#9888; [Specific missed sales behavior or compliance deviation]</div>\n' +
    '      <div class="ai-flag-detail" contenteditable="true">' +
    '[2-3 sentences: what was missed, why it matters for the sale or customer trust, ' +
    'verbatim example from transcript]' +
    '</div>\n' +
    '    </div>\n' +
    '  </div>\n' +
    '</div>\n\n' +

    '<!-- FOOTER -->\n' +
    '<div class="report-footer">\n' +
    '  <span class="report-footer-ref">Audit Ref: __AUDIT_REF__</span>\n' +
    '  <span class="report-footer-obs">Observer: __OBSERVER__</span>\n' +
    '</div>\n\n' +

    '</div>\n\n' +

    'RULES:\n' +
    '1. Replace ALL [placeholder] text with real analysis from the transcript.\n' +
    '2. Issue Resolution badge and Pitched a Sale? badge: class="report-badge-yes" for Yes, class="report-badge-no" for No.\n' +
    '3. Keep __AUDIT_REF__ and __OBSERVER__ as literal text — do NOT replace them.\n' +
    '4. Keep contenteditable="true" on all elements shown with it.\n' +
    '5. Exactly 2 li items in What\'s Working and What Needs to Change.\n' +
    '6. AI Spotted Flags: 2-5 flags. Each title MUST have class="ai-flag-title".\n' +
    '7. Do NOT include any text outside the HTML. Do NOT use markdown.\n\n' +
    kb + lobBlock +
    'TRANSCRIPT:\n\n' + transcriptText;
}

// ── LOB-specific process evaluation knowledge ──────────────────────────────────
function getLOBKnowledge(selectedLOB) {
  var lob = (selectedLOB || '').toString().trim().toUpperCase().replace(/\s+/g,' ');

  if (lob === 'PF CXSS') {
    return 'LOB ROLE: PF CxSS — PureFibre Frontline Care\n\n' +
    'TRANSFER RULES (flag violations):\n' +
    '• Always probe and exhaust options before transferring. Never transfer without trying to resolve.\n' +
    '• To Retention/CLS: ONLY for genuine churn risk. INVALID: Copper Compass customers, non-churn issues, deceased, domestic violence, billing-only, tech support, natural disasters.\n' +
    '• Do NOT place any orders before transferring to Retention.\n' +
    '• Do NOT promise what Retention can do (never say "they can waive fees" or "they can match that offer").\n' +
    '• Always inform customer they are being transferred. Warm transfer: CRMT, PRC, Abusive customer, CLS1→CLS2 only. All others = informed (cold) transfer.\n' +
    '• Verify customer in Casa before transferring. Leave detailed account notes.\n' +
    '• Before transferring to TS: verify service is provisioned in CSR; confirm customer is at home with time for troubleshooting; probe first; select correct TS department.\n\n' +
    'PLATFORM ROUTING (flag if wrong):\n' +
    '• Compass customers → Compass CxSS. FIFA/PureFibre → FIFA CxSS.\n' +
    '• Cease/cancel on PureFibre FIFA → transfer to FFH PureFibre Loyalty CLS.\n' +
    '• Cease/cancel on Compass → transfer to FFH Copper Loyalty CLS.\n' +
    '• Escalations → Consumer Escalation Process → CRMT (WLN CRMT EN) if necessary.\n\n' +
    'SELF-SERVE POLICY (flag violations):\n' +
    '• Self-serve ONLY transactions (NOT available at call centre): payment arrangements, appointment reschedule, billing address change, My TELUS password reset, WiFi/SSID change, one-time payment, PIN change, e.bill to paper, pre-authorized payments.\n' +
    '• Check Self Serve Only indicator in Personal profile tab first.\n' +
    '• Required positioning: "This transaction now needs to be completed online. Would you like me to show you how?"\n' +
    '• Escalate to CRMT ONLY if customer requests manager/escalation or mentions CCTS. Simply refusing self-serve = NOT valid escalation reason.\n\n' +
    'CALLBACKS (flag violations):\n' +
    '• Confirm callback number at start of EVERY call. Update mobile number on profile.\n' +
    '• Scheduled callbacks: ONLY offer when customer specifically requests (never proactively).\n' +
    '• INVALID follow-up scenarios: sales/re-contracting, bill review after adjustments, remote resolve, callback for another dept (unless HS CxSS/TS specific exception).\n\n' +
    'AVOID REPEAT CALLS (flag if missing):\n' +
    '• MANDATORY end-of-call: "I really value your time — I want to make sure you have everything you need so you don\'t have to call back. Do you need help with any other TELUS services?"\n' +
    '• Check Repeat Indicator in Casa/Genesys before ending call.\n' +
    '• Complete ALL commitments made during the call (tickets, follow-ups).\n\n' +
    'CEB BEHAVIORS (flag if missing):\n' +
    '• PARAPHRASE & CONFIRM: Paraphrase request back ("Just to be clear, you are calling to..."), check Repeat Indicator, share relevant account context.\n' +
    '• DISCOVERY: Ask open-ended WHY question FIRST before presenting solution. Use benefit-led transition ("If I was able to...would you consider?").\n' +
    '• HOLD TECHNIQUES: Lead with benefit before holding, ask permission, get callback number, check in within 2 minutes, thank on return. Use HOLD not MUTE for processing.\n' +
    '• ENSURING UNDERSTANDING: Give permission to interrupt, ask permission to probe, suggest pen/paper before complex info, ask "What was clear and what needs more review?" after explaining.\n' +
    '• CAN DO SOLUTIONS: Thank customer → summarize root cause → state what I CAN do → confirm satisfaction before proceeding.\n' +
    '• CHECK FOR SATISFACTION (MANDATORY every call): "I really value your time... Is there anything else? Do you need help with other TELUS services?" End with name + personalized closing.\n' +
    '• GREAT RECAPPING: Paraphrase reason for calling → recap steps taken → reinforce extras done for customer → offer to write down.\n';
  }

  if (lob === 'PF CLS') {
    return 'LOB ROLE: PF CLS — PureFibre Retention/Loyalty\n\n' +
    'OFFER BUILDING SEQUENCE (flag violations):\n' +
    '• STEP 1: Make emotional connection FIRST (empathy, willingness to help) BEFORE any offer.\n' +
    '• STEP 2: Probe customer needs (What changed? What matters most? What are they paying now? Competitor offer details?).\n' +
    '• STEP 3: Check Casa Offers FIRST — always before manual CLS codes.\n' +
    '• STEP 4: Must always offer speed UPGRADE. Present 2 upgrade options: Priority 1 ($20+ lift), Priority 2 ($10+ lift), Priority 3 (same price). Never just one option.\n' +
    '• Lead with VALUE (5-year price lock, Simple Everyday Pricing) NOT discounts. Call it "loyalty rate" not "regular price."\n' +
    '• ATL (recommended tile) first; BTL only as back-pocket to close.\n' +
    '• Do NOT offer legacy/non-current plans — always upgrade to current PureFibre West 2026 plans.\n' +
    '• OTC: ONLY for genuine competitor offer match. NEVER to close a price gap or when customer just asks for a credit.\n\n' +
    'COMPETITOR OFFER HANDLING:\n' +
    '• Ask: total price, ongoing vs promo duration, regular price after promo, equipment/installation fees, additional charges.\n' +
    '• Lead with TELUS advantages (5-year price lock, symmetrical speeds, direct fibre), not by attacking competitors.\n\n' +
    'CEASE/CANCEL BEST PRACTICES (flag violations):\n' +
    '• Date ceases to END of current bill cycle. If customer wants sooner, process and advise partial charges.\n' +
    '• When applying new offer/renewal: REMOVE ALL existing discounts FIRST.\n' +
    '• Cease standalone equipment — not doing so prevents BAN closure (CCTS risk).\n' +
    '• Check TELUS Rewards before ceasing: advise customer they lose points; date cease so they can use outstanding points (do NOT advise to call back).\n' +
    '• Do NOT proactively offer manager callback. Never refuse manager request.\n' +
    '• Do NOT place orders before transferring to CLS2.\n\n' +
    'ETF/SATF RULES (flag if wrong amount or inappropriate waiver):\n' +
    '• Check SA start date: on/after April 11 2026 = $20/month; before = $15/month per service.\n' +
    '• Cannot stack Internet ETF + Boost Wi-Fi Easy Payment ETF.\n' +
    '• GWP cancellation fees: auto-waived for SA entered before July 28 2023 (CRTC). Do not manually charge.\n' +
    '• ETF waiver sensitive scenarios (NEVER proactively ask): Armed Forces transferred out of province, deceased customer, customer escaping domestic abuse.\n\n' +
    'PRE-INSTALL SAVE: Max $100 in Casa (not CSR). Tag: #Preinstallsave in comments. No stacking with CSR retention offers.\n\n' +
    'CCTS HANDLING (flag violations):\n' +
    '• If customer mentions CCTS/CRTC/BBB/Legal/Office of President: acknowledge → diffuse → warm transfer to CRMT immediately.\n' +
    '• If customer declines transfer: advise callback + connect with CRMT queue to inform them.\n\n' +
    'CEB BEHAVIORS (same as PF CxSS — all 7 apply, especially Overcoming Objections and Negotiations).\n';
  }

  if (lob === 'PF TS') {
    return 'LOB ROLE: PF TS — PureFibre Technical Support\n\n' +
    'SUPPORT SCOPE:\n' +
    'IN SCOPE: Internet/TV troubleshooting, modem issues, wireless connectivity/speed, Optik TV, remote troubleshooting/replacement, port profile changes, outage support, appointment management.\n' +
    'OUT OF SCOPE (flag if handled incorrectly): Port forwarding/bridge mode, order-related provisioning, Smart Hub, customer device issues (laptops, smart TVs — refer to manufacturer), domain/server support, inside wiring appointments.\n\n' +
    'BEFORE TRANSFERRING TS → CARE (flag if skipped):\n' +
    '• Complete all available TS troubleshooting first.\n' +
    '• Check Dispatch Status Tool for install date before transferring.\n' +
    '• Transfer to Care for: account changes, billing inquiries, equipment adds/deletions.\n\n' +
    'ORDER STATUS OWNERSHIP:\n' +
    '• Initial/Negotiation/Cancelled → always CARE.\n' +
    '• Delivery (RW/SW/FW completed) → TS owns.\n' +
    '• Completion/Done → TS owns.\n' +
    '• Same-day order script: "Your order is dated for today. If services not working by tomorrow morning, please call 310-2255 and select Repair."\n\n' +
    'CEB BEHAVIORS (apply all 7 CEB items).\n';
  }

  if (lob === 'SHS CXSS') {
    return 'LOB ROLE: SHS CxSS — SmartHome Security Frontline\n\n' +
    'PRE-INSTALL REQUIREMENTS (flag if not covered):\n' +
    '• Minimum 2 bars cellular coverage required for professionally monitored plans — technician cannot install without it.\n' +
    '• Authorized user 18+ must be present for full installation — failure = incomplete install.\n' +
    '• Emergency contacts: 1 site number + 2 alternatives (all unique, Canadian, not toll-free, with person name).\n' +
    '• 48-hour test mode after install: "Services monitored; however 2-way voice not active and no emergency services dispatched via CMS."\n' +
    '• Ask about pets (calibrate sensor sensitivity), firearms (inform CMS), age of home (asbestos).\n' +
    '• Quebec: Smart Thermostat CANNOT be installed by TELUS technicians.\n' +
    '• MDU: Outdoor cameras and doorbell cameras CANNOT be installed.\n\n' +
    'EQUIPMENT RETURN (flag if wrong advice):\n' +
    '• NEVER tell customer to throw away or recycle TELUS equipment.\n' +
    '• Cease: return Qolsys IQ panel, keypads, cameras, thermostat, garage opener. Leave: door locks, smoke/CO/motion/flood sensors, 4-button remote, legacy panels.\n' +
    '• 45-day return window, free Canada Post shipping.\n' +
    '• Both CxSS and CLS can credit unreturned equipment — no need to transfer to CLS for this.\n' +
    '• Can credit even without waybill if customer confirms equipment returned.\n' +
    '• SHS wall-mounted panel → issue credit IMMEDIATELY (removing damages wall). Table-mounted → customer returns first, THEN credit.\n\n' +
    'CANCEL/CEASE ROUTING: ALL SHS cancellations → transfer to SHS Retention (no exceptions).\n\n' +
    'DECEASED CUSTOMER (flag violations):\n' +
    '• Death certificate NOT required (removed June 2026).\n' +
    '• Service MUST be cancelled or transferred within 30 days — beyond 30 days = legally considered fraud.\n' +
    '• Mandatory: change billing to "Estate of", set up paper billing, send CAM notification via Casa Notify (Level 2 - Related disputes), remove from marketing lists.\n\n' +
    'CEB BEHAVIORS (apply all 7 CEB items).\n';
  }

  if (lob === 'SHS CLS') {
    return 'LOB ROLE: SHS CLS — SmartHome Security Retention\n\n' +
    'ROUTING RULE: ALL SHS cancellation requests → SHS Retention handles (no exceptions).\n\n' +
    'SAVE PROCESS (flag if skipped):\n' +
    '• Attempt save offer FIRST before processing cease. Review cancellation fees with customer.\n' +
    '• 30-day satisfaction guarantee cancels: attempt save → waive all fees → submit Google form → advise charges on closing statement waived within 1 week.\n' +
    '• SHS downgrade to Smart Automation Plus (LAST RESORT — CLS only): try renewal offers first (SHSRENEW5/10/15/20). Only use SAP15MTM if all else fails.\n' +
    '  - Must read mandatory legal script before proceeding (removes 24/7 monitoring — customer responsible for calling 911).\n' +
    '  - Get verbal consent from customer. Remove ALL existing discounts first. Send go/send confirmation email after order.\n' +
    '  - Advise customer to contact home insurance and check alarm permit implications.\n\n' +
    'SHS ETF (flag if wrong):\n' +
    '• Check SA start date: after Nov 13 2019: Secure/Control = $15/month, Smart Auto Plus = $10/month, Smart Camera = $5/month.\n' +
    '• Before Jan 26 2019: Secure = $24, Protect = $32, Control = $39/month remaining.\n' +
    '• If CSA-quoted ETF is LOWER than guidelines → honor the CSA fee (do not adjust upward).\n' +
    '• Quebec ILEC/NILEC: Maximum $50 cancellation.\n' +
    '• ETF exceptions (NEVER proactively ask): deceased, Armed Forces transferred out of province, customer escaping domestic abuse.\n\n' +
    'RENEWALS (flag violations):\n' +
    '• Never say "contract" — always "Service Agreement."\n' +
    '• Do NOT select email for CSA notification — advise customer to access via MyTELUS.\n' +
    '• CxSS cannot renew Migrated customers — transfer to SHS CLS.\n' +
    '• Migrated renewal: do NOT add any equipment (forces to in-market plan).\n\n' +
    'CEASE CHECKLIST (flag if missing):\n' +
    '• Advise customer to download/save videos BEFORE ceasing (cannot retrieve after cancellation).\n' +
    '• Return equipment: Qolsys panel, keypads, cameras, thermostat, garage opener. Leave: sensors, door locks, legacy panels.\n\n' +
    'CEB BEHAVIORS (apply all 7 — especially Overcoming Objections and Offering Solutions).\n';
  }

  if (lob === 'SHS TS') {
    return 'LOB ROLE: SHS TS — SmartHome Security Technical Support\n\n' +
    'CAMERA ROUTING (flag if wrong):\n' +
    '• Each camera needs 2.5 Mbps upload. Up to 4 cameras → SHS Team (BAU). 5+ cameras → Custom Home Team (go/customleads).\n' +
    '• alarm.com auto-configures to 4 cameras by default — must manually adjust for 5+ or extra cameras will NOT pair.\n' +
    '• Video Expansion Pack (8,000 clips): offer REACTIVELY ONLY (customer calls about exceeding clips or is churn risk). 24/7 Recording is first offer. Max 8,000 clips cap.\n\n' +
    'DIRECT FULFILL / SHIPPING (flag violations):\n' +
    '• Must manually select Shipping in CSR (defaults to Installer Supplied).\n' +
    '• Sensors CANNOT be shipped (Phase 1) — technician needed.\n' +
    '• Equipment damage/missing/stolen: ALWAYS waive equipment AND installation charges (regardless of who is liable).\n\n' +
    'BATTERIES (flag violations):\n' +
    '• Do NOT charge customer for batteries. Free shipping for most devices.\n' +
    '• Qolsys Panel battery (IQ Battery): technician dispatch ONLY (complex), waive truck roll fee.\n' +
    '• Low battery alert = approx. 2 weeks before dies.\n\n' +
    'HARDWARE SWAPS (flag violations):\n' +
    '• Complete all troubleshooting FIRST before offering swap.\n' +
    '• Warranty: TELUS branded equipment = 1 year; Hardware Store devices = 30 days; Non-TELUS Hardware Store = NOT eligible.\n\n' +
    'CEB BEHAVIORS (apply all 7 CEB items).\n';
  }

  if (lob === 'MOB CXSS' || lob === 'MOB CLS' || lob === 'MOB TS' ||
      lob === 'KOODO' || lob === 'KOODO TS') {
    return 'LOB ROLE: ' + selectedLOB + ' — Wireless\nNote: Wireless-specific evaluation guidelines will be added in a future update. Apply standard CEB behavior evaluation and general TELUS process principles.\n';
  }

  return '';
}

// ── Main orchestrator ─────────────────────────────────────────────────────────
function analyzeTranscript(transcriptText, analysisType, agentName, selectedLOB) {
  var knowledgeText = '';
  try {
    knowledgeText = analysisType === 'sales'
      ? getPDFKnowledgeFromFolder(SALES_FOLDER_ID,   PDF_CACHE_KEY_SALES)
      : getPDFKnowledgeFromFolder(REPEATS_FOLDER_ID, PDF_CACHE_KEY_REPEATS);
  } catch(e) { Logger.log('PDF fetch failed (non-fatal): ' + e); }

  var prompt = analysisType === 'sales'
    ? buildSalesPrompt(transcriptText, knowledgeText, agentName, selectedLOB)
    : buildRepeatsPrompt(transcriptText, knowledgeText, agentName, selectedLOB);

  return callFuelIX(prompt);
}

