# NH FCR, Transfer or Sales Call Analyzer — Complete Documentation

> **Version:** Current (October 2026)  
> **Platform:** Google Apps Script Web App  
> **Owner:** Lobell Bantaya — TELUS Quality Analyst Team

---

## Table of Contents

1. [Project Overview](#1-project-overview)
2. [Architecture](#2-architecture)
3. [Setup & Deployment](#3-setup--deployment)
4. [Configuration — `Config.gs`](#4-configuration--configgs)
5. [Sheet Service — `SheetService.gs`](#5-sheet-service--sheetservicegs)
6. [AI Service — `AIService.gs`](#6-ai-service--aiservicegs)
7. [Main Application — `Code.gs`](#7-main-application--codegs)
8. [Data Schemas](#8-data-schemas)
9. [Caching Strategy](#9-caching-strategy)
10. [Dashboard — `index.html`](#10-dashboard--indexhtml)
11. [Error Handling](#11-error-handling)
12. [Access Control](#12-access-control)
13. [HTML Sanitiser — `HtmlSanitizer.gs`](#13-html-sanitiser--htmlsanitizergs)
14. [Audit Check — `AuditCheck.gs`](#14-audit-check--auditcheckgs)

---

## 1. Project Overview

A Google Apps Script web application that enables TELUS contact center Quality Analysts to:

- **Submit call transcripts** for AI-powered evaluation (Repeats & Transfer Analyzer or Sales Analyzer)
- **Auto-fill agent details** from roster files using SAP ID lookup
- **Generate structured audit reports** with coaching recommendations, positioning statements, and RCA categorization
- **Track evaluations** in a Google Sheets database with Audit Reference numbers
- **Send notifications** to Admin/Dev users and email audit reports to Team Leaders
- **Visualize trends** on an interactive dashboard with charts, filters, and AI analytics briefings

### Key Features

| Feature | Description |
|---|---|
| SAP ID Lookup | Auto-fills Team Member, Team Leader, Ops Manager, LOB, Locale, VTID from roster |
| AI Analysis | FuelIX API (OpenAI-compatible) generates structured HTML evaluation reports |
| Audit Reference | Auto-generates `NHA-YYYYMMDD-XXXX` unique IDs per evaluation |
| Caching | Multi-layer caching (CacheService + localStorage) for sub-second loads |
| Dashboard | Interactive filters, bar charts, KPI cards, AI analytics briefings |
| Admin Notifications | Auto-emails all `Admin/Dev` roster members on each submission |
| PDF Generation | Structured evaluation form PDF attached to emails |

---

## 2. Architecture

```
┌─────────────────────────────────────────────────────────┐
│                     index.html                          │
│   (Single-page web app — HTML + CSS + JavaScript)       │
│   - Form: SAP ID, Team Member, Call Details, Transcript │
│   - Dashboard: Filters, Charts, Tables, AI Analytics    │
│   - Chat Assistant widget                               │
└────────────────────┬────────────────────────────────────┘
                     │ google.script.run (async RPC)
┌────────────────────▼────────────────────────────────────┐
│                    Code.gs                              │
│   - doGet() — web app entry point                       │
│   - submitTranscript() — main submission handler        │
│   - sendAuditEmail() — email with PDF attachment        │
│   - notifyAdmins() — Admin/Dev notifications            │
│   - getDashboardData() — cached dashboard rows          │
│   - getAIAnalytics() — executive briefing generator     │
│   - generateAuditPDF() / buildEvalFormHTML()            │
│   - enrichDashboardData() — RCA field extraction        │
└──────┬──────────────┬───────────────┬───────────────────┘
       │              │               │
┌──────▼──────┐ ┌─────▼──────┐ ┌────▼──────────────────┐
│  Config.gs  │ │SheetService│ │     AIService.gs       │
│  Constants  │ │    .gs     │ │  - callFuelIX_()        │
│  Headers    │ │ Sheet R/W  │ │  - lookupBySapId_()     │
│  Column idx │ │ Caching    │ │  - getAllRosterData_()  │
└─────────────┘ └────────────┘ │  - lookupVTID_()        │
                               │  - buildRepeatsPrompt()│
                               │  - buildSalesPrompt()  │
                               │  - analyzeTranscript() │
                               └────────────────────────┘
```

### External Dependencies

| Service | Purpose | ID/Key |
|---|---|---|
| FuelIX API | AI analysis & chat | `ANTHROPIC_AUTH_TOKEN` Script Property |
| Roster Google Sheet | Agent lookup | `ROSTER_SHEET_ID` |
| AT Data GCP Sheet | VTID lookup | `AT_DATA_GCP_SS_ID` |
| FCR Dashboard Sheet | Team Leader/Ops lookup | `FCR_DASHBOARD_SS_ID` |
| Repeats RAG Drive | PDF knowledge base | `REPEATS_FOLDER_ID` |
| Sales RAG Drive | PDF knowledge base | `SALES_FOLDER_ID` |
| Main Spreadsheet | All data storage | Created automatically |

---

## 3. Setup & Deployment

### Prerequisites

1. Google account with access to the TELUS organization
2. Apps Script project linked to a Google Drive folder
3. FuelIX API key
4. Access to the roster, AT Data GCP, and FCR Dashboard spreadsheets

### Step 1 — Script Properties

In Apps Script editor → Project Settings → Script Properties, add:

| Property | Value |
|---|---|
| `ANTHROPIC_AUTH_TOKEN` | Your FuelIX API key |
| `ANTHROPIC_BASE_URL` | `https://api.fuelix.ai` |
| `ANTHROPIC_MODEL` | `claude-sonnet-4` |
| `REGISTRY_SECRET` | Shared secret the Action Registry poster sends to the web app (`doPost`). **Set this.** Until it exists, the old secret that is committed in `Code.gs` (`REGISTRY_SECRET_LEGACY`) is still accepted, and it must be treated as public. Once it is set the old one stops working and can be deleted from the code. |

### Step 2 — Enable Advanced Services

Apps Script editor → Extensions → Services → Enable:
- **Drive API v2**

### Step 3 — OAuth Scopes

The `appsscript.json` declares these scopes (already configured):
- `https://www.googleapis.com/auth/drive`
- `https://www.googleapis.com/auth/spreadsheets`
- `https://www.googleapis.com/auth/documents`
- `https://www.googleapis.com/auth/script.external_request`
- `https://mail.google.com/`
- `https://www.googleapis.com/auth/gmail.send`
- `https://www.googleapis.com/auth/script.send_mail`

### Step 4 — Deploy

Deploy → New deployment → Web App:
- **Execute as:** Me
- **Who has access:** Anyone within your org (DOMAIN)

### Step 5 — One-Time Functions

Run these from the editor after first deployment. They are **admin-only**: each one starts with `requireAdmin_()`, so the account running it must be on `ADMIN_USERNAMES` (see [Access Control](#12-access-control)).

```javascript
setupScriptProperties()    // Removed — set manually
repairSheetHeaders()       // Fix column order in all sheets
addDashboardColumnNotes()  // Add hover tooltips to spreadsheet columns
```

---

## 4. Configuration — `Config.gs`

All global constants used across the application.

### `FUELIX_CONFIG`

```javascript
var FUELIX_CONFIG = { baseUrl, model, apiKey }
```

Reads FuelIX API credentials from Script Properties at startup. Falls back to defaults if properties are missing.

**Properties:**

| Key | Type | Default | Description |
|---|---|---|---|
| `baseUrl` | `string` | `https://api.fuelix.ai` | FuelIX API base URL |
| `model` | `string` | `claude-sonnet-4` | Model identifier |
| `apiKey` | `string` | `''` | Bearer token (must be set in Script Properties) |

---

### Sheet Name Constants

| Constant | Value | Purpose |
|---|---|---|
| `SPREADSHEET_NAME` | `'NH FCR, Transfer or Sales Call Analyzer'` | Main data spreadsheet name |
| `TRANSCRIPTS_SHEET` | `'Transcripts'` | Raw transcript submissions |
| `ANALYSIS_SHEET` | `'Analysis'` | AI analysis plain text results |
| `CACHE_SHEET` | `'Cache'` | Per-interaction HTML cache |
| `DASHBOARD_DATA_SHEET` | `'Dashboard_Data'` | Structured analytics data |
| `AUDIT_LOG_SHEET` | `'Audit_Log'` | Permanent audit archive |
| `AI_ANALYTICS_SHEET` | `'AI_Analytics'` | AI briefing cache |

---

### External Resource IDs

| Constant | Description |
|---|---|
| `ROSTER_SHEET_ID` | Google Sheet ID for the agent roster |
| `AT_DATA_GCP_SS_ID` | Google Sheet ID for AT Data GCP (VTID source) |
| `FCR_DASHBOARD_SS_ID` | Google Sheet ID for FCR Dashboard Data |
| `REPEATS_FOLDER_ID` | Google Drive folder ID for Repeats RAG PDFs |
| `SALES_FOLDER_ID` | Google Drive folder ID for Sales RAG PDFs |

---

### Roster Column Indices (0-based)

| Constant | Column | Header |
|---|---|---|
| `ROSTER_COL_SAP_ID` | A (0) | `sap_ID` |
| `ROSTER_COL_AGENT_NAME` | C (2) | `Agent_Name` |
| `ROSTER_COL_DEPT_CD` | D (3) | `Dept_Cd` |
| `ROSTER_COL_DOMAIN_NAME` | F (5) | `Agent Role (Scorecard)` → Line of Business |
| `ROSTER_COL_TEAM_MGR` | I (8) | `Team_Mgr_Name` → Team Leader |
| `ROSTER_COL_OPS_MGR` | K (10) | `Ops_Mgr_Name` → Operations Manager |
| `ROSTER_COL_LOCALE` | S (18) | `Location` → Locale (Site) |

---

## 5. Sheet Service — `SheetService.gs`

Handles all Google Sheets read/write operations.

---

### `getOrCreateSpreadsheet()`

**Description:** Finds the main spreadsheet by name or creates it if it doesn't exist. Opens it via `openSheetWithRetry()` rather than a bare `SpreadsheetApp.openById()`.

**Parameters:** None

**Returns:** `SpreadsheetApp.Spreadsheet`

**Example:**
```javascript
var ss = getOrCreateSpreadsheet();
```

---

### `openSheetWithRetry(id, label)`

**Description:** Opens a Spreadsheet by ID, retrying up to 3 times (with 750ms/1500ms backoff) if the error looks permission-related (`"permission"` or `"do not have access"` in the message). `SpreadsheetApp.openById()` can throw that exact error under heavy concurrent access even for a file's own owner — this deployment runs as "Execute as: Me", so a genuine per-user sharing gap isn't the likely cause of such reports. Retrying clears the transient case; a truly broken share still surfaces the same error after these attempts, just slightly delayed. Used by every lookup source the analyzer depends on: the main spreadsheet, Roster, AT Data GCP / Trainee Roster, Global Roster, and Users / Trainer Lookup.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `id` | `string` | Yes | Spreadsheet ID to open |
| `label` | `string` | Yes | Human-readable name used in the thrown error if every attempt fails |

**Returns:** `SpreadsheetApp.Spreadsheet`

**Throws:** If every attempt fails, an `Error` naming both `label` and `id`.

---

### `getOrCreateSheet(spreadsheet, sheetName)`

**Description:** Returns a named sheet tab, creating it if it doesn't exist.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `spreadsheet` | `Spreadsheet` | Yes | Parent spreadsheet object |
| `sheetName` | `string` | Yes | Tab name to find or create |

**Returns:** `SpreadsheetApp.Sheet`

**Example:**
```javascript
var sheet = getOrCreateSheet(ss, 'Dashboard_Data');
```

---

### `getTeamDetailsFromAuditHistory_(sapId)`

**Description:** Backfill source for `submitTranscript()` when the client-side roster autofill left Team Leader / Ops Manager / LOB / Locale / VTID blank (e.g. an external Roster/AT-Data lookup failure). Reads `Audit_Log` — which lives in the analyzer's own spreadsheet, not an external file — and returns the values from this SAP ID's most recent prior audit, scanning newest-first. Never throws; returns `null` on any error or if the SAP ID has no prior history. Data may be stale if the team member was reassigned since their last audit — callers should warn when a returned value is actually used, not apply it silently.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `sapId` | `string` | Yes | SAP ID to look up in Audit_Log |

**Returns:** `Object|null` — `{ teamLeader, opsManager, lineOfBusiness, locale, vtid }` from the most recent matching row, or `null` if none found

---

### `ensureHeaders(sheet, headers)`

**Description:** Writes the header row with purple styling and frozen row if the sheet is empty.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `sheet` | `Sheet` | Yes | Target sheet |
| `headers` | `string[]` | Yes | Array of column header names |

**Returns:** `void`

**Example:**
```javascript
ensureHeaders(sheet, TRANSCRIPTS_HEADERS);
```

---

### `appendRow(sheet, rowData)`

**Description:** Appends a single row of data to a sheet.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `sheet` | `Sheet` | Yes | Target sheet |
| `rowData` | `any[]` | Yes | Array of values matching column order |

**Returns:** `void`

---

### `getSpreadsheetUrl()`

**Description:** Returns the URL of the main spreadsheet.

**Returns:** `string` — full Google Sheets URL

---

### `logSubmissionError(functionName, formData, e)`

**Description:** Writes a `submitTranscript()` failure to the `Error_Log` tab in the main spreadsheet. Reading Stackdriver/GCP execution logs requires a non-default GCP project linked to the script, which this project does not have, so `Logger.log()` output alone is not actually reachable by the admin — this gives instant visibility with no extra setup. Never throws; a logging failure must never mask the original error being returned to the analyst.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `functionName` | `string` | Yes | Name of the function that failed, e.g. `'submitTranscript'` |
| `formData` | `Object` | No | The submission's form data, for context (observer, SAP ID, interaction ID) |
| `e` | `Error` | No | The caught exception |

**Returns:** `void`

---

### `findCachedResult_(interactionId, analysisType, sapId)`

**Description:** Searches the Cache sheet for a previously processed audit using `TextFinder` (O(1) lookup, not linear scan). A cached evaluation belongs to ONE agent: several agents can share one Interaction ID (a transferred call) and each is audited separately, so the match is Interaction ID + analysis type + SAP ID (Cache sheet column E). Rows written before column E existed have a blank owner. A blank-owner row belongs to the agent whose audit created it, the earliest Audit_Log row for that Interaction ID and analysis type, and is returned to that agent only. Each matching Cache row is read once and the continuation chunks are assembled from memory.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `interactionId` | `string` | Yes | The interaction/conversation ID (shared by every agent on the call) |
| `analysisType` | `string` | Yes | `'repeats'` or `'sales'` |
| `sapId` | `string` | Yes | SAP ID of the agent being audited; a blank value matches any agent |

**Returns:** `string|null` — cached HTML string, or `null` on cache miss

**Example:**
```javascript
var cached = findCachedResult_('230dcc38-4d04...', 'repeats');
if (cached) return { success: true, html: cached, fromCache: true };
```

---

### `saveCachedResult_(interactionId, analysisType, htmlResult, sapId)`

**Description:** Writes a new cache entry to the Cache sheet, tagged with the agent's SAP ID in column E. Replaces only that agent's existing rows for the Interaction ID and analysis type (and blank-owner rows only when that agent owns them), so saving or editing one agent's evaluation never overwrites another agent's on the same call. It reads only columns A, B and E of the Cache sheet, never the HTML column, and removes the old CacheService copy before writing the new one.

**Private** (trailing underscore): only server code calls it, so Apps Script does not expose it to the browser. The HTML is **sanitised first** (`sanitizeEvalHtml_`), so no caller can plant markup that runs; if nothing is left after sanitising, the existing cache is kept. A report over 48,000 characters is stored as consecutive chunk rows (`type`, `type_2`, `type_3` …); **every chunk is written with Sheets' own text marker (a leading apostrophe)**, because Sheets reads what `appendRow` writes as if typed, and a chunk that started with `=`, `+`, `-` or `@` became a formula and one that started with an apostrophe lost it. Reading the old rows, deleting them by index and appending the new ones runs **under the script lock**, so two saves at once cannot delete the wrong rows; if the lock cannot be had the evaluation is saved anyway. The same applies to `updateCachedResult_`, which calls this function.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `interactionId` | `string` | Yes | Interaction ID the entry belongs to |
| `analysisType` | `string` | Yes | `'repeats'` or `'sales'` |
| `htmlResult` | `string` | Yes | Full AI-generated HTML |
| `sapId` | `string` | Yes | SAP ID of the agent the evaluation is for |

**Returns:** `void`

---

## 6. AI Service — `AIService.gs`

Handles FuelIX API calls, data lookups, and prompt construction.

> **Private functions.** `callFuelIX_`, `getAllRosterData_`, `lookupBySapId_`, `lookupSapId_` and `lookupVTID_` end with an underscore, so Apps Script does **not** expose them to `google.script.run`. Before, any signed-in user could call `callFuelIX` as a free proxy to the AI key, or dump the roster with `getAllRosterData`. The browser reaches the lookups through `getRosterBySapId`, `getSapIdForParticipant` and `getTranscriptMetadataAndRoster` in `Code.gs`.
>
> **Why a lookup can come back empty.** The roster loaders record a source they could not read (`_noteSourceError_`, `_lookupSourceErrors`). `getRosterBySapId` then returns `{ _error: 'source_unavailable', sources: [...] }` instead of `null`, and the page says which source could not be read. A genuine miss is still `null`.

---

### `_getRosterSheetData()`

**Description:** Loads the full roster sheet into memory, cached 2 hours in CacheService and within the current execution via `_rosterDataInMemory`. Eliminates repeated sheet reads.

**Parameters:** None

**Returns:** `any[][]` — 2D array of roster rows (excluding header)

**Performance:** First call ~4s (sheet read), subsequent calls 0ms (in-memory)

---

### `_getATDataGCPSheetData()`

**Description:** Loads the full AT Data GCP sheet into memory, cached 2 hours. Fixes the 52-second timeout bug where every SAP lookup read 4,491 rows from scratch.

**Parameters:** None

**Returns:** `any[][]` — Full sheet data including header row at index 0

---

### `callFuelIX_(prompt)`

**Description:** Sends a prompt to the FuelIX API and returns the response content. Called completely unguarded from `submitTranscript()` (Code.gs), so every error it throws is tagged `AI_SERVICE_ERROR` — a FuelIX auth/outage error's response body often contains wording like "permission" or "access", which used to be misclassified by `submitTranscript()`'s generic error handling as a Sheets/Drive lookup failure. Retries up to 3x (1s, then 2s backoff) for transient-looking failures — 429 (rate limit), 5xx (server-side), or a network-level exception (timeout/DNS) — but not for a bad/expired API key (401/403) or a malformed request (400), since those fail identically every attempt.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `prompt` | `string` | Yes | Full prompt string |

**Returns:** `string` — AI response text (trimmed)

**Errors:** Throws `Error` tagged `AI_SERVICE_ERROR[<code>]` if every attempt's HTTP status is non-retryable or non-200, or the response structure is unexpected; `AI_SERVICE_ERROR[network]` if `UrlFetchApp` itself fails on every attempt.

**Example:**
```javascript
var response = callFuelIX_('Analyze this transcript: ...');
```

---

### `getPDFKnowledgeFromFolder(folderId, cacheKey)`

**Description:** Reads PDFs from a Drive folder, converts each to a Google Doc, extracts text, concatenates, and caches for 6 hours (up to 100KB).

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `folderId` | `string` | Yes | Google Drive folder ID |
| `cacheKey` | `string` | Yes | CacheService key for this folder |

**Returns:** `string` — concatenated plain text from all PDFs

---

### `lookupSapId_(participantName)`

**Description:** Reverse lookup — finds a SAP ID by agent name using the cached roster.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `participantName` | `string` | Yes | Agent's full name |

**Returns:** `string` — SAP ID or empty string if not found

---

### `getAllRosterData_()`

**Description:** Returns a structured array of all roster agents, cached 2 hours. Used to populate dashboard filter dropdowns.

**Parameters:** None

**Returns:** `Object[]` — array of `{ sapId, participant, lineOfBusiness, teamLeader, opsManager, locale }`

---

### `lookupVTID_(sapId)`

**Description:** Looks up the VTID for a SAP ID from the AT Data GCP sheet. Uses `_getATDataGCPSheetData()` for fast cached access. Results cached per-SAP-ID for 6 hours.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `sapId` | `string` | Yes | Agent SAP ID (Production ID in AT Data GCP) |

**Returns:** `string` — VTID (e.g. `VTI2026000599`) or empty string

---

### `lookupBySapId_(sapId)`

**Description:** Three-source lookup chain for agent details. Uses cached data — no live sheet reads after first call.

**Lookup order:**
1. **Roster** (5-digit SAP IDs) — returns LOB, Team Leader, Ops Manager, Locale
2. **AT Data GCP** (7-digit SAP IDs) — returns Agent Name, LOB, Locale, VTID, Facilitator as Team Leader
3. **FCR Dashboard** — enriches Team Leader, Ops Manager if missing

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `sapId` | `string` | Yes | Agent SAP ID |

**Returns:** `Object|null` — `{ participant, lineOfBusiness, teamLeader, opsManager, locale, vtid }` or `null`

**Example:**
```javascript
var row = lookupBySapId_('2007888');
// Returns: { participant: 'James Collado', lineOfBusiness: 'FFH', ... }
```

---

### `lookupFromFCRDashboard(targetStr, targetNum)`

**Description:** Searches all sheets in the FCR Dashboard spreadsheet for a SAP ID match, returning Team Leader, Ops Manager, LOB, and Locale. Results cached 6 hours.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `targetStr` | `string` | Yes | SAP ID as string |
| `targetNum` | `number` | Yes | SAP ID as number |

**Returns:** `Object|null` — `{ teamLeader, opsManager, lineOfBusiness, locale }` or `null`

---

### `getAllRosterData_()`

**Description:** Returns all roster agents as structured objects, used for dashboard filter pills.

**Returns:** `Object[]`

---

### `parseTranscriptMetadata(transcriptText)`

**Description:** Extracts metadata fields from transcript header text using regex patterns.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `transcriptText` | `string` | Yes | Raw transcript content |

**Returns:** `Object` — `{ interactionId, startTime, duration, direction, participant, customerBAN?, customerName? }`

**Extracted patterns:**
- `Interaction ID: <value>`
- `Transcript Start Time: <value>`
- `Transcript Duration: <value>`
- `Direction: <value>`
- `Internal Participant(s): <value>`
- BAN/account/phone numbers
- Customer names from External participant lines

---

### `buildRepeatsPrompt(transcriptText, knowledgeText)`

**Description:** Constructs the full Repeats & Transfer Analyzer prompt instructing the AI to return structured HTML with call summary, analysis tables, critical flags, and positioning statements.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `transcriptText` | `string` | Yes | Call transcript(s) |
| `knowledgeText` | `string` | No | Extracted PDF policy text |

**Returns:** `string` — Full prompt string

---

### `buildSalesPrompt(transcriptText, knowledgeText)`

**Description:** Constructs the Sales Analyzer prompt with 5-column framework tables (Skill Name, Score, Analysis & Strengths, Areas of Opportunity, Recommendation & Sample Positioning Statement).

**Parameters:** Same as `buildRepeatsPrompt`

**Returns:** `string` — Full prompt string

---

### `analyzeTranscript(transcriptText, analysisType)`

**Description:** Orchestrates the full AI analysis: loads PDF knowledge, builds the appropriate prompt, calls FuelIX, returns HTML.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `transcriptText` | `string` | Yes | Full transcript text |
| `analysisType` | `string` | Yes | `'repeats'` or `'sales'` |

**Returns:** `string` — AI-generated HTML response

---

### `sharedCSS()`

**Description:** Returns the full CSS string for the AI analysis HTML output (shared by modal view, PDF, and email).

**Returns:** `string` — `<style>` block HTML

---

### `lookupVTID_(sapId)` (also in AT Data)

See above under AT Data GCP section.

---

## 7. Main Application — `Code.gs`

The core server-side logic exposed to the web app via `google.script.run`.

---

### `doGet(e)`

**Description:** Web app entry point. Routes to dashboard or main analyzer based on `?page=` parameter.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `e` | `Event` | No | Request event with `.parameter` |

**Returns:** `HtmlOutput` — serves `index.html` or `Dashboard.html`

**Routing:**
- `?page=dashboard` → serves `Dashboard.html`
- Default → serves `index.html` (main analyzer)

---

### `generateAuditRef()`

**Description:** Generates a unique audit reference number in format `<PREFIX>-YYYYMMDD-XXXX`. The prefix comes from `getAuditRefPrefix_()` (Config.gs) and depends on which script is running: `NHA` for the NH Analyzer, `TIA` for the TI Analyzer, `NHA` for any unlisted script. The two projects share one spreadsheet but each keeps its own counter, so without separate prefixes they issued the same references. Uses `PropertiesService` (via `LockService.getScriptLock()` for atomicity) to persist today's sequence counter, avoiding a full Audit_Log scan. `PropertiesService` has no expiry, unlike `CacheService` — a per-day counter needs to survive up to 24 hours, which `CacheService.put()`'s hard 21600s (6-hour) TTL cap cannot hold without throwing. If the counter path throws (e.g. a lock-acquisition timeout under concurrent submissions), falls back to `_uniqueFallbackAuditRef()` rather than returning an unchecked value.

**Returns:** `string` — e.g. `NHA-20260804-0042` (NH) or `TIA-20260804-0042` (TI)

**Example:**
```javascript
var ref = generateAuditRef(); // 'NHA-20260804-0001'
```

---

### `_uniqueFallbackAuditRef(dateStr)`

**Description:** Last-resort ref generator used only when `generateAuditRef()`'s primary counter path throws. A bare random guess with no collision check previously caused the same-day audit-ref collision bug (two unrelated submissions landing on the identical `NHA-YYYYMMDD-####`, with `getEvalViewData()`/`getEvaluationByAuditRef()` silently returning the first matching row for both). This checks `Audit_Log` column A for the candidate ref before returning it, retrying up to 10 times, and only drops to a millisecond-timestamp suffix (`NHA-YYYYMMDD-T######`) if every guess collides or the sheet check itself fails.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `dateStr` | `string` | Yes | `YYYYMMDD` date prefix from the caller |

**Returns:** `string` — a ref guaranteed not to already exist in `Audit_Log`

---

### `resolveObserver()`

**Description:** Identifies the currently logged-in user by their Google account email, looks them up in the roster, and returns their name and SAP ID. Cached 8 hours in CacheService (user scope) and sessionStorage client-side.

**Returns:** `Object` — `{ name, sapId, email }`

**Fallback:** If email not found in roster, returns `{ name: email_prefix, sapId: '', email }`

---

### `readAuditLog()`

**Description:** Reads all Audit_Log rows, returning structured objects. Cached 10 minutes in CacheService. Cache is invalidated whenever a new row is appended.

**Returns:** `Object[]` — array of objects keyed by `AUDIT_LOG_HEADERS`

---

### `submitTranscript(formData)`

**Description:** Main form submission handler. The most critical function — orchestrates the complete evaluation workflow.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `formData` | `Object` | Yes | All form fields |
| `formData.sapId` | `string` | Yes | Agent SAP ID |
| `formData.transcript` | `string` | Yes | Full call transcript |
| `formData.participant` | `string` | No | Agent name (auto-filled from SAP lookup) |
| `formData.analysisType` | `string` | No | `'repeats'` or `'sales'` |
| `formData.interactionId` | `string` | No | Auto-extracted from transcript |
| `formData.teamLeader` | `string` | No | Team Leader name |
| `formData.opsManager` | `string` | No | Operations Manager name |
| `formData.lineOfBusiness` | `string` | No | Line of Business |
| `formData.locale` | `string` | No | Site/locale |
| `formData.vtid` | `string` | No | VTID from AT Data GCP |
| `formData.observerName` | `string` | No | QA observer full name |
| `formData.auditRef` | `string` | No | Pre-generated audit ref |

**Workflow:**
1. Validate required fields (`sapId`, `transcript`)
2. Parse transcript metadata server-side (fallback if client autofill missed) and resolve the Interaction ID (the form's, else the one read from the transcript)
3. Take the per-Interaction-ID "being processed" flag (see below); refuse with "already being processed" if another submit holds it
4. Check the Cache sheet:
   - cached **and** the audit has an Audit_Log row: return it (`fromCache: true`)
   - cached but **no** Audit_Log row (an earlier attempt failed after caching): **complete the records** from the cached report (steps 7 and 8), reusing the original reference from the Transcripts row, without calling the AI again
5. Generate the Audit Reference (or reuse the one from step 4)
6. Save the raw transcript row to Transcripts (skipped when completing an audit whose row exists)
7. `produceEvaluation_`: translate (if asked), narrow to the agent, run the AI, tidy the output, **sanitise it** (`sanitizeEvalHtml_`), and save it to the Cache sheet
8. `writeAuditRecords_`: write the Analysis, Dashboard_Data and Audit_Log rows. Each write is independent: one failing never stops the others, and adds a warning that names the reference. When completing, a row that already exists is not written twice.
9. Release the flag, in a `finally` block, on every way out (success, cache hit, access denied, AI failure, unexpected error). Only the flag this call took is released.

**The "being processed" flag.** One CacheService entry per Interaction ID **and** SAP ID (`_procKey_`), 10-minute TTL. It stops a double click from running the paid AI call twice. It is keyed on the **resolved** Interaction ID, so it also works when the form field is blank.

**Verifying an audit.** `verifyMyLatestAudit()` (see [Audit Check](#14-audit-check--auditcheckgs)) proves that one audit landed in every sheet and in the cache.

**Returns:** `Object` — `{ success, html, sheetUrl, fromCache, auditRef, observerName, warnings, filterStats, resolvedFields }`

---

### `sendAuditEmail(formData, htmlResult)`

**Description:** Generates a PDF and sends it to the Team Leader and Operations Manager. Before sending it checks that the form still describes the audit saved under `formData.auditRef`: if the form's SAP ID or Interaction ID differs from the saved Audit_Log row it sends nothing and returns an error telling the analyst to Clear All and re-process (a blank field on either side is not compared). Without this the subject and recipients could be about one agent while the evaluation link opened another agent's. It also refuses a reference that two Audit_Log rows share, because the link carries only the reference and would open the first row, and it updates the cache under the saved audit's analysis type, not the form's Repeats/Sales radio.

**Recipients and TEST mode.** `formData.emailMode` is `live` (default) or `test`. TEST sends only to the `Admin/Dev` rows of the Roster tab. The server decides with `isAdminCaller_()`, the same rule the screen uses to show the TEST toggle. A request for TEST from anyone who is not an admin (including a caller whose email is hidden, or an admin name on another domain) is **refused with a message**; it is never turned into a real send. In LIVE mode the agent's address is taken from the roster row for the **SAP ID**, then from the address the analyst was shown, then from the name, and only when exactly one person has that name. A name shared by people with different addresses resolves to nobody. If no address is found the error says whether the TEST or the LIVE list was empty.

`htmlResult` is sanitised (`sanitizeEvalHtml_`) before it is used, emailed or saved.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `formData` | `Object` | Yes | Form data (same as submitTranscript) |
| `htmlResult` | `string` | Yes | AI analysis HTML to attach |

**Returns:** `Object` — `{ success, recipients, auditRef }`

---

### `reportDuplicateAuditRefs()` and `applyDuplicateAuditRefRenames()` (`RefMaintenance.gs`)

**Description:** One-off admin tools (**admin-only**, `requireAdmin_()`), run from the Apps Script editor, for audit references that two audits already share (from the time the NH and TI scripts used the same prefix). Run `reportDuplicateAuditRefs()` first: it only reads and logs what would change. `applyDuplicateAuditRefRenames()` keeps the earliest row's reference and gives each later row a new one (the old reference plus `-D2`, `-D3`, ...) in Audit_Log, Dashboard_Data, Transcripts and Analysis, matching rows by reference and SAP ID (and Interaction ID where the sheet has that column). Rows that cannot be told apart are reported for manual review and left unchanged. Every change is written to the `Ref_Rename_Log` sheet, and running it again changes nothing. An evaluation link emailed for a renamed row stops working instead of showing another agent's evaluation.

**Returns:** `Object` — report: `{ duplicateRefs, renames, manual, plan }`; apply: `{ renamed, manual }`

---

### `notifyAdmins(formData, auditRef)`

**Description:** Reads the Roster tab for `Admin/Dev` role users and sends each a formatted HTML email with the evaluation summary. Sends individually so one bad address doesn't block others.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `formData` | `Object` | Yes | Submitted evaluation data |
| `auditRef` | `string` | Yes | Audit reference number |

**Returns:** `void` (non-fatal — errors are logged, not thrown)

---

### `getAdminRecipients()`

**Description:** Reads the Roster sheet (from the main spreadsheet) for rows where Column C = `Admin/Dev` (case-insensitive). Validates email format before including.

**Returns:** `Object[]` — `[{ email, name }, ...]`

---

### `generateAuditPDF(formData, htmlResult)`

**Description:** Converts the evaluation form HTML to a PDF blob using Drive's Google Doc conversion. Uses `try/finally` to guarantee the temporary Doc is always trashed.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `formData` | `Object` | Yes | Evaluation form data |
| `htmlResult` | `string` | Yes | AI analysis HTML |

**Returns:** `Blob|null` — PDF blob ready for email attachment, or `null` on error

---

### `buildEvalFormHTML(formData, auditRef, analysisResult)`

**Description:** Generates a structured, print-ready HTML evaluation form matching the FCR Live Call Listening Evaluation layout. Includes two-column detail grid, call summary, highlights/lowlights boxes, and full AI analysis.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `formData` | `Object` | Yes | All evaluation fields |
| `auditRef` | `string` | Yes | Audit reference number |
| `analysisResult` | `string` | Yes | AI HTML or plain text |

**Returns:** `string` — Complete `<!DOCTYPE html>` document string

---

### `getDashboardData()`

**Description:** Returns Dashboard_Data rows with server-side CacheService (30 min TTL). Merges stale cache with fresh fetch in background on client side.

**Returns:** `Object` — `{ rows, ts, rowCount, fromCache }`

---

### `getDashboardDataFresh()`

**Description:** Force-refreshes the dashboard cache, invalidates CacheService, reads sheet, re-caches. Called by ⟳ Refresh button.

**Returns:** `Object` — same as `getDashboardData()`

---

### `getAIAnalytics(rows, filters, analysisType)`

**Description:** Generates an AI executive intelligence briefing from filtered dashboard rows. Checks AI_Analytics sheet cache (24h TTL) before calling FuelIX.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `rows` | `Object[]` | Yes | Filtered dashboard data rows |
| `filters` | `Object` | Yes | Active filter state `{ dateFrom, dateTo, members, ... }` |
| `analysisType` | `string` | Yes | `'repeats'` or `'sales'` |

**Returns:** `Object` — `{ success, html, fromCache, generatedAt }`

---

### `sendAnalyticsBriefing(toEmail, analyticsHTML, filters, analysisType, rowCount)`

**Description:** Emails the AI analytics briefing HTML to a specified address.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `toEmail` | `string` | Yes | Recipient email address |
| `analyticsHTML` | `string` | Yes | Rendered briefing HTML |
| `filters` | `Object` | Yes | Active filters for subject line |
| `analysisType` | `string` | Yes | `'repeats'` or `'sales'` |
| `rowCount` | `number` | Yes | Number of records analyzed |

**Returns:** `Object` — `{ success }` or `{ success: false, error }`

---

### `enrichDashboardData()`

**Description:** One-time function (run from editor). Re-processes cached AI HTML for each Dashboard_Data row, extracting structured RCA fields: Call Driver, RCA Category, RCA Sub-Parameter, Top Opportunity, Product Opportunity, Critical Flags, Repeat Projection %.

**Returns:** `void` (logs enrichment count)

---

### `extractStructuredRCA(html, analysisType, auditRef)`

**Description:** Calls FuelIX AI on plain-text evaluation content to extract 7 structured JSON fields for the dashboard.

**Returns:** `Object|null` — `{ callDriver, rcaCategory, rcaSubParameter, topOpportunity, productOpportunity, callSummaryShort, criticalFlags, repeatPct }`

---

### `repairSheetHeaders()`

**Description:** One-time repair function. Reads each sheet's current headers, detects mismatches with the target definition in Config.gs, and rewrites the sheet with correct column order. Safe to run multiple times.

**Returns:** `void`

---

### `htmlToPlainText(html)`

**Description:** Converts AI-generated HTML to readable plain text for spreadsheet storage. Strips `<style>` and `<script>` blocks, converts block elements to newlines, decodes HTML entities.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `html` | `string` | Yes | Raw HTML string |

**Returns:** `string` — Clean plain text

---

### `fixBadgeClasses(html)`

**Description:** Post-processes AI HTML to correct score badge CSS classes based on actual numeric values. The AI is unreliable at assigning correct classes (e.g. it might put a score of 1.5 in `ai-badge-good`). This regex pass enforces: 3–5=green, 2–2.9=amber, 1–1.9=red, 0=grey.

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `html` | `string` | Yes | AI-generated HTML |

**Returns:** `string` — HTML with corrected badge classes

---

### `sanitiseHTMLForPDF(html)`

**Description:** Strips dangerous HTML tags from AI output before embedding in PDFs: `<script>`, `<iframe>`, `<object>`, `<form>`, inline event handlers.

**Returns:** `string` — Sanitised HTML

---

### `extractTextBlock(html, keyword)`

**Description:** Extracts plain text from a named section in AI HTML using 6 regex strategies (table cells, ai-title sections, chip values, broad text match).

**Parameters:**

| Name | Type | Required | Description |
|---|---|---|---|
| `html` | `string` | Yes | AI HTML |
| `keyword` | `string` | Yes | Section name to search for |

**Returns:** `string` — Extracted plain text (max 500 chars) or empty string

---

### `backfillVTIDFromATData()`

**Description:** One-time function. Reads unique SAP IDs from Dashboard_Data with blank VTID, looks each up in AT Data GCP, writes results back. Fixes historical rows.

**Returns:** `void`

---

### `fixRepeatProjectionPct()`

**Description:** One-time function. Extracts Repeat % numbers from the Issue Resolved text column and Audit_Log, writing clean integers to the Repeat Projection % column.

**Returns:** `void`

---

## 8. Data Schemas

### Transcripts Sheet

| Column | Header | Description |
|---|---|---|
| A | Audit Ref | `NHA-YYYYMMDD-XXXX` |
| B | Timestamp | Submission datetime |
| C | SAP ID | Agent SAP ID |
| D | Team Member | Agent full name |
| E | Team Leader | TL from roster |
| F | Operations Manager | Ops Mgr from roster |
| G | Line of Business | From roster col F |
| H | Locale | From roster col S |
| I | Observer Name | QA analyst name |
| J | Transcript Start Time | From transcript metadata |
| K | Customer BAN | BAN or customer name |
| L | Interaction ID | UUID from transcript |
| M | Direction | Inbound/Outbound/Chat/Email |
| N | Transcript Duration | e.g. `34:50` |
| O | Analysis Type | Repeats & Transfer Analyzer or Sales Analyzer |
| P | Transcript | Full raw transcript text |

### Analysis Sheet

| Column | Header | Description |
|---|---|---|
| A | Audit Ref | Links to Transcripts |
| B | Timestamp | Same as submission |
| C | SAP ID | Agent SAP ID |
| D | Team Member | Agent name |
| E | Observer Name | QA analyst |
| F | Analysis Type | Repeats or Sales |
| G | AI Analysis Result | Plain text (HTML-stripped) |

### Dashboard_Data Sheet (30 columns)

Key columns for analytics:

| Column | Header | Source |
|---|---|---|
| A | Audit Ref | Generated |
| D | SAP ID | Form |
| E | Team Member | Roster |
| J | VTID | AT Data GCP |
| O | Analysis Type | Form |
| P | Call Reason | Extracted from AI HTML |
| Q | Call Summary | Extracted from AI HTML |
| T | Critical Flags | AI-enriched |
| U | Repeat Projection % | AI-enriched (integer 0–100) |
| X | Call Driver | AI-enriched (short label) |
| Y | RCA Category | AI-enriched (one of 4 categories) |
| Z | RCA Sub-Parameter | AI-enriched (specific gap) |
| AA | Top Opportunity | AI-enriched (coaching action) |
| last | Auditor Email | The signed-in caller as the server reads it (`resolveAuditorEmail_`); if Google hides it, the Users sheet address(es) for the observer name (all of them, separated by `, `, when the name has several); blank when the name is not there |

### Audit_Log Sheet

Permanent archive written on every submission, never modified.

| Column | Header |
|---|---|
| A | Audit Ref |
| B | Submitted At |
| C | Interaction ID |
| D | SAP ID |
| E | Team Member |
| F | Team Leader |
| G | Operations Manager |
| H | Line of Business |
| I | Locale |
| J | VTID |
| K | Observer Name |
| L | Analysis Type |
| M | Direction |
| N | Duration |
| O | Repeat Projection % |
| P | Issue Resolved |
| Q | Transfer Occurred |
| R | Email Status |
| S | Recipients |

### Roster Tab (for Admin Notifications)

The Roster tab in the main spreadsheet (not the TELUS roster file) controls who receives Admin/Dev notifications:

| Column | Header | Values |
|---|---|---|
| A | Email Address | Full email |
| B | Name | Display name |
| C | Role | `Admin/Dev` triggers notifications |
| D | Supervisor | Supervisor name |
| E | Supervisor Email | Supervisor email |
| F | Status | Active/Inactive |
| G | ccList | CC recipients |

---

## 9. Caching Strategy

Multi-layer caching ensures fast load times across all components:

| Layer | Storage | TTL | Data |
|---|---|---|---|
| In-process variable | JavaScript memory | Current execution | Roster rows, AT Data GCP rows |
| `CacheService.getScriptCache()` | Apps Script shared | 30 min (dashboard), 2h (roster/AT Data), 6h (VTID/PDFs), 10 min (audit log) | All server-side data |
| `CacheService.getUserCache()` | Per-user | 8h | Observer identity |
| Google Sheet tab | `Cache` sheet | Permanent | Per-interaction AI HTML |
| Google Sheet tab | `AI_Analytics` sheet | 24h | AI briefing HTML per filter combo |
| `localStorage` | Browser | 15 min | Full dashboard rows |
| `sessionStorage` | Browser session | Session | Observer name |

### Cache Invalidation

| Event | Caches invalidated |
|---|---|
| New transcript submission | Dashboard CacheService, Audit_Log CacheService |
| Dashboard_Data write | Dashboard CacheService |
| Force Refresh button | Dashboard CacheService + localStorage |
| Audit_Log write | Audit_Log CacheService |

---

## 10. Dashboard — `index.html`

The single-page web app. Key JavaScript functions:

### Form Functions

| Function | Description |
|---|---|
| `debounceSapLookup()` | Fires 600ms after SAP ID input; calls `getRosterBySapId` |
| `scheduleAutofill()` | Fires 800ms after transcript paste; extracts metadata |
| `handleSubmit()` | Validates form, calls `submitTranscript`, opens analysis modal |
| `lockFormAfterSubmit(auditRef)` | Disables all fields after Submit & Email |
| `clearAllFields()` | Resets entire form, restores locked state |
| `saveAsDraft(btn)` | Persists form to localStorage |
| `restoreDraftIfAvailable()` | Prompts to restore on page load |
| `downloadAuditPDF()` | Calls `generateAndServePDF`, opens result in new tab |

### Dashboard Functions

| Function | Description |
|---|---|
| `openDashboard()` | Shows dashboard panel, loads data, shows Analytics FAB |
| `loadDashboardData()` | Checks localStorage (15 min), renders stale + fetches fresh |
| `dbHandleData(result, isLocalCache)` | Merges Audit_Log + Dashboard_Data, renders all tabs |
| `dbApplyAll()` | Applies shared filters to all 3 tabs simultaneously |
| `dbSwitchTab(tab, btn)` | Switches tab, updates stats row, updates FAB target |
| `dbPopulateFilters()` | Builds pill dropdown options from merged data |
| `dbSelectAllPills()` | Pre-selects all options on dashboard open |
| `dbGetSharedFilters()` | Reads current pill selections (all-selected = no filter) |
| `dbRenderStats(pfx, rows)` | Renders KPI stat cards with icons and sub-text |
| `dbRenderTable(pfx, rows)` | Renders critical flag details table with Audit Ref links |
| `dbDrawChart(pfx, rows)` | Draws Google Charts bar charts |
| `openAnalyticsPanel(pfx)` | Opens AI analytics briefing panel |
| `anRunAnalytics()` | Filters rows by panel filters, calls `getAIAnalytics` |
| `sendAnalyticsBriefingFromUI()` | Emails briefing from panel footer |

### Pill Filter Functions

| Function | Description |
|---|---|
| `dbTogglePill(key, pillEl)` | Opens/closes a filter dropdown |
| `dbPillMemberChange(name, checkbox)` | Cross-fills Leader, Ops, LOB, Locale when member selected |
| `dbPreTickPill(key, vals)` | Pre-ticks specific values in a pill store |
| `dbClearPill(key)` | Resets pill to "All" (re-selects everything) |
| `dbClearAll()` | Resets all pills and date inputs |
| `dbPillOnly(key, val)` | Deselects all others, selects just one ("only" button) |
| `dbPillSelectAll(key, selectAll)` | Selects or deselects all options in a pill |

### Chat Assistant Functions

| Function | Description |
|---|---|
| `toggleChat()` | Opens/closes the QA Assistant chat widget |
| `sendChatMessage()` | Sends user message to `chatAssistant` server function |
| `appendChatMsg(role, text)` | Adds a message bubble to the chat UI |

---

### Views, Home and the loading screen (new look, October 2026)

- **One bar, six views:** `showView(name)` (`index.html`) switches between `home`, `overview`, `repeats`, `sales`, `auditlog` and `form`. `<body data-view>` decides what shows (CSS), the dashboard views call `openDashboard()` and `dbSwitchTab()`, and the form page and the dashboard are only hidden, so nothing typed or chosen is lost. The tab bar inside the dashboard (`.db-tab-bar`) is kept in the page, hidden, because `dbSwitchTab` and the tests use it.
- **Home:** the numbers (`homeUpdate()`) come from the rows the dashboard holds and are counted like the Overview tiles; the texts about targets, cache ages and the deck limit are tested against `SCORECARD_THRESHOLDS`, the `LS_*` constants and `SC_MAX_AGENTS`, so change both together.
- **Loading screen:** the first element of the body (`#app-splash`), taken away by `hideSplash()` when the page init finishes, at least 600 ms after the start and always after 8 s.
- **KPI tiles:** the row (`.db-stats-row`) is full width and the tiles are `flex: 1 1 140px; max-width: 230px`. Do not put a width limit around the stats rows; an auto-fit grid with a fixed maximum counts its columns by the maximum and fits fewer tiles.
- **Critical Flags per Audit:** `dbDrawFlagsPerAudit` (after the shared helper block, not inside it) draws audits by number of flags (0, 1, 2, 3, 4 or more). One slice per flag text does not work because nearly every flag text is different.
- **The dashboard panel is in the page flow:** `#dashboard-panel` is `position: relative` under the sticky bar, and `#db-topbar` sticks at `top: var(--nav-h)`. Do not make the panel `position: fixed` again: the Apps Script frame can put the page about 16 px lower than the top of the window, and a fixed panel then starts behind the bar. `tests/navigation-test.js` checks both rules.
- **One filter, the same audits on every tab (`tests/filter-consistency-test.js`):** the test loads generated audits through `dbHandleData`, applies 23 filter combinations with `dbApplyAll` and compares every tab, tile, table, chart, export, the scorecard selection, Home and the AI text with an answer it works out itself. When a number is added to the dashboard, add its check there. The rules it holds: the Audit Log tab filters the merged rows (`_dbAuditView`: a blank Audit_Log cell is filled from Dashboard_Data like on Repeats and Sales); "Audits with Flags" (Repeats and Sales) counts audits and "Total Critical Flags" (Overview) counts flags; a flag written twice in one audit is one flag (`_dbSplitFlags`, `_scSplitFlags_`); the AI is sent the applied filters and the tab (`_dbAiFilters`), and the Overview is "both analyzers together"; the AI Analytics panel uses the same local-midnight day boundaries as the tabs; Home ignores filters.
- **Limits that the code respects (`tests/review-round2-test.js`):** a Sheets cell holds 50,000 characters, so the saved transcript is cut at 49,000 (`_cellSafeTranscript_`; the AI still gets up to 150,000); text from the browser that is written to a sheet goes through `appendRow` (the safe wrapper) or is reduced to a known value first (`getAIAnalytics` takes only `sales` or `repeats`); every function that sends mail from the owner's mailbox counts against `_emailRateTake_`; the AI calls send ten fields per audit (`AI_ROW_FIELDS`, `_dbAiRows`) and `buildAnalyticsContext_` must not read a field that is not in that list (a test compares the text for whole rows and for ten fields).
- **Resize:** Google Charts keep the width they were drawn at, so `dbRedrawActiveView` clears and draws the open dashboard view again 400 ms after the window size settles.
- **Colours:** the design tokens are in `:root`; add a colour as a token and add its pairs to the contrast test in `tests/accessibility-test.js`. The report documents keep their TELUS purple on purpose.
- **The palette is the TELUS table:** purples `DFD5E5 C1AFCD B9A0CA 946DAF 6E3D8F 4B286D`, greens `DBF9BD C7EF9F AAE074 84D533 66CC00 5AB400 248700`, greys `54595F 6A6E74 9DA5AA`. Navigation and deep purple are `4B286D`, the open tab `6E3D8F` (white text, green underline), the accent `66CC00` (on purple only, 5.6:1 on `4B286D`), green text `248700`. Do not take colours from screenshots: a wide-gamut screen shows them more saturated. Test 7 of `tests/accessibility-test.js` fails if a brand token leaves the table. The hero start `3A1F57`, the tier tints and the status colours are derived shades, not table colours. Green text on the mint box uses `094218` (`248700` on `F0FFF4` is only 4.47:1).

## 11. Error Handling

### Non-Fatal Patterns

Functions that must not break the main submission flow catch and log errors:
- `notifyAdmins()` — email failures are logged, not re-thrown
- `Analysis write`, `Dashboard_Data write`, `Audit_Log write` — each in its own `try/catch` (`writeAuditRecords_`), so one failing never stops the others; each adds a warning naming the audit reference
- Dashboard cache write (`writeDashboardCache`) — a failure is logged and the dashboard is served from the sheet; it can never blank the dashboard
- `generateAuditPDF()` — returns `null` on failure; email sends without attachment

### Timeout Prevention

| Risk | Solution |
|---|---|
| Roster (10k rows) | Cached in-memory + CacheService |
| AT Data GCP (4.5k rows) | Cached in-memory + CacheService |
| FuelIX API latency | Already async; processing overlay shown to user |
| PDF generation | `try/finally` ensures temp Doc always trashed |
| CacheService 100KB limit | The limit is in **bytes**, not characters, and accented or non-Latin text is up to 3 bytes a character. Dashboard cache chunks are 30,000 characters; result and roster payloads over 99,000 characters are not put in the script cache at all |

### Client-Side Error Display

- Form errors → red banner at top of form
- Dashboard load errors → shown in loading div
- Analytics errors → shown in panel body
- Email send errors → `alert()` + button re-enabled

---

## 12. Access Control

Every public function (no trailing underscore) in a `.gs` file can be called by **any signed-in user** from the browser console with `google.script.run`. So access is decided function by function, in three groups. `tests/access-control-test.js` enforces all three and fails if a page starts calling a function it does not list.

| Group | Rule | Examples |
|---|---|---|
| **Browser** — the pages really call them | Public, **not** admin-gated | `submitTranscript`, `getRosterBySapId`, `sendAuditEmail`, `getDashboardData`, `getEvalViewData`, `generateAndServePDF`, `sendAnalyticsBriefing` (24 in all) |
| **Maintenance** — an admin runs them by hand | Start with `requireAdmin_()` | `repairSheetHeaders`, `backfill*`, `diagnose*`, `test*`, `applyDuplicateAuditRefRenames`, `appendRegistryRow`, `verifyLatestAudit` |
| **Internal** — only server code calls them | Renamed with a trailing `_`, which Apps Script does not expose to the browser | `callFuelIX_`, `getAllRosterData_`, `lookupBySapId_`, `lookupSapId_`, `lookupVTID_`, `saveCachedResult_`, `updateCachedResult_`, `findCachedResult_`, `getTeamDetailsFromAuditHistory_`, and (code review batch 1) everything that sends mail, writes the shared caches, spends AI quota or returns rosters, addresses or the audit log: `notifyAdmins_`, `writeDashboardCache_`, `readAuditLog_`, `resolveObserver_`, `analyzeTranscript_`, `lookupAgentEmail_`, `logSubmissionError_` and 25 more |
| **Pure** — only turn arguments into text or HTML | Public and harmless; must not touch a sheet, mail, cache, Drive or the network | `esc`, `cell`, `section`, `build*HTML`, `sanitiseHTMLForPDF`, `htmlToPlainText` |
| **Accepted** — public for a stated reason | `doGet` and `doPost` (Apps Script entry points); helpers that need a Sheet object, which the browser cannot send (`appendRow`, `ensureHeaders`, `getOrCreateSheet`, `getOrCreateSpreadsheet`, `openSheetWithRetry`) | each has its reason written next to it in `tests/access-control-test.js` |

Test 8 in `tests/access-control-test.js` fails when a public function is in no group or in two, when a group names a function that is no longer public, or when a **Pure** function uses a service (`SpreadsheetApp`, `MailApp`, `UrlFetchApp`, `CacheService`, …). A new public function therefore cannot appear by accident.

### `isAdminCaller_()` and `requireAdmin_()` (`Config.gs`)

The person **actually running the call** must be on `ADMIN_USERNAMES`, signed in on a domain in `ADMIN_EMAIL_DOMAINS` (`telus.com`, `telusinternational.com`). It deliberately does **not** fall back to `Session.getEffectiveUser()`: the web app runs as *Execute as: Me*, so the effective user is the **owner** for every caller, and a fallback would make anyone whose own email is hidden look like an admin. A blank or unreadable identity is refused. `requireAdmin_()` throws "This function is for the QA and LS admin team only."

The screen's TEST toggle (`getObserverInfo`) and the server's TEST decision (`sendAuditEmail`) both use `isAdminCaller_()`.

### Adding a new function

- Called by a page → add it to the `BROWSER` list in `tests/access-control-test.js` and keep it public and ungated.
- Run by hand by an admin → put `requireAdmin_();` first and list it under `MAINTENANCE`.
- Only called by other server code → give it a trailing `_`.

### The e-mail relay guard (`EmailGuard.gs`, `sendAuditEmail`)

`sendAuditEmail` sends from the owner account, so what the browser sends is not trusted:

- **Who:** only an admin, or the analyst who saved the audit (matched by observer name, e-mail or the part before the `@`), may email it. Anyone else is refused, nothing is sent, and the attempt is logged as `sendAuditEmail:notYourAudit`.
- **To whom:** recipients must be on a company domain (`telus.com`, `telusinternational.com`, plus the Script Property `EMAIL_EXTRA_DOMAINS`, comma separated). Others are dropped, logged (`sendAuditEmail:externalRecipient`) and returned as `droppedRecipients`; if none is left nothing is sent.
- **How often:** 30 e-mails per person per hour (admins 100); the cache key is hashed.
- **What:** agent name, SAP ID, interaction id and audit ref come from the saved row first; the observer e-mail is the signed-in caller; form text is reduced to one line (control characters become spaces) so it cannot add a mail header.
- **Sheets:** `updateDashboardPDFLink_` stores the status (cut to 40 characters) and recipients as text, and accepts only Yes, No or Partially resolved as the analyst's Issue Resolved override.

### Limits, folders and AI replies (code review batch 3)

- **Tries:** `createScorecardDeck` and `exportAudits` count every try per person per hour *before* the lock and the sheet reads (30 and 60), on top of the finished-build limits (5 decks, 10 exports). The helper is `_hourlyTake_(prefix, who, max)` in `Config.gs`; it fails open with a log line when the cache is down.
- **Folders:** `_ownedFolder_(name)` (`Config.gs`) returns the folder the app owner *owns*; a same-named folder that was merely shared with the owner is ignored. Used for NH Analyzer Scorecards, NH Analyzer Exports and the PDF folder.
- **AI replies:** a 200 that is not JSON, or has empty text, is `AI_SERVICE_ERROR[invalid]` / `[empty]` and is retried inside the shared budget. The 300 s budget counts from the start of the execution (`_AI_EXEC_STARTED_AT_`). The transcript goes through `_fenceTranscript_`: cut at 150,000 characters, a forged end marker defused.
- **Repair tools:** `repairSheetHeaders` keeps unknown columns after the standard ones and holds the script lock. `repairNextTruncated` skips an interaction shared by more than one agent and only touches rows with the same SAP ID.

### Accessibility (code review batch 4)

- **Names:** every form label has `for`; every control, icon button and search box has a label or `aria-label`; `tests/accessibility-test.js` fails if one loses it.
- **Dialogs:** `A11Y_DIALOGS` in `index.html` lists them. When one opens the focus moves in, Tab stays inside, Escape closes it (`close: null` = a required choice, no Escape), and the focus returns to the opener. A new dialog needs `role="dialog"`, `aria-modal`, a name, and an entry in the list.
- **Pills, tabs, badge, stars:** pills are `role="button"` with `aria-expanded`; tabs use `role="tab"` with arrow, Home and End keys; the Yes/No badge and the EvalView stars work from the keyboard. Anything that is a button only by role (a span) gets Enter and Space from `a11yKeydown`.
- **Look:** keyboard focus is a 3 px purple ring with a white halo; body text is never under 12 px; grey text is `#6B6B6B` or darker.

### Names and references in the dashboard

A name or an Audit Ref from the sheet is never put inside an inline handler. The browser decodes an attribute before running it, so `escAttr` does not protect a handler: a name containing a quote and a closing bracket would run as code. Handlers are fixed text (`dbPillOnly('member',this.getAttribute('data-v'))`) and the value travels in `value` or a `data-*` attribute. `tests/dashboard-handlers-test.js` decodes each handler the way the browser does and runs it with hostile names.

### Registry endpoint

`doPost` accepts an Action Registry entry only when `payload.key` equals the Script Property `REGISTRY_SECRET`. Until that property is set it falls back to the old committed constant (logged as a warning). `appendRegistryRow` is admin-only; `doPost` calls the private `appendRegistryRow_`.

---

## 13. HTML Sanitiser — `HtmlSanitizer.gs`

`sanitizeEvalHtml_(html)` reduces model output and an analyst's edits to a short allow-list before they are stored, emailed, turned into a PDF or put on a page. The model output contains whatever was said on the call, so a transcript could ask for markup that runs.

It is a **pure function of a string** (no DOM): it reads the input as a token stream and **rebuilds** the markup, so only listed tags and attributes come out, every attribute value is re-escaped, anything it does not understand becomes text, and unclosed tags are closed. The same text lives in three files and `tests/html-sanitizer-test.js` asserts they are identical: **change `HtmlSanitizer.gs`, `index.html` and `EvalView.html` together.**

| Allowed | Notes |
|---|---|
| Layout and text tags | `div span p br hr ul ol li b strong i em u s small sub sup mark abbr cite q del ins code pre blockquote h1–h6 table thead tbody tfoot tr th td caption colgroup col a section article header footer aside dl dt dd details summary figure figcaption` |
| `button` and plain SVG shapes | `svg g path polygon polyline circle rect line`. Only for the Play button: a button always comes out `type="button"` and there is no script, style, link, image, animation or `foreignObject` |
| Attributes | `class title colspan rowspan align role scope`, `data-*`, `aria-*`, a filtered `style`, SVG presentation attributes, and `href` on `a` only |

| Always removed | |
|---|---|
| Elements **with their content** | `script style iframe object embed form textarea select template noscript math` and the like |
| Attributes | every `on*` handler, `id`, `contenteditable`, `srcdoc`, `formaction`, `xlink:href` |
| URLs | anything but `http`, `https`, `mailto`, `tel` or `#` after entity decoding and removal of whitespace and control characters |
| `style` values | `url(`, `expression`, `@import`, `javascript:`, comments, backslashes, `position: fixed/absolute/sticky` |

**Where it runs on the server:** `saveCachedResult_` (so `updateCachedResult_` too), `sendAuditEmail`, `generateAndServePDF`, `sendAnalyticsBriefing`, and the model output in `produceEvaluation_`.

**The report stylesheet.** The server sends the report as `sharedCSS() + html`, with its own `<style>` block first. The sanitiser removes every `<style>`, so the page wrapper `sanitiseResultHTML` keeps the **one block at the very start**, and only if it holds nothing but plain rules; a later or model-supplied `<style>` is dropped. `generateAndServePDF` puts the server's own `sharedCSS()` back instead of trusting the browser's.

**`contenteditable` is stripped on purpose**, so the evaluation page agents open is read-only. The report page sets it again itself (`_makeReportEditable`) for the modal and the applied view. `?ref=` is validated with `isValidAuditRef_` (letters, digits, dot, underscore, hyphen, up to 64) and everything printed into the evaluation page's script block has `<`, `>`, `&` and the line separators escaped.

### Editing the report, and the Play button (`index.html`)

- `_makeReportEditable(root)` marks the call summary, the bullets under *What's Working* and *What Needs to Change*, the sample statements and the flag text as editable. The *How to Change It* list item is **not** editable, only the statement inside it: the outermost `contenteditable` is the editing host, so an editable list item made Select-All plus typing replace the label, the statement and the Play button with plain text.
- A Play button speaks what its statement says **now** (`_ttsTextOf`), not its `data-tts`, which is only what the statement said when the button was made. An emptied statement plays nothing. All double quotation marks are removed from the text sent to the voice.
- `_syncTtsText(root)` copies the edited text into `data-tts` before the HTML is applied, emailed or saved, so the Play button on the saved evaluation page says what the analyst wrote.
- A Play button deleted by backspacing is put back when the statement loses focus (`_ensurePlayButton`).

---

## 14. Audit Check — `AuditCheck.gs`

`verifyLatestAudit(ref)` and `verifyMyLatestAudit()` are **admin-only and read-only**: they write no row and no cache entry. Run one from the Apps Script editor (open the **AuditCheck** file, choose the function in the dropdown, Run) and read the **Execution log**. No new version is needed, because the editor runs the saved code.

- `verifyMyLatestAudit()` checks the newest audit whose **Observer Name** is the person running it (worked out with `resolveObserver`). It never falls back to someone else's audit: no match is a failure that names the observer.
- `verifyLatestAudit()` checks the newest audit in the whole sheet, which on a busy day is usually another analyst's. Pass a reference to check a specific audit.

| Check | Fails when |
|---|---|
| Audit_Log | no row, or the same reference more than once (the evaluation link would open the first) |
| Transcripts, Analysis | no row for the reference, or the text is empty |
| Dashboard_Data | no row, a duplicate, or `Interaction ID`, `SAP ID`, `Team Member`, `Call Summary` blank (other blank fields, such as VTID, are reported as information) |
| Sheets agree | the SAP ID or Interaction ID differs between sheets |
| Saved evaluation | none in the **Cache sheet** (it reads the sheet directly, because the script-cache copy would hide a damaged sheet copy), contains `#ERROR!`, or looks cut off |
| Email | Email Status is a failure reason, or `Sent` with no recipients (`Not Sent` is information) |
| Flag | a "being processed" flag is still set |
| Error_Log | an error was logged for the interaction since the audit |

It checks the **data**, not the screen: whether the email reached an inbox, how the report looks and whether the Play button works must still be looked at.

### `verifyDashboardData()` — the data behind the dashboard

Admin-only and read-only, like the two above. The dashboard is built **from Audit_Log** and joined to Dashboard_Data by Audit Ref, then computes its KPIs from the merged rows; this function checks that join and the values the charts depend on, and works out the KPIs **the way the dashboard does**, so they can be compared with the screen. It reports:

- row counts, and missing columns (the code reads columns by name, so a missing header reads as blank)
- audits in Dashboard_Data but **not** in Audit_Log (the dashboard hides them), audits in Audit_Log with no Dashboard_Data row (counted with blank analysis), duplicate refs, rows with no ref, and a different SAP ID for the same ref in the two sheets
- audits whose Analysis Type is neither Repeats nor Sales (in no tab)
- Issue Resolved: the counts the dashboard shows (Resolved / Partially resolved / Unresolved / Not recorded). The text is read by its START ("Yes", "Partially resolved - ...", "Not resolved ..."); text with no verdict is a warning and shows as Not Recorded
- Repeat Projection %: numeric, 0%, blank, not a number, over 100, and the average (the dashboard uses the same one: a real 0% counts, blank and text do not)
- blank, invalid or future timestamps (a date filter hides them), a Team Member over 80 characters (a shifted column), email statuses, and Call Drivers written several ways

`tests/dashboard-correctness-test.js` runs the real dashboard code (index.html and Dashboard.html) against small datasets with obvious answers. It began as the audit behind Action Registry #163 and now fails if any of the 18 confirmed defects returns.

### How the dashboard counts (Action Registry #164)

- **Issue Resolved** is read by the start of the text: Resolved, Partially resolved (its own count), Unresolved, or Not Recorded (blank, or text with no verdict; never guessed). The KPI cards and the Resolution donut use the same rule.
- **Avg Repeat Risk** averages the audits that have a number; a real 0% counts, a blank or text does not.
- **Critical Flags** are stored in one cell. New audits join flags with ` | `; older rows used `, ` and are split so that a piece starting with a lower-case letter stays with the flag before it. `None` is no flag.
- **Top Agent** for a flag is the agent with the most of it. A top performer or a most-flags agent is named only when one agent is strictly lowest or highest.
- **Merge:** every Dashboard_Data column reaches the dashboard, an audit that exists only in Dashboard_Data is kept, and the status line reports audits with no analysis row or no Analysis Type.
- The helper functions are one block, identical in `index.html` and `Dashboard.html`; a test compares them, so change both together.

### Data repair tools (`DataRepair.gs`, Action Registry #165)

Admin-only. Run from the Apps Script editor.

| Function | What it does |
|---|---|
| `addMissingHeaders()` | Writes a header that is missing at the **end** of a sheet's header row (for example `Agent Email`). It moves and clears no data, and it does nothing if the existing headers are not exactly the first part of the expected list (it logs why). Safe to run twice. |
| `backfillAuditorEmail()` | Fills a **blank** `Auditor Email` in Dashboard_Data from the Users sheet, by the row's Observer Name. Writes only that column and never overwrites an email. A name that is not on the Users sheet is left blank and listed in the log. A name on the Users sheet with several different addresses gets all of them, separated by `, `, and is listed in the log as "several addresses" (it may be one person with two addresses or two people with one name). Run `addMissingHeaders()` first so the column exists. |
| `diagnoseDataIssues()` | Read-only. Prints the exact rows behind the duplicate ref, an audit that is only in Dashboard_Data, SAP ID mismatches, non-numeric Repeat % values and shifted Team Member rows. |

| `diagnoseDashboardGaps()` | Read-only. Three reports about Audit_Log: how many Duration cells Sheets turned into a time of day, where the "Not Sent" audits are (by month, type and observer), and how the Issue Resolved answers with no clear verdict look and start. |
| `previewDataRepairs()` | Read-only. Lists what `applyDataRepairs()` would change and what it would leave alone. |
| `applyDataRepairs()` | Rebuilds an Audit_Log row for an audit that exists only in Dashboard_Data (needs SAP ID, Interaction ID, Team Member and Timestamp; Recipients stays blank). Corrects a Dashboard_Data SAP ID only when the other three sheets agree and Audit_Log is not the odd one out. Turns a Repeat Projection % sentence ("95% repeat risk. ...") into the number and keeps the sentence as a cell note. Skips and lists anything unclear. Stops before the time limit; run again if it says it did not finish. Safe to run twice. |

Use `addMissingHeaders()` instead of `repairSheetHeaders()` when only a header is missing: `repairSheetHeaders()` clears and rewrites the whole sheet.

A duplicate audit reference (two audits sharing one ref) is handled by `reportDuplicateAuditRefs()` (read-only) then `applyDuplicateAuditRefRenames()`, described above.

The dashboard reads a Repeat Projection % cell as a number, or as the first `NN%` inside a sentence, because older audits stored a sentence there. It also reads a bare number above 0 and up to 1 as a fraction (`0.78` is 78%), because Sheets turns the text "78%" into a percent number; a written `1%` stays 1 and a bare `1` means 100%.

`applyDataRepairs()` also cleans cells that hold more than they should: a Team Member that lists every participant or has transcript header text glued on becomes the clean Team Member the **same SAP ID** has on its other audits in Audit_Log (only if that name is inside the messy cell; otherwise the row is skipped and listed); an Interaction ID longer than a UUID is cut to 36 characters; Direction becomes Inbound or Outbound; a Duration with text glued on becomes `h:mm` or `h:mm:ss` and is stored as text. The original value of every cleaned cell is kept in a cell note. A Duration that Sheets stored as a time of day (a Date on 1899-12-30) is turned back into text such as `2:15:05` the same way; anything that is not a time within one day is skipped and listed.

### Scorecard Slides (Action Registry #168)

The **Generate Slides** button in the header of the **Team Member Trends & Repeat Offenders** section (Overview tab, beside the High / Moderate / Low legend) builds a Google Slides deck for the agents in the **current dashboard filters**: a cover slide (period, filters, a legend of the targets) and one slide per agent. It is open to every signed-in analyst. The page sends only the Audit Refs; the server reads the sheets and computes everything.

| On the slide | Where it comes from |
|---|---|
| Avg Repeat Risk | Average of the audits that have a number (a real 0% counts; a sentence or `0.78` is read like the dashboard does). Target 30% or lower. |
| Resolution rate | Resolved of the audits that have a clear verdict (Not recorded is left out). Target 80% or higher. Partial and Unresolved are shown under the number. |
| Critical flags | Flags per audit. Target 0, judged on the real average (1 flag in 21 audits is Below threshold and is shown as `<0.1`, not `0.0`). The AI template writes 2 to 5 flags per audit, so most agents show red here. |
| Transfer rate / Sales attempted | Transfer rate for Repeats audits (Yes or No; a sentence is not counted), target 20% or lower. A Sales-only agent gets Sales attempted instead (Yes or No; N/A is not counted), target 80% or higher. |
| Badge | No tile below target: On track. One: Needs attention. Two or more: Priority support. |
| Strengths | The "What's Working" items of the agent's newest 3 reports; if there are none, the areas that meet their target. The Analysis sheet holds the report as PLAIN TEXT (`htmlToPlainText`), so the bullets under the heading are read; the HTML form is still understood. Dates on the slide use the script time zone, not UTC. |
| Gaps / Coaching | The most common Critical Flags (then RCA sub-parameters) / the most common Top SMART Opportunity. |

The targets are in `SCORECARD_THRESHOLDS` in `Config.gs`. The look (positions, colours, sizes, text) is in `ScorecardLayout.gs`; the numbers and wording are in `ScorecardMetrics.gs`; the PowerPoint package is written by `ScorecardPptx.gs` and handed to Drive by `ScorecardSlides.gs`, which converts it to Google Slides (the Google Slides API is not used).

**Limits:** up to 40 agents per deck, 2,000 audits, 5 decks per person per hour, one deck built at a time. The deck is filed in the folder **NH Analyzer Scorecards** in the app owner's Drive and shared, as editor, only with the person who asked.

**Admin tools** (Apps Script editor): `previewScorecards(refs)` logs the numbers without creating anything, and `testSlidesAccess()` builds a one-slide sample deck.

**No Google setting or new permission is needed.** The deck is a PowerPoint file that Drive converts to Google Slides, the same Drive conversion the PDF export uses. After a deploy, the app owner can run `testSlidesAccess` once to see a sample deck.

### Export to Google Sheets (Action Registry #169)

The Export All / Export Filtered buttons call `exportAudits(view, auditRefs, title)` in `ExportService.gs`. The page sends only the view (`repeats`, `sales` or `auditlog`) and the Audit Refs; the server builds the rows from the sheets (the same merge the dashboard uses), stores every cell as text, and files the new sheet in the folder **NH Analyzer Exports** in the app owner's Drive, shared as **viewer** with the person who asked, and nobody else. Limits: up to 5,000 audits, 10 exports per person per hour, one export at a time. The old `exportToSpreadsheet` (rows from the browser, shared with the whole domain) is gone.

### AI reliability (Action Registry #169)

- An AI answer that was **cut off** by the token limit is refused: nothing is saved and the analyst is told to submit again.
- The retries share one **300 s budget**, so they can never run past the 360 s Apps Script limit; a retry is not made when under 60 s is left.
- A **saved draft** keeps its Audit Ref when the page loads.

---

*Documentation updated October 2026 — NH FCR, Transfer or Sales Call Analyzer (Action Registry #155–#169)*
