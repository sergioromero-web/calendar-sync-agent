/**
 * B2B calendar sync — read-only JSON export for the adoption dashboard.
 *
 * Paste into an Apps Script project with read access to the Central Ops Availabilities
 * spreadsheet, then Deploy → New deployment → Web app. See SETUP below.
 *
 * DESIGN: returns tabs RAW — tab names, header row, value rows — and interprets nothing. All
 * interpretation lives in sheets.js on the dashboard side, so a header change is a mapping fix
 * rather than a redeploy here, and no data is dropped for not being understood yet.
 *
 * Domain truth (what the tabs mean, which ones to skip, how to schedule the ingest) is in
 * docs/agent-reference.md and docs/limitations.md. Two things from there that constrain THIS file:
 *
 *   - Timestamps in the sheet are Date objects rendered by cell format. We read with getValues()
 *     and serialize to ISO here. NEVER display values: the same logical column is formatted
 *     differently across tabs, so display strings are not comparable.
 *   - We cannot synchronise with the agent. Its ScriptLock is scoped to ITS Apps Script project
 *     and this is a different project, so no lock we take here would mean anything. The agent
 *     calls clearContents() while rewriting tabs, so a read can land mid-write. Consistency comes
 *     from WHEN the ingest runs (y:30 — the agent starts at :44:30 and has been observed running
 *     27 min, so the quiet window is :15-:40) plus snapshot validation on the dashboard side —
 *     never from locking. Read-only and lock-free by necessity, not by preference.
 *
 * SETUP
 *   1. Project Settings → Script properties:
 *        SHEET_READ_SECRET     (required) long random string; must match the dashboard env var
 *        CENTRAL_SPREADSHEET_ID (optional) Central Ops Availabilities; omit to use the bound one
 *   2. Deploy → New deployment → Web app → Execute as: Me · Who has access: Anyone.
 *      "Anyone" is required for the Hive backend to reach it; the secret is the access control.
 *   3. After editing this file: Deploy → Manage deployments → Edit (pencil) → New version.
 *      Without a new version the deployed URL keeps serving the OLD code.
 *
 * API (always HTTP 200 — Apps Script web apps cannot set status codes, so the caller MUST branch
 * on the `ok` field in the body, never on the HTTP status. `status` mirrors what the code would
 * have been.)
 *
 *   ?secret=…&resource=tabs
 *     Discovery. Every tab with its header row, column count and data row count. Cheap — reads
 *     only row 1 per tab. `data_rows` is what lets the ingest spot a tab caught mid-clearContents
 *     (empty or header-only) and discard that source rather than the whole run.
 *
 *   ?secret=…&resource=rows&tab=<name>[&offset=0][&limit=2000]
 *     Paginated data rows for one tab. offset is 0-based over DATA rows (the header is not a data
 *     row). Paginated from day one: the 6-minute execution limit and the response size cap are not
 *     problems worth discovering the day a tab gets big.
 */

var PROP_SECRET = "SHEET_READ_SECRET";
var PROP_SPREADSHEET_ID = "CENTRAL_SPREADSHEET_ID";

var DEFAULT_LIMIT = 2000;
var MAX_LIMIT = 5000;

// A single runaway cell must not blow up the whole response. Truncated cells are marked so the
// ingest can tell "this value is long" from "this value ends here". Distinct from the agent's own
// truncateText suffix ("...") so error grouping can strip both without confusing them.
var MAX_CELL_CHARS = 4000;
var TRUNCATED_SUFFIX = "…[truncated]";

function doGet(e) {
  var params = (e && e.parameter) || {};
  try {
    requireSecret_(params.secret);
    var resource = String(params.resource || "tabs").toLowerCase();
    if (resource === "tabs") return json_(listTabs_());
    if (resource === "rows") return json_(readRows_(params));
    throw httpError_("unknown resource '" + resource + "' (use: tabs, rows)", 400);
  } catch (err) {
    var status = err && err.status ? err.status : 500;
    return json_({
      ok: false,
      status: status,
      error: err && err.message ? err.message : String(err),
      fetched_at: nowIso_(),
    });
  }
}

// ── Auth ─────────────────────────────────────────────────────────────────────

function requireSecret_(provided) {
  var expected = PropertiesService.getScriptProperties().getProperty(PROP_SECRET);
  if (!expected) {
    // Misconfiguration, not a bad request: fail closed rather than serving the sheet open.
    throw httpError_("Script property " + PROP_SECRET + " is not set", 503);
  }
  if (!provided || !timingSafeEqual_(String(provided), String(expected))) {
    throw httpError_("forbidden", 403);
  }
}

// Compares in time independent of where the first mismatch is. Overkill for an internal tool,
// but it costs three lines and removes the question entirely.
function timingSafeEqual_(a, b) {
  if (a.length !== b.length) return false;
  var diff = 0;
  for (var i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

// ── Spreadsheet ──────────────────────────────────────────────────────────────

/**
 * Everything the dashboard reads lives in ONE spreadsheet (Central Ops Availabilities): the queue,
 * the per-source baselines, the telemetry log and the audit tabs. The Bulk Update spreadsheet is
 * the agent's write destination and is not read here, so there is deliberately no second book.
 */
function spreadsheet_() {
  var id = (PropertiesService.getScriptProperties().getProperty(PROP_SPREADSHEET_ID) || "").trim();
  if (!id) {
    var bound = SpreadsheetApp.getActiveSpreadsheet();
    if (!bound) {
      throw httpError_(
        "No spreadsheet available: set the " + PROP_SPREADSHEET_ID +
        " script property, or bind this script to the spreadsheet",
        503
      );
    }
    return bound;
  }
  try {
    return SpreadsheetApp.openById(id);
  } catch (err) {
    throw httpError_(
      "Cannot open spreadsheet " + id + " (check the ID and that this script's owner has access)",
      502
    );
  }
}

// ── resource=tabs ────────────────────────────────────────────────────────────

function listTabs_() {
  var ss = spreadsheet_();
  return {
    ok: true,
    status: 200,
    resource: "tabs",
    fetched_at: nowIso_(),
    spreadsheet_id: ss.getId(),
    spreadsheet_name: ss.getName(),
    timezone: ss.getSpreadsheetTimeZone(),
    tabs: ss.getSheets().map(describeTab_),
  };
}

function describeTab_(sh) {
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  return {
    name: sh.getName(),
    index: sh.getIndex(),
    hidden: sh.isSheetHidden(),
    columns: lastCol,
    // The header row is not data. An empty tab reports 0 rather than -1. A tab read mid-write
    // shows up here as 0 (or a header-only 0), which is the ingest's cue to skip that source.
    data_rows: Math.max(0, lastRow - 1),
    headers: readHeaders_(sh, lastRow, lastCol),
  };
}

// ── resource=rows ────────────────────────────────────────────────────────────

function readRows_(params) {
  var tabName = String(params.tab || "").trim();
  if (!tabName) throw httpError_("missing 'tab' (call resource=tabs to list them)", 400);

  var offset = toInt_(params.offset, 0);
  if (offset < 0) offset = 0;
  var limit = toInt_(params.limit, DEFAULT_LIMIT);
  if (limit < 1) limit = DEFAULT_LIMIT;
  if (limit > MAX_LIMIT) limit = MAX_LIMIT;

  var ss = spreadsheet_();
  var sh = findTab_(ss, tabName);
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  var total = Math.max(0, lastRow - 1);

  var count = Math.min(limit, Math.max(0, total - offset));
  var rows = [];
  if (count > 0 && lastCol > 0) {
    rows = sh.getRange(2 + offset, 1, count, lastCol).getValues().map(function (row) {
      return row.map(cellOut_);
    });
  }

  return {
    ok: true,
    status: 200,
    resource: "rows",
    fetched_at: nowIso_(),
    spreadsheet_id: ss.getId(),
    // The sheet's timezone is what a naive Date cell is relative to. Sent so the ingest never has
    // to assume one.
    timezone: ss.getSpreadsheetTimeZone(),
    tab: sh.getName(),
    headers: readHeaders_(sh, lastRow, lastCol),
    rows: rows,
    offset: offset,
    limit: limit,
    returned: rows.length,
    total: total,
    has_more: offset + rows.length < total,
  };
}

/**
 * Exact tab name only, and a loud failure listing what IS there. No case-insensitive or fuzzy
 * fallback: tab names come from the source registry, which sets them explicitly, so a name that
 * does not match means the registry and the sheet have diverged. That is exactly the kind of thing
 * we want to see rather than paper over.
 */
function findTab_(ss, tabName) {
  var sh = ss.getSheetByName(tabName);
  if (sh) return sh;
  var available = ss.getSheets().map(function (s) { return s.getName(); });
  throw httpError_("tab '" + tabName + "' not found. Available: " + available.join(", "), 404);
}

/**
 * Header row via getValues(), like everything else — never display values. Headers are plain
 * strings that ensureSheetHeader rewrites when they drift, so the typed read and the rendered
 * read are identical here; using getValues() keeps a single rule with no exception to remember.
 */
function readHeaders_(sh, lastRow, lastCol) {
  if (lastRow < 1 || lastCol < 1) return [];
  return sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return h === null || h === undefined ? "" : String(h);
  });
}

/**
 * Serialize one cell. Dates become ISO instants: the sheet stores real Date objects and renders
 * them by cell format, so the same logical column is formatted differently across tabs — ISO is
 * the only representation that is comparable. Numbers and booleans keep their type so the mapping
 * layer never has to re-parse strings.
 */
function cellOut_(v) {
  if (v === null || v === undefined || v === "") return "";
  if (v instanceof Date) {
    return isNaN(v.getTime()) ? "" : v.toISOString();
  }
  if (typeof v === "number" || typeof v === "boolean") return v;
  var s = String(v);
  return s.length > MAX_CELL_CHARS ? s.slice(0, MAX_CELL_CHARS) + TRUNCATED_SUFFIX : s;
}

// ── Helpers ──────────────────────────────────────────────────────────────────

function toInt_(v, dflt) {
  var n = parseInt(v, 10);
  return isNaN(n) ? dflt : n;
}

function nowIso_() {
  return new Date().toISOString();
}

function httpError_(message, status) {
  var err = new Error(message);
  err.status = status;
  return err;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
