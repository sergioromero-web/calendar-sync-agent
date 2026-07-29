function syncSource(cfg, runId) {
  validateSourceConfig(cfg);

  const landlord = SpreadsheetApp.openById(cfg.sheetID);
  const sheet = getSheetOrThrow(landlord, cfg.tab, '[' + cfg.key + ']');
  const timezone = landlord.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'UTC';
  const sourceRange = sheet.getDataRange();
  const data = sourceRange.getValues();
  const displayData = sourceRange.getDisplayValues();

  if (!displayData.length || isBlankRow(displayData[0])) {
    throw new Error('[' + cfg.key + '] Source sheet has no header row');
  }

  const header = displayData[0];
  const ruIdCol = findHeaderIndex(header, cfg.colRuId);
  const availFromCol = findHeaderIndex(header, cfg.colAvail);

  Logger.log('[' + cfg.key + '] ruId col: ' + ruIdCol + ' | availFrom col: ' + availFromCol);

  if (ruIdCol === -1) {
    throw new Error('[' + cfg.key + '] Column not found: "' + cfg.colRuId + '"');
  }
  if (availFromCol === -1) {
    throw new Error('[' + cfg.key + '] Column not found: "' + cfg.colAvail + '"');
  }

  const central = SpreadsheetApp.openById(CENTRAL_ID);
  const stateSheet = getOrCreateTab(central, cfg.stateTab);
  ensureSheetHeader(stateSheet, STATE_HEADERS);

  const stateMap = readStateMap(stateSheet, timezone);
  const queueIndex = loadQueueIndex(central, true);
  const now = new Date();
  const isFirstRun = Object.keys(stateMap).length === 0;
  const newState = Object.assign({}, stateMap);

  const queueRows = [];
  const telemetryRows = [];
  const stats = {
    rows: 0,
    queued: 0,
    baseline: 0,
    unchanged: 0,
    duplicateQueued: 0,
    invalidRuId: 0,
    invalidDate: 0
  };

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const displayRow = displayData[i];
    const ruId = normalizeRuId(displayRow[ruIdCol] || row[ruIdCol]);
    const availFrom = normalizeDate(displayRow[availFromCol] || row[availFromCol], timezone);

    if (!ruId) {
      stats.invalidRuId++;
      continue;
    }

    if (!availFrom) {
      stats.invalidDate++;
      continue;
    }

    stats.rows++;

    if (isFirstRun) {
      newState[ruId] = availFrom;
      stats.baseline++;
      continue;
    }

    if (stateMap[ruId] === availFrom) {
      stats.unchanged++;
      continue;
    }

    const key = queueKey(cfg.key, ruId, availFrom);
    if (queueIndex[key]) {
      stats.duplicateQueued++;
      continue;
    }

    queueRows.push(buildQueueRow(runId, cfg, ruId, availFrom, now));
    telemetryRows.push([
      now,
      runId,
      cfg.key,
      ruId,
      stateMap[ruId] || '',
      availFrom,
      'queued',
      'Queued for Bulk'
    ]);
    stats.queued++;
    queueIndex[key] = true;
  }

  if (isFirstRun) {
    writeStateMap(stateSheet, newState, now);
    appendTelemetryRows(central, [[
      now,
      runId,
      cfg.key,
      '',
      '',
      '',
      'baseline',
      'Baseline built with ' + stats.baseline + ' rows; nothing sent to Bulk'
    ]]);
    Logger.log('[' + cfg.key + '] First run baseline built: ' + stats.baseline + ' rows');
    return stats;
  }

  if (queueRows.length) {
    appendQueueRows(central, queueRows);
    appendTelemetryRows(central, telemetryRows);
    Logger.log('[' + cfg.key + '] Queued rows: ' + queueRows.length);
  } else {
    Logger.log('[' + cfg.key + '] No changes to queue');
  }

  return stats;
}

function pushToBulk(runId) {
  const central = SpreadsheetApp.openById(CENTRAL_ID);
  const queueSheet = ensureExportQueue(central);
  pruneOldSentQueueRows(queueSheet);
  const allQueueRows = readQueueRows(queueSheet, { skipOldSent: true });
  const invalidRows = markInvalidPendingQueueRowsFailed(queueSheet, allQueueRows);
  const queueRows = allQueueRows
    .filter(queueRow => isProcessableQueueRow(queueRow))
    .slice(0, MAX_BULK_ROWS_PER_RUN);

  if (!queueRows.length) {
    Logger.log('pushToBulk: nothing to push');
    return { pushed: 0, failed: invalidRows, pending: countPendingQueueRows(queueSheet) };
  }

  Logger.log('pushToBulk: ' + queueRows.length + ' rows to push');

  const now = new Date();
  markQueueRowsProcessing(queueSheet, queueRows, runId, now);

  try {
    const bulkSheet = getSheetOrThrow(SpreadsheetApp.openById(BULK_ID), BULK_TAB, '[Bulk]');
    const start = bulkSheet.getLastRow() + 1;
    const values = queueRows.map(queueRow => [
      queueRow.agentEmail,
      '',
      '',
      toBulkRuIdValue(queueRow.ruId),
      queueRow.availFrom
    ]);

    bulkSheet.getRange(start, 1, values.length, 1).setNumberFormat('@');
    bulkSheet.getRange(start, 5, values.length, 1).setNumberFormat('@');
    bulkSheet.getRange(start, 1, values.length, 5).setValues(values);
    Logger.log('pushToBulk: wrote ' + values.length + ' rows to Bulk at row ' + start);

  } catch (e) {
    const error = truncateText(formatError(e), 1000);
    markQueueRowsPending(queueSheet, queueRows, error);
    throw new Error('pushToBulk failed; rows left pending for retry. ' + error);
  }

  markQueueRowsSent(queueSheet, queueRows, new Date());
  applySentQueueRowsToState(central, queueRows, new Date());

  return { pushed: queueRows.length, failed: invalidRows, pending: countPendingQueueRows(queueSheet) };
}

function ensureExportQueue(central) {
  const ss = central || SpreadsheetApp.openById(CENTRAL_ID);
  const sheet = getOrCreateTab(ss, EXPORT_TAB);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, QUEUE_HEADERS.length).setValues([QUEUE_HEADERS]);
    return sheet;
  }

  const width = Math.max(sheet.getLastColumn(), QUEUE_HEADERS.length);
  const header = sheet.getRange(1, 1, 1, width).getValues()[0];

  if (headersMatch(header, QUEUE_HEADERS)) {
    return sheet;
  }

  if (isBlankRow(header)) {
    sheet.getRange(1, 1, 1, QUEUE_HEADERS.length).setValues([QUEUE_HEADERS]);
    return sheet;
  }

  if (headersMatch(header, LEGACY_EXPORT_HEADERS)) {
    migrateLegacyExportRows(sheet);
    return sheet;
  }

  throw new Error('Queue tab "' + EXPORT_TAB + '" has an unexpected header. Refusing to overwrite it.');
}

function migrateLegacyExportRows(sheet) {
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const migratedRows = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const agentEmail = String(row[0] || '').trim();
    const ruId = normalizeRuId(row[1]);
    const availFrom = normalizeDate(row[2], Session.getScriptTimeZone());
    const source = String(row[3] || 'legacy').trim();

    if (!agentEmail || !ruId || !availFrom) continue;

    migratedRows.push([
      Utilities.getUuid(),
      'legacy-migration',
      now,
      source,
      agentEmail,
      ruId,
      availFrom,
      QUEUE_STATUS.PENDING,
      0,
      '',
      '',
      'Migrated from legacy export format'
    ]);
  }

  sheet.clearContents();
  sheet.getRange(1, 1, 1, QUEUE_HEADERS.length).setValues([QUEUE_HEADERS]);
  if (migratedRows.length) {
    sheet.getRange(2, 1, migratedRows.length, QUEUE_HEADERS.length).setValues(migratedRows);
  }

  Logger.log('migrateLegacyExportRows: migrated ' + migratedRows.length + ' rows');
}

function buildQueueRow(runId, cfg, ruId, availFrom, now) {
  return [
    Utilities.getUuid(),
    runId,
    now,
    cfg.key,
    cfg.agentEmail,
    ruId,
    availFrom,
    QUEUE_STATUS.PENDING,
    0,
    '',
    '',
    ''
  ];
}

function appendQueueRows(central, rows) {
  if (!rows.length) return;

  const sheet = ensureExportQueue(central);
  const start = sheet.getLastRow() + 1;
  sheet.getRange(start, 5, rows.length, 3).setNumberFormat('@');
  sheet.getRange(start, 1, rows.length, QUEUE_HEADERS.length).setValues(rows);
}

function readQueueRows(queueSheet, options) {
  const opts = options || {};
  const now = opts.now || new Date();
  const range = queueSheet.getDataRange();
  const data = range.getValues();
  const displayData = range.getDisplayValues();
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const displayRow = displayData[i];
    if (isBlankRow(row)) continue;

    const raw = normalizeQueueRawRow(row);
    const rawDisplay = normalizeQueueRawRow(displayRow);

    const queueRow = {
      sheetRow: i + 1,
      raw: raw,
      queueId: String(raw[QUEUE_COL.QUEUE_ID] || '').trim(),
      runId: String(raw[QUEUE_COL.RUN_ID] || '').trim(),
      queuedAt: raw[QUEUE_COL.QUEUED_AT],
      source: String(raw[QUEUE_COL.SOURCE] || '').trim(),
      agentEmail: String(rawDisplay[QUEUE_COL.AGENT_EMAIL] || raw[QUEUE_COL.AGENT_EMAIL] || '').trim(),
      ruId: normalizeRuId(rawDisplay[QUEUE_COL.RU_ID] || raw[QUEUE_COL.RU_ID]),
      availFrom: normalizeDate(rawDisplay[QUEUE_COL.AVAIL_FROM] || raw[QUEUE_COL.AVAIL_FROM], Session.getScriptTimeZone()),
      status: String(raw[QUEUE_COL.STATUS] || '').trim().toLowerCase(),
      attempts: Number(raw[QUEUE_COL.ATTEMPTS] || 0),
      lastAttemptAt: raw[QUEUE_COL.LAST_ATTEMPT_AT],
      sentAt: raw[QUEUE_COL.SENT_AT],
      lastError: String(raw[QUEUE_COL.LAST_ERROR] || '')
    };

    if (opts.skipOldSent && queueRow.status === QUEUE_STATUS.SENT && !isRecentSentQueueRow(queueRow, now)) {
      continue;
    }

    rows.push(queueRow);
  }

  return rows;
}

function loadQueueIndex(central, includeSent) {
  const sheet = ensureExportQueue(central);
  const index = {};
  const now = new Date();

  readQueueRows(sheet, { skipOldSent: true, now: now }).forEach(queueRow => {
    if (!queueRow.source || !queueRow.ruId || !queueRow.availFrom) return;
    if (queueRow.status === QUEUE_STATUS.FAILED) return;
    if (!includeSent && queueRow.status === QUEUE_STATUS.SENT) return;
    index[queueKey(queueRow.source, queueRow.ruId, queueRow.availFrom)] = true;
  });

  return index;
}

function normalizeQueueRawRow(row) {
  const normalized = row.slice(0, QUEUE_HEADERS.length);
  while (normalized.length < QUEUE_HEADERS.length) {
    normalized.push('');
  }
  return normalized;
}

function isProcessableQueueRow(queueRow) {
  if (!queueRow.agentEmail || !queueRow.ruId || !queueRow.availFrom || !queueRow.source) return false;
  return queueRow.status === QUEUE_STATUS.PENDING;
}

function markInvalidPendingQueueRowsFailed(queueSheet, queueRows) {
  let failed = 0;
  const updates = [];

  queueRows.forEach(queueRow => {
    if (queueRow.status !== QUEUE_STATUS.PENDING) return;
    if (queueRow.agentEmail && queueRow.ruId && queueRow.availFrom && queueRow.source) return;

    const row = queueRow.raw.slice(0, QUEUE_HEADERS.length);
    row[QUEUE_COL.STATUS] = QUEUE_STATUS.FAILED;
    row[QUEUE_COL.LAST_ERROR] = 'Missing required queue data';
    updates.push({ sheetRow: queueRow.sheetRow, row: row });
    failed++;
  });

  writeQueueRowUpdates(queueSheet, updates);

  if (failed) {
    Logger.log('markInvalidPendingQueueRowsFailed: marked ' + failed + ' invalid rows as failed');
  }

  return failed;
}

function countPendingQueueRows(queueSheet) {
  return readQueueRows(queueSheet, { skipOldSent: true })
    .filter(queueRow => queueRow.status === QUEUE_STATUS.PENDING)
    .length;
}

function markQueueRowsProcessing(queueSheet, queueRows, runId, now) {
  const updates = [];

  queueRows.forEach(queueRow => {
    const row = queueRow.raw.slice(0, QUEUE_HEADERS.length);
    row[QUEUE_COL.RUN_ID] = runId;
    row[QUEUE_COL.STATUS] = QUEUE_STATUS.PROCESSING;
    row[QUEUE_COL.ATTEMPTS] = Number(row[QUEUE_COL.ATTEMPTS] || 0) + 1;
    row[QUEUE_COL.LAST_ATTEMPT_AT] = now;
    row[QUEUE_COL.LAST_ERROR] = '';
    updates.push({ sheetRow: queueRow.sheetRow, row: row });
    queueRow.raw = row;
    queueRow.runId = runId;
    queueRow.status = QUEUE_STATUS.PROCESSING;
    queueRow.attempts = Number(row[QUEUE_COL.ATTEMPTS] || 0);
    queueRow.lastAttemptAt = now;
    queueRow.lastError = '';
  });

  writeQueueRowUpdates(queueSheet, updates);
}

function markQueueRowsPending(queueSheet, queueRows, error) {
  const updates = [];

  queueRows.forEach(queueRow => {
    const row = queueRow.raw.slice(0, QUEUE_HEADERS.length);
    row[QUEUE_COL.STATUS] = QUEUE_STATUS.PENDING;
    row[QUEUE_COL.LAST_ERROR] = error;
    updates.push({ sheetRow: queueRow.sheetRow, row: row });
    queueRow.raw = row;
    queueRow.status = QUEUE_STATUS.PENDING;
    queueRow.lastError = error;
  });

  writeQueueRowUpdates(queueSheet, updates);
}

function markQueueRowsSent(queueSheet, queueRows, sentAt) {
  const updates = [];

  queueRows.forEach(queueRow => {
    const row = queueRow.raw.slice(0, QUEUE_HEADERS.length);
    row[QUEUE_COL.STATUS] = QUEUE_STATUS.SENT;
    row[QUEUE_COL.SENT_AT] = sentAt;
    row[QUEUE_COL.LAST_ERROR] = '';
    updates.push({ sheetRow: queueRow.sheetRow, row: row });
    queueRow.raw = row;
    queueRow.status = QUEUE_STATUS.SENT;
    queueRow.sentAt = sentAt;
    queueRow.lastError = '';
  });

  writeQueueRowUpdates(queueSheet, updates);
}

function writeQueueRowUpdates(queueSheet, updates) {
  if (!updates.length) return;

  updates.sort((a, b) => a.sheetRow - b.sheetRow);

  let index = 0;
  while (index < updates.length) {
    const startRow = updates[index].sheetRow;
    const values = [updates[index].row];
    let previousRow = startRow;
    index++;

    while (index < updates.length && updates[index].sheetRow === previousRow + 1) {
      values.push(updates[index].row);
      previousRow = updates[index].sheetRow;
      index++;
    }

    queueSheet.getRange(startRow, 1, values.length, QUEUE_HEADERS.length).setValues(values);
  }
}

function reclaimStaleQueueRows(runId) {
  const central = SpreadsheetApp.openById(CENTRAL_ID);
  const queueSheet = ensureExportQueue(central);
  const cutoff = new Date(new Date().getTime() - PROCESSING_STALE_MINUTES * 60 * 1000);
  let reclaimed = 0;
  const updates = [];

  readQueueRows(queueSheet, { skipOldSent: true }).forEach(queueRow => {
    if (queueRow.status !== QUEUE_STATUS.PROCESSING) return;
    if (queueRow.lastAttemptAt instanceof Date && queueRow.lastAttemptAt > cutoff) return;

    const row = queueRow.raw.slice(0, QUEUE_HEADERS.length);
    row[QUEUE_COL.RUN_ID] = runId;
    row[QUEUE_COL.STATUS] = QUEUE_STATUS.PENDING;
    row[QUEUE_COL.LAST_ERROR] = 'Requeued after stale processing status';
    updates.push({ sheetRow: queueRow.sheetRow, row: row });
    reclaimed++;
  });

  writeQueueRowUpdates(queueSheet, updates);

  if (reclaimed) {
    Logger.log('reclaimStaleQueueRows: reclaimed ' + reclaimed + ' stale rows');
  }

  return reclaimed;
}

function reconcileSentQueueToState() {
  const central = SpreadsheetApp.openById(CENTRAL_ID);
  const queueSheet = ensureExportQueue(central);
  const sentRows = readQueueRows(queueSheet)
    .filter(queueRow => queueRow.status === QUEUE_STATUS.SENT);

  return applySentQueueRowsToState(central, sentRows, new Date());
}

function pruneOldSentQueueRows(queueSheet, now) {
  const sheet = queueSheet || ensureExportQueue();
  const cutoff = getSentQueueRetentionCutoff(now || new Date());
  if (!cutoff) return 0;

  const rowsToDelete = readQueueRows(sheet)
    .filter(queueRow => queueRow.status === QUEUE_STATUS.SENT && isQueueRowOlderThan(queueRow, cutoff))
    .map(queueRow => queueRow.sheetRow);

  deleteSheetRows(sheet, rowsToDelete);

  if (rowsToDelete.length) {
    Logger.log('pruneOldSentQueueRows: pruned ' + rowsToDelete.length + ' sent rows older than ' + SENT_QUEUE_RETENTION_DAYS + ' days');
  }

  return rowsToDelete.length;
}

function isRecentSentQueueRow(queueRow, now) {
  const cutoff = getSentQueueRetentionCutoff(now || new Date());
  if (!cutoff) return true;
  return !isQueueRowOlderThan(queueRow, cutoff);
}

function isQueueRowOlderThan(queueRow, cutoff) {
  const timestamp = getQueueRowTimestamp(queueRow);
  return timestamp ? timestamp < cutoff : false;
}

function getSentQueueRetentionCutoff(now) {
  const retentionDays = Number(SENT_QUEUE_RETENTION_DAYS || 0);
  if (retentionDays <= 0) return null;
  return new Date(now.getTime() - retentionDays * 24 * 60 * 60 * 1000);
}

function getQueueRowTimestamp(queueRow) {
  return parseQueueTimestamp(queueRow.sentAt)
    || parseQueueTimestamp(queueRow.lastAttemptAt)
    || parseQueueTimestamp(queueRow.queuedAt);
}

function parseQueueTimestamp(value) {
  if (!value) return null;

  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : value;
  }

  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function deleteSheetRows(sheet, sheetRows) {
  if (!sheetRows.length) return;

  const sortedRows = sheetRows.slice().sort((a, b) => b - a);
  let topRow = sortedRows[0];
  let rowCount = 1;

  for (let i = 1; i < sortedRows.length; i++) {
    const sheetRow = sortedRows[i];
    if (sheetRow === topRow - rowCount) {
      rowCount++;
      continue;
    }

    sheet.deleteRows(topRow - rowCount + 1, rowCount);
    topRow = sheetRow;
    rowCount = 1;
  }

  sheet.deleteRows(topRow - rowCount + 1, rowCount);
}

function applySentQueueRowsToState(central, queueRows, updatedAt) {
  const rowsBySource = {};

  queueRows.forEach(queueRow => {
    if (!queueRow.source || !queueRow.ruId || !queueRow.availFrom) return;
    if (!rowsBySource[queueRow.source]) rowsBySource[queueRow.source] = [];
    rowsBySource[queueRow.source].push(queueRow);
  });

  let updated = 0;

  Object.keys(rowsBySource).forEach(sourceKey => {
    const cfg = getSourceConfig(sourceKey);
    if (!cfg) {
      Logger.log('applySentQueueRowsToState: source config not found for ' + sourceKey);
      return;
    }

    const stateSheet = getOrCreateTab(central, cfg.stateTab);
    ensureSheetHeader(stateSheet, STATE_HEADERS);
    const stateMap = readStateMap(stateSheet, Session.getScriptTimeZone());
    let changed = false;

    rowsBySource[sourceKey].forEach(queueRow => {
      if (stateMap[queueRow.ruId] !== queueRow.availFrom) {
        stateMap[queueRow.ruId] = queueRow.availFrom;
        changed = true;
        updated++;
      }
    });

    if (changed) {
      writeStateMap(stateSheet, stateMap, updatedAt);
    }
  });

  if (updated) {
    Logger.log('applySentQueueRowsToState: updated ' + updated + ' state entries');
  }

  return updated;
}

function readStateMap(stateSheet, timezone) {
  const range = stateSheet.getDataRange();
  const data = range.getValues();
  const displayData = range.getDisplayValues();
  const stateMap = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const displayRow = displayData[i];
    const ruId = normalizeRuId(displayRow[0] || row[0]);
    const availFrom = normalizeDate(displayRow[1] || row[1], timezone);
    if (!ruId || !availFrom) continue;
    stateMap[ruId] = availFrom;
  }

  return stateMap;
}

function writeStateMap(stateSheet, stateMap, updatedAt) {
  const ids = Object.keys(stateMap).sort();
  const rows = ids.map(id => [id, stateMap[id], updatedAt]);

  stateSheet.clearContents();
  stateSheet.getRange(1, 1, Math.max(rows.length + 1, 1), 2).setNumberFormat('@');
  stateSheet.getRange(1, 3, Math.max(rows.length + 1, 1), 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  stateSheet.getRange(1, 1, 1, STATE_HEADERS.length).setValues([STATE_HEADERS]);
  if (rows.length) {
    stateSheet.getRange(2, 1, rows.length, STATE_HEADERS.length).setValues(rows);
  }
}

function appendTelemetryRows(central, rows) {
  if (!rows.length) return;

  const sheet = getOrCreateTab(central, TELEMETRY_TAB);
  ensureSheetHeader(sheet, TELEMETRY_HEADERS);
  const start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, rows.length, TELEMETRY_HEADERS.length).setValues(rows);
}

function validateSourceConfig(cfg) {
  const required = ['key', 'sheetID', 'tab', 'agentEmail', 'colRuId', 'colAvail', 'stateTab'];
  required.forEach(field => {
    if (!cfg[field]) {
      throw new Error('Invalid source config: missing ' + field);
    }
  });
}

function toBulkRuIdValue(ruId) {
  const value = String(ruId || '').trim();
  if (/^\d+$/.test(value)) {
    const numeric = Number(value);
    if (Number.isSafeInteger(numeric)) return numeric;
  }
  return value;
}

// Kept for manual compatibility. It no longer deletes pending work.
function clearExportTab() {
  ensureExportQueue();
  Logger.log('clearExportTab: queue ensured; no data cleared');
}
