function auditDispoCompartida() {
  return auditSourceCounts('dispo_compartida');
}

function auditAllSources() {
  return auditSources(getActiveSources());
}

function auditAllSourceConfigs() {
  return auditSources(SOURCE);
}

function auditSources(configs) {
  const central = SpreadsheetApp.openById(CENTRAL_ID);
  const selectedConfigs = normalizeAuditConfigInputs(configs);
  const results = [];

  selectedConfigs.forEach(sourceOrConfig => {
    let cfg = null;
    try {
      cfg = resolveAuditConfig(sourceOrConfig);
      results.push(Object.assign({ success: true, error: '' }, auditSourceConfig(cfg, central)));
    } catch (e) {
      const result = buildAuditErrorResult(cfg || sourceOrConfig, e);
      results.push(result);
      Logger.log('[' + result.source + '] audit ERROR: ' + result.error);
    }
  });

  writeAuditSummarySheet(central, results);
  Logger.log(JSON.stringify(results, null, 2));
  return results;
}

function auditSourceCounts(sourceOrConfig) {
  const cfg = resolveAuditConfig(sourceOrConfig);
  const central = SpreadsheetApp.openById(CENTRAL_ID);
  return auditSourceConfig(cfg, central);
}

function auditSourceConfig(cfg, central) {
  validateAuditConfig(cfg);

  const landlord = SpreadsheetApp.openById(cfg.sheetID);
  const sheet = getSheetOrThrow(landlord, cfg.tab, '[' + cfg.key + ']');
  const timezone = landlord.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'UTC';
  const range = sheet.getDataRange();
  const values = range.getValues();
  const displayValues = range.getDisplayValues();

  if (!displayValues.length || isBlankRow(displayValues[0])) {
    throw new Error('[' + cfg.key + '] Source sheet has no header row');
  }

  const header = displayValues[0];
  const ruIdCol = findHeaderIndex(header, cfg.colRuId);
  const availFromCol = findHeaderIndex(header, cfg.colAvail);

  if (ruIdCol === -1) {
    throw new Error('[' + cfg.key + '] Column not found: "' + cfg.colRuId + '"');
  }
  if (availFromCol === -1) {
    throw new Error('[' + cfg.key + '] Column not found: "' + cfg.colAvail + '"');
  }

  const sourceMap = {};
  const seenDatesByRu = {};
  const duplicateRows = [];
  const invalidRows = [];

  let validRows = 0;
  let invalidRuIdRows = 0;
  let invalidDateRows = 0;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const displayRow = displayValues[i];
    const ruId = normalizeRuId(displayRow[ruIdCol] || row[ruIdCol]);
    const availFrom = normalizeDate(displayRow[availFromCol] || row[availFromCol], timezone);

    if (!ruId) {
      invalidRuIdRows++;
      invalidRows.push([i + 1, displayRow[ruIdCol] || row[ruIdCol], displayRow[availFromCol], 'invalid RU id']);
      continue;
    }

    if (!availFrom) {
      invalidDateRows++;
      invalidRows.push([i + 1, ruId, displayRow[availFromCol], 'invalid available from']);
      continue;
    }

    validRows++;

    if (!seenDatesByRu[ruId]) {
      seenDatesByRu[ruId] = {};
    }
    if (sourceMap[ruId]) {
      duplicateRows.push([i + 1, ruId, sourceMap[ruId], availFrom]);
    }

    seenDatesByRu[ruId][availFrom] = true;
    sourceMap[ruId] = availFrom;
  }

  const stateSheet = getOrCreateTab(central, cfg.stateTab);
  ensureSheetHeader(stateSheet, STATE_HEADERS);
  const stateMap = readStateMap(stateSheet, timezone);

  const sourceRuIds = Object.keys(sourceMap).sort();
  const stateRuIds = Object.keys(stateMap).sort();
  const sourceIndex = sourceRuIds.reduce((acc, ruId) => {
    acc[ruId] = true;
    return acc;
  }, {});
  const stateIndex = stateRuIds.reduce((acc, ruId) => {
    acc[ruId] = true;
    return acc;
  }, {});

  const missingInState = sourceRuIds.filter(ruId => !stateIndex[ruId]);
  const extraInState = stateRuIds.filter(ruId => !sourceIndex[ruId]);
  const mismatches = sourceRuIds
    .filter(ruId => stateIndex[ruId] && stateMap[ruId] !== sourceMap[ruId])
    .map(ruId => [ruId, sourceMap[ruId], stateMap[ruId]]);
  const conflictingRus = Object.keys(seenDatesByRu)
    .filter(ruId => Object.keys(seenDatesByRu[ruId]).length > 1)
    .sort()
    .map(ruId => [ruId, Object.keys(seenDatesByRu[ruId]).sort().join(', ')]);

  const result = {
    source: cfg.key,
    success: true,
    error: '',
    active: cfg.active === false ? false : true,
    sheetId: cfg.sheetID,
    tab: cfg.tab,
    stateTab: cfg.stateTab,
    agentEmail: cfg.agentEmail || '',
    ruIdHeader: cfg.colRuId,
    availableFromHeader: cfg.colAvail,
    ruIdColumn: ruIdCol + 1,
    availableFromColumn: availFromCol + 1,
    scannedRowsExcludingHeader: Math.max(values.length - 1, 0),
    validRows: validRows,
    uniqueValidRus: sourceRuIds.length,
    duplicateValidRows: validRows - sourceRuIds.length,
    conflictingRus: conflictingRus.length,
    invalidRuIdRows: invalidRuIdRows,
    invalidDateRows: invalidDateRows,
    stateRows: stateRuIds.length,
    missingInState: missingInState.length,
    extraInState: extraInState.length,
    dateMismatches: mismatches.length
  };

  writeAuditSheet(central, cfg.key, result, duplicateRows, conflictingRus, missingInState, extraInState, mismatches, invalidRows);
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function normalizeAuditConfigInputs(configs) {
  if (!configs) return getActiveSources();

  if (Array.isArray(configs)) {
    return configs;
  }

  if (isAuditConfigObject(configs)) {
    return [configs];
  }

  if (isAuditConfigMap(configs)) {
    return Object.keys(configs).map(key => configs[key]);
  }

  return [configs];
}

function isAuditConfigObject(value) {
  if (!value || typeof value !== 'object') return false;
  return Object.prototype.hasOwnProperty.call(value, 'key')
    || Object.prototype.hasOwnProperty.call(value, 'sheetID')
    || Object.prototype.hasOwnProperty.call(value, 'tab')
    || Object.prototype.hasOwnProperty.call(value, 'colRuId')
    || Object.prototype.hasOwnProperty.call(value, 'colAvail')
    || Object.prototype.hasOwnProperty.call(value, 'stateTab');
}

function isAuditConfigMap(value) {
  if (!value || typeof value !== 'object') return false;

  const keys = Object.keys(value);
  if (!keys.length) return false;

  return keys.every(key => {
    const entry = value[key];
    return typeof entry === 'string' || isAuditConfigObject(entry);
  });
}

function resolveAuditConfig(sourceOrConfig) {
  if (typeof sourceOrConfig === 'string') {
    const cfg = getSourceConfig(sourceOrConfig);
    if (!cfg) {
      throw new Error('Source not found: ' + sourceOrConfig);
    }
    return cfg;
  }

  if (sourceOrConfig && typeof sourceOrConfig === 'object') {
    validateAuditConfig(sourceOrConfig);
    return sourceOrConfig;
  }

  throw new Error('Invalid audit source config: expected source key or config object');
}

function validateAuditConfig(cfg) {
  const required = ['key', 'sheetID', 'tab', 'colRuId', 'colAvail', 'stateTab'];
  required.forEach(field => {
    if (!cfg[field]) {
      throw new Error('Invalid audit source config: missing ' + field);
    }
  });
}

function buildAuditErrorResult(sourceOrConfig, error) {
  const cfg = sourceOrConfig && typeof sourceOrConfig === 'object' ? sourceOrConfig : null;
  const source = typeof sourceOrConfig === 'string'
    ? sourceOrConfig
    : cfg && cfg.key ? cfg.key : 'unknown';

  return {
    source: source,
    success: false,
    error: formatError(error),
    active: cfg && cfg.active === false ? false : true,
    sheetId: cfg && cfg.sheetID ? cfg.sheetID : '',
    tab: cfg && cfg.tab ? cfg.tab : '',
    stateTab: cfg && cfg.stateTab ? cfg.stateTab : '',
    agentEmail: cfg && cfg.agentEmail ? cfg.agentEmail : '',
    ruIdHeader: cfg && cfg.colRuId ? cfg.colRuId : '',
    availableFromHeader: cfg && cfg.colAvail ? cfg.colAvail : '',
    ruIdColumn: '',
    availableFromColumn: '',
    scannedRowsExcludingHeader: '',
    validRows: '',
    uniqueValidRus: '',
    duplicateValidRows: '',
    conflictingRus: '',
    invalidRuIdRows: '',
    invalidDateRows: '',
    stateRows: '',
    missingInState: '',
    extraInState: '',
    dateMismatches: ''
  };
}

function writeAuditSummarySheet(central, results) {
  const sheet = getOrCreateTab(central, '_audit_summary');
  sheet.clearContents();

  const headers = [
    'source',
    'success',
    'active',
    'sheetId',
    'tab',
    'stateTab',
    'agentEmail',
    'ruIdHeader',
    'availableFromHeader',
    'ruIdColumn',
    'availableFromColumn',
    'scannedRowsExcludingHeader',
    'validRows',
    'uniqueValidRus',
    'duplicateValidRows',
    'conflictingRus',
    'invalidRuIdRows',
    'invalidDateRows',
    'stateRows',
    'missingInState',
    'extraInState',
    'dateMismatches',
    'error'
  ];

  const rows = results.map(result => headers.map(header => result[header]));
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

function writeAuditSheet(central, sourceKey, result, duplicateRows, conflictingRus, missingInState, extraInState, mismatches, invalidRows) {
  const sheet = getOrCreateTab(central, '_audit_' + sourceKey);
  sheet.clearContents();

  const summaryRows = Object.keys(result).map(key => [key, result[key]]);
  sheet.getRange(1, 1, 1, 2).setValues([['metric', 'value']]);
  if (summaryRows.length) {
    sheet.getRange(2, 1, summaryRows.length, 2).setValues(summaryRows);
  }

  let row = summaryRows.length + 4;
  row = writeAuditSection(sheet, row, 'duplicate valid rows', ['sheet_row', 'RU id', 'previous_date', 'date_at_row'], duplicateRows);
  row = writeAuditSection(sheet, row, 'conflicting RUs', ['RU id', 'dates_seen'], conflictingRus);
  row = writeAuditSection(sheet, row, 'missing in state', ['RU id'], missingInState.map(ruId => [ruId]));
  row = writeAuditSection(sheet, row, 'extra in state', ['RU id'], extraInState.map(ruId => [ruId]));
  row = writeAuditSection(sheet, row, 'date mismatches', ['RU id', 'source_date', 'state_date'], mismatches);
  writeAuditSection(sheet, row, 'invalid rows', ['sheet_row', 'RU id', 'available_from', 'reason'], invalidRows);
}

function writeAuditSection(sheet, startRow, title, headers, rows) {
  sheet.getRange(startRow, 1, 1, 1).setValue(title);
  sheet.getRange(startRow + 1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    sheet.getRange(startRow + 2, 1, rows.length, headers.length).setValues(rows);
  }
  return startRow + rows.length + 4;
}
