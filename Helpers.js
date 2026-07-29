function normalizeDate(value, timezone) {
  if (value === null || value === undefined || value === '' || value === '-') return null;

  try {
    const tz = timezone || Session.getScriptTimeZone() || 'UTC';

    if (value instanceof Date) {
      if (isNaN(value.getTime())) return null;
      return buildDateString(
        Number(Utilities.formatDate(value, tz, 'yyyy')),
        Number(Utilities.formatDate(value, tz, 'MM')),
        Number(Utilities.formatDate(value, tz, 'dd'))
      );
    }

    if (typeof value === 'number') {
      if (value <= 0) return null;
      const ms = Math.round((value - 25569) * 86400 * 1000);
      const date = new Date(ms);
      return buildDateString(date.getUTCFullYear(), date.getUTCMonth() + 1, date.getUTCDate());
    }

    const str = String(value).trim();
    if (!str) return null;

    const invalid = [
      'LIBERA',
      'LIBRE',
      'CHIUSURA',
      'N/A',
      '#N/A',
      '#REF!',
      '#VALUE!',
      'TBD',
      'TBC',
      'PENDING',
      '-',
      '?'
    ];
    if (invalid.indexOf(str.toUpperCase()) !== -1) return null;

    let match = str.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
    if (match) {
      return buildDateString(Number(match[1]), Number(match[2]), Number(match[3]));
    }

    match = str.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
    if (match) {
      const first = Number(match[1]);
      const second = Number(match[2]);
      const year = Number(match[3]);

      if (first > 12) return buildDateString(year, second, first);
      if (second > 12) return buildDateString(year, first, second);

      // Ambiguous numeric dates default to European day/month format.
      return buildDateString(year, second, first);
    }

    const parsed = new Date(str);
    if (isNaN(parsed.getTime())) {
      Logger.log('normalizeDate: unrecognized format value=' + str);
      return null;
    }
    return buildDateString(
      Number(Utilities.formatDate(parsed, tz, 'yyyy')),
      Number(Utilities.formatDate(parsed, tz, 'MM')),
      Number(Utilities.formatDate(parsed, tz, 'dd'))
    );

  } catch (e) {
    Logger.log('normalizeDate ERROR: value=' + value + ' | type=' + typeof value + ' | ' + formatError(e));
    return null;
  }
}

function buildDateString(year, month, day) {
  if (!year || !month || !day) return null;
  if (year < 2000 || year > 2035) return null;

  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year) return null;
  if (date.getUTCMonth() !== month - 1) return null;
  if (date.getUTCDate() !== day) return null;

  return Utilities.formatDate(date, 'UTC', 'yyyy-MM-dd');
}

function normalizeRuId(value) {
  if (value === null || value === undefined) return '';

  if (typeof value === 'number') {
    if (!isFinite(value) || Math.floor(value) !== value) return '';
    return String(value);
  }

  const ruId = String(value).trim();
  if (!/^\d+$/.test(ruId)) return '';

  return ruId;
}

function normalizeHeaderName(value) {
  return String(value || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function findHeaderIndex(headerRow, expectedName) {
  const expected = normalizeHeaderName(expectedName);
  for (let i = 0; i < headerRow.length; i++) {
    if (normalizeHeaderName(headerRow[i]) === expected) return i;
  }
  return -1;
}

function getOrCreateTab(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function getSheetOrThrow(ss, name, context) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error((context ? context + ' ' : '') + 'Tab "' + name + '" not found');
  }
  return sheet;
}

function ensureSheetHeader(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (isBlankRow(current) || !headersMatch(current, headers)) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function headersMatch(row, expectedHeaders) {
  if (!row || row.length < expectedHeaders.length) return false;

  for (let i = 0; i < expectedHeaders.length; i++) {
    if (normalizeHeaderName(row[i]) !== normalizeHeaderName(expectedHeaders[i])) {
      return false;
    }
  }
  return true;
}

function isBlankRow(row) {
  return !row || row.every(cell => cell === '' || cell === null || cell === undefined);
}

function queueKey(source, ruId, availFrom) {
  return [source, ruId, availFrom].map(part => encodeURIComponent(String(part))).join('|');
}

function getSourceConfig(sourceKey) {
  return SOURCE.find(cfg => cfg.key === sourceKey) || null;
}

function formatError(error) {
  if (!error) return 'Unknown error';
  if (error.stack) return error.stack;
  if (error.message) return error.message;
  return String(error);
}

function truncateText(value, maxLength) {
  const text = String(value || '');
  if (text.length <= maxLength) return text;
  return text.slice(0, maxLength - 3) + '...';
}
