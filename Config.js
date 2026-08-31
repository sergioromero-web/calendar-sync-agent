const CENTRAL_ID = '15FtTM7QgVfh40b89jvHMI5DmwGpmkMUuqaJi-yFhisI';
const EXPORT_TAB = '_state_dispo';
const BULK_ID = '1BMZoA-5cXdY8cmYcxYlMcbqkvoU3yMgFmqOK6jJ1a2U';
const BULK_TAB = 'To process';

const OPS_EMAIL = 'sergioromero@spotahome.com';
const SEND_OPS_EMAIL = true;
const NOTIFY_ON_SUCCESS = false;

const LOCK_TIMEOUT_MS = 30000;
const TRIGGER_HANDLER = 'main';
const TRIGGER_EVERY_HOURS = 1;
const MAX_BULK_ROWS_PER_RUN = 500;
const PROCESSING_STALE_MINUTES = 120;
const SENT_QUEUE_RETENTION_DAYS = 30;

const QUEUE_STATUS = {
  PENDING: 'pending',
  PROCESSING: 'processing',
  SENT: 'sent',
  FAILED: 'failed'
};

const QUEUE_HEADERS = [
  'Queue Id',
  'Run Id',
  'Queued At',
  'Source',
  'Agent email',
  'Rentable Unit Id',
  'Available From (yyyy-mm-dd)',
  'Status',
  'Attempts',
  'Last Attempt At',
  'Sent At',
  'Last Error'
];

const QUEUE_COL = {
  QUEUE_ID: 0,
  RUN_ID: 1,
  QUEUED_AT: 2,
  SOURCE: 3,
  AGENT_EMAIL: 4,
  RU_ID: 5,
  AVAIL_FROM: 6,
  STATUS: 7,
  ATTEMPTS: 8,
  LAST_ATTEMPT_AT: 9,
  SENT_AT: 10,
  LAST_ERROR: 11
};

const LEGACY_EXPORT_HEADERS = [
  'Agent email',
  'Rentable Unit Id',
  'Available From (yyyy-mm-dd)',
  'Source'
];

const STATE_HEADERS = ['RU id', 'available_from', 'updatedAt'];
const TELEMETRY_TAB = '_telemetry';
const TELEMETRY_HEADERS = [
  'timestamp',
  'run_id',
  'source',
  'RU id',
  'old_date',
  'new_date',
  'event',
  'details'
];

const SOURCE = [
  {
    key:        'smartflats',
    sheetID:    '1J0mQUyA9D4gaF4qPgSIyOZ8YJTvg-MXw',
    tab:        'Sheet1',
    agentEmail: 'holliewalker@spotahome.com',
    colRuId:    'Flat reference',
    colAvail:   'next available date',
    stateTab:   '_state_smartflats'
  },
  {
    key:        'dispo_compartida',
    sheetID:    '1cC-rcyrXChnP5nPXMsvftQI1dWCt91YPGWcVUmlkv00',
    tab:        'Dispo',
    agentEmail: 'sofialaukamp@spotahome.com',
    colRuId:    'RU ID',
    colAvail:   'available from',
    stateTab:   '_state_dispo_compartida'
  },
  {
    key:        'roomie_units',
    sheetID:    '1yG0Rt3673gvhKo_DZDmc2xk2akXAyBlU1TwQzjQZz78',
    tab:        'Main table',
    agentEmail: 'luca@spotahome.com',
    colRuId:    'spotahome RU id',
    colAvail:   'Available date',
    stateTab:   '_state_roomie'
  },
  {
    key:        'room_house',
    sheetID:    '1PI1sveQf9pZY_fjLKYDgnOv1PgUpBNweKYsvSDUyu9M',
    tab:        'Icals',
    agentEmail: 'marfernandez@spotahome.com',
    colRuId:    'RU_ID',
    colAvail:   'available_from',
    stateTab:   '_state_room_house'
  },
  {
    key:        'file_doomod',
    sheetID:    '1BNTa3bqXEL7v7dyvgMeEHHZlN8M-EdW0KUTPrug_0OQ',
    tab:        'Affitti semestrali-annuali',
    agentEmail: 'simonenepa@spotahome.com',
    colRuId:    'RU',
    colAvail:   'Libero da',
    stateTab:   '_state_file_doomod'
  },
  {
    key:        'hello_guest',
    sheetID:    '1nrdtfTQ3sFSeEE7lA_TlKmtrDYLh96KoK9mdMnYCMn0',
    tab:        'Availability update',
    agentEmail: 'holliewalker@spotahome.com',
    colRuId:    'sah ID',
    colAvail:   'AVAILABLE FROM',
    stateTab:   '_state_hello_guest'
  },
  {
    key:        'ngb_living',
    sheetID:    '1wk_hf-LjdGfXogwjlr97cvhHjI7DZF1SVvWvpmOWbAw',
    tab:        'Updated Info',
    agentEmail: 'marilenamaierholzner@spotahome.com',
    colRuId:    'Spotahome ID',
    colAvail:   'Available from',
    stateTab:   '_state_ngb_living'
  },
  {
    key:        'disponibilidad_pl',
    sheetID:    '1u8rqDq5UxhtSmSjkkhvhh7sBn9iAVREEJGBCz0rtHrQ',
    tab:        'Granada 2026',
    agentEmail: 'karlaaraque@spotahome.com',
    colRuId:    'SPOTAHOME ID',
    colAvail:   'DISPONIBILIDAD',
    stateTab:   '_state_disponibilidad_pl'
  },
  {
  key: 'habyt_france',
  sheetID: '11omVgpwyzc1nWdsHKreUVzSnLpGeQ1kfXZ7pmP8I814',
  tab: 'Feuille 3',
  agentEmail: 'luca@spotahome.com',
  colRuId: 'ru_id',
  colAvail: 'Status',
  stateTab: '_state_habyt_france'
},// more sources from here
];
