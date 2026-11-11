/* CodeSim Lab — copia automática. Este archivo se ejecuta en Google Apps Script.
 * SHARED_SECRET y SHEET_ID se configuran en Propiedades del script, nunca en el HTML.
 */
var CODESIM_ACTIVITY_HEADERS = [
  'Fecha del dispositivo', 'Fecha de recepción', 'Nombre', 'Correo', 'Grupo',
  'Actividad', 'Misión', 'Intento', 'Resultado', 'Puntos', 'Total',
  'ID evaluación', 'ID registro', 'ID evento', 'Programa', 'Datos de sincronización'
];
var CODESIM_ASSESSMENT_HEADERS = [
  'Nombre', 'Correo', 'Grupo', 'Estado', 'Misión actual', 'Intentos usados',
  'Misiones resueltas', 'Puntos', 'Total', 'Porcentaje', 'Inicio',
  'Última actividad', 'Entrega', 'Última recepción', 'ID evaluación', 'ID registro'
];
var CODESIM_EVENT_TYPES = [
  'study_opened', 'case_opened', 'program_run', 'case_completed',
  'exam_start', 'exam_run', 'exam_result', 'exam_abandoned'
];
var CODESIM_HTML_STATUS = 'Entregada · HTML (sin verificación)';
var CODESIM_PAGES_SOURCE = 'github_pages';
var CODESIM_PAGES_STATUS = ' · Navegador (sin verificación)';

function doGet() {
  return codesimJson_({ok: true, service: 'CodeSim Lab — recepción de registros', offlineResult: true, offlineActivity: true});
}

function doPost(e) {
  var accepted = [];
  var lock;
  try {
    var properties = PropertiesService.getScriptProperties();
    var secret = properties.getProperty('SHARED_SECRET');
    var sheetId = properties.getProperty('SHEET_ID');
    var raw = e && e.postData && e.postData.contents;
    if (typeof raw !== 'string' || raw.length > 900000) return codesimJson_({ok: false, error: 'INVALID_BATCH', accepted: []});
    var payload;
    try { payload = JSON.parse(raw); } catch (_) { return codesimJson_({ok: false, error: 'INVALID_BATCH', accepted: []}); }
    if (payload && payload.kind === 'codesim-offline-result-v1') {
      return codesimReceiveOffline_(properties, sheetId, raw, payload.row);
    }
    if (payload && payload.kind === 'codesim-activity-v1') {
      return codesimReceivePages_(properties, sheetId, raw, payload.rows);
    }
    if (!secret || secret.length < 32 || !sheetId) return codesimJson_({ok: false, error: 'NOT_CONFIGURED', accepted: []});
    if (!payload || !codesimSecretEquals_(secret, payload.secret)) return codesimJson_({ok: false, error: 'AUTH_REQUIRED', accepted: []});
    if (!Array.isArray(payload.rows) || !payload.rows.length || payload.rows.length > 20) return codesimJson_({ok: false, error: 'INVALID_BATCH', accepted: []});
    var rows = payload.rows.map(function (value) { return codesimCleanEvent_(value); });
    var ids = Object.create(null);
    rows.forEach(function (row) {
      var canonical = JSON.stringify(row);
      if (ids[row.id] && ids[row.id] !== canonical) throw new Error('EVENT_CONFLICT');
      ids[row.id] = canonical;
    });
    lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return codesimJson_({ok: false, error: 'BUSY', accepted: []});
    var book = SpreadsheetApp.openById(sheetId);
    var activity = codesimSheet_(book, 'Actividad', CODESIM_ACTIVITY_HEADERS);
    var assessments = codesimSheet_(book, 'Evaluaciones', CODESIM_ASSESSMENT_HEADERS);
    var known = Object.create(null);
    if (activity.getLastRow() > 1) {
      activity.getRange(2, 14, activity.getLastRow() - 1, 1).getValues().forEach(function (entry, i) {
        if (entry[0]) known[String(entry[0])] = i + 2;
      });
    }
    // Reject a reused ID with different data before adding any rows from this batch.
    rows.forEach(function (row) {
      if (known[row.id] && activity.getRange(known[row.id], 16).getValue() !== JSON.stringify(row)) throw new Error('EVENT_CONFLICT');
    });
    var affected = Object.create(null);
    rows.forEach(function (row) {
      if (!known[row.id]) {
        var position = activity.getLastRow() + 1;
        activity.getRange(position, 1, 1, 16).setValues([codesimActivityValues_(row)]).setFontFamily('Times New Roman').setFontSize(12);
        activity.getRange(position, 1, 1, 2).setNumberFormat('yyyy-mm-dd hh:mm:ss');
        known[row.id] = position;
      }
      if (row.assessmentId) affected[row.assessmentId] = true;
    });
    // Rebuild affected summaries from the durable activity log. This also repairs
    // a previous attempt that appended the event but failed before the summary.
    Object.keys(affected).forEach(function (assessmentId) {
      codesimRebuildAssessment_(activity, assessments, assessmentId);
    });
    SpreadsheetApp.flush();
    accepted = Object.keys(ids);
    return codesimJson_({ok: true, accepted: accepted});
  } catch (error) {
    var knownErrors = ['INVALID_BATCH', 'INVALID_EVENT', 'INVALID_RESULT', 'EVENT_CONFLICT', 'HEADERS_CHANGED', 'NOT_CONFIGURED', 'BUSY', 'RETRY_LATER'];
    return codesimJson_({ok: false, error: knownErrors.indexOf(error.message) >= 0 ? error.message : 'COPY_FAILED', accepted: []});
  } finally {
    if (lock && lock.hasLock()) lock.releaseLock();
  }
}

// GitHub Pages can send a bounded batch without exposing the server secret.
// The receiver never returns names, grades, or anything already in the sheet.
function codesimReceivePages_(properties, sheetId, raw, values) {
  if (!sheetId) throw new Error('NOT_CONFIGURED');
  if (raw.length > 180000 || !Array.isArray(values) || !values.length || values.length > 15) {
    throw new Error('INVALID_BATCH');
  }
  var now = Date.now();
  var rows = Object.create(null);
  values.forEach(function (value) {
    var row = codesimCleanPagesEvent_(value, now);
    if (rows[row.id] && JSON.stringify(rows[row.id]) !== JSON.stringify(row)) throw new Error('EVENT_CONFLICT');
    rows[row.id] = row;
  });
  var ids = Object.keys(rows);
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) throw new Error('BUSY');
  try {
    codesimLimitIntake_(properties, 'CODESIM_PAGES_RATE', now, ids.length, 600, 15000);
    var book = SpreadsheetApp.openById(sheetId);
    var activity = codesimSheet_(book, 'Actividad', CODESIM_ACTIVITY_HEADERS);
    var assessments = codesimSheet_(book, 'Evaluaciones', CODESIM_ASSESSMENT_HEADERS);
    var added = [];
    var affected = Object.create(null);
    // Verify all conflicts before appending the batch. Receipt time is assigned
    // by the server and remains unchanged on a retry.
    ids.forEach(function (id) {
      var row = rows[id];
      var oldPosition = codesimFindResult_(activity, 14, id);
      if (oldPosition) {
        var old = JSON.parse(activity.getRange(oldPosition, 16).getValue());
        row.receivedAt = old.receivedAt;
        if (JSON.stringify(old) !== JSON.stringify(row)) throw new Error('EVENT_CONFLICT');
      } else {
        added.push(codesimActivityValues_(row));
      }
      if (!row.assessmentId) return;
      if (!affected[row.assessmentId]) {
        var prior = codesimAssessmentEvents_(activity, row.assessmentId);
        var summary = codesimFindResult_(assessments, 15, row.assessmentId);
        if (summary && (!prior.length ||
          !String(assessments.getRange(summary, 4).getValue()).endsWith(CODESIM_PAGES_STATUS))) {
          throw new Error('EVENT_CONFLICT');
        }
        affected[row.assessmentId] = {events: prior, summary: summary};
      }
      var events = affected[row.assessmentId].events;
      if (events.some(function (event) {
        return event.source !== CODESIM_PAGES_SOURCE || event.registrationId !== row.registrationId || event.email !== row.email;
      })) throw new Error('EVENT_CONFLICT');
      if (!oldPosition) events.push(row);
    });
    if (added.length) {
      var start = codesimAppendRow_(activity, added.length);
      activity.getRange(start, 1, added.length, 16).setValues(added).setFontFamily('Times New Roman').setFontSize(12);
      activity.getRange(start, 1, added.length, 2).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    }
    Object.keys(affected).forEach(function (id) {
      var entry = affected[id];
      codesimWriteAssessment_(assessments, id, entry.events, entry.summary);
    });
    SpreadsheetApp.flush();
    return codesimJson_({ok: true, accepted: ids.map(function (id) { return id.slice(6); })});
  } finally {
    lock.releaseLock();
  }
}

function codesimCleanPagesEvent_(value, now) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('INVALID_EVENT');
  var idPattern = /^[A-Za-z0-9_-]{1,120}$/;
  var name = typeof value.name === 'string' && value.name.trim();
  var email = typeof value.email === 'string' && value.email.trim().toLowerCase();
  if (typeof value.id !== 'string' || !idPattern.test(value.id) ||
      typeof value.registrationId !== 'string' || !idPattern.test(value.registrationId) ||
      (value.assessmentId && (typeof value.assessmentId !== 'string' || !idPattern.test(value.assessmentId))) ||
      (value.caseId && (typeof value.caseId !== 'string' || !idPattern.test(value.caseId))) ||
      !name || name.length > 120 || !email || email.length > 254 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ||
      !Number.isSafeInteger(value.createdAt) || value.createdAt < 1577836800000 || value.createdAt > now + 86400000 ||
      (value.total != null && value.total !== 25) ||
      (value.feedback != null && (typeof value.feedback !== 'string' || value.feedback.length > 1000)) ||
      (value.program != null && (!Array.isArray(value.program) || JSON.stringify(value.program).length > 65000))) {
    throw new Error('INVALID_EVENT');
  }
  var row = codesimCleanEvent_({
    id: 'pages:' + value.id, type: value.type, name: name, email: email,
    registrationId: 'pages:' + value.registrationId,
    assessmentId: value.assessmentId ? 'pages:' + value.assessmentId : '',
    caseId: value.caseId || '', attemptNumber: value.attemptNumber, passed: value.passed,
    score: value.score, total: value.total, createdAt: value.createdAt, receivedAt: now, program: value.program
  }, true);
  row.source = CODESIM_PAGES_SOURCE;
  if (value.feedback) row.feedback = value.feedback;
  return row;
}

// The portable HTML contains no server secret. Accept only a completed score,
// keep it distinct from authenticated Site data, and return no student data.
function codesimReceiveOffline_(properties, sheetId, raw, value) {
  if (!sheetId) throw new Error('NOT_CONFIGURED');
  if (raw.length > 4096) throw new Error('INVALID_RESULT');
  var now = Date.now();
  var row = codesimCleanOfflineResult_(value, now);
  var sheetKey = 'html:' + row.id;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) throw new Error('BUSY');
  try {
    codesimLimitOffline_(properties, now);
    var book = SpreadsheetApp.openById(sheetId);
    var activity = codesimSheet_(book, 'Actividad', CODESIM_ACTIVITY_HEADERS);
    var assessments = codesimSheet_(book, 'Evaluaciones', CODESIM_ASSESSMENT_HEADERS);
    var activityRow = codesimFindResult_(activity, 14, sheetKey);
    var summaryRow = codesimFindResult_(assessments, 15, sheetKey);
    var receivedAt = now;
    if (activityRow) {
      var stored = JSON.parse(activity.getRange(activityRow, 16).getValue());
      if (stored.type !== 'exam_result_html' || JSON.stringify(stored.offlineResult) !== JSON.stringify(row)) {
        throw new Error('EVENT_CONFLICT');
      }
      receivedAt = stored.receivedAt;
    }
    if (summaryRow && (!activityRow || assessments.getRange(summaryRow, 4).getValue() !== CODESIM_HTML_STATUS ||
        assessments.getRange(summaryRow, 16).getValue() !== sheetKey)) {
      throw new Error('EVENT_CONFLICT');
    }
    if (!activityRow) {
      var event = {
        id: sheetKey, type: 'exam_result_html', group: '', name: row.name, email: row.email,
        registrationId: sheetKey, assessmentId: sheetKey, caseId: '', attemptNumber: null, passed: null,
        score: row.score, total: row.total, createdAt: row.submittedAt, receivedAt: receivedAt,
        program: '', offlineResult: row
      };
      activityRow = codesimAppendRow_(activity);
      activity.getRange(activityRow, 1, 1, 16).setValues([codesimActivityValues_(event)]);
      activity.getRange(activityRow, 1, 1, 2).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    }
    if (!summaryRow) summaryRow = codesimAppendRow_(assessments);
    var output = [codesimCell_(row.name), codesimCell_(row.email), '', CODESIM_HTML_STATUS, '',
      row.attemptsUsed, row.score, row.score, row.total, row.score / row.total,
      new Date(row.startedAt), new Date(row.submittedAt), new Date(row.submittedAt), new Date(receivedAt), sheetKey, sheetKey];
    assessments.getRange(summaryRow, 1, 1, 16).setValues([output]);
    assessments.getRange(summaryRow, 10).setNumberFormat('0.0%');
    assessments.getRange(summaryRow, 11, 1, 4).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    SpreadsheetApp.flush();
    return codesimJson_({ok: true, accepted: [row.id]});
  } finally {
    lock.releaseLock();
  }
}

function codesimCleanOfflineResult_(value, now) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('INVALID_RESULT');
  var id = value.id;
  var name = typeof value.name === 'string' && value.name.trim();
  var email = typeof value.email === 'string' && value.email.trim().toLowerCase();
  var attemptsUsed = value.attemptsUsed === undefined ? value.score : value.attemptsUsed;
  if (typeof id !== 'string' || !/^[A-Za-z0-9_-]{1,120}$/.test(id) || !name || name.length > 120 ||
      !email || email.length > 254 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ||
      value.total !== 25 || !Number.isSafeInteger(value.score) || value.score < 0 || value.score > 25 ||
      !Number.isSafeInteger(attemptsUsed) || attemptsUsed < value.score || attemptsUsed > 75 ||
      !Number.isSafeInteger(value.startedAt) || !Number.isSafeInteger(value.submittedAt) ||
      value.startedAt < 1577836800000 || value.submittedAt > now + 24 * 60 * 60 * 1000 ||
      value.submittedAt < value.startedAt || value.submittedAt - value.startedAt > 90 * 60 * 1000) {
    throw new Error('INVALID_RESULT');
  }
  return {id: id, name: name, email: email, score: value.score, total: value.total,
    startedAt: value.startedAt, submittedAt: value.submittedAt, attemptsUsed: attemptsUsed};
}

function codesimLimitOffline_(properties, now) {
  // Requests share one limit. Apps Script does not expose the remote IP here.
  codesimLimitIntake_(properties, 'CODESIM_HTML_RATE', now, 1, 120, 1500);
}

function codesimLimitIntake_(properties, key, now, count, perMinute, perDay) {
  var minute = Math.floor(now / 60000);
  var prior = (properties.getProperty(key) || '').split(':');
  var priorMinute = Number(prior[0]);
  var minuteCount = priorMinute === minute ? Number(prior[1]) || 0 : 0;
  var dayCount = Math.floor(priorMinute / 1440) === Math.floor(minute / 1440) ? Number(prior[2]) || 0 : 0;
  if (minuteCount + count > perMinute || dayCount + count > perDay) throw new Error('RETRY_LATER');
  properties.setProperty(key, [minute, minuteCount + count, dayCount + count].join(':'));
}

function codesimFindResult_(sheet, column, value) {
  var count = sheet.getLastRow() - 1;
  if (count < 1) return 0;
  var matches = sheet.getRange(2, column, count, 1).createTextFinder(value).matchEntireCell(true).matchCase(true).findAll();
  if (matches.length > 1) throw new Error('EVENT_CONFLICT');
  return matches.length ? matches[0].getRow() : 0;
}

function codesimAppendRow_(sheet, count) {
  var row = sheet.getLastRow() + 1;
  var capacity = sheet.getMaxRows();
  var end = row + (count || 1) - 1;
  if (end > capacity) sheet.insertRowsAfter(capacity, Math.max(100, end - capacity));
  return row;
}

function codesimJson_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}

function codesimSecretEquals_(expected, actual) {
  if (typeof actual !== 'string' || actual.length > 1024) return false;
  var different = expected.length ^ actual.length;
  for (var i = 0; i < expected.length; i++) different |= expected.charCodeAt(i) ^ (actual.charCodeAt(i) || 0);
  return different === 0;
}

function codesimCleanEvent_(value, fromPages) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('INVALID_EVENT');
  function text(key, limit, optional) {
    var result = value[key];
    if (optional && (result === undefined || result === null)) return '';
    if (typeof result !== 'string' || result.length > limit || (!optional && !result.trim())) throw new Error('INVALID_EVENT');
    return result;
  }
  function id(key, optional) {
    var result = text(key, 160, optional);
    if (result && !/^[A-Za-z0-9:_-]+$/.test(result)) throw new Error('INVALID_EVENT');
    return result;
  }
  function number(key, min, max, optional) {
    var result = value[key];
    if (optional && (result === undefined || result === null || result === '')) return null;
    if (typeof result !== 'number' || !Number.isSafeInteger(result) || result < min || result > max) throw new Error('INVALID_EVENT');
    return result;
  }
  var row = {
    id: id('id'), type: text('type', 40), group: text('group', 120, true),
    registrationId: id('registrationId'), assessmentId: id('assessmentId', true),
    name: text('name', 120), email: text('email', 254), caseId: id('caseId', true),
    attemptNumber: number('attemptNumber', 1, 3, true), passed: value.passed == null ? null : value.passed,
    score: number('score', 0, 25, true), total: number('total', 1, 25, true),
    createdAt: number('createdAt', 1, 8640000000000000), receivedAt: number('receivedAt', 1, 8640000000000000),
    program: value.program == null ? null : value.program
  };
  if ((CODESIM_EVENT_TYPES.indexOf(row.type) < 0 && !(fromPages && row.type === 'login')) || (row.passed !== null && typeof row.passed !== 'boolean') ||
      (row.score !== null && (row.total === null || row.score > row.total)) ||
      (row.type.indexOf('exam_') === 0 && !row.assessmentId) ||
      (row.type === 'exam_run' && (!row.caseId || row.attemptNumber === null || row.passed === null)) ||
      (row.type === 'exam_result' && (row.score === null || row.total === null)) ||
      (row.program !== null && !Array.isArray(row.program))) throw new Error('INVALID_EVENT');
  var programText = row.program === null ? '' : JSON.stringify(row.program);
  if (programText.length > 350000) throw new Error('INVALID_EVENT');
  row.programDigest = programText ? Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, programText, Utilities.Charset.UTF_8)
    .map(function (byte) { return ('0' + ((byte + 256) % 256).toString(16)).slice(-2); }).join('') : '';
  var previewLimit = fromPages ? 15000 : 30000;
  row.programTruncated = programText.length > previewLimit;
  row.program = programText.slice(0, previewLimit);
  return row;
}

function codesimCell_(value) {
  var text = value == null ? '' : String(value);
  // Quoted text is kept as text by Sheets, including formulas after whitespace.
  return /^[\s\u0000-\u001f]*[=+@-]/.test(text) ? "'" + text : text;
}

function codesimActivityValues_(row) {
  var note = row.source === CODESIM_PAGES_SOURCE ? '\n[Vista abreviada del programa]' : '\n[Vista abreviada: el programa completo está en el panel docente.]';
  return [new Date(row.createdAt), new Date(row.receivedAt), codesimCell_(row.name), codesimCell_(row.email), codesimCell_(row.group),
    row.type, row.caseId, row.attemptNumber === null ? '' : row.attemptNumber,
    row.passed === null ? '' : row.passed ? 'Correcto' : 'Incorrecto', row.score === null ? '' : row.score,
    row.total === null ? '' : row.total, row.assessmentId, row.registrationId, row.id,
    row.program + (row.programTruncated ? note : ''), JSON.stringify(row)];
}

function codesimSheet_(book, name, headers) {
  var sheet = book.getSheetByName(name) || book.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var actual = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (headers.some(function (header, i) { return actual[i] !== header; })) throw new Error('HEADERS_CHANGED');
  return sheet;
}

function codesimRebuildAssessment_(activity, assessments, assessmentId) {
  var events = codesimAssessmentEvents_(activity, assessmentId);
  if (!events.length) return;
  codesimWriteAssessment_(assessments, assessmentId, events, codesimFindResult_(assessments, 15, assessmentId));
}

function codesimAssessmentEvents_(activity, assessmentId) {
  var count = activity.getLastRow() - 1;
  if (count <= 0) return [];
  var matches = activity.getRange(2, 12, count, 1).createTextFinder(assessmentId).matchEntireCell(true).matchCase(true).findAll();
  return matches.map(function (cell) { return JSON.parse(activity.getRange(cell.getRow(), 16).getValue()); });
}

function codesimWriteAssessment_(assessments, assessmentId, events, position) {
  events.sort(function (a, b) { return a.createdAt - b.createdAt || a.receivedAt - b.receivedAt || a.id.localeCompare(b.id); });
  var first = events[0], latest = events[events.length - 1];
  var terminal = events.filter(function (row) { return row.type === 'exam_result'; }).pop();
  var abandoned = events.filter(function (row) { return row.type === 'exam_abandoned'; }).pop();
  var runs = events.filter(function (row) { return row.type === 'exam_run'; });
  var passedCases = Object.create(null);
  runs.forEach(function (row) { if (row.passed) passedCases[row.caseId] = true; });
  var total = (terminal && terminal.total) || latest.total || first.total || 25;
  var score = terminal ? terminal.score : Object.keys(passedCases).length;
  var receivedAt = events.reduce(function (maximum, row) { return Math.max(maximum, row.receivedAt); }, 0);
  var started = events.filter(function (row) { return row.type === 'exam_start'; })[0];
  var status = terminal ? 'Entregada' : abandoned ? 'Reiniciada o abandonada' : 'En curso';
  if (latest.source === CODESIM_PAGES_SOURCE) status += CODESIM_PAGES_STATUS;
  var lastRun = runs[runs.length - 1];
  var output = [codesimCell_(latest.name), codesimCell_(latest.email), codesimCell_(latest.group), status,
    lastRun ? lastRun.caseId : '', runs.length, Object.keys(passedCases).length, score, total,
    score / total, new Date(started ? started.createdAt : first.createdAt), new Date(latest.createdAt),
    terminal ? new Date(terminal.createdAt) : '', new Date(receivedAt), assessmentId, latest.registrationId];
  if (!position) position = codesimAppendRow_(assessments);
  assessments.getRange(position, 1, 1, 16).setValues([output]).setFontFamily('Times New Roman').setFontSize(12);
  assessments.getRange(position, 10).setNumberFormat('0.0%');
  assessments.getRange(position, 11, 1, 4).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}

// Ejecutar una vez desde el editor, después de configurar las propiedades.
function prepararHoja() {
  var properties = PropertiesService.getScriptProperties();
  var id = properties.getProperty('SHEET_ID');
  var book = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!book) throw new Error('Configura SHEET_ID en Propiedades del script.');
  if (!id) properties.setProperty('SHEET_ID', book.getId());
  book.setSpreadsheetTimeZone('America/Bogota');
  var pairs = [['Evaluaciones', CODESIM_ASSESSMENT_HEADERS], ['Actividad', CODESIM_ACTIVITY_HEADERS]];
  pairs.forEach(function (pair) {
    var sheet = codesimSheet_(book, pair[0], pair[1]);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(pair[0] === 'Evaluaciones' ? 1 : 0);
    sheet.setHiddenGridlines(true);
    sheet.setColumnWidths(1, 16, 160);
    sheet.setColumnWidth(pair[0] === 'Evaluaciones' ? 1 : 3, 230);
    sheet.setColumnWidth(pair[0] === 'Evaluaciones' ? 2 : 4, 260);
    sheet.getRange(1, 1, 1, 16).setBackground('#2457d7').setFontColor('#ffffff').setFontFamily('Times New Roman').setFontSize(12).setFontWeight('bold').setWrap(true);
    sheet.setRowHeight(1, 44);
  });
  SpreadsheetApp.flush();
  return {ok: true, message: 'Hoja preparada. Falta comprobar una copia desde el panel docente.'};
}
