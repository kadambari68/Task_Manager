// ============================================================
// recurringtaskengine.js — TaskFlow v6  Sprint 6
//
// Recurring tasks: permanent tasks that send a reminder email
// on a configurable schedule. The task itself never closes —
// it is a standing obligation. No new row is created per cycle;
// the engine advances NextTriggerDate after each fire.
//
// Public API (called from frontend via google.script.run):
//   createRecurringTask(data)   Owner + Manager
//   updateRecurringTask(id, updates)
//   pauseRecurringTask(id)      Owner + Manager
//   resumeRecurringTask(id)     Owner + Manager
//   deleteRecurringTask(id)     Owner only
//   getRecurringTasks()         Owner sees all; Manager sees team; Member sees own
//
// Internal (called from runHourlyEngine in reminderengine):
//   runRecurringReminderEngine_()
//
// Sheet columns (0-based, A=0):
//   0  RecurringID     1  Title          2  Description
//   3  AssigneeEmail   4  AssigneeName   5  Team
//   6  Frequency       7  IntervalValue  8  IntervalUnit
//   9  NextTriggerDate 10 ReminderTime   11 StartDate
//   12 EndDate         13 CreatedBy      14 CreatedAt
//   15 Status          16 SourceTaskId
//
// Frequency constants: DAILY | WEEKLY | MONTHLY | YEARLY | CUSTOM
// Status constants:    ACTIVE | PAUSED
// ============================================================

var REC_COL = {
  ID: 0,
  TITLE: 1,
  DESCRIPTION: 2,
  ASSIGNEE_EMAIL: 3,
  ASSIGNEE_NAME: 4,
  TEAM: 5,
  FREQUENCY: 6,
  INTERVAL_VAL: 7,
  INTERVAL_UNIT: 8,
  NEXT_TRIGGER: 9,
  REM_TIME: 10,
  START_DATE: 11,
  END_DATE: 12,
  CREATED_BY: 13,
  CREATED_AT: 14,
  STATUS: 15,
  SOURCE_TASK_ID: 16
};

var REC_FREQ = { DAILY: 'DAILY', WEEKLY: 'WEEKLY', MONTHLY: 'MONTHLY', YEARLY: 'YEARLY', CUSTOM: 'CUSTOM' };
var REC_STATUS = { ACTIVE: 'ACTIVE', PAUSED: 'PAUSED' };
var VALID_FREQUENCIES = ['DAILY', 'WEEKLY', 'MONTHLY', 'YEARLY', 'CUSTOM'];
var VALID_UNITS = ['DAY', 'WEEK', 'MONTH', 'YEAR'];

// ------------------------------------------------------------------
// PRIVATE HELPERS
// ------------------------------------------------------------------

function getRecSheet_() {
  var sh = getSheet(SHEETS.RECURRING_TASKS);
  if (sh) return ensureRecurringSourceTaskColumn_(sh);

  // Sheet missing -- create it gracefully instead of throwing.
  return setupRecurringTasksSheet_Internal_();
}

/**
 * setupRecurringTasksSheet()
 * Public API to initialize the RecurringTasks sheet. (Owner/Manager only)
 */
function setupRecurringTasksSheet() {
  try {
    requireRole_(['Owner', 'Manager']);
    var sheet = setupRecurringTasksSheet_Internal_();
    if (sheet) return ok_({ message: 'RecurringTasks sheet ready.' });
    return err_('SYSTEM_ERROR');
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('setupRecurringTasksSheet: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

/**
 * Internal logic for sheet creation to be reused by getRecSheet_ and setupRecurringTasksSheet
 */
function setupRecurringTasksSheet_Internal_() {
  try {
    var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!id) return null;
    var ss = SpreadsheetApp.openById(id);
    var existing = ss.getSheetByName(SHEETS.RECURRING_TASKS);
    if (!existing) {
      existing = ss.insertSheet(SHEETS.RECURRING_TASKS);
      existing.appendRow([
        'RecurringID', 'Title', 'Description', 'AssigneeEmail',
        'AssigneeName', 'Team', 'Frequency', 'IntervalValue',
        'IntervalUnit', 'NextTriggerDate', 'ReminderTime', 'StartDate',
        'EndDate', 'CreatedBy', 'CreatedAt', 'Status', 'SourceTaskId'
      ]);
      existing.getRange(1, 1, 1, 17).setFontWeight('bold');
      existing.setFrozenRows(1);
      existing.setColumnWidth(2, 200);
      existing.setColumnWidth(10, 140);
      // Ensure it's in the cache if global cache exists
      if (typeof _sheetCache !== 'undefined') _sheetCache[SHEETS.RECURRING_TASKS] = existing;
    }
    return ensureRecurringSourceTaskColumn_(existing);
  } catch (e) {
    console.error('setupRecurringTasksSheet_Internal_: failed to create RecurringTasks sheet: ' + e.message);
    return null;
  }
}

function ensureRecurringSourceTaskColumn_(sheet) {
  if (!sheet) return sheet;
  if (sheet.getLastColumn() < REC_COL.SOURCE_TASK_ID + 1) {
    sheet.getRange(1, REC_COL.SOURCE_TASK_ID + 1).setValue('SourceTaskId');
    sheet.getRange(1, 1, 1, REC_COL.SOURCE_TASK_ID + 1).setFontWeight('bold');
  }
  return sheet;
}


function generateRecId_() {
  return 'REC-' + new Date().getFullYear() + '-' + String(Date.now()).slice(-6);
}

// ------------------------------------------------------------------
// calculateNextTriggerDate_
// Advances a date by one interval given the frequency/unit config.
// Month-end safety: MONTHLY/YEARLY cap at last day of target month.
// base  — Date object (current NextTriggerDate)
// freq  — one of VALID_FREQUENCIES
// val   — integer interval value (used when freq === CUSTOM)
// unit  — one of VALID_UNITS (used when freq === CUSTOM)
// ------------------------------------------------------------------
function calculateNextTriggerDate_(base, freq, val, unit) {
  var d = new Date(base);

  if (freq === REC_FREQ.DAILY) {
    d.setDate(d.getDate() + 1);
    return d;
  }
  if (freq === REC_FREQ.WEEKLY) {
    d.setDate(d.getDate() + 7);
    return d;
  }
  if (freq === REC_FREQ.YEARLY) {
    return advanceMonths_(d, 12);
  }
  if (freq === REC_FREQ.MONTHLY) {
    return advanceMonths_(d, 1);
  }
  if (freq === REC_FREQ.CUSTOM) {
    var n = parseInt(val, 10) || 1;
    var u = String(unit || 'DAY').toUpperCase();
    if (u === 'DAY') { d.setDate(d.getDate() + n); return d; }
    if (u === 'WEEK') { d.setDate(d.getDate() + n * 7); return d; }
    if (u === 'MONTH') { return advanceMonths_(d, n); }
    if (u === 'YEAR') { return advanceMonths_(d, n * 12); }
  }
  // Fallback: 1 month
  return advanceMonths_(d, 1);
}

// Month-end safety helper.
// Adding months naively (setMonth) overflows: Jan 31 + 1 month = Mar 3.
// This caps at the last valid day of the target month.
function advanceMonths_(base, months) {
  var d = new Date(base);
  var day = d.getDate();
  d.setDate(1);                            // set to 1st to avoid overflow during month change
  d.setMonth(d.getMonth() + months);
  var lastDay = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
  d.setDate(Math.min(day, lastDay));       // cap at last day of target month
  return d;
}

// Format a Date as YYYY-MM-DD string (date only — no time component stored)
function toDateStr_(d) {
  try {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    return '';
  }
}

// Safely normalize sheet values that may be Date objects or strings.
function toDateStrSafe_(raw) {
  var dateStr = '';
  if (raw) {
    var d = (raw instanceof Date) ? raw : new Date(raw);
    if (!isNaN(d.getTime())) {
      dateStr = toDateStr_(d);
    }
  }
  return dateStr;
}

// Find a recurring task row by ID — returns { rowIndex, data } or null
function findRecRow_(recId) {
  if (!recId) return null;
  var sheet = getRecSheet_();
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][REC_COL.ID] === recId) return { rowIndex: i + 1, data: data[i] };
  }
  return null;
}

function recurringLegacyKey_(title, assigneeEmail) {
  return String(assigneeEmail || '').toLowerCase().trim() + '|' + String(title || '').toLowerCase().trim();
}

function recurringTaskHasSourceTag_(row) {
  if (!row) return false;
  if (typeof isRecurringTaskRow_ === 'function') return isRecurringTaskRow_(row);
  return String(row[COL.TAGS] || '').split(',').map(function (t) {
    return String(t).trim();
  }).indexOf('__recurring') !== -1;
}

function buildRecurringSourceTaskMap_() {
  var out = { byId: {}, byLegacyKey: {} };
  var taskSheet = getSheet(SHEETS.TASKS);
  if (!taskSheet) return out;

  var rows = taskSheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var taskId = row[COL.TASK_ID];
    if (!taskId || !recurringTaskHasSourceTag_(row)) continue;

    var source = {
      taskId: taskId,
      title: row[COL.TASK_NAME] || '',
      assigneeEmail: row[COL.OWNER_EMAIL] || '',
      status: row[COL.STATUS] || '',
      row: row
    };
    out.byId[String(taskId)] = source;
    out.byLegacyKey[recurringLegacyKey_(source.title, source.assigneeEmail)] = source;
  }
  return out;
}

function resolveRecurringSourceTask_(row, sourceMap) {
  var sourceTaskId = row[REC_COL.SOURCE_TASK_ID] || '';
  if (sourceTaskId && sourceMap.byId[String(sourceTaskId)]) return sourceMap.byId[String(sourceTaskId)];
  return sourceMap.byLegacyKey[recurringLegacyKey_(row[REC_COL.TITLE], row[REC_COL.ASSIGNEE_EMAIL])] || null;
}

function isRecurringSourceReminderEligible_(source) {
  if (!source) return false;
  return source.status === STATUS.ON_HOLD;
}

// ------------------------------------------------------------------
// createRecurringTask
// Owner or Manager. Manager can only create for their own team.
// ------------------------------------------------------------------
function createRecurringTask(taskData) {
  try {
    var actor = requireRole_(['Owner', 'Manager']);
    var clean = sanitize_(taskData || {});

    if (!clean.title || !String(clean.title).trim()) return err_('INVALID_INPUT');

    var freq = String(clean.frequency || 'MONTHLY').toUpperCase();
    if (VALID_FREQUENCIES.indexOf(freq) === -1) return err_('INVALID_INPUT');

    var intervalVal = parseInt(clean.intervalValue, 10) || 1;
    var intervalUnit = String(clean.intervalUnit || 'MONTH').toUpperCase();
    if (freq === REC_FREQ.CUSTOM && VALID_UNITS.indexOf(intervalUnit) === -1) return err_('INVALID_INPUT');

    // Assignee validation
    var assigneeEmail = String(clean.assigneeEmail || '').toLowerCase().trim();
    if (!assigneeEmail) return err_('INVALID_INPUT');
    var memberMap = getMemberMap_();
    var assignee = memberMap[assigneeEmail];
    if (!assignee) return err_('NOT_FOUND');

    // Manager scope guard: can only assign to own team
    if (actor.role === 'Manager' && assignee.team !== actor.team) return err_('UNAUTHORIZED');

    // StartDate — default today
    var startDate = clean.startDate ? new Date(clean.startDate) : new Date();
    startDate.setHours(0, 0, 0, 0);

    // EndDate — optional
    var endDate = clean.endDate ? new Date(clean.endDate) : null;
    if (endDate && endDate <= startDate) return err_('INVALID_INPUT');

    // ReminderTime — default 09:00
    var remTime = clean.reminderTime || '09:00';
    if (!/^\d{2}:\d{2}$/.test(remTime)) remTime = '09:00';

    // NextTriggerDate — first fire = startDate
    var nextTrigger = toDateStr_(startDate);

    var recId = generateRecId_();
    var now = new Date();

    getRecSheet_().appendRow([
      recId,                            // A RecurringID
      String(clean.title).trim(),       // B Title
      clean.description || '',          // C Description
      assigneeEmail,                    // D AssigneeEmail
      assignee.name || assigneeEmail,   // E AssigneeName
      assignee.team || '',              // F Team
      freq,                             // G Frequency
      intervalVal,                      // H IntervalValue
      intervalUnit,                     // I IntervalUnit
      nextTrigger,                      // J NextTriggerDate (YYYY-MM-DD)
      remTime,                          // K ReminderTime
      toDateStr_(startDate),            // L StartDate
      endDate ? toDateStr_(endDate) : '',// M EndDate
      actor.email,                      // N CreatedBy
      now.toISOString(),                // O CreatedAt
      REC_STATUS.ACTIVE,                // P Status
      ''                                // Q SourceTaskId
    ]);

    emitEvent_({
      type: EVENT.RECURRING_REMINDER_SENT,
      actorEmail: actor.email,
      actorTeam: actor.team,
      notes: 'Recurring task created: ' + String(clean.title).trim()
    });

    return ok_({ recId: recId, message: 'Recurring task "' + String(clean.title).trim() + '" created.' });

  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('createRecurringTask: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

// ------------------------------------------------------------------
// updateRecurringTask
// Owner can update any. Manager can only update their team's tasks.
// ------------------------------------------------------------------
function updateRecurringTask(recId, updates) {
  try {
    var actor = requireRole_(['Owner', 'Manager']);
    if (!recId) return err_('INVALID_INPUT');

    var sheet = getRecSheet_();
    var found = findRecRow_(recId);
    if (!found) return err_('NOT_FOUND');

    var row = found.data;
    var rIdx = found.rowIndex;
    var clean = sanitize_(updates || {});

    // Manager scope guard
    if (actor.role === 'Manager' && row[REC_COL.TEAM] !== actor.team) return err_('UNAUTHORIZED');

    if (clean.title !== undefined) sheet.getRange(rIdx, REC_COL.TITLE + 1).setValue(String(clean.title).trim() || row[REC_COL.TITLE]);
    if (clean.description !== undefined) sheet.getRange(rIdx, REC_COL.DESCRIPTION + 1).setValue(clean.description);
    if (clean.reminderTime !== undefined && /^\d{2}:\d{2}$/.test(clean.reminderTime)) {
      sheet.getRange(rIdx, REC_COL.REM_TIME + 1).setValue(clean.reminderTime);
    }
    if (clean.endDate !== undefined) {
      var ed = clean.endDate ? new Date(clean.endDate) : null;
      sheet.getRange(rIdx, REC_COL.END_DATE + 1).setValue(ed ? toDateStr_(ed) : '');
    }
    // Allow rescheduling next trigger (e.g. skip a cycle)
    if (clean.nextTriggerDate !== undefined) {
      var nt = new Date(clean.nextTriggerDate);
      if (!isNaN(nt.getTime())) sheet.getRange(rIdx, REC_COL.NEXT_TRIGGER + 1).setValue(toDateStr_(nt));
    }
    if (clean.intervalValue !== undefined && parseInt(clean.intervalValue, 10) > 0) {
      sheet.getRange(rIdx, REC_COL.INTERVAL_VAL + 1).setValue(parseInt(clean.intervalValue, 10));
    }
    if (clean.intervalUnit !== undefined && VALID_UNITS.indexOf(String(clean.intervalUnit).toUpperCase()) !== -1) {
      sheet.getRange(rIdx, REC_COL.INTERVAL_UNIT + 1).setValue(String(clean.intervalUnit).toUpperCase());
    }

    return ok_({ message: 'Recurring task updated.' });

  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('updateRecurringTask: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

// ------------------------------------------------------------------
// pauseRecurringTask / resumeRecurringTask
// ------------------------------------------------------------------
function pauseRecurringTask(recId) {
  return setRecurringStatus_(recId, REC_STATUS.PAUSED);
}

function resumeRecurringTask(recId) {
  return setRecurringStatus_(recId, REC_STATUS.ACTIVE);
}

function setRecurringStatus_(recId, newStatus) {
  try {
    var actor = requireRole_(['Owner', 'Manager']);
    var found = findRecRow_(recId);
    if (!found) return err_('NOT_FOUND');
    if (actor.role === 'Manager' && found.data[REC_COL.TEAM] !== actor.team) return err_('UNAUTHORIZED');
    getRecSheet_().getRange(found.rowIndex, REC_COL.STATUS + 1).setValue(newStatus);
    return ok_({ message: 'Recurring task ' + (newStatus === REC_STATUS.PAUSED ? 'paused' : 'resumed') + '.' });
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('setRecurringStatus_: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

// ------------------------------------------------------------------
// deleteRecurringTask — Owner only
// ------------------------------------------------------------------
function deleteRecurringTask(recId) {
  try {
    requireRole_(['Owner']);
    if (!recId) return err_('INVALID_INPUT');
    var sheet = getRecSheet_();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][REC_COL.ID] !== recId) continue;
      sheet.deleteRow(i + 1);
      return ok_({ message: 'Recurring task deleted.' });
    }
    return err_('NOT_FOUND');
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('deleteRecurringTask: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

// ------------------------------------------------------------------
// getRecurringTasks
// Owner: all rows. Manager: only their team's rows. Member: only their own.
// ------------------------------------------------------------------
function getRecurringTasks() {
  try {
    var actor = requireRole_(['Owner', 'Manager', 'Member']);
    var sheet = getRecSheet_();
    if (!sheet) return ok_([]);
    var data = sheet.getDataRange().getValues();
    var out = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[REC_COL.ID]) continue;
      if (actor.role === 'Manager' && row[REC_COL.TEAM] !== actor.team) continue;
      if (actor.role === 'Member' && String(row[REC_COL.ASSIGNEE_EMAIL] || '').toLowerCase() !== String(actor.email || '').toLowerCase()) continue;

      out.push({
        recId: row[REC_COL.ID],
        title: row[REC_COL.TITLE],
        description: row[REC_COL.DESCRIPTION],
        assigneeEmail: row[REC_COL.ASSIGNEE_EMAIL],
        assigneeName: row[REC_COL.ASSIGNEE_NAME],
        team: row[REC_COL.TEAM],
        frequency: row[REC_COL.FREQUENCY],
        intervalValue: row[REC_COL.INTERVAL_VAL],
        intervalUnit: row[REC_COL.INTERVAL_UNIT],
        nextTriggerDate: toDateStrSafe_(row[REC_COL.NEXT_TRIGGER]),
        reminderTime: row[REC_COL.REM_TIME],
        startDate: toDateStrSafe_(row[REC_COL.START_DATE]),
        endDate: toDateStrSafe_(row[REC_COL.END_DATE]),
        createdBy: row[REC_COL.CREATED_BY],
        status: row[REC_COL.STATUS],
        sourceTaskId: row[REC_COL.SOURCE_TASK_ID] || ''
      });
    }

    return ok_(out);

  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('getRecurringTasks: ' + e.message);
    return { success: false, error: 'SYSTEM_ERROR', message: 'Recurring tasks failed to load: ' + e.message };
  }
}

// ------------------------------------------------------------------
// runRecurringReminderEngine_  (called by runHourlyEngine)
//
// Safety model — two layers:
//   1. tryLock(0): exits immediately if another engine instance runs
//   2. NextTriggerDate advanced per-row immediately after send,
//      before processing next row — crash-safe, no silent misses
//
// Month-end safety: calculateNextTriggerDate_ caps day at last
// valid day of target month.
// ------------------------------------------------------------------
function runRecurringReminderEngine_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    console.log('runRecurringReminderEngine_: lock held, skipping.');
    return 0;
  }
  try {
    var sheet = getRecSheet_();
    if (!sheet) return 0;

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return 0;

    var todayStr = toDateStr_(new Date());
    var sourceTaskMap = buildRecurringSourceTaskMap_();
    var fired = 0;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[REC_COL.ID]) continue;

      // Skip paused
      if (row[REC_COL.STATUS] !== REC_STATUS.ACTIVE) continue;

      // A recurring series may email only while its source task is still On Hold.
      // Done, Archived, moved, or unlinked legacy rows are paused before they can fire.
      var sourceTask = resolveRecurringSourceTask_(row, sourceTaskMap);
      if (!isRecurringSourceReminderEligible_(sourceTask)) {
        sheet.getRange(i + 1, REC_COL.STATUS + 1).setValue(REC_STATUS.PAUSED);
        console.log('runRecurringReminderEngine_: auto-paused inactive/unlinked recurring task ' + row[REC_COL.ID]);
        continue;
      }
      if (!row[REC_COL.SOURCE_TASK_ID] && sourceTask.taskId) {
        sheet.getRange(i + 1, REC_COL.SOURCE_TASK_ID + 1).setValue(sourceTask.taskId);
      }

      // Check end date — if past, auto-pause
      var endDateStr = toDateStrSafe_(row[REC_COL.END_DATE]);
      if (endDateStr && endDateStr < todayStr) {
        sheet.getRange(i + 1, REC_COL.STATUS + 1).setValue(REC_STATUS.PAUSED);
        console.log('runRecurringReminderEngine_: auto-paused expired task ' + row[REC_COL.ID]);
        continue;
      }

      // Check NextTriggerDate
      var nextStr = toDateStrSafe_(row[REC_COL.NEXT_TRIGGER]);
      if (!nextStr || nextStr > todayStr) continue;  // not yet time

      var recId = row[REC_COL.ID];
      var title = row[REC_COL.TITLE] || '(Untitled)';
      var description = row[REC_COL.DESCRIPTION] || '';
      var assigneeEmail = row[REC_COL.ASSIGNEE_EMAIL] || '';
      var assigneeName = row[REC_COL.ASSIGNEE_NAME] || assigneeEmail;
      var freq = row[REC_COL.FREQUENCY] || REC_FREQ.MONTHLY;
      var intVal = row[REC_COL.INTERVAL_VAL] || 1;
      var intUnit = row[REC_COL.INTERVAL_UNIT] || 'MONTH';

      // Send reminder email — wrapped individually so one failure doesn't halt the loop
      try {
        sendRecurringReminderEmail_(assigneeEmail, assigneeName, recId, title, description);
      } catch (emailErr) {
        console.warn('sendRecurringReminderEmail_ failed for ' + recId + ': ' + emailErr.message);
      }

      // Advance NextTriggerDate IMMEDIATELY after send (crash-safe)
      var nextDate = calculateNextTriggerDate_(new Date(nextStr), freq, intVal, intUnit);
      sheet.getRange(i + 1, REC_COL.NEXT_TRIGGER + 1).setValue(toDateStr_(nextDate));

      emitEvent_({
        type: EVENT.RECURRING_REMINDER_SENT,
        actorEmail: assigneeEmail,
        actorTeam: row[REC_COL.TEAM] || '',
        notes: 'Recurring reminder fired: ' + title + ' | Next: ' + toDateStr_(nextDate)
      });

      fired++;
    }

    return fired;

  } finally {
    lock.releaseLock();
  }
}
