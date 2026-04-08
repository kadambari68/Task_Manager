// ============================================================
// Code.js — TaskFlow v6 | Foundation Layer
// Phase 1: Core architecture — all other files depend on this.
//
// Responsibilities (one per section, SOLID compliant):
//   §1  bootstrapApp()     Single page-load call
//   §2  requireRole_()     Server-side auth guard
//   §3  sanitize_()        Input cleaning
//   §4  ok_() / err_()     Response envelope (no internal leakage)
//   §5  canTransition_()   State machine rules
//   §6  emitEvent_()       Central EventLog writer
//   §7  getSheet()         Sheet accessor (reads ID from Properties)
//   §8  O(n) map helpers   Build once, lookup O(1)
//   §9  Reference fetchers SLA, task types, teams (cached)
//   §10 getTaskPage_()     Paginated, role-filtered task fetcher
//   §11 Finders            findTaskRow_() — internal only
//   §12 ID generators
//   §13 Legacy compat      Keeps CalendarEngine/old frontend working
//   §14 Migration helper   One-time status rename script
// ============================================================

// ------------------------------------------------------------------
// SHEET CONSTANTS — must match tab names exactly
// ------------------------------------------------------------------
const SHEETS = {
  TASKS: 'Tasks',
  ATTACHMENTS: 'Attachments',
  EVENT_LOG: 'EventLog',
  HANDOFF_LOG: 'HandoffLog',
  REMINDERS_LOG: 'RemindersLog',
  TEAM_MEMBERS: 'TeamMembers',
  TEAMS: 'Teams',
  SLA_CONFIG: 'SLAConfig',
  TASK_TYPES: 'TaskTypes',
  COMMENTS: 'Comments',
  PROJECTS: 'Projects',   // Phase 3
  CHAT: 'Chat',        // Phase 4
  CHAT_SPACES: 'ChatSpaces',
  GOALS: 'Goals',
  USER_REMINDERS: 'UserReminders',  // Sprint 5: user-configurable reminders
  RECURRING_TASKS: 'RecurringTasks', // Sprint 6: recurring task definitions
  CHECKLISTS: 'Checklists' // Sprint 9: subtasks
};

// ------------------------------------------------------------------
// STATUS CONSTANTS — 6-state machine, only valid strings in system
// ------------------------------------------------------------------
const STATUS = {
  TODO: 'To Do',
  IN_PROGRESS: 'In Progress',
  IN_REVIEW: 'In Review',
  ON_HOLD: 'On Hold',
  DONE: 'Done',
  ARCHIVED: 'Archived'
};

// ------------------------------------------------------------------
// STATE MACHINE — valid transitions (single source of truth)
// Key = current status. Value = array of allowed next statuses.
// ------------------------------------------------------------------
const TASK_TRANSITIONS = {
  'To Do': ['In Progress', 'On Hold'],
  'In Progress': ['In Review', 'On Hold', 'Done'],
  'In Review': ['In Progress', 'On Hold', 'Done'],
  'On Hold': ['To Do', 'In Progress'],
  'Done': ['To Do', 'Archived'],
  'Archived': ['To Do']
};

// ------------------------------------------------------------------
// EVENT TYPE CONSTANTS
// ------------------------------------------------------------------
const EVENT = {
  TASK_CREATED: 'TASK_CREATED',
  TASK_ROUTED: 'TASK_ROUTED',
  TASK_STATUS_CHANGED: 'TASK_STATUS_CHANGED',
  TASK_COMPLETED: 'TASK_COMPLETED',
  TASK_REOPENED: 'TASK_REOPENED',
  SLA_BREACHED: 'SLA_BREACHED',
  SLA_AT_RISK: 'SLA_AT_RISK',
  REMINDER_SENT: 'REMINDER_SENT',
  TASK_IDLE: 'TASK_IDLE',
  USER_REMINDER_SENT: 'USER_REMINDER_SENT',  // Sprint 5: user-set reminder fired
  GOAL_CREATED: 'GOAL_CREATED',               // Sprint 4 compat
  RECURRING_REMINDER_SENT: 'RECURRING_REMINDER_SENT', // Sprint 6: recurring task fired
  ATTACHMENT_ADDED: 'ATTACHMENT_ADDED',
  ATTACHMENT_DUPLICATE_SKIPPED: 'ATTACHMENT_DUPLICATE_SKIPPED',
  ATTACHMENT_UPLOAD_FAILED: 'ATTACHMENT_UPLOAD_FAILED'
};

const REMINDER = { GENTLE: 'Gentle', FIRM: 'Firm', ESCALATION: 'Escalation', IDLE: 'Idle' };

// ------------------------------------------------------------------
// COLUMN INDEX CONSTANTS - Tasks sheet (0-based)
// Single source of truth. If columns move, update ONLY here.
// ------------------------------------------------------------------
const COL = {
  TASK_ID: 0,   // A
  TASK_NAME: 1,   // B
  TASK_TYPE: 2,   // C
  OWNER_EMAIL: 3,   // D
  OWNER_NAME: 4,   // E
  CURRENT_TEAM: 5,   // F
  HOME_TEAM: 6,   // G
  STATUS: 7,   // H
  CREATED_BY: 8,   // I
  CREATOR_EMAIL: 9,   // J
  CREATED_AT: 10,  // K
  DEADLINE: 11,  // L
  COMPLETED_AT: 12,  // M
  TOTAL_HOURS: 13,  // N
  DRIVE_URL: 14,  // O
  NOTES: 15,  // P
  REMINDER_CT: 16,  // Q
  LAST_REMINDER: 17,  // R
  SLA_BREACHED: 18,  // S
  SLA_HOURS: 19,  // T
  CAL_EVENT_ID: 20,  // U
  PRIORITY: 21,  // V
  TAGS: 22,  // W
  LAST_ACTION: 23,  // X
  PROJECT_ID: 24,  // Y
  ON_HOLD_SINCE: 25,  // Z
  TOTAL_COLS: 26   // total column count
};

// ------------------------------------------------------------------
// EXECUTION-SCOPED TASK INDEX - build once, O(1) lookups
// Eliminates repeated full-table scans in findTaskRow_().
// ------------------------------------------------------------------
var _taskIndex = null;

function buildTaskIndex_() {
  if (_taskIndex) return _taskIndex;

  var cache = CacheService.getScriptCache();
  try {
    var cached = cache.get('task_index_v6');
    if (cached) {
      _taskIndex = JSON.parse(cached);
      return _taskIndex;
    }
  } catch (e) { }

  var sheet = getSheet(SHEETS.TASKS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    _taskIndex = {};
    try { cache.put('task_index_v6', '{}', 300); } catch (e) { }
    return _taskIndex;
  }

  var _taskData = sheet.getRange(2, 1, lastRow - 1, COL.TOTAL_COLS).getValues();
  _taskIndex = {};
  for (var i = 0; i < _taskData.length; i++) {
    var id = _taskData[i][COL.TASK_ID];
    if (id) _taskIndex[id] = { rowIndex: i + 2, data: _taskData[i] };
  }

  try {
    var json = JSON.stringify(_taskIndex);
    if (json.length < 100000) { cache.put('task_index_v6', json, 300); }
  } catch (e) { }

  return _taskIndex;
}

function invalidateTaskIndex_() {
  _taskIndex = null;
  try { CacheService.getScriptCache().remove('task_index_v6'); } catch (e) { }
}

// ------------------------------------------------------------------
// WEB APP ENTRY POINT
// ------------------------------------------------------------------
function doGet(e) {
  // SPRINT 9: PWA Manifest serving
  if (e && e.parameter && e.parameter.pwa === 'manifest.json') {
    var url = ScriptApp.getService().getUrl();
    var manifest = {
      "name": "TaskFlow v6",
      "short_name": "TaskFlow",
      "start_url": url,
      "display": "standalone",
      "background_color": "#ffffff",
      "theme_color": "#172b4d",
      "icons": [{
        "src": "https://fonts.gstatic.com/s/i/short-term/release/googlesymbols/task_alt/default/192px.svg",
        "sizes": "192x192",
        "type": "image/svg+xml"
      }]
    };
    return ContentService.createTextOutput(JSON.stringify(manifest))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var tmpl = HtmlService.createTemplateFromFile('index');
  tmpl.SCRIPT_URL = ScriptApp.getService().getUrl();
  return tmpl.evaluate().setTitle('TaskFlow').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ══════════════════════════════════════════════════════════
// §1 BOOTSTRAP — one call, all data the client needs
// ══════════════════════════════════════════════════════════

function bootstrapApp() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) return err_('UNAUTHORIZED');

    // SPRINT 8B: Batch sheet read. Pre-fetch all sheets into _sheetCache 
    // to avoid N sequential SpreadsheetApp API calls during app startup.
    if (!_ss) {
      var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
      if (id) {
        _ss = SpreadsheetApp.openById(id);
        var allSheets = _ss.getSheets();
        for (var i = 0; i < allSheets.length; i++) {
          _sheetCache[allSheets[i].getName()] = allSheets[i];
        }
      }
    }

    const memberMap = getMemberMap_();
    const user = memberMap[email.toLowerCase().trim()];
    if (!user || !user.active) return err_('UNAUTHORIZED');

    // Load projects (non-fatal — new Projects sheet may not exist yet)
    var projects = [];
    try { projects = getProjectsPayload_(email, user.role); } catch (e) { /* sheet not yet created */ }

    // Chat unread count (non-fatal — Chat sheet may not exist yet)
    var unreadCount = 0;
    try { var uc = getUnreadCount(); if (uc && uc.success) unreadCount = uc.data.total; } catch (e) { }

    return ok_({
      user: user,
      members: Object.values(memberMap).filter(function (m) { return m.active; }).sort(function (a, b) { return a.name.localeCompare(b.name); }),
      teams: getAllTeams_(),
      taskTypes: getTaskTypes_(),
      slaConfig: getSLAConfig_(),
      tasks: getTaskPage_(email, user.role, user.team, memberMap, 0, 150), // Phase 4: Pagination limit
      projects: projects,
      unreadCount: unreadCount
    });
  } catch (e) {
    console.error('bootstrapApp: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

// ══════════════════════════════════════════════════════════
// §2 SECURITY GUARD — every public write function calls this
// ══════════════════════════════════════════════════════════

function requireRole_(allowedRoles) {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('UNAUTHORIZED');

  const cacheKey = 'mbr_' + email.replace(/[^a-z0-9]/gi, '_');
  const cache = CacheService.getUserCache();
  let member = null;

  try {
    const hit = cache.get(cacheKey);
    if (hit) member = JSON.parse(hit);
  } catch (e) { }

  if (!member) {
    const map = getMemberMap_();
    member = map[email.toLowerCase().trim()] || null;
    if (member) {
      try { cache.put(cacheKey, JSON.stringify(member), 300); } catch (e) { }
    }
  }

  if (!member || !member.active) throw new Error('UNAUTHORIZED');
  if (!allowedRoles.includes(member.role)) throw new Error('UNAUTHORIZED');
  return member;
}

function withScriptLock_(waitMs, fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(waitMs || 10000);
  try {
    return fn();
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

function canActorAccessTaskRow_(actor, row) {
  if (!actor || !row) return false;
  if (actor.role === 'Owner') return true;
  var ownerEmail = String(row[COL.OWNER_EMAIL] || '').toLowerCase();
  var creatorEmail = String(row[COL.CREATOR_EMAIL] || '').toLowerCase();
  var rowTeam = String(row[COL.CURRENT_TEAM] || '');
  var homeTeam = String(row[COL.HOME_TEAM] || row[COL.CURRENT_TEAM] || '');
  var creatorTeam = '';
  if (creatorEmail) {
    var creator = getMemberByEmail_(creatorEmail);
    creatorTeam = creator ? String(creator.team || '') : '';
  }
  var actorEmail = String(actor.email || '').toLowerCase();
  if (actor.role === 'Manager') {
    return rowTeam === actor.team || homeTeam === actor.team || creatorTeam === actor.team || ownerEmail === actorEmail || creatorEmail === actorEmail;
  }
  return ownerEmail === actorEmail || creatorEmail === actorEmail;
}

function getCurrentUser() {
  var email = Session.getActiveUser().getEmail();
  return getMemberByEmail_(email);
}

function isAdmin() {
  try { requireRole_(['Owner', 'Manager']); return true; } catch (e) { return false; }
}

// ══════════════════════════════════════════════════════════
// §3 SANITIZE — all user input passes through here
// ══════════════════════════════════════════════════════════

function sanitize_(input) {
  if (input === null || input === undefined) return '';
  if (typeof input === 'string') {
    // Preserve legitimate business text like "R&D" while still stripping unsafe tag brackets.
    return input.replace(/[<>]/g, '').replace(/[\r\n\t]/g, ' ').trim().substring(0, 500);
  }
  if (typeof input === 'number' || typeof input === 'boolean') return input;
  if (Array.isArray(input)) return input.map(sanitize_);
  if (typeof input === 'object') {
    var out = {};
    Object.keys(input).forEach(function (k) { out[k] = sanitize_(input[k]); });
    return out;
  }
  return '';
}

// ══════════════════════════════════════════════════════════
// §4 RESPONSE ENVELOPE — no internal details ever leaked
// ══════════════════════════════════════════════════════════

function ok_(data) { return { success: true, data: data }; }

function err_(code) {
  var MESSAGES = {
    UNAUTHORIZED: 'You do not have permission to perform this action.',
    NOT_FOUND: 'The requested item could not be found.',
    INVALID_INPUT: 'One or more fields are invalid.',
    INVALID_TRANSITION: 'This status change is not allowed.',
    UPLOAD_FAILED: 'Files could not be uploaded. Please try again.',
    SYSTEM_ERROR: 'Something went wrong. Please try again.',
    DUPLICATE: 'This item already exists.'
  };
  return { success: false, error: code, message: MESSAGES[code] || MESSAGES.SYSTEM_ERROR };
}

// ══════════════════════════════════════════════════════════
// §5 STATE MACHINE
// ══════════════════════════════════════════════════════════

function canTransition_(fromStatus, toStatus) {
  var allowed = TASK_TRANSITIONS[fromStatus];
  return Array.isArray(allowed) && allowed.indexOf(toStatus) !== -1;
}

// ══════════════════════════════════════════════════════════
// §6 EVENT EMITTER — ONLY writer to EventLog sheet
//    Non-throwing: a log failure must not block the action.
// ══════════════════════════════════════════════════════════

function emitEvent_(params) {
  if (!params || !params.type || !params.actorEmail) {
    console.warn('emitEvent_: missing required fields');
    return null;
  }
  try {
    var eventId = 'EVT-' + Date.now();
    withScriptLock_(5000, function () {
      getSheet(SHEETS.EVENT_LOG).appendRow([
        eventId,
        params.type,
        params.taskId || '',
        params.projectId || '',
        params.actorEmail,
        params.actorTeam || '',
        params.targetEmail || '',
        params.targetTeam || '',
        params.fromStatus || '',
        params.toStatus || '',
        Number(params.timeSpentHrs || 0).toFixed(2),
        Number(params.slaHours || 0),
        params.slaBreached === true,
        sanitize_(params.notes || ''),
        new Date().toISOString(),
        JSON.stringify(params.meta || {})
      ]);
    });
    return { id: eventId };
  } catch (e) {
    console.error('emitEvent_ write failed: ' + e.message);
    return null;
  }
}

// ══════════════════════════════════════════════════════════
// §7 SHEET ACCESSOR — reads SPREADSHEET_ID from Script Properties
// ══════════════════════════════════════════════════════════

var _ss = null; // execution-scoped spreadsheet cache
var _sheetCache = {}; // execution-scoped sheet cache

function getSheet(name) {
  if (_sheetCache[name]) return _sheetCache[name];

  if (!_ss) {
    var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!id) throw new Error('SPREADSHEET_ID not set in Script Properties');
    _ss = SpreadsheetApp.openById(id);
  }

  _sheetCache[name] = _ss.getSheetByName(name);
  return _sheetCache[name];
}

// ══════════════════════════════════════════════════════════
// §8 O(n) MAP HELPERS — build once, O(1) lookup
// ══════════════════════════════════════════════════════════

// S3.1 SPRINT 3: Returns { projectId → status } for all projects.
// Used by reminderengine to skip tasks linked to On Hold projects.
// Sheet col G (index 6, 0-based) = Status.
function getProjectStatusMap_() {
  try {
    var sheet = getSheet('Projects');
    if (!sheet) return {};
    var data = sheet.getDataRange().getValues();
    var map = {};
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) map[data[i][0]] = data[i][6] || 'Active';
    }
    return map;
  } catch (e) {
    return {};
  }
}

function getMemberMap_() {
  var cache = CacheService.getScriptCache();
  var key = 'member_map_v6';
  try {
    var hit = cache.get(key);
    if (hit) return JSON.parse(hit);
  } catch (e) { }

  var data = getSheet(SHEETS.TEAM_MEMBERS).getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[2]) continue;
    var em = String(row[2]).toLowerCase().trim();
    map[em] = {
      id: row[0] || '',
      name: row[1] || '',
      email: em,
      team: row[3] || '',
      role: row[4] || 'Member',
      active: row[5] === true,
      phone: row[6] ? String(row[6]).trim() : ''
    };
  }
  try { cache.put(key, JSON.stringify(map), 600); } catch (e) { }
  return map;
}

function invalidateMemberCache_() {
  CacheService.getScriptCache().remove('member_map_v6');
}

function getMemberByEmail_(email) {
  if (!email) return null;
  return getMemberMap_()[email.toLowerCase().trim()] || null;
}

function getActiveMembers_() {
  return Object.values(getMemberMap_())
    .filter(function (m) { return m.active; })
    .sort(function (a, b) { return a.name.localeCompare(b.name); });
}

// ══════════════════════════════════════════════════════════
// §9 REFERENCE DATA — SLA, task types, teams (cached)
// ══════════════════════════════════════════════════════════

function getSLAConfig_() {
  var cache = CacheService.getScriptCache();
  var key = 'sla_config_v6';
  try { var hit = cache.get(key); if (hit) return JSON.parse(hit); } catch (e) { }

  var data = getSheet(SHEETS.SLA_CONFIG).getDataRange().getValues();
  var config = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    config[row[0]] = {
      fromTeam: row[1] || '',
      toTeam: row[2] || '',
      slaHours: Number(row[3]) || 24,
      gentleReminderHours: Number(row[4]) || 6,
      escalationEmail: row[5] || ''
    };
  }
  try { cache.put(key, JSON.stringify(config), 900); } catch (e) { }
  return config;
}

function getSLAConfig() { return getSLAConfig_(); }

function getTaskTypes_() {
  return getSheet(SHEETS.TASK_TYPES).getDataRange().getValues().slice(1)
    .filter(function (r) { return r[0]; })
    .map(function (r) { return { id: r[0], label: r[1], fromTeam: r[2] || '', toTeam: r[3] || '', description: r[4] || '' }; });
}
function getTaskTypes() { return getTaskTypes_(); }

function getAllTeams_() {
  var sheet = getSheet(SHEETS.TEAMS);
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1)
    .filter(function (r) { return r[0]; })
    .map(function (r) { return { id: r[0], name: r[1], color: r[2] || '#0052CC', description: r[3] || '', managerEmail: r[4] || '', createdAt: r[5] ? new Date(r[5]).toISOString() : '' }; });
}

function getTeamMembers() {
  try {
    var actor = requireRole_(['Owner', 'Manager']);
    var out = Object.values(getMemberMap_());
    if (actor.role === 'Manager') {
      out = out.filter(function (m) { return m.team === actor.team; });
    }
    out.sort(function (a, b) { return a.name.localeCompare(b.name); });
    return out.map(function (m) {
      return {
        id: m.id || '',
        name: m.name || '',
        email: m.email || '',
        team: m.team || '',
        role: m.role || 'Member',
        phone: m.phone || '',
        isActive: m.active === true,
        active: m.active === true
      };
    });
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return [];
    console.error('getTeamMembers: ' + e.message);
    return [];
  }
}
function getMemberByEmail(email) { return getMemberByEmail_(email); }

// ══════════════════════════════════════════════════════════
// §10 PAGINATED TASK FETCHER
//     Reads all rows ONCE, filters in memory, returns one page.
//     memberMap passed in — never re-read inside this function.
// ══════════════════════════════════════════════════════════

function getTaskPage_(email, role, userTeam, memberMap, offset, limit) {
  limit = limit || 50;
  offset = offset || 0;

  var sheet = getSheet(SHEETS.TASKS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { tasks: [], total: 0, hasMore: false, offset: 0 };

  var allData = sheet.getRange(2, 1, lastRow - 1, 26).getValues();
  var now = new Date();
  var isOwner = role === 'Owner';
  var isMgr = role === 'Manager';
  userTeam = userTeam || '';
  var emailLower = email.toLowerCase();
  var visible = [];

  for (var i = 0; i < allData.length; i++) {
    var row = allData[i];
    if (!row[0]) continue;
    var ownerEmail = (row[COL.OWNER_EMAIL] || '').toLowerCase();
    var creatorEmail = (row[COL.CREATOR_EMAIL] || '').toLowerCase();
    var creatorTeam = (memberMap && memberMap[creatorEmail]) ? (memberMap[creatorEmail].team || '') : '';

    var rowTeam = row[COL.CURRENT_TEAM] || '';
    var homeTeam = row[COL.HOME_TEAM] || rowTeam || '';
    var canSee = isOwner
      || (isMgr && (rowTeam === userTeam || homeTeam === userTeam || creatorTeam === userTeam || ownerEmail === emailLower || creatorEmail === emailLower))
      || (!isOwner && !isMgr && (ownerEmail === emailLower || creatorEmail === emailLower));

    if (!canSee) continue;
    visible.push(row);
  }

  // Sort: non-done tasks by urgency first, done tasks last
  visible.sort(function (a, b) {
    var doneA = a[7] === STATUS.DONE || a[7] === STATUS.ARCHIVED || a[7] === 'Completed';
    var doneB = b[7] === STATUS.DONE || b[7] === STATUS.ARCHIVED || b[7] === 'Completed';
    if (doneA !== doneB) return doneA ? 1 : -1;
    var hA = a[11] ? (new Date(a[11]) - now) / 3600000 : 999;
    var hB = b[11] ? (new Date(b[11]) - now) / 3600000 : 999;
    return hA - hB;
  });

  var total = visible.length;
  var page = visible.slice(offset, offset + limit);
  var hasMore = (offset + limit) < total;
  var isMemberRole = (role === 'Member');
  return {
    tasks: page.map(function (row) {
      var t = parseTaskRow_(row, memberMap, now);
      // S3.4 SPRINT 3: Strip phone PII for Members — they have no escalation use case.
      if (isMemberRole) t.ownerPhone = '';
      return t;
    }),
    total: total,
    hasMore: hasMore,
    offset: offset
  };
}

function getTasksForCurrentUser() {
  var email = Session.getActiveUser().getEmail();
  var memberMap = getMemberMap_();
  var user = memberMap[email.toLowerCase()] || { role: 'Member', team: '' };
  return getTaskPage_(email, user.role, user.team, memberMap, 0, 1000).tasks;
}

// Column map (0-based): A=0 TaskID, B=1 Name, C=2 Type, D=3 OwnerEmail,
// E=4 OwnerName, F=5 CurrentTeam, G=6 HomeTeam, H=7 Status, I=8 CreatedBy,
// J=9 CreatedByEmail, K=10 CreatedAt, L=11 Deadline, M=12 CompletedAt,
// N=13 TotalHours, O=14 DriveURL, P=15 Notes, Q=16 ReminderCount,
// R=17 LastReminderSent, S=18 SLABreached, T=19 SLAHours,
// U=20 CalendarEventID, V=21 Priority, W=22 Tags,
// X=23 LastActionAt, Y=24 ProjectID, Z=25 OnHoldSince
function parseTaskRow_(row, memberMap, now) {
  now = now || new Date();
  var dl = row[COL.DEADLINE] ? new Date(row[COL.DEADLINE]) : null;
  var hoursL = dl ? (dl - now) / 3600000 : 999;
  var ownerEm = (row[COL.OWNER_EMAIL] || '').toLowerCase().trim();
  var ownerM = memberMap ? (memberMap[ownerEm] || null) : null;
  var rawTags = row[COL.TAGS] || '';
  var isRecurring = rawTags ? rawTags.split(',').some(function (t) { return t.trim() === '__recurring'; }) : false;
  var tags = rawTags ? rawTags.split(',').map(function (t) { return String(t).trim(); }).filter(function (t) { return t && t !== '__recurring'; }).join(',') : '';
  return {
    taskId: row[COL.TASK_ID] || '',
    taskName: row[COL.TASK_NAME] || '',
    taskType: row[COL.TASK_TYPE] || '',
    currentOwnerEmail: row[COL.OWNER_EMAIL] || '',
    currentOwnerName: row[COL.OWNER_NAME] || '',
    currentTeam: row[COL.CURRENT_TEAM] || '',
    homeTeam: row[COL.HOME_TEAM] || row[COL.CURRENT_TEAM] || '',
    status: row[COL.STATUS] || STATUS.TODO,
    createdBy: row[COL.CREATED_BY] || '',
    createdByEmail: row[COL.CREATOR_EMAIL] || '',
    createdTimestamp: row[COL.CREATED_AT] ? new Date(row[COL.CREATED_AT]).toISOString() : '',
    deadline: row[COL.DEADLINE] ? new Date(row[COL.DEADLINE]).toISOString() : '',
    completedAt: row[COL.COMPLETED_AT] ? new Date(row[COL.COMPLETED_AT]).toISOString() : '',
    totalHours: Number(row[COL.TOTAL_HOURS]) || 0,
    driveFileUrl: row[COL.DRIVE_URL] || '',
    notes: row[COL.NOTES] || '',
    reminderCount: Number(row[COL.REMINDER_CT]) || 0,
    slaBreached: row[COL.SLA_BREACHED] === true,
    slaHours: Number(row[COL.SLA_HOURS]) || 24,
    priority: row[COL.PRIORITY] || 'Medium',
    tags: tags,
    lastActionAt: row[COL.LAST_ACTION] ? new Date(row[COL.LAST_ACTION]).toISOString() : '',
    projectId: row[COL.PROJECT_ID] || '',
    hoursLeft: hoursL.toFixed(1),
    isOverdue: hoursL < 0,
    ownerPhone: ownerM ? (ownerM.phone || '') : '',
    isRecurring: isRecurring
  };
}

// ══════════════════════════════════════════════════════════
// §11 FINDERS — internal only, row indices never sent to client
// ══════════════════════════════════════════════════════════

function findTaskRow_(taskId) {
  if (!taskId) return null;
  var idx = buildTaskIndex_();
  return idx[taskId] || null;
}

function findTaskRow(taskId) { return findTaskRow_(taskId); } // legacy

// ══════════════════════════════════════════════════════════
// §12 ID GENERATORS
// ══════════════════════════════════════════════════════════

// S3.3 SPRINT 3: Lock-protected ScriptProperties counter.
// Previous: getLastRow() returned the same value for concurrent creates and
// wrong IDs after deletions. Fix: atomic increment under ScriptLock.
// Seed: on first call, seeds from sheet row count so existing IDs are preserved.
function generateTaskId() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var props = PropertiesService.getScriptProperties();
    var PROP = 'TASK_COUNTER_' + new Date().getFullYear();
    var current = parseInt(props.getProperty(PROP) || '0', 10);
    if (current === 0) {
      // First call this year — seed from existing row count so IDs stay sequential
      current = Math.max(getSheet(SHEETS.TASKS).getLastRow() - 1, 0);
    }
    var next = current + 1;
    props.setProperty(PROP, String(next));
    return 'TASK-' + new Date().getFullYear() + '-' + String(next).padStart(4, '0');
  } finally {
    lock.releaseLock();
  }
}

function generateLogId(prefix) { return (prefix || 'LOG') + '-' + Date.now(); }

// ══════════════════════════════════════════════════════════
// §13 LEGACY COMPAT — keeps CalendarEngine and v5 frontend working
// ══════════════════════════════════════════════════════════

function getLastHandoffTime(taskId) {
  // Check new EventLog first, then fall back to HandoffLog
  var latest = null;
  try {
    var elData = getSheet(SHEETS.EVENT_LOG).getDataRange().getValues();
    for (var i = 1; i < elData.length; i++) {
      if (elData[i][2] === taskId && elData[i][14]) {
        var t = new Date(elData[i][14]);
        if (!latest || t > latest) latest = t;
      }
    }
  } catch (e) { }

  if (latest) return latest;

  try {
    var hlData = getSheet(SHEETS.HANDOFF_LOG).getDataRange().getValues();
    for (var j = 1; j < hlData.length; j++) {
      if (hlData[j][1] === taskId && hlData[j][6]) {
        var t2 = new Date(hlData[j][6]);
        if (!latest || t2 > latest) latest = t2;
      }
    }
  } catch (e) { }

  return latest || new Date();
}

function isAdminEmail(email) {
  var m = getMemberByEmail_(email);
  return m && (m.role === 'Owner' || m.role === 'Manager');
}

function getAllTasksAdmin() {
  var email = Session.getActiveUser().getEmail();
  if (!isAdminEmail(email)) return [];
  return getTaskPage_(email, 'Owner', '', getMemberMap_(), 0, 1000).tasks;
}

function getMyCreatedTasks() {
  var email = Session.getActiveUser().getEmail().toLowerCase();
  var sheet = getSheet(SHEETS.TASKS);
  var data = sheet.getDataRange().getValues();
  var memberMap = getMemberMap_();
  var now = new Date();
  var tasks = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    if ((row[9] || '').toLowerCase() !== email) continue;
    tasks.push(parseTaskRow_(row, memberMap, now));
  }
  tasks.sort(function (a, b) {
    var doneA = a.status === STATUS.DONE || a.status === STATUS.ARCHIVED || a.status === 'Completed';
    var doneB = b.status === STATUS.DONE || b.status === STATUS.ARCHIVED || b.status === 'Completed';
    if (doneA && !doneB) return 1;
    if (!doneA && doneB) return -1;
    return parseFloat(a.hoursLeft) - parseFloat(b.hoursLeft);
  });
  return tasks;
}

function getSLAConfigRows() {
  var config = getSLAConfig_();
  return Object.keys(config).map(function (k) {
    var v = config[k];
    return {
      taskType: k, fromTeam: v.fromTeam, toTeam: v.toTeam, slaHours: v.slaHours,
      gentleAfterHours: v.gentleReminderHours, escalationEmail: v.escalationEmail
    };
  });
}

function getTaskTimeline(taskId) {
  try {
    var actor = requireRole_(['Owner', 'Manager', 'Member']);
    var tr = findTaskRow_(taskId);
    if (!tr) return [];
    if (!canActorAccessTaskRow_(actor, tr.data)) return [];

    var entries = [];
    // EventLog (new events) - Phase 4: TextFinder optimization to avoid full scan
    try {
      var sheet = getSheet(SHEETS.EVENT_LOG);
      var matches = sheet.getRange("C:C").createTextFinder(taskId).matchEntireCell(true).findAll();
      var lastCol = sheet.getLastColumn();
      for (var i = 0; i < matches.length; i++) {
        var r = sheet.getRange(matches[i].getRow(), 1, 1, lastCol).getValues()[0];
        entries.push({
          source: 'event', type: r[1], fromTeam: r[5], fromEmail: r[4],
          toTeam: r[7], toEmail: r[6], fromStatus: r[8], toStatus: r[9],
          timeSpentHours: Number(r[10] || 0).toFixed(1),
          slaHours: Number(r[11] || 0), slaBreached: r[12] === true,
          notes: r[13] || '', timestamp: r[14] || ''
        });
      }
    } catch (e) { }
    // HandoffLog (legacy events)
    try {
      var hlData = getSheet(SHEETS.HANDOFF_LOG).getDataRange().getValues();
      for (var j = 1; j < hlData.length; j++) {
        var hl = hlData[j];
        if (hl[1] !== taskId) continue;
        entries.push({
          source: 'legacy', type: 'TASK_ROUTED', fromTeam: hl[2] || '', fromEmail: hl[3] || '',
          toTeam: hl[4] || '', toEmail: hl[5] || '', fromStatus: '', toStatus: '',
          timeSpentHours: Number(hl[7] || 0).toFixed(1),
          slaHours: Number(hl[8] || 0), slaBreached: hl[9] === true,
          notes: hl[10] || '', timestamp: hl[6] ? new Date(hl[6]).toISOString() : ''
        });
      }
    } catch (e) { }
    entries.sort(function (a, b) { return new Date(a.timestamp) - new Date(b.timestamp); });
    return entries.slice(-50); // Phase 4 Payload limit
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return [];
    console.error('getTaskTimeline: ' + e.message);
    return [];
  }
}

// ══════════════════════════════════════════════════════════
// §15 SPRINT 5 SETUP — run once from Apps Script editor
// ══════════════════════════════════════════════════════════

// Creates the UserReminders sheet with correct headers.
// Safe to re-run — exits cleanly if sheet already exists.
// Schema (A→I):
//   A ReminderID   B TaskID       C TaskName
//   D Recipients   E RemindAt     F CreatedBy
//   G CreatedAt    H Fired        I Note (optional context)
function setupUserRemindersSheet() {
  requireRole_(['Owner']);
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  var ss = SpreadsheetApp.openById(id);
  if (ss.getSheetByName(SHEETS.USER_REMINDERS)) {
    return ok_({ message: 'UserReminders sheet already exists.' });
  }
  var sh = ss.insertSheet(SHEETS.USER_REMINDERS);
  sh.appendRow(['ReminderID', 'TaskID', 'TaskName', 'Recipients',
    'RemindAt', 'CreatedBy', 'CreatedAt', 'Fired', 'Note', 'CalendarEventId']);
  sh.getRange(1, 1, 1, 10).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setColumnWidth(4, 220);
  sh.setColumnWidth(5, 160);
  Logger.log('UserReminders sheet created.');
  return ok_({ message: 'UserReminders sheet created successfully.' });
}

// NOTE: setupRecurringTasksSheet() is defined in recurringtasksengine.js.
// It was removed from here (Sprint 7 fix) to prevent "Identifier already
// declared" GAS deployment error. Run setupRecurringTasksSheet() from
// the GAS editor once to create the RecurringTasks sheet.

// ------------------------------------------------------------------
// installTriggers â€” Sprint 6
// Creates/refreshes the hourly trigger for runHourlyEngine.
// Run manually once from the Apps Script editor.
// ------------------------------------------------------------------
function installTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'runHourlyEngine' || fn === 'runFrequentReminderEngine' || fn === 'warmDashCache') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('runHourlyEngine').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('warmDashCache').timeBased().everyMinutes(4).create();
  return ok_({ message: 'Triggers installed: runHourlyEngine (hourly) + warmDashCache (every 4 min).' });
}

// ══════════════════════════════════════════════════════════
// §14 ONE-TIME MIGRATION — run once from Apps Script editor
// ══════════════════════════════════════════════════════════

function migrateStatusValues() {
  var user = getCurrentUser();
  if (!user || user.role !== 'Owner') return err_('UNAUTHORIZED');

  var sheet = getSheet(SHEETS.TASKS);
  var data = sheet.getDataRange().getValues();
  var MAP = { 'Open': 'To Do', 'Completed': 'Done', 'Escalated': 'In Progress' };
  var n = 0;

  for (var i = 1; i < data.length; i++) {
    var nw = MAP[data[i][7]];
    if (!nw) continue;
    sheet.getRange(i + 1, 8).setValue(nw);
    if (!data[i][23]) sheet.getRange(i + 1, 24).setValue(new Date().toISOString());
    n++;
  }
  console.log('Migrated ' + n + ' rows');
  return ok_({ migrated: n });
}

// One-time helper: restores HomeTeam (col G) from EventLog TASK_CREATED target team.
// Run manually after deploying the HomeTeam visibility patch.
function migrateTaskHomeTeamFromEvents() {
  var user = getCurrentUser();
  if (!user || user.role !== 'Owner') return err_('UNAUTHORIZED');

  var tasksSheet = getSheet(SHEETS.TASKS);
  var data = tasksSheet.getDataRange().getValues();
  if (data.length < 2) return ok_({ updated: 0 });

  var createdTeamByTask = {};
  try {
    var ev = getSheet(SHEETS.EVENT_LOG).getDataRange().getValues();
    for (var i = 1; i < ev.length; i++) {
      var row = ev[i];
      var type = row[1] || '';
      var taskId = row[2] || '';
      if (!taskId || type !== EVENT.TASK_CREATED) continue;
      var targetTeam = row[7] || row[5] || '';
      if (targetTeam && !createdTeamByTask[taskId]) createdTeamByTask[taskId] = targetTeam;
    }
  } catch (e) { }

  var updated = 0;
  for (var r = 1; r < data.length; r++) {
    var taskId = data[r][0];
    if (!taskId) continue;
    var currentTeam = data[r][5] || '';
    var desiredHomeTeam = createdTeamByTask[taskId] || data[r][6] || currentTeam || '';
    if (!desiredHomeTeam) continue;
    if (data[r][6] !== desiredHomeTeam) {
      tasksSheet.getRange(r + 1, 7).setValue(desiredHomeTeam); // G HomeTeam
      updated++;
    }
  }
  return ok_({ updated: updated });
}
