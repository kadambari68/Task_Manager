// ============================================================
// goalsengine.js — TaskFlow v6 Sprint 4
//
// Simple goal tracking similar to HubSpot goals:
//   GOALS sheet columns (A→H):
//   A GoalID       B GoalName       C OwnerEmail
//   D Target       E MetricType     F StartDate
//   G EndDate      H Description
//
// MetricType values: tasksCompleted | tasksClosed
// Progress is always computed live from task data — never stored.
// ============================================================

var GOALS_SHEET = 'Goals';
var GOALS_COLS = 9; // Added CreatedBy column
var GOAL_SCOPE_ALL = 'all';

// ------------------------------------------------------------------
// setupGoalsSheet
// Run once from GAS editor to create the sheet.
// ------------------------------------------------------------------
function setupGoalsSheet() {
  requireRole_(['Owner']);
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  var ss = SpreadsheetApp.openById(id);
  if (ss.getSheetByName(GOALS_SHEET)) {
    var sh = ss.getSheetByName(GOALS_SHEET);
    if (sh.getLastColumn() < GOALS_COLS) {
      sh.getRange(1, 9).setValue('CreatedBy').setFontWeight('bold');
    }
    return ok_({ message: 'Goals sheet updated/checked.' });
  }
  var sh = ss.insertSheet(GOALS_SHEET);
  sh.appendRow(['GoalID', 'GoalName', 'OwnerEmail', 'Target', 'MetricType', 'StartDate', 'EndDate', 'Description', 'CreatedBy']);
  sh.getRange(1, 1, 1, GOALS_COLS).setFontWeight('bold');
  sh.setFrozenRows(1);
  return ok_({ message: 'Goals sheet created successfully.' });
}

// ------------------------------------------------------------------
// getGoalSheet_ helper
// ------------------------------------------------------------------
function getGoalSheet_() {
  var sh = getSheet(GOALS_SHEET);
  if (!sh) throw new Error('Goals sheet not found. Run setupGoalsSheet() first.');
  return sh;
}

// ------------------------------------------------------------------
// generateGoalId_
// ------------------------------------------------------------------
function generateGoalId_() {
  return 'GOAL-' + new Date().getFullYear() + '-' + String(Date.now()).slice(-5);
}

function normalizeGoalScope_(rawScope, actor) {
  var raw = String(rawScope || GOAL_SCOPE_ALL).trim();
  if (!raw || raw === 'Team') raw = GOAL_SCOPE_ALL;
  var rawLower = raw.toLowerCase();

  if (rawLower === GOAL_SCOPE_ALL) {
    if (actor.role !== 'Owner') throw new Error('UNAUTHORIZED');
    return GOAL_SCOPE_ALL;
  }

  if (rawLower.indexOf('t:') === 0) {
    var teamIn = raw.substring(2).trim();
    if (!teamIn) throw new Error('INVALID_INPUT');
    var teams = getAllTeams_() || [];
    var team = teams.find(function (t) {
      return String(t.name || '').toLowerCase() === teamIn.toLowerCase();
    });
    if (!team) throw new Error('INVALID_INPUT');
    if (actor.role === 'Manager' && String(team.name || '').toLowerCase() !== String(actor.team || '').toLowerCase()) throw new Error('UNAUTHORIZED');
    return 't:' + team.name;
  }

  var email = rawLower.indexOf('m:') === 0 ? raw.substring(2) : raw;
  email = String(email || '').toLowerCase().trim();
  var member = getMemberByEmail_(email);
  if (!member || member.active === false) throw new Error('INVALID_INPUT');
  if (actor.role === 'Manager' && String(member.team || '').toLowerCase() !== String(actor.team || '').toLowerCase()) throw new Error('UNAUTHORIZED');
  return 'm:' + member.email.toLowerCase();
}

function canActorManageGoal_(actor, row) {
  if (actor.role === 'Owner') return true;
  var scope = String(row[2] || '').toLowerCase().trim();
  var creator = String(row[8] || '').toLowerCase().trim();
  var actorEm = String(actor.email || '').toLowerCase();
  if (creator === actorEm) return true;
  if (scope.indexOf('t:') === 0) return scope.substring(2).toLowerCase() === String(actor.team || '').toLowerCase();
  if (scope.indexOf('m:') === 0) {
    var member = getMemberByEmail_(scope.substring(2));
    return !!(member && String(member.team || '').toLowerCase() === String(actor.team || '').toLowerCase());
  }
  return false;
}

// ------------------------------------------------------------------
// createGoal
// Owner or Manager only.
// ------------------------------------------------------------------
function createGoal(goalData) {
  try {
    var actor = requireRole_(['Owner', 'Manager']);
    var clean = sanitize_(goalData || {});

    if (!clean.goalName || !clean.goalName.trim()) return err_('INVALID_INPUT');
    if (!clean.target || Number(clean.target) <= 0) return err_('INVALID_INPUT');

    var validMetrics = ['tasksCompleted', 'tasksClosed'];
    var metricType = clean.metricType || 'tasksCompleted';
    if (validMetrics.indexOf(metricType) === -1) return err_('INVALID_INPUT');

    var goalId = generateGoalId_();
    var ownerEmail = normalizeGoalScope_(clean.ownerEmail, actor);
    var startDate = clean.startDate ? new Date(clean.startDate) : new Date();
    // Normalize startDate to beginning of day so tasks completed earlier today count
    startDate.setHours(0, 0, 0, 0);

    var endDate = clean.endDate ? new Date(clean.endDate) : null;
    if (endDate) endDate.setHours(23, 59, 59, 999);
    if (endDate && endDate <= startDate) return err_('INVALID_INPUT');

    var sheet = getGoalSheet_();
    sheet.appendRow([
      goalId,
      clean.goalName.trim(),
      ownerEmail,
      Number(clean.target),
      metricType,
      startDate.toISOString(),
      endDate ? endDate.toISOString() : '',
      clean.description || '',
      actor.email // Store creator
    ]);

    invalidateDashCache_();

    emitEvent_({
      type: 'GOAL_CREATED',
      actorEmail: actor.email,
      actorTeam: actor.team,
      notes: 'Goal created: ' + clean.goalName.trim()
    });

    return ok_({ goalId: goalId, message: 'Goal "' + clean.goalName.trim() + '" created.' });

  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    if (e.message === 'INVALID_INPUT') return err_('INVALID_INPUT');
    console.error('createGoal: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

// ------------------------------------------------------------------
// updateGoal
// Owner or goal creator (Manager) only.
// ------------------------------------------------------------------
function updateGoal(goalId, updates) {
  try {
    var actor = requireRole_(['Owner', 'Manager']);
    if (!goalId) return err_('INVALID_INPUT');

    var sheet = getGoalSheet_();
    var data = sheet.getDataRange().getValues();
    var clean = sanitize_(updates || {});

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] !== goalId) continue;

      if (!canActorManageGoal_(actor, data[i])) {
        return err_('UNAUTHORIZED');
      }

      if (clean.goalName !== undefined) sheet.getRange(i + 1, 2).setValue(clean.goalName.trim() || data[i][1]);
      if (clean.ownerEmail !== undefined) sheet.getRange(i + 1, 3).setValue(normalizeGoalScope_(clean.ownerEmail, actor));
      if (clean.target !== undefined && Number(clean.target) > 0) sheet.getRange(i + 1, 4).setValue(Number(clean.target));
      if (clean.metricType !== undefined) {
        var validMetrics = ['tasksCompleted', 'tasksClosed'];
        if (validMetrics.indexOf(clean.metricType) === -1) return err_('INVALID_INPUT');
        sheet.getRange(i + 1, 5).setValue(clean.metricType);
      }
      if (clean.startDate !== undefined) {
        var sd = new Date(clean.startDate); sd.setHours(0, 0, 0, 0);
        sheet.getRange(i + 1, 6).setValue(sd.toISOString());
      }
      if (clean.endDate !== undefined) {
        var ed = clean.endDate ? new Date(clean.endDate) : null;
        if (ed) ed.setHours(23, 59, 59, 999);
        sheet.getRange(i + 1, 7).setValue(ed ? ed.toISOString() : '');
      }
      if (clean.description !== undefined) sheet.getRange(i + 1, 8).setValue(clean.description);

      invalidateDashCache_();
      return ok_({ message: 'Goal updated.' });
    }
    return err_('NOT_FOUND');

  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    if (e.message === 'INVALID_INPUT') return err_('INVALID_INPUT');
    console.error('updateGoal: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

// ------------------------------------------------------------------
// deleteGoal
// Owner only.
// ------------------------------------------------------------------
function deleteGoal(goalId) {
  try {
    requireRole_(['Owner']);
    if (!goalId) return err_('INVALID_INPUT');

    var sheet = getGoalSheet_();
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] !== goalId) continue;
      sheet.deleteRow(i + 1);
      invalidateDashCache_();
      return ok_({ message: 'Goal deleted.' });
    }
    return err_('NOT_FOUND');

  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('deleteGoal: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

// ------------------------------------------------------------------
// getGoalProgress
// Returns all goals with live-calculated progress.
// Uses getAnalyticsDashboard internally so no extra sheet reads.
// ------------------------------------------------------------------
function getGoalProgress(filters) {
  try {
    requireRole_(['Owner', 'Manager']);
    var dash = getAnalyticsDashboard(filters || {});
    if (!dash.success) return dash;
    return ok_(dash.data.goalProgress || []);
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('getGoalProgress: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}
