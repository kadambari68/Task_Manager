// ============================================================
// ProjectEngine.js — TaskFlow v6 (NEW FILE — Phase 3)
//
// Projects layer. Tasks optionally belong to a project via
// the ProjectID field (Tasks col Y).
//
// Sheet: Projects
//   A  ProjectID   — AUTO: PROJ-001
//   B  ProjectName — sanitized
//   C  Description — optional
//   D  OwnerEmail  — creator
//   E  OwnerName
//   F  TeamScope   — team name or "All"
//   G  Status      — Active | On Hold | Completed | Archived
//   H  Health      — Green | Yellow | Red (auto-computed on read)
//   I  StartDate
//   J  DueDate     — optional
//   K  CompletedAt
//   L  CreatedAt
//
// Every public function:
//   1. requireRole_()
//   2. sanitize_() all inputs
//   3. ok_() / err_() response
// ============================================================

// ------------------------------------------------------------------
// SHEET CONSTANT — add to SHEETS object in Code.js
// We define it here as a fallback in case Code.js hasn't been
// updated yet; SHEETS.PROJECTS from Code.js takes precedence.
// ------------------------------------------------------------------
var PROJ_SHEET = 'Projects';

function getProjSheet_() {
  var name = (typeof SHEETS !== 'undefined' && SHEETS.PROJECTS) ? SHEETS.PROJECTS : PROJ_SHEET;
  var sheet = getSheet(name);
  if (!sheet) throw new Error('Projects sheet not found. Run setupProjectsSheet() first.');
  return sheet;
}

function canActorAccessProjectRow_(actor, row) {
  if (!actor || !row) return false;
  var teamScope = row[5] || 'All';
  var ownerEmail = String(row[3] || '').toLowerCase();
  var actorEmail = String(actor.email || '').toLowerCase();
  if (actor.role === 'Owner') return true;
  if (actor.role === 'Manager') {
    return teamScope === 'All' || teamScope === actor.team || ownerEmail === actorEmail;
  }
  return teamScope === 'All' || teamScope === actor.team || ownerEmail === actorEmail;
}

function findProjectRow_(projectId) {
  if (!projectId) return null;
  var data = getProjSheet_().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === projectId) return { rowIndex: i + 1, data: data[i] };
  }
  return null;
}


// ------------------------------------------------------------------
// CREATE PROJECT
// Owner or Manager only.
// ------------------------------------------------------------------
function createProject(projectData) {
  try {
    var actor = requireRole_(['Owner', 'Manager']);
    var clean = sanitize_(projectData);

    if (!clean.name || !clean.name.trim()) return err_('INVALID_INPUT');
    if (actor.role === 'Manager') {
      if (!actor.team) return err_('UNAUTHORIZED');
      if (clean.teamScope && clean.teamScope !== actor.team) return err_('UNAUTHORIZED');
      clean.teamScope = actor.team;
    }

    var sheet = getProjSheet_();
    var existing = sheet.getDataRange().getValues();

    // Uniqueness check
    for (var i = 1; i < existing.length; i++) {
      if ((existing[i][1] || '').toLowerCase() === clean.name.trim().toLowerCase()) {
        return err_('DUPLICATE');
      }
    }

    // Generate ID
    var ids = existing.slice(1)
      .map(function (r) { return r[0] || ''; })
      .filter(function (id) { return /^PROJ-\d+$/.test(id); })
      .map(function (id) { return parseInt(id.replace('PROJ-', ''), 10); });
    var nextNum = ids.length ? Math.max.apply(null, ids) + 1 : 1;
    var projectId = 'PROJ-' + String(nextNum).padStart(3, '0');
    var now = new Date();

    var startDate = clean.startDate ? new Date(clean.startDate) : now;
    var dueDate = clean.dueDate ? new Date(clean.dueDate) : '';

    sheet.appendRow([
      projectId,           // A ProjectID
      clean.name.trim(),   // B ProjectName
      clean.description || '', // C Description
      actor.email,         // D OwnerEmail
      actor.name,          // E OwnerName
      clean.teamScope || 'All', // F TeamScope
      'Active',            // G Status
      'Green',             // H Health
      startDate,           // I StartDate
      dueDate,             // J DueDate
      '',                  // K CompletedAt
      now.toISOString()    // L CreatedAt
    ]);

    // Emit event so analytics can track project creation
    emitEvent_({
      type: 'PROJECT_CREATED',
      taskId: '',
      projectId: projectId,
      actorEmail: actor.email,
      actorTeam: actor.team,
      notes: 'Project created: ' + clean.name.trim()
    });

    return ok_({ projectId: projectId, message: 'Project "' + clean.name.trim() + '" created.' });

  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('createProject: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}


// ------------------------------------------------------------------
// GET PROJECTS FOR USER
// All roles. Members see projects in their team or "All" scope.
// Managers/Owners see all projects.
// Enriches each project with live task stats (one Tasks read).
// ------------------------------------------------------------------
function getProjectsForUser() {
  try {
    var actor = requireRole_(['Owner', 'Manager', 'Member']);

    var sheet = getProjSheet_();
    var projData = sheet.getDataRange().getValues();

    // Build project map from sheet
    var projects = [];
    for (var i = 1; i < projData.length; i++) {
      var r = projData[i];
      if (!r[0]) continue;
      if (!canActorAccessProjectRow_(actor, r)) continue;
      var teamScope = r[5] || 'All';

      projects.push({
        projectId: r[0],
        name: r[1] || '',
        description: r[2] || '',
        ownerEmail: r[3] || '',
        ownerName: r[4] || '',
        teamScope: teamScope,
        status: r[6] || 'Active',
        health: r[7] || 'Green',
        startDate: r[8] ? new Date(r[8]).toISOString() : '',
        dueDate: r[9] ? new Date(r[9]).toISOString() : '',
        completedAt: r[10] ? new Date(r[10]).toISOString() : '',
        createdAt: r[11] ? new Date(r[11]).toISOString() : '',
        // Task stats computed below
        totalTasks: 0,
        doneTasks: 0,
        activeTasks: 0,
        progress: 0
      });
    }

    if (!projects.length) return ok_(projects);

    // Enrich with task counts — single Tasks read ──────────────
    var taskData = getSheet(SHEETS.TASKS).getDataRange().getValues();
    var projMap = {};
    projects.forEach(function (p) { projMap[p.projectId] = p; });

    for (var j = 1; j < taskData.length; j++) {
      var row = taskData[j];
      var pid = row[24]; // Y ProjectID
      if (!pid || !projMap[pid]) continue;
      if (!canActorAccessTaskRow_(actor, row)) continue;

      var proj = projMap[pid];
      var st = row[7] || 'To Do';
      proj.totalTasks++;
      if (st === 'Done' || st === 'Completed' || st === 'Archived') proj.doneTasks++;
      else proj.activeTasks++;
    }

    // Compute progress and auto-health for each project
    projects.forEach(function (p) {
      p.progress = p.totalTasks > 0 ? Math.round((p.doneTasks / p.totalTasks) * 100) : 0;
      if (p.status === 'Active') {
        if (p.progress >= 80) p.computedHealth = 'Green';
        else if (p.activeTasks === 0 && p.totalTasks > 0) p.computedHealth = 'Yellow';
        else p.computedHealth = p.health;
      } else {
        p.computedHealth = p.health;
      }
    });

    // Sort: Active first, then by due date
    projects.sort(function (a, b) {
      if (a.status === 'Active' && b.status !== 'Active') return -1;
      if (a.status !== 'Active' && b.status === 'Active') return 1;
      if (a.dueDate && b.dueDate) return new Date(a.dueDate) - new Date(b.dueDate);
      return a.name.localeCompare(b.name);
    });

    return ok_(projects);

  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('getProjectsForUser: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}


// ------------------------------------------------------------------
// GET PROJECT TASKS
// Returns all tasks belonging to a project.
// Used by Project Detail page.
// ------------------------------------------------------------------
function getProjectTasks(projectId) {
  try {
    var actor = requireRole_(['Owner', 'Manager', 'Member']);
    if (!projectId) return err_('INVALID_INPUT');

    var proj = findProjectRow_(projectId);
    if (!proj) return err_('NOT_FOUND');
    if (!canActorAccessProjectRow_(actor, proj.data)) return err_('UNAUTHORIZED');

    var memberMap = getMemberMap_();
    var now = new Date();
    var taskData = getSheet(SHEETS.TASKS).getDataRange().getValues();
    var tasks = [];

    for (var i = 1; i < taskData.length; i++) {
      var row = taskData[i];
      if (!row[0] || row[24] !== projectId) continue;
      if (!canActorAccessTaskRow_(actor, row)) continue;
      tasks.push(parseTaskRow_(row, memberMap, now));
    }

    // Sort: active first by urgency, done last
    tasks.sort(function (a, b) {
      var doneA = a.status === 'Done' || a.status === 'Completed';
      var doneB = b.status === 'Done' || b.status === 'Completed';
      if (doneA !== doneB) return doneA ? 1 : -1;
      return parseFloat(a.hoursLeft) - parseFloat(b.hoursLeft);
    });

    return ok_(tasks);

  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('getProjectTasks: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}


// ------------------------------------------------------------------
// UPDATE PROJECT
// Owner or Manager. Can update: name, description, status, health,
// teamScope, dueDate. ProjectID and owner are immutable.
// ------------------------------------------------------------------
function updateProject(projectId, updates) {
  try {
    var actor = requireRole_(['Owner', 'Manager']);
    if (!projectId) return err_('INVALID_INPUT');

    var clean = sanitize_(updates || {});
    var sheet = getProjSheet_();
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] !== projectId) continue;

      // Owner can change anything; Manager can only change their own projects
      if (actor.role === 'Manager' && (data[i][3] || '').toLowerCase() !== actor.email.toLowerCase()) {
        return err_('UNAUTHORIZED');
      }

      var validStatuses = ['Active', 'On Hold', 'Completed', 'Archived'];
      var validHealth = ['Green', 'Yellow', 'Red'];

      if (clean.name !== undefined) sheet.getRange(i + 1, 2).setValue(clean.name.trim() || data[i][1]);
      if (clean.description !== undefined) sheet.getRange(i + 1, 3).setValue(clean.description);
      if (clean.teamScope !== undefined) {
        if (actor.role === 'Manager' && clean.teamScope !== actor.team) return err_('UNAUTHORIZED');
        sheet.getRange(i + 1, 6).setValue(clean.teamScope);
      }
      if (clean.status && validStatuses.indexOf(clean.status) !== -1) {
        var prevStatus = data[i][6] || 'Active';
        sheet.getRange(i + 1, 7).setValue(clean.status);
        if (clean.status === 'Completed') sheet.getRange(i + 1, 11).setValue(new Date().toISOString());

        // S3.1 SPRINT 3: Project On Hold → Active transition.
        // Refresh LastActionAt (col X, index 24) on all linked tasks so their
        // SLA clock restarts from now rather than from original creation time.
        if (prevStatus === 'On Hold' && clean.status === 'Active') {
          try {
            var taskSheet = getSheet(SHEETS.TASKS);
            var taskData = taskSheet.getDataRange().getValues();
            var nowIso = new Date().toISOString();
            for (var ti = 1; ti < taskData.length; ti++) {
              if (taskData[ti][24] === projectId && taskData[ti][0]) {
                taskSheet.getRange(ti + 1, 24).setValue(nowIso); // col X = LastActionAt
              }
            }
          } catch (ex) { console.warn('S3.1 task resume: ' + ex.message); }
        }
      }
      if (clean.health && validHealth.indexOf(clean.health) !== -1) {
        sheet.getRange(i + 1, 8).setValue(clean.health);
      }
      if (clean.dueDate) sheet.getRange(i + 1, 10).setValue(new Date(clean.dueDate));

      emitEvent_({
        type: 'PROJECT_UPDATED',
        projectId: projectId,
        actorEmail: actor.email,
        actorTeam: actor.team,
        notes: 'Project updated: ' + JSON.stringify(updates)
      });

      return ok_({ message: 'Project updated.' });
    }

    return err_('NOT_FOUND');

  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('updateProject: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}


// ------------------------------------------------------------------
// ARCHIVE PROJECT — Owner only
// ------------------------------------------------------------------
function archiveProject(projectId) {
  try {
    requireRole_(['Owner']);
    return updateProject(projectId, { status: 'Archived' });
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    return err_('SYSTEM_ERROR');
  }
}


// ------------------------------------------------------------------
// DELETE PROJECT
// Owner role OR project creator
// ------------------------------------------------------------------
function deleteProject(projectId) {
  try {
    var actor = requireRole_(['Owner', 'Manager']);
    if (!projectId) return err_('INVALID_INPUT');

    var sheet = getProjSheet_();
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] !== projectId) continue;

      var projectOwnerEmail = (data[i][3] || '').trim().toLowerCase();
      var actorEmail = actor.email.toLowerCase();

      if (actor.role !== 'Owner' && projectOwnerEmail !== actorEmail) {
        return err_('UNAUTHORIZED');
      }

      sheet.deleteRow(i + 1);

      emitEvent_({
        type: 'PROJECT_DELETED',
        projectId: projectId,
        actorEmail: actor.email,
        actorTeam: actor.team,
        notes: 'Project deleted: ' + projectId
      });

      return ok_({ message: 'Project deleted successfully.' });
    }

    return err_('NOT_FOUND');
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('deleteProject: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}


// ------------------------------------------------------------------
// SETUP PROJECTS SHEET — run once from Apps Script editor
// Add to setupV6Sheets() in AdminEngine.js as well.
// ------------------------------------------------------------------
function setupProjectsSheet() {
  requireRole_(['Owner']);

  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  var ss = SpreadsheetApp.openById(id);

  if (!ss.getSheetByName(PROJ_SHEET)) {
    var ps = ss.insertSheet(PROJ_SHEET);
    ps.appendRow([
      'ProjectID', 'ProjectName', 'Description',
      'OwnerEmail', 'OwnerName', 'TeamScope',
      'Status', 'Health', 'StartDate', 'DueDate',
      'CompletedAt', 'CreatedAt'
    ]);
    ps.getRange(1, 1, 1, 12).setFontWeight('bold');
    Logger.log('Projects sheet created.');
  } else {
    Logger.log('Projects sheet already exists.');
  }
  return ok_({ message: 'Projects sheet ready.' });
}


// ------------------------------------------------------------------
// bootstrapApp PATCH — adds projects to the startup payload.
// Call this from within bootstrapApp() in Code_v6.js:
//
//   projects: getProjectsPayload_(email, user.role)
//
// Or replace the relevant line with a direct call. This helper
// is kept separate so Code.js doesn't have to import this file.
// ------------------------------------------------------------------
function getProjectsPayload_(email, role) {
  try {
    var projResult = getProjectsForUser();
    return projResult.success ? projResult.data : [];
  } catch (e) {
    console.warn('getProjectsPayload_: ' + e.message);
    return [];
  }
}