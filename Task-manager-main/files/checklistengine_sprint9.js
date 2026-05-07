// ============================================================
// ChecklistEngine.js - per-task subtasks/checklists
// Sheet: Checklists
// Columns: ChecklistID, TaskID, Item, Checked, CreatedAt
// ============================================================

function getChecklistSheet_() {
  var name = (typeof SHEETS !== 'undefined' && SHEETS.CHECKLISTS) ? SHEETS.CHECKLISTS : 'Checklists';
  var sh = getSheet(name);
  if (!sh) {
    var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!id) throw new Error('SPREADSHEET_ID not set in Script Properties');
    var ss = SpreadsheetApp.openById(id);
    sh = ss.insertSheet(name);
    sh.appendRow(['ChecklistID', 'TaskID', 'Item', 'Checked', 'CreatedAt']);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold');
    sh.setFrozenRows(1);
    _sheetCache[name] = sh;
  }
  return sh;
}

function canActorUseChecklist_(actor, taskRow) {
  if (!actor || !taskRow) return false;
  if (actor.role === 'Owner' || actor.role === 'Manager') return canActorAccessTaskRow_(actor, taskRow);
  var actorEmail = String(actor.email || '').toLowerCase();
  var ownerEmail = String(taskRow[COL.OWNER_EMAIL] || '').toLowerCase();
  var creatorEmail = String(taskRow[COL.CREATOR_EMAIL] || '').toLowerCase();
  return actorEmail === ownerEmail || actorEmail === creatorEmail;
}

function getChecklists(taskId) {
  try {
    var actor = requireRole_(['Owner', 'Manager', 'Member']);
    if (!taskId) return err_('INVALID_INPUT');
    var tr = findTaskRow_(taskId);
    if (!tr) return err_('NOT_FOUND');
    if (!canActorUseChecklist_(actor, tr.data)) return err_('UNAUTHORIZED');

    var sheet = getChecklistSheet_();
    var lastRow = sheet.getLastRow();
    var out = [];
    if (lastRow < 2) return ok_(out);

    var matches = sheet.getRange(2, 2, lastRow - 1, 1)
      .createTextFinder(String(taskId))
      .matchEntireCell(true)
      .findAll();

    for (var i = 0; i < matches.length; i++) {
      var row = sheet.getRange(matches[i].getRow(), 1, 1, 5).getValues()[0];
      out.push({
        checklistId: row[0],
        taskId: row[1],
        item: row[2],
        checked: row[3] === true || String(row[3]).toLowerCase() === 'true',
        createdAt: row[4] ? new Date(row[4]).toISOString() : ''
      });
    }
    out.sort(function (a, b) { return new Date(a.createdAt || 0) - new Date(b.createdAt || 0); });
    return ok_(out);
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('getChecklists: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

function addChecklistItem(taskId, item) {
  try {
    var actor = requireRole_(['Owner', 'Manager', 'Member']);
    var cleanItem = sanitize_(item || '').trim();
    if (!taskId || !cleanItem) return err_('INVALID_INPUT');
    var tr = findTaskRow_(taskId);
    if (!tr) return err_('NOT_FOUND');
    if (!canActorUseChecklist_(actor, tr.data)) return err_('UNAUTHORIZED');

    var checklistId = 'CHK-' + Date.now();
    getChecklistSheet_().appendRow([checklistId, taskId, cleanItem, false, new Date()]);
    invalidateTaskIndex_();
    return ok_({ checklistId: checklistId, message: 'Subtask added.' });
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('addChecklistItem: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

function toggleChecklistItem(checklistId, checked) {
  try {
    var actor = requireRole_(['Owner', 'Manager', 'Member']);
    if (!checklistId) return err_('INVALID_INPUT');
    var sheet = getChecklistSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return err_('NOT_FOUND');
    var match = sheet.getRange(2, 1, lastRow - 1, 1)
      .createTextFinder(String(checklistId))
      .matchEntireCell(true)
      .findNext();
    if (!match) return err_('NOT_FOUND');

    var rowNum = match.getRow();
    var row = sheet.getRange(rowNum, 1, 1, 5).getValues()[0];
    var tr = findTaskRow_(row[1]);
    if (!tr) return err_('NOT_FOUND');
    if (!canActorUseChecklist_(actor, tr.data)) return err_('UNAUTHORIZED');

    sheet.getRange(rowNum, 4).setValue(checked === true);
    invalidateTaskIndex_();
    return ok_({ message: checked === true ? 'Subtask checked.' : 'Subtask unchecked.' });
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('toggleChecklistItem: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}
