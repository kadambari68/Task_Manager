// ============================================================
// TaskEngine.js — TaskFlow v6
// Task lifecycle engine. Handles ALL task state changes.
//
// Design rules:
//   - requireRole_() called at top of every public function
//   - All inputs pass through sanitize_() before sheet writes
//   - ALL status changes go through transitionTask_() — no exceptions
//   - emitEvent_() called for every state change
//   - Returns ok_() or err_() — never throws to client
//   - No sheet name strings — always SHEETS.* constants
//   - No member lookups inside loops — getMemberMap_() called once
// ============================================================


// ------------------------------------------------------------------
// CREATE TASK
// Called when user submits New Task form.
//
// Phase 1 improvements over v5:
//   - requireRole_() verifies caller server-side (any authenticated member)
//   - sanitize_() cleans all string inputs
//   - emitEvent_(TASK_CREATED) writes to EventLog
//   - Returns ok_() / err_() envelope
//   - LastActionAt (col X) set on creation
// ------------------------------------------------------------------
var REC_TAG = '__recurring';

function hasRecurringTag_(tags) {
  if (!tags) return false;
  return String(tags).split(',').map(function (t) { return String(t).trim(); }).indexOf(REC_TAG) !== -1;
}

function isRecurringTaskRow_(row) {
  return hasRecurringTag_(row && row[COL.TAGS]);
}

function parseDateOnly_(dateStr) {
  if (!dateStr) return null;
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(dateStr));
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  var d = new Date(dateStr);
  return isNaN(d.getTime()) ? null : d;
}

function createTask(taskData) {
  try {
    var actor = requireRole_(['Owner', 'Manager', 'Member']);
    var clean = sanitize_(taskData);

    // HOTFIX carried to sprint 4: validate taskType against TaskTypes labels.
    // SLAConfig is optional per type — types like "Urgent", "R&D" etc. may have
    // no custom SLA rule but are still valid. Fall back to 24h default.
    var validTypes = getTaskTypes_().map(function (t) { return t.label; });
    var slaConfig = getSLAConfig_();
    var sla = slaConfig[clean.taskType] || { slaHours: 24, gentleReminderHours: 6, escalationEmail: '' };

    if (!clean.taskName || !clean.taskName.trim()) return err_('INVALID_INPUT');
    if (!clean.taskType || validTypes.indexOf(clean.taskType) === -1) return err_('INVALID_INPUT');

    var assignee = getMemberByEmail_(clean.assigneeEmail);
    if (!assignee) return err_('INVALID_INPUT');

    // Assignment scope rules — enforced server-side:
    //   Owner   → can assign to any active member, any team, any role
    //   Manager → can assign to: own team members OR other Managers (cross-dept coordination)
    //             CANNOT assign to Owner role (Owner is not an operational resource)
    //             CANNOT assign to Members of other teams
    //   Member  → can only self-assign (creates task for themselves)
    if (!canActorAssignTarget_(actor, assignee)) return err_('UNAUTHORIZED');
    // End assignment scope rules

    // S3 SPRINT 2: Validate projectId if provided.
    // Previously: any string was accepted, including invented IDs and archived projects.
    if (clean.projectId) {
      var projRow = findProjectRow_(clean.projectId);
      if (!projRow) return err_('INVALID_INPUT');
      if (projRow.data[6] === 'Archived') return err_('INVALID_INPUT');
    }

    var now = null;
    var taskId = null;
    var deadline = null;
    var rec = clean.recurring || null;
    var isRecurring = rec && rec.frequency;

    withScriptLock_(10000, function () {
      now = new Date();
      taskId = generateTaskId();
      deadline = new Date(now.getTime() + sla.slaHours * 3600000);

      if (isRecurring) {
        var recStart = parseDateOnly_(rec && rec.startDate);
        if (recStart) {
          recStart.setHours(0, 0, 0, 0);
          var today = new Date(); today.setHours(0, 0, 0, 0);
          if (recStart < today) {
            if (typeof advanceMonths_ === 'function') recStart = advanceMonths_(recStart, 1);
            else recStart.setMonth(recStart.getMonth() + 1);
          }
          var remTime = (rec && rec.reminderTime) || '09:00';
          var tm = /^(\d{2}):(\d{2})$/.exec(remTime);
          if (tm) recStart.setHours(Number(tm[1]), Number(tm[2]), 0, 0);
          else recStart.setHours(9, 0, 0, 0);
          deadline = recStart;
        } else {
          deadline = now;
        }
      }

      var tags = clean.tags || '';
      if (isRecurring && !hasRecurringTag_(tags)) {
        tags = tags ? (tags + ',' + REC_TAG) : REC_TAG;
      }
      var status = isRecurring ? STATUS.ON_HOLD : STATUS.TODO;
      var onHoldSince = isRecurring ? now.toISOString() : '';

      getSheet(SHEETS.TASKS).appendRow([
        taskId,                             // A TaskID
        clean.taskName.trim(),              // B TaskName
        clean.taskType,                     // C TaskType
        assignee.email,                     // D CurrentOwnerEmail
        assignee.name,                      // E CurrentOwnerName
        assignee.team,                      // F CurrentTeam
        assignee.team,                      // G HomeTeam (sticky team ownership)
        status,                             // H Status — To Do or On Hold for recurring
        actor.name,                         // I CreatedBy
        actor.email,                        // J CreatedByEmail
        now,                                // K CreatedAt
        deadline,                           // L Deadline
        '',                                 // M CompletedAt
        '',                                 // N TotalHours
        clean.driveFileUrl || '',           // O DriveFileURL
        clean.notes || '',                  // P Notes
        0,                                  // Q ReminderCount
        '',                                 // R LastReminderSent
        false,                              // S SLABreached
        sla.slaHours,                       // T SLAHours
        '',                                 // U CalendarEventID
        clean.priority || 'Medium',         // V Priority
        tags,                               // W Tags
        now.toISOString(),                  // X LastActionAt
        clean.projectId || '',              // Y ProjectID
        onHoldSince                          // Z OnHoldSince
      ]);
    });

    // Emit creation event to EventLog
    emitEvent_({
      type: EVENT.TASK_CREATED,
      taskId: taskId,
      projectId: clean.projectId || '',
      actorEmail: actor.email,
      actorTeam: actor.team,
      targetEmail: assignee.email,
      targetTeam: assignee.team,
      fromStatus: '',
      toStatus: isRecurring ? STATUS.ON_HOLD : STATUS.TODO,
      timeSpentHrs: 0,
      slaHours: sla.slaHours,
      slaBreached: false,
      notes: 'Task created for ' + assignee.name
    });

    // Send assignment email (non-blocking)
    try { sendTaskAssignmentEmail(assignee, taskId, clean.taskName.trim(), actor.name, deadline); } catch (e) { }

    // Calendar event for long-SLA tasks (non-blocking)
    if (sla.slaHours > 24) {
      try { createCalendarEvent(taskId, clean.taskName.trim(), assignee.email, deadline); } catch (e) { }
    }

    // Phase 4 — auto-post to team chat (non-fatal if ChatEngine not deployed)
    try { emitSystemMessage_('team', assignee.team, '🆕 ' + taskId + ' assigned to ' + assignee.name + ' by ' + actor.name); } catch (e) { }

    // Recurring: create a recurring schedule row tied to this task (non-blocking)
    if (isRecurring) {
      try { createRecurringFromTask_(taskId, clean, assignee, actor); } catch (e) { console.warn('createRecurringFromTask_ failed for ' + taskId + ': ' + e.message); }
    }

    invalidateTaskIndex_(); // Phase 4
    return ok_({ taskId: taskId, message: 'Task created.' });

  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('createTask: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}


// ------------------------------------------------------------------
// TRANSITION TASK (internal engine — all status changes call this)
//
// This is the ONLY function that changes task status.
// Enforces the state machine, emits events, manages SLA clock.
//
// @param {string}  taskId
// @param {string}  toStatus    must be a STATUS.* value
// @param {Object}  actor       verified user object from requireRole_()
// @param {string}  notes       optional context
// @param {Object}  overrides   optional { targetEmail, targetName, targetTeam }
//                              used by routeTask() to change ownership
// @returns {Object} ok_() or err_()
// ------------------------------------------------------------------
// ------------------------------------------------------------------
// Recurring schedule creation (from New Task form recurring options)
// Creates a RecurringTasks row tied to the task.
// ------------------------------------------------------------------
function createRecurringFromTask_(taskId, clean, assignee, actor) {
  if (!clean || !clean.recurring) return null;
  var rec = clean.recurring || {};

  var freq = (rec && rec.frequency) || 'MONTHLY';
  var intervalVal = parseInt(rec.intervalValue, 10) || 1;
  var intervalUnit = (rec.intervalUnit || 'MONTH').toUpperCase();

  var startDate = parseDateOnly_(rec.startDate) || new Date();
  startDate.setHours(0, 0, 0, 0);

  var today = new Date(); today.setHours(0, 0, 0, 0);
  if (startDate < today) {
    if (typeof advanceMonths_ === 'function') startDate = advanceMonths_(startDate, 1);
    else startDate.setMonth(startDate.getMonth() + 1);
  }

  var endDate = rec.endDate ? parseDateOnly_(rec.endDate) : null;
  if (endDate && isNaN(endDate.getTime())) endDate = null;
  if (endDate && endDate <= startDate) endDate = null;

  var remTime = '09:00';

  var startStr = (typeof toDateStr_ === 'function') ? toDateStr_(startDate) : startDate.toISOString().substring(0, 10);
  var endStr = endDate ? ((typeof toDateStr_ === 'function') ? toDateStr_(endDate) : endDate.toISOString().substring(0, 10)) : '';
  var nextTrigger = startStr;

  var recId = (typeof generateRecId_ === 'function')
    ? generateRecId_()
    : ('REC-' + new Date().getFullYear() + '-' + String(Date.now()).slice(-6));
  var now = new Date();
  var title = String(clean.taskName || '').trim() || '(Untitled)';
  var description = clean.notes || '';
  var status = (typeof REC_STATUS !== 'undefined' && REC_STATUS.ACTIVE) ? REC_STATUS.ACTIVE : 'ACTIVE';

  var sheet = getRecSheet_();
  if (!sheet) return null;

  sheet.appendRow([
    recId,                           // A RecurringID
    title,                           // B Title
    description,                     // C Description
    assignee.email || '',            // D AssigneeEmail
    assignee.name || assignee.email || '', // E AssigneeName
    assignee.team || '',             // F Team
    freq,                            // G Frequency
    intervalVal,                     // H IntervalValue
    intervalUnit,                    // I IntervalUnit
    nextTrigger,                     // J NextTriggerDate
    remTime,                         // K ReminderTime
    startStr,                        // L StartDate
    endStr,                          // M EndDate
    actor.email || '',               // N CreatedBy
    now.toISOString(),               // O CreatedAt
    status,                          // P Status
    taskId                           // Q SourceTaskId
  ]);

  try {
    emitEvent_({
      type: EVENT.RECURRING_REMINDER_SENT,
      actorEmail: actor.email,
      actorTeam: actor.team,
      notes: 'Recurring task created from ' + (taskId || title)
    });
  } catch (e) { }

  return recId;
}

// ------------------------------------------------------------------
// TRANSITION TASK (internal engine â€” all status changes call this)
// ------------------------------------------------------------------
function transitionTask_(taskId, toStatus, actor, notes, overrides) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var taskRow = findTaskRow_(taskId);
    if (!taskRow) return err_('NOT_FOUND');

    var row = taskRow.data;
    var rowIndex = taskRow.rowIndex;
    var fromStatus = row[COL.STATUS] || STATUS.TODO;
    var isRoutingToNewPerson = overrides && overrides.targetEmail && overrides.targetEmail !== row[COL.OWNER_EMAIL];
    var isRecurringTask = isRecurringTaskRow_(row);
    var forceTransition = !!(overrides && overrides.forceTransition === true && actor && actor.role === 'Owner');

    // Recurring tasks: keep On Hold, allow only Done/Archived transitions.
    if (isRecurringTask && [STATUS.ON_HOLD, STATUS.DONE, STATUS.ARCHIVED].indexOf(toStatus) === -1) {
      return err_('INVALID_TRANSITION');
    }

    // State machine check.
    // Exception: allow same-status reassignment when owner actually changes.
    // Recurring override: allow On Hold -> Done/Archived.
    var allowRecurring = isRecurringTask && (toStatus === STATUS.ON_HOLD || toStatus === STATUS.DONE || toStatus === STATUS.ARCHIVED);
    if (!forceTransition && !canTransition_(fromStatus, toStatus) && !allowRecurring && !(isRoutingToNewPerson && fromStatus === toStatus)) {
      return err_('INVALID_TRANSITION');
    }

    var sheet = getSheet(SHEETS.TASKS);
    var now = new Date();
    var slaConfig = getSLAConfig_();
    var sla = slaConfig[row[2]] || { slaHours: 24 };

    // Calculate time spent in previous status for EventLog
    var lastAction = row[COL.LAST_ACTION] ? new Date(row[COL.LAST_ACTION]) : (row[COL.CREATED_AT] ? new Date(row[COL.CREATED_AT]) : now);
    var timeSpentHrs = (now - lastAction) / 3600000;

    // SLA breach check: was task held longer than allowed?
    var slaBreached = (fromStatus !== STATUS.ON_HOLD) && (timeSpentHrs > sla.slaHours);

    // Determine new owner (overrides supplied by routeTask)
    var newOwnerEmail = (overrides && overrides.targetEmail) || row[COL.OWNER_EMAIL];
    var newOwnerName = (overrides && overrides.targetName) || row[COL.OWNER_NAME];
    var newOwnerTeam = (overrides && overrides.targetTeam) || row[COL.CURRENT_TEAM];
    var homeTeam = row[COL.HOME_TEAM] || row[COL.CURRENT_TEAM] || newOwnerTeam;

    // Compute new deadline (reset on routing, preserved on status-only changes)
    var newDeadline = isRoutingToNewPerson
      ? new Date(now.getTime() + sla.slaHours * 3600000)
      : (row[COL.DEADLINE] ? new Date(row[COL.DEADLINE]) : new Date(now.getTime() + sla.slaHours * 3600000));

    // Batch write — single setValues call for performance
    var updates = [
      [newOwnerEmail],  // D col 4
      [newOwnerName],   // E col 5
      [newOwnerTeam],   // F col 6
      [homeTeam],       // G col 7
      [toStatus],       // H col 8
      [newDeadline],    // L col 12 — only deadline updated here, handled via rowIndex
    ];

    // Write to memory array
    row[COL.OWNER_EMAIL] = newOwnerEmail;
    row[COL.OWNER_NAME] = newOwnerName;
    row[COL.CURRENT_TEAM] = newOwnerTeam;
    row[COL.HOME_TEAM] = homeTeam;
    row[COL.STATUS] = toStatus;
    row[COL.SLA_BREACHED] = slaBreached;
    row[COL.LAST_ACTION] = now.toISOString();

    if (isRoutingToNewPerson) {
      row[COL.DEADLINE] = newDeadline;
      row[COL.REMINDER_CT] = 0;
      row[COL.LAST_REMINDER] = '';
    }

    if (toStatus === STATUS.ON_HOLD) {
      row[COL.ON_HOLD_SINCE] = now.toISOString();
    }
    if (fromStatus === STATUS.ON_HOLD) {
      row[COL.ON_HOLD_SINCE] = '';
    }

    if (toStatus === STATUS.DONE) {
      finalizeTask_(taskId, rowIndex, row, actor, timeSpentHrs, sla.slaHours, slaBreached);
    }

    // SPRINT 1 FIX: also reset SLABreached so reopened task starts clean.
    // Previously the breach flag survived the entire new lifecycle of the task.
    if ((fromStatus === STATUS.DONE || fromStatus === STATUS.ARCHIVED) && toStatus !== STATUS.DONE && toStatus !== STATUS.ARCHIVED) {
      row[COL.COMPLETED_AT] = '';
      row[COL.TOTAL_HOURS] = '';
      row[COL.SLA_BREACHED] = false; // clear breach flag
    }

    // BATCHED WRITE
    sheet.getRange(rowIndex, 1, 1, COL.TOTAL_COLS).setValues([row]);

    // Determine event type for EventLog
    var eventType = isRoutingToNewPerson ? EVENT.TASK_ROUTED
      : toStatus === STATUS.DONE ? EVENT.TASK_COMPLETED
        : ((fromStatus === STATUS.DONE || fromStatus === STATUS.ARCHIVED) && toStatus === STATUS.TODO) ? EVENT.TASK_REOPENED
          : EVENT.TASK_STATUS_CHANGED;

    // Emit event — non-blocking
    emitEvent_({
      type: eventType,
      taskId: taskId,
      projectId: row[COL.PROJECT_ID] || '',
      actorEmail: actor.email,
      actorTeam: actor.team,
      targetEmail: newOwnerEmail,
      targetTeam: newOwnerTeam,
      fromStatus: fromStatus,
      toStatus: toStatus,
      timeSpentHrs: timeSpentHrs,
      slaHours: sla.slaHours,
      slaBreached: slaBreached,
      notes: notes || ''
    });

    // Phase 3: HandoffLog write removed. EventLog is now strictly authoritative.

    invalidateTaskIndex_(); // Phase 4

    return ok_({
      taskId: taskId,
      fromStatus: fromStatus,
      toStatus: toStatus,
      taskName: row[1] || ''
    });
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}


// ------------------------------------------------------------------
// ROUTE TASK
// Public function called from UI when user assigns task to someone.
// Wraps transitionTask_() with ownership change.
//
// Phase 1 improvements:
//   - Server-side role check via requireRole_()
//   - Uses transitionTask_() — state machine enforced
//   - SLA breach recorded in EventLog, not just HandoffLog
// ------------------------------------------------------------------
function routeTask(routeData) {
  try {
    var actor = requireRole_(['Owner', 'Manager', 'Member']);
    var clean = sanitize_(routeData);

    if (!clean.taskId || !clean.toEmail) return err_('INVALID_INPUT');

    var taskRow = findTaskRow_(clean.taskId);
    if (!taskRow) return err_('NOT_FOUND');

    var ownerEmail = (taskRow.data[3] || '').toLowerCase();
    var creatorEmail = (taskRow.data[9] || '').toLowerCase();
    var taskTeam = taskRow.data[5] || '';
    var homeTeam = taskRow.data[6] || taskTeam;
    var actorEmail = (actor.email || '').toLowerCase();

    if (actor.role === 'Member' && ownerEmail !== actorEmail) {
      return err_('UNAUTHORIZED');
    }
    if (actor.role === 'Manager' && taskTeam !== actor.team && homeTeam !== actor.team && ownerEmail !== actorEmail && creatorEmail !== actorEmail) {
      return err_('UNAUTHORIZED');
    }

    var newAssignee = getMemberByEmail_(clean.toEmail);
    if (!newAssignee) return err_('INVALID_INPUT');

    // Routing scope rules (mirrors createTask assignment rules):
    //   Owner   → can route to anyone
    //   Manager → can route to own team members OR other Managers; never to Owner
    //   Member  → cannot route (handled above — only current owner can route)
    if (!canActorAssignTarget_(actor, newAssignee)) return err_('UNAUTHORIZED');

    // ── SELF-ROUTE: targetEmail === currentOwner ──────────────────────
    // routeTask always forces IN_PROGRESS, but transitionTask_ flags
    // isRoutingToNewPerson=false (same email) so the state-machine
    // exception (L143) never fires, causing INVALID_TRANSITION when the
    // task is already In Progress.
    var isSelfRoute = newAssignee.email.toLowerCase() === ownerEmail;
    if (isSelfRoute) {
      var currentStatus = taskRow.data[7] || STATUS.TODO;
      // Already In Progress with the same person → no-op, return success
      if (currentStatus === STATUS.IN_PROGRESS) {
        return ok_({ message: taskRow.data[1] + ' is already in progress with you.' });
      }
      // If transition to In Progress is valid → status-only change (no owner swap)
      if (canTransition_(currentStatus, STATUS.IN_PROGRESS)) {
        return transitionTask_(clean.taskId, STATUS.IN_PROGRESS, actor, clean.notes || '', null);
      }
      // Any other status (e.g. Done, Archived) → genuine invalid
      return err_('INVALID_TRANSITION');
    }
    // ── END SELF-ROUTE ────────────────────────────────────────────────

    // Determine target status: routing to someone else → In Progress
    var toStatus = STATUS.IN_PROGRESS;

    var result = transitionTask_(
      clean.taskId,
      toStatus,
      actor,
      clean.notes || '',
      {
        targetEmail: newAssignee.email,
        targetName: newAssignee.name,
        targetTeam: newAssignee.team
      }
    );

    if (!result.success) return result;

    // Send notification to new owner (non-blocking)
    try {
      var deadline = taskRow && taskRow.data[COL.DEADLINE] ? new Date(taskRow.data[COL.DEADLINE]) : null;
      sendTaskRoutedEmail(newAssignee, clean.taskId, result.data.taskName, actor.name, deadline);
    } catch (e) { }

    // Transfer calendar event (non-blocking)
    // SPRINT 1 FIX: result.data has no taskType field — was always undefined, so transfer never fired.
    // Read taskType from taskRow.data[COL.TASK_TYPE] which is populated before transitionTask_.
    try {
      var taskTypeName = taskRow.data[COL.TASK_TYPE] || '';
      var slaH = getSLAConfig_()[taskTypeName] || { slaHours: 24 };
      if (slaH.slaHours > 24) transferCalendarEvent(clean.taskId, newAssignee.email, deadline);
    } catch (e) { }

    // Phase 4 — auto-post to team chat
    try { emitSystemMessage_('team', newAssignee.team, '📋 ' + clean.taskId + ' routed to ' + newAssignee.name + ' by ' + actor.name); } catch (e) { }
    return ok_({ message: result.data.taskName + ' routed to ' + newAssignee.name });

  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('routeTask: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}


// ------------------------------------------------------------------
// CHANGE STATUS (no ownership change — inline board drag/dropdown)
// Called when user drags card between columns OR uses status dropdown.
// Does NOT change assignment. Use routeTask() to change assignee.
// ------------------------------------------------------------------
function changeStatus(taskId, toStatus, notes) {
  try {
    var actor = requireRole_(['Owner', 'Manager', 'Member']);

    if (!taskId || !toStatus) return err_('INVALID_INPUT');

    // Validate toStatus is a known value
    var validStatuses = Object.values(STATUS);
    if (validStatuses.indexOf(toStatus) === -1) return err_('INVALID_INPUT');
    if (toStatus === STATUS.ARCHIVED && actor.role === 'Member') return err_('UNAUTHORIZED');

    // Members can only change status on tasks they own
    if (actor.role === 'Member') {
      var taskRow = findTaskRow_(taskId);
      if (!taskRow) return err_('NOT_FOUND');
      if ((taskRow.data[COL.OWNER_EMAIL] || '').toLowerCase() !== actor.email.toLowerCase()) {
        return err_('UNAUTHORIZED');
      }
    }

    var result = transitionTask_(taskId, toStatus, actor, notes || '', null);
    if (result.success && (toStatus === STATUS.DONE || toStatus === STATUS.ARCHIVED)) {
      try { syncRecurringStatusWithTask_(taskId, 'PAUSED'); } catch (e) { }
    }
    if (result.success && toStatus === STATUS.ON_HOLD) {
      try { syncRecurringStatusWithTask_(taskId, 'ACTIVE'); } catch (e) { }
    }
    return result;

  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('changeStatus: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

// ------------------------------------------------------------------
// UPDATE TASK DETAILS (post-creation edits)
// Editable: assignee, deadline, priority, tags, notes, taskType
// ------------------------------------------------------------------
function updateTaskDetails(payload) {
  try {
    var actor = requireRole_(['Owner', 'Manager', 'Member']);
    var clean = sanitize_(payload || {});
    var taskId = clean.taskId;
    if (!taskId) return err_('INVALID_INPUT');

    var tr = findTaskRow_(taskId);
    if (!tr) return err_('NOT_FOUND');
    var row = tr.data;
    var rowIndex = tr.rowIndex;

    if (actor.role === 'Member') {
      if (String(row[COL.OWNER_EMAIL] || '').toLowerCase() !== String(actor.email || '').toLowerCase()) return err_('UNAUTHORIZED');
    } else if (actor.role === 'Manager') {
      if (!canActorAccessTaskRow_(actor, row)) return err_('UNAUTHORIZED');
    }

    var updates = {};
    var newAssignee = null;

    if (clean.assigneeEmail) {
      var desiredEmail = String(clean.assigneeEmail || '').toLowerCase().trim();
      var currentEmail = String(row[COL.OWNER_EMAIL] || '').toLowerCase().trim();
      if (desiredEmail && desiredEmail !== currentEmail) {
        newAssignee = getMemberByEmail_(desiredEmail);
        if (!newAssignee) return err_('INVALID_INPUT');
        if (!canActorAssignTarget_(actor, newAssignee)) return err_('UNAUTHORIZED');
        updates.assignee = newAssignee;
      }
    }

    if (clean.taskType) {
      var validTypesUpd = getTaskTypes_().map(function (t) { return t.label; });
      if (validTypesUpd.indexOf(clean.taskType) === -1) return err_('INVALID_INPUT');
      var slaCfgUpd = getSLAConfig_();
      var typeSlA = slaCfgUpd[clean.taskType] || { slaHours: 24 };
      updates.taskType = clean.taskType;
      updates.slaHours = Number(typeSlA.slaHours || row[COL.SLA_HOURS] || 24);
    }

    if (clean.priority) {
      var p = String(clean.priority);
      if (['Low', 'Medium', 'High', 'Critical'].indexOf(p) === -1) return err_('INVALID_INPUT');
      updates.priority = p;
    }

    if (clean.deadline) {
      var dl = new Date(clean.deadline);
      if (isNaN(dl.getTime())) return err_('INVALID_INPUT');
      updates.deadline = dl;
    }

    if (clean.tags !== undefined) {
      var tagStr = String(clean.tags || '');
      if (isRecurringTaskRow_(row) && !hasRecurringTag_(tagStr)) {
        tagStr = tagStr ? (tagStr + ',' + REC_TAG) : REC_TAG;
      }
      updates.tags = tagStr;
    }
    // SPRINT 7: Members cannot edit notes, deadline, priority, or task type.
    // Only Owner/Manager can change those fields — prevents unaudited description changes.
    // Members retain ability to update tags only (personal organisation, low-risk).
    if (actor.role !== 'Member') {
      if (clean.notes !== undefined) updates.notes = String(clean.notes || '');
    }

    if (!Object.keys(updates).length) return err_('INVALID_INPUT');

    var now = new Date();
    withScriptLock_(10000, function () {
      var sheet = getSheet(SHEETS.TASKS);
      if (updates.taskType !== undefined) row[COL.TASK_TYPE] = updates.taskType;
      if (updates.assignee) {
        row[COL.OWNER_EMAIL] = updates.assignee.email;
        row[COL.OWNER_NAME] = updates.assignee.name;
        row[COL.CURRENT_TEAM] = updates.assignee.team;
      }
      if (updates.deadline) row[COL.DEADLINE] = updates.deadline;
      if (updates.notes !== undefined) row[COL.NOTES] = updates.notes;
      if (updates.slaHours !== undefined) row[COL.SLA_HOURS] = updates.slaHours;
      if (updates.priority !== undefined) row[COL.PRIORITY] = updates.priority;
      if (updates.tags !== undefined) row[COL.TAGS] = updates.tags;
      row[COL.LAST_ACTION] = now.toISOString();
      sheet.getRange(rowIndex, 1, 1, COL.TOTAL_COLS).setValues([row]);
    });

    emitEvent_({
      type: updates.assignee ? EVENT.TASK_ROUTED : EVENT.TASK_STATUS_CHANGED,
      taskId: taskId,
      projectId: row[COL.PROJECT_ID] || '',
      actorEmail: actor.email,
      actorTeam: actor.team,
      targetEmail: (updates.assignee && updates.assignee.email) || row[COL.OWNER_EMAIL] || '',
      targetTeam: (updates.assignee && updates.assignee.team) || row[COL.CURRENT_TEAM] || '',
      fromStatus: row[COL.STATUS] || '',
      toStatus: row[COL.STATUS] || '',
      timeSpentHrs: 0,
      slaHours: Number((updates.slaHours !== undefined ? updates.slaHours : row[19]) || 24),
      slaBreached: row[COL.SLA_BREACHED] === true,
      notes: (function () {
        var parts = [];
        if (updates.assignee) parts.push('Reassigned to ' + updates.assignee.name);
        if (updates.taskType !== undefined) parts.push('Type -> ' + updates.taskType);
        if (updates.priority !== undefined) parts.push('Priority -> ' + updates.priority);
        if (updates.deadline) parts.push('Deadline updated');
        if (updates.notes !== undefined) parts.push('Notes edited');
        if (updates.tags !== undefined) parts.push('Tags updated');
        return parts.length ? parts.join(', ') + ' by ' + actor.name : 'Task updated by ' + actor.name;
      })()
    });

    // W3 SPRINT 2: Send assignment email to new owner when reassigned via Edit Task modal.
    // Previously: no email sent — new owner had no notification of the change.
    if (updates.assignee) {
      try {
        var taskDeadline = updates.deadline || (row[COL.DEADLINE] ? new Date(row[COL.DEADLINE]) : null);
        sendTaskRoutedEmail(updates.assignee, taskId, row[COL.TASK_NAME] || taskId, actor.name + ' (reassigned)', taskDeadline);
      } catch (e) { }
    }

    invalidateTaskIndex_(); // Phase 4

    return ok_({ taskId: taskId, message: 'Task updated successfully.' });
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('updateTaskDetails: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}


// ------------------------------------------------------------------
// COMPLETE TASK
// Explicit completion. Wraps changeStatus to Done.
// ------------------------------------------------------------------
function completeTask(taskId) {
  try {
    var actor = requireRole_(['Owner', 'Manager', 'Member']);

    // Members can only complete tasks they own
    if (actor.role === 'Member') {
      var taskRow = findTaskRow_(taskId);
      if (!taskRow) return err_('NOT_FOUND');
      if ((taskRow.data[COL.OWNER_EMAIL] || '').toLowerCase() !== actor.email.toLowerCase()) {
        return err_('UNAUTHORIZED');
      }
    }

    var result = transitionTask_(taskId, STATUS.DONE, actor, 'Marked complete', null);
    if (!result.success) return result;
    try { syncRecurringStatusWithTask_(taskId, 'PAUSED'); } catch (e) { }

    // Notify creator — SPRINT 1 FIX: was passing raw row.data array.
    // sendTaskCompletedEmail expects a parsed object {taskName,taskId,createdBy,createdByEmail,totalHours}.
    try {
      var rowAfter = findTaskRow_(taskId);
      if (rowAfter) {
        var parsedTask = parseTaskRow_(rowAfter.data, getMemberMap_(), new Date());
        sendTaskCompletedEmail(parsedTask, actor.name, new Date(), parsedTask.totalHours);
      }
    } catch (e) { }

    // Phase 4 — auto-post to team chat
    try { emitSystemMessage_('team', actor.team, '✅ ' + taskId + ' completed by ' + actor.name); } catch (e) { }
    return ok_({ message: result.data.taskName + ' marked as done.' });

  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('completeTask: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}


// ------------------------------------------------------------------
// ARCHIVE TASK
// Owner/Manager only. Moves Done -> Archived.
// ------------------------------------------------------------------
function archiveTask(taskId, notes) {
  try {
    var actor = requireRole_(['Owner', 'Manager']);
    if (!taskId) return err_('INVALID_INPUT');

    var tr = findTaskRow_(taskId);
    if (!tr) return err_('NOT_FOUND');
    var fromStatus = tr.data[7] || STATUS.TODO;
    var isRecurringTask = isRecurringTaskRow_(tr.data);
    if (fromStatus !== STATUS.DONE && fromStatus !== 'Completed' && !(isRecurringTask && fromStatus === STATUS.ON_HOLD)) {
      return err_('INVALID_TRANSITION');
    }

    var result = transitionTask_(taskId, STATUS.ARCHIVED, actor, notes || 'Task archived', null);
    if (!result.success) return result;

    try { syncRecurringStatusWithTask_(result.data.taskId, 'PAUSED'); } catch (e) { }
    try { emitSystemMessage_('team', actor.team, 'Task ' + taskId + ' archived by ' + actor.name); } catch (e) { }
    return ok_({ message: result.data.taskName + ' archived.' });

  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('archiveTask: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}


// ------------------------------------------------------------------
// REOPEN TASK (drag from Done back to To Do, or admin restart)
// Phase 2 behaviour: filter toggle shows closed tasks on board,
// dragging them left calls this function.
// ------------------------------------------------------------------
function reopenTask(taskId, notes) {
  try {
    // Only Owner/Manager can reopen completed tasks
    var actor = requireRole_(['Owner', 'Manager']);

    var taskRow = findTaskRow_(taskId);
    if (!taskRow) return err_('NOT_FOUND');
    var targetStatus = isRecurringTaskRow_(taskRow.data) ? STATUS.ON_HOLD : STATUS.TODO;
    var result = transitionTask_(taskId, targetStatus, actor, notes || 'Task reopened', null);
    if (!result.success) return result;

    // Notify current assignee (non-blocking)
    try {
      var row = findTaskRow_(taskId);
      if (row) {
        var assignee = getMemberByEmail_(row.data[3]);
        if (assignee) sendTaskAssignmentEmail(assignee, taskId, row.data[1], actor.name + ' (Reopened)', new Date(row.data[11]));
      }
    } catch (e) { }

    try { syncRecurringStatusWithTask_(taskId, 'ACTIVE'); } catch (e) { }
    // Phase 4 — auto-post to team chat
    try { emitSystemMessage_('team', actor.team, '↩️ ' + taskId + ' reopened by ' + actor.name); } catch (e) { }
    return ok_({ message: result.data.taskName + ' has been reopened.' });

  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('reopenTask: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

// Legacy alias kept for frontend backward compat
function restartTask(taskId) { return reopenTask(taskId, 'Restarted'); }

function getLastStatusEventForTask_(taskId) {
  try {
    var sheet = getSheet(SHEETS.EVENT_LOG);
    if (!sheet) return null;

    var matches = sheet.getRange('C:C').createTextFinder(taskId).matchEntireCell(true).findAll();
    var lastCol = sheet.getLastColumn();
    var latest = null;

    for (var i = 0; i < matches.length; i++) {
      var row = sheet.getRange(matches[i].getRow(), 1, 1, lastCol).getValues()[0];
      var type = row[1] || '';
      if (['TASK_STATUS_CHANGED', 'TASK_COMPLETED', 'TASK_REOPENED'].indexOf(type) === -1) continue;
      if (!row[8] || !row[9]) continue;

      var ts = row[14] ? new Date(row[14]) : new Date(0);
      if (!latest || ts > latest.timestamp) {
        latest = {
          type: type,
          fromStatus: row[8],
          toStatus: row[9],
          timestamp: ts
        };
      }
    }
    return latest;
  } catch (e) {
    console.warn('getLastStatusEventForTask_ failed: ' + e.message);
    return null;
  }
}

function reverseLastStatusChange(taskId, notes) {
  try {
    var actor = requireRole_(['Owner']);
    if (!taskId) return err_('INVALID_INPUT');

    var tr = findTaskRow_(taskId);
    if (!tr) return err_('NOT_FOUND');

    var lastEvent = getLastStatusEventForTask_(taskId);
    if (!lastEvent || !lastEvent.fromStatus || !lastEvent.toStatus) {
      return { success: false, message: 'No recent status change was found to reverse.' };
    }

    var currentStatus = String(tr.data[COL.STATUS] || STATUS.TODO);
    if (currentStatus !== String(lastEvent.toStatus || '')) {
      return { success: false, message: 'This task changed again, so the last status move cannot be safely reversed now.' };
    }

    var cleanNote = sanitize_(notes || '').trim();
    var note = 'Owner reversed the last status change from ' + currentStatus + ' to ' + lastEvent.fromStatus + '.';
    if (cleanNote) note += ' Note: ' + cleanNote;

    var result = transitionTask_(taskId, lastEvent.fromStatus, actor, note, { forceTransition: true });
    if (!result.success) return result;

    return ok_({ message: result.data.taskName + ' moved back to ' + lastEvent.fromStatus + '.' });
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('reverseLastStatusChange: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}


// ------------------------------------------------------------------
// FINALIZE TASK (internal — called by transitionTask_ on Done)
// Records completion timestamp and total TAT hours.
// ------------------------------------------------------------------
function finalizeTask_(taskId, rowIndex, row, actor, stageHours, slaHours, slaBreached) {
  var sheet = getSheet(SHEETS.TASKS);
  var now = new Date();
  var createdAt = row[COL.CREATED_AT] ? new Date(row[COL.CREATED_AT]) : now;
  var totalHrs = (now - createdAt) / 3600000;

  row[COL.COMPLETED_AT] = now.toISOString();
  row[COL.TOTAL_HOURS] = totalHrs.toFixed(2);

  // Emit separate SLA_BREACHED event if applicable — makes it queryable in analytics
  if (slaBreached) {
    emitEvent_({
      type: EVENT.SLA_BREACHED,
      taskId: taskId,
      actorEmail: actor.email,
      actorTeam: actor.team,
      slaHours: slaHours,
      slaBreached: true,
      notes: 'SLA breached at completion stage. Held ' + stageHours.toFixed(1) + 'h vs ' + slaHours + 'h SLA'
    });
  }
}


// ------------------------------------------------------------------
// Phase 3: logHandoff_ and logHandoff deprecated.
// All historical data tracking flows through emitEvent_ to EventLog.


// ------------------------------------------------------------------
// GET MY TASKS (legacy alias — now delegates to bootstrapApp pattern)
// ------------------------------------------------------------------
function getMyTasks() {
  return getTasksForCurrentUser();
}

// ══════════════════════════════════════════════════════════
// SPRINT 5 — USER-CONFIGURABLE REMINDERS
// Three public functions: set, get, delete.
// All write to / read from SHEETS.USER_REMINDERS.
//
// Recipients stored as semicolon-delimited string:
//   "a@company.com;b@company.com"
// RemindAt stored as ISO datetime string (absolute, calculated
// by frontend so no server-side timezone arithmetic needed).
// ══════════════════════════════════════════════════════════

// ------------------------------------------------------------------
// setUserReminder(taskId, remindAt, recipients, note)
// Who can call: Owner, Manager, Member (any active user on their task)
// Validates: task exists, actor has access, remindAt is in the future
// ------------------------------------------------------------------
function setUserReminder(taskId, remindAt, recipients, note) {
  try {
    var actor = requireRole_(['Owner', 'Manager', 'Member']);
    if (!taskId) return err_('INVALID_INPUT');
    if (!remindAt) return err_('INVALID_INPUT');

    // Validate remindAt is a real future datetime
    var fireTime = new Date(remindAt);
    if (isNaN(fireTime.getTime())) return err_('INVALID_INPUT');
    if (fireTime <= new Date()) return err_('INVALID_INPUT');  // no past reminders

    // Validate task exists and actor has access
    var taskRow = findTaskRow_(taskId);
    if (!taskRow) return err_('NOT_FOUND');
    if (!canActorAccessTaskRow_(actor, taskRow.data)) return err_('UNAUTHORIZED');

    // Task must still be open — no point reminding on Done/Archived
    var taskStatus = taskRow.data[COL.STATUS] || '';
    if (taskStatus === STATUS.DONE || taskStatus === STATUS.ARCHIVED) {
      return err_('INVALID_INPUT');
    }

    // Build recipients string — default to actor if none supplied
    var recipientList = [];
    if (Array.isArray(recipients) && recipients.length) {
      recipientList = recipients.map(function (e) { return String(e).trim().toLowerCase(); })
        .filter(function (e) { return e.indexOf('@') > 0; });
    }
    if (!recipientList.length) recipientList = [actor.email.toLowerCase()];
    var recipientsStr = recipientList.join(';');

    var taskName = taskRow.data[COL.TASK_NAME] || taskId;
    var reminderId = 'UREM-' + new Date().getFullYear() + '-' + String(Date.now()).slice(-6);
    var now = new Date();
    var eventId = '';

    // Create a calendar event on the creator's calendar.
    // removeAllReminders() clears Google Calendar's default reminders (e.g. default
    // 10-minute email) before adding our single explicit one, preventing double-send.
    try {
      var cal = CalendarApp.getDefaultCalendar();
      var endTime = new Date(fireTime.getTime() + 15 * 60000);
      var desc = 'TaskFlow reminder for ' + taskName + ' (' + taskId + ')' + (note ? ('\\n\\nNote: ' + note) : '');
      var ev = cal.createEvent('[TaskFlow] Reminder: ' + taskName, fireTime, endTime, { description: desc });
      ev.removeAllReminders();
      ev.addEmailReminder(15);
      eventId = ev.getId();
    } catch (e) {
      console.warn('setUserReminder calendar event failed: ' + e.message);
    }

    var reminderSheet = getSheet(SHEETS.USER_REMINDERS);
    if (reminderSheet.getLastColumn() < 10) {
      reminderSheet.getRange(1, 10).setValue('CalendarEventId').setFontWeight('bold');
    }

    reminderSheet.appendRow([
      reminderId,            // A ReminderID
      taskId,                // B TaskID
      taskName,              // C TaskName
      recipientsStr,         // D Recipients (semicolon-delimited)
      fireTime.toISOString(),// E RemindAt (absolute ISO)
      actor.email,           // F CreatedBy
      now.toISOString(),     // G CreatedAt
      false,                 // H Fired
      sanitize_(note || ''), // I Note
      eventId                // J CalendarEventId
    ]);

    return ok_({ reminderId: reminderId, message: 'Reminder set for ' + fireTime.toLocaleString() + '.' });

  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('setUserReminder: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

// ------------------------------------------------------------------
// getUserReminders(taskId)
// Returns non-fired reminders for a task (for display in edit modal).
// Any active user with task access can read.
// ------------------------------------------------------------------
function getUserReminders(taskId) {
  try {
    var actor = requireRole_(['Owner', 'Manager', 'Member']);
    if (!taskId) return err_('INVALID_INPUT');

    var taskRow = findTaskRow_(taskId);
    if (!taskRow) return err_('NOT_FOUND');
    if (!canActorAccessTaskRow_(actor, taskRow.data)) return err_('UNAUTHORIZED');

    var sheet = getSheet(SHEETS.USER_REMINDERS);
    var data = sheet.getDataRange().getValues();
    var out = [];

    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (!r[0]) continue;
      if (r[1] !== taskId) continue;
      if (r[7] === true) continue;  // skip already-fired

      out.push({
        reminderId: r[0],
        taskId: r[1],
        recipients: r[3],            // semicolon string — frontend splits it
        remindAt: r[4],
        createdBy: r[5],
        note: r[8] || ''
      });
    }

    return ok_(out);

  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('getUserReminders: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

// ------------------------------------------------------------------
// deleteUserReminder(reminderId)
// Owner can delete any. Manager/Member can only delete their own.
// Cannot delete already-fired reminders (nothing to cancel).
// ------------------------------------------------------------------
function deleteUserReminder(reminderId) {
  try {
    var actor = requireRole_(['Owner', 'Manager', 'Member']);
    if (!reminderId) return err_('INVALID_INPUT');

    var sheet = getSheet(SHEETS.USER_REMINDERS);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] !== reminderId) continue;

      // Already fired — nothing to cancel
      if (data[i][7] === true) return err_('INVALID_INPUT');

      // Non-Owner can only delete their own reminders
      if (actor.role !== 'Owner') {
        var createdBy = (data[i][5] || '').toLowerCase();
        if (createdBy !== actor.email.toLowerCase()) return err_('UNAUTHORIZED');
      }

      // Remove linked calendar event if present
      var eventId = data[i][9] || '';
      if (eventId) {
        try {
          var ev = CalendarApp.getEventById(eventId);
          if (ev) ev.deleteEvent();
        } catch (e) {
          console.warn('deleteUserReminder calendar event failed: ' + e.message);
        }
      }

      sheet.deleteRow(i + 1);
      return ok_({ message: 'Reminder cancelled.' });
    }

    return err_('NOT_FOUND');

  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('deleteUserReminder: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

/**
 * syncRecurringStatusWithTask_
 * Bridge between Task status and RecurringTasks series.
 * If taskId is archived, PAUSE the series. If reopened, ACTIVATE it.
 */
function syncRecurringStatusWithTask_(taskId, newStatus) {
  try {
    var taskRow = findTaskRow_(taskId);
    if (!taskRow) return;
    if (!isRecurringTaskRow_(taskRow.data)) return;

    var title = taskRow.data[1];
    var owner = taskRow.data[3];

    var recSheet = (typeof getRecSheet_ === 'function') ? getRecSheet_() : getSheet(SHEETS.RECURRING_TASKS);
    if (!recSheet) return;
    var data = recSheet.getDataRange().getValues();
    var statusCol = (typeof REC_COL !== 'undefined' && REC_COL.STATUS !== undefined) ? REC_COL.STATUS + 1 : 16;
    var sourceCol = (typeof REC_COL !== 'undefined' && REC_COL.SOURCE_TASK_ID !== undefined) ? REC_COL.SOURCE_TASK_ID + 1 : 17;
    var ownerLower = (owner || '').toLowerCase();

    for (var i = 1; i < data.length; i++) {
      var sourceTaskId = data[i][sourceCol - 1] || '';
      var sourceMatch = String(sourceTaskId) === String(taskId);
      var legacyMatch = !sourceTaskId && data[i][1] === title && (data[i][3] || '').toLowerCase() === ownerLower;

      if (sourceMatch || legacyMatch) {
        recSheet.getRange(i + 1, statusCol).setValue(newStatus);
        if (!sourceTaskId) recSheet.getRange(i + 1, sourceCol).setValue(taskId);
        console.log('syncRecurringStatusWithTask_: ' + taskId + ' (' + title + ') -> Series ' + data[i][0] + ' set to ' + newStatus);
      }
    }
  } catch (e) {
    console.warn('syncRecurringStatusWithTask_ failed: ' + e.message);
  }
}
