// ============================================================
// ReminderEngine.js — TaskFlow v6
// Time-triggered engine. Runs every hour via Apps Script trigger.
//
// Phase 1 improvements over v5:
//   - checkIdleTasks()  NEW: detects tasks with no action > 24h
//   - checkSLARisk()    NEW: early warning at 80% SLA consumption
//   - sendWeeklyDigest() Updated to use EventLog via AnalyticsEngine
//   - runDailyReminderEngine() Updated: skips On Hold tasks (SLA paused)
//   - All reads use getMemberMap_() — no getMemberByEmail inside loops
//   - emitEvent_() called for all alerts (makes them queryable)
//
// Trigger setup:
//   runHourlyEngine      → Time-driven, Hour timer
//   sendWeeklyDigest     → Time-driven, Weekly, Monday 8am
// ============================================================


// ------------------------------------------------------------------
// MASTER HOURLY ENGINE
// Single trigger entry point. Runs all checks in sequence.
// Individual check failures are caught so one failure doesn't
// prevent the others from running.
// ------------------------------------------------------------------
function runHourlyEngine() {
  var now = new Date();
  console.log('Hourly engine started: ' + now.toISOString());

  var results = {
    slaReminders: 0, slaRisk: 0, idleAlerts: 0,
    userReminders: 0, recurringReminders: 0, cleaned: 0, errors: []
  };

  try { results.slaReminders = runDailyReminderEngine_(); }
  catch (e) { results.errors.push('SLA reminders: ' + e.message); }

  try { results.slaRisk = checkSLARisk_(); }
  catch (e) { results.errors.push('SLA risk: ' + e.message); }

  try { results.idleAlerts = checkIdleTasks_(); }
  catch (e) { results.errors.push('Idle tasks: ' + e.message); }

  // Sprint 5: user-configurable reminders — runs every hour
  try { results.userReminders = runUserReminderEngine_(); }
  catch (e) { results.errors.push('User reminders: ' + e.message); }

  // Sprint 6: recurring task reminders — runs every hour, date-gated internally
  try { results.recurringReminders = runRecurringReminderEngine_(); }
  catch (e) { results.errors.push('Recurring reminders: ' + e.message); }

  // Sprint 5: cleanup — only at midnight (hour 0) to avoid repeated deletes
  if (now.getHours() === 0) {
    try { results.cleaned = cleanupOldReminders_(); }
    catch (e) { results.errors.push('Cleanup: ' + e.message); }
  }

  console.log('Hourly engine done: ' + JSON.stringify(results));
}

// ------------------------------------------------------------------
// FREQUENT REMINDER ENGINE (lightweight)
// Runs user + recurring reminders more often to reduce drift.
// ------------------------------------------------------------------
// Legacy trigger alias — if someone already has this set as their trigger
function runDailyReminderEngine() { return runHourlyEngine(); }


// ------------------------------------------------------------------
// SLA REMINDERS (Gentle → Firm → Escalation)
// Scans all active tasks and sends appropriate reminder tier.
// Uses getMemberMap_() once — no per-row sheet reads.
//
// Phase 2 change: skips ON HOLD tasks (SLA clock paused).
// ------------------------------------------------------------------
function runDailyReminderEngine_() {
  // Sprint 5: tryLock(0) — exit immediately if another instance is running.
  // Triggers must skip, not queue. Prevents duplicate SLA emails on overlap.
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    console.log('runDailyReminderEngine_: lock held, skipping this run.');
    return 0;
  }
  try {
    var sheet = getSheet(SHEETS.TASKS);
    var data = sheet.getDataRange().getValues();
    var slaConfig = getSLAConfig_();
    var memberMap = getMemberMap_();        // built once — O(1) per lookup below
    var now = new Date();
    var sent = 0;

    // S3.2 SPRINT 3: Build today-sent set once — O(n) total instead of O(tasks × types).
    // Previously wasReminderSentToday_() re-read the entire RemindersLog per task per type.
    // At ~500 rows / 6 months this would time out the GAS 6-minute limit.
    var todayStr = now.toISOString().substring(0, 10);
    var todayAlerted = {};  // key: taskId + '|' + reminderType → true
    try {
      var remData = getSheet(SHEETS.REMINDERS_LOG).getDataRange().getValues();
      for (var r = 1; r < remData.length; r++) {
        if (!remData[r][5]) continue;
        if (new Date(remData[r][5]).toISOString().substring(0, 10) !== todayStr) continue;
        todayAlerted[remData[r][1] + '|' + remData[r][4]] = true;
      }
    } catch (e) { /* RemindersLog may not exist yet */ }

    // S3.1 SPRINT 3: Build project status map once — skip tasks linked to On Hold projects.
    var projStatusMap = getProjectStatusMap_();  // { projectId → status }

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var status = row[7];

      // Skip: done, on hold (SLA clock paused), or empty row
      if (!row[0]) continue;
      if (status === STATUS.DONE || status === STATUS.ARCHIVED) continue;
      if (status === STATUS.ON_HOLD) continue;  // Phase 2 — On Hold pauses SLA

      var taskId = row[0];
      var taskName = row[1];
      var taskType = row[2];
      var ownerEmail = row[3];
      var ownerName = row[4];
      var deadline = row[11] ? new Date(row[11]) : null;
      var reminderCount = Number(row[16]) || 0;

      // S3.1 SPRINT 3: Skip if this task belongs to an On Hold project.
      var taskProjId = row[24] || '';
      if (taskProjId && projStatusMap[taskProjId] === 'On Hold') continue;

      if (!deadline) continue;

      var sla = slaConfig[taskType] || { slaHours: 24, gentleReminderHours: 6, escalationEmail: '' };

      var hoursLeft = (deadline - now) / 3600000;
      var escalationEmail = sla.escalationEmail;

      // Use memberMap for phone lookup — O(1), no extra sheet read
      var ownerMember = memberMap[ownerEmail ? ownerEmail.toLowerCase() : ''];
      var ownerPhone = ownerMember ? (ownerMember.phone || '') : '';

      // CASE 1: OVERDUE — send escalation to manager + direct alert to owner (once/day)
      if (hoursLeft < 0) {
        if (!todayAlerted[taskId + '|' + REMINDER.ESCALATION]) {  // S3.2: O(1) lookup
          var overdue = Math.abs(hoursLeft).toFixed(1);
          sendEscalationEmail(escalationEmail, ownerName, ownerEmail, taskId, taskName, overdue);
          // sendOverdueOwnerEmail_ is now handled inside sendFirmReminderEmail stub (isOverdue=true)
          sendFirmReminderEmail(ownerEmail, ownerName, taskId, taskName, hoursLeft.toFixed(1), true);
          logReminder_(taskId, ownerEmail, ownerName, REMINDER.ESCALATION, reminderCount + 1);
          updateReminderFields_(sheet, i + 1, reminderCount + 1, now);
          emitEvent_({
            type: EVENT.REMINDER_SENT, taskId: taskId,
            actorEmail: ownerEmail, actorTeam: row[5],
            slaBreached: true, notes: 'Escalation sent. Overdue by ' + overdue + 'h',
            meta: { reminderType: REMINDER.ESCALATION, channel: 'Email', ownerPhone: ownerPhone }
          });
          sent++;
        }

        // CASE 2: APPROACHING DEADLINE — no email. Board shows SLA chip. Log only.
      } else if (hoursLeft <= sla.gentleReminderHours) {
        // Sprint 1: Firm reminder email removed (noise). Board overdue/SLA chips are sufficient.
        // Event still logged so analytics can count near-miss tasks.
        if (!todayAlerted[taskId + '|' + REMINDER.FIRM]) {  // S3.2: O(1) lookup
          logReminder_(taskId, ownerEmail, ownerName, REMINDER.FIRM, reminderCount + 1);
          updateReminderFields_(sheet, i + 1, reminderCount + 1, now);
          emitEvent_({
            type: EVENT.REMINDER_SENT, taskId: taskId,
            actorEmail: ownerEmail, actorTeam: row[5],
            notes: 'Approaching deadline (no email): ' + hoursLeft.toFixed(1) + 'h left',
            meta: { reminderType: REMINDER.FIRM }
          });
          sent++;
        }

        // CASE 3: FIRST OPEN — no email. Gentle reminder removed (noise). Log only.
      } else if (reminderCount === 0) {
        // Sprint 1: Gentle reminder email removed. Task is visible on board.
        if (!todayAlerted[taskId + '|' + REMINDER.GENTLE]) {  // S3.2: O(1) lookup
          logReminder_(taskId, ownerEmail, ownerName, REMINDER.GENTLE, 1);
          updateReminderFields_(sheet, i + 1, 1, now);
          emitEvent_({
            type: EVENT.REMINDER_SENT, taskId: taskId,
            actorEmail: ownerEmail, actorTeam: row[5],
            notes: 'Task open (no email sent — board visible)',
            meta: { reminderType: REMINDER.GENTLE }
          });
          sent++;
        }
      }
    }

    return sent;
  } finally {
    lock.releaseLock();
  }
}
// Identifies tasks where (time held / SLA hours) > 0.8.
// Sends one early-warning alert before the task actually breaches.
// Prevents floods: only alerts once per task (uses EventLog check).
// ------------------------------------------------------------------
function checkSLARisk_() {
  var sheet = getSheet(SHEETS.TASKS);
  var data = sheet.getDataRange().getValues();
  var slaConfig = getSLAConfig_();
  var memberMap = getMemberMap_();
  var now = new Date();
  var alerted = 0;

  // Build a set of tasks that already received an AT_RISK alert today
  var atRiskToday = buildAtRiskAlertedSet_();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    if (row[7] === STATUS.DONE || row[7] === STATUS.ARCHIVED || row[7] === STATUS.ON_HOLD) continue;

    var taskId = row[0];
    var taskType = row[2];
    var sla = slaConfig[taskType];
    if (!sla) continue;

    // Time held = now minus LastActionAt (or CreatedAt for first stage)
    var lastAction = row[23] ? new Date(row[23]) : (row[10] ? new Date(row[10]) : now);
    var timeHeld = (now - lastAction) / 3600000;
    var riskRatio = timeHeld / sla.slaHours;

    // Only alert if: risk > 80%, not already Done, not already alerted today
    if (riskRatio < 0.8) continue;
    if (atRiskToday[taskId]) continue;

    var ownerEmail = row[3];
    var ownerMember = memberMap[(ownerEmail || '').toLowerCase()];
    var ownerName = ownerMember ? ownerMember.name : (row[4] || '');

    // Sprint 1: SLA risk email removed — board SLA chip is sufficient, no email needed.

    emitEvent_({
      type: EVENT.SLA_AT_RISK,
      taskId: taskId,
      actorEmail: ownerEmail,
      actorTeam: row[5],
      slaHours: sla.slaHours,
      notes: 'SLA risk: ' + (riskRatio * 100).toFixed(0) + '% of SLA consumed (' + timeHeld.toFixed(1) + 'h / ' + sla.slaHours + 'h)',
      meta: { riskRatio: riskRatio, timeHeld: timeHeld }
    });

    alerted++;
  }

  return alerted;
}

// Reads today's EventLog entries to find tasks already alerted at-risk today
function buildAtRiskAlertedSet_() {
  var set = {};
  var todayStr = new Date().toISOString().substring(0, 10); // 'YYYY-MM-DD'
  try {
    var sheet = getSheet(SHEETS.EVENT_LOG);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] !== EVENT.SLA_AT_RISK) continue;
      if (!data[i][14]) continue;
      var ts = String(data[i][14]);
      if (ts.substring(0, 10) === todayStr) set[data[i][2]] = true;
    }
  } catch (e) { }
  return set;
}

// sendSLARiskEmail_ removed from reminderengine.js — stub in calendarnotifications.js


// ------------------------------------------------------------------
// IDLE TASK DETECTION (NEW — Phase 1)
// Tasks with no state change for > 24h that are not On Hold or Done.
// Alerts the task owner and the owner's manager.
// Uses LastActionAt (col X) — set on every transitionTask_() call.
// ------------------------------------------------------------------
function checkIdleTasks_() {
  var sheet = getSheet(SHEETS.TASKS);
  var data = sheet.getDataRange().getValues();
  var memberMap = getMemberMap_();
  var now = new Date();
  var IDLE_THRESHOLD_HOURS = 24;
  var alerted = 0;

  // Build set of tasks already idle-alerted today
  var idleAlertedToday = buildIdleAlertedSet_();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    if (row[7] === STATUS.DONE || row[7] === STATUS.ARCHIVED || row[7] === STATUS.ON_HOLD) continue;

    var taskId = row[0];
    var lastAction = row[23] ? new Date(row[23]) : (row[10] ? new Date(row[10]) : null);
    if (!lastAction) continue;

    var idleHours = (now - lastAction) / 3600000;
    if (idleHours < IDLE_THRESHOLD_HOURS) continue;
    if (idleAlertedToday[taskId]) continue;

    var ownerEmail = row[3];
    var ownerMember = memberMap[(ownerEmail || '').toLowerCase()];
    var ownerName = ownerMember ? ownerMember.name : (row[4] || '');
    var ownerTeam = row[5];

    // Find manager of this team to also notify
    var managerEmail = getTeamManagerEmail_(ownerTeam, memberMap);

    // Sprint 1: Idle alert emails removed — board idle badge is sufficient.
    // EventLog event still emitted for analytics.

    emitEvent_({
      type: EVENT.TASK_IDLE,
      taskId: taskId,
      actorEmail: ownerEmail,
      actorTeam: ownerTeam,
      notes: 'Task idle for ' + idleHours.toFixed(1) + ' hours',
      meta: { idleHours: idleHours, threshold: IDLE_THRESHOLD_HOURS }
    });

    alerted++;
  }

  return alerted;
}

function buildIdleAlertedSet_() {
  var set = {};
  var todayStr = new Date().toISOString().substring(0, 10);
  try {
    var data = getSheet(SHEETS.EVENT_LOG).getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] !== EVENT.TASK_IDLE) continue;
      if (String(data[i][14]).substring(0, 10) === todayStr) set[data[i][2]] = true;
    }
  } catch (e) { }
  return set;
}

function getTeamManagerEmail_(teamName, memberMap) {
  var members = Object.values(memberMap);
  for (var i = 0; i < members.length; i++) {
    var m = members[i];
    if (m.team === teamName && m.role === 'Manager' && m.active) return m.email;
  }
  // Fall back to Owner
  for (var j = 0; j < members.length; j++) {
    if (members[j].role === 'Owner' && members[j].active) return members[j].email;
  }
  return null;
}

// sendIdleAlertEmail_ removed from reminderengine.js — stub in calendarnotifications.js


// ------------------------------------------------------------------
// WEEKLY DIGEST
// Trigger: Sunday 8am
// Sends a performance summary to the Owner.
// In Phase 1, uses HandoffLog for metrics (AnalyticsEngine is Phase 3).
// ------------------------------------------------------------------
function sendWeeklyDigest() {
  var ownerEmail = getOwnerEmail_();
  if (!ownerEmail) return;

  var data = buildBasicWeeklyMetrics_();
  var html = buildDigestHtml_(data);

  MailApp.sendEmail({
    to: ownerEmail,
    subject: 'TaskFlow — Weekly Summary ' + new Date().toDateString(),
    htmlBody: html
  });
}

function buildBasicWeeklyMetrics_() {
  var sheet = getSheet(SHEETS.TASKS);
  var data = sheet.getDataRange().getValues();
  var memberMap = getMemberMap_();
  var now = new Date();
  var weekAgo = new Date(now.getTime() - 7 * 24 * 3600000);

  var total = 0, completed = 0, breached = 0;
  var totalTAT = 0, completedCount = 0;
  var byPerson = {};  // email → { name, tasks, breaches, tatHours }

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;

    var createdAt = row[10] ? new Date(row[10]) : null;
    if (!createdAt || createdAt < weekAgo) continue;

    total++;
    var ownerEmail = (row[3] || '').toLowerCase();
    var ownerMember = memberMap[ownerEmail];
    var ownerName = ownerMember ? ownerMember.name : (row[4] || ownerEmail);

    if (!byPerson[ownerEmail]) byPerson[ownerEmail] = { name: ownerName, tasks: 0, breaches: 0, tatHours: [] };
    byPerson[ownerEmail].tasks++;
    if (row[18] === true) { breached++; byPerson[ownerEmail].breaches++; }
    if (row[7] === STATUS.DONE || row[7] === STATUS.ARCHIVED) {
      completed++;
      if (row[13]) {
        var tat = Number(row[13]);
        totalTAT += tat;
        completedCount++;
        byPerson[ownerEmail].tatHours.push(tat);
      }
    }
  }

  var avgTAT = completedCount > 0 ? (totalTAT / completedCount).toFixed(1) : 'N/A';
  var onTime = total > 0 ? (((total - breached) / total) * 100).toFixed(0) : 100;

  var leaderboard = Object.values(byPerson)
    .map(function (p) {
      return {
        name: p.name,
        tasksHandled: p.tasks,
        breaches: p.breaches,
        avgTAT: p.tatHours.length ? (p.tatHours.reduce(function (a, b) { return a + b; }, 0) / p.tatHours.length).toFixed(1) : 'N/A'
      };
    })
    .sort(function (a, b) { return a.breaches - b.breaches || b.tasksHandled - a.tasksHandled; });

  return {
    kpis: {
      totalTasks: total, completedTasks: completed,
      escalatedTasks: breached,
      onTimeRate: onTime, breachRate: (100 - Number(onTime)),
      avgTATHours: avgTAT
    },
    leaderboard: leaderboard
  };
}


// ------------------------------------------------------------------
// PRIVATE HELPERS
// ------------------------------------------------------------------

function wasReminderSentToday_(taskId, reminderType) {
  var sheet = getSheet(SHEETS.REMINDERS_LOG);
  var data = sheet.getDataRange().getValues();
  var todayStr = new Date().toISOString().substring(0, 10);

  for (var i = 1; i < data.length; i++) {
    if (data[i][1] !== taskId) continue;
    if (data[i][4] !== reminderType) continue;
    if (!data[i][5]) continue;
    if (new Date(data[i][5]).toISOString().substring(0, 10) === todayStr) return true;
  }
  return false;
}

function logReminder_(taskId, email, name, type, count) {
  getSheet(SHEETS.REMINDERS_LOG).appendRow([
    generateLogId('REM'), taskId, email, name, type, new Date(), count
  ]);
}

function updateReminderFields_(sheet, rowIndex, newCount, timestamp) {
  sheet.getRange(rowIndex, 17).setValue(newCount);
  sheet.getRange(rowIndex, 18).setValue(timestamp);
}

function getOwnerEmail_() {
  try {
    var data = getSheet(SHEETS.TEAM_MEMBERS).getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][4] === 'Owner' && data[i][5] === true) return data[i][2];
    }
  } catch (e) { }
  return Session.getActiveUser().getEmail();
}

// Legacy aliases
function wasReminderSentToday(taskId, type) { return wasReminderSentToday_(taskId, type); }
function logReminder(taskId, email, name, type, count) { logReminder_(taskId, email, name, type, count); }
function updateReminderFields(sheet, rowIndex, count, ts) { updateReminderFields_(sheet, rowIndex, count, ts); }


// ------------------------------------------------------------------
// DIGEST HTML BUILDER (unchanged from v5, kept for compat)
// ------------------------------------------------------------------

function buildDigestHtml_(d) {
  var k = d.kpis || {};
  var ldr = (d.leaderboard || []).slice(0, 5);
  var rows = ldr.map(function (r, i) {
    return '<tr style="border-bottom:1px solid #eee;">' +
      '<td style="padding:8px 12px;">' + (i + 1) + '</td>' +
      '<td style="padding:8px 12px;font-weight:600;">' + r.name + '</td>' +
      '<td style="padding:8px 12px;">' + r.avgTAT + 'h avg</td>' +
      '<td style="padding:8px 12px;">' + r.tasksHandled + ' tasks</td>' +
      '<td style="padding:8px 12px;color:' + (r.breaches > 0 ? '#de350b' : '#00875a') + ';">' + r.breaches + ' breaches</td>' +
      '</tr>';
  }).join('');

  return '<!DOCTYPE html><html><body style="font-family:sans-serif;max-width:600px;margin:0 auto;">' +
    '<div style="background:#0052cc;padding:20px 24px;border-radius:8px 8px 0 0;">' +
    '<h1 style="color:#fff;margin:0;font-size:20px;">TaskFlow Weekly Digest</h1>' +
    '<p style="color:rgba(255,255,255,.7);margin:4px 0 0;font-size:13px;">Week ending ' + new Date().toDateString() + '</p>' +
    '</div><div style="background:#f4f5f7;padding:20px 24px;">' +
    '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:20px;">' +
    kpiBox_('#0052cc', k.totalTasks, 'Total Tasks') +
    kpiBox_('#00875a', k.onTimeRate + '%', 'On-Time Rate') +
    kpiBox_('#de350b', k.breachRate + '%', 'SLA Breaches') +
    kpiBox_('#172b4d', k.avgTATHours + 'h', 'Avg Turnaround') +
    kpiBox_('#00875a', k.completedTasks, 'Completed') +
    kpiBox_('#ff8b00', k.escalatedTasks, 'Escalated') +
    '</div>' +
    '<div style="background:#fff;border-radius:8px;padding:16px;">' +
    '<h3 style="margin:0 0 12px;font-size:14px;">Performance This Week</h3>' +
    '<table style="width:100%;border-collapse:collapse;font-size:13px;">' +
    '<thead><tr style="background:#f4f5f7;"><th style="padding:8px 12px;text-align:left;">Rank</th>' +
    '<th style="padding:8px 12px;text-align:left;">Name</th>' +
    '<th style="padding:8px 12px;text-align:left;">Avg TAT</th>' +
    '<th style="padding:8px 12px;text-align:left;">Tasks</th>' +
    '<th style="padding:8px 12px;text-align:left;">Breaches</th></tr></thead>' +
    '<tbody>' + rows + '</tbody></table></div></div></body></html>';
}

function kpiBox_(color, val, label) {
  return '<div style="background:#fff;border-radius:6px;padding:14px;text-align:center;">' +
    '<div style="font-size:24px;font-weight:700;color:' + color + ';">' + val + '</div>' +
    '<div style="font-size:11px;color:#5e6c84;margin-top:3px;">' + label + '</div></div>';
}

// ══════════════════════════════════════════════════════════
// SPRINT 5 — USER REMINDER ENGINE
// Scans UserReminders sheet every hour.
// Fires emails for rows where: Fired=FALSE AND RemindAt <= now.
// Skips tasks that are already Done or Archived.
// tryLock(0): exits immediately if SLA engine is running.
// ══════════════════════════════════════════════════════════
function runUserReminderEngine_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    console.log('runUserReminderEngine_: lock held, skipping.');
    return 0;
  }
  try {
    var sheet = getSheet(SHEETS.USER_REMINDERS);
    if (!sheet) return 0;

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return 0;

    var now = new Date();
    var sent = 0;

    // Build a task-status lookup once — avoids per-row sheet reads
    var taskSheet = getSheet(SHEETS.TASKS);
    var taskData = taskSheet.getDataRange().getValues();
    var taskStatusMap = {};  // taskId → status
    for (var t = 1; t < taskData.length; t++) {
      if (taskData[t][0]) taskStatusMap[taskData[t][0]] = taskData[t][COL.STATUS] || '';
    }

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;  // empty row
      if (row[7] === true) continue;  // already fired

      var remindAt = row[4] ? new Date(row[4]) : null;
      if (!remindAt || remindAt > now) continue;  // not yet time

      var taskId = row[1];
      var taskName = row[2] || taskId;
      var recipStr = row[3] || '';
      var note = row[8] || '';
      var eventId = row[9] || '';

      // Skip if task is already closed — no point reminding
      var tStatus = taskStatusMap[taskId] || '';
      if (tStatus === STATUS.DONE || tStatus === STATUS.ARCHIVED) {
        // Mark fired so engine never rechecks it
        sheet.getRange(i + 1, 8).setValue(true);
        continue;
      }

      // Calendar-based reminders: rely on Calendar email and avoid TaskFlow email
      if (eventId) {
        sheet.getRange(i + 1, 8).setValue(true);
        emitEvent_({
          type: EVENT.USER_REMINDER_SENT,
          taskId: taskId,
          actorEmail: row[5],  // CreatedBy
          notes: 'User reminder fired via Calendar.',
          meta: { recipients: recipStr, taskName: taskName, calendarEventId: eventId, channel: 'Calendar' }
        });
        sent++;
        continue;
      }

      // Parse recipients — semicolon-delimited, trim, filter valid emails
      var recipients = recipStr.split(';')
        .map(function (e) { return e.trim(); })
        .filter(function (e) { return e.indexOf('@') > 0; });

      if (!recipients.length) {
        sheet.getRange(i + 1, 8).setValue(true);  // nothing to send, mark done
        continue;
      }

      // Send one email per recipient
      var allSent = true;
      for (var r = 0; r < recipients.length; r++) {
        try {
          sendUserReminderEmail_(recipients[r], taskId, taskName, remindAt, note);
        } catch (emailErr) {
          console.warn('sendUserReminderEmail_ failed for ' + recipients[r] + ': ' + emailErr.message);
          allSent = false;
        }
      }

      // Mark fired regardless — partial send is acceptable; retrying would double-email
      sheet.getRange(i + 1, 8).setValue(true);

      emitEvent_({
        type: EVENT.USER_REMINDER_SENT,
        taskId: taskId,
        actorEmail: row[5],  // CreatedBy
        notes: 'User reminder fired to: ' + recipStr,
        meta: { recipients: recipStr, taskName: taskName }
      });

      sent++;
    }

    return sent;

  } finally {
    lock.releaseLock();
  }
}


// ══════════════════════════════════════════════════════════
// SPRINT 5 — CLEANUP OLD REMINDERS
// Runs once daily at midnight (gated in runHourlyEngine).
// Rewrites sheet in 3 API calls regardless of rows deleted:
//   1. getDataRange().getValues()  — read all
//   2. clearContents()             — wipe data rows
//   3. setValues()                 — write kept rows
// 60-day window keeps data covering the analytics lookback period.
// ══════════════════════════════════════════════════════════
function cleanupOldReminders_() {
  var sheet = getSheet(SHEETS.USER_REMINDERS);
  if (!sheet) return 0;

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return 0;

  var cutoff = new Date(Date.now() - 60 * 24 * 3600000);  // 60 days ago
  var header = [data[0]];  // always keep header row
  var kept = [];
  var deleted = 0;
  var colCount = sheet.getLastColumn();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;  // skip empty rows — don't keep blank garbage

    var isFired = row[7] === true;
    var createdAt = row[6] ? new Date(row[6]) : null;
    var isOld = createdAt && createdAt < cutoff;

    if (isFired && isOld) {
      deleted++;
    } else {
      kept.push(row);
    }
  }

  if (!deleted) return 0;  // nothing to do — skip API calls

  // Rewrite sheet: header + kept rows (3 API calls total)
  var allRows = header.concat(kept);
  var lastRow = sheet.getLastRow();

  // Clear all data rows (row 2 onwards)
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, colCount).clearContent();
  }

  // Write kept rows back (only if there are any)
  if (kept.length > 0) {
    sheet.getRange(2, 1, kept.length, colCount).setValues(kept);
  }

  console.log('cleanupOldReminders_: deleted ' + deleted + ', kept ' + kept.length);
  return deleted;
}
