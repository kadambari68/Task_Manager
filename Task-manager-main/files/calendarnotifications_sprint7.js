// ============================================================
// CalendarAndNotifications.gs — TaskFlow v6
// SPRINT 1 REWRITE
//
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// WHY THE CALENDAR FEATURE WAS REMOVED
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// The previous code called:
//   calendar.createEvent(title, start, end, { guests: email, sendInvites: true })
//
// Google Workspace (unlike free Gmail) automatically attaches a
// Google Meet link to ANY calendar event that has guest invitees.
// This is a Workspace org-level setting and CANNOT be overridden
// from Apps Script — the Meet URL is injected by Google's backend.
//
// The result: every task with SLA > 24 hours sent the assignee
// an unexpected Google Meet calendar invite. The "sudden Google Meet
// message" was this invite. If the event was for a task deadline
// 30 min block, the Meet link was real and joinable — causing confusion.
//
// Additionally, the calendar transfer on routing was broken (Bug 4):
// the old owner's calendar event was never removed (CalendarApp has
// no remove-guest API), leaving phantom deadlines on old owners' calendars.
//
// DECISION: Entire calendar integration removed. SLA deadlines are
// tracked in the Tasks sheet (col L) and surfaced via the board's
// overdue chip and email reminders — no calendar noise needed.
//
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// EMAIL REDUCTION — SPRINT 1
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// REMOVED (all caused inbox noise with no action value):
//   × Gentle Reminder    — board shows overdue chip already
//   × Firm Reminder      — consolidated into escalation email
//   × SLA Risk Alert     — board shows SLA % chip already
//   × Idle Alert         — board shows idle badge already
//
// KEPT (direct action required, no other channel):
//   ✓ Task Assignment    — new owner must know they have a task
//   ✓ Task Routed        — new owner must know task was passed to them
//   ✓ Task Completion    — creator confirmation (now fixed — was broken)
//   ✓ SLA Breach Escalation — manager/owner must intervene
//   ✓ Weekly Digest      — owned by ReminderEngine, not changed here
//     Firm Reminder      — consolidated into escalation email
//
// All emails upgraded from plain text to branded HTML.
// ============================================================


// ──────────────────────────────────────────────────────────
// CALENDAR STUBS — safe no-ops so TaskEngine calls don't crash
// ──────────────────────────────────────────────────────────
// These replace the deleted createCalendarEvent / transferCalendarEvent.
// TaskEngine calls them inside try/catch so they're already non-blocking,
// but having stubs avoids "function not defined" errors during transition.

function createCalendarEvent(taskId, taskName, assigneeEmail, deadline) {
  // REMOVED: was creating Google Calendar events with Meet links.
  // See file header for full explanation.
  // No-op stub kept for backward compatibility with TaskEngine calls.
}

function transferCalendarEvent(taskId, newOwnerEmail, newDeadline) {
  // REMOVED: calendar transfer also removed.
  // No-op stub.
}

function storeCalendarEventId(taskId, eventId) {
  // No-op stub.
}

function getCalendarEventId(taskId) {
  return null;
}


// ──────────────────────────────────────────────────────────
// SHARED EMAIL HELPERS
// ──────────────────────────────────────────────────────────

var BRAND = {
  name: 'TaskFlow',
  blue: '#0052CC',
  dark: '#172B4D',
  green: '#00875A',
  red: '#DE350B',
  orange: '#FF8B00',
  grey: '#5E6C84',
  lightBg: '#F4F5F7'
};

/**
 * Builds the common HTML email shell.
 * @param {string} accentColor  — hex color for the header bar
 * @param {string} label        — short uppercase label in header (e.g. "NEW TASK")
 * @param {string} bodyHtml     — inner HTML content
 */
function buildEmailHtml_(accentColor, label, bodyHtml) {
  return '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>' +
    'body{margin:0;padding:0;background:#F4F5F7;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;}' +
    '.wrap{max-width:560px;margin:32px auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.12);}' +
    '.hdr{background:' + accentColor + ';padding:20px 28px;}' +
    '.hdr-brand{color:rgba(255,255,255,.7);font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;}' +
    '.hdr-label{color:#fff;font-size:20px;font-weight:700;margin-top:4px;}' +
    '.body{padding:28px;}' +
    '.greeting{font-size:15px;color:' + BRAND.dark + ';margin-bottom:16px;}' +
    '.card{background:' + BRAND.lightBg + ';border-radius:6px;padding:16px 20px;margin:16px 0;}' +
    '.card-row{display:flex;justify-content:space-between;align-items:flex-start;padding:6px 0;border-bottom:1px solid #DFE1E6;}' +
    '.card-row:last-child{border-bottom:none;}' +
    '.card-key{font-size:11px;font-weight:700;color:' + BRAND.grey + ';text-transform:uppercase;letter-spacing:.05em;min-width:110px;}' +
    '.card-val{font-size:13px;color:' + BRAND.dark + ';font-weight:500;text-align:right;}' +
    '.btn{display:inline-block;background:' + accentColor + ';color:#fff;text-decoration:none;padding:11px 24px;border-radius:4px;font-size:13px;font-weight:700;margin:20px 0;}' +
    '.note{font-size:12px;color:' + BRAND.grey + ';margin-top:20px;line-height:1.6;}' +
    '.footer{background:' + BRAND.lightBg + ';padding:14px 28px;font-size:11px;color:' + BRAND.grey + ';border-top:1px solid #DFE1E6;}' +
    '</style></head><body>' +
    '<div class="wrap">' +
    '<div class="hdr">' +
    '<div class="hdr-brand">' + BRAND.name + '</div>' +
    '<div class="hdr-label">' + label + '</div>' +
    '</div>' +
    '<div class="body">' + bodyHtml +
    '<div style="margin-top:24px;"><a href="' + ScriptApp.getService().getUrl() + '" class="btn" style="text-decoration:none; color:#ffffff !important; display:inline-block; background:' + accentColor + '; padding:11px 24px; border-radius:4px; font-weight:700;"><span style="color:#ffffff;">Open TaskFlow</span></a></div>' +
    '</div>' +
    '<div class="footer">This is an automated message from TaskFlow &mdash; do not reply.</div>' +
    '</div></body></html>';
}

/** Single card row helper */
function row_(key, val) {
  return '<div class="card-row"><span class="card-key">' + key + '</span><span class="card-val">' + (val || '—') + '</span></div>';
}

/** Format deadline date cleanly */
function fmtDate_(d) {
  if (!d) return '—';
  try { return new Date(d).toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' }); }
  catch (e) { return String(d); }
}


// ──────────────────────────────────────────────────────────
// 1. TASK ASSIGNMENT EMAIL
//    Sent: createTask() when a task is first created
//    Recipient: assignee
// ──────────────────────────────────────────────────────────
function sendTaskAssignmentEmail(assignee, taskId, taskName, createdBy, deadline) {
  try {
    var subject = '[TaskFlow] New Task Assigned — ' + taskName;
    subject = '[TaskFlow] Task Created - ' + taskName;
    var body = buildEmailHtml_(BRAND.blue, 'Task Created',
      '<p class="greeting">Hi ' + (assignee.name || 'there') + ',</p>' +
      '<p style="font-size:14px;color:' + BRAND.dark + ';">A new task has been created for you and is awaiting your action.</p>' +
      '<div class="card">' +
      row_('Task ID', taskId) +
      row_('Task', taskName) +
      row_('Created by', createdBy) +
      row_('Deadline', fmtDate_(deadline)) +
      '</div>' +
      '<p class="note">Please log in to TaskFlow to view the full task details and begin work.</p>'
    );
    GmailApp.sendEmail(assignee.email, subject, 'Task created: ' + taskName, { htmlBody: body });
  } catch (e) { console.warn('sendTaskAssignmentEmail failed: ' + e.message); }
}


// ──────────────────────────────────────────────────────────
// 2. TASK ROUTED EMAIL
//    Sent: routeTask() when a task changes hands
//    Recipient: new assignee
// ──────────────────────────────────────────────────────────
function sendTaskRoutedEmail(newAssignee, taskId, taskName, routedBy, deadline) {
  try {
    var subject = '[TaskFlow] Task Routed to You — ' + taskName;
    var body = buildEmailHtml_(BRAND.blue, 'Task Routed to You',
      '<p class="greeting">Hi ' + (newAssignee.name || 'there') + ',</p>' +
      '<p style="font-size:14px;color:' + BRAND.dark + ';">A task has been routed to you by <strong>' + routedBy + '</strong> and requires your attention.</p>' +
      '<div class="card">' +
      row_('Task ID', taskId) +
      row_('Task', taskName) +
      row_('Routed by', routedBy) +
      row_('Deadline', fmtDate_(deadline)) +
      '</div>' +
      '<p class="note">Log in to TaskFlow to pick up this task and mark it In Progress.</p>'
    );
    GmailApp.sendEmail(newAssignee.email, subject, 'Task routed to you: ' + taskName, { htmlBody: body });
  } catch (e) { console.warn('sendTaskRoutedEmail failed: ' + e.message); }
}


// ──────────────────────────────────────────────────────────
// 3. TASK COMPLETION EMAIL
//    Sent: completeTask() on successful completion
//    Recipient: task creator (confirmation)
// ──────────────────────────────────────────────────────────
function sendTaskCompletedEmail(task, completedBy, completedAt, totalHours) {
  try {
    var subject = '[TaskFlow] Task Completed — ' + task.taskName;
    var hrs = Number(totalHours || 0).toFixed(1);
    var body = buildEmailHtml_(BRAND.green, 'Task Completed',
      '<p class="greeting">Hi ' + (task.createdBy || 'there') + ',</p>' +
      '<p style="font-size:14px;color:' + BRAND.dark + ';">A task you created has been marked as complete.</p>' +
      '<div class="card">' +
      row_('Task ID', task.taskId) +
      row_('Task', task.taskName) +
      row_('Completed by', completedBy) +
      row_('Completed at', fmtDate_(completedAt)) +
      row_('Duration', hrs + ' hours') +
      '</div>'
    );
    GmailApp.sendEmail(task.createdByEmail, subject, 'Task completed: ' + task.taskName, { htmlBody: body });
  } catch (e) { console.warn('sendTaskCompletedEmail failed: ' + e.message); }
}


// ──────────────────────────────────────────────────────────
// 4. COMMENT NOTIFICATION EMAIL
//    Sent: addComment() when someone replies to a task thread
//    Recipients: current owner, creator, etc.
// ──────────────────────────────────────────────────────────
function sendCommentNotificationEmail(recipient, taskName, taskId, actorName, commentText) {
  try {
    var subject = '[TaskFlow] New Comment on: ' + taskName;
    var body = buildEmailHtml_(BRAND.blue, 'New Comment',
      '<p class="greeting">Hi ' + (recipient.name || 'there') + ',</p>' +
      '<p style="font-size:14px;color:' + BRAND.dark + ';"><strong>' + actorName + '</strong> added a new comment to a task you are involved with:</p>' +
      '<div class="card" style="font-style:italic;color:' + BRAND.dark + ';">"' + String(commentText).replace(/\n/g, '<br>') + '"</div>' +
      '<div class="card">' +
      row_('Task ID', taskId) +
      row_('Task', taskName) +
      '</div>' +
      '<p class="note">Log in to TaskFlow to view the full thread and reply.</p>'
    );
    GmailApp.sendEmail(recipient.email, subject, 'New comment on: ' + taskName, { htmlBody: body });
  } catch (e) {
    console.warn('sendCommentNotificationEmail failed: ' + e.message);
  }
}

// ──────────────────────────────────────────────────────────
// 5. SLA BREACH ESCALATION EMAIL
//    Sent: runDailyReminderEngine_() hourly, once per day per task
//    Recipient: escalationEmail in SLAConfig (manager / owner)
//    Also sends a direct alert to the task owner
// ──────────────────────────────────────────────────────────
function sendEscalationEmail(escalationEmail, ownerName, ownerEmail, taskId, taskName, hoursOverdue) {
  try {
    // WhatsApp follow-up link if phone is stored
    var owner = getMemberByEmail(ownerEmail);
    var phone = owner && owner.phone ? String(owner.phone).replace(/\D/g, '') : '';
    var waBtn = phone
      ? '<a class="btn" style="background:' + BRAND.orange + ';text-decoration:none;" href="https://wa.me/' + phone +
      '?text=' + encodeURIComponent('Hi ' + ownerName + ', following up on overdue task ' + taskId + ': ' + taskName) +
      '">WhatsApp ' + ownerName + '</a>'
      : '';

    var subject = '[TaskFlow] SLA Breach — ' + taskName + ' (' + taskId + ')';
    var body = buildEmailHtml_(BRAND.red, 'SLA Breach — Action Required',
      '<p class="greeting">This task has exceeded its SLA deadline and requires your immediate attention.</p>' +
      '<div class="card">' +
      row_('Task ID', taskId) +
      row_('Task', taskName) +
      row_('Assigned to', ownerName + ' (' + ownerEmail + ')') +
      row_('Hours overdue', hoursOverdue + 'h') +
      '</div>' +
      (waBtn ? '<p>' + waBtn + '</p>' : '') +
      '<p class="note">Please follow up with the assigned team member directly. You can view and reassign the task in TaskFlow.</p>'
    );
    GmailApp.sendEmail(escalationEmail, subject, 'SLA breach: ' + taskName, { htmlBody: body });
  } catch (e) { console.warn('sendEscalationEmail failed: ' + e.message); }
}

// Direct overdue alert sent to the task owner alongside the manager escalation
function sendOverdueOwnerEmail_(ownerEmail, ownerName, taskId, taskName, hoursOverdue) {
  try {
    var subject = '[TaskFlow] Your Task Is Overdue — ' + taskName;
    var body = buildEmailHtml_(BRAND.red, 'Task Overdue',
      '<p class="greeting">Hi ' + ownerName + ',</p>' +
      '<p style="font-size:14px;color:' + BRAND.dark + ';">The task below is <strong>' + hoursOverdue + ' hours overdue</strong>. Your manager has been notified.</p>' +
      '<div class="card">' +
      row_('Task ID', taskId) +
      row_('Task', taskName) +
      row_('Overdue', hoursOverdue + 'h') +
      '</div>' +
      '<p class="note">Please log in to TaskFlow immediately to update the status or contact your manager.</p>'
    );
    GmailApp.sendEmail(ownerEmail, subject, 'Your task is overdue: ' + taskName, { htmlBody: body });
  } catch (e) { console.warn('sendOverdueOwnerEmail_ failed: ' + e.message); }
}


// ──────────────────────────────────────────────────────────
// REMOVED EMAIL STUBS
// These are kept as no-ops so any legacy calls in old trigger
// functions don't throw "function not defined" errors.
// All were generating inbox noise without actionable value —
// the board's overdue chips and SLA badges show the same info.
// ──────────────────────────────────────────────────────────

function sendGentleReminderEmail(email, name, taskId, taskName, hoursLeft) {
  // REMOVED Sprint 1 — board shows overdue chip. No email sent.
}

// ------------------------------------------------------------------
// FIRM REMINDER EMAIL — STUBBED Sprint 5
// sendOverdueOwnerEmail_ (above) covers the owner-direct overdue alert
// with proper branded HTML. This plain-text version referenced SIGNATURE
// (undefined) and would have thrown a ReferenceError on execution.
// Kept as a no-op so reminderengine.js call on line 117 does not crash.
// ------------------------------------------------------------------
function sendFirmReminderEmail(email, name, taskId, taskName, hoursValue, isOverdue) {
  // No-op — sendOverdueOwnerEmail_ handles the owner overdue alert.
  // sendEscalationEmail handles the manager escalation alert.
}

function sendSLARiskEmail_(toEmail, toName, taskId, taskName, timeHeld, slaHours) {
  // REMOVED — board shows SLA risk chip. Not sending.
}

function sendIdleAlertEmail_(toEmail, toName, taskId, taskName, idleHours) {
  // REMOVED — board shows idle badge. Not sending.
}


// ──────────────────────────────────────────────────────────
// SPRINT 5 — USER REMINDER EMAIL
// Sent by runUserReminderEngine_() when a user-set reminder fires.
// Recipient: each email in the semicolon-delimited recipients list.
// Uses same buildEmailHtml_ template as all other system emails.
// ──────────────────────────────────────────────────────────
function sendUserReminderEmail_(toEmail, taskId, taskName, remindAt, note) {
  try {
    var subject = '[TaskFlow] Reminder: ' + taskName;
    var timeStr = fmtDate_(remindAt);
    var body = buildEmailHtml_(BRAND.orange, 'Task Reminder',
      '<p class="greeting">Hi,</p>' +
      '<p style="font-size:14px;color:' + BRAND.dark + ';">This is a reminder for the following task.</p>' +
      '<div class="card">' +
      row_('Task ID', taskId) +
      row_('Task', taskName) +
      row_('Reminder', timeStr) +
      (note ? row_('Note', note) : '') +
      '</div>' +
      '<p class="note">Log in to TaskFlow to view the full task details and update its status.</p>'
    );
    GmailApp.sendEmail(toEmail, subject, 'Reminder: ' + taskName, { htmlBody: body });
  } catch (e) {
    console.warn('sendUserReminderEmail_ failed for ' + toEmail + ': ' + e.message);
  }
}

// ──────────────────────────────────────────────────────────
// SPRINT 6 — RECURRING TASK REMINDER EMAIL
// Sent by runRecurringReminderEngine_() on each cycle.
// ──────────────────────────────────────────────────────────
function sendRecurringReminderEmail_(toEmail, toName, recId, title, description) {
  try {
    var subject = '[TaskFlow] Recurring Task Due: ' + title;
    var body = buildEmailHtml_(BRAND.orange, 'Recurring Task Reminder',
      '<p class="greeting">Hi ' + (toName || 'there') + ',</p>' +
      '<p style="font-size:14px;color:' + BRAND.dark + ';">This is your scheduled reminder for the following recurring task.</p>' +
      '<div class="card">' +
      row_('Task', title) +
      row_('Ref', recId) +
      (description ? row_('Details', description) : '') +
      '</div>' +
      '<p class="note">Log in to TaskFlow to update your progress or contact your manager if needed.</p>'
    );
    GmailApp.sendEmail(toEmail, subject, 'Recurring task reminder: ' + title, { htmlBody: body });
  } catch (e) {
    console.warn('sendRecurringReminderEmail_ failed for ' + toEmail + ': ' + e.message);
  }
}
