// ============================================================
// 10_BackendAddons.gs — TaskFlow v6  (DEPLOY ORDER: 8th, after all engines)
//
// Contains ONLY functions unique to this file:
//   • setupAddons()        one-time sheet scaffolding
//   • File upload helpers  Google Drive attachment store
//   • Comment thread       addComment / getComments
//   • getMemberPhone_()    thin wrapper used by idle alert
//   • getDateDaysAgo_()    date util (used locally in weekly digest)
//
// REMOVED from this file (now canonical in their engine files):
//   getMemberByEmail_()   → Code_v6.gs
//   getSLAConfigRows()    → Code_v6.gs
//   getMyCreatedTasks()   → Code_v6.gs
//   getLastHandoffTime()  → Code_v6.gs
//   saveSLAConfig()       → AdminEngine_v6.gs
//   getIdleTasksAlert()   → AdminEngine_v6.gs
//   updateMemberPhone()   → AdminEngine_v6.gs
//   addMemberWithPhone()  → AdminEngine_v6.gs (deprecated, use addMember)
//   sendWeeklyDigest()    → ReminderEngine_v6.gs
//   getOwnerEmail_()      → ReminderEngine_v6.gs
//   buildDigestHtml_()    → ReminderEngine_v6.gs
//   kpiBox_()             → ReminderEngine_v6.gs
// ============================================================


// ──────────────────────────────────────────────────────────
// ONE-TIME SETUP — run manually once from Apps Script editor
// Creates any sheets not yet present and seeds default data
// ──────────────────────────────────────────────────────────
function setupAddons() {
  try {
    requireRole_(['Owner']);
    var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!id) return err_('SYSTEM_ERROR');
    var ss = SpreadsheetApp.openById(id);

    // Comments sheet
    if (!ss.getSheetByName('Comments')) {
      var cs = ss.insertSheet('Comments');
      cs.appendRow(['CommentID', 'TaskID', 'AuthorEmail', 'AuthorName', 'Text', 'Timestamp', 'AttachmentURLs']);
      cs.getRange(1, 1, 1, 7).setFontWeight('bold');
      cs.setFrozenRows(1);
    }

    if (!ss.getSheetByName('Attachments')) {
      var ats = ss.insertSheet('Attachments');
      ats.appendRow(['AttachmentID', 'TaskID', 'DriveFileID', 'DriveURL', 'FileName', 'MimeType', 'SizeBytes', 'Checksum', 'Status', 'UploadedByEmail', 'UploadedByName', 'UploadedAt', 'SourceFlow', 'SourceRef', 'IsDeleted', 'DeletedAt', 'DeletedBy']);
      ats.getRange(1, 1, 1, 17).setFontWeight('bold');
      ats.setFrozenRows(1);
    }

    // SLAConfig seed row if empty
    var sla = ss.getSheetByName('SLAConfig');
    if (sla && sla.getLastRow() <= 1) {
      sla.appendRow(['General', 'All', 'All', '24', '6', '']);
    }

    // EventLog sheet (also created lazily by emitEvent_, but good to pre-create)
    if (!ss.getSheetByName('EventLog')) {
      var el = ss.insertSheet('EventLog');
      el.appendRow(['EventID', 'EventType', 'TaskID', 'ProjectID', 'ActorEmail', 'ActorTeam',
        'TargetEmail', 'TargetTeam', 'FromStatus', 'ToStatus',
        'TimeSpentHours', 'SLAHours', 'SLABreached', 'Notes', 'Timestamp', 'Meta']);
      el.getRange(1, 1, 1, 16).setFontWeight('bold');
      el.setFrozenRows(1);
    }

    // RecurringTasks sheet
    if (typeof setupRecurringTasksSheet === 'function') {
      setupRecurringTasksSheet();
    }

    Logger.log('Addon setup complete.');
    return ok_({ message: 'Setup complete.' });
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('setupAddons: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

function setupAttachmentsSheet() {
  try {
    requireRole_(['Owner']);
    ensureAttachmentSheet_();
    return ok_({ message: 'Attachments sheet is ready.' });
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('setupAttachmentsSheet: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

function setupAttachmentsSystem() {
  try {
    requireRole_(['Owner']);
    ensureAttachmentSheet_();
    var folder = getOrCreateAttachmentRoot_();
    return ok_({
      message: 'Attachment system is ready.',
      folderId: folder.getId(),
      folderUrl: folder.getUrl()
    });
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('setupAttachmentsSystem: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}


// ══════════════════════════════════════════════════════════
// FILE UPLOADS — Google Drive attachment store
// Files are stored under TaskFlow Attachments/<TaskID>/
// Folder ID is cached in Script Properties after first run.
// ══════════════════════════════════════════════════════════

var ATT_COL = {
  ATTACHMENT_ID: 0,
  TASK_ID: 1,
  DRIVE_FILE_ID: 2,
  DRIVE_URL: 3,
  FILE_NAME: 4,
  MIME_TYPE: 5,
  SIZE_BYTES: 6,
  CHECKSUM: 7,
  STATUS: 8,
  UPLOADED_BY_EMAIL: 9,
  UPLOADED_BY_NAME: 10,
  UPLOADED_AT: 11,
  SOURCE_FLOW: 12,
  SOURCE_REF: 13,
  IS_DELETED: 14,
  DELETED_AT: 15,
  DELETED_BY: 16,
  TOTAL_COLS: 17
};

var ATT_STATUS = {
  ACTIVE: 'ACTIVE',
  DELETED: 'DELETED'
};

var ALLOWED_ATTACHMENT_TYPES = [
  'image/jpeg', 'image/png', 'image/gif', 'image/webp',
  'application/pdf',
  'application/msword',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.ms-excel',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-powerpoint',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  'text/plain', 'text/csv',
  'application/zip', 'application/x-zip-compressed'
];

var MAX_ATTACHMENT_BYTES = 10 * 1024 * 1024; // 10 MB
var MAX_ATTACHMENTS_PER_TASK = 25;

function ensureAttachmentSheet_() {
  var sh = getSheet(SHEETS.ATTACHMENTS);
  if (sh) return sh;

  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('SPREADSHEET_ID not set in Script Properties');
  var ss = SpreadsheetApp.openById(id);
  sh = ss.getSheetByName(SHEETS.ATTACHMENTS) || ss.insertSheet(SHEETS.ATTACHMENTS);

  if (sh.getLastRow() === 0) {
    sh.appendRow(['AttachmentID', 'TaskID', 'DriveFileID', 'DriveURL', 'FileName', 'MimeType', 'SizeBytes', 'Checksum', 'Status', 'UploadedByEmail', 'UploadedByName', 'UploadedAt', 'SourceFlow', 'SourceRef', 'IsDeleted', 'DeletedAt', 'DeletedBy']);
    sh.getRange(1, 1, 1, ATT_COL.TOTAL_COLS).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function getOrCreateAttachmentRoot_() {
  var props = PropertiesService.getScriptProperties();
  var fid = props.getProperty('ATTACHMENT_FOLDER_ID');
  if (fid) {
    try { return DriveApp.getFolderById(fid); } catch (e) { /* folder deleted */ }
  }
  var folder = DriveApp.createFolder('TaskFlow Attachments');
  try { folder.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { }
  props.setProperty('ATTACHMENT_FOLDER_ID', folder.getId());
  return folder;
}

function getOrCreateTaskFolder_(taskId) {
  var root = getOrCreateAttachmentRoot_();
  var it = root.getFoldersByName(taskId);
  return it.hasNext() ? it.next() : root.createFolder(taskId);
}

function generateAttachmentId_() {
  return 'ATT-' + Date.now() + '-' + Utilities.getUuid().substring(0, 8);
}

function normalizeAttachmentName_(name) {
  return String(name || 'file')
    .replace(/[\\/:*?"<>|]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 120) || 'file';
}

function bytesToHex_(bytes) {
  var hex = [];
  for (var i = 0; i < bytes.length; i++) {
    var v = (bytes[i] + 256) % 256;
    hex.push((v < 16 ? '0' : '') + v.toString(16));
  }
  return hex.join('');
}

function computeAttachmentChecksum_(decodedBytes) {
  return bytesToHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, decodedBytes));
}

function parseAttachmentRow_(row) {
  return {
    attachmentId: row[ATT_COL.ATTACHMENT_ID] || '',
    taskId: row[ATT_COL.TASK_ID] || '',
    driveFileId: row[ATT_COL.DRIVE_FILE_ID] || '',
    driveUrl: row[ATT_COL.DRIVE_URL] || '',
    url: row[ATT_COL.DRIVE_URL] || '',
    fileName: row[ATT_COL.FILE_NAME] || '',
    name: row[ATT_COL.FILE_NAME] || '',
    mimeType: row[ATT_COL.MIME_TYPE] || '',
    sizeBytes: Number(row[ATT_COL.SIZE_BYTES]) || 0,
    checksum: row[ATT_COL.CHECKSUM] || '',
    status: row[ATT_COL.STATUS] || ATT_STATUS.ACTIVE,
    uploadedByEmail: row[ATT_COL.UPLOADED_BY_EMAIL] || '',
    uploadedByName: row[ATT_COL.UPLOADED_BY_NAME] || '',
    uploadedAt: row[ATT_COL.UPLOADED_AT] ? new Date(row[ATT_COL.UPLOADED_AT]).toISOString() : '',
    sourceFlow: row[ATT_COL.SOURCE_FLOW] || 'task',
    sourceRef: row[ATT_COL.SOURCE_REF] || '',
    isDeleted: row[ATT_COL.IS_DELETED] === true || row[ATT_COL.STATUS] === ATT_STATUS.DELETED,
    deletedAt: row[ATT_COL.DELETED_AT] ? new Date(row[ATT_COL.DELETED_AT]).toISOString() : '',
    deletedBy: row[ATT_COL.DELETED_BY] || '',
    isImage: String(row[ATT_COL.MIME_TYPE] || '').indexOf('image/') === 0
  };
}

function getAttachmentRowsForTask_(taskId) {
  var sh = ensureAttachmentSheet_();
  if (sh.getLastRow() < 2) return [];
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][ATT_COL.TASK_ID] !== taskId) continue;
    out.push({ rowIndex: i + 1, data: data[i] });
  }
  return out;
}

function listTaskAttachments_(taskId, includeLegacyFallback) {
  var rows = getAttachmentRowsForTask_(taskId).map(function (r) { return parseAttachmentRow_(r.data); })
    .filter(function (r) { return !r.isDeleted && r.status === ATT_STATUS.ACTIVE; })
    .sort(function (a, b) { return new Date(b.uploadedAt || 0) - new Date(a.uploadedAt || 0); });

  if (includeLegacyFallback !== false) {
    var tr = findTaskRow_(taskId);
    var legacyUrl = tr && tr.data ? String(tr.data[COL.DRIVE_URL] || '') : '';
    if (legacyUrl) {
      var alreadyListed = rows.some(function (r) { return r.driveUrl === legacyUrl; });
      if (!alreadyListed) {
        rows.push({
          attachmentId: 'LEGACY-' + taskId,
          taskId: taskId,
          driveFileId: '',
          driveUrl: legacyUrl,
          url: legacyUrl,
          fileName: 'Existing File',
          name: 'Existing File',
          mimeType: '',
          sizeBytes: 0,
          checksum: '',
          status: ATT_STATUS.ACTIVE,
          uploadedByEmail: tr.data[COL.CREATOR_EMAIL] || '',
          uploadedByName: tr.data[COL.CREATED_BY] || '',
          uploadedAt: tr.data[COL.LAST_ACTION] ? new Date(tr.data[COL.LAST_ACTION]).toISOString() : (tr.data[COL.CREATED_AT] ? new Date(tr.data[COL.CREATED_AT]).toISOString() : ''),
          sourceFlow: 'legacy',
          sourceRef: '',
          isDeleted: false,
          deletedAt: '',
          deletedBy: '',
          isImage: false
        });
      }
    }
  }

  return rows;
}

function saveTaskAttachments_(taskId, filesData, sourceFlow, sourceRef, actor) {
  sourceFlow = sourceFlow || 'task';
  sourceRef = sourceRef || '';

  var tr = findTaskRow_(taskId);
  if (!tr) return { error: 'NOT_FOUND' };
  if (!canActorAccessTaskRow_(actor, tr.data)) return { error: 'UNAUTHORIZED' };

  var sheet = ensureAttachmentSheet_();
  var existing = getAttachmentRowsForTask_(taskId).map(function (r) { return parseAttachmentRow_(r.data); });
  var activeCount = existing.filter(function (r) { return !r.isDeleted && r.status === ATT_STATUS.ACTIVE; }).length;
  var checksumMap = {};
  existing.forEach(function (r) {
    if (!r.isDeleted && r.status === ATT_STATUS.ACTIVE && r.checksum) checksumMap[r.checksum] = r;
  });

  var rowsToAppend = [];
  var uploaded = [];
  var skippedDuplicates = [];
  var failed = [];
  var folder = null;

  (filesData || []).forEach(function (f) {
    var cleanName = normalizeAttachmentName_(f && f.name);
    if (!f || !f.base64) {
      failed.push({ fileName: cleanName || 'file', reason: 'Missing file data.' });
      return;
    }
    if (activeCount + uploaded.length >= MAX_ATTACHMENTS_PER_TASK) {
      failed.push({ fileName: cleanName, reason: 'Task attachment limit reached.' });
      return;
    }
    if (ALLOWED_ATTACHMENT_TYPES.indexOf(f.mimeType || '') === -1) {
      failed.push({ fileName: cleanName, reason: 'File type not allowed.' });
      return;
    }

    try {
      var decoded = Utilities.base64Decode(f.base64);
      var sizeBytes = decoded.length;
      if (sizeBytes > MAX_ATTACHMENT_BYTES) {
        failed.push({ fileName: cleanName, reason: 'File exceeds 10 MB.' });
        return;
      }

      var checksum = computeAttachmentChecksum_(decoded);
      if (checksum && checksumMap[checksum]) {
        skippedDuplicates.push({
          fileName: cleanName,
          attachmentId: checksumMap[checksum].attachmentId,
          driveUrl: checksumMap[checksum].driveUrl,
          reason: 'Duplicate already attached to this task.'
        });
        emitEvent_({
          type: EVENT.ATTACHMENT_DUPLICATE_SKIPPED,
          taskId: taskId,
          projectId: tr.data[COL.PROJECT_ID] || '',
          actorEmail: actor.email,
          actorTeam: actor.team || '',
          notes: cleanName,
          meta: { sourceFlow: sourceFlow, checksum: checksum }
        });
        return;
      }

      if (!folder) folder = getOrCreateTaskFolder_(taskId);
      var attachmentId = generateAttachmentId_();
      var file = folder.createFile(Utilities.newBlob(decoded, f.mimeType, attachmentId + '__' + cleanName));
      try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (shareErr) { }

      var uploadedAt = new Date().toISOString();
      var item = {
        attachmentId: attachmentId,
        taskId: taskId,
        driveFileId: file.getId(),
        driveUrl: file.getUrl(),
        url: file.getUrl(),
        fileName: cleanName,
        name: cleanName,
        mimeType: f.mimeType || '',
        sizeBytes: sizeBytes,
        checksum: checksum,
        status: ATT_STATUS.ACTIVE,
        uploadedByEmail: actor.email,
        uploadedByName: actor.name,
        uploadedAt: uploadedAt,
        sourceFlow: sourceFlow,
        sourceRef: sourceRef,
        isDeleted: false,
        deletedAt: '',
        deletedBy: '',
        isImage: String(f.mimeType || '').indexOf('image/') === 0
      };

      rowsToAppend.push([
        item.attachmentId,
        item.taskId,
        item.driveFileId,
        item.driveUrl,
        item.fileName,
        item.mimeType,
        item.sizeBytes,
        item.checksum,
        item.status,
        item.uploadedByEmail,
        item.uploadedByName,
        item.uploadedAt,
        item.sourceFlow,
        item.sourceRef,
        false,
        '',
        ''
      ]);
      uploaded.push(item);
      if (checksum) checksumMap[checksum] = item;

      emitEvent_({
        type: EVENT.ATTACHMENT_ADDED,
        taskId: taskId,
        projectId: tr.data[COL.PROJECT_ID] || '',
        actorEmail: actor.email,
        actorTeam: actor.team || '',
        notes: cleanName,
        meta: {
          attachmentId: attachmentId,
          driveFileId: item.driveFileId,
          sourceFlow: sourceFlow,
          mimeType: item.mimeType,
          sizeBytes: item.sizeBytes
        }
      });
    } catch (err) {
      failed.push({ fileName: cleanName, reason: err.message || 'Upload failed.' });
    }
  });

  if (rowsToAppend.length) {
    withScriptLock_(10000, function () {
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rowsToAppend.length, ATT_COL.TOTAL_COLS).setValues(rowsToAppend);

      var taskSheet = getSheet(SHEETS.TASKS);
      var row = tr.data;
      if (!row[COL.DRIVE_URL]) row[COL.DRIVE_URL] = uploaded[0].driveUrl;
      row[COL.LAST_ACTION] = new Date().toISOString();
      taskSheet.getRange(tr.rowIndex, 1, 1, COL.TOTAL_COLS).setValues([row]);
    });
    invalidateTaskIndex_();
    if (typeof invalidateDashCache_ === 'function') invalidateDashCache_();
  }

  failed.forEach(function (f) {
    emitEvent_({
      type: EVENT.ATTACHMENT_UPLOAD_FAILED,
      taskId: taskId,
      projectId: tr.data[COL.PROJECT_ID] || '',
      actorEmail: actor.email,
      actorTeam: actor.team || '',
      notes: f.fileName,
      meta: { sourceFlow: sourceFlow, reason: f.reason }
    });
  });

  return {
    uploaded: uploaded,
    skippedDuplicates: skippedDuplicates,
    failed: failed
  };
}

/**
 * Upload one or more files for a task.
 * @param {string}   taskId     — task the files belong to
 * @param {Array}    filesData  — [{name, mimeType, base64}]
 * @returns {Array}  [{name, url, id}]
 */
function uploadTaskFiles(taskId, filesData, sourceFlow, sourceRef) {
  if (!taskId || !filesData || !filesData.length) return [];
  var actor = requireRole_(['Owner', 'Manager', 'Member']);
  var result = saveTaskAttachments_(taskId, filesData, sourceFlow || 'comment', sourceRef || '', actor);
  if (result.error) return [];
  return result.uploaded || [];
}

/**
 * Finalize attachments for a newly created task.
 * Stores files in Drive, exposes the first file in task details,
 * and logs all uploaded files into the task comment thread.
 */
function attachFilesToTask(taskId, filesData, sourceFlow, sourceRef) {
  try {
    if (!taskId || !filesData || !filesData.length) return ok_([]);
    var actor = requireRole_(['Owner', 'Manager', 'Member']);
    var result = saveTaskAttachments_(taskId, filesData, sourceFlow || 'task', sourceRef || '', actor);
    if (result.error === 'NOT_FOUND') return err_('NOT_FOUND');
    if (result.error === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    if ((result.uploaded || []).length) addComment(taskId, '', result.uploaded);
    return ok_(result);
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('attachFilesToTask: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

function getTaskAttachments(taskId) {
  try {
    var actor = requireRole_(['Owner', 'Manager', 'Member']);
    if (!taskId) return err_('INVALID_INPUT');
    var tr = findTaskRow_(taskId);
    if (!tr) return err_('NOT_FOUND');
    if (!canActorAccessTaskRow_(actor, tr.data)) return err_('UNAUTHORIZED');
    return ok_(listTaskAttachments_(taskId, true));
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('getTaskAttachments: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}


// ══════════════════════════════════════════════════════════
// COMMENTS / ACTIVITY THREAD
// Thread is per-task; supports text + Drive attachment URLs.
// ══════════════════════════════════════════════════════════

function getCommentSheet_() {
  var sh = getSheet(SHEETS.COMMENTS);
  if (!sh) {
    var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!id) throw new Error('SPREADSHEET_ID not set in Script Properties');
    var ss = SpreadsheetApp.openById(id);
    sh = ss.insertSheet('Comments');
    sh.appendRow(['CommentID', 'TaskID', 'AuthorEmail', 'AuthorName', 'Text', 'Timestamp', 'AttachmentURLs']);
    sh.getRange(1, 1, 1, 7).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/**
 * Add a comment to a task thread.
 * @param {string} taskId
 * @param {string} text
 * @param {Array}  attachmentUrls — optional Drive URLs from uploadTaskFiles
 */
function addComment(taskId, text, attachmentUrls) {
  try {
    var actor = requireRole_(['Owner', 'Manager', 'Member']);
    var cleanTxt = sanitize_(text || '');
    var hasText = !!cleanTxt.trim();
    var hasAttachments = !!(attachmentUrls && attachmentUrls.length);
    if (!taskId || (!hasText && !hasAttachments)) return err_('INVALID_INPUT');

    if (actor.role !== 'Owner') {
      var tr = findTaskRow_(taskId);
      if (!tr) return err_('NOT_FOUND');
      if (!canActorAccessTaskRow_(actor, tr.data)) return err_('UNAUTHORIZED');
    }

    var commentId = 'CMT-' + Date.now();
    var ts = new Date();
    var urlsJson = JSON.stringify(attachmentUrls || []);

    getCommentSheet_().appendRow([
      commentId, taskId, actor.email, actor.name, cleanTxt.trim(), ts, urlsJson
    ]);

    // Send email notification
    try {
      if (typeof sendCommentNotificationEmail === 'function') {
        var trToNotify = (typeof tr !== 'undefined' && tr) ? tr : findTaskRow_(taskId);
        if (trToNotify) {
          var taskName = trToNotify.data[COL.TASK_NAME] || taskId;
          var ownerEmail = String(trToNotify.data[COL.OWNER_EMAIL] || '').toLowerCase();
          var creatorEmail = String(trToNotify.data[COL.CREATOR_EMAIL] || '').toLowerCase();
          var actorE = String(actor.email || '').toLowerCase();

          var recipients = [];
          // Standard Participants
          if (ownerEmail && ownerEmail !== actorE) recipients.push(ownerEmail);
          if (creatorEmail && creatorEmail !== actorE && creatorEmail !== ownerEmail) recipients.push(creatorEmail);

          // Parse @mentions
          // Supports @email@domain.com or @FirstName (if unique in memMap)
          var memMap = getMemberMap_();
          var mentionRegex = /@([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}|[a-zA-Z0-9._-]+)/g;
          var match;
          while ((match = mentionRegex.exec(cleanTxt)) !== null) {
            var mention = match[1].toLowerCase();
            var targetEmail = '';

            if (mention.indexOf('@') > 0) {
              targetEmail = mention;
            } else {
              // Try to find by name (case-insensitive check against all members)
              for (var em in memMap) {
                var m = memMap[em];
                if (m && (m.name || '').toLowerCase().indexOf(mention) !== -1) {
                  targetEmail = em.toLowerCase();
                  break;
                }
              }
            }

            if (targetEmail && targetEmail !== actorE && recipients.indexOf(targetEmail) === -1) {
              recipients.push(targetEmail);
            }
          }

          // Also notify if only an attachment was added with no text (optional but good)
          // The user specifically asked for text comments, but let's be robust.

          for (var i = 0; i < recipients.length; i++) {
            var mem = memMap[recipients[i]];
            var recipientObj = mem ? { email: recipients[i], name: mem.name } : { email: recipients[i], name: 'User' };
            sendCommentNotificationEmail(recipientObj, taskName, taskId, actor.name, cleanTxt.trim() || '(Attachment added)');
          }
        }
      }
    } catch (emailErr) {
      console.warn('Silent addComment email fail: ' + emailErr.message);
    }

    return ok_({ commentId: commentId, timestamp: ts.toISOString() });

  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('addComment: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

function getDeptManagerForTeam_(teamName) {
  var team = String(teamName || '').trim();
  var teams = getAllTeams_() || [];
  var memberMap = getMemberMap_();
  var activeMembers = Object.values(memberMap).filter(function (m) { return m && m.active; });

  var teamCfg = null;
  for (var i = 0; i < teams.length; i++) {
    if (String(teams[i].name || '') === team) {
      teamCfg = teams[i];
      break;
    }
  }

  var explicitEmail = teamCfg && teamCfg.managerEmail ? String(teamCfg.managerEmail).toLowerCase().trim() : '';
  if (explicitEmail && memberMap[explicitEmail] && memberMap[explicitEmail].active) {
    return memberMap[explicitEmail];
  }

  for (var j = 0; j < activeMembers.length; j++) {
    if (activeMembers[j].team === team && activeMembers[j].role === 'Manager') return activeMembers[j];
  }

  for (var k = 0; k < activeMembers.length; k++) {
    if (activeMembers[k].role === 'Owner') return activeMembers[k];
  }

  return null;
}

function sendTaskEscalationEmail_(recipient, payload) {
  try {
    if (!recipient || !recipient.email) return;
    var subject = '[TaskFlow] Escalation - ' + (payload.taskName || payload.taskId || 'Task');
    var lines = [
      'Task ID: ' + (payload.taskId || ''),
      'Task: ' + (payload.taskName || ''),
      'Escalated by: ' + (payload.actorName || ''),
      'Current team: ' + (payload.teamName || ''),
      'Status: ' + (payload.status || ''),
      'Reason: ' + (payload.note || 'No note provided')
    ];
    GmailApp.sendEmail(recipient.email, subject, lines.join('\n'));
  } catch (e) {
    console.warn('sendTaskEscalationEmail_ failed: ' + e.message);
  }
}

function escalateTaskToDeptManager(taskId, note) {
  try {
    var actor = requireRole_(['Owner', 'Manager', 'Member']);
    if (!taskId) return err_('INVALID_INPUT');
    var tr = findTaskRow_(taskId);
    if (!tr) return err_('NOT_FOUND');
    if (!canActorAccessTaskRow_(actor, tr.data)) return err_('UNAUTHORIZED');

    var status = String(tr.data[COL.STATUS] || '');
    if (status === STATUS.ARCHIVED) return err_('INVALID_INPUT');

    var taskName = tr.data[COL.TASK_NAME] || taskId;
    var taskTeam = tr.data[COL.CURRENT_TEAM] || tr.data[COL.HOME_TEAM] || actor.team || '';
    var manager = getDeptManagerForTeam_(taskTeam);
    if (!manager) return err_('NOT_FOUND');

    var cleanNote = sanitize_(note || '').trim();
    var commentText = '[Dept Manager Escalation] ' + actor.name + ' escalated this task to ' + manager.name
      + (cleanNote ? '. Context: ' + cleanNote : '.');
    addComment(taskId, commentText, []);
    sendTaskEscalationEmail_(manager, {
      taskId: taskId,
      taskName: taskName,
      actorName: actor.name,
      teamName: taskTeam,
      status: status,
      note: cleanNote
    });

    try {
      emitSystemMessage_('team', taskTeam,
        'Department escalation: ' + taskId + ' needs manager attention.',
        {
          mode: 'dept_manager',
          taskId: taskId,
          taskName: taskName,
          status: status,
          teamName: taskTeam,
          targetName: manager.name,
          targetEmail: manager.email,
          escalatedBy: actor.name,
          note: cleanNote
        });
    } catch (e) { }

    return ok_({
      message: manager.email.toLowerCase() === String(actor.email || '').toLowerCase()
        ? 'You are already the department manager for this task.'
        : ('Escalated to ' + manager.name + '.'),
      managerName: manager.name,
      managerEmail: manager.email,
      teamName: taskTeam
    });
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('escalateTaskToDeptManager: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

function escalateTaskToHodDesk(taskId, note, targetType, targetId) {
  try {
    var actor = requireRole_(['Owner', 'Manager']);
    if (!taskId) return err_('INVALID_INPUT');
    var tr = findTaskRow_(taskId);
    if (!tr) return err_('NOT_FOUND');
    if (!canActorAccessTaskRow_(actor, tr.data)) return err_('UNAUTHORIZED');

    var status = String(tr.data[COL.STATUS] || '');
    if (status === STATUS.ARCHIVED) return err_('INVALID_INPUT');

    var taskName = tr.data[COL.TASK_NAME] || taskId;
    var taskTeam = tr.data[COL.CURRENT_TEAM] || tr.data[COL.HOME_TEAM] || actor.team || '';
    var deadline = tr.data[COL.DEADLINE] ? new Date(tr.data[COL.DEADLINE]).toISOString() : '';
    var cleanNote = sanitize_(note || '').trim();
    var ctxType = targetType === 'space' ? 'space' : 'managers';
    var ctxId = ctxType === 'space' ? String(targetId || '').trim() : 'managers';
    var targetLabel = 'HOD Desk';
    var targetSpace = null;

    if (ctxType === 'space') {
      targetSpace = getChatSpaceById_(ctxId);
      if (!targetSpace) return err_('NOT_FOUND');
      if (!canAccessChatSpace_(actor, targetSpace)) return err_('UNAUTHORIZED');
      targetLabel = targetSpace.name || 'Private HOD Space';
    } else {
      ctxId = 'managers';
    }

    addComment(taskId,
      '[' + targetLabel + ' Escalation] ' + actor.name + ' raised this task for cross-department review.'
      + (cleanNote ? ' Context: ' + cleanNote : ''),
      []);

    emitSystemMessage_(ctxType, ctxId,
      'HOD escalation: ' + taskId + ' - ' + taskName,
      {
        mode: ctxType === 'space' ? 'hod_space' : 'hod_desk',
        taskId: taskId,
        taskName: taskName,
        status: status,
        teamName: taskTeam,
        targetType: ctxType,
        targetId: ctxId,
        targetLabel: targetLabel,
        spaceId: targetSpace ? targetSpace.spaceId : '',
        spaceName: targetSpace ? targetSpace.name : '',
        ownerName: tr.data[COL.OWNER_NAME] || '',
        ownerEmail: tr.data[COL.OWNER_EMAIL] || '',
        createdBy: tr.data[COL.CREATED_BY] || '',
        deadline: deadline,
        escalatedBy: actor.name,
        escalatedByEmail: actor.email,
        note: cleanNote
      });

    return ok_({ message: 'Escalated to ' + targetLabel + '.' });
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('escalateTaskToHodDesk: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

/**
 * Retrieve all comments for a task, chronological.
 * @param {string} taskId
 */
function getComments(taskId) {
  try {
    var actor = requireRole_(['Owner', 'Manager', 'Member']);
    if (!taskId) return err_('INVALID_INPUT');

    var tr = findTaskRow_(taskId);
    if (!tr) return err_('NOT_FOUND');

    // W6 SPRINT 2: Tighten visibility to current owner + creator + Owners/Managers.
    // Previously used canActorAccessTaskRow_ which also allowed past owners and
    // teammates via team-scope — too broad per business decision.
    if (actor.role !== 'Owner' && actor.role !== 'Manager') {
      var ownerEmail = String(tr.data[COL.OWNER_EMAIL] || '').toLowerCase();
      var creatorEmail = String(tr.data[COL.CREATOR_EMAIL] || '').toLowerCase();
      var actEmail = String(actor.email || '').toLowerCase();
      if (actEmail !== ownerEmail && actEmail !== creatorEmail) return err_('UNAUTHORIZED');
    }

    var sheet = getCommentSheet_();
    var data = sheet.getDataRange().getValues();
    var out = [];

    for (var i = 1; i < data.length; i++) {
      if (data[i][1] !== taskId) continue;
      var urls = [];
      try { urls = JSON.parse(data[i][6] || '[]'); } catch (e) { }
      out.push({
        commentId: data[i][0],
        taskId: data[i][1],
        authorEmail: data[i][2],
        authorName: data[i][3],
        text: data[i][4],
        timestamp: data[i][5] ? new Date(data[i][5]).toISOString() : '',
        attachments: urls
      });
    }

    out.sort(function (a, b) { return new Date(a.timestamp) - new Date(b.timestamp); });
    return ok_(out);

  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('getComments: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}


// ══════════════════════════════════════════════════════════
// UTILITY HELPERS
// ══════════════════════════════════════════════════════════

/**
 * getMemberPhone_ — thin wrapper used by idle alert email templates.
 * Kept here so calendarnotifications.gs doesn't need importing.
 */
function getMemberPhone_(email) {
  var m = getMemberByEmail_(email);
  return m ? String(m.phone || '') : '';
}

// getDateDaysAgo_ lives in AnalyticsEngine_v6.gs — callable from all files in same GAS project
