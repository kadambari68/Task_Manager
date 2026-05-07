// chatengine_sprint8.js - TaskFlow v6
// ============================================================

function ensureChatSheet_() {
  var sh = getSheet(SHEETS.CHAT);
  if (sh) return sh;

  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  var ss = SpreadsheetApp.openById(id);
  sh = ss.insertSheet(SHEETS.CHAT);
  sh.appendRow(['MessageID', 'ContextType', 'ContextID', 'Text', 'AuthorEmail', 'AuthorName', 'Timestamp', 'IsSystem', 'MetaJson']);
  sh.getRange(1, 1, 1, 9).setFontWeight('bold');
  sh.setFrozenRows(1);
  _sheetCache[SHEETS.CHAT] = sh;
  return sh;
}

function ensureChatSpacesSheet_() {
  var sh = getSheet(SHEETS.CHAT_SPACES);
  if (sh) return sh;

  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  var ss = SpreadsheetApp.openById(id);
  sh = ss.insertSheet(SHEETS.CHAT_SPACES);
  sh.appendRow(['SpaceID', 'SpaceName', 'ParticipantEmailsJson', 'CreatedByEmail', 'CreatedByName', 'CreatedAt', 'IsActive']);
  sh.getRange(1, 1, 1, 7).setFontWeight('bold');
  sh.setFrozenRows(1);
  _sheetCache[SHEETS.CHAT_SPACES] = sh;
  return sh;
}

function setupChatSheet() {
  requireRole_(['Owner']);
  ensureChatSheet_();
  ensureChatSpacesSheet_();
  return ok_({ message: 'Chat sheets are ready.' });
}

function parseJsonArray_(raw) {
  try {
    var arr = raw ? JSON.parse(raw) : [];
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    return [];
  }
}

function uniqueEmails_(emails) {
  var seen = {};
  var out = [];
  (emails || []).forEach(function (email) {
    var em = String(email || '').toLowerCase().trim();
    if (!em || seen[em]) return;
    seen[em] = true;
    out.push(em);
  });
  return out;
}

function parseChatSpaceRow_(row, memberMap) {
  memberMap = memberMap || getMemberMap_();
  var participantEmails = uniqueEmails_(parseJsonArray_(row[2]));
  var participantNames = participantEmails.map(function (email) {
    var member = memberMap[email];
    return member ? (member.name || email) : email;
  });

  return {
    spaceId: row[0] || '',
    name: row[1] || 'Private Space',
    participantEmails: participantEmails,
    participantNames: participantNames,
    createdByEmail: String(row[3] || '').toLowerCase(),
    createdByName: row[4] || '',
    createdAt: row[5] || '',
    isActive: row[6] !== false && row[6] !== 'FALSE'
  };
}

function getAllChatSpaces_() {
  var sheet = ensureChatSpacesSheet_();
  var data = sheet.getDataRange().getValues();
  var memberMap = getMemberMap_();
  var spaces = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var space = parseChatSpaceRow_(row, memberMap);
    if (!space.isActive) continue;
    spaces.push(space);
  }
  return spaces;
}

function getChatSpaceById_(spaceId) {
  if (!spaceId) return null;
  var spaces = getAllChatSpaces_();
  for (var i = 0; i < spaces.length; i++) {
    if (spaces[i].spaceId === spaceId) return spaces[i];
  }
  return null;
}

function canAccessChatSpace_(user, space) {
  if (!user || !space) return false;
  if (user.role !== 'Owner' && user.role !== 'Manager') return false;
  return space.participantEmails.indexOf(String(user.email || '').toLowerCase()) !== -1;
}

function canAccessChatContext_(user, ctxType, ctxId) {
  if (!user) return false;

  if (ctxType === 'managers') {
    return user.role === 'Owner' || user.role === 'Manager';
  }

  if (ctxType === 'team') {
    if (user.role === 'Owner') return true;
    return String(user.team || '') === String(ctxId || '');
  }

  if (ctxType === 'space') {
    return canAccessChatSpace_(user, getChatSpaceById_(ctxId));
  }

  return false;
}

function getAccessibleChatSpaces() {
  try {
    var user = requireRole_(['Owner', 'Manager']);
    var spaces = getAllChatSpaces_().filter(function (space) {
      return canAccessChatSpace_(user, space);
    }).sort(function (a, b) {
      return String(a.name || '').localeCompare(String(b.name || ''));
    });

    return ok_(spaces.map(function (space) {
      return {
        spaceId: space.spaceId,
        name: space.name,
        participantEmails: space.participantEmails,
        participantNames: space.participantNames,
        createdByEmail: space.createdByEmail,
        createdByName: space.createdByName,
        createdAt: space.createdAt
      };
    }));
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('getAccessibleChatSpaces: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

function createManagerChatSpace(name, participantEmails) {
  try {
    var actor = requireRole_(['Owner', 'Manager']);
    var cleanName = sanitize_(name || '').trim();
    if (!cleanName) return err_('INVALID_INPUT');

    var normalized = uniqueEmails_(participantEmails || []);
    normalized.push(String(actor.email || '').toLowerCase());
    normalized = uniqueEmails_(normalized);

    if (normalized.length < 2) {
      return { success: false, message: 'Select at least one manager to invite.' };
    }

    var memberMap = getMemberMap_();
    var chosen = [];
    for (var i = 0; i < normalized.length; i++) {
      var member = memberMap[normalized[i]];
      if (!member || !member.active) return err_('INVALID_INPUT');
      if (member.role !== 'Owner' && member.role !== 'Manager') return err_('INVALID_INPUT');
      chosen.push(member);
    }

    var sheet = ensureChatSpacesSheet_();
    var now = new Date().toISOString();
    var spaceId = generateLogId('SPACE');
    sheet.appendRow([
      spaceId,
      cleanName,
      JSON.stringify(chosen.map(function (m) { return m.email; })),
      actor.email,
      actor.name,
      now,
      true
    ]);

    return ok_({
      spaceId: spaceId,
      name: cleanName,
      participantEmails: chosen.map(function (m) { return m.email; }),
      participantNames: chosen.map(function (m) { return m.name; }),
      createdByEmail: actor.email,
      createdByName: actor.name,
      createdAt: now
    });
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('createManagerChatSpace: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

function findChatMessageRow_(messageId) {
  if (!messageId) return null;
  var sheet = ensureChatSheet_();
  var hit = sheet.getRange('A:A').createTextFinder(String(messageId)).matchEntireCell(true).findNext();
  if (!hit) return null;
  return {
    sheet: sheet,
    rowIndex: hit.getRow(),
    data: sheet.getRange(hit.getRow(), 1, 1, 9).getValues()[0]
  };
}

function canDeleteChatMessage_(user, msgRow) {
  if (!user || !msgRow) return false;
  var isSystem = msgRow[7] === true || msgRow[7] === 'TRUE';
  if (isSystem) return false;

  var authorEmail = String(msgRow[4] || '').toLowerCase();
  var actorEmail = String(user.email || '').toLowerCase();
  if (authorEmail === actorEmail) return true;

  if (user.role === 'Owner' && canAccessChatContext_(user, msgRow[1], msgRow[2])) return true;

  if (msgRow[1] === 'space') {
    var space = getChatSpaceById_(msgRow[2]);
    if (space && space.createdByEmail === actorEmail) return true;
  }

  return false;
}

function postMessage(ctxType, ctxId, text) {
  try {
    var user = getCurrentUser();
    if (!user) return err_('UNAUTHORIZED');
    if (!canAccessChatContext_(user, ctxType, ctxId)) return err_('UNAUTHORIZED');

    var cleanText = sanitize_(text || '').trim();
    if (!cleanText) return err_('INVALID_INPUT');

    var sheet = ensureChatSheet_();
    var msgId = generateLogId('MSG');
    var tsStr = new Date().toISOString();

    sheet.appendRow([
      msgId,
      ctxType,
      ctxId,
      cleanText,
      user.email,
      user.name,
      tsStr,
      false,
      ''
    ]);

    return ok_({ messageId: msgId, timestamp: tsStr });
  } catch (e) {
    console.error('postMessage: ' + e.message);
    return err_(e.message === 'UNAUTHORIZED' ? 'UNAUTHORIZED' : 'SYSTEM_ERROR');
  }
}

function deleteChatMessage(messageId) {
  try {
    var user = getCurrentUser();
    if (!user) return err_('UNAUTHORIZED');

    var hit = findChatMessageRow_(messageId);
    if (!hit) return err_('NOT_FOUND');
    if (!canDeleteChatMessage_(user, hit.data)) return err_('UNAUTHORIZED');

    var meta = {};
    try { meta = hit.data[8] ? JSON.parse(hit.data[8]) : {}; } catch (e) { meta = {}; }
    var now = new Date().toISOString();
    meta.deleted = true;
    meta.deletedAt = now;
    meta.deletedBy = user.name;
    meta.deletedByEmail = user.email;

    hit.data[3] = '';
    hit.data[6] = now;
    hit.data[8] = JSON.stringify(meta);
    hit.sheet.getRange(hit.rowIndex, 1, 1, 9).setValues([hit.data]);

    return ok_({ message: 'Message deleted.' });
  } catch (e) {
    console.error('deleteChatMessage: ' + e.message);
    return err_(e.message === 'UNAUTHORIZED' ? 'UNAUTHORIZED' : 'SYSTEM_ERROR');
  }
}

function getMessages(ctxType, ctxId, afterTs) {
  try {
    var user = getCurrentUser();
    if (!user) return err_('UNAUTHORIZED');
    if (!canAccessChatContext_(user, ctxType, ctxId)) return err_('UNAUTHORIZED');

    var sheet = ensureChatSheet_();
    var data = sheet.getDataRange().getValues();
    var msgs = [];
    var after = afterTs ? new Date(afterTs) : null;

    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (r[1] !== ctxType || r[2] !== ctxId) continue;

      var ts = new Date(r[6]);
      if (after && ts <= after) continue;

      var meta = {};
      try { meta = r[8] ? JSON.parse(r[8]) : {}; } catch (e) { meta = {}; }

      msgs.push({
        messageId: r[0],
        type: r[1],
        contextId: r[2],
        text: meta.deleted ? '' : (r[3] || ''),
        authorEmail: r[4],
        authorName: r[5],
        timestamp: r[6],
        isSystem: r[7] === true || r[7] === 'TRUE',
        isOwn: String(r[4] || '').toLowerCase() === String(user.email || '').toLowerCase(),
        deleted: meta.deleted === true,
        meta: meta || {}
      });
    }

    if (!afterTs && msgs.length > 200) msgs = msgs.slice(-200);
    return ok_(msgs);
  } catch (e) {
    console.error('getMessages: ' + e.message);
    return err_(e.message === 'UNAUTHORIZED' ? 'UNAUTHORIZED' : 'SYSTEM_ERROR');
  }
}

function emitSystemMessage_(ctxType, ctxId, text, meta) {
  try {
    var sheet = ensureChatSheet_();
    var cleanText = sanitize_(text || '');
    cleanText = cleanText.replace(/^.*?(TASK-\S+)\s+assigned to\s+(.+)\s+by\s+(.+)$/, '$1 created for $2 by $3');
    sheet.appendRow([
      generateLogId('SYSMSG'),
      ctxType,
      ctxId,
      cleanText,
      'system@taskflow',
      'System',
      new Date().toISOString(),
      true,
      JSON.stringify(meta || {})
    ]);
  } catch (e) {
    console.warn('emitSystemMessage_ failed: ' + e.message);
  }
}
