/**
 * Wonder+ Portal master API stabilizer.
 *
 * Load this file last in Apps Script. It keeps the spreadsheet as the source of
 * truth for schedules, assignments, and important notices, while forwarding
 * existing auth/event-guide endpoints to the already installed functions.
 */

var WONDER_PORTAL_SCHEDULE_SS_ID = '1SRb3_nwgPWEj2Kb44SDD38SI0HmkllE4G_P7DPYKGKk';
var WONDER_PORTAL_BOARD_FALLBACK_SHEET_NAMES = ['重要事項共有専用', '重要事項', 'ImportantNotices', 'BoardPosts'];

function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = String(params.action || 'schedule');
  var callback = params.callback;
  var result;

  try {
    if (action === 'schedule') {
      result = buildPortalSchedulePayloadMasterStable_();
    } else if (action === 'boardPosts') {
      result = { ok: true, posts: getPortalBoardPostsMasterStable_(), updatedAt: new Date().toISOString(), source: 'spreadsheet-master' };
    } else if (action === 'createBoardPost') {
      result = createPortalBoardPostMasterStable_(params);
    } else if (action === 'eventGuides' && typeof buildPortalEventGuidesPayloadFinal_ === 'function') {
      result = buildPortalEventGuidesPayloadFinal_();
    } else if (action === 'minutes' && typeof getPortalMinutesPayload_ === 'function') {
      result = getPortalMinutesPayload_();
    } else if (action === 'login' && typeof loginPortalUser_ === 'function') {
      result = loginPortalUser_(params);
    } else if (action === 'register' && typeof registerPortalUser_ === 'function') {
      result = registerPortalUser_(params);
    } else if (action === 'resetPassword' && typeof resetPortalPasswordRequest_ === 'function') {
      result = resetPortalPasswordRequest_(params);
    } else if (action === 'toggleAssignment' && typeof togglePortalAssignment_ === 'function') {
      result = togglePortalAssignment_(params);
    } else if (action === 'admin' && typeof handlePortalAdminAction_ === 'function') {
      result = handlePortalAdminAction_(params);
    } else {
      result = { ok: false, error: 'unsupported action: ' + action };
    }
  } catch (error) {
    result = { ok: false, error: String(error && error.message ? error.message : error), action: action };
  }

  return portalJsonOutputMasterStable_(result, callback);
}

function doPost(e) {
  return doGet(e);
}

function portalJsonOutputMasterStable_(payload, callback) {
  var json = JSON.stringify(payload || {});
  if (callback) {
    return ContentService.createTextOutput(String(callback) + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function buildPortalSchedulePayloadMasterStable_() {
  var ss = openPortalScheduleSpreadsheetMasterStable_();
  var sheet = getPortalFormattedScheduleSheetMasterStable_(ss);
  if (!sheet) throw new Error('formatted schedule sheet not found');

  var schedules = createPortalEmptyScheduleMapMasterStable_();
  var values = sheet.getDataRange().getValues();
  var header = values[0] || [];
  var indexes = getPortalScheduleHeaderIndexesMasterStable_(header);
  var seen = {};
  var counts = {};

  for (var i = 1; i < values.length; i += 1) {
    var row = values[i];
    var item = makePortalScheduleItemFromRowMasterStable_(row, indexes, i + 1);
    if (!item) continue;

    var key = portalScheduleKeyMasterStable_(item.month, item.category);
    if (!schedules[key]) continue;
    if (seen[item.id]) continue;
    seen[item.id] = true;
    schedules[key].items.push(item);
    counts[key] = (counts[key] || 0) + 1;
  }

  Object.keys(schedules).forEach(function(key) {
    schedules[key].items.sort(function(a, b) {
      return a.day - b.day || String(a.timeRange || a.meta).localeCompare(String(b.timeRange || b.meta), 'ja') || String(a.title).localeCompare(String(b.title), 'ja');
    });
  });

  var assignmentResult = getPortalAssignmentMapMasterStable_();

  return {
    ok: true,
    updatedAt: new Date().toISOString(),
    source: 'spreadsheet-master-stable',
    masterSource: 'spreadsheet-master-stable',
    scheduleCounts: counts,
    assignments: assignmentResult.map,
    assignmentSource: assignmentResult.source,
    schedules: schedules
  };
}

function openPortalScheduleSpreadsheetMasterStable_() {
  var props = PropertiesService.getScriptProperties();
  var ids = [
    props.getProperty('SCHEDULE_SPREADSHEET_ID'),
    props.getProperty('WONDER_PORTAL_SCHEDULE_SPREADSHEET_ID'),
    WONDER_PORTAL_SCHEDULE_SS_ID
  ];
  for (var i = 0; i < ids.length; i += 1) {
    if (!ids[i]) continue;
    try {
      var ss = SpreadsheetApp.openById(ids[i]);
      if (getPortalFormattedScheduleSheetMasterStable_(ss)) return ss;
    } catch (error) {
      // Continue to the next candidate.
    }
  }
  throw new Error('schedule spreadsheet not found');
}

function getPortalFormattedScheduleSheetMasterStable_(ss) {
  if (!ss) return null;
  return ss.getSheetByName('整形済み') || ss.getSheetByName('整形') || ss.getSheetByName('Formatted');
}

function createPortalEmptyScheduleMapMasterStable_() {
  var map = {};
  var monthNames = {
    1: 'january', 2: 'february', 3: 'march', 4: 'april', 5: 'may', 6: 'june',
    7: 'july', 8: 'august', 9: 'september', 10: 'october', 11: 'november', 12: 'december'
  };
  var types = [
    ['wonder', 'Wonder+'],
    ['regular', 'レギュラー'],
    ['miraiba', 'ミライバ'],
    ['meeting', 'ミーティング']
  ];
  Object.keys(monthNames).forEach(function(monthText) {
    var month = Number(monthText);
    types.forEach(function(type) {
      var key = monthNames[month] + type[0].charAt(0).toUpperCase() + type[0].slice(1);
      map[key] = { month: month, category: type[0], label: month + '月 ' + type[1], items: [] };
    });
  });
  return map;
}

function getPortalScheduleHeaderIndexesMasterStable_(header) {
  var idx = {};
  var names = header.map(function(value) { return String(value || '').trim(); });
  function find(patterns, fallback) {
    for (var i = 0; i < names.length; i += 1) {
      for (var j = 0; j < patterns.length; j += 1) {
        if (patterns[j].test(names[i])) return i;
      }
    }
    return fallback;
  }
  idx.date = find([/開催日|日付|date/i], 3);
  idx.category = find([/種別|区分|category|type/i], 2);
  idx.place = find([/会場|場所|開催地|place|venue/i], 4);
  idx.eventName = find([/イベント|名称|タイトル|event|name/i], 5);
  idx.time = find([/時間|時刻|time/i], 6);
  idx.note = find([/備考|メモ|note/i], 8);
  idx.raw = find([/原文|raw/i], 9);
  return idx;
}

function makePortalScheduleItemFromRowMasterStable_(row, indexes, rowNumber) {
  var date = portalDateValueMasterStable_(row[indexes.date]);
  if (!date) return null;

  var allText = row.map(function(value) { return String(value || ''); }).join(' ');
  var category = portalScheduleCategoryMasterStable_(row, indexes);
  var place = portalCleanTextMasterStable_(row[indexes.place]);
  var eventName = portalCleanTextMasterStable_(row[indexes.eventName]);
  var raw = portalCleanTextMasterStable_(row[indexes.raw] || allText);
  var timeRange = portalNormalizeTimeRangeMasterStable_(row[indexes.time], raw, category);
  var note = portalCleanTextMasterStable_(row[indexes.note]);
  var title = portalScheduleTitleMasterStable_(category, place, eventName, raw);
  var metaParts = [timeRange];
  if (note && !/^\(?\d+\s*名\)?$/.test(note)) metaParts.push(note.replace(/\(?\d+\s*名\)?/g, '').trim());
  var meta = metaParts.filter(Boolean).join(' / ').replace(/\s*\/\s*$/g, '').trim();
  var id = [
    date.getFullYear(),
    date.getMonth() + 1,
    date.getDate(),
    category,
    portalIdTextMasterStable_(title),
    portalIdTextMasterStable_(timeRange),
    rowNumber
  ].join('|');

  return {
    id: id,
    year: date.getFullYear(),
    month: date.getMonth() + 1,
    category: category,
    day: date.getDate(),
    title: title,
    meta: meta,
    timeRange: timeRange,
    sourceRow: rowNumber
  };
}

function portalDateValueMasterStable_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  var text = String(value || '').trim();
  if (!text) return null;
  var date = new Date(text);
  if (!isNaN(date.getTime())) return date;
  var match = text.match(/(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
  if (match) return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  return null;
}

function portalScheduleCategoryMasterStable_(row, indexes) {
  var categoryText = String(row[indexes.category] || '').toLowerCase();
  var text = row.map(function(value) { return String(value || ''); }).join(' ');
  if (/meeting|ミーティング|mtg/i.test(categoryText + ' ' + text)) return 'meeting';
  if (/ミライバ/i.test(categoryText + ' ' + text)) return 'miraiba';
  if (/朝活|レギュラー|異業種交流会/.test(categoryText + ' ' + text) && !/wonder|w\+|wonder\+|\+story|\+leaders|\+ladies|\+lady|\+gravity|\+beauty|\+finance|\+entertainment|\+cxo|\+alliance|\+night|\+100/i.test(text)) return 'regular';
  return 'wonder';
}

function portalScheduleTitleMasterStable_(category, place, eventName, raw) {
  if (category === 'meeting') return 'ミーティング';
  if (category === 'miraiba') return 'ミライバ';
  if (category === 'regular') {
    return /朝活/.test(raw) ? '朝活異業種交流会' : 'レギュラー異業種交流会';
  }
  var cleanPlace = place.replace(/銀座$/g, '銀座').trim();
  var cleanName = eventName
    .replace(/^W\+/i, 'Wonder+')
    .replace(/^Wonder\+/i, 'Wonder+')
    .replace(/^\+/g, 'Wonder+')
    .replace(/\s+/g, '');
  if (!cleanName || /^イベント$/.test(cleanName)) cleanName = 'Wonder+';
  if (!/Wonder\+/i.test(cleanName)) cleanName = 'Wonder+' + cleanName.replace(/^\+/g, '');
  return [cleanPlace, cleanName].filter(Boolean).join(' ').trim();
}

function portalNormalizeTimeRangeMasterStable_(value, raw, category) {
  var text = String(value || '').trim();
  var rawText = String(raw || '');
  var range = '';
  var rangeMatch = (text + ' ' + rawText).match(/(\d{1,2}:\d{2})\s*[-ーｰ－―–—~〜～]+\s*(\d{1,2}:\d{2})/);
  if (rangeMatch) range = rangeMatch[1] + '-' + rangeMatch[2];
  if (!range) {
    var startMatch = (text + ' ' + rawText).match(/(\d{1,2}:\d{2})/);
    if (startMatch) {
      var start = startMatch[1];
      var minutes = category === 'meeting' ? 90 : category === 'regular' ? 60 : 90;
      range = start + '-' + portalAddMinutesMasterStable_(start, minutes);
    }
  }
  return range.replace(/\s+/g, '').replace(/--+/g, '-');
}

function portalAddMinutesMasterStable_(time, minutes) {
  var match = String(time || '').match(/^(\d{1,2}):(\d{2})$/);
  if (!match) return '';
  var date = new Date(2000, 0, 1, Number(match[1]), Number(match[2]) + minutes, 0);
  return ('0' + date.getHours()).slice(-2) + ':' + ('0' + date.getMinutes()).slice(-2);
}

function portalCleanTextMasterStable_(value) {
  return String(value || '')
    .replace(/\u3000/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/\(?\d+\s*名\)?/g, '')
    .trim();
}

function portalIdTextMasterStable_(value) {
  return String(value || '').toLowerCase().replace(/\s+/g, '').replace(/[＋]/g, '+').trim();
}

function portalScheduleKeyMasterStable_(month, category) {
  var monthNames = {
    1: 'january', 2: 'february', 3: 'march', 4: 'april', 5: 'may', 6: 'june',
    7: 'july', 8: 'august', 9: 'september', 10: 'october', 11: 'november', 12: 'december'
  };
  var suffix = { wonder: 'Wonder', regular: 'Regular', miraiba: 'Miraiba', meeting: 'Meeting' }[category || 'wonder'];
  return monthNames[Number(month)] && suffix ? monthNames[Number(month)] + suffix : '';
}

function getPortalAssignmentMapMasterStable_() {
  if (typeof getPortalAssignmentMap_ === 'function') {
    try {
      var existing = getPortalAssignmentMap_();
      if (existing && Object.keys(existing).length) return { map: existing, source: 'existing-getPortalAssignmentMap' };
    } catch (error) {
      // Fall through to generic sheet parsing.
    }
  }

  var ss = openPortalScheduleSpreadsheetMasterStable_();
  var result = {};
  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    if (!/(担当|参加|assignment|assign)/i.test(name)) return;
    var values = sheet.getDataRange().getValues();
    if (values.length < 2) return;
    var headers = values[0].map(function(value) { return String(value || '').trim(); });
    var eventIdx = findHeaderIndexMasterStable_(headers, [/event.?id|schedule.?id|予定ID|イベントID|id/i], 0);
    var nameIdx = findHeaderIndexMasterStable_(headers, [/氏名|名前|name|担当者|user/i], 1);
    var checkedIdx = findHeaderIndexMasterStable_(headers, [/checked|参加|status|状態/i], -1);
    for (var i = 1; i < values.length; i += 1) {
      var row = values[i];
      var eventId = String(row[eventIdx] || '').trim();
      var person = String(row[nameIdx] || '').trim();
      if (!eventId || !person) continue;
      if (checkedIdx >= 0 && /^(false|0|no|off|削除|不参加|cancel)/i.test(String(row[checkedIdx] || '').trim())) continue;
      if (!result[eventId]) result[eventId] = [];
      if (result[eventId].indexOf(person) === -1) result[eventId].push(person);
    }
  });
  return { map: result, source: 'generic-assignment-sheets' };
}

function findHeaderIndexMasterStable_(headers, patterns, fallback) {
  for (var i = 0; i < headers.length; i += 1) {
    for (var j = 0; j < patterns.length; j += 1) {
      if (patterns[j].test(headers[i])) return i;
    }
  }
  return fallback;
}

function getPortalBoardPostsMasterStable_() {
  var sheets = findPortalBoardSheetsMasterStable_();
  var posts = [];
  sheets.forEach(function(sheet) {
    var values = sheet.getDataRange().getValues();
    if (values.length < 2) return;
    var headers = values[0].map(function(value) { return String(value || '').trim(); });
    var messageIdx = findHeaderIndexMasterStable_(headers, [/内容|本文|message|body|重要/i], 1);
    var authorIdx = findHeaderIndexMasterStable_(headers, [/投稿者|氏名|名前|author|name/i], 0);
    var createdIdx = findHeaderIndexMasterStable_(headers, [/日時|作成|created|timestamp|time/i], 2);
    for (var i = 1; i < values.length; i += 1) {
      var row = values[i];
      var message = String(row[messageIdx] || '').trim();
      if (!message) continue;
      var created = row[createdIdx] instanceof Date ? row[createdIdx].toISOString() : String(row[createdIdx] || '').trim();
      posts.push({
        id: sheet.getSheetId() + ':' + (i + 1),
        author: String(row[authorIdx] || 'ポータル利用者').trim(),
        message: message,
        createdAt: created || new Date().toISOString()
      });
    }
  });
  posts.sort(function(a, b) { return String(b.createdAt).localeCompare(String(a.createdAt)); });
  return posts.slice(0, 50);
}

function findPortalBoardSheetsMasterStable_() {
  var candidates = [];
  var props = PropertiesService.getScriptProperties();
  var ids = [
    props.getProperty('PORTAL_BOARD_SPREADSHEET_ID'),
    props.getProperty('WONDER_PORTAL_BOARD_SPREADSHEET_ID'),
    props.getProperty('SCHEDULE_SPREADSHEET_ID'),
    props.getProperty('WONDER_PORTAL_SCHEDULE_SPREADSHEET_ID'),
    WONDER_PORTAL_SCHEDULE_SS_ID
  ];
  ids.forEach(function(id) {
    if (!id) return;
    try {
      var ss = SpreadsheetApp.openById(id);
      ss.getSheets().forEach(function(sheet) {
        var name = sheet.getName();
        if (WONDER_PORTAL_BOARD_FALLBACK_SHEET_NAMES.indexOf(name) >= 0 || /重要|notice|board/i.test(name)) {
          candidates.push(sheet);
        }
      });
    } catch (error) {
      // Ignore inaccessible candidate.
    }
  });
  return candidates;
}

function createPortalBoardPostMasterStable_(params) {
  var sheets = findPortalBoardSheetsMasterStable_();
  var sheet = sheets[0];
  if (!sheet) {
    var ss = openPortalScheduleSpreadsheetMasterStable_();
    sheet = ss.insertSheet('重要事項共有専用');
    sheet.appendRow(['投稿者', '内容', '日時']);
  }
  var author = String(params.author || params.name || 'ポータル利用者').trim();
  var message = String(params.message || params.body || params.text || '').trim();
  if (!message) return { ok: false, error: 'message is required' };
  sheet.appendRow([author, message, new Date()]);

  try {
    if (typeof sendImportantNoticeToLineWorks_ === 'function') {
      sendImportantNoticeToLineWorks_(message, author);
    }
  } catch (notifyError) {
    return { ok: true, notifyError: String(notifyError && notifyError.message ? notifyError.message : notifyError), posts: getPortalBoardPostsMasterStable_() };
  }
  return { ok: true, posts: getPortalBoardPostsMasterStable_() };
}
