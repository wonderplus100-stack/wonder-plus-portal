/**
 * Wonder+ Portal event guide realtime refresh add-on.
 *
 * Paste this file at the bottom of the existing Apps Script project, save it,
 * run installEventGuideRealtimeTrigger() once, then deploy a new web app
 * version. This version avoids mojibake-prone source text by using Unicode
 * escapes for Japanese labels and by matching form headers with several
 * fallbacks.
 */

var PORTAL_EVENT_GUIDE_REALTIME_SHEET_NAME = '\u30a4\u30d9\u30f3\u30c8\u6848\u5185';
var PORTAL_EVENT_GUIDE_PRIMARY_SPREADSHEET_ID = '1rXae1o13ucNAI6VfPhXKIFi3LTq3uV0g4wKm_zI4zKM';
var PORTAL_EVENT_GUIDE_REALTIME_HEADERS = [
  'CreatedAt',
  'Title',
  'Month',
  'Day',
  'EventDate',
  'StartTime',
  'EndTime',
  'Body',
  'MaterialUrl',
  'AttachmentUrls',
  'Category',
  'UpdatedAt',
  'Source'
];

var PORTAL_EVENT_GUIDE_TITLE_KEYS = [
  '\u30a4\u30d9\u30f3\u30c8\u540d',
  '\u30a4\u30d9\u30f3\u30c8\u30bf\u30a4\u30c8\u30eb',
  '\u6848\u5185\u30bf\u30a4\u30c8\u30eb',
  '\u6848\u5185\u540d',
  '\u30bf\u30a4\u30c8\u30eb',
  '\u30d6\u30e9\u30f3\u30c9',
  'event name',
  'eventname',
  'title',
  'name',
  'brand'
];

var PORTAL_EVENT_GUIDE_BODY_KEYS = [
  '\u6848\u5185\u6587',
  '\u6848\u5185\u672c\u6587',
  '\u672c\u6587',
  '\u6587\u7ae0',
  '\u6295\u7a3f\u6587',
  '\u30e1\u30c3\u30bb\u30fc\u30b8',
  '\u5185\u5bb9',
  'body',
  'text',
  'content',
  'message',
  'description'
];

var PORTAL_EVENT_GUIDE_MONTH_KEYS = [
  '\u6708',
  '\u5bfe\u8c61\u6708',
  '\u958b\u50ac\u6708',
  '\u5e74\u6708',
  '\u958b\u50ac\u5e74\u6708',
  'month',
  'eventmonth',
  'targetmonth'
];

var PORTAL_EVENT_GUIDE_DAY_KEYS = [
  '\u65e5',
  '\u958b\u50ac\u65e5',
  '\u5bfe\u8c61\u65e5',
  'day',
  'eventday',
  'targetday'
];

var PORTAL_EVENT_GUIDE_DATE_KEYS = [
  '\u958b\u50ac\u65e5',
  '\u958b\u50ac\u65e5\u6642',
  '\u65e5\u6642',
  '\u65e5\u4ed8',
  '\u30a4\u30d9\u30f3\u30c8\u65e5',
  '\u30a4\u30d9\u30f3\u30c8\u65e5\u6642',
  '\u5bfe\u8c61\u65e5',
  'event date',
  'eventdate',
  'date',
  'datetime',
  'eventdatetime'
];

var PORTAL_EVENT_GUIDE_START_TIME_KEYS = [
  '\u958b\u59cb\u6642\u9593',
  '\u958b\u59cb\u6642\u523b',
  '\u30b9\u30bf\u30fc\u30c8',
  '\u958b\u59cb',
  'starttime',
  'start time',
  'start'
];

var PORTAL_EVENT_GUIDE_END_TIME_KEYS = [
  '\u7d42\u4e86\u6642\u9593',
  '\u7d42\u4e86\u6642\u523b',
  '\u30af\u30ed\u30fc\u30ba',
  '\u7d42\u4e86',
  'endtime',
  'end time',
  'close',
  'end'
];

var PORTAL_EVENT_GUIDE_MATERIAL_KEYS = [
  '\u7d20\u6750URL',
  '\u7d20\u6750\u30ea\u30f3\u30af',
  '\u8cc7\u6599URL',
  'drive url',
  'drive',
  'materialurl',
  'material url',
  'url'
];

var PORTAL_EVENT_GUIDE_ATTACHMENT_KEYS = [
  '\u6dfb\u4ed8\u30d5\u30a1\u30a4\u30eb',
  '\u6dfb\u4ed8\u30d5\u30a1\u30a4\u30ebURL',
  '\u753b\u50cf',
  '\u753b\u50cfURL',
  '\u52d5\u753b',
  '\u7d20\u6750\u30d5\u30a1\u30a4\u30eb',
  'attachment',
  'attachments',
  'attachmenturls',
  'file',
  'files',
  'image',
  'images'
];

var PORTAL_EVENT_GUIDE_CATEGORY_KEYS = [
  '\u30ab\u30c6\u30b4\u30ea',
  '\u7a2e\u5225',
  '\u533a\u5206',
  '\u30d6\u30e9\u30f3\u30c9',
  'category',
  'type',
  'brand'
];

function installEventGuideRealtimeTrigger() {
  var formUrl = getPortalEventGuideFormUrl_();
  if (!formUrl) {
    throw new Error('EVENT_GUIDE_FORM_URL or CONFIG.eventGuideFormUrl is not set.');
  }
  deletePortalRealtimeTriggers_(['onEventGuideFormSubmit']);
  ScriptApp.newTrigger('onEventGuideFormSubmit')
    .forForm(openPortalFormByAnyUrl_(formUrl))
    .onFormSubmit()
    .create();
  Logger.log('Event guide realtime trigger installed: ' + formUrl);
}

function onEventGuideFormSubmit(e) {
  var answers = getPortalNamedFormAnswers_(e);
  var ss = getPortalEventGuideSpreadsheet_();
  var sheet = ensurePortalEventGuideRealtimeSheet_(ss);
  var row = normalizePortalEventGuideRecord_(answers, 'form-submit');
  sheet.appendRow([
    row.createdAt,
    row.title,
    row.month,
    row.day,
    row.eventDate,
    row.startTime,
    row.endTime,
    row.body,
    row.materialUrl,
    row.attachmentUrls,
    row.category,
    row.updatedAt,
    row.source
  ]);
  SpreadsheetApp.flush();
  var refresh = refreshPortalEventGuideCache_();
  Logger.log(JSON.stringify({ ok: true, eventGuideTitle: row.title, month: row.month, refresh: refresh }));
  return { ok: true, guide: row, refresh: refresh };
}

function buildPortalEventGuidesSafe_() {
  var guideLibrary = buildPortalEventGuideLibrary_();
  var scheduledGuides = buildPortalEventGuidesFromSchedule_(guideLibrary);
  if (scheduledGuides.length) {
    return scheduledGuides.sort(function(a, b) {
      return (a.month - b.month) || (a.day - b.day) || String(a.title).localeCompare(String(b.title), 'ja') || String(a.timeRange).localeCompare(String(b.timeRange), 'ja');
    });
  }
  return dedupePortalEventGuides_(guideLibrary.rawGuides)
    .sort(function(a, b) { return getPortalGuideUpdatedTime_(b) - getPortalGuideUpdatedTime_(a); });
}

function buildPortalEventGuideLibrary_() {
  var ss = getPortalEventGuideSpreadsheet_();
  var rawGuides = [];
  getPortalEventGuideCandidateSheets_(ss).forEach(function(sheet) {
    var values = sheet.getDataRange().getValues();
    if (values.length < 2) return;
    var headers = values[0].map(function(value) { return String(value || '').trim(); });
    for (var r = 1; r < values.length; r += 1) {
      var raw = {};
      headers.forEach(function(header, c) {
        if (header) raw[header] = values[r][c];
      });
      var guide = normalizePortalEventGuideRecord_(raw, sheet.getName(), values[r], headers);
      if (isUsablePortalEventGuide_(guide)) rawGuides.push(guide);
    }
  });
  var guides = dedupePortalEventGuides_(rawGuides);
  var byExact = {};
  var byVenueGeneric = {};
  var byTypeOnly = {};
  var byGeneric = [];
  guides.forEach(function(guide) {
    var identity = getPortalGuideIdentity_(guide);
    guide.guideVenueKey = identity.venueKey;
    guide.guideTypeKey = identity.typeKey;
    guide.guideIdentityKey = identity.key;
    if (identity.venueKey && identity.typeKey !== 'generic') setLatestPortalGuide_(byExact, identity.venueKey + '|' + identity.typeKey, guide);
    if (identity.venueKey && identity.typeKey === 'generic') setLatestPortalGuide_(byVenueGeneric, identity.venueKey, guide);
    if (!identity.venueKey && identity.typeKey !== 'generic') setLatestPortalGuide_(byTypeOnly, identity.typeKey, guide);
    if (!identity.venueKey && identity.typeKey === 'generic') byGeneric.push(guide);
  });
  byGeneric.sort(function(a, b) { return getPortalGuideUpdatedTime_(b) - getPortalGuideUpdatedTime_(a); });
  return {
    rawGuides: rawGuides,
    guides: guides,
    byExact: byExact,
    byVenueGeneric: byVenueGeneric,
    byTypeOnly: byTypeOnly,
    byGeneric: byGeneric
  };
}

function setLatestPortalGuide_(map, key, guide) {
  if (!key) return;
  if (!map[key] || getPortalGuideUpdatedTime_(guide) >= getPortalGuideUpdatedTime_(map[key])) {
    map[key] = guide;
  }
}

function buildPortalEventGuidesFromSchedule_(library) {
  var sheet = getPortalFormattedScheduleSheetForGuides_();
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  var guides = [];
  var seen = {};
  for (var i = 1; i < values.length; i += 1) {
    var row = values[i];
    var date = row[3] instanceof Date ? row[3] : new Date(row[3]);
    if (isNaN(date.getTime())) continue;
    var category = typeof scheduleTypeKey_ === 'function' ? scheduleTypeKey_(row) : 'wonder';
    if (category !== 'wonder') continue;

    var place = String(row[4] || '').trim();
    var eventName = String(row[5] || '').trim();
    var raw = String(row[9] || '').trim();
    var scheduleTime = typeof portalFormatScheduleTime_ === 'function'
      ? portalFormatScheduleTime_(row[6])
      : String(row[6] || '').trim();
    var displayTime = typeof portalDisplayTimeRangeFinal_ === 'function'
      ? portalDisplayTimeRangeFinal_(scheduleTime, raw)
      : scheduleTime;
    var timeParts = splitPortalTimeRange_(displayTime);
    var title = typeof portalScheduleTitle_ === 'function'
      ? portalScheduleTitle_(place, eventName)
      : [place, eventName].filter(Boolean).join(' ');
    var contextText = [place, eventName, title, raw].join(' ');
    var identity = getPortalGuideIdentity_({ title: title, body: contextText, category: 'wonder' });
    var matchedGuide = findPortalGuideForSchedule_(library, identity);
    var fallbackBody = '\u3053\u306e\u30a4\u30d9\u30f3\u30c8\u306b\u7d10\u3065\u304f\u6848\u5185\u6587\u7ae0\u306f\u307e\u3060\u8aad\u307f\u8fbc\u307e\u308c\u3066\u3044\u307e\u305b\u3093\u3002';
    var body = matchedGuide ? String(matchedGuide.body || matchedGuide.text || '').trim() : fallbackBody;
    body = syncPortalGuideBodyWithSchedule_(body, title, date.getMonth() + 1, date.getDate(), displayTime);
    var attachmentUrls = matchedGuide ? (matchedGuide.attachmentUrls || '') : '';
    var materialUrl = matchedGuide ? (matchedGuide.materialUrl || '') : '';
    var key = [date.getMonth() + 1, date.getDate(), title, displayTime].join('|');
    if (seen[key]) continue;
    seen[key] = true;
    guides.push({
      title: title,
      eventName: title,
      name: title,
      month: date.getMonth() + 1,
      day: date.getDate(),
      eventMonth: date.getMonth() + 1,
      eventDay: date.getDate(),
      eventDate: (date.getMonth() + 1) + '\u6708' + date.getDate() + '\u65e5',
      startTime: timeParts.start,
      endTime: timeParts.end,
      timeRange: displayTime,
      startEndTime: displayTime,
      eventTime: displayTime,
      dateTime: buildPortalEventGuideDateTime_(date.getMonth() + 1, date.getDate(), displayTime, ''),
      eventDateTime: buildPortalEventGuideDateTime_(date.getMonth() + 1, date.getDate(), displayTime, ''),
      body: body,
      text: body,
      materialUrl: materialUrl,
      attachmentUrls: attachmentUrls,
      attachments: matchedGuide ? (matchedGuide.attachments || makePortalGuideAttachmentList_(attachmentUrls || materialUrl)) : [],
      files: matchedGuide ? (matchedGuide.files || matchedGuide.attachments || makePortalGuideAttachmentList_(attachmentUrls || materialUrl)) : [],
      category: 'wonder',
      createdAt: matchedGuide ? matchedGuide.createdAt : '',
      updatedAt: matchedGuide ? matchedGuide.updatedAt : '',
      source: matchedGuide ? matchedGuide.source : 'schedule-without-guide',
      guideMatched: !!matchedGuide,
      guideIdentityKey: identity.key
    });
  }
  return guides;
}

function getPortalFormattedScheduleSheetForGuides_() {
  try {
    var ss = typeof getPortalScheduleSpreadsheetFull_ === 'function'
      ? getPortalScheduleSpreadsheetFull_()
      : SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SCHEDULE_SPREADSHEET_ID'));
    return ss.getSheetByName('\u6574\u5f62\u6e08\u307f') || ss.getSheetByName('\u6574\u5f62');
  } catch (error) {
    Logger.log('event guide schedule sheet skipped: ' + error.message);
    return null;
  }
}

function findPortalGuideForSchedule_(library, identity) {
  if (!library || !identity) return null;
  var exactKey = identity.venueKey + '|' + identity.typeKey;
  if (identity.venueKey && identity.typeKey !== 'generic' && library.byExact[exactKey]) return library.byExact[exactKey];
  if (identity.venueKey && library.byVenueGeneric[identity.venueKey]) return library.byVenueGeneric[identity.venueKey];
  if (identity.typeKey !== 'generic' && library.byTypeOnly[identity.typeKey]) return library.byTypeOnly[identity.typeKey];
  return null;
}

function getPortalGuideIdentity_(guide) {
  var text = [
    guide && (guide.title || guide.eventName || guide.name || ''),
    guide && (guide.body || guide.text || ''),
    guide && (guide.materialUrl || ''),
    guide && (guide.attachmentUrls || '')
  ].join(' ');
  var category = normalizePortalEventGuideCategory_(guide && guide.category || text || 'wonder');
  var venueKey = normalizePortalVenueKey_(text);
  var typeKey = normalizePortalGuideTypeKey_(text);
  return {
    category: category,
    venueKey: venueKey,
    typeKey: typeKey,
    key: [category, venueKey || 'anywhere', typeKey || 'generic'].join('|')
  };
}

function normalizePortalGuideTypeKey_(value) {
  var key = normalizePortalTextKey_(value);
  if (/realestate|\u4e0d\u52d5\u7523/.test(key)) return 'realestate';
  if (/alliance/.test(key)) return 'alliance';
  if (/entertainment/.test(key)) return 'entertainment';
  if (/leaders|leader/.test(key)) return 'leaders';
  if (/ladies|lady/.test(key)) return 'ladies';
  if (/story/.test(key)) return 'story';
  if (/gravity/.test(key)) return 'gravity';
  if (/beauty/.test(key)) return 'beauty';
  if (/finance/.test(key)) return 'finance';
  if (/cxo/.test(key)) return 'cxo';
  if (/night/.test(key)) return 'night';
  if (/(?:^|[^0-9])100(?:[^0-9]|$)|100\u4eba/.test(key)) return '100';
  return 'generic';
}

function splitPortalTimeRange_(value) {
  var text = String(value || '').replace(/[\u301c\uff5e\uff0d\u2014\u2015]/g, '-');
  var match = text.match(/(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})/);
  if (match) return { start: match[1], end: match[2] };
  var single = text.match(/(\d{1,2}:\d{2})/);
  return { start: single ? single[1] : '', end: '' };
}

function syncPortalGuideBodyWithSchedule_(body, title, month, day, timeRange) {
  var text = stripPortalEventGuideMaterialFooter_(body || '');
  var dateText = month && day ? month + '\u6708' + day + '\u65e5 ' + timeRange : timeRange;
  if (dateText) {
    if (/\u3010\u958b\u50ac\u65e5\u6642\u3011/.test(text)) {
      text = text.replace(/\u3010\u958b\u50ac\u65e5\u6642\u3011\s*[^\n]*/g, '\u3010\u958b\u50ac\u65e5\u6642\u3011 ' + dateText);
    }
    var parts = splitPortalTimeRange_(timeRange);
    if (parts.start && parts.end) {
      var open = subtractPortalMinutes_(parts.start, 15);
      var line = open + ' \u30aa\u30fc\u30d7\u30f3 | ' + parts.start + ' \u30b9\u30bf\u30fc\u30c8 | ' + parts.end + ' \u30af\u30ed\u30fc\u30ba';
      text = text.replace(/\d{1,2}:\d{2}\s*\u30aa\u30fc\u30d7\u30f3\s*[|\uff5c]\s*\d{1,2}:\d{2}\s*\u30b9\u30bf\u30fc\u30c8\s*[|\uff5c]\s*\d{1,2}:\d{2}\s*\u30af\u30ed\u30fc\u30ba/g, line);
    }
  }
  return text;
}

function subtractPortalMinutes_(time, minutes) {
  var match = String(time || '').match(/^(\d{1,2}):(\d{2})$/);
  if (!match) return time;
  var date = new Date(2000, 0, 1, Number(match[1]), Number(match[2]));
  date.setMinutes(date.getMinutes() - Number(minutes || 0));
  return Utilities.formatDate(date, 'Asia/Tokyo', 'H:mm');
}

function buildPortalEventGuidesRealtimeFallback_() {
  return buildPortalEventGuidesSafe_();
}

function buildPortalEventGuidesPayload_() {
  var guides = buildPortalEventGuidesSafe_();
  return {
    ok: true,
    updatedAt: new Date().toISOString(),
    source: 'schedule-joined-event-guides',
    sourceSpreadsheetId: PORTAL_EVENT_GUIDE_PRIMARY_SPREADSHEET_ID,
    count: Array.isArray(guides) ? guides.length : 0,
    eventGuides: guides
  };
}

function getPortalEventGuidesForApi_() {
  return buildPortalEventGuidesPayload_();
}

function getPortalEventGuideDiagnostics_() {
  var ss = getPortalEventGuideSpreadsheet_();
  var sheets = getPortalEventGuideCandidateSheets_(ss).map(function(sheet) {
    var headers = [];
    if (sheet.getLastRow() >= 1 && sheet.getLastColumn() >= 1) {
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
        .map(function(value) { return String(value || '').trim(); });
    }
    var samples = [];
    if (sheet.getLastRow() >= 2 && headers.length) {
      var rowCount = Math.min(3, sheet.getLastRow() - 1);
      var values = sheet.getRange(2, 1, rowCount, sheet.getLastColumn()).getValues();
      values.forEach(function(row) {
        var raw = {};
        headers.forEach(function(header, index) {
          if (header) raw[header] = row[index];
        });
        var guide = normalizePortalEventGuideRecord_(raw, sheet.getName(), row, headers);
        samples.push({
          title: guide.title,
          month: guide.month,
          day: guide.day,
          startTime: guide.startTime,
          endTime: guide.endTime,
          updatedAt: String(guide.updatedAt || ''),
          source: guide.source
        });
      });
    }
    return {
      name: sheet.getName(),
      rows: sheet.getLastRow(),
      columns: sheet.getLastColumn(),
      priority: getPortalEventGuideSheetPriority_(sheet),
      headers: headers,
      samples: samples
    };
  });
  return {
    ok: true,
    sourceSpreadsheetId: PORTAL_EVENT_GUIDE_PRIMARY_SPREADSHEET_ID,
    updatedAt: new Date().toISOString(),
    sheets: sheets
  };
}

function refreshPortalEventGuideCache_() {
  var props = PropertiesService.getScriptProperties();
  var now = new Date();
  var count = 0;
  var scheduleCount = 0;
  try {
    var guides = buildPortalEventGuidesSafe_();
    count = Array.isArray(guides) ? guides.length : 0;
  } catch (error) {
    Logger.log('Event guide refresh skipped: ' + error.message);
  }
  clearPortalScheduleCaches_();
  try {
    if (typeof buildPortalSchedulePayload_ === 'function') {
      scheduleCount = countPortalScheduleItems_(buildPortalSchedulePayload_());
    }
  } catch (error) {
    Logger.log('Schedule payload warmup skipped: ' + error.message);
  }
  props.setProperties({
    PORTAL_EVENT_GUIDES_LAST_REFRESH_AT: now.toISOString(),
    PORTAL_EVENT_GUIDES_LAST_COUNT: String(count),
    PORTAL_SCHEDULE_LAST_REFRESH_AT: now.toISOString(),
    PORTAL_SCHEDULE_LAST_COUNT: String(scheduleCount)
  }, true);
  return {
    ok: true,
    eventGuideCount: count,
    scheduleCount: scheduleCount,
    refreshedAt: Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
  };
}

function forceRefreshPortalDataAfterFormSubmit() {
  return refreshPortalEventGuideCache_();
}

function getPortalEventGuideFormUrl_() {
  var props = PropertiesService.getScriptProperties();
  var configured = String(props.getProperty('EVENT_GUIDE_FORM_URL') || '').trim();
  if (configured) return configured;
  try {
    return String(CONFIG && CONFIG.eventGuideFormUrl || '').trim();
  } catch (error) {
    return '';
  }
}

function getPortalEventGuideSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var candidates = [
    PORTAL_EVENT_GUIDE_PRIMARY_SPREADSHEET_ID,
    props.getProperty('EVENT_GUIDE_SPREADSHEET_ID'),
    props.getProperty('EVENT_INFO_SPREADSHEET_ID'),
    props.getProperty('SCHEDULE_SPREADSHEET_ID'),
    props.getProperty('SPREADSHEET_ID'),
    props.getProperty('PORTAL_SPREADSHEET_ID')
  ].filter(Boolean);
  for (var i = 0; i < candidates.length; i += 1) {
    try {
      return SpreadsheetApp.openById(String(candidates[i]).trim());
    } catch (error) {}
  }
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  throw new Error('Event guide spreadsheet was not found. Set EVENT_INFO_SPREADSHEET_ID.');
}

function ensurePortalEventGuideRealtimeSheet_(ss) {
  var sheet = ss.getSheetByName(PORTAL_EVENT_GUIDE_REALTIME_SHEET_NAME) || ss.insertSheet(PORTAL_EVENT_GUIDE_REALTIME_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, PORTAL_EVENT_GUIDE_REALTIME_HEADERS.length).setValues([PORTAL_EVENT_GUIDE_REALTIME_HEADERS]);
  }
  return sheet;
}

function getPortalEventGuideCandidateSheets_(ss) {
  var sheets = ss.getSheets();
  var candidates = [];
  sheets.forEach(function(sheet) {
    var nameKey = normalizePortalTextKey_(sheet.getName());
    if (nameKey.indexOf(normalizePortalTextKey_(PORTAL_EVENT_GUIDE_REALTIME_SHEET_NAME)) !== -1 ||
        /event|guide|form/.test(nameKey)) {
      candidates.push(sheet);
      return;
    }
    if (sheet.getLastRow() < 2 || sheet.getLastColumn() < 2) return;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var score = headers.reduce(function(total, header) {
      var key = normalizePortalTextKey_(header);
      return total + (matchesAnyPortalGuideKey_(key) ? 1 : 0);
    }, 0);
    if (score >= 2) candidates.push(sheet);
  });
  var selected = candidates.length ? candidates : sheets;
  return selected.sort(function(a, b) {
    return getPortalEventGuideSheetPriority_(b) - getPortalEventGuideSheetPriority_(a);
  });
}

function getPortalEventGuideSheetPriority_(sheet) {
  var priority = 0;
  var nameKey = normalizePortalTextKey_(sheet.getName());
  if (/form|response|\u30d5\u30a9\u30fc\u30e0|\u56de\u7b54/.test(nameKey)) priority += 30;
  if (nameKey.indexOf(normalizePortalTextKey_(PORTAL_EVENT_GUIDE_REALTIME_SHEET_NAME)) !== -1) priority += 10;
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return priority;
  try {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (pickPortalHeaderName_(headers, PORTAL_EVENT_GUIDE_DATE_KEYS)) priority += 30;
    if (pickPortalHeaderName_(headers, PORTAL_EVENT_GUIDE_START_TIME_KEYS)) priority += 12;
    if (pickPortalHeaderName_(headers, PORTAL_EVENT_GUIDE_END_TIME_KEYS)) priority += 12;
    if (pickPortalHeaderName_(headers, PORTAL_EVENT_GUIDE_TITLE_KEYS)) priority += 12;
    if (pickPortalHeaderName_(headers, PORTAL_EVENT_GUIDE_BODY_KEYS)) priority += 12;
  } catch (error) {}
  return priority;
}

function matchesAnyPortalGuideKey_(normalizedHeader) {
  var groups = [
    PORTAL_EVENT_GUIDE_TITLE_KEYS,
    PORTAL_EVENT_GUIDE_BODY_KEYS,
    PORTAL_EVENT_GUIDE_MONTH_KEYS,
    PORTAL_EVENT_GUIDE_DAY_KEYS,
    PORTAL_EVENT_GUIDE_DATE_KEYS,
    PORTAL_EVENT_GUIDE_START_TIME_KEYS,
    PORTAL_EVENT_GUIDE_END_TIME_KEYS,
    PORTAL_EVENT_GUIDE_MATERIAL_KEYS,
    PORTAL_EVENT_GUIDE_ATTACHMENT_KEYS,
    PORTAL_EVENT_GUIDE_CATEGORY_KEYS
  ];
  return groups.some(function(keys) {
    return keys.some(function(key) {
      var needle = normalizePortalTextKey_(key);
      return normalizedHeader === needle || isSafePortalHeaderPartialMatch_(normalizedHeader, needle);
    });
  });
}

function getPortalNamedFormAnswers_(e) {
  var answers = {};
  if (e && e.namedValues) {
    Object.keys(e.namedValues).forEach(function(key) {
      var value = e.namedValues[key];
      answers[key] = Array.isArray(value) ? value.join('\n') : String(value || '');
    });
  }
  if (e && e.response && typeof e.response.getItemResponses === 'function') {
    e.response.getItemResponses().forEach(function(itemResponse) {
      var title = itemResponse.getItem().getTitle();
      var response = itemResponse.getResponse();
      answers[title] = Array.isArray(response) ? response.join('\n') : String(response || '');
    });
  }
  return answers;
}

function normalizePortalEventGuideRecord_(record, source, rowValues, headers) {
  record = record || {};
  var body = stripPortalEventGuideMaterialFooter_(
    pickPortalAnswerStrict_(record, PORTAL_EVENT_GUIDE_BODY_KEYS) ||
    pickPortalAnswer_(record, PORTAL_EVENT_GUIDE_BODY_KEYS) ||
    inferPortalGuideBodyFromRow_(rowValues || [])
  );
  var title = pickPortalAnswerStrict_(record, PORTAL_EVENT_GUIDE_TITLE_KEYS) || pickPortalAnswer_(record, PORTAL_EVENT_GUIDE_TITLE_KEYS) || inferPortalEventGuideTitleFromBody_(body) || inferPortalGuideTitleFromRow_(rowValues || []);
  var materialUrl = pickPortalAnswerStrict_(record, PORTAL_EVENT_GUIDE_MATERIAL_KEYS) || pickPortalAnswer_(record, PORTAL_EVENT_GUIDE_MATERIAL_KEYS);
  var attachmentUrls = pickPortalAnswerStrict_(record, PORTAL_EVENT_GUIDE_ATTACHMENT_KEYS) || pickPortalAnswer_(record, PORTAL_EVENT_GUIDE_ATTACHMENT_KEYS) || inferPortalGuideUrlsFromRow_(rowValues || []);
  var eventDateRaw = pickPortalRawValueStrict_(record, PORTAL_EVENT_GUIDE_DATE_KEYS);
  var eventDate = formatPortalEventGuideDate_(eventDateRaw);
  var startTime = normalizePortalEventGuideTime_(pickPortalRawValueStrict_(record, PORTAL_EVENT_GUIDE_START_TIME_KEYS));
  var endTime = normalizePortalEventGuideTime_(pickPortalRawValueStrict_(record, PORTAL_EVENT_GUIDE_END_TIME_KEYS));
  var dateFromValue = parsePortalEventGuideMonthDayFromValue_(eventDateRaw);
  var explicitMonth = parsePortalEventGuideMonth_(pickPortalAnswerStrict_(record, PORTAL_EVENT_GUIDE_MONTH_KEYS));
  var explicitDay = parsePortalEventGuideDay_(pickPortalAnswerStrict_(record, PORTAL_EVENT_GUIDE_DAY_KEYS));
  var month = explicitMonth || dateFromValue.month || 0;
  var day = explicitDay || dateFromValue.day || 0;
  var category = normalizePortalEventGuideCategory_(pickPortalAnswerStrict_(record, PORTAL_EVENT_GUIDE_CATEGORY_KEYS) || pickPortalAnswer_(record, PORTAL_EVENT_GUIDE_CATEGORY_KEYS) || title || body || 'wonder');
  var updatedAt = pickPortalAnswerStrict_(record, ['UpdatedAt', 'updatedAt', '\u66f4\u65b0\u65e5\u6642', '\u30bf\u30a4\u30e0\u30b9\u30bf\u30f3\u30d7', 'Timestamp']) ||
    pickPortalAnswerStrict_(record, ['CreatedAt', 'createdAt', '\u767b\u9332\u65e5\u6642', '\u9001\u4fe1\u65e5\u6642']);
  var createdAt = pickPortalAnswerStrict_(record, ['CreatedAt', 'createdAt', '\u767b\u9332\u65e5\u6642', '\u9001\u4fe1\u65e5\u6642', 'Timestamp']) || updatedAt || new Date();
  var cleanTitle = normalizePortalEventGuideTitle_(title || inferPortalEventGuideTitle_(materialUrl || attachmentUrls) || 'Wonder+');
  var attachments = makePortalGuideAttachmentList_(attachmentUrls || materialUrl);
  var timeRange = startTime && endTime ? startTime + '-' + endTime : '';
  return {
    title: cleanTitle,
    eventName: cleanTitle,
    name: cleanTitle,
    month: month,
    day: day,
    eventDate: eventDate,
    eventMonth: month,
    eventDay: day,
    startTime: startTime,
    endTime: endTime,
    timeRange: timeRange,
    startEndTime: timeRange,
    eventTime: timeRange,
    dateTime: buildPortalEventGuideDateTime_(month, day, timeRange, eventDate),
    eventDateTime: buildPortalEventGuideDateTime_(month, day, timeRange, eventDate),
    body: body,
    text: body,
    materialUrl: materialUrl,
    attachmentUrls: attachmentUrls,
    attachments: attachments,
    files: attachments,
    category: category,
    createdAt: createdAt,
    updatedAt: updatedAt || createdAt,
    source: source || 'event-guide'
  };
}

function isUsablePortalEventGuide_(guide) {
  if (!guide || typeof guide !== 'object') return false;
  var title = String(guide.title || guide.eventName || guide.name || '').trim();
  var body = String(guide.body || guide.text || '').trim();
  var url = String(guide.materialUrl || guide.attachmentUrls || '').trim();
  if (!title && !body) return false;
  if (body.length < 8 && !url) return false;
  return true;
}

function dedupePortalEventGuides_(guides) {
  var map = {};
  guides.forEach(function(guide) {
    var identity = getPortalGuideIdentity_(guide);
    var key = identity.key;
    if (!key || key === 'wonder|anywhere|generic') key = 'fallback|' + normalizePortalTextKey_(guide.title || guide.eventName || guide.body || '');
    if (!key.replace(/[|]/g, '')) return;
    if (!map[key] || getPortalGuideUpdatedTime_(guide) >= getPortalGuideUpdatedTime_(map[key])) {
      map[key] = guide;
    }
  });
  return Object.keys(map).map(function(key) { return map[key]; });
}

function pickPortalAnswer_(object, names) {
  if (!object) return '';
  var keys = Object.keys(object);
  var normalizedKeys = keys.map(function(key) {
    return { key: key, normalized: normalizePortalTextKey_(key) };
  });
  for (var i = 0; i < names.length; i += 1) {
    var needle = normalizePortalTextKey_(names[i]);
    for (var j = 0; j < normalizedKeys.length; j += 1) {
      if (normalizedKeys[j].normalized === needle ||
          normalizedKeys[j].normalized.indexOf(needle) !== -1 ||
          needle.indexOf(normalizedKeys[j].normalized) !== -1) {
        return String(object[normalizedKeys[j].key] || '').trim();
      }
    }
  }
  return '';
}

function pickPortalAnswerStrict_(object, names) {
  var value = pickPortalRawValueStrict_(object, names);
  return value === null || typeof value === 'undefined' ? '' : String(value || '').trim();
}

function pickPortalRawValueStrict_(object, names) {
  if (!object) return '';
  var keys = Object.keys(object);
  var normalizedKeys = keys.map(function(key) {
    return { key: key, normalized: normalizePortalTextKey_(key) };
  });
  for (var phase = 0; phase < 2; phase += 1) {
    for (var i = 0; i < names.length; i += 1) {
      var needle = normalizePortalTextKey_(names[i]);
      if (!needle) continue;
      for (var j = 0; j < normalizedKeys.length; j += 1) {
        var candidate = normalizedKeys[j].normalized;
        if (phase === 0 && candidate === needle) {
          return object[normalizedKeys[j].key];
        }
        if (phase === 1 && isSafePortalHeaderPartialMatch_(candidate, needle)) {
          return object[normalizedKeys[j].key];
        }
      }
    }
  }
  return '';
}

function pickPortalHeaderName_(headers, names) {
  var record = {};
  (headers || []).forEach(function(header) {
    if (header) record[String(header)] = String(header);
  });
  return pickPortalAnswerStrict_(record, names);
}

function isSafePortalHeaderPartialMatch_(candidate, needle) {
  if (!candidate || !needle) return false;
  if (needle.length < 3 && candidate !== needle) return false;
  if (candidate.indexOf(needle) !== -1) return true;
  return needle.length >= 5 && needle.indexOf(candidate) !== -1;
}

function normalizePortalTextKey_(value) {
  var text = String(value || '').toLowerCase();
  try {
    text = text.normalize('NFKC');
  } catch (error) {}
  return text
    .replace(/[＋+]/g, '+')
    .replace(/\s+/g, '')
    .replace(/[【】\[\]（）()「」『』:：・,，.。_\-ー]/g, '');
}

function normalizePortalVenueKey_(value) {
  var key = normalizePortalTextKey_(value);
  var venues = [
    ['ginza', '\u9280\u5ea7'],
    ['shinjuku', '\u65b0\u5bbf'],
    ['nagoya', '\u540d\u53e4\u5c4b'],
    ['sendai', '\u4ed9\u53f0'],
    ['fukushima', '\u798f\u5cf6', 'fukushim'],
    ['koriyama', '\u90e1\u5c71'],
    ['kanazawa', '\u91d1\u6ca2'],
    ['sapporo', '\u672d\u5e4c'],
    ['fukuoka', '\u798f\u5ca1', 'hakata', '\u535a\u591a'],
    ['kitakyushu', '\u5317\u4e5d\u5dde', 'kitakyusyu'],
    ['kokura', '\u5c0f\u5009'],
    ['kumamoto', '\u718a\u672c'],
    ['niigata', '\u65b0\u6f5f'],
    ['yokohama', '\u6a2a\u6d5c'],
    ['funabashi', '\u8239\u6a4b'],
    ['osaka', '\u5927\u962a'],
    ['hiroshima', '\u5e83\u5cf6'],
    ['okayama', '\u5ca1\u5c71'],
    ['kobe', '\u795e\u6238'],
    ['kyoto', '\u4eac\u90fd'],
    ['takasaki', '\u9ad8\u5d0e'],
    ['maebashi', '\u524d\u6a4b'],
    ['utsunomiya', '\u5b87\u90fd\u5bae', 'utunomiya'],
    ['machida', '\u753a\u7530'],
    ['shizuoka', '\u9759\u5ca1'],
    ['hamamatsu', '\u6d5c\u677e'],
    ['omiya', '\u5927\u5bae'],
    ['chiba', '\u5343\u8449'],
    ['kofu', '\u7532\u5e9c']
  ];
  for (var i = 0; i < venues.length; i += 1) {
    for (var j = 0; j < venues[i].length; j += 1) {
      if (key.indexOf(normalizePortalTextKey_(venues[i][j])) !== -1) {
        return venues[i][0];
      }
    }
  }
  return '';
}

function parsePortalEventGuideMonth_(value) {
  var text = String(value || '');
  try {
    text = text.normalize('NFKC');
  } catch (error) {}
  var ym = text.match(/20\d{2}\s*\u5e74\s*(1[0-2]|[1-9])\s*\u6708/);
  if (ym) return Number(ym[1]);
  var month = text.match(/(1[0-2]|[1-9])\s*\u6708/);
  if (month) return Number(month[1]);
  var slash = text.match(/(?:^|[^\d])(1[0-2]|[1-9])\s*[\/.-]\s*\d{1,2}(?:[^\d]|$)/);
  if (slash) return Number(slash[1]);
  return 0;
}

function parsePortalEventGuideDay_(value) {
  var fromValue = parsePortalEventGuideMonthDayFromValue_(value);
  if (fromValue.day) return fromValue.day;
  var text = String(value || '');
  try {
    text = text.normalize('NFKC');
  } catch (error) {}
  var md = text.match(/(?:1[0-2]|[1-9])\s*\u6708\s*(3[01]|[12]\d|[1-9])\s*\u65e5/);
  if (md) return Number(md[1]);
  var slash = text.match(/(?:^|[^\d])(?:1[0-2]|[1-9])\s*[\/.-]\s*(3[01]|[12]\d|[1-9])(?:[^\d]|$)/);
  if (slash) return Number(slash[1]);
  var day = text.match(/(?:^|[^\d])(3[01]|[12]\d|[1-9])\s*\u65e5(?:[^\d]|$)/);
  return day ? Number(day[1]) : 0;
}

function parsePortalEventGuideMonthDay_(value) {
  var fromValue = parsePortalEventGuideMonthDayFromValue_(value);
  if (fromValue.month || fromValue.day) return fromValue;
  var text = String(value || '');
  try {
    text = text.normalize('NFKC');
  } catch (error) {}
  var md = text.match(/(?:20\d{2}\s*\u5e74\s*)?(1[0-2]|[1-9])\s*\u6708\s*(3[01]|[12]\d|[1-9])\s*\u65e5/);
  if (md) return { month: Number(md[1]), day: Number(md[2]) };
  var slash = text.match(/(?:^|[^\d])(1[0-2]|[1-9])\s*[\/.-]\s*(3[01]|[12]\d|[1-9])(?:[^\d]|$)/);
  if (slash) return { month: Number(slash[1]), day: Number(slash[2]) };
  return { month: parsePortalEventGuideMonth_(text), day: parsePortalEventGuideDay_(text) };
}

function normalizePortalEventGuideTime_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'HH:mm');
  }
  var text = String(value || '').trim();
  if (!text) return '';
  try {
    text = text.normalize('NFKC');
  } catch (error) {}
  var match = text.match(/([01]?\d|2[0-3])\s*[:\u6642]\s*([0-5]\d)?/);
  if (!match) return text;
  var hour = ('0' + Number(match[1])).slice(-2);
  var minute = match[2] ? ('0' + Number(match[2])).slice(-2) : '00';
  return hour + ':' + minute;
}

function parsePortalEventGuideMonthDayFromValue_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return {
      month: Number(Utilities.formatDate(value, 'Asia/Tokyo', 'M')),
      day: Number(Utilities.formatDate(value, 'Asia/Tokyo', 'd'))
    };
  }
  return { month: 0, day: 0 };
}

function formatPortalEventGuideDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'M\u6708d\u65e5');
  }
  return String(value || '').trim();
}

function buildPortalEventGuideDateTime_(month, day, timeRange, eventDate) {
  var dateText = '';
  if (month && day) dateText = month + '\u6708' + day + '\u65e5';
  else dateText = String(eventDate || '').trim();
  return [dateText, timeRange].filter(Boolean).join(' ');
}

function normalizePortalEventGuideCategory_(value) {
  var text = normalizePortalTextKey_(value);
  if (/regular|\u30ec\u30ae\u30e5\u30e9\u30fc|\u7570\u696d\u7a2e/.test(text)) return 'regular';
  if (/morning|\u671d\u6d3b/.test(text)) return 'morning';
  if (/meeting|\u30df\u30fc\u30c6\u30a3\u30f3\u30b0/.test(text)) return 'meeting';
  if (/miraiba|mirai|\u30df\u30e9\u30a4\u30d0/.test(text)) return 'miraiba';
  return 'wonder';
}

function stripPortalEventGuideMaterialFooter_(value) {
  return String(value || '')
    .replace(/\r\n/g, '\n')
    .replace(/\n{0,2}\u3010\u7d20\u6750\u3011[\s\S]*$/m, '')
    .replace(/\n{0,2}\[\u7d20\u6750\][\s\S]*$/mi, '')
    .replace(/\n{0,2}\u7d20\u6750\s*[:\uff1a]\s*(?:https?:\/\/\S+|Drive|Google Drive)[\s\S]*$/mi, '')
    .trim();
}

function normalizePortalEventGuideTitle_(value) {
  return String(value || '')
    .replace(/\s+/g, ' ')
    .replace(/[＋]/g, '+')
    .trim();
}

function inferPortalEventGuideTitle_(value) {
  var text = String(value || '').trim();
  if (!text) return '';
  var match = text.match(/([^/?#\\]+?)(?:\.[a-z0-9]+)?(?:[?#].*)?$/i);
  return match ? decodeURIComponent(match[1]).replace(/[_-]+/g, ' ').trim() : '';
}

function inferPortalEventGuideTitleFromBody_(value) {
  var text = String(value || '').trim();
  if (!text) return '';
  var labeled = text.match(/\u3010\u30a4\u30d9\u30f3\u30c8\u540d\u3011\s*([^\n]+)/);
  if (labeled) return labeled[1].trim();
  var quotedWonder = text.match(/[「\u3010\[]\s*(Wonder\s*[+＋]\s*[A-Za-z0-9\u3040-\u30ff\u3400-\u9fff ]+)\s*[」\u3011\]]/i);
  if (quotedWonder) return quotedWonder[1].trim();
  var inlineWonder = text.match(/Wonder\s*[+＋]\s*[A-Za-z0-9\u3040-\u30ff\u3400-\u9fff ]+/i);
  return inlineWonder ? inlineWonder[0].trim() : '';
}

function inferPortalGuideBodyFromRow_(rowValues) {
  var strings = rowValues.map(function(value) { return String(value || '').trim(); }).filter(Boolean);
  strings.sort(function(a, b) { return b.length - a.length; });
  for (var i = 0; i < strings.length; i += 1) {
    if (strings[i].length >= 40 && /(Wonder|W\+|\u6848\u5185|\u958b\u50ac|\u3054\u6848\u5185)/i.test(strings[i])) {
      return strings[i];
    }
  }
  return '';
}

function inferPortalGuideTitleFromRow_(rowValues) {
  var strings = rowValues.map(function(value) { return String(value || '').trim(); }).filter(Boolean);
  for (var i = 0; i < strings.length; i += 1) {
    if (strings[i].length <= 80 && /(Wonder|W\+|\u9280\u5ea7|\u4ed9\u53f0|\u798f\u5cf6|\u91d1\u6ca2)/i.test(strings[i])) {
      return strings[i];
    }
  }
  return '';
}

function inferPortalGuideUrlsFromRow_(rowValues) {
  return rowValues.map(function(value) { return String(value || ''); })
    .join('\n')
    .match(/https?:\/\/\S+/g);
}

function makePortalGuideAttachmentList_(value) {
  var urls = [];
  if (Array.isArray(value)) {
    value.forEach(function(item) {
      if (item && typeof item === 'object' && item.url) urls.push(String(item.url));
      else urls = urls.concat(String(item || '').match(/https?:\/\/\S+/g) || []);
    });
  } else {
    urls = String(value || '').match(/https?:\/\/\S+/g) || [];
  }
  return urls.map(function(url, index) {
    return { name: 'file-' + (index + 1), url: url };
  });
}

function getPortalGuideUpdatedTime_(guide) {
  var raw = guide && (guide.updatedAt || guide.createdAt);
  var date = raw instanceof Date ? raw : new Date(raw);
  var time = date.getTime();
  return isNaN(time) ? 0 : time;
}

function clearPortalScheduleCaches_() {
  try {
    CacheService.getScriptCache().removeAll([
      'PORTAL_SCHEDULE_PAYLOAD',
      'PORTAL_SCHEDULE_CACHE',
      'PORTAL_EVENT_GUIDES',
      'PORTAL_EVENT_GUIDE_CACHE',
      'SCHEDULE_PAYLOAD',
      'EVENT_GUIDES'
    ]);
  } catch (error) {
    Logger.log('Cache clear skipped: ' + error.message);
  }
}

function countPortalScheduleItems_(payload) {
  var count = 0;
  var schedules = payload && payload.schedules;
  if (!schedules || typeof schedules !== 'object') return 0;
  Object.keys(schedules).forEach(function(key) {
    var value = schedules[key];
    if (Array.isArray(value)) count += value.length;
    else if (value && Array.isArray(value.items)) count += value.items.length;
  });
  return count;
}

function deletePortalRealtimeTriggers_(handlerNames) {
  var names = handlerNames || [];
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (names.indexOf(trigger.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function openPortalFormByAnyUrl_(url) {
  var text = resolvePortalFormShortUrl_(String(url || '').trim());
  if (!text) throw new Error('Form URL is empty.');
  try {
    return FormApp.openByUrl(text);
  } catch (error) {}
  var idMatch = text.match(/\/forms\/d\/(?:e\/)?([a-zA-Z0-9_-]+)/) || text.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  var id = idMatch ? idMatch[1] : text;
  return FormApp.openById(id);
}

function resolvePortalFormShortUrl_(url) {
  if (!/^https?:\/\/forms\.gle\//i.test(url)) return url;
  try {
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      followRedirects: false,
      muteHttpExceptions: true
    });
    var headers = response.getAllHeaders();
    var location = headers.Location || headers.location;
    return location ? String(location).trim() : url;
  } catch (error) {
    return url;
  }
}
