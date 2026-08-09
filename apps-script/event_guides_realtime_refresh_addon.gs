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
var PORTAL_EVENT_GUIDE_REALTIME_HEADERS = [
  'CreatedAt',
  'Title',
  'Month',
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
  var ss = getPortalEventGuideSpreadsheet_();
  var guides = [];
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
      if (isUsablePortalEventGuide_(guide)) guides.push(guide);
    }
  });
  return dedupePortalEventGuides_(guides)
    .sort(function(a, b) { return getPortalGuideUpdatedTime_(b) - getPortalGuideUpdatedTime_(a); });
}

function buildPortalEventGuidesRealtimeFallback_() {
  return buildPortalEventGuidesSafe_();
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
    props.getProperty('EVENT_INFO_SPREADSHEET_ID'),
    props.getProperty('EVENT_GUIDE_SPREADSHEET_ID'),
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
  return candidates.length ? candidates : sheets;
}

function matchesAnyPortalGuideKey_(normalizedHeader) {
  var groups = [
    PORTAL_EVENT_GUIDE_TITLE_KEYS,
    PORTAL_EVENT_GUIDE_BODY_KEYS,
    PORTAL_EVENT_GUIDE_MONTH_KEYS,
    PORTAL_EVENT_GUIDE_MATERIAL_KEYS,
    PORTAL_EVENT_GUIDE_ATTACHMENT_KEYS,
    PORTAL_EVENT_GUIDE_CATEGORY_KEYS
  ];
  return groups.some(function(keys) {
    return keys.some(function(key) {
      var needle = normalizePortalTextKey_(key);
      return normalizedHeader === needle || normalizedHeader.indexOf(needle) !== -1 || needle.indexOf(normalizedHeader) !== -1;
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
  var body = stripPortalEventGuideMaterialFooter_(
    pickPortalAnswer_(record, PORTAL_EVENT_GUIDE_BODY_KEYS) ||
    inferPortalGuideBodyFromRow_(rowValues || [])
  );
  var title = pickPortalAnswer_(record, PORTAL_EVENT_GUIDE_TITLE_KEYS) || inferPortalEventGuideTitleFromBody_(body) || inferPortalGuideTitleFromRow_(rowValues || []);
  var materialUrl = pickPortalAnswer_(record, PORTAL_EVENT_GUIDE_MATERIAL_KEYS);
  var attachmentUrls = pickPortalAnswer_(record, PORTAL_EVENT_GUIDE_ATTACHMENT_KEYS) || inferPortalGuideUrlsFromRow_(rowValues || []);
  var month = parsePortalEventGuideMonth_(
    pickPortalAnswer_(record, PORTAL_EVENT_GUIDE_MONTH_KEYS) + '\n' + title + '\n' + body
  );
  var category = normalizePortalEventGuideCategory_(pickPortalAnswer_(record, PORTAL_EVENT_GUIDE_CATEGORY_KEYS) || title || body || 'wonder');
  var updatedAt = pickPortalAnswer_(record, ['UpdatedAt', 'updatedAt', '\u66f4\u65b0\u65e5\u6642', '\u30bf\u30a4\u30e0\u30b9\u30bf\u30f3\u30d7']) ||
    pickPortalAnswer_(record, ['CreatedAt', 'createdAt', '\u767b\u9332\u65e5\u6642', '\u9001\u4fe1\u65e5\u6642']);
  var createdAt = pickPortalAnswer_(record, ['CreatedAt', 'createdAt', '\u767b\u9332\u65e5\u6642', '\u9001\u4fe1\u65e5\u6642']) || updatedAt || new Date();
  var cleanTitle = normalizePortalEventGuideTitle_(title || inferPortalEventGuideTitle_(materialUrl || attachmentUrls) || 'Wonder+');
  var attachments = makePortalGuideAttachmentList_(attachmentUrls || materialUrl);
  return {
    title: cleanTitle,
    eventName: cleanTitle,
    name: cleanTitle,
    month: month,
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
    var key = [
      guide.month || 0,
      normalizePortalTextKey_(guide.title || guide.eventName || ''),
      normalizePortalVenueKey_((guide.title || '') + '\n' + (guide.body || guide.text || ''))
    ].join('|');
    if (!key.replace(/[|0]/g, '')) return;
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
    ['fukushima', '\u798f\u5cf6'],
    ['koriyama', '\u90e1\u5c71'],
    ['kanazawa', '\u91d1\u6ca2'],
    ['sapporo', '\u672d\u5e4c'],
    ['fukuoka', '\u798f\u5ca1'],
    ['kitakyushu', '\u5317\u4e5d\u5dde'],
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
    ['utsunomiya', '\u5b87\u90fd\u5bae'],
    ['machida', '\u753a\u7530'],
    ['shizuoka', '\u9759\u5ca1'],
    ['hamamatsu', '\u6d5c\u677e'],
    ['omiya', '\u5927\u5bae'],
    ['chiba', '\u5343\u8449'],
    ['kofu', '\u7532\u5e9c']
  ];
  for (var i = 0; i < venues.length; i += 1) {
    if (key.indexOf(venues[i][0]) !== -1 || key.indexOf(normalizePortalTextKey_(venues[i][1])) !== -1) {
      return venues[i][0];
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
