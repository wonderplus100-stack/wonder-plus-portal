/**
 * Wonder+ Portal event guide realtime refresh add-on.
 *
 * Paste this file at the bottom of the existing Apps Script project, save it,
 * run installEventGuideRealtimeTrigger() once, then deploy a new web app
 * version. The form submit trigger writes normalized guide rows and warms the
 * portal data so the next portal API request can show the new guide.
 */

var PORTAL_EVENT_GUIDE_REALTIME_SHEET_NAME = 'イベント案内';
var PORTAL_EVENT_GUIDE_REALTIME_HEADERS = [
  '登録日時',
  'イベント名',
  '月',
  '案内文',
  '素材URL',
  '添付ファイルURL',
  'カテゴリ',
  '更新日時'
];

function installEventGuideRealtimeTrigger() {
  var formUrl = getPortalEventGuideFormUrl_();
  if (!formUrl) {
    throw new Error('EVENT_GUIDE_FORM_URL or CONFIG.eventGuideFormUrl is not set.');
  }
  deletePortalRealtimeTriggers_(['onEventGuideFormSubmit']);
  var form = openPortalFormByAnyUrl_(formUrl);
  ScriptApp.newTrigger('onEventGuideFormSubmit')
    .forForm(form)
    .onFormSubmit()
    .create();
  Logger.log('Event guide realtime trigger installed: ' + formUrl);
}

function onEventGuideFormSubmit(e) {
  var answers = getPortalNamedFormAnswers_(e);
  var ss = getPortalEventGuideSpreadsheet_();
  var sheet = ensurePortalEventGuideRealtimeSheet_(ss);
  var row = normalizePortalEventGuideAnswer_(answers);
  sheet.appendRow([
    row.createdAt,
    row.title,
    row.month,
    row.body,
    row.materialUrl,
    row.attachmentUrls,
    row.category,
    row.updatedAt
  ]);
  SpreadsheetApp.flush();
  var refresh = refreshPortalEventGuideCache_();
  Logger.log(JSON.stringify({
    ok: true,
    eventGuideTitle: row.title,
    month: row.month,
    refresh: refresh
  }));
  return { ok: true, guide: row, refresh: refresh };
}

function refreshPortalEventGuideCache_() {
  var props = PropertiesService.getScriptProperties();
  var now = new Date();
  var count = 0;
  var scheduleCount = 0;
  try {
    var guides = typeof buildPortalEventGuides_ === 'function'
      ? buildPortalEventGuides_()
      : buildPortalEventGuidesRealtimeFallback_();
    count = Array.isArray(guides) ? guides.length : 0;
  } catch (error) {
    Logger.log('Event guide refresh skipped: ' + error.message);
  }
  try {
    if (typeof buildPortalSchedulePayload_ === 'function') {
      var payload = buildPortalSchedulePayload_();
      scheduleCount = countPortalScheduleItems_(payload);
    }
  } catch (error) {
    Logger.log('Schedule payload warmup skipped: ' + error.message);
  }
  clearPortalScheduleCaches_();
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
  var sheet = ss.getSheetByName(PORTAL_EVENT_GUIDE_REALTIME_SHEET_NAME) ||
    ss.getSheets().filter(function(candidate) {
      return /イベント|案内|event|guide|form|フォーム|回答/i.test(candidate.getName());
    })[0] ||
    ss.insertSheet(PORTAL_EVENT_GUIDE_REALTIME_SHEET_NAME);
  var existing = sheet.getLastRow() > 0
    ? sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), PORTAL_EVENT_GUIDE_REALTIME_HEADERS.length)).getValues()[0]
    : [];
  var hasUsefulHeader = existing.some(function(value) {
    return String(value || '').trim();
  });
  if (!hasUsefulHeader) {
    sheet.getRange(1, 1, 1, PORTAL_EVENT_GUIDE_REALTIME_HEADERS.length)
      .setValues([PORTAL_EVENT_GUIDE_REALTIME_HEADERS]);
  }
  return sheet;
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

function normalizePortalEventGuideAnswer_(answers) {
  var body = stripPortalEventGuideMaterialFooter_(pickPortalAnswer_(answers, [
    '案内文', '案内本文', '本文', '投稿文', '文章', 'メッセージ'
  ]));
  var title = pickPortalAnswer_(answers, [
    'イベント名', 'イベントタイトル', '案内タイトル', 'イベント', 'タイトル', '案内名', 'イベントブランド'
  ]) || inferPortalEventGuideTitleFromBody_(body);
  var month = parsePortalEventGuideMonth_(pickPortalAnswer_(answers, [
    '月', '対象月', '開催月', '年月', '開催年月'
  ]) || body || title);
  var attachmentUrls = pickPortalAnswer_(answers, [
    '添付ファイル', '添付ファイルURL', 'ファイル', '画像', '画像・動画', '素材ファイル'
  ]);
  var materialUrl = pickPortalAnswer_(answers, [
    '素材URL', '素材リンク', 'Drive URL', 'Drive資料URL', '資料URL', 'URL'
  ]);
  var category = pickPortalAnswer_(answers, ['カテゴリ', '種別', 'ブランド']) || 'wonder';
  var now = new Date();
  return {
    createdAt: now,
    updatedAt: now,
    title: title || inferPortalEventGuideTitle_(materialUrl || attachmentUrls) || 'Wonder+',
    month: month,
    body: body,
    materialUrl: materialUrl,
    attachmentUrls: attachmentUrls,
    category: normalizePortalEventGuideCategory_(category)
  };
}

function buildPortalEventGuidesRealtimeFallback_() {
  var ss = getPortalEventGuideSpreadsheet_();
  var sheet = ensurePortalEventGuideRealtimeSheet_(ss);
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0].map(function(value) { return String(value || '').trim(); });
  var guides = [];
  for (var i = 1; i < values.length; i += 1) {
    var row = {};
    headers.forEach(function(header, index) {
      if (header) row[header] = values[i][index];
    });
    var body = stripPortalEventGuideMaterialFooter_(pickPortalAnswer_(row, ['案内文', '案内本文', '本文', 'Body', 'body', 'Text', 'text']));
    var title = pickPortalAnswer_(row, ['イベント名', 'イベントタイトル', '案内タイトル', 'EventName', 'eventName', 'イベント', 'タイトル', '案内名']);
    if (!title && !body) continue;
    guides.push({
      title: title || inferPortalEventGuideTitleFromBody_(body),
      eventName: title || inferPortalEventGuideTitleFromBody_(body),
      month: parsePortalEventGuideMonth_(pickPortalAnswer_(row, ['月', 'Month', 'eventMonth', '対象月', '開催月'])),
      body: body,
      text: body,
      materialUrl: pickPortalAnswer_(row, ['素材URL', 'MaterialUrl', 'materialUrl', 'Drive資料URL', '資料URL', 'URL', 'url', 'Drive URL']),
      attachmentUrls: pickPortalAnswer_(row, ['添付ファイルURL', '添付ファイル', 'AttachmentUrls', 'attachmentUrls', 'ファイル', '画像', '素材ファイル']),
      category: normalizePortalEventGuideCategory_(pickPortalAnswer_(row, ['カテゴリ', 'Category', 'category', '種別', 'ブランド']) || 'wonder'),
      updatedAt: pickPortalAnswer_(row, ['更新日時', 'UpdatedAt', 'updatedAt', 'タイムスタンプ', '登録日時'])
    });
  }
  return guides;
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
      if (normalizedKeys[j].normalized === needle || normalizedKeys[j].normalized.indexOf(needle) !== -1) {
        return String(object[normalizedKeys[j].key] || '').trim();
      }
    }
  }
  return '';
}

function normalizePortalTextKey_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[＋]/g, '+')
    .replace(/\s+/g, '')
    .replace(/[【】\[\]（）()・:：/／_-]/g, '');
}

function parsePortalEventGuideMonth_(value) {
  var text = String(value || '');
  var ym = text.match(/20\d{2}\s*年\s*(1[0-2]|[1-9])\s*月/);
  if (ym) return Number(ym[1]);
  var month = text.match(/(1[0-2]|[1-9])\s*月/);
  if (month) return Number(month[1]);
  var slash = text.match(/(?:^|[^\d])(1[0-2]|[1-9])\s*[\/.-]\s*\d{1,2}(?:[^\d]|$)/);
  if (slash) return Number(slash[1]);
  return 0;
}

function normalizePortalEventGuideCategory_(value) {
  var text = String(value || '').toLowerCase();
  if (/regular|レギュラー|異業種/.test(text)) return 'regular';
  if (/morning|朝活/.test(text)) return 'morning';
  if (/meeting|ミーティング/.test(text)) return 'meeting';
  if (/miraiba|mirai|ミライバ/.test(text)) return 'miraiba';
  return 'wonder';
}

function stripPortalEventGuideMaterialFooter_(value) {
  return String(value || '')
    .replace(/\r\n/g, '\n')
    .replace(/\n{0,2}【素材】[\s\S]*$/m, '')
    .replace(/\n{0,2}\[素材\][\s\S]*$/mi, '')
    .replace(/\n{0,2}素材\s*[:：]\s*(?:https?:\/\/\S+|Drive|Google Drive)[\s\S]*$/mi, '')
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
  var labeled = text.match(/【イベント名】\s*([^\n]+)/);
  if (labeled) return labeled[1].trim();
  var quotedWonder = text.match(/[「『【\[]\s*(Wonder\s*[＋+]\s*[A-Za-z0-9一-龠ぁ-んァ-ンー ]+)\s*[」』】\]]/i);
  if (quotedWonder) return quotedWonder[1].trim();
  var inlineWonder = text.match(/Wonder\s*[＋+]\s*[A-Za-z0-9一-龠ぁ-んァ-ンー ]+/i);
  return inlineWonder ? inlineWonder[0].trim() : '';
}

function clearPortalScheduleCaches_() {
  try {
    var cache = CacheService.getScriptCache();
    cache.removeAll([
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
