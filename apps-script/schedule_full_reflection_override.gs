/**
 * Wonder+ Portal full schedule reflection override.
 * ASCII-only source to avoid mojibake in Apps Script.
 */

function runScheduleFullReflectionNow() {
  importLatestScheduleFormResponse();
}

function importLatestScheduleFormResponse() {
  const ss = getPortalScheduleSpreadsheetFull_();
  const form = findScheduleFormFull_();
  const responses = form.getResponses();
  if (!responses.length) throw new Error('No schedule form responses found.');

  const selected = selectBestScheduleResponseFull_(responses);
  if (!selected.text) throw new Error('Schedule text was empty.');

  writeScheduleTextAndSync_(ss, selected.text);
  Logger.log('schedule imported: ' + form.getTitle() + ' / ' + selected.response.getTimestamp() + ' / score=' + selected.score + ' / length=' + selected.text.length);
}

function getPortalScheduleSpreadsheetFull_() {
  const props = PropertiesService.getScriptProperties();
  const candidates = [
    props.getProperty('SCHEDULE_SPREADSHEET_ID'),
    props.getProperty('SPREADSHEET_ID'),
    props.getProperty('PORTAL_SPREADSHEET_ID'),
    typeof CONFIG !== 'undefined' ? CONFIG.scheduleSpreadsheetId : '',
    typeof CONFIG !== 'undefined' ? CONFIG.spreadsheetId : ''
  ].filter(Boolean);

  for (let i = 0; i < candidates.length; i += 1) {
    try {
      return SpreadsheetApp.openById(String(candidates[i]).trim());
    } catch (error) {}
  }

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  throw new Error('Schedule spreadsheet was not found. Set SCHEDULE_SPREADSHEET_ID.');
}

function findScheduleFormFull_() {
  const publishedId = '1FAIpQLSddWQGQj1KvM1lgmsBJsVfYsKefZ2xIvR-kD_lVnUdnlhjfLw';
  const publishedMatch = findScheduleFormByPublishedIdFull_(publishedId);
  if (publishedMatch) return publishedMatch;

  const explicitUrls = [
    'https://forms.gle/38UzoYYjz5DQzr5o6',
    typeof CONFIG !== 'undefined' ? CONFIG.scheduleFormUrl : ''
  ].filter(Boolean);

  for (let i = 0; i < explicitUrls.length; i += 1) {
    const form = openScheduleFormByUrlFull_(explicitUrls[i]);
    if (form) return form;
  }

  const files = DriveApp.searchFiles('title contains "\\u30b9\\u30b1\\u30b8\\u30e5\\u30fc\\u30eb\\u5171\\u6709\\u30d5\\u30a9\\u30fc\\u30e0" and mimeType = "application/vnd.google-apps.form" and trashed = false');
  let selected = null;
  while (files.hasNext()) {
    const file = files.next();
    if (!selected || file.getLastUpdated().getTime() > selected.getLastUpdated().getTime()) selected = file;
  }
  if (selected) return FormApp.openById(selected.getId());

  throw new Error('Schedule form was not found.');
}

function findScheduleFormByPublishedIdFull_(publishedId) {
  const files = DriveApp.searchFiles('mimeType = "application/vnd.google-apps.form" and trashed = false');
  while (files.hasNext()) {
    const file = files.next();
    try {
      const form = FormApp.openById(file.getId());
      const urls = [form.getPublishedUrl(), form.getEditUrl()].join('\n');
      if (urls.indexOf(publishedId) >= 0) return form;
    } catch (error) {}
  }
  return null;
}

function openScheduleFormByUrlFull_(url) {
  try {
    const resolved = resolveShortUrlFull_(url);
    const id = extractGoogleFormIdFull_(resolved) || extractGoogleFormIdFull_(url);
    if (!id) return null;
    return FormApp.openById(id);
  } catch (error) {
    return null;
  }
}

function resolveShortUrlFull_(url) {
  const value = String(url || '').trim();
  if (!/^https:\/\/forms\.gle\//.test(value)) return value;
  const response = UrlFetchApp.fetch(value, { followRedirects: false, muteHttpExceptions: true });
  return response.getHeaders().Location || value;
}

function extractGoogleFormIdFull_(url) {
  const value = String(url || '');
  const editMatch = value.match(/\/forms\/d\/([a-zA-Z0-9_-]+)/);
  if (editMatch && !/^e$/.test(editMatch[1])) return editMatch[1];
  return '';
}

function extractScheduleTextFromResponseFull_(response) {
  const answers = response.getItemResponses().map(function(itemResponse) {
    const answer = itemResponse.getResponse();
    return Array.isArray(answer) ? answer.join('\n') : String(answer || '');
  });

  const scheduleLike = answers
    .map(function(value) { return String(value || '').trim(); })
    .filter(function(value) {
      return /\d{4}\u5e74\s*\d{1,2}\u6708/.test(value) || /Wonder\+|W\+|\u30ec\u30ae\u30e5\u30e9\u30fc\u7570\u696d\u7a2e\u4ea4\u6d41\u4f1a|\u671d\u6d3b\u7570\u696d\u7a2e\u4ea4\u6d41\u4f1a|\u30df\u30fc\u30c6\u30a3\u30f3\u30b0/i.test(value);
    })
    .sort(function(a, b) { return b.length - a.length; });

  return scheduleLike[0] || answers.sort(function(a, b) { return String(b || '').length - String(a || '').length; })[0] || '';
}

function selectBestScheduleResponseFull_(responses) {
  const sorted = responses.slice().sort(function(a, b) {
    return b.getTimestamp().getTime() - a.getTimestamp().getTime();
  });

  for (let i = 0; i < sorted.length; i += 1) {
    const response = sorted[i];
    const text = extractScheduleTextFromResponseFull_(response);
    if (scoreScheduleTextFull_(text) > 0) {
      return { response: response, text: text, score: scoreScheduleTextFull_(text) };
    }
  }

  const fallback = sorted[0] || responses[responses.length - 1];
  const fallbackText = fallback ? extractScheduleTextFromResponseFull_(fallback) : '';
  return { response: fallback, text: fallbackText, score: scoreScheduleTextFull_(fallbackText) };
}

function scoreScheduleTextFull_(text) {
  const value = String(text || '');
  const monthMatches = value.match(/\d{4}\u5e74\s*\d{1,2}\u6708/g) || [];
  const sectionMatches = value.match(/Wonder\+\u958b\u50ac\u30b9\u30b1\u30b8\u30e5\u30fc\u30eb|\u30ec\u30ae\u30e5\u30e9\u30fc\u7570\u696d\u7a2e\u4ea4\u6d41\u4f1a|\u671d\u6d3b\u7570\u696d\u7a2e\u4ea4\u6d41\u4f1a|\u30df\u30fc\u30c6\u30a3\u30f3\u30b0/g) || [];
  const timeMatches = value.match(/\d{1,2}:\d{2}/g) || [];
  return monthMatches.length * 1000000 + sectionMatches.length * 100000 + timeMatches.length * 1000 + value.length;
}

function writeScheduleTextAndSync_(ss, text) {
  const pasteSheet = ss.getSheetByName('\u8cbc\u308a\u4ed8\u3051') || ss.insertSheet('\u8cbc\u308a\u4ed8\u3051');
  pasteSheet.getRange('A2').setValue(text);

  const outputSheet = ss.getSheetByName('\u6574\u5f62\u6e08\u307f') || ss.insertSheet('\u6574\u5f62\u6e08\u307f');
  ensureScheduleHeaderFull_(outputSheet);
  const rows = parseScheduleText_(text);
  logParsedScheduleCountsFull_(rows);
  mergeScheduleRowsByMonthAndType_(outputSheet, rows);
  syncScheduleRows(ss);
}

function ensureScheduleHeaderFull_(sheet) {
  const header = ['\u6708', '\u65e5', '\u66dc\u65e5', '\u958b\u50ac\u65e5', '\u4f1a\u5834', '\u30a4\u30d9\u30f3\u30c8\u540d', '\u6642\u9593', '\u4eba\u6570', '\u30e1\u30e2', '\u539f\u6587', 'Google\u30ab\u30ec\u30f3\u30c0\u30fc', 'Notion\u53cd\u6620', '\u66f4\u65b0\u30ad\u30fc'];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    return;
  }
  const existing = sheet.getRange(1, 1, 1, header.length).getValues()[0];
  if (existing.join('') === '') sheet.getRange(1, 1, 1, header.length).setValues([header]);
}

function logParsedScheduleCountsFull_(rows) {
  const counts = {};
  rows.forEach(function(row) {
    const key = scheduleMergeGroupKey_(row);
    counts[key] = (counts[key] || 0) + 1;
  });
  Logger.log('parsed schedule counts: ' + JSON.stringify(counts));
}

function mergeScheduleRowsByMonthAndType_(sheet, newRows) {
  const width = 13;
  const lastRow = sheet.getLastRow();
  const existingRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, width).getValues() : [];

  const targetGroups = {};
  newRows.forEach(function(row) {
    const key = scheduleMergeGroupKey_(row);
    targetGroups[key] = true;
  });

  const keptRows = existingRows.filter(function(row) {
    const key = scheduleMergeGroupKey_(row);
    return !targetGroups[key];
  });

  const mergedRows = keptRows.concat(dedupeScheduleRowsFull_(newRows));
  mergedRows.sort(function(a, b) {
    const da = a[3] instanceof Date ? a[3].getTime() : new Date(a[3]).getTime();
    const db = b[3] instanceof Date ? b[3].getTime() : new Date(b[3]).getTime();
    return da - db || String(a[5]).localeCompare(String(b[5]), 'ja') || String(a[6]).localeCompare(String(b[6]), 'ja');
  });

  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, width).clearContent();
  if (mergedRows.length) sheet.getRange(2, 1, mergedRows.length, width).setValues(mergedRows);
}

function scheduleMergeGroupKey_(row) {
  const date = row[3] instanceof Date ? row[3] : new Date(row[3]);
  const month = isNaN(date.getTime()) ? String(row[0] || '') : String(date.getMonth() + 1);
  return month + ':' + scheduleTypeKey_(row);
}

function scheduleTypeKey_(row) {
  const text = row.map(function(value) { return String(value || ''); }).join(' ');
  if (new RegExp('\u30df\u30fc\u30c6\u30a3\u30f3\u30b0|MTG|meeting', 'i').test(text)) return 'meeting';
  if (new RegExp('\u30df\u30e9\u30a4\u30d0', 'i').test(text)) return 'miraiba';
  if (new RegExp('\u671d\u6d3b|\u30ec\u30ae\u30e5\u30e9\u30fc').test(text)) return 'regular';
  if (/Wonder|Wonder\+|W\+|Leaders|Ladies|Lady|Story|Gravity|Beauty|Finance|Entertainment|Executive|CXO|CxO|Alliance|Night|Real Estate|\+100|100/i.test(text)) return 'wonder';
  return 'regular';
}

function parseScheduleText_(text) {
  const rawText = normalizeScheduleTextBlocksFull_(String(text || '').replace(/\r\n?/g, '\n'));
  let currentYear = new Date().getFullYear();
  let currentMonth = null;
  let currentCategory = 'wonder';
  let currentDay = '';
  let currentDow = '';
  const rows = [];

  rawText.split('\n').forEach(function(originalLine) {
    let line = normalizeScheduleLineFull_(originalLine);
    if (!line) return;
    if (/^[-\u30fc\u2014\u2015]+$/.test(line)) return;
    if (/^["']+$/.test(line)) return;
    if (/^\*\s*\d+/.test(line) && !/\d{1,2}:\d{2}/.test(line)) return;

    const monthHeading = line.match(/(?:(\d{4})\u5e74\s*)?(\d{1,2})\u6708/);
    if (monthHeading && /Wonder|W\+|\u30ec\u30ae\u30e5\u30e9\u30fc|\u671d\u6d3b|\u30df\u30fc\u30c6\u30a3\u30f3\u30b0|\u30df\u30e9\u30a4\u30d0|\u30b9\u30b1\u30b8\u30e5\u30fc\u30eb|\u958b\u50ac/i.test(line)) {
      if (monthHeading[1]) currentYear = Number(monthHeading[1]);
      currentMonth = Number(monthHeading[2]);
      currentDay = '';
      currentDow = '';
      if (/\u30ec\u30ae\u30e5\u30e9\u30fc/.test(line)) currentCategory = 'regular';
      else if (/\u671d\u6d3b/.test(line)) currentCategory = 'morning';
      else if (/\u30df\u30fc\u30c6\u30a3\u30f3\u30b0/.test(line)) currentCategory = 'meeting';
      else if (/\u30df\u30e9\u30a4\u30d0/.test(line)) currentCategory = 'miraiba';
      else currentCategory = 'wonder';
      return;
    }

    const headingName = line.replace(/^[\u300a<\[\(\s]+|[\u300b>\]\)\s]+$/g, '');
    if (/^\u30ec\u30ae\u30e5\u30e9\u30fc\u7570\u696d\u7a2e\u4ea4\u6d41\u4f1a$/.test(headingName)) {
      currentCategory = 'regular';
      currentDay = '';
      return;
    }
    if (/^\u671d\u6d3b\u7570\u696d\u7a2e\u4ea4\u6d41\u4f1a$/.test(headingName)) {
      currentCategory = 'morning';
      currentDay = '';
      return;
    }
    if (/^\u30df\u30fc\u30c6\u30a3\u30f3\u30b0$/.test(headingName)) {
      currentCategory = 'meeting';
      currentDay = '';
      return;
    }
    if (/^\u30df\u30e9\u30a4\u30d0$/.test(headingName)) {
      currentCategory = 'miraiba';
      currentDay = '';
      return;
    }

    const parsedDate = extractScheduleDateFull_(line, currentMonth);
    let body = line;
    if (parsedDate) {
      currentMonth = parsedDate.month || currentMonth;
      currentDay = parsedDate.day;
      currentDow = parsedDate.dow || '';
      body = parsedDate.body;
    }

    if (!currentMonth || !currentDay) return;

    if (currentCategory === 'regular' || currentCategory === 'morning' || currentCategory === 'meeting' || currentCategory === 'miraiba') {
      const simple = parseSimpleScheduleLineFull_(body, currentCategory);
      if (simple) rows.push(scheduleRowFull_(currentYear, currentMonth, currentDay, currentDow, simple.place, simple.name, simple.time, '', simple.note, line));
      return;
    }

    parseWonderScheduleBodyFull_(body).forEach(function(item) {
      rows.push(scheduleRowFull_(currentYear, currentMonth, currentDay, currentDow, item.place, item.name, item.time, item.capacity, item.note, line));
    });
  });

  return dedupeScheduleRowsFull_(rows);
}

function normalizeScheduleTextBlocksFull_(text) {
  return String(text || '')
    .replace(/(20\d{2}\u5e74\s*\d{1,2}\u6708\s*Wonder\+?[^\n]*)/gi, '\n$1\n')
    .replace(/(\n\s*)(\u30ec\u30ae\u30e5\u30e9\u30fc\u7570\u696d\u7a2e\u4ea4\u6d41\u4f1a)(\s*\n)/g, '\n$2\n')
    .replace(/(\n\s*)(\u671d\u6d3b\u7570\u696d\u7a2e\u4ea4\u6d41\u4f1a)(\s*\n)/g, '\n$2\n')
    .replace(/(\n\s*)(\u30df\u30fc\u30c6\u30a3\u30f3\u30b0)(\s*\n)/g, '\n$2\n')
    .replace(/(\n\s*)(\u30df\u30e9\u30a4\u30d0)(\s*\n)/g, '\n$2\n');
}

function normalizeScheduleLineFull_(line) {
  return String(line || '')
    .replace(/\u3000/g, ' ')
    .replace(/[\uff0b]/g, '+')
    .replace(/[\u301c\uff5e\uff0d\u2014\u2015]/g, '-')
    .replace(/--+/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}

function extractScheduleDateFull_(line, currentMonth) {
  let m = line.match(/^(\d{1,2})\/(\d{1,2})[\uff08(]([^\uff09)]+)[\uff09)]\s*(.*)$/);
  if (m) return { month: Number(m[1]), day: Number(m[2]), dow: m[3], body: m[4].trim() };

  m = line.match(/^(\d{1,2})[\uff08(]([^\uff09)]+)[\uff09)]\s*(.*)$/);
  if (m) return { month: currentMonth, day: Number(m[1]), dow: m[2], body: m[3].trim() };

  m = line.match(/^(\d{1,2})\u65e5\s*(.*)$/);
  if (m) return { month: currentMonth, day: Number(m[1]), dow: '', body: m[2].trim() };

  return null;
}

function parseSimpleScheduleLineFull_(body, category) {
  const text = String(body || '').replace(/[\u301c\uff5e\uff0d\u2014\u2015]/g, '-');
  const timeMatch = text.match(/(\d{1,2}:\d{2})(?:\s*-\s*(\d{1,2}:\d{2}))?/);
  if (!timeMatch) return null;

  const names = {
    regular: '\u30ec\u30ae\u30e5\u30e9\u30fc\u7570\u696d\u7a2e\u4ea4\u6d41\u4f1a',
    morning: '\u671d\u6d3b\u7570\u696d\u7a2e\u4ea4\u6d41\u4f1a',
    meeting: '\u30df\u30fc\u30c6\u30a3\u30f3\u30b0',
    miraiba: '\u30df\u30e9\u30a4\u30d0'
  };

  return {
    place: '',
    name: names[category] || '\u4e88\u5b9a',
    time: timeMatch[2] ? timeMatch[1] + '-' + timeMatch[2] : timeMatch[1],
    note: text.replace(timeMatch[0], '').replace(/[\uff08(]\d+\u540d[\uff09)]/g, '').trim()
  };
}

function parseWonderScheduleBodyFull_(body) {
  const source = String(body || '')
    .replace(/[\u301c\uff5e\uff0d\u2014\u2015]/g, '-')
    .replace(/--+/g, '-')
    .replace(/[\uff06]/g, '&')
    .trim();
  const items = [];
  const timeRegex = /(\d{1,2}:\d{2})(?:\s*-\s*(\d{1,2}:\d{2}))?/g;
  let match;
  let previousEnd = 0;
  let previousPlace = '';

  while ((match = timeRegex.exec(source)) !== null) {
    const titleRaw = cleanWonderTitleCandidateFull_(source.slice(previousEnd, match.index));
    const start = match[1];
    const end = match[2] || defaultWonderEndTimeFull_(start);
    const after = source.slice(timeRegex.lastIndex);
    const capacityMatch = after.match(/^\s*[\uff08(](?:\u5408\u8a08)?(\d+)\s*\u540d[\uff09)]/);
    const capacity = capacityMatch ? capacityMatch[1] : '';
    const afterCapacityIndex = timeRegex.lastIndex + (capacityMatch ? capacityMatch[0].length : 0);
    const untilNextTime = source.slice(afterCapacityIndex);
    const nextTimeIndex = untilNextTime.search(/\d{1,2}:\d{2}/);
    const between = nextTimeIndex >= 0 ? untilNextTime.slice(0, nextTimeIndex) : untilNextTime;
    const suffixMatch = between.match(/^\s*(?:[\u3001,]\s*)?(\+|W\+|Wonder\+)\s*([A-Za-z][A-Za-z ]*|100\u4eba?|CXO|CxO)/i);
    let title = titleRaw;
    if (suffixMatch) title += '+' + suffixMatch[2].trim();

    const parsed = parseWonderTitleFull_(title, previousPlace);
    if (parsed.place || parsed.name) {
      items.push({
        place: parsed.place,
        name: parsed.name,
        time: start + '-' + end,
        capacity: capacity,
        note: extractWonderNoteFull_(between)
      });
      previousPlace = parsed.place || previousPlace;
    }

    previousEnd = afterCapacityIndex;
  }

  return dedupeWonderItemsFull_(items);
}

function cleanWonderTitleCandidateFull_(value) {
  let text = String(value || '')
    .replace(/^[\u3001,\s&]+|[\u3001,\s&]+$/g, '')
    .trim();
  if (!text) return '';

  const segments = text
    .split(/[\u3001,]/)
    .map(function(segment) {
      return String(segment || '')
        .replace(/^\*[^\u3001,]*$/g, '')
        .replace(/^\*.*?(?=(?:[^\s\u3001,]+)?(?:Wonder|W\s*\+|W\+|Leaders|Ladies|Lady|Story|Gravity|Beauty|Finance|Entertainment|Executive|CXO|CxO|Alliance|Night|Real Estate|\+100|100\u4eba))/i, '')
        .replace(/^[\u3001,\s&]+|[\u3001,\s&]+$/g, '')
        .trim();
    })
    .filter(Boolean);

  text = segments.length ? segments[segments.length - 1] : text;
  return text
    .replace(/^\*.*?[\u3001,]\s*/g, '')
    .replace(/^[\u3001,\s&]+|[\u3001,\s&]+$/g, '')
    .trim();
}

function parseWonderTitleFull_(title, previousPlace) {
  let core = String(title || '')
    .replace(/[\uff08(](?:\u5408\u8a08)?\d+\s*\u540d[\uff09)]/g, '')
    .replace(/\uff0a.*$/, '')
    .replace(/\s+/g, ' ')
    .trim();

  core = core.replace(/W\s*\+/i, '+');
  const plusIndex = core.indexOf('+');
  if (plusIndex >= 0) {
    const place = core.slice(0, plusIndex).replace(/Wonder$/i, '').trim() || previousPlace || '';
    const suffix = core.slice(plusIndex + 1).trim().replace(/^100\u4eba?$/, '100');
    return { place: place, name: suffix ? 'Wonder+' + suffix : 'Wonder+' };
  }

  core = core.replace(/Wonder\s*\+?/i, '').trim();
  return { place: core || previousPlace || '', name: 'Wonder+' };
}

function extractWonderNoteFull_(value) {
  const note = String(value || '').match(/\uff0a([^\u3001,\d]+)/);
  return note ? note[1].trim() : '';
}

function defaultWonderEndTimeFull_(start) {
  const parts = String(start || '').split(':').map(Number);
  if (parts.length !== 2 || isNaN(parts[0]) || isNaN(parts[1])) return '';
  const date = new Date(2000, 0, 1, parts[0], parts[1]);
  date.setMinutes(date.getMinutes() + 90);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'H:mm');
}

function scheduleRowFull_(year, month, day, dow, place, name, time, capacity, note, raw) {
  const date = new Date(Number(year), Number(month) - 1, Number(day));
  return [
    Number(month) || '',
    Number(day) || '',
    dow || '',
    Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    place || '',
    name || '',
    time || '',
    capacity || '',
    note || '',
    raw || '',
    '',
    '',
    ''
  ];
}

function dedupeScheduleRowsFull_(rows) {
  const seen = {};
  return rows.filter(function(row) {
    const key = [row[3], row[4], row[5], row[6], scheduleTypeKey_(row)].join('|');
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function dedupeWonderItemsFull_(items) {
  const seen = {};
  return items.filter(function(item) {
    const key = [item.place, item.name, item.time].join('|');
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function debugScheduleSheetCountsNow() {
  const ss = getPortalScheduleSpreadsheetFull_();
  const sheet = ss.getSheetByName('\u6574\u5f62\u6e08\u307f');
  if (!sheet) throw new Error('formatted sheet not found');
  const values = sheet.getDataRange().getValues();
  const counts = {};
  for (let i = 1; i < values.length; i += 1) {
    const row = values[i];
    const date = row[3] instanceof Date ? row[3] : new Date(row[3]);
    if (isNaN(date.getTime())) continue;
    const key = (date.getMonth() + 1) + ':' + scheduleTypeKey_(row);
    counts[key] = (counts[key] || 0) + 1;
  }
  Logger.log('sheet schedule counts: ' + JSON.stringify(counts));
  if (typeof buildPortalSchedulePayload_ === 'function') {
    const payload = buildPortalSchedulePayload_();
    const apiCounts = {};
    Object.keys(payload.schedules || {}).forEach(function(key) {
      const items = payload.schedules[key].items || [];
      if (items.length) apiCounts[key] = items.length;
    });
    Logger.log('api payload counts: ' + JSON.stringify(apiCounts));
  }
}

/**
 * Final API override.
 *
 * The portal must reflect every row produced in the formatted schedule sheet.
 * Older API builders deduped items by title/time and could drop meetings,
 * regular events, or same-day Wonder+ sessions. This late-loading override
 * keeps the sheet as the source of truth and only skips rows without a valid
 * date/category.
 */
function buildPortalSchedulePayload_() {
  return buildPortalSchedulePayloadFromFormattedSheetFull_();
}

function buildPortalSchedulePayloadFromFormattedSheetFull_() {
  const ss = getPortalScheduleSpreadsheetFull_();
  const sheet = ss.getSheetByName('\u6574\u5f62\u6e08\u307f');
  const schedules = portalEmptySchedulesFinal_();
  if (!sheet) {
    return {
      updatedAt: new Date().toISOString(),
      schedules: schedules,
      eventGuides: typeof buildPortalEventGuidesSafe_ === 'function' ? buildPortalEventGuidesSafe_() : {},
      error: 'formatted schedule sheet not found'
    };
  }

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i += 1) {
    const row = values[i];
    const date = row[3] instanceof Date ? row[3] : new Date(row[3]);
    if (isNaN(date.getTime())) continue;

    const category = scheduleTypeKey_(row);
    const key = portalScheduleKey_(date.getMonth() + 1, category);
    if (!key || !schedules[key]) continue;

    const place = String(row[4] || '').trim();
    const eventName = String(row[5] || '').trim();
    const scheduleTime = typeof portalFormatScheduleTime_ === 'function'
      ? portalFormatScheduleTime_(row[6])
      : String(row[6] || '').trim();
    const note = String(row[8] || '').trim();
    const raw = String(row[9] || '').trim();
    const title = typeof portalScheduleTitle_ === 'function'
      ? portalScheduleTitle_(place, eventName)
      : [place, eventName].filter(Boolean).join(' ');
    const displayTime = portalDisplayTimeRangeFinal_(scheduleTime, raw);
    const displayNote = portalDisplayNoteFinal_(note);

    schedules[key].items.push({
      day: date.getDate(),
      title: title,
      meta: [displayTime, displayNote].filter(Boolean).join(' / ')
    });
  }

  Object.keys(schedules).forEach(function(key) {
    schedules[key].items.sort(function(a, b) {
      return a.day - b.day || String(a.title).localeCompare(String(b.title), 'ja') || String(a.meta).localeCompare(String(b.meta), 'ja');
    });
  });

  return {
    updatedAt: new Date().toISOString(),
    schedules: schedules,
    eventGuides: typeof buildPortalEventGuidesSafe_ === 'function' ? buildPortalEventGuidesSafe_() : {}
  };
}
