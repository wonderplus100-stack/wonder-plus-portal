/**
 * Wonder+ Portal schedule API dynamic month override.
 *
 * Paste this file at the very bottom of the existing Apps Script project.
 * It intentionally overrides only the portal schedule payload helpers so that
 * newly submitted months such as September and October are returned to the
 * portal without changing login, board, minutes, or approval features.
 */

function buildPortalSchedulePayload_() {
  return buildPortalScheduleDataSafe_();
}

function buildPortalScheduleDataSafe_() {
  const ssId = PropertiesService.getScriptProperties().getProperty('SCHEDULE_SPREADSHEET_ID');
  if (!ssId) {
    return {
      updatedAt: new Date().toISOString(),
      schedules: portalEmptySchedulesFinal_(),
      eventGuides: buildPortalEventGuidesSafe_(),
      error: 'SCHEDULE_SPREADSHEET_ID is not set'
    };
  }

  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName('整形済み') || ss.getSheetByName('謨ｴ蠖｢貂医∩') || ss.getSheets().find(function(candidate) {
    return String(candidate.getName()).indexOf('整') >= 0 || String(candidate.getName()).indexOf('済') >= 0;
  });
  if (!sheet) {
    return {
      updatedAt: new Date().toISOString(),
      schedules: portalEmptySchedulesFinal_(),
      eventGuides: buildPortalEventGuidesSafe_(),
      error: 'formatted schedule sheet was not found'
    };
  }

  const values = sheet.getDataRange().getValues();
  const schedules = portalEmptySchedulesFinal_();
  const seen = {};

  for (let i = 1; i < values.length; i += 1) {
    const row = values[i];
    const dateValue = row[3];
    if (!dateValue) continue;

    const date = dateValue instanceof Date ? dateValue : new Date(dateValue);
    if (isNaN(date.getTime())) continue;

    const month = date.getMonth() + 1;
    if (month < 1 || month > 12) continue;

    const category = portalScheduleTypeKeyDynamic_(row);
    const key = portalScheduleKey_(month, category);
    if (!key || !schedules[key]) continue;

    const place = String(row[4] || '').trim();
    const eventName = String(row[5] || '').trim();
    const startTime = portalFormatScheduleTime_(row[6]);
    const note = String(row[8] || '').trim();
    const raw = String(row[9] || '').trim();
    const title = portalScheduleTitle_(place, eventName);
    const displayTime = portalDisplayTimeRangeFinal_(startTime, raw);
    const itemKey = [date.getDate(), title, displayTime, category].join('|');
    if (seen[itemKey]) continue;
    seen[itemKey] = true;

    schedules[key].items.push({
      day: date.getDate(),
      title: title,
      meta: [displayTime, portalDisplayNoteFinal_(note)].filter(Boolean).join(' / ')
    });
  }

  Object.keys(schedules).forEach(function(key) {
    schedules[key].items.sort(function(a, b) {
      return a.day - b.day || String(a.title).localeCompare(String(b.title), 'ja');
    });
  });

  return {
    updatedAt: new Date().toISOString(),
    schedules: schedules,
    eventGuides: buildPortalEventGuidesSafe_()
  };
}

function portalScheduleTypeKeyDynamic_(row) {
  const text = row.map(function(value) { return String(value || ''); }).join(' ');
  if (/ミーティング|MTG|meeting/i.test(text)) return 'meeting';
  if (/ミライバ/i.test(text)) return 'miraiba';
  if (/朝活/.test(text)) return 'regular';
  if (/Wonder|Wonder\+|Wonder＋|W\+|W＋|Leaders|Ladies|Lady|Story|Gravity|Beauty|Finance|Entertainment|Executive|CXO|CxO|\+100|100人/i.test(text)) return 'wonder';
  return 'regular';
}

function portalScheduleKey_(month, category) {
  const safeMonth = Number(month);
  if (!Number.isInteger(safeMonth) || safeMonth < 1 || safeMonth > 12) return '';
  const suffix = {
    wonder: 'Wonder',
    regular: 'Regular',
    miraiba: 'Miraiba',
    meeting: 'Meeting'
  }[category] || 'Wonder';
  return 'month' + safeMonth + suffix;
}

function portalEmptySchedulesFinal_() {
  const schedules = {};
  const labels = {
    wonder: 'Wonder+',
    regular: 'レギュラー',
    miraiba: 'ミライバ',
    meeting: 'ミーティング'
  };
  ['wonder', 'regular', 'miraiba', 'meeting'].forEach(function(category) {
    for (let month = 1; month <= 12; month += 1) {
      schedules[portalScheduleKey_(month, category)] = {
        month: month,
        label: month + '月 ' + labels[category],
        items: []
      };
    }
  });
  return schedules;
}

function emptyPortalSchedules_() {
  return portalEmptySchedulesFinal_();
}

function portalDisplayTimeRangeFinal_(time, raw) {
  const timeText = String(time || '').trim().replace(/[〜～–—−]/g, '-');
  if (/\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}/.test(timeText)) {
    return timeText.replace(/\s*-\s*/g, '-');
  }
  const source = String(raw || '').replace(/[〜～–—−]/g, '-');
  if (timeText) {
    const escaped = timeText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const near = source.match(new RegExp(escaped + '\\s*-\\s*(\\d{1,2}:\\d{2})'));
    if (near) return timeText + '-' + near[1];
  }
  const anyRange = source.match(/(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})/);
  if (anyRange) return anyRange[1] + '-' + anyRange[2];
  return timeText;
}

function portalDisplayNoteFinal_(note) {
  return String(note || '')
    .replace(/\(?\d+\s*(名|人|席|枠)\)?/g, '')
    .replace(/\s*\/\s*/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}
