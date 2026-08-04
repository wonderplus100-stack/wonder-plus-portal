/**
 * Wonder+ Portal schedule API dynamic month override.
 *
 * ASCII-safe source. Japanese labels are written with Unicode escapes so
 * Apps Script paste/import does not corrupt category matching.
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
  const sheet = getSheetByNamePartFinal_(ss, ['\u6574\u5f62\u6e08\u307f', '\u6574\u5f62']);
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
    const scheduleTime = portalFormatScheduleTime_(row[6]);
    const note = String(row[8] || '').trim();
    const raw = String(row[9] || '').trim();
    const title = portalScheduleTitle_(place, eventName);
    const displayTime = portalDisplayTimeRangeFinal_(scheduleTime, raw);
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

function getSheetByNamePartFinal_(ss, names) {
  for (let i = 0; i < names.length; i += 1) {
    const exact = ss.getSheetByName(names[i]);
    if (exact) return exact;
  }
  const sheets = ss.getSheets();
  for (let s = 0; s < sheets.length; s += 1) {
    const name = String(sheets[s].getName() || '');
    for (let i = 0; i < names.length; i += 1) {
      if (name.indexOf(names[i]) >= 0) return sheets[s];
    }
  }
  return null;
}

function portalScheduleTypeKeyDynamic_(row) {
  const text = row.map(function(value) { return String(value || ''); }).join(' ');
  if (new RegExp('\u30df\u30fc\u30c6\u30a3\u30f3\u30b0|MTG|meeting', 'i').test(text)) return 'meeting';
  if (new RegExp('\u30df\u30e9\u30a4\u30d0', 'i').test(text)) return 'miraiba';
  if (new RegExp('\u671d\u6d3b|\u30ec\u30ae\u30e5\u30e9\u30fc\u7570\u696d\u7a2e\u4ea4\u6d41\u4f1a|\u30ec\u30ae\u30e5\u30e9\u30fc').test(text)) return 'regular';
  if (/Wonder|Wonder\+|W\+|Leaders|Ladies|Lady|Story|Gravity|Beauty|Finance|Entertainment|Executive|CXO|CxO|Alliance|Night|Real Estate|\+100|100/i.test(text)) return 'wonder';
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
    regular: '\u30ec\u30ae\u30e5\u30e9\u30fc',
    miraiba: '\u30df\u30e9\u30a4\u30d0',
    meeting: '\u30df\u30fc\u30c6\u30a3\u30f3\u30b0'
  };
  ['wonder', 'regular', 'miraiba', 'meeting'].forEach(function(category) {
    for (let month = 1; month <= 12; month += 1) {
      schedules[portalScheduleKey_(month, category)] = {
        month: month,
        label: month + '\u6708 ' + labels[category],
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
  const timeText = String(time || '').trim().replace(/[\u301c\uff5e\uff0d\u2014\u2015]/g, '-');
  if (/\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}/.test(timeText)) return timeText.replace(/\s*-\s*/g, '-');
  const source = String(raw || '').replace(/[\u301c\uff5e\uff0d\u2014\u2015]/g, '-');
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
    .replace(/\(?\d+\s*(\u540d|\u4eba|\u5e2d|\u67a0)\)?/g, '')
    .replace(/\s*\/\s*/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}
