/**
 * Wonder+ Portal full schedule reflection override.
 *
 * Paste this file at the very bottom of the Apps Script project.
 * It replaces only the schedule import/parser helpers. Other portal features
 * such as login, board, minutes, and event-guide APIs are left untouched.
 */

function writeScheduleTextAndSync_(ss, text) {
  const pasteSheet = ss.getSheetByName('貼り付け') || ss.insertSheet('貼り付け');
  pasteSheet.getRange('A2').setValue(text);

  const outputSheet = ss.getSheetByName('整形済み') || ss.insertSheet('整形済み');
  const rows = parseScheduleText_(text);
  mergeScheduleRowsByMonthAndType_(outputSheet, rows);
  syncScheduleRows(ss);
}

function mergeScheduleRowsByMonthAndType_(sheet, newRows) {
  const width = 13;
  const lastRow = sheet.getLastRow();
  const existingRows = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, width).getValues()
    : [];

  const targetGroups = {};
  newRows.forEach(function(row) {
    const key = scheduleMergeGroupKey_(row);
    if (key) targetGroups[key] = true;
  });

  const keptRows = existingRows.filter(function(row) {
    const key = scheduleMergeGroupKey_(row);
    return key && !targetGroups[key];
  });

  const mergedRows = keptRows.concat(newRows);
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, width).clearContent();
  if (mergedRows.length) sheet.getRange(2, 1, mergedRows.length, width).setValues(mergedRows);
}

function scheduleMergeGroupKey_(row) {
  const month = Number(row[0] || '');
  if (!month) return '';
  return month + ':' + scheduleTypeKey_(row);
}

function scheduleTypeKey_(row) {
  const text = [row[4], row[5], row[8], row[9]].join(' ');
  if (/ミーティング|MTG|meeting/i.test(text)) return 'meeting';
  if (/ミライバ/i.test(text)) return 'miraiba';
  if (/朝活|レギュラー/.test(text)) return 'regular';
  if (/Wonder|Wonder\+|Wonder＋|W\+|W＋|Leaders|Ladies|Lady|Story|Gravity|Beauty|Finance|Entertainment|Executive|CXO|CxO|\+100|100人/i.test(text)) return 'wonder';
  return 'regular';
}

function parseScheduleText_(text) {
  const rawText = String(text || '')
    .replace(/\r/g, '')
    .replace(/[“”]/g, '"')
    .replace(/[＋]/g, '+');

  const firstHeading = rawText.match(/(\d{4})年\s*(\d{1,2})月/);
  let currentYear = firstHeading ? Number(firstHeading[1]) : new Date().getFullYear();
  let currentMonth = firstHeading ? Number(firstHeading[2]) : new Date().getMonth() + 1;
  let currentCategory = 'wonder';
  let currentDay = '';
  let currentDow = '';
  const rows = [];

  rawText.split('\n').forEach(function(originalLine) {
    let line = normalizeScheduleLineFull_(originalLine);
    if (!line) return;
    if (/^[-ー－—―]+$/.test(line)) return;
    if (/^["']+$/.test(line)) return;
    if (/^＊\s*\d+/.test(line) && !/\d{1,2}:\d{2}/.test(line)) return;

    const heading = line.match(/(?:(\d{4})年\s*)?(\d{1,2})月.*?(Wonder|W\+|W＋|レギュラー|朝活|ミーティング|ミライバ|スケジュール|開催)/i);
    if (heading) {
      if (heading[1]) currentYear = Number(heading[1]);
      currentMonth = Number(heading[2]);
      currentDay = '';
      currentDow = '';
      if (/レギュラー/.test(line)) currentCategory = 'regular';
      else if (/朝活/.test(line)) currentCategory = 'morning';
      else if (/ミーティング/.test(line)) currentCategory = 'meeting';
      else if (/ミライバ/.test(line)) currentCategory = 'miraiba';
      else if (/Wonder|W\+|W＋/i.test(line)) currentCategory = 'wonder';
      return;
    }

    if (/^《?レギュラー異業種交流会》?$/.test(line) || /^レギュラー異業種交流会$/.test(line)) {
      currentCategory = 'regular';
      currentDay = '';
      return;
    }
    if (/^《?朝活異業種交流会》?$/.test(line) || /^朝活異業種交流会$/.test(line)) {
      currentCategory = 'morning';
      currentDay = '';
      return;
    }
    if (/^《?ミーティング》?$/.test(line) || /^ミーティング$/.test(line)) {
      currentCategory = 'meeting';
      currentDay = '';
      return;
    }
    if (/^《?ミライバ》?$/.test(line) || /^ミライバ$/.test(line)) {
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

    if (!currentDay) return;

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

function normalizeScheduleLineFull_(line) {
  return String(line || '')
    .replace(/\u3000/g, ' ')
    .replace(/[＋]/g, '+')
    .replace(/[～〜–—−]/g, '-')
    .replace(/--+/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}

function extractScheduleDateFull_(line, currentMonth) {
  let m = line.match(/^(\d{1,2})\/(\d{1,2})[（(]([^）)]+)[）)]\s*(.*)$/);
  if (m) return { month: Number(m[1]), day: Number(m[2]), dow: m[3], body: m[4].trim() };

  m = line.match(/^(\d{1,2})[（(]([^）)]+)[）)]\s*(.*)$/);
  if (m) return { month: currentMonth, day: Number(m[1]), dow: m[2], body: m[3].trim() };

  m = line.match(/^(\d{1,2})日\s*(.*)$/);
  if (m) return { month: currentMonth, day: Number(m[1]), dow: '', body: m[2].trim() };

  return null;
}

function parseSimpleScheduleLineFull_(body, category) {
  const text = String(body || '').replace(/[～〜–—−]/g, '-');
  const timeMatch = text.match(/(\d{1,2}:\d{2})(?:\s*-\s*(\d{1,2}:\d{2}))?/);
  if (!timeMatch) return null;

  const names = {
    regular: 'レギュラー異業種交流会',
    morning: '朝活異業種交流会',
    meeting: 'ミーティング',
    miraiba: 'ミライバ'
  };

  return {
    place: '',
    name: names[category] || '予定',
    time: timeMatch[2] ? timeMatch[1] + '-' + timeMatch[2] : timeMatch[1],
    note: text.replace(timeMatch[0], '').replace(/[（(]\d+名[）)]/g, '').trim()
  };
}

function parseWonderScheduleBodyFull_(body) {
  const source = String(body || '')
    .replace(/[～〜–—−]/g, '-')
    .replace(/--+/g, '-')
    .replace(/＆/g, ' & ')
    .trim();
  const items = [];
  const timeRegex = /(\d{1,2}:\d{2})(?:\s*-\s*(\d{1,2}:\d{2}))?/g;
  let match;
  let previousEnd = 0;
  let previousTitle = '';

  while ((match = timeRegex.exec(source)) !== null) {
    const titleRaw = source.slice(previousEnd, match.index).replace(/^[、,，\s&]+|[、,，\s&]+$/g, '').trim();
    const start = match[1];
    const end = match[2] || defaultWonderEndTimeFull_(start);
    let after = source.slice(timeRegex.lastIndex);
    const capacityMatch = after.match(/^\s*[（(](?:合計)?(\d+)\s*名[）)]/);
    const capacity = capacityMatch ? capacityMatch[1] : '';
    let note = '';
    let nextStart = timeRegex.lastIndex + (capacityMatch ? capacityMatch[0].length : 0);
    const betweenNext = source.slice(nextStart);
    const nextTime = betweenNext.search(/\d{1,2}:\d{2}/);
    const noteSource = nextTime >= 0 ? betweenNext.slice(0, nextTime) : betweenNext;
    const noteMatch = noteSource.match(/＊([^、,，&]+)/);
    if (noteMatch) note = noteMatch[1].trim();

    let title = titleRaw || previousTitle;
    if (title) {
      if (/^\+/.test(title) && previousTitle) {
        title = previousTitle.replace(/\+.*$/, '') + title;
      }
      const parsed = parseWonderTitleFull_(title);
      items.push({
        place: parsed.place,
        name: parsed.name,
        time: start + '-' + end,
        capacity: capacity,
        note: note
      });
      previousTitle = title;
    }

    previousEnd = nextStart;
  }

  return dedupeWonderItemsFull_(items);
}

function parseWonderTitleFull_(title) {
  let core = String(title || '')
    .replace(/[（(](?:合計)?\d+\s*名[）)]/g, '')
    .replace(/[＊].*$/, '')
    .replace(/\s+/g, ' ')
    .trim();

  core = core.replace(/W\s*\+/i, '+');
  const plusIndex = core.indexOf('+');
  if (plusIndex >= 0) {
    const place = core.slice(0, plusIndex).replace(/Wonder$/i, '').trim();
    const suffix = core.slice(plusIndex + 1).trim();
    return {
      place: place || '銀座',
      name: suffix ? 'Wonder+' + suffix : 'Wonder+'
    };
  }

  core = core.replace(/Wonder\s*\+?/i, '').trim();
  return {
    place: core || '',
    name: 'Wonder+'
  };
}

function defaultWonderEndTimeFull_(start) {
  const parts = String(start || '').split(':');
  if (parts.length !== 2) return '';
  const date = new Date(2000, 0, 1, Number(parts[0]), Number(parts[1]));
  date.setMinutes(date.getMinutes() + 90);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'HH:mm');
}

function scheduleRowFull_(year, month, day, dow, place, name, time, capacity, note, raw) {
  const date = new Date(Number(year), Number(month) - 1, Number(day));
  return [
    Number(month),
    Number(day),
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
