/**
 * Wonder+ Portal shared board API add-on.
 * Paste this file into the existing portal Apps Script project and redeploy the
 * web app. It provides the functions already called by the current doGet router.
 */

function getPortalBoardPosts_() {
  const sheet = getPortalBoardSheet_();
  const values = sheet.getDataRange().getValues();
  const posts = [];
  for (let i = 1; i < values.length; i += 1) {
    const row = values[i];
    const createdAt = row[0];
    const name = String(row[1] || '').trim() || 'ポータル利用者';
    const message = String(row[2] || '').trim();
    if (!message) continue;
    posts.push({
      id: String(row[3] || ''),
      name: name,
      message: message,
      createdAt: formatPortalBoardDate_(createdAt)
    });
  }
  posts.sort(function(a, b) {
    return String(b.id || '').localeCompare(String(a.id || ''));
  });
  return { ok: true, posts: posts.slice(0, 50) };
}

function createPortalBoardPost_(params) {
  params = params || {};
  const message = String(params.message || '').trim();
  if (!message) return { ok: false, message: 'message is required' };

  const sheet = getPortalBoardSheet_();
  const now = new Date();
  const post = {
    id: Utilities.getUuid(),
    name: String(params.name || '').trim() || 'ポータル利用者',
    message: message,
    createdAt: formatPortalBoardDate_(now)
  };

  sheet.appendRow([now, post.name, post.message, post.id]);
  return { ok: true, post: post };
}

function getPortalBoardSheet_() {
  const ss = getPortalBoardSpreadsheet_();
  const sheetName = '重要共有事項';
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  ensurePortalBoardHeader_(sheet);
  return sheet;
}

function getPortalBoardSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const candidates = [
    props.getProperty('BOARD_SPREADSHEET_ID'),
    props.getProperty('SCHEDULE_SPREADSHEET_ID'),
    props.getProperty('SPREADSHEET_ID'),
    props.getProperty('PORTAL_SPREADSHEET_ID')
  ].filter(Boolean);

  for (let i = 0; i < candidates.length; i += 1) {
    try {
      return SpreadsheetApp.openById(String(candidates[i]).trim());
    } catch (error) {}
  }

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  throw new Error('Portal spreadsheet was not found. Set BOARD_SPREADSHEET_ID or SCHEDULE_SPREADSHEET_ID.');
}

function ensurePortalBoardHeader_(sheet) {
  const header = ['投稿日時', '投稿者', '重要共有事項', 'ID'];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    return;
  }
  const existing = sheet.getRange(1, 1, 1, header.length).getValues()[0];
  if (existing.join('') === '') sheet.getRange(1, 1, 1, header.length).setValues([header]);
}

function formatPortalBoardDate_(value) {
  const date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return String(value || '');
  return Utilities.formatDate(date, 'Asia/Tokyo', 'M/d HH:mm');
}
