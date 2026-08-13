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
      createdAt: formatPortalBoardDate_(createdAt),
      updatedAt: formatPortalBoardDate_(row[4] || '')
    });
  }
  posts.sort(function(a, b) {
    return String(b.id || '').localeCompare(String(a.id || ''));
  });
  return { ok: true, posts: posts.slice(0, 50) };
}

function createPortalBoardPost_(params) {
  params = params || {};
  const id = String(params.id || params.postId || '').trim();
  if (id) return updatePortalBoardPost_(params);

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

  sheet.appendRow([now, post.name, post.message, post.id, '']);
  post.lineWorks = notifyImportantNoticeToLineWorks_({
    mode: 'created',
    id: post.id,
    name: post.name,
    message: post.message,
    createdAt: post.createdAt
  });
  return { ok: true, post: post };
}

function updatePortalBoardPost_(params) {
  params = params || {};
  const id = String(params.id || params.postId || '').trim();
  const message = String(params.message || '').trim();
  if (!id) return { ok: false, message: 'id is required' };
  if (!message) return { ok: false, message: 'message is required' };

  const sheet = getPortalBoardSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i += 1) {
    const row = values[i];
    if (String(row[3] || '').trim() !== id) continue;

    const now = new Date();
    const name = String(row[1] || '').trim() || String(params.name || '').trim() || 'ポータル利用者';
    sheet.getRange(i + 1, 3).setValue(message);
    sheet.getRange(i + 1, 5).setValue(now);
    const lineWorks = notifyImportantNoticeToLineWorks_({
      mode: 'updated',
      id: id,
      name: name,
      message: message,
      createdAt: formatPortalBoardDate_(row[0]),
      updatedAt: formatPortalBoardDate_(now)
    });
    return {
      ok: true,
      post: {
        id: id,
        name: name,
        message: message,
        createdAt: formatPortalBoardDate_(row[0]),
        updatedAt: formatPortalBoardDate_(now),
        lineWorks: lineWorks
      }
    };
  }
  return { ok: false, message: 'post was not found' };
}

function notifyImportantNoticeToLineWorks_(notice) {
  const result = {
    lineWorks: { ok: false, skipped: true, message: 'not attempted' },
    line: { ok: false, skipped: true, message: 'not attempted' }
  };
  try {
    result.lineWorks = sendImportantNoticeToLineWorks_(notice);
  } catch (error) {
    Logger.log('LINE WORKS important notice notification skipped: ' + error.message);
    result.lineWorks = { ok: false, skipped: true, message: error.message };
  }
  try {
    result.line = sendImportantNoticeToLine_(notice);
  } catch (error) {
    Logger.log('LINE important notice notification skipped: ' + error.message);
    result.line = { ok: false, skipped: true, message: error.message };
  }
  result.ok = Boolean(result.lineWorks.ok || result.line.ok);
  return result;
}

function sendImportantNoticeToLineWorks_(notice) {
  const botId = getPortalLineWorksBotId_();
  const accessToken = getPortalLineWorksAccessToken_();
  const channelIds = getPortalLineWorksNoticeChannelIds_();
  if (!channelIds.length) throw new Error('WONDER_PORTAL_LINEWORKS_CHANNEL_IDS is empty.');

  const text = buildImportantNoticeLineWorksText_(notice);
  const results = [];
  channelIds.forEach(function(channelId) {
    const url = 'https://www.worksapis.com/v1.0/bots/' + encodeURIComponent(botId) +
      '/channels/' + encodeURIComponent(channelId) + '/messages';
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + accessToken },
      payload: JSON.stringify({
        content: {
          type: 'text',
          text: text
        }
      }),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    const body = res.getContentText();
    results.push({ channelId: channelId, status: code, ok: code >= 200 && code < 300 });
    if (code < 200 || code >= 300) {
      Logger.log('LINE WORKS send failed: channelId=' + channelId + ' status=' + code + ' body=' + body);
    }
  });

  return {
    ok: results.every(function(result) { return result.ok; }),
    results: results
  };
}

function sendImportantNoticeToLine_(notice) {
  const accessToken = getOptionalPortalLineProperty_('WONDER_PORTAL_LINE_CHANNEL_ACCESS_TOKEN');
  if (!accessToken) return { ok: false, skipped: true, message: 'WONDER_PORTAL_LINE_CHANNEL_ACCESS_TOKEN is not set.' };

  const toIds = getPortalLineNoticeToIds_();
  if (!toIds.length) return { ok: false, skipped: true, message: 'WONDER_PORTAL_LINE_TO_IDS is empty.' };

  const text = buildImportantNoticeLineWorksText_(notice);
  const results = [];
  toIds.forEach(function(to) {
    const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + accessToken },
      payload: JSON.stringify({
        to: to,
        messages: [{ type: 'text', text: text }]
      }),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    const body = res.getContentText();
    results.push({ to: to, status: code, ok: code >= 200 && code < 300 });
    if (code < 200 || code >= 300) {
      Logger.log('LINE send failed: to=' + to + ' status=' + code + ' body=' + body);
    }
  });

  return {
    ok: results.every(function(result) { return result.ok; }),
    results: results
  };
}

function buildImportantNoticeLineWorksText_(notice) {
  const modeLabel = notice.mode === 'updated' ? '更新' : '投稿';
  const name = String(notice.name || 'Portal').trim();
  const timestamp = String(notice.updatedAt || notice.createdAt || '').trim();
  const message = String(notice.message || '').trim();
  const header = [
    '【重要事項共有専用】' + modeLabel,
    '投稿者: ' + name,
    timestamp ? '日時: ' + timestamp : ''
  ].filter(Boolean).join('\n');
  const maxBodyLength = 1800 - header.length;
  const body = message.length > maxBodyLength
    ? message.slice(0, Math.max(0, maxBodyLength - 20)) + '\n...(省略)'
    : message;
  return header + '\n\n' + body;
}

function getPortalLineWorksNoticeChannelIds_() {
  const props = PropertiesService.getScriptProperties();
  const configured = String(props.getProperty('WONDER_PORTAL_LINEWORKS_CHANNEL_IDS') || '').trim();
  const fallback = [
    'b0c28afd-26dc-55c3-531d-564ee5088d56',
    '86be568e-ca15-b86a-3623-fbcbabe68431',
    '4bec9198-ca7b-346c-092f-7d1e2d87eaf8'
  ];
  const raw = configured ? configured.split(/[\s,]+/) : fallback;
  return raw.map(function(channelId) {
    return String(channelId || '').trim();
  }).filter(Boolean);
}

function getPortalLineWorksBotId_() {
  const configured = String(PropertiesService.getScriptProperties().getProperty('WONDER_PORTAL_LINEWORKS_BOT_ID') || '').trim();
  return configured || '12897792';
}

function getPortalLineNoticeToIds_() {
  const props = PropertiesService.getScriptProperties();
  const configured = [
    props.getProperty('WONDER_PORTAL_LINE_TO_IDS'),
    props.getProperty('WONDER_PORTAL_LINE_GROUP_IDS')
  ].filter(Boolean).join(',');
  return String(configured || '').split(/[\s,]+/).map(function(to) {
    return String(to || '').trim();
  }).filter(Boolean);
}

function getOptionalPortalLineProperty_(name) {
  return String(PropertiesService.getScriptProperties().getProperty(name) || '').trim();
}

function getPortalLineWorksAccessToken_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('WONDER_PORTAL_LINEWORKS_ACCESS_TOKEN');
  if (cached) return cached;

  const clientId = getRequiredPortalLineWorksProperty_('WONDER_PORTAL_LINEWORKS_CLIENT_ID');
  const clientSecret = getRequiredPortalLineWorksProperty_('WONDER_PORTAL_LINEWORKS_CLIENT_SECRET');
  const serviceAccount = getRequiredPortalLineWorksProperty_('WONDER_PORTAL_LINEWORKS_SERVICE_ACCOUNT');
  const assertion = createPortalLineWorksJwt_(clientId, serviceAccount);
  const payload = {
    assertion: assertion,
    grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
    client_id: clientId,
    client_secret: clientSecret,
    scope: 'bot.message'
  };
  const res = UrlFetchApp.fetch('https://auth.worksmobile.com/oauth2/v2.0/token', {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('LINE WORKS token request failed: HTTP ' + code + ' ' + body);
  }
  const token = JSON.parse(body);
  if (!token.access_token) throw new Error('LINE WORKS token response did not include access_token.');
  const expiresIn = Number(token.expires_in || 3600);
  cache.put('WONDER_PORTAL_LINEWORKS_ACCESS_TOKEN', token.access_token, Math.max(60, Math.min(expiresIn - 120, 21600)));
  return token.access_token;
}

function createPortalLineWorksJwt_(clientId, serviceAccount) {
  const now = Math.floor(Date.now() / 1000);
  const header = {
    alg: 'RS256',
    typ: 'JWT'
  };
  const claims = {
    iss: clientId,
    sub: serviceAccount,
    iat: now,
    exp: now + 3600
  };
  const unsigned = base64UrlEncodePortalLineWorks_(JSON.stringify(header)) + '.' +
    base64UrlEncodePortalLineWorks_(JSON.stringify(claims));
  const privateKey = normalizePortalLineWorksPrivateKey_(getRequiredPortalLineWorksProperty_('WONDER_PORTAL_LINEWORKS_PRIVATE_KEY'));
  const signature = Utilities.computeRsaSha256Signature(unsigned, privateKey);
  return unsigned + '.' + base64UrlEncodePortalLineWorks_(signature);
}

function base64UrlEncodePortalLineWorks_(value) {
  const bytes = typeof value === 'string' ? Utilities.newBlob(value).getBytes() : value;
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/g, '');
}

function normalizePortalLineWorksPrivateKey_(key) {
  const value = String(key || '').trim().replace(/\\n/g, '\n');
  if (value.indexOf('BEGIN PRIVATE KEY') !== -1) return value;
  return '-----BEGIN PRIVATE KEY-----\n' + value.replace(/\s+/g, '\n') + '\n-----END PRIVATE KEY-----';
}

function getRequiredPortalLineWorksProperty_(name) {
  const value = String(PropertiesService.getScriptProperties().getProperty(name) || '').trim();
  if (!value) throw new Error(name + ' is not set.');
  return value;
}

function testImportantNoticeLineWorks() {
  const result = sendImportantNoticeToLineWorks_({
    mode: 'created',
    id: 'test-' + Date.now(),
    name: 'Wonder+ Portal Test',
    message: 'LINE WORKS important notice notification test.',
    createdAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d HH:mm')
  });
  Logger.log(JSON.stringify(result));
  return result;
}

function testImportantNoticeLine() {
  const result = sendImportantNoticeToLine_({
    mode: 'created',
    id: 'line-test-' + Date.now(),
    name: 'Wonder+ Portal Test',
    message: 'LINE important notice notification test.',
    createdAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d HH:mm')
  });
  Logger.log(JSON.stringify(result));
  return result;
}

function debugImportantNoticeNotificationSetup() {
  const props = PropertiesService.getScriptProperties();
  const lineWorksProperties = [
    'WONDER_PORTAL_LINEWORKS_BOT_ID',
    'WONDER_PORTAL_LINEWORKS_CLIENT_ID',
    'WONDER_PORTAL_LINEWORKS_CLIENT_SECRET',
    'WONDER_PORTAL_LINEWORKS_SERVICE_ACCOUNT',
    'WONDER_PORTAL_LINEWORKS_PRIVATE_KEY',
    'WONDER_PORTAL_LINEWORKS_CHANNEL_IDS'
  ];
  const lineProperties = [
    'WONDER_PORTAL_LINE_CHANNEL_ACCESS_TOKEN',
    'WONDER_PORTAL_LINE_TO_IDS',
    'WONDER_PORTAL_LINE_GROUP_IDS'
  ];
  const mask = function(value) {
    const text = String(value || '').trim();
    if (!text) return { set: false, length: 0 };
    return {
      set: true,
      length: text.length,
      preview: text.slice(0, 4) + '...' + text.slice(-4)
    };
  };
  const result = {
    lineWorks: {},
    line: {},
    resolvedLineWorksBotId: getPortalLineWorksBotId_(),
    resolvedLineWorksChannelIds: getPortalLineWorksNoticeChannelIds_(),
    resolvedLineToIds: getPortalLineNoticeToIds_()
  };
  lineWorksProperties.forEach(function(name) {
    result.lineWorks[name] = mask(props.getProperty(name));
  });
  lineProperties.forEach(function(name) {
    result.line[name] = mask(props.getProperty(name));
  });
  Logger.log(JSON.stringify(result));
  return result;
}

function getPortalBoardSpreadsheetCandidates_() {
  const props = PropertiesService.getScriptProperties();
  const candidates = [
    props.getProperty('BOARD_SPREADSHEET_ID'),
    props.getProperty('SCHEDULE_SPREADSHEET_ID'),
    props.getProperty('SPREADSHEET_ID'),
    props.getProperty('PORTAL_SPREADSHEET_ID')
  ].filter(Boolean);

  const spreadsheets = [];
  for (let i = 0; i < candidates.length; i += 1) {
    try {
      spreadsheets.push(SpreadsheetApp.openById(String(candidates[i]).trim()));
    } catch (error) {}
  }

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) spreadsheets.push(active);
  return spreadsheets;
}

function getPortalBoardSheet_() {
  const names = ['重要事項共有専用', '重要共有事項', 'ImportantNotices', 'BoardPosts'];
  const spreadsheets = getPortalBoardSpreadsheetCandidates_();
  for (let i = 0; i < spreadsheets.length; i += 1) {
    for (let j = 0; j < names.length; j += 1) {
      const sheet = spreadsheets[i].getSheetByName(names[j]);
      if (sheet) {
        ensurePortalBoardHeader_(sheet);
        return sheet;
      }
    }
  }
  for (let i = 0; i < spreadsheets.length; i += 1) {
    const sheets = spreadsheets[i].getSheets();
    for (let j = 0; j < sheets.length; j += 1) {
      const sheet = sheets[j];
      const name = sheet.getName();
      if (name.indexOf('重要') >= 0 || name.toLowerCase().indexOf('notice') >= 0 || name.toLowerCase().indexOf('board') >= 0) {
        ensurePortalBoardHeader_(sheet);
        return sheet;
      }
    }
  }
  if (!spreadsheets.length) throw new Error('Portal spreadsheet was not found. Set BOARD_SPREADSHEET_ID or SCHEDULE_SPREADSHEET_ID.');
  const sheet = spreadsheets[0].insertSheet(names[0]);
  ensurePortalBoardHeader_(sheet);
  return sheet;
}

function getPortalBoardSpreadsheet_() {
  const spreadsheets = getPortalBoardSpreadsheetCandidates_();
  if (spreadsheets.length) return spreadsheets[0];
  throw new Error('Portal spreadsheet was not found. Set BOARD_SPREADSHEET_ID or SCHEDULE_SPREADSHEET_ID.');
}

function ensurePortalBoardHeader_(sheet) {
  const header = ['投稿日時', '投稿者', '重要事項共有専用', 'ID', '更新日時'];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    return;
  }
  const existing = sheet.getRange(1, 1, 1, header.length).getValues()[0];
  if (existing.join('') === '') {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
  } else {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
  }
}

function formatPortalBoardDate_(value) {
  const date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return String(value || '');
  return Utilities.formatDate(date, 'Asia/Tokyo', 'M/d HH:mm');
}
