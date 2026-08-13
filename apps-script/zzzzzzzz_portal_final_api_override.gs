/**
 * True final API override for Wonder+ Portal.
 *
 * The portal had multiple historic doGet/buildPortalSchedulePayload_ overrides.
 * This file must be placed last in Apps Script so every public endpoint uses the
 * current formatted schedule sheet and the schedule-joined event guide payload.
 */

function ensurePortalFormattedScheduleReadyFinal_(options) {
  options = options || {};
  var result = {
    ok: false,
    refreshed: false,
    rows: 0,
    message: ''
  };

  function getSpreadsheet_() {
    var props = PropertiesService.getScriptProperties();
    var ids = [
      props.getProperty('SCHEDULE_SPREADSHEET_ID'),
      props.getProperty('WONDER_PORTAL_SCHEDULE_SPREADSHEET_ID'),
      '1SRb3_nwgPWEj2Kb44SDD38SI0HmkllE4G_P7DPYKGKk'
    ].filter(function(id, index, all) {
      return id && all.indexOf(id) === index;
    });

    for (var i = 0; i < ids.length; i++) {
      try {
        var ss = SpreadsheetApp.openById(ids[i]);
        if (getFormattedSheet_(ss)) return ss;
      } catch (idError) {
        // Try the next configured ID.
      }
    }

    if (typeof getPortalScheduleSpreadsheetFull_ === 'function') {
      var fallbackSs = getPortalScheduleSpreadsheetFull_();
      if (getFormattedSheet_(fallbackSs)) return fallbackSs;
    }

    throw new Error('formatted schedule sheet not found in configured spreadsheets.');
  }

  function getFormattedSheet_(ss) {
    return ss.getSheetByName('\u6574\u5f62\u6e08\u307f') || ss.getSheetByName('\u6574\u5f62');
  }

  if (!options.skipRefresh && typeof importLatestScheduleFormResponse === 'function') {
    try {
      importLatestScheduleFormResponse();
      result.refreshed = true;
    } catch (refreshError) {
      result.message = String(refreshError && refreshError.message ? refreshError.message : refreshError);
    }
  }

  try {
    var refreshedSs = getSpreadsheet_();
    var refreshedSheet = getFormattedSheet_(refreshedSs);
    if (refreshedSheet && refreshedSheet.getLastRow() > 1) {
      result.ok = true;
      result.rows = refreshedSheet.getLastRow() - 1;
    } else if (!result.message) {
      result.message = 'Formatted schedule sheet was not found after refresh.';
    }
  } catch (finalError) {
    result.message = String(finalError && finalError.message ? finalError.message : finalError);
  }

  return result;
}

function buildPortalSchedulePayloadFinal_() {
  var readiness = ensurePortalFormattedScheduleReadyFinal_();
  var payload = null;
  var selectedSource = '';
  var buildWarnings = [];
  var builders = [
    {
      name: 'formatted-schedule-sheet',
      available: typeof buildPortalSchedulePayloadFromFormattedSheetFull_ === 'function',
      run: function() { return buildPortalSchedulePayloadFromFormattedSheetFull_(); }
    },
    {
      name: 'dynamic-schedule-sheet',
      available: typeof buildPortalScheduleDataSafe_ === 'function',
      run: function() { return buildPortalScheduleDataSafe_(); }
    },
    {
      name: 'fast-cache-fallback',
      available: typeof buildPortalSchedulePayloadFastFinal_ === 'function',
      run: function() { return buildPortalSchedulePayloadFastFinal_(); }
    }
  ];

  for (var i = 0; i < builders.length; i++) {
    if (!builders[i].available) continue;
    try {
      var candidate = builders[i].run();
      if (!payload && candidate) {
        payload = candidate;
        selectedSource = builders[i].name;
      }
      if (countPortalFinalScheduleItems_(candidate && candidate.schedules) > 0) {
        payload = candidate;
        selectedSource = builders[i].name;
        break;
      }
    } catch (builderError) {
      buildWarnings.push(builders[i].name + ': ' + String(builderError && builderError.message ? builderError.message : builderError));
    }
  }

  if (!payload) {
    payload = {
      ok: false,
      updatedAt: new Date().toISOString(),
      schedules: {},
      events: []
    };
  }

  payload = payload || {};
  payload.ok = payload.ok !== false && readiness.ok;
  payload.source = selectedSource || 'formatted-schedule-sheet';
  payload.masterSource = 'formatted-schedule-sheet';
  payload.scheduleReadiness = readiness;
  if (buildWarnings.length) payload.buildWarnings = buildWarnings;

  // Event guides have their own endpoint. Keeping old guide rows in schedule
  // responses caused the portal to fall back to stale guide data.
  if (payload.eventGuides) delete payload.eventGuides;

  try {
    payload.assignments = typeof getPortalAssignmentMap_ === 'function' ? getPortalAssignmentMap_() : {};
  } catch (assignmentError) {
    payload.assignments = {};
    payload.assignmentError = String(assignmentError && assignmentError.message ? assignmentError.message : assignmentError);
  }

  return payload;
}

function countPortalFinalScheduleItems_(schedules) {
  var count = 0;
  if (!schedules || typeof schedules !== 'object') return 0;
  Object.keys(schedules).forEach(function(monthKey) {
    var byType = schedules[monthKey];
    if (!byType || typeof byType !== 'object') return;
    Object.keys(byType).forEach(function(typeKey) {
      var rows = byType[typeKey];
      if (Array.isArray(rows)) count += rows.length;
    });
  });
  return count;
}

function buildPortalEventGuidesPayloadFinal_() {
  var readiness = ensurePortalFormattedScheduleReadyFinal_();
  var eventGuides = [];
  var buildMessage = '';
  var refreshWarning = '';

  try {
    if (typeof refreshPortalEventGuideCache_ === 'function') {
      try {
        refreshPortalEventGuideCache_();
      } catch (refreshError) {
        refreshWarning = String(refreshError && refreshError.message ? refreshError.message : refreshError);
      }
    }
    if (typeof buildPortalEventGuideLibrary_ === 'function' && typeof buildPortalEventGuidesFromSchedule_ === 'function') {
      eventGuides = buildPortalEventGuidesFromSchedule_(buildPortalEventGuideLibrary_());
    } else if (typeof buildPortalEventGuidesSafe_ === 'function') {
      eventGuides = buildPortalEventGuidesSafe_();
    } else {
      buildMessage = 'event guide builder is not installed';
    }
  } catch (error) {
    buildMessage = String(error && error.message ? error.message : error);
    eventGuides = [];
  }

  eventGuides = sanitizePortalFinalEventGuides_(eventGuides);

  return {
    ok: readiness.ok && !buildMessage,
    updatedAt: new Date().toISOString(),
    source: 'schedule-joined-event-guides',
    scheduleReadiness: readiness,
    count: eventGuides.length,
    eventGuides: eventGuides,
    message: buildMessage,
    refreshWarning: refreshWarning
  };
}

function sanitizePortalFinalEventGuides_(guides) {
  if (!Array.isArray(guides)) return [];

  var seen = {};
  var cleaned = [];

  guides.forEach(function(guide) {
    if (!guide || typeof guide !== 'object') return;

    var month = Number(guide.month || guide.eventMonth);
    var day = Number(guide.day || guide.eventDay);
    if (!(month >= 1 && month <= 12 && day >= 1 && day <= 31)) return;

    var title = String(guide.title || guide.eventName || guide.name || '').trim();
    if (!title || /^(open|untitled|title)$/i.test(title)) return;

    var body = String(guide.body || guide.text || guide.content || guide.description || guide.guideText || guide.message || '');
    if (typeof stripPortalEventGuideMaterialFooter_ === 'function') {
      body = stripPortalEventGuideMaterialFooter_(body);
    }

    var timeRange = String(guide.timeRange || guide.startEndTime || guide.eventTime || '').trim();
    var key = [month, day, title.toLowerCase(), timeRange].join('|');
    var normalized = {};
    Object.keys(guide).forEach(function(prop) {
      normalized[prop] = guide[prop];
    });

    normalized.month = month;
    normalized.day = day;
    normalized.eventMonth = month;
    normalized.eventDay = day;
    normalized.category = 'wonder';
    normalized.title = title;
    normalized.eventName = title;
    normalized.name = title;
    normalized.timeRange = timeRange;
    normalized.startEndTime = timeRange;
    normalized.eventTime = timeRange;
    normalized.body = body;
    normalized.text = body;
    normalized.source = 'schedule-joined-event-guides';

    if (!seen[key]) {
      seen[key] = normalized;
      cleaned.push(normalized);
      return;
    }

    var current = seen[key];
    var currentRank = current.guideMatched === false ? 0 : 1;
    var nextRank = normalized.guideMatched === false ? 0 : 1;
    if (
      nextRank > currentRank ||
      (nextRank === currentRank && getPortalFinalGuideUpdatedTime_(normalized) >= getPortalFinalGuideUpdatedTime_(current))
    ) {
      seen[key] = normalized;
      var index = cleaned.indexOf(current);
      if (index >= 0) cleaned[index] = normalized;
    }
  });

  cleaned.sort(function(a, b) {
    return (a.month - b.month) ||
      (a.day - b.day) ||
      String(a.timeRange || '').localeCompare(String(b.timeRange || '')) ||
      String(a.title || '').localeCompare(String(b.title || ''), 'ja');
  });

  return cleaned;
}

function getPortalFinalGuideUpdatedTime_(guide) {
  if (!guide) return 0;
  if (typeof getPortalGuideUpdatedTime_ === 'function') {
    try {
      return Number(getPortalGuideUpdatedTime_(guide)) || 0;
    } catch (error) {
      return 0;
    }
  }
  return Date.parse(guide.updatedAt || guide.createdAt || guide.timestamp || '') || 0;
}

function refreshPortalDataFinal_() {
  var schedule = { ok: false };
  var guides = { ok: false };
  if (typeof importLatestScheduleFormResponse === 'function') {
    try {
      importLatestScheduleFormResponse();
      schedule = { ok: true };
    } catch (scheduleError) {
      schedule = { ok: false, message: String(scheduleError && scheduleError.message ? scheduleError.message : scheduleError) };
    }
  }
  if (typeof refreshPortalEventGuideCache_ === 'function') {
    try {
      var refreshResult = refreshPortalEventGuideCache_();
      guides = buildPortalEventGuidesPayloadFinal_();
      guides.refresh = refreshResult;
    } catch (guideError) {
      guides = { ok: false, message: String(guideError && guideError.message ? guideError.message : guideError) };
    }
  } else {
    guides = buildPortalEventGuidesPayloadFinal_();
  }
  return {
    ok: schedule.ok && guides.ok !== false,
    updatedAt: new Date().toISOString(),
    schedule: schedule,
    eventGuides: guides
  };
}

function getPortalBoardSpreadsheetCandidates_() {
  var props = PropertiesService.getScriptProperties();
  var candidates = [
    props.getProperty('BOARD_SPREADSHEET_ID'),
    props.getProperty('WONDER_PORTAL_BOARD_SPREADSHEET_ID'),
    props.getProperty('SCHEDULE_SPREADSHEET_ID'),
    props.getProperty('WONDER_PORTAL_SCHEDULE_SPREADSHEET_ID'),
    props.getProperty('SPREADSHEET_ID'),
    props.getProperty('PORTAL_SPREADSHEET_ID'),
    '1SRb3_nwgPWEj2Kb44SDD38SI0HmkllE4G_P7DPYKGKk'
  ].filter(function(id, index, all) {
    return id && all.indexOf(id) === index;
  });

  var spreadsheets = [];
  for (var i = 0; i < candidates.length; i += 1) {
    try {
      spreadsheets.push(SpreadsheetApp.openById(String(candidates[i]).trim()));
    } catch (error) {}
  }

  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) spreadsheets.push(active);
  } catch (activeError) {}

  return spreadsheets;
}

function getPortalBoardSheet_() {
  var names = ['重要事項共有専用', '重要共有事項', 'ImportantNotices', 'BoardPosts'];
  var spreadsheets = getPortalBoardSpreadsheetCandidates_();
  for (var i = 0; i < spreadsheets.length; i += 1) {
    for (var j = 0; j < names.length; j += 1) {
      var sheet = spreadsheets[i].getSheetByName(names[j]);
      if (sheet) {
        ensurePortalBoardHeader_(sheet);
        return sheet;
      }
    }
  }

  for (var s = 0; s < spreadsheets.length; s += 1) {
    var sheets = spreadsheets[s].getSheets();
    for (var k = 0; k < sheets.length; k += 1) {
      var candidateSheet = sheets[k];
      var name = String(candidateSheet.getName() || '');
      var lower = name.toLowerCase();
      if (name.indexOf('重要') >= 0 || lower.indexOf('notice') >= 0 || lower.indexOf('board') >= 0) {
        ensurePortalBoardHeader_(candidateSheet);
        return candidateSheet;
      }
    }
  }

  if (!spreadsheets.length) throw new Error('Portal spreadsheet was not found. Set BOARD_SPREADSHEET_ID or SCHEDULE_SPREADSHEET_ID.');
  var newSheet = spreadsheets[0].insertSheet(names[0]);
  ensurePortalBoardHeader_(newSheet);
  return newSheet;
}

function getPortalBoardSpreadsheet_() {
  var spreadsheets = getPortalBoardSpreadsheetCandidates_();
  if (spreadsheets.length) return spreadsheets[0];
  throw new Error('Portal spreadsheet was not found. Set BOARD_SPREADSHEET_ID or SCHEDULE_SPREADSHEET_ID.');
}

var doGet = function(e) {
  var params = (e && e.parameter) || {};
  try {
    if (params.action === 'approveUser') {
      return approvePortalUserFromEmail_(params);
    }
    if (params.action === 'approvePasswordReset') {
      return approvePortalPasswordReset_(params);
    }

    var payload;
    switch (params.action) {
      case 'schedule':
        payload = buildPortalSchedulePayloadFinal_();
        break;
      case 'eventGuides':
        payload = buildPortalEventGuidesPayloadFinal_();
        break;
      case 'refreshPortalData':
        payload = refreshPortalDataFinal_();
        break;
      case 'eventGuideDiagnostics':
        payload = typeof getPortalEventGuideDiagnostics_ === 'function'
          ? getPortalEventGuideDiagnostics_()
          : { ok: false, message: 'event guide diagnostics is not installed' };
        break;
      case 'register':
        payload = registerPortalUser_(params);
        break;
      case 'requestPasswordReset':
      case 'resetPassword':
        payload = requestPortalPasswordReset_(params);
        break;
      case 'login':
        payload = loginPortalUser_(params);
        break;
      case 'me':
        payload = getPortalSession_(params.token);
        break;
      case 'assignments':
        payload = {
          ok: true,
          assignments: getPortalAssignmentMap_(),
          mine: params.token ? getPortalUserAssignmentIds_(validatePortalToken_(params.token).email) : []
        };
        break;
      case 'boardPosts':
        payload = getPortalBoardPosts_();
        break;
      case 'createBoardPost':
        payload = createPortalBoardPost_(params);
        break;
      case 'minutes':
        payload = getPortalMinutes_();
        break;
      case 'toggleAssignment':
        payload = togglePortalAssignment_(params);
        break;
      case 'adminUsers':
        payload = listPortalUsers_(params.token);
        break;
      case 'adminSetUserStatus':
        payload = setPortalUserStatus_(params);
        break;
      default:
        payload = { ok: true, name: 'WonderPlus Portal Automation' };
    }
    return portalJsonOutput_(payload, params.callback);
  } catch (error) {
    return portalJsonOutput_({
      ok: false,
      source: params.action === 'eventGuides' ? 'schedule-joined-event-guides' : 'portal-final-api',
      updatedAt: new Date().toISOString(),
      message: String(error && error.message ? error.message : error)
    }, params.callback);
  }
};
