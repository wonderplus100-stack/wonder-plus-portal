/**
 * Final web API override for Wonder+ Portal.
 *
 * This adds a dedicated event guide endpoint so the portal can read the latest
 * Google Form response sheet directly instead of relying on a stale schedule
 * payload or browser cache.
 */

var doGet = function(e) {
  const params = (e && e.parameter) || {};
  try {
    if (params.action === 'approveUser') {
      return approvePortalUserFromEmail_(params);
    }
    if (params.action === 'approvePasswordReset') {
      return approvePortalPasswordReset_(params);
    }

    let payload;
    switch (params.action) {
      case 'schedule':
        payload = buildPortalSchedulePayload_();
        payload.assignments = getPortalAssignmentMap_();
        break;
      case 'eventGuides':
        payload = typeof getPortalEventGuidesForApi_ === 'function'
          ? getPortalEventGuidesForApi_()
          : { ok: true, updatedAt: new Date().toISOString(), eventGuides: buildPortalEventGuidesSafe_() };
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
    return portalJsonOutput_({ ok: false, message: String(error && error.message ? error.message : error) }, params.callback);
  }
};
