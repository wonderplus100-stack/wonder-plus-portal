/**
 * Deprecated hotfix.
 *
 * Event guides are now joined from the authoritative schedule rows and the
 * latest guide form responses. The previous Alliance title rewrite made
 * unrelated guides collide, so this file intentionally leaves payloads intact.
 */

function buildPortalSchedulePayload_() {
  return buildPortalSchedulePayloadFromFormattedSheetFull_();
}

function normalizePortalAllianceGuideTitles_(payload) {
  return payload;
}

function normalizePortalAllianceGuides_(guides) {
  return guides;
}
