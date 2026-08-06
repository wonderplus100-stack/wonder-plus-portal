/**
 * Ensures Wonder+Alliance guide text matches the Ginza schedule title.
 * This override runs after the full schedule builder and only adjusts guide
 * titles in the API response; schedule rows and guide bodies are untouched.
 */

function buildPortalSchedulePayload_() {
  const payload = buildPortalSchedulePayloadFromFormattedSheetFull_();
  return normalizePortalAllianceGuideTitles_(payload);
}

function normalizePortalAllianceGuideTitles_(payload) {
  if (!payload || typeof payload !== 'object') return payload;
  payload.eventGuides = normalizePortalAllianceGuides_(payload.eventGuides);
  payload.eventGuideDocuments = normalizePortalAllianceGuides_(payload.eventGuideDocuments);
  payload.guides = normalizePortalAllianceGuides_(payload.guides);
  return payload;
}

function normalizePortalAllianceGuides_(guides) {
  if (!Array.isArray(guides)) return guides;
  return guides.map(function(guide) {
    if (!guide || typeof guide !== 'object') return guide;
    const titleText = [
      guide.title,
      guide.eventName,
      guide.name,
      guide.fileName,
      guide.body,
      guide.text,
      guide.content
    ].map(function(value) {
      return String(value || '').normalize('NFKC').replace(/[＋+]/g, '+').toLowerCase();
    }).join(' ');
    if (titleText.indexOf('wonder+alliance') < 0) return guide;
    const copy = {};
    Object.keys(guide).forEach(function(key) { copy[key] = guide[key]; });
    copy.title = '\u9280\u5ea7 Wonder\uff0bAlliance';
    copy.eventName = '\u9280\u5ea7 Wonder\uff0bAlliance';
    return copy;
  });
}
