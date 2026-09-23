// Refreshes schedule-cache.json / event-guide-cache.json (the static
// fallback the site paints from instantly, and falls back to whenever the
// live Apps Script call is slow, rate-limited, or briefly down) straight
// from the live portal API. Run on a schedule by
// .github/workflows/refresh-portal-cache.yml - unlike the older
// build-portal-caches.mjs, this doesn't need public Sheets CSV export
// access, and it can never drift from the real backend's matching logic
// since it's just saving that backend's own response.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const SCHEDULE_API_URL =
  "https://script.google.com/macros/s/AKfycbyYU-Hz_2taLQ3ke3KRGWOuP-szI8TwuZwV-tpXeeZImMNrV-S4O4lfFNX1s3b4LM3jZg/exec?action=schedule";
const REQUIRED_SCHEDULE_SOURCE = "formatted-schedule-sheet";
const REQUIRED_EVENT_GUIDE_SOURCE = "schedule-joined-event-guides";
const MAX_ATTEMPTS = 5;
const RETRY_DELAY_MS = 20000;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function fetchSchedulePayload() {
  let lastError;
  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt += 1) {
    try {
      const res = await fetch(`${SCHEDULE_API_URL}&_=${Date.now()}`, { redirect: "follow" });
      const text = await res.text();
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      let payload;
      try {
        payload = JSON.parse(text);
      } catch {
        // Apps Script returns an HTML error page (Drive/quota hiccup, not
        // valid JSON) under transient rate limiting - retry rather than
        // treat it as a real, permanent failure.
        throw new Error("non-JSON response (likely a transient Apps Script/Drive hiccup)");
      }
      return payload;
    } catch (error) {
      lastError = error;
      console.warn(`Attempt ${attempt}/${MAX_ATTEMPTS} failed: ${error.message}`);
      if (attempt < MAX_ATTEMPTS) await sleep(RETRY_DELAY_MS);
    }
  }
  throw lastError;
}

function isValidSchedulePayload(payload) {
  return Boolean(
    payload &&
      typeof payload === "object" &&
      payload.ok !== false &&
      payload.schedules &&
      payload.source === REQUIRED_SCHEDULE_SOURCE
  );
}

function isValidEventGuide(guide) {
  if (!guide || typeof guide !== "object") return false;
  const month = Number(guide.month ?? guide.eventMonth);
  const day = Number(guide.day ?? guide.eventDay);
  const title = String(guide.title || guide.eventName || guide.name || "").trim();
  return (
    guide.source === REQUIRED_EVENT_GUIDE_SOURCE &&
    Number.isInteger(month) &&
    month >= 1 &&
    month <= 12 &&
    Number.isInteger(day) &&
    day >= 1 &&
    day <= 31 &&
    Boolean(title)
  );
}

function writeJson(file, value) {
  fs.writeFileSync(path.join(ROOT, file), `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

const payload = await fetchSchedulePayload();

if (!isValidSchedulePayload(payload)) {
  console.error("Fetched schedule payload failed validation; leaving existing cache files untouched.");
  console.error(JSON.stringify({ ok: payload?.ok, source: payload?.source, message: payload?.message }, null, 2));
  process.exit(1);
}

const eventGuides = Array.isArray(payload.eventGuides) ? payload.eventGuides : [];
const badGuides = eventGuides.filter((guide) => !isValidEventGuide(guide));
if (badGuides.length) {
  console.error(`${badGuides.length} of ${eventGuides.length} event guides failed validation; leaving existing cache files untouched.`);
  console.error(JSON.stringify(badGuides.slice(0, 5), null, 2));
  process.exit(1);
}

const schedulePayload = {
  ok: payload.ok,
  updatedAt: payload.updatedAt,
  source: payload.source,
  masterSource: payload.masterSource,
  scheduleCounts: payload.scheduleCounts,
  assignments: payload.assignments,
  schedules: payload.schedules
};
writeJson("schedule-cache.json", schedulePayload);

const eventGuidePayload = {
  ok: true,
  updatedAt: payload.updatedAt,
  source: REQUIRED_EVENT_GUIDE_SOURCE,
  count: eventGuides.length,
  eventGuides,
  scheduleReadiness: { ok: true, source: "live-api-snapshot", missing: [] }
};
writeJson("event-guide-cache.json", eventGuidePayload);

console.log(`Wrote schedule-cache.json and event-guide-cache.json (updatedAt: ${payload.updatedAt}, ${eventGuides.length} event guides).`);
