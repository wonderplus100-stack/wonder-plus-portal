import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const SCHEDULE_SPREADSHEET_ID = "1SRb3_nwgPWEj2Kb44SDD38SI0HmkllE4G_P7DPYKGKk";
const GUIDE_SPREADSHEET_ID = "1rXae1o13ucNAI6VfPhXKIFi3LTq3uV0g4wKm_zI4zKM";
const SCHEDULE_SHEET = "整形済み";
const GUIDE_SHEET = "イベント案内";
const SOURCE = "schedule-joined-event-guides";

const VENUES = [
  ["北九州", "小倉", "kitakyushu", "kitakyusyu", "kokura"],
  ["福岡", "博多", "fukuoka", "hakata"],
  ["熊本", "kumamoto"],
  ["銀座", "ginza"],
  ["新宿", "shinjuku"],
  ["名古屋", "nagoya"],
  ["大阪", "osaka"],
  ["岡山", "okayama"],
  ["広島", "hiroshima"],
  ["京都", "kyoto"],
  ["神戸", "kobe"],
  ["横浜", "yokohama"],
  ["船橋", "funabashi"],
  ["千葉", "chiba"],
  ["仙台", "sendai"],
  ["郡山", "koriyama"],
  ["福島", "fukushima", "fukushim"],
  ["金沢", "kanazawa"],
  ["新潟", "niigata"],
  ["浜松", "hamamatsu"],
  ["久留米", "kurume"],
  ["前橋", "maebashi"],
  ["高崎", "takasaki", "labi"],
  ["静岡", "shizuoka"],
  ["札幌", "sapporo", "saporo"],
  ["宇都宮", "utsunomiya", "utunomiya"],
  ["町田", "machida"],
  ["甲府", "kofu"],
  ["大宮", "omiya"]
];

const BRANDS = [
  ["Real Estate", "realestate", "real estate", "不動産"],
  ["Entertainment", "entertainment"],
  ["Executive", "executive"],
  ["Alliance", "alliance"],
  ["Finance", "finance"],
  ["Leaders", "leaders"],
  ["Ladies", "ladies", "lady"],
  ["Gravity", "gravity"],
  ["Beauty", "beauty"],
  ["Story", "story"],
  ["Night", "night"],
  ["CXO", "cxo"],
  ["100", "100人", "+100"]
];

function csvUrl(id, sheet) {
  return `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheet)}`;
}

async function fetchText(url) {
  const res = await fetch(url, { cache: "no-store" });
  if (!res.ok) throw new Error(`HTTP ${res.status}: ${url}`);
  return res.text();
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;
  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];
    if (quoted) {
      if (char === "\"" && next === "\"") {
        field += "\"";
        i += 1;
      } else if (char === "\"") {
        quoted = false;
      } else {
        field += char;
      }
    } else if (char === "\"") {
      quoted = true;
    } else if (char === ",") {
      row.push(field);
      field = "";
    } else if (char === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else if (char !== "\r") {
      field += char;
    }
  }
  row.push(field);
  rows.push(row);
  return rows.filter((cells) => cells.some((cell) => String(cell || "").trim()));
}

function tableFromCsv(text) {
  const rows = parseCsv(text.replace(/^\uFEFF/, ""));
  const headers = rows.shift() || [];
  return rows.map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] || ""])));
}

function normalizeText(value) {
  return String(value || "")
    .normalize("NFKC")
    .toLowerCase()
    .replace(/wonder\s*[＋+]/g, "wonder+")
    .replace(/w\s*[＋+]/g, "wonder+")
    .replace(/＆/g, "&")
    .replace(/&/g, "and")
    .replace(/[ーｰ−–—]/g, "-")
    .replace(/\s+/g, "")
    .replace(/[()（）「」『』【】\[\]、，,。]/g, "");
}

function findVenueKeys(...values) {
  const text = normalizeText(values.join(" "));
  return VENUES
    .filter((aliases) => aliases.some((alias) => text.includes(normalizeText(alias))))
    .map((aliases) => aliases[0]);
}

function getVenueKey(...values) {
  const matches = findVenueKeys(...values);
  if (matches.includes("高崎")) return "高崎";
  if (matches.includes("前橋")) return "前橋";
  return matches[0] || "";
}

function getBrandKey(...values) {
  const text = normalizeText(values.join(" "));
  if (text.includes("realestate") || text.includes("不動産")) return "Real Estate";
  const brand = BRANDS.find((aliases) => aliases.some((alias) => text.includes(normalizeText(alias))));
  return brand ? brand[0] : "";
}

function getGuideVenueKey(guide) {
  return getVenueKey(guide.title) || getVenueKey(guide.materialUrl) || getVenueKey(guide.body);
}

function getGuideBrandKey(guide) {
  return getBrandKey(guide.title) || getBrandKey(guide.body);
}

function createEmptySchedules() {
  const schedules = {};
  for (let month = 1; month <= 12; month += 1) {
    schedules[`month${month}Wonder`] = { month, title: `${month}月 Wonder+`, items: [] };
    schedules[`month${month}Regular`] = { month, title: `${month}月 レギュラー`, items: [] };
    schedules[`month${month}Miraiba`] = { month, title: `${month}月 ミライバ`, items: [] };
    schedules[`month${month}Meeting`] = { month, title: `${month}月 ミーティング`, items: [] };
  }
  return schedules;
}

function scheduleCategory(eventName) {
  const text = normalizeText(eventName);
  if (text.includes("meeting") || eventName.includes("ミーティング")) return "Meeting";
  if (eventName.includes("ミライバ")) return "Miraiba";
  if (text.includes("wonder+")) return "Wonder";
  return "Regular";
}

function titleForSchedule(row) {
  const venue = String(row["会場"] || "").trim();
  const eventName = String(row["イベント名"] || "").trim();
  if (venue && eventName) return `${venue} ${eventName}`.replace(/\s+/g, " ").trim();
  return eventName || venue || "イベント";
}

function buildSchedulePayload(rows, previousAssignments) {
  const schedules = createEmptySchedules();
  for (const row of rows) {
    const month = Number(row["月"]);
    const day = Number(row["日"]);
    if (!Number.isInteger(month) || month < 1 || month > 12 || !Number.isInteger(day) || day < 1 || day > 31) continue;
    const category = scheduleCategory(row["イベント名"] || "");
    const key = `month${month}${category}`;
    const time = String(row["時間"] || "").trim();
    const memo = String(row["メモ"] || "").trim();
    const meta = [time, memo].filter(Boolean).join(" / ");
    schedules[key].items.push({
      day,
      title: titleForSchedule(row),
      meta,
      venue: String(row["会場"] || "").trim(),
      eventName: String(row["イベント名"] || "").trim(),
      time,
      capacity: String(row["人数"] || "").trim(),
      sourceKey: String(row["更新キー"] || row["原文"] || "").trim()
    });
  }
  for (const group of Object.values(schedules)) {
    group.items.sort((a, b) => (a.day - b.day) || String(a.time).localeCompare(String(b.time), "ja") || a.title.localeCompare(b.title, "ja"));
  }
  return {
    updatedAt: new Date().toISOString(),
    schedules,
    ok: true,
    source: "formatted-schedule-sheet",
    scheduleReadiness: { ok: true, source: "local-cache-builder", missing: [] },
    assignments: previousAssignments || {}
  };
}

function parseMaybeDate(value) {
  const text = String(value || "").trim();
  if (!text) return 0;
  const date = new Date(text);
  if (Number.isFinite(date.getTime())) return date.getTime();
  const match = text.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (!match) return 0;
  return new Date(Number(match[3]), Number(match[1]) - 1, Number(match[2]), Number(match[4] || 0), Number(match[5] || 0)).getTime();
}

function latestGuideRows(rows) {
  return rows
    .map((row, index) => ({
      row,
      index,
      title: String(row["イベント名"] || "").trim(),
      month: Number(String(row["月"] || "").replace(/[^\d]/g, "")) || null,
      body: String(row["案内文"] || "").trim(),
      materialUrl: String(row["素材URL"] || "").trim(),
      attachmentUrl: String(row["添付ファイルURL"] || "").trim(),
      category: String(row["カテゴリ"] || "").trim(),
      timestamp: Math.max(parseMaybeDate(row["更新日時"]), parseMaybeDate(row["登録日時"]), index)
    }))
    .filter((guide) => guide.title || guide.body || guide.materialUrl || guide.attachmentUrl)
    .sort((a, b) => b.timestamp - a.timestamp || b.index - a.index);
}

function scoreGuide(schedule, guide) {
  const scheduleVenue = getVenueKey(schedule.title, schedule.venue);
  const guideVenue = getGuideVenueKey(guide);
  const scheduleBrand = getBrandKey(schedule.title, schedule.eventName);
  const guideBrand = getGuideBrandKey(guide);
  let score = 0;

  if (scheduleBrand && guideBrand && scheduleBrand !== guideBrand) return -1000;
  if (scheduleVenue && guideVenue && scheduleVenue !== guideVenue) return -1000;

  if (scheduleVenue && guideVenue === scheduleVenue) score += 130;
  if (scheduleVenue && !guideVenue) score -= 30;
  if (scheduleBrand && guideBrand === scheduleBrand) score += 90;
  if (scheduleBrand && !guideBrand && scheduleBrand !== "100") score -= 25;
  if (!scheduleBrand && !guideBrand) score += 5;

  if (guide.month && Number(guide.month) === Number(schedule.month)) score += 25;
  if (guide.month && Number(guide.month) !== Number(schedule.month)) score -= 30;

  const scheduleText = normalizeText(`${schedule.title} ${schedule.eventName}`);
  const guideText = normalizeText(`${guide.title} ${guide.body}`);
  if (guideText.includes(scheduleText) || scheduleText.includes(normalizeText(guide.title))) score += 10;
  return score;
}

function getBestGuide(schedule, guides) {
  let best = null;
  for (const guide of guides) {
    const score = scoreGuide(schedule, guide);
    if (score < 70) continue;
    if (!best || score > best.score || (score === best.score && guide.timestamp > best.guide.timestamp)) {
      best = { guide, score };
    }
  }
  return best?.guide || null;
}

function stripMaterialFooter(body) {
  return String(body || "")
    .replace(/\r\n?/g, "\n")
    .replace(/\n?【素材】\s\S*$/u, "")
    .replace(/\n?https?:\/\/drive\.google\.com\/[^\s]+[\s\S]*$/u, "")
    .trim();
}

function splitTime(meta) {
  const match = String(meta || "").match(/(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})/);
  if (!match) return null;
  return {
    start: `${match[1].padStart(2, "0")}:${match[2]}`,
    end: `${match[3].padStart(2, "0")}:${match[4]}`
  };
}

function minusMinutes(time, minutes) {
  const [hour, minute] = time.split(":").map(Number);
  const date = new Date(2000, 0, 1, hour, minute - minutes);
  return `${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}`;
}

function replaceEventTimes(body, schedule) {
  const time = splitTime(schedule.meta || schedule.time);
  if (!time) return body;
  const open = minusMinutes(time.start, 15);
  const line = `${open} オープン｜${time.start} スタート｜${time.end} クローズ`;
  const normalized = String(body || "");
  if (/オープン.*スタート.*クローズ/u.test(normalized)) {
    return normalized.replace(/\d{1,2}:\d{2}\s*オープン\s*[|｜]\s*\d{1,2}:\d{2}\s*スタート\s*[|｜]\s*\d{1,2}:\d{2}\s*クローズ/u, line);
  }
  return normalized.replace(/(■\s*日時\s*\n?)/u, `$1${line}\n`);
}

function buildGuideBody(schedule, guide) {
  const timeText = schedule.meta || schedule.time || "";
  const body = guide ? replaceEventTimes(stripMaterialFooter(guide.body), schedule) : "";
  const fallback = body || "このイベントに紐づくDrive案内文章はまだ読み込まれていません。イベント案内共有フォームから案内本文を投稿してください。";
  return [
    `【イベント名】${guide?.title || schedule.title}`,
    `【開催日時】${schedule.month}月${schedule.day}日 ${timeText}`,
    "【案内文】",
    fallback
  ].join("\n");
}

function attachmentFiles(guide) {
  if (!guide) return [];
  return [guide.materialUrl, guide.attachmentUrl]
    .flatMap((value) => String(value || "").split(/\s*,\s*|\n+/))
    .map((url) => url.trim())
    .filter(Boolean)
    .map((url, index) => ({ name: `投稿素材${index + 1}`, url, downloadUrl: url }));
}

function eventId(schedule) {
  return `month${schedule.month}-wonder-${schedule.day}-${normalizeText(schedule.title)}-${normalizeText(schedule.meta)}`;
}

function buildEventGuides(schedulePayload, guideRows) {
  const guides = latestGuideRows(guideRows);
  const entries = Object.entries(schedulePayload.schedules)
    .filter(([key]) => /Wonder$/.test(key))
    .flatMap(([, group]) => group.items.map((item) => ({ ...item, month: group.month })));

  return entries.map((schedule) => {
    const guide = getBestGuide(schedule, guides);
    const body = buildGuideBody(schedule, guide);
    return {
      id: eventId(schedule),
      source: SOURCE,
      month: Number(schedule.month),
      day: Number(schedule.day),
      title: schedule.title,
      eventName: guide?.title || schedule.title,
      scheduleTitle: schedule.title,
      scheduleMeta: schedule.meta,
      venue: schedule.venue || getVenueKey(schedule.title),
      brand: getBrandKey(schedule.title, schedule.eventName) || "Wonder+",
      body,
      text: body,
      materialUrl: guide?.materialUrl || "",
      attachments: attachmentFiles(guide),
      files: attachmentFiles(guide),
      guideMatched: Boolean(guide),
      matchedGuideTitle: guide?.title || "",
      matchedSource: guide ? GUIDE_SHEET : "",
      guideUpdatedAt: guide ? new Date(guide.timestamp).toISOString() : "",
      updatedAt: new Date().toISOString()
    };
  });
}

function readJsonIfExists(file) {
  try {
    return JSON.parse(fs.readFileSync(path.join(ROOT, file), "utf8").replace(/^\uFEFF/, ""));
  } catch {
    return null;
  }
}

function writeJson(file, value) {
  fs.writeFileSync(path.join(ROOT, file), `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

const [scheduleCsv, guideCsv] = await Promise.all([
  fetchText(csvUrl(SCHEDULE_SPREADSHEET_ID, SCHEDULE_SHEET)),
  fetchText(csvUrl(GUIDE_SPREADSHEET_ID, GUIDE_SHEET))
]);

const previousSchedule = readJsonIfExists("schedule-cache.json");
const schedulePayload = buildSchedulePayload(tableFromCsv(scheduleCsv), previousSchedule?.assignments);
const eventGuides = buildEventGuides(schedulePayload, tableFromCsv(guideCsv));
const eventGuidePayload = {
  ok: true,
  updatedAt: new Date().toISOString(),
  source: SOURCE,
  sourceSpreadsheetId: GUIDE_SPREADSHEET_ID,
  sourceSheet: GUIDE_SHEET,
  count: eventGuides.length,
  eventGuides,
  scheduleReadiness: { ok: true, source: "local-cache-builder", missing: [] }
};

writeJson("schedule-cache.json", schedulePayload);
writeJson("event-guide-cache.json", eventGuidePayload);

const unmatched = eventGuides.filter((guide) => !guide.guideMatched);
console.log(`schedule groups: ${Object.keys(schedulePayload.schedules).length}`);
console.log(`event guides: ${eventGuides.length}, unmatched: ${unmatched.length}`);
console.log(unmatched.slice(0, 20).map((guide) => `${guide.month}/${guide.day} ${guide.title}`).join("\n"));
