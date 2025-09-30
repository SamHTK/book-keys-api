// Helper to read per-slug booking page configuration from env PAGES_JSON
// Example PAGES_JSON value:
// {
//   "exec-a": {
//     "schedulerUpn": "exec-a@example.com",
//     "calendars": ["exec-a@example.com", "room-1@example.com"],
//     "timeZone": "Eastern Standard Time",
//     "businessHours": { "start": "09:00", "end": "17:00" }
//   }
// }

function parsePagesJson() {
  const raw = process.env.PAGES_JSON || "";
  if (!raw.trim()) return {};
  try {
    const obj = JSON.parse(raw);
    if (obj && typeof obj === "object") return obj;
    return {};
  } catch (e) {
    throw new Error("Invalid PAGES_JSON: " + e.message);
  }
}

function validatePageConfig(slug, cfg) {
  if (!cfg || typeof cfg !== "object") {
    throw new Error(`Missing config for slug '${slug}'`);
  }
  const schedulerUpn = String(cfg.schedulerUpn || "").trim().toLowerCase();
  const calendars = Array.isArray(cfg.calendars) ? cfg.calendars.map(s => String(s).trim().toLowerCase()).filter(Boolean) : [];
  const timeZone = String(cfg.timeZone || "").trim() || "Eastern Standard Time";
  const bh = cfg.businessHours || {};
  const businessHours = {
    start: String(bh.start || "09:00"),
    end: String(bh.end || "17:00")
  };
  if (!schedulerUpn) throw new Error(`schedulerUpn missing for slug '${slug}'`);
  if (!calendars.length) throw new Error(`calendars missing for slug '${slug}'`);
  return { slug, schedulerUpn, calendars, timeZone, businessHours };
}

function getPageConfig(slug) {
  const pages = parsePagesJson();
  const cfg = pages[slug];
  return validatePageConfig(slug, cfg);
}

function getAllPages() {
  const pages = parsePagesJson();
  const out = [];
  for (const [slug, cfg] of Object.entries(pages)) {
    try {
      out.push(validatePageConfig(slug, cfg));
    } catch (_) {
      // skip invalid entries
    }
  }
  return out;
}

module.exports = { getPageConfig, getAllPages };
