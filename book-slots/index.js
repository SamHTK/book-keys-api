const { ClientCertificateCredential } = require("@azure/identity");
const { getPageConfig } = require("../lib/pages");
const { DateTime } = require("luxon");

async function getToken() {
  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const pem = process.env.CERT_PEM;
  if (!tenantId || !clientId || !pem) throw new Error("Missing TENANT_ID, CLIENT_ID, or CERT_PEM");
  const cred = new ClientCertificateCredential(tenantId, clientId, { certificate: pem });
  const token = await cred.getToken("https://graph.microsoft.com/.default");
  return token.token;
}

function clampToBusinessHours(dateIso, { start, end }, timeZone) {
  // dateIso is YYYY-MM-DD in local browser, we construct start/end in tz
  // We'll use Luxon to build the local start/end in the correct timezone
  const d = DateTime.fromISO(dateIso, { zone: timeZone });
  const startLocal = d.set({ hour: Number(start.split(':')[0]), minute: Number(start.split(':')[1]), second: 0, millisecond: 0 });
  const endLocal = d.set({ hour: Number(end.split(':')[0]), minute: Number(end.split(':')[1]), second: 0, millisecond: 0 });
  return { startLocal, endLocal };
}

function intersectBusy(views) {
  // views: array of strings like "0.." Busy=1, Free=0 etc. availabilityView uses chars per 15-min block
  if (!views.length) return "";
  let acc = views[0];
  for (let i = 1; i < views.length; i++) {
    const b = views[i];
    let out = '';
    for (let j = 0; j < Math.min(acc.length, b.length); j++) {
      // consider '0' free, other (1-3) as busy
      const aC = acc[j];
      const bC = b[j];
      out += (aC !== '0' || bC !== '0') ? '1' : '0';
    }
    acc = out;
  }
  return acc;
}

function slotsFromAvailability(availabilityView, windowStart, intervalMinutes, durationMinutes, timeZone) {
  // availabilityView: string per 15-min block from windowStart
  // Build slots aligned to intervalMinutes (e.g., 15/30) where durationMinutes fits with all '0'
  const blockMinutes = 15;
  const blocks = availabilityView.split('').map(c => c === '0'); // true=free
  const start = DateTime.fromISO(windowStart, { zone: 'utc' });

  const step = intervalMinutes; // step between slot starts
  const out = [];

  for (let minute = 0; minute <= (blocks.length * blockMinutes - durationMinutes); minute += step) {
    const startDt = start.plus({ minutes: minute });
    const neededBlocks = Math.ceil(durationMinutes / blockMinutes);
    const startBlock = Math.floor(minute / blockMinutes);
    let ok = true;
    for (let k = 0; k < neededBlocks; k++) {
      if (!blocks[startBlock + k]) { ok = false; break; }
    }
    if (ok) {
      const endDt = startDt.plus({ minutes: durationMinutes });
      out.push({ start: startDt.toUTC().toISO(), end: endDt.toUTC().toISO() });
    }
  }
  return out;
}

module.exports = async function (context, req) {
  try {
    const slug = context.bindingData.slug;
    const cfg = getPageConfig(slug);
    const date = String(req.query.date || '').trim(); // YYYY-MM-DD
    const duration = Math.max(15, Math.min(240, Number(req.query.duration || 30)));

    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      context.res = { status: 400, body: { error: 'Invalid or missing date' } };
      return;
    }

    // Block historical dates
    const now = DateTime.now().setZone(cfg.timeZone);
    const requestedDate = DateTime.fromISO(date, { zone: cfg.timeZone }).startOf('day');
    if (requestedDate < now.startOf('day')) {
      context.res = { status: 400, body: { error: 'Cannot book past dates' } };
      return;
    }

    // Compute window in configured time zone business hours
    const { startLocal, endLocal } = clampToBusinessHours(date, cfg.businessHours, cfg.timeZone);

    // Convert local start to UTC ISO for querying availability
    const windowStartIso = startLocal.toUTC().toISO();
    const windowEndIso = endLocal.toUTC().toISO();

    // ... (call your calendar API and build availabilityView string here) ...

    // Example: placeholder combined availability string
    const combined = "0".repeat(32); // All free for 8 hours in 15-min blocks

    let slots = slotsFromAvailability(combined, windowStartIso, 15, duration, cfg.timeZone);

    // Filter out slots whose end is before now in the target timezone
    slots = slots.filter(slot => {
      const slotEnd = DateTime.fromISO(slot.end, { zone: cfg.timeZone });
      return slotEnd > now;
    });

    context.res = { status: 200, headers: { 'content-type': 'application/json' }, body: { slug, timeZone: cfg.timeZone, date, duration, slots } };
  } catch (err) {
    context.log.error(err.stack || err.message || String(err));
    context.res = { status: 500, body: { error: err.message || 'Internal error' } };
  }
};
