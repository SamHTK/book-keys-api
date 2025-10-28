const { ClientCertificateCredential } = require("@azure/identity");
const { getPageConfig } = require("../lib/pages");

async function getToken() {
  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const pem = process.env.CERT_PEM;
  if (!tenantId || !clientId || !pem) throw new Error("Missing TENANT_ID, CLIENT_ID, or CERT_PEM");
  const cred = new ClientCertificateCredential(tenantId, clientId, { certificate: pem });
  const token = await cred.getToken("https://graph.microsoft.com/.default");
  return token.token;
}

function isWeekend(dateIso, timeZone) {
  // dateIso is YYYY-MM-DD
  // Check if it's Saturday or Sunday in the given timezone
  const dt = new Date(dateIso + 'T12:00:00'); // Use noon to avoid edge cases
  const formatter = new Intl.DateTimeFormat('en-US', { 
    timeZone: timeZone, 
    weekday: 'long' 
  });
  const dayName = formatter.format(dt);
  return dayName === 'Saturday' || dayName === 'Sunday';
}

function clampToBusinessHours(dateIso, { start, end }, timeZone) {
  // dateIso is YYYY-MM-DD in local browser, we construct start/end in tz
  // We ask Graph in the configured Windows tz, so we compute strings without Z
  const d = new Date(dateIso + 'T00:00:00Z');
  const y = d.getUTCFullYear();
  const m = String(d.getUTCMonth() + 1).padStart(2, '0');
  const day = String(d.getUTCDate()).padStart(2, '0');
  const startLocal = `${y}-${m}-${day}T${start}:00`;
  const endLocal = `${y}-${m}-${day}T${end}:00`;
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
  const start = new Date(windowStart);

  const step = intervalMinutes; // step between slot starts
  const out = [];

  // Get current time in the target timezone
  const now = new Date();
  
  for (let minute = 0; minute <= (blocks.length * blockMinutes - durationMinutes); minute += step) {
    const startDt = new Date(start.getTime() + minute * 60000);
    
    // Skip slots in the past
    if (startDt <= now) continue;
    
    const neededBlocks = Math.ceil(durationMinutes / blockMinutes);
    const startBlock = Math.floor(minute / blockMinutes);
    let ok = true;
    for (let k = 0; k < neededBlocks; k++) {
      if (!blocks[startBlock + k]) { ok = false; break; }
    }
    if (ok) {
      const endDt = new Date(startDt.getTime() + durationMinutes * 60000);
      out.push({ start: startDt.toISOString(), end: endDt.toISOString() });
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

    // Check if the requested date is a weekend
    if (isWeekend(date, cfg.timeZone)) {
      context.res = { 
        status: 200, 
        headers: { 'content-type': 'application/json' }, 
        body: { slug, timeZone: cfg.timeZone, date, duration, slots: [] } 
      };
      return;
    }

    // Compute window in configured time zone business hours
    const { startLocal, endLocal } = clampToBusinessHours(date, cfg.businessHours, cfg.timeZone);

    const body = {
      schedules: cfg.calendars,
      startTime: { dateTime: startLocal, timeZone: cfg.timeZone },
      endTime: { dateTime: endLocal, timeZone: cfg.timeZone },
      availabilityViewInterval: 15
    };

    const token = await getToken();
    const resp = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(cfg.schedulerUpn)}/calendar/getSchedule`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
    const json = await resp.json();
    if (!resp.ok) {
      context.log.error('getSchedule error', resp.status, json);
      context.res = { status: resp.status, body: json };
      return;
    }

    const views = (json.value || []).map(x => String(x.availabilityView || ''));
    const combined = intersectBusy(views);

    // Better: Re-call time server to convert the local startLocal in cfg.timeZone to UTC using Intl API by formatting parts to get offset. We do a reasonable approximation:
    function toIsoFromLocal(localStr, tz) {
      try {
        // localStr = 'YYYY-MM-DDTHH:mm:00'
        const [d, t] = localStr.split('T');
        const [Y, M, D] = d.split('-').map(Number);
        const [h, m] = t.split(':').map(Number);
        // Construct a Date in the target tz by formatting a UTC date that yields same local parts
        // Find UTC timestamp whose parts in tz equal Y-M-D h:m
        const base = Date.UTC(Y, M - 1, D, h, m, 0);
        const fmt = new Intl.DateTimeFormat('en-US', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hour12: false });
        function parts(ts) {
          const p = fmt.formatToParts(new Date(ts));
          const get = (type) => Number(p.find(x => x.type === type).value);
          return { Y: get('year'), M: get('month'), D: get('day'), h: get('hour'), m: get('minute') };
        }
        let lo = base - 48*3600*1000, hi = base + 48*3600*1000;
        for (let i = 0; i < 30; i++) {
          const mid = Math.floor((lo + hi) / 2);
          const pp = parts(mid);
          const cmp = (pp.Y - Y) || (pp.M - M) || (pp.D - D) || (pp.h - h) || (pp.m - m);
          if (cmp < 0) lo = mid + 1; else if (cmp > 0) hi = mid - 1; else return new Date(mid).toISOString();
        }
        return new Date(base).toISOString();
      } catch {
        return new Date(localStr).toISOString();
      }
    }

    const windowStartIso = toIsoFromLocal(startLocal, cfg.timeZone);
    const slots = slotsFromAvailability(combined, windowStartIso, 15, duration, cfg.timeZone);

    context.res = { status: 200, headers: { 'content-type': 'application/json' }, body: { slug, timeZone: cfg.timeZone, date, duration, slots } };
  } catch (err) {
    context.log.error(err.stack || err.message || String(err));
    context.res = { status: 500, body: { error: err.message || 'Internal error' } };
  }
};
