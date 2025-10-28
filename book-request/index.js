const { getPageConfig } = require("../lib/pages");

module.exports = async function (context, req) {
  try {
    const slug = context.bindingData.slug;
    const cfg = getPageConfig(slug);

    const body = req.body || {};
    const startIso = String(body.start || '').trim();
    const duration = Math.max(15, Math.min(240, Number(body.duration || 30)));
    const title = String(body.title || '').trim();
    const email = String(body.email || '').trim().toLowerCase();
    const additionalAttendees = Array.isArray(body.additionalAttendees) ? body.additionalAttendees.map(s => String(s).trim().toLowerCase()).filter(Boolean) : [];
    const wantsTeams = !!body.wantsTeams;
    const notes = String(body.notes || '');

    if (!startIso || !title || !email) {
      context.res = { status: 400, body: { error: 'Missing start, title, or email' } };
      return;
    }

    const start = new Date(startIso);
    if (isNaN(start.getTime())) {
      context.res = { status: 400, body: { error: 'Invalid start time' } };
      return;
    }
    const end = new Date(start.getTime() + duration * 60000);

    const payload = {
      slug,
      execEmail: cfg.schedulerUpn,
      title,
      start: start.toISOString(),
      end: end.toISOString(),
      wantsTeams,
      timeZone: cfg.timeZone,
      customer: { name: email, email },
      attendees: additionalAttendees,
      notes,
      allMailboxes: cfg.calendars
    };

    // Proxy to the existing function logic by invoking it in-process.
    // Simulate the same handler signature as ../request/index.js expects.
    const createRequest = require("../request/index.js");

    const fakeReq = { body: payload, query: {}, headers: {} };
    let status = 500; let bodyOut = { error: 'Internal error' };
    const fakeContext = {
      bindingData: {},
      log: context.log,
      res: undefined,
      set res(v) { this._res = v; },
      get res() { return this._res; }
    };

    await createRequest(fakeContext, fakeReq);

    const res = fakeContext.res || {};
    status = res.status || 500;
    bodyOut = res.body || bodyOut;

    context.res = { status, headers: { 'content-type': 'application/json' }, body: bodyOut };
  } catch (err) {
    context.log.error(err.stack || err.message || String(err));
    context.res = { status: 500, body: { error: err.message || 'Internal error' } };
  }
};
