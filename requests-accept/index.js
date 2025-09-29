const { getGraphToken } = require("../lib/auth");
const { getRequestEntity, updateRequestStatus } = require("../lib/storage");
const { computeTokenHash, timingSafeEqualHex } = require("../lib/tokens");

function htmlPage(title, body) {
  return `<!doctype html><html><head><meta charset="utf-8"><title>${title}</title></head><body style="font-family:Segoe UI,Arial,sans-serif">${body}</body></html>`;
}

function toIsoNoZ(dtIso) {
  return new Date(dtIso).toISOString().slice(0, 19);
}

module.exports = async function (context, req) {
  try {
    const id = context.bindingData.id;
    const token = (req.query && req.query.t) || "";
    const pepper = process.env.APPROVAL_SIGNING_KEY;
    if (!id || !token || !pepper) {
      context.res = {
        status: 400,
        headers: { "content-type": "text/html" },
        body: htmlPage("Invalid", "<p>Invalid approval link.</p>")
      };
      return;
    }

    let ent;
    try {
      ent = await getRequestEntity(id);
    } catch {
      ent = null;
    }

    if (!ent) {
      context.res = {
        status: 404,
        headers: { "content-type": "text/html" },
        body: htmlPage("Not found", "<p>Request not found.</p>")
      };
      return;
    }

    if (ent.status !== "pending") {
      context.res = {
        status: 200,
        headers: { "content-type": "text/html" },
        body: htmlPage("Processed", "<p>This request has already been processed.</p>")
      };
      return;
    }

    const now = new Date();
    if (ent.expiresAt && new Date(ent.expiresAt) < now) {
      await updateRequestStatus(id, { status: "expired", lastActionAt: now.toISOString() });
      context.res = {
        status: 200,
        headers: { "content-type": "text/html" },
        body: htmlPage("Expired", "<p>This approval link has expired.</p>")
      };
      return;
    }

    const expectedHash = computeTokenHash(id, token, pepper);
    if (!timingSafeEqualHex(expectedHash, ent.tokenHash)) {
      context.res = {
        status: 400,
        headers: { "content-type": "text/html" },
        body: htmlPage("Invalid", "<p>Invalid approval token.</p>")
      };
      return;
    }

    // Create final event in exec's calendar
    const exec = ent.targetExecUpn;
    const title = ent.title;
    const bodyHtml = `
  <div style="font-family:Segoe UI,Arial,sans-serif">
    <p>Approved booking</p>
    <p><b>Title:</b> ${title}</p>
    <p><b>Customer:</b> ${ent.customerName} (${ent.customerEmail})</p>
  </div>
`;
    const attendees = [ent.customerEmail, ...(ent.attendeesCsv ? ent.attendeesCsv.split(",").filter(Boolean) : [])];
    const eventPayload = {
      subject: title,
      body: { contentType: "HTML", content: bodyHtml },
      start: { dateTime: toIsoNoZ(ent.start), timeZone: "UTC" },
      end: { dateTime: toIsoNoZ(ent.end), timeZone: "UTC" },
      attendees: attendees.map(a => ({ emailAddress: { address: a }, type: "required" })),
      responseRequested: true,
      transactionId: ent.transactionId || id
    };

    if (ent.wantsTeams === "true") {
      // Try Teams meeting; if it fails (license), we'll catch and fallback
      eventPayload.isOnlineMeeting = true;
      eventPayload.onlineMeetingProvider = "teamsForBusiness";
    }

    const graphToken = await getGraphToken();
    let resp = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(exec)}/events`, {
      method: "POST",
      headers: { Authorization: `Bearer ${graphToken}`, "Content-Type": "application/json" },
      body: JSON.stringify(eventPayload)
    });

    if (!resp.ok && ent.wantsTeams === "true") {
      // Fallback without Teams if license error
      const errJson = await resp.json().catch(() => ({}));
      context.log.warn("Teams creation failed, fallback", resp.status, errJson);
      delete eventPayload.isOnlineMeeting;
      delete eventPayload.onlineMeetingProvider;
      resp = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(exec)}/events`, {
        method: "POST",
        headers: { Authorization: `Bearer ${graphToken}`, "Content-Type": "application/json" },
        body: JSON.stringify(eventPayload)
      });
    }

    const result = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      context.log.error("Create exec event error", resp.status, result);
      context.res = {
        status: 502,
        headers: { "content-type": "text/html" },
        body: htmlPage("Failed", "<p>Could not finalize the meeting. The slot may be unavailable or license is missing.</p>")
      };
      return;
    }

    await updateRequestStatus(id, {
      status: "accepted",
      lastActionAt: new Date().toISOString(),
      execEventId: result.id || ""
    });

    context.res = {
      status: 200,
      headers: { "content-type": "text/html" },
      body: htmlPage("Approved", "<p>Meeting approved. Invites have been sent from the exec's calendar.</p>")
    };

  } catch (err) {
    context.log.error(err.stack || err.message || String(err));
    context.res = {
      status: 500,
      headers: { "content-type": "text/html" },
      body: htmlPage("Error", "<p>Internal error.</p>")
    };
  }
};
