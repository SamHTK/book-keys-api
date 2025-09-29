const { getRequestEntity, updateRequestStatus } = require("../lib/storage");
const { computeTokenHash, timingSafeEqualHex } = require("../lib/tokens");

function htmlPage(title, body) {
  return `<!doctype html><html><head><meta charset="utf-8"><title>${title}</title></head><body style="font-family:Segoe UI,Arial,sans-serif">${body}</body></html>`;
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
        body: htmlPage("Invalid", "<p>Invalid decline link.</p>")
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
        body: htmlPage("Expired", "<p>This link has expired.</p>")
      };
      return;
    }

    const expectedHash = computeTokenHash(id, token, pepper);
    if (!timingSafeEqualHex(expectedHash, ent.tokenHash)) {
      context.res = {
        status: 400,
        headers: { "content-type": "text/html" },
        body: htmlPage("Invalid", "<p>Invalid token.</p>")
      };
      return;
    }

    await updateRequestStatus(id, { status: "declined", lastActionAt: new Date().toISOString() });
    context.res = {
      status: 200,
      headers: { "content-type": "text/html" },
      body: htmlPage("Declined", "<p>Request declined.</p>")
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
