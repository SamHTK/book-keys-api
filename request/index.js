const crypto = require("crypto");
const { getGraphToken } = require("../lib/auth");
const { ensureTables, putRequestEntity, putRequestByUser } = require("../lib/storage");
const { generateApprovalToken, computeTokenHash } = require("../lib/tokens");

function htmlEscape(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function toIsoNoZ(dtIso) {
  return new Date(dtIso).toISOString().slice(0, 19);
}

module.exports = async function (context, req) {
  try {
    const scheduler = process.env.REQUEST_FROM_UPN || process.env.SCHEDULER_UPN;
    const brand = process.env.BRAND_NAME || "Booking";
    const baseUrl = process.env.BASE_URL;
    const pepper = process.env.APPROVAL_SIGNING_KEY;
    if (!scheduler || !baseUrl || !pepper) {
      throw new Error("Missing REQUEST_FROM_UPN/SCHEDULER_UPN or BASE_URL or APPROVAL_SIGNING_KEY");
    }

    const body = req.body || {};
    const ownerUserId = String(body.ownerUserId || "single-tenant"); // tie to your auth later
    const slug = String(body.slug || "").trim();
    const execEmail = String(body.execEmail || "").trim().toLowerCase();
    const title = String(body.title || "").trim();
    const start = body.start;
    const end = body.end;
    const wantsTeams = !!body.wantsTeams;
    const customerEmail = String((body.customer && body.customer.email) || "").trim().toLowerCase();
    const customerName = String((body.customer && body.customer.name) || "").trim();
    const attendees = Array.isArray(body.attendees) 
      ? body.attendees.map(a => String(a).trim().toLowerCase()).filter(Boolean) 
      : [];
    const notes = String(body.notes || "");

    if (!slug || !execEmail || !title || !start || !end || !customerEmail || !customerName) {
      context.res = {
        status: 400,
        body: { error: "Missing required fields: slug, execEmail, title, start, end, customer{name,email}" }
      };
      return;
    }

    // Create pending request row
    await ensureTables();
    const id = crypto.randomUUID();
    const token = generateApprovalToken();
    const tokenHash = computeTokenHash(id, token, pepper);
    const nowIso = new Date().toISOString();
    const expiresAt = new Date(Date.now() + 48 * 3600 * 1000).toISOString(); // 48h

    const allMailboxesCsv = (Array.isArray(body.allMailboxes) ? body.allMailboxes : [])
      .map(s => String(s).trim().toLowerCase())
      .filter(Boolean)
      .join(",");

    const entity = {
      partitionKey: id,
      rowKey: "r",
      ownerUserId,
      slug,
      targetExecUpn: execEmail,
      title,
      start,
      end,
      customerEmail,
      customerName,
      attendeesCsv: attendees.join(","),
      wantsTeams: wantsTeams ? "true" : "false",
      status: "pending",
      createdAt: nowIso,
      expiresAt,
      tokenHash,
      allMailboxesCsv,
      transactionId: id
    };
    await putRequestEntity(entity);
    await putRequestByUser(ownerUserId, id, {
      createdAt: nowIso,
      status: "pending",
      targetExecUpn: execEmail,
      start,
      end
    });

    // Build approval links
    const approveUrl = `${baseUrl}/api/requests/${encodeURIComponent(id)}/accept?t=${encodeURIComponent(token)}`;
    const declineUrl = `${baseUrl}/api/requests/${encodeURIComponent(id)}/decline?t=${encodeURIComponent(token)}`;

    // Build email HTML
    const whenHtml = `${htmlEscape(start)} – ${htmlEscape(end)} (UTC)`;
    const attendeesHtml = [customerEmail, ...attendees].map(htmlEscape).join(", ");
    const html = `
  <div style="font-family:Segoe UI,Arial,sans-serif">
    <h2 style="margin:0 0 8px 0">${htmlEscape(brand)}: Approval required</h2>
    <p style="margin:4px 0"><b>Title:</b> ${htmlEscape(title)}</p>
    <p style="margin:4px 0"><b>When:</b> ${whenHtml}</p>
    <p style="margin:4px 0"><b>Organizer (target):</b> ${htmlEscape(execEmail)}</p>
    <p style="margin:4px 0"><b>Attendees:</b> ${attendeesHtml}</p>
    <p style="margin:4px 0"><b>Teams requested:</b> ${wantsTeams ? "Yes" : "No"}</p>
    ${notes ? `<p style="margin:4px 0"><b>Notes:</b> ${htmlEscape(notes)}</p>` : ""}
    <p style="margin:16px 0">
      <a href="${approveUrl}" style="background:#2563eb;color:#fff;padding:10px 16px;border-radius:6px;text-decoration:none">Approve</a>
      <a href="${declineUrl}" style="margin-left:12px;color:#6b7280;text-decoration:none">Decline</a>
    </p>
  </div>
`;

    // Send email to exec (from scheduler)
    const graphToken = await getGraphToken();
    const payload = {
      message: {
        subject: `Approval needed: ${title}`,
        body: { contentType: "HTML", content: html },
        toRecipients: [{ emailAddress: { address: execEmail } }],
        internetMessageHeaders: [{ name: "X-BookKeys-Request-Id", value: id }]
      },
      saveToSentItems: true
    };

    const resp = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(scheduler)}/sendMail`, {
      method: "POST",
      headers: { Authorization: `Bearer ${graphToken}`, "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    const result = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      context.log.error("sendMail error", resp.status, result);
      context.res = { status: resp.status, body: result };
      return;
    }

    context.res = {
      status: 200,
      headers: { "content-type": "application/json" },
      body: { requestId: id }
    };

  } catch (err) {
    context.log.error(err.stack || err.message || String(err));
    context.res = {
      status: 500,
      body: { error: err.message || "Internal error" }
    };
  }
};
