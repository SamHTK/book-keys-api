const { ClientCertificateCredential } = require("@azure/identity");
const crypto = require("crypto");

function toUtcBasic(dtIso) {
  const d = new Date(dtIso);
  const yyyy = d.getUTCFullYear();
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  const dd = String(d.getUTCDate()).padStart(2, "0");
  const hh = String(d.getUTCHours()).padStart(2, "0");
  const mi = String(d.getUTCMinutes()).padStart(2, "0");
  const ss = String(d.getUTCSeconds()).padStart(2, "0");
  return `${yyyy}${mm}${dd}T${hh}${mi}${ss}Z`;
}

function icsEscape(s) {
  return String(s)
    .replace(/\\/g, "\\\\") // escape backslashes
    .replace(/\n/g, "\\n")  // proper newline escape
    .replace(/,/g, "\\,")
    .replace(/;/g, "\\;");
}

function foldIcsLine(line) {
  const max = 75;
  if (line.length <= max) return line;
  let out = line.slice(0, max);
  for (let i = max; i < line.length; i += max) {
    out += "\r\n " + line.slice(i, i + max);
  }
  return out;
}

function buildIcsRequest({ uid, organizer, attendee, startIso, endIso, subject, description }) {
  const dtstamp = toUtcBasic(new Date().toISOString());
  const dtstart = toUtcBasic(startIso);
  const dtend = toUtcBasic(endIso);
  const lines = [
    "BEGIN:VCALENDAR",
    "PRODID:-//BookKeys//EN",
    "VERSION:2.0",
    "CALSCALE:GREGORIAN",
    "METHOD:REQUEST",
    "BEGIN:VEVENT",
    `UID:${uid}`,
    `DTSTAMP:${dtstamp}`,
    `DTSTART:${dtstart}`,
    `DTEND:${dtend}`,
    foldIcsLine(`SUMMARY:${icsEscape(subject)}`),
    foldIcsLine(`DESCRIPTION:${icsEscape(description)}`),
    `ORGANIZER:MAILTO:${organizer}`,
    foldIcsLine(`ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION:MAILTO:${attendee}`),
    "END:VEVENT",
    "END:VCALENDAR"
  ];
  return lines.join("\r\n") + "\r\n";
}

async function getToken() {
  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const pem = process.env.CERT_PEM;
  if (!tenantId || !clientId || !pem) throw new Error("Missing TENANT_ID, CLIENT_ID, or CERT_PEM");

  const cred = new ClientCertificateCredential(tenantId, clientId, { certificate: pem });
  const token = await cred.getToken("https://graph.microsoft.com/.default");
  return token.token;
}

module.exports = async function (context, req) {
  try {
    const fromUpn = process.env.REQUEST_FROM_UPN || process.env.SCHEDULER_UPN;
    if (!fromUpn) throw new Error("Missing REQUEST_FROM_UPN or SCHEDULER_UPN app setting.");

    const body = req.body || {};
    const execEmail = String(body.execEmail || "").trim().toLowerCase();
    const start = body.start;
    const end = body.end;
    const customer = body.customer || {};
    const custName = String(customer.name || "").trim();
    const custEmail = String(customer.email || "").trim().toLowerCase();
    const notes = String(body.notes || "").trim();

    if (!execEmail || !start || !end || !custName || !custEmail) {
      context.res = {
        status: 400,
        body: { error: "Missing required fields: execEmail, start, end, customer{name,email}" }
      };
      return;
    }

    const uid = crypto.randomUUID();
    const brand = process.env.BRAND_NAME || "Booking";
    const subject = `${brand} request: ${custName}`;
    const desc = [
      `Customer: ${custName}`,
      `Email: ${custEmail}`,
      notes ? `Notes: ${notes}` : ""
    ]
      .filter(Boolean)
      .join("\n");

    const ics = buildIcsRequest({
      uid,
      organizer: fromUpn,
      attendee: execEmail,
      startIso: start,
      endIso: end,
      subject,
      description: desc
    });

    const base64Ics = Buffer.from(ics, "utf8").toString("base64");
    const message = {
      message: {
        subject,
        body: { contentType: "Text", content: `${brand} meeting request attached.` },
        toRecipients: [{ emailAddress: { address: execEmail } }],
        internetMessageHeaders: [{ name: "X-BookKeys-Request-Uid", value: uid }],
        attachments: [
          {
            "@odata.type": "#microsoft.graph.fileAttachment",
            name: "request.ics",
            contentType: "text/calendar; method=REQUEST; charset=UTF-8",
            contentBytes: base64Ics
          }
        ]
      },
      saveToSentItems: true
    };

    const token = await getToken();
    const resp = await fetch(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(fromUpn)}/sendMail`,
      {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify(message)
      }
    );

    if (!resp.ok) {
      const errJson = await resp.json().catch(() => ({}));
      context.log.error("sendMail error", resp.status, errJson);
      context.res = { status: resp.status, body: errJson };
      return;
    }

    context.res = {
      status: 200,
      headers: { "content-type": "application/json" },
      body: { uid }
    };
  } catch (err) {
    context.log.error(err.stack || err.message || String(err));
    context.res = { status: 500, body: { error: err.message || "Internal error" } };
  }
};
