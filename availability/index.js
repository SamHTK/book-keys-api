const { ClientCertificateCredential } = require("@azure/identity");
async function getToken() { const tenantId = process.env.TENANT_ID; const clientId = process.env.CLIENT_ID; const pem = process.env.CERT_PEM; if (!tenantId || !clientId || !pem) throw new Error("Missing TENANT_ID, CLIENT_ID, or CERT_PEM"); const cred = new ClientCertificateCredential(tenantId, clientId, { certificate: pem }); const token = await cred.getToken("https://graph.microsoft.com/.default"); return token.token; }

module.exports = async function (context, req) { try { const schedules = (process.env.AVAILABILITY_SCHEDULES || "") .split(",").map(s => s.trim()).filter(Boolean); const tz = process.env.AVAILABILITY_TIMEZONE || "Eastern Standard Time"; if (schedules.length === 0) throw new Error("AVAILABILITY_SCHEDULES is empty");

const userInUrl = process.env.SCHEDULER_UPN || schedules[0];

// Defaults if not provided
const now = new Date();
const start = (req.query.start || new Date(now.setHours(8, 0, 0, 0)).toISOString().slice(0, 19));
const end = (req.query.end || new Date(Date.now() + 24 * 3600 * 1000).toISOString().slice(0, 19));
const interval = Number(req.query.interval || 15);

const token = await getToken();

const body = {
  schedules,
  startTime: { dateTime: start, timeZone: tz },
  endTime:   { dateTime: end,   timeZone: tz },
  availabilityViewInterval: interval
};

const resp = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userInUrl)}/calendar/getSchedule`, {
  method: "POST",
  headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
  body: JSON.stringify(body)
});

const data = await resp.json();
if (!resp.ok) {
  context.log.error("getSchedule error", resp.status, data);
  context.res = { status: resp.status, body: data };
  return;
}

context.res = { status: 200, headers: { "content-type": "application/json" }, body: data };

} catch (err) { context.log.error(err); context.res = { status: 500, body: { error: err.message } }; } };
