const { ClientCertificateCredential } = require("@azure/identity");

let cached = { token: null, exp: 0 };

async function getGraphToken() {
  const now = Math.floor(Date.now() / 1000);
  if (cached.token && now < cached.exp - 60) return cached.token;
  
  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const pem = process.env.CERT_PEM;
  
  if (!tenantId || !clientId || !pem) {
    throw new Error("Missing TENANT_ID, CLIENT_ID, or CERT_PEM");
  }
  
  const cred = new ClientCertificateCredential(tenantId, clientId, {
    certificate: pem
  });
  
  const scope = "https://graph.microsoft.com/.default";
  const tok = await cred.getToken(scope);
  
  cached.token = tok.token;
  cached.exp = Math.floor(tok.expiresOnTimestamp / 1000);
  
  return cached.token;
}

module.exports = { getGraphToken };
