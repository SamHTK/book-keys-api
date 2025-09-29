const crypto = require("crypto");

// Base64url encode without padding
function b64url(buf) {
  return buf.toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

function generateApprovalToken() {
  return b64url(crypto.randomBytes(32));
}

// Non-reversible hash: SHA-256 over version|requestId|pepper|token
function computeTokenHash(requestId, token, pepper) {
  const h = crypto.createHash("sha256");
  h.update("v1|");
  h.update(requestId);
  h.update("|");
  h.update(pepper);
  h.update("|");
  h.update(token);
  return h.digest("hex");
}

function timingSafeEqualHex(aHex, bHex) {
  try {
    const a = Buffer.from(aHex, "hex");
    const b = Buffer.from(bHex, "hex");
    if (a.length !== b.length) return false;
    return crypto.timingSafeEqual(a, b);
  } catch {
    return false;
  }
}

module.exports = { generateApprovalToken, computeTokenHash, timingSafeEqualHex };
