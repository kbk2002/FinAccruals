import crypto from "crypto";

const SESSION_COOKIE = "finaccruals_qbo_session";
const STATE_COOKIE = "finaccruals_qbo_state";

function secretKey() {
  const secret = process.env.SESSION_SECRET || process.env.QBO_CLIENT_SECRET;

  if (!secret) {
    throw new Error("SESSION_SECRET or QBO_CLIENT_SECRET is required.");
  }

  return crypto.createHash("sha256").update(secret).digest();
}

function parseCookies(req) {
  return String(req.headers.cookie || "")
    .split(";")
    .map((part) => part.trim())
    .filter(Boolean)
    .reduce((cookies, part) => {
      const separator = part.indexOf("=");
      if (separator === -1) return cookies;
      cookies[part.slice(0, separator)] = decodeURIComponent(part.slice(separator + 1));
      return cookies;
    }, {});
}

function serializeCookie(name, value, maxAge) {
  return [
    `${name}=${encodeURIComponent(value)}`,
    "Path=/",
    "HttpOnly",
    "Secure",
    "SameSite=None",
    `Max-Age=${maxAge}`,
  ].join("; ");
}

function appendCookie(res, cookie) {
  const current = res.getHeader("Set-Cookie");
  const cookies = current ? (Array.isArray(current) ? current : [current]) : [];
  res.setHeader("Set-Cookie", [...cookies, cookie]);
}

function encrypt(payload) {
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv("aes-256-gcm", secretKey(), iv);
  const encrypted = Buffer.concat([
    cipher.update(JSON.stringify(payload), "utf8"),
    cipher.final(),
  ]);
  const tag = cipher.getAuthTag();
  return Buffer.concat([iv, tag, encrypted]).toString("base64url");
}

function decrypt(value) {
  try {
    const data = Buffer.from(value, "base64url");
    const iv = data.subarray(0, 12);
    const tag = data.subarray(12, 28);
    const encrypted = data.subarray(28);
    const decipher = crypto.createDecipheriv("aes-256-gcm", secretKey(), iv);
    decipher.setAuthTag(tag);
    return JSON.parse(Buffer.concat([decipher.update(encrypted), decipher.final()]).toString("utf8"));
  } catch {
    return null;
  }
}

export function createOAuthState(res) {
  const state = crypto.randomBytes(32).toString("base64url");
  appendCookie(res, serializeCookie(STATE_COOKIE, encrypt({ state }), 10 * 60));
  return state;
}

export function consumeOAuthState(req, res, receivedState) {
  const stored = decrypt(parseCookies(req)[STATE_COOKIE]);
  appendCookie(res, serializeCookie(STATE_COOKIE, "", 0));
  return Boolean(stored?.state && receivedState && stored.state === receivedState);
}

export function readSession(req) {
  const value = parseCookies(req)[SESSION_COOKIE];
  return value ? decrypt(value) : null;
}

export function writeSession(res, session) {
  appendCookie(res, serializeCookie(SESSION_COOKIE, encrypt(session), 60 * 60 * 24 * 90));
}

export function clearSession(res) {
  appendCookie(res, serializeCookie(SESSION_COOKIE, "", 0));
}
