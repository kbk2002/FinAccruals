import { clearSession, readSession, writeSession } from "./session.js";

const AUTH_URL = "https://appcenter.intuit.com/connect/oauth2";
const TOKEN_URL = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer";
const REVOKE_URL = "https://developer.api.intuit.com/v2/oauth2/tokens/revoke";

function requiredEnvironment() {
  const clientId = process.env.QBO_CLIENT_ID;
  const clientSecret = process.env.QBO_CLIENT_SECRET;
  const redirectUri = process.env.QBO_REDIRECT_URI;

  if (!clientId || !clientSecret || !redirectUri) {
    throw new Error("QuickBooks environment variables are incomplete.");
  }

  return { clientId, clientSecret, redirectUri };
}

function basicAuthorization() {
  const { clientId, clientSecret } = requiredEnvironment();
  return `Basic ${Buffer.from(`${clientId}:${clientSecret}`).toString("base64")}`;
}

async function tokenRequest(parameters) {
  const response = await fetch(TOKEN_URL, {
    method: "POST",
    headers: {
      Accept: "application/json",
      Authorization: basicAuthorization(),
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: new URLSearchParams(parameters),
  });
  const payload = await response.json();

  if (!response.ok) {
    throw new Error(payload.error_description || payload.error || "QuickBooks token request failed.");
  }

  return payload;
}

function createSession(tokens, realmId, previous = {}) {
  const now = Date.now();
  return {
    realmId,
    accessToken: tokens.access_token,
    refreshToken: tokens.refresh_token || previous.refreshToken,
    accessTokenExpiresAt: now + Number(tokens.expires_in || 3600) * 1000,
    refreshTokenExpiresAt:
      now + Number(tokens.x_refresh_token_expires_in || 8726400) * 1000,
    companyName: previous.companyName || "",
  };
}

export function authorizationUrl(state) {
  const { clientId, redirectUri } = requiredEnvironment();
  const parameters = new URLSearchParams({
    client_id: clientId,
    response_type: "code",
    scope: "com.intuit.quickbooks.accounting",
    redirect_uri: redirectUri,
    state,
  });
  return `${AUTH_URL}?${parameters.toString()}`;
}

export async function exchangeCode(code, realmId) {
  const { redirectUri } = requiredEnvironment();
  const tokens = await tokenRequest({
    grant_type: "authorization_code",
    code,
    redirect_uri: redirectUri,
  });
  return createSession(tokens, realmId);
}

export async function activeSession(req, res) {
  let session = readSession(req);
  if (!session?.accessToken || !session?.realmId) return null;

  if (session.accessTokenExpiresAt > Date.now() + 5 * 60 * 1000) {
    return session;
  }

  if (!session.refreshToken || session.refreshTokenExpiresAt <= Date.now()) {
    clearSession(res);
    return null;
  }

  let tokens;
  try {
    tokens = await tokenRequest({
      grant_type: "refresh_token",
      refresh_token: session.refreshToken,
    });
  } catch {
    clearSession(res);
    return null;
  }

  session = createSession(tokens, session.realmId, session);
  writeSession(res, session);
  return session;
}

function apiBase(realmId) {
  const host =
    process.env.QBO_ENVIRONMENT === "production"
      ? "https://quickbooks.api.intuit.com"
      : "https://sandbox-quickbooks.api.intuit.com";
  return `${host}/v3/company/${encodeURIComponent(realmId)}`;
}

export async function qboQuery(session, query) {
  const url = new URL(`${apiBase(session.realmId)}/query`);
  url.searchParams.set("query", query);
  url.searchParams.set("minorversion", "75");

  const response = await fetch(url, {
    headers: {
      Accept: "application/json",
      Authorization: `Bearer ${session.accessToken}`,
    },
  });
  const payload = await response.json();

  if (!response.ok || payload.Fault) {
    const detail = payload.Fault?.Error?.[0]?.Detail;
    throw new Error(detail || "QuickBooks API request failed.");
  }

  return payload.QueryResponse || {};
}

export async function loadCompanyName(session) {
  const result = await qboQuery(session, "select * from CompanyInfo");
  return result.CompanyInfo?.[0]?.CompanyName || "";
}

export async function revokeSession(session) {
  if (!session?.refreshToken && !session?.accessToken) return;

  await fetch(REVOKE_URL, {
    method: "POST",
    headers: {
      Accept: "application/json",
      Authorization: basicAuthorization(),
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ token: session.refreshToken || session.accessToken }),
  });
}
