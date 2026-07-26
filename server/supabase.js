import { decryptSecret, encryptSecret } from "./crypto.js";

const QBO_CONNECTIONS_TABLE = "qbo_connections";

function qboEnvironment() {
  return process.env.QBO_ENVIRONMENT === "production" ? "production" : "sandbox";
}

function requiredSupabaseConfig() {
  const url = process.env.SUPABASE_URL;
  const serviceRoleKey = process.env.SUPABASE_SERVICE_ROLE_KEY;

  if (!url || !serviceRoleKey) {
    throw new Error("Supabase environment variables are incomplete.");
  }

  return {
    url: url.replace(/\/$/, ""),
    serviceRoleKey,
  };
}

async function supabaseRequest(path, options = {}) {
  const { url, serviceRoleKey } = requiredSupabaseConfig();
  const response = await fetch(`${url}/rest/v1/${path}`, {
    ...options,
    headers: {
      apikey: serviceRoleKey,
      Authorization: `Bearer ${serviceRoleKey}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });

  if (response.status === 204) return null;

  const payload = await response.json().catch(() => null);

  if (!response.ok) {
    const message = payload?.message || payload?.hint || "Supabase request failed.";
    throw new Error(message);
  }

  return payload;
}

function toIsoDate(value) {
  return value ? new Date(value).toISOString() : null;
}

function sessionFromRecord(record) {
  if (!record) return null;

  return {
    realmId: record.realm_id,
    companyName: record.company_name || "",
    accessToken: decryptSecret(record.encrypted_access_token),
    refreshToken: decryptSecret(record.encrypted_refresh_token),
    accessTokenExpiresAt: record.access_token_expires_at
      ? new Date(record.access_token_expires_at).getTime()
      : 0,
    refreshTokenExpiresAt: record.refresh_token_expires_at
      ? new Date(record.refresh_token_expires_at).getTime()
      : 0,
  };
}

export function safeSession(session) {
  if (!session?.realmId) return null;

  return {
    realmId: session.realmId,
    companyName: session.companyName || "",
  };
}

export async function saveQboConnection(session) {
  if (!session?.realmId || !session?.accessToken || !session?.refreshToken) {
    throw new Error("Cannot save incomplete QuickBooks connection.");
  }

  await supabaseRequest(`${QBO_CONNECTIONS_TABLE}?on_conflict=realm_id`, {
    method: "POST",
    headers: {
      Prefer: "resolution=merge-duplicates",
    },
    body: JSON.stringify({
      realm_id: session.realmId,
      company_name: session.companyName || "",
      environment: qboEnvironment(),
      encrypted_access_token: encryptSecret(session.accessToken),
      encrypted_refresh_token: encryptSecret(session.refreshToken),
      access_token_expires_at: toIsoDate(session.accessTokenExpiresAt),
      refresh_token_expires_at: toIsoDate(session.refreshTokenExpiresAt),
      disconnected_at: null,
      updated_at: new Date().toISOString(),
    }),
  });
}

export async function loadQboConnection(realmId) {
  if (!realmId) return null;

  const environment = encodeURIComponent(qboEnvironment());
  const encodedRealmId = encodeURIComponent(realmId);
  const rows = await supabaseRequest(
    `${QBO_CONNECTIONS_TABLE}?realm_id=eq.${encodedRealmId}&environment=eq.${environment}&disconnected_at=is.null&select=*`
  );

  return sessionFromRecord(rows?.[0]);
}

export async function loadLatestQboConnection() {
  const environment = encodeURIComponent(qboEnvironment());
  const rows = await supabaseRequest(
    `${QBO_CONNECTIONS_TABLE}?environment=eq.${environment}&disconnected_at=is.null&select=*&order=updated_at.desc&limit=1`
  );

  return sessionFromRecord(rows?.[0]);
}

export async function markQboDisconnected(realmId) {
  if (!realmId) return;

  const encodedRealmId = encodeURIComponent(realmId);
  await supabaseRequest(`${QBO_CONNECTIONS_TABLE}?realm_id=eq.${encodedRealmId}`, {
    method: "PATCH",
    body: JSON.stringify({
      disconnected_at: new Date().toISOString(),
      updated_at: new Date().toISOString(),
    }),
  });
}
