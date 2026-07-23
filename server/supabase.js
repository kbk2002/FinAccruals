const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;

if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY) {
  throw new Error("SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY are required for secure token storage.");
}

const API_BASE = `${SUPABASE_URL.replace(/\/$/, "")}/rest/v1`;
const DEFAULT_HEADERS = {
  apikey: SUPABASE_SERVICE_ROLE_KEY,
  Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`,
  Accept: "application/json",
  "Content-Type": "application/json",
};

async function supabaseFetch(path, options = {}) {
  const url = `${API_BASE}${path}`;
  const response = await fetch(url, {
    headers: { ...DEFAULT_HEADERS, ...(options.headers || {}) },
    ...options,
  });

  const text = await response.text();
  const body = text ? JSON.parse(text) : null;

  if (!response.ok) {
    const message = body?.message || body?.error || response.statusText;
    throw new Error(`Supabase request failed: ${message}`);
  }

  return body;
}

export async function getQboSession(sessionId) {
  if (!sessionId) return null;
  const encoded = encodeURIComponent(sessionId);
  const rows = await supabaseFetch(`/qbo_sessions?session_id=eq.${encoded}&select=session_id,payload`);
  return Array.isArray(rows) && rows.length > 0 ? rows[0] : null;
}

export async function upsertQboSession(sessionId, payload) {
  if (!sessionId) throw new Error("sessionId is required to save a QBO session.");
  await supabaseFetch(`/qbo_sessions?on_conflict=session_id`, {
    method: "POST",
    headers: { Prefer: "return=minimal" },
    body: JSON.stringify([{ session_id: sessionId, payload }]),
  });
}

export async function deleteQboSession(sessionId) {
  if (!sessionId) return;
  const encoded = encodeURIComponent(sessionId);
  await supabaseFetch(`/qbo_sessions?session_id=eq.${encoded}`, {
    method: "DELETE",
  });
}
