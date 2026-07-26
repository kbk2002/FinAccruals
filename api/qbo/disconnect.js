import { traceApiCall } from "../../server/logging.js";
import { activeSession, revokeSession } from "../../server/qbo.js";
import { clearSession } from "../../server/session.js";
import { markQboDisconnected } from "../../server/supabase.js";

async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Method not allowed." });
  }

  const session = await activeSession(req, res);

  try {
    await revokeSession(session);
    await markQboDisconnected(session?.realmId);
  } catch {
    // The local session must still be removed if Intuit or Supabase already invalidated the token.
  }

  clearSession(res);
  return res.status(200).json({ success: true });
}

export default traceApiCall(handler);
