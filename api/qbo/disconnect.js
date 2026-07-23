import { revokeSession } from "../../server/qbo.js";
import { clearSession, readSession } from "../../server/session.js";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Method not allowed." });
  }

  const session = await readSession(req);

  try {
    await revokeSession(session);
  } catch {
    // The local session must still be removed if Intuit already invalidated the token.
  }

  await clearSession(res, req);
  return res.status(200).json({ success: true });
}
