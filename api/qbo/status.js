import { activeSession, loadCompanyName } from "../lib/qbo.js";
import { writeSession } from "../lib/session.js";

export default async function handler(req, res) {
  if (req.method !== "GET") {
    res.setHeader("Allow", "GET");
    return res.status(405).json({ error: "Method not allowed." });
  }

  try {
    const session = await activeSession(req, res);
    if (!session) return res.status(200).json({ connected: false });

    if (!session.companyName) {
      session.companyName = await loadCompanyName(session);
      writeSession(res, session);
    }

    return res.status(200).json({
      connected: true,
      companyName: session.companyName,
      realmId: session.realmId,
      environment: process.env.QBO_ENVIRONMENT === "production" ? "production" : "sandbox",
    });
  } catch (error) {
    return res.status(200).json({ connected: false, error: error.message });
  }
}
