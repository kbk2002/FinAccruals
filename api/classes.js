import { activeSession, qboQuery } from "../server/qbo.js";

export default async function handler(req, res) {
  try {
    const session = await activeSession(req, res);
    if (!session) return res.status(401).json({ error: "QuickBooks is not connected." });

    const result = await qboQuery(session, "select * from Class where Active = true maxresults 1000");
    const classes = (result.Class || []).map((item) => ({
      id: item.Id,
      name: item.FullyQualifiedName || item.Name,
    }));
    return res.status(200).json({ success: true, classes });
  } catch (error) {
    return res.status(502).json({ error: error.message });
  }
}
