import { activeSession, qboQuery } from "../server/qbo.js";

export default async function handler(req, res) {
  try {
    const session = await activeSession(req, res);
    if (!session) return res.status(401).json({ error: "QuickBooks is not connected." });

    const result = await qboQuery(session, "select * from Account where Active = true maxresults 1000");
    const accounts = (result.Account || []).map((account) => ({
      id: account.Id,
      code: account.AcctNum || account.Id,
      name: account.FullyQualifiedName || account.Name,
      type: account.AccountType || "",
    }));
    return res.status(200).json({ success: true, accounts });
  } catch (error) {
    return res.status(502).json({ error: error.message });
  }
}
