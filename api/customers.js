import { activeSession, qboQuery } from "../server/qbo.js";

export default async function handler(req, res) {
  try {
    const session = await activeSession(req, res);
    if (!session) return res.status(401).json({ error: "QuickBooks is not connected." });

    const result = await qboQuery(session, "select * from Customer where Active = true maxresults 1000");
    const customers = (result.Customer || []).map((customer) => ({
      id: customer.Id,
      name: customer.DisplayName || customer.CompanyName || "",
      company: customer.CompanyName || "",
      email: customer.PrimaryEmailAddr?.Address || "",
      phone: customer.PrimaryPhone?.FreeFormNumber || "",
      balance: customer.Balance || 0,
    }));
    return res.status(200).json({ success: true, customers });
  } catch (error) {
    return res.status(502).json({ error: error.message });
  }
}
