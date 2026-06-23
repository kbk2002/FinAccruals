import { activeSession, qboQuery } from "./lib/qbo.js";

export default async function handler(req, res) {
  try {
    const session = await activeSession(req, res);
    if (!session) return res.status(401).json({ error: "QuickBooks is not connected." });

    const result = await qboQuery(session, "select * from Vendor where Active = true maxresults 1000");
    const vendors = (result.Vendor || []).map((vendor) => ({
      id: vendor.Id,
      name: vendor.DisplayName || vendor.CompanyName || "",
      email: vendor.PrimaryEmailAddr?.Address || "",
    }));
    return res.status(200).json({ success: true, vendors });
  } catch (error) {
    return res.status(502).json({ error: error.message });
  }
}
