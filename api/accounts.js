import { activeSession, qboQuery } from "../server/qbo.js";

const DEFAULT_MAX_RESULTS = 1000;

function parsePaging(req) {
  const requestedLimit = Number.parseInt(String(req.query.limit || ""), 10);
  const requestedOffset = Number.parseInt(String(req.query.offset || ""), 10);

  const limit = Number.isInteger(requestedLimit) && requestedLimit > 0 ? Math.min(requestedLimit, DEFAULT_MAX_RESULTS) : DEFAULT_MAX_RESULTS;
  const offset = Number.isInteger(requestedOffset) && requestedOffset >= 0 ? requestedOffset : 0;

  return { limit, offset };
}

export default async function handler(req, res) {
  try {
    const session = await activeSession(req, res);
    if (!session) return res.status(401).json({ error: "QuickBooks is not connected." });

    const { limit, offset } = parsePaging(req);
    const query = `select * from Account where Active = true maxresults ${limit} startposition ${offset + 1}`;
    const result = await qboQuery(session, query);

    const accounts = (result.Account || []).map((account) => ({
      id: account.Id,
      code: account.AcctNum || account.Id,
      name: account.FullyQualifiedName || account.Name,
      type: account.AccountType || "",
    }));

    return res.status(200).json({
      success: true,
      accounts,
      limit,
      offset,
      count: accounts.length,
      totalCount: result.totalCount ?? null,
    });
  } catch (error) {
    return res.status(502).json({ error: error.message });
  }
}
