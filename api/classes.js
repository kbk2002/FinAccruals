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
    const query = `select * from Class where Active = true maxresults ${limit} startposition ${offset + 1}`;
    const result = await qboQuery(session, query);

    const classes = (result.Class || []).map((item) => ({
      id: item.Id,
      name: item.FullyQualifiedName || item.Name,
    }));

    return res.status(200).json({
      success: true,
      classes,
      limit,
      offset,
      count: classes.length,
      totalCount: result.totalCount ?? null,
    });
  } catch (error) {
    return res.status(502).json({ error: error.message });
  }
}
