import { authorizationUrl } from "../../server/qbo.js";
import { createOAuthState } from "../../server/session.js";

export default function handler(req, res) {
  if (req.method !== "GET") {
    res.setHeader("Allow", "GET");
    return res.status(405).json({ error: "Method not allowed." });
  }

  try {
    const state = createOAuthState(res);
    return res.redirect(302, authorizationUrl(state));
  } catch (error) {
    return res.status(500).send(`Unable to begin QuickBooks authentication: ${error.message}`);
  }
}
