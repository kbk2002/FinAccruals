import { authorizationUrl } from "../lib/qbo.js";
import { createOAuthState } from "../lib/session.js";

export default function handler(req, res) {
  if (req.method !== "GET") {
    res.setHeader("Allow", "GET");
    return res.status(405).json({ error: "Method not allowed." });
  }

  try {
    const state = createOAuthState(res);
    return res.status(200).json({ url: authorizationUrl(state) });
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
}
