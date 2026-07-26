import { traceApiCall } from "../server/logging.js";

async function handler(req, res) {
  return res.status(200).json({ success: true, history: [] });
}

export default traceApiCall(handler);
