import { traceApiCall } from "../server/logging.js";

async function handler(req, res) {
  res.status(200).json({
    status: "ok",
    app: "FinAccruals API",
    message: "Backend API is running"
  });
}

export default traceApiCall(handler);
