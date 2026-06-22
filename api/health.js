export default function handler(req, res) {
  res.status(200).json({
    status: "ok",
    app: "FinAccruals API",
    message: "Backend API is running"
  });
}
