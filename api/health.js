export default function handler(req, res) {
  res.status(200).json({
    status: "healthy",
    app: "FinAccruals API",
    message: "Vercel backend is running"
  });
}
