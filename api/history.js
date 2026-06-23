export default function handler(req, res) {
  res.status(200).json({
    success: true,
    history: [
      {
        id: "JE-001",
        date: "2026-06-22",
        status: "Demo Submitted",
        amount: 500,
      },
    ],
  });
}
