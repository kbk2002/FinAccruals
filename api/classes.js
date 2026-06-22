export default function handler(req, res) {
  res.status(200).json({
    success: true,
    classes: [
      { name: "Consulting" },
      { name: "Retail" },
      { name: "Operations" }
    ]
  });
}
