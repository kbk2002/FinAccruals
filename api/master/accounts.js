export default function handler(req, res) {
  if (req.method === "GET") {
    res.status(200).json({
      success: true,
      data: []   // empty placeholder
    })
  } else {
    res.status(405).json({ error: "Method not allowed" })
  }
}
