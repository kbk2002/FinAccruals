const http = require("http");
const url = require("url");

// Import all handlers
const health = require("./health");
const accounts = require("./accounts");
const vendors = require("./vendors");
const classes = require("./classes");
const history = require("./history");

// Helper to send JSON
function sendJSON(res, statusCode, data) {
  res.writeHead(statusCode, {
    "Content-Type": "application/json",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET,OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  });
  res.end(JSON.stringify(data));
}

// Wrapper to adapt your handler(req, res) to pure Node.js
function runHandler(handler, req, res) {
  return handler(
    {
      method: req.method,
      url: req.url,
      query: url.parse(req.url, true).query
    },
    {
      status: (code) => ({
        json: (data) => sendJSON(res, code, data)
      })
    }
  );
}

const server = http.createServer((req, res) => {
  const parsed = url.parse(req.url, true);
  const path = parsed.pathname;

  if (req.method === "OPTIONS") {
    res.writeHead(204, {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET,OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type"
    });
    return res.end();
  }

  // Routing
  if (req.method === "GET" && path === "/api/health") {
    return runHandler(health, req, res);
  }

  if (req.method === "GET" && path === "/api/accounts") {
    return runHandler(accounts, req, res);
  }

  if (req.method === "GET" && path === "/api/vendors") {
    return runHandler(vendors, req, res);
  }

  if (req.method === "GET" && path === "/api/classes") {
    return runHandler(classes, req, res);
  }

  if (req.method === "GET" && path === "/api/history") {
    return runHandler(history, req, res);
  }

  // Default 404
  sendJSON(res, 404, { error: "Route not found", path });
});

server.listen(3001, () => {
  console.log("FinAccruals API running on http://localhost:3001");
});
