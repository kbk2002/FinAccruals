function safeSerialize(value) {
  try {
    const serialized = JSON.stringify(value);
    return serialized.length > 2000 ? `${serialized.slice(0, 2000)}...` : JSON.parse(serialized);
  } catch {
    return String(value);
  }
}

export function traceApiCall(handler) {
  return async function tracedHandler(req, res) {
    const requestId = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    const startedAt = Date.now();

    res.setHeader("X-Request-Id", requestId);
    res.on("finish", () => {
      const durationMs = Date.now() - startedAt;
      console.log(
        JSON.stringify({
          event: "api-response",
          requestId,
          timestamp: new Date().toISOString(),
          method: req.method,
          path: req.url || req.originalUrl || "",
          statusCode: res.statusCode,
          durationMs,
        })
      );
    });

    console.log(
      JSON.stringify({
        event: "api-request",
        requestId,
        timestamp: new Date().toISOString(),
        method: req.method,
        path: req.url || req.originalUrl || "",
        query: req.query,
        body: safeSerialize(req.body),
        headers: {
          host: req.headers?.host,
          "user-agent": req.headers?.["user-agent"],
        },
      })
    );

    try {
      await handler(req, res);
    } catch (error) {
      console.error(
        JSON.stringify({
          event: "api-error",
          requestId,
          timestamp: new Date().toISOString(),
          method: req.method,
          path: req.url || req.originalUrl || "",
          error: error?.message || String(error),
          stack: error?.stack,
        })
      );
      throw error;
    }
  };
}
