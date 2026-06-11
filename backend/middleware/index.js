// backend/middleware/index.js

import { config } from "../config/index.js";

// ── Guard: reject if GROQ_API_KEY is missing ──────────────────
export function requireGroqKey(req, res, next) {
  if (!config.groqApiKey) {
    return res.status(500).json({
      error: "GROQ_API_KEY is not set. Add it to your .env file.",
    });
  }
  next();
}

// ── Global error handler (register last in server.js) ─────────
export function errorHandler(err, req, res, _next) {
  console.error("Unhandled error:", err.message);
  res.status(500).json({ error: err.message || "Internal server error" });
}
