// backend/server.js
// ─────────────────────────────────────────────────────────────
// Entry point for the BA Productivity Hub backend.
// Run:  node server.js   or   npm run dev  (with nodemon)
// ─────────────────────────────────────────────────────────────

import "dotenv/config";
import express from "express";
import cors from "cors";
import path from "path";
import { fileURLToPath } from "url";
import { config } from "./config/index.js";

import claudeRouter      from "./routes/claude.js";
import reportRouter      from "./routes/report.js";
import momRouter         from "./routes/mom.js";
import standupRouter     from "./routes/standup.js";
import sharepointRouter  from "./routes/sharepoint.js";
import teamsRouter       from "./routes/teams.js";
import outlookRouter     from "./routes/outlook.js";
import msCalendarRouter  from "./routes/mscalendar.js";
import ewsRouter         from "./routes/ews.js";
import documentsRouter   from "./routes/documents.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app = express();
const PORT = process.env.PORT || 3000;

// ── Middleware ─────────────────────────────────────────────────
app.use(cors({ origin: "*" }));
app.use(express.json({ limit: "10mb" }));
app.use(express.urlencoded({ extended: true }));

// ── Serve frontend static files ────────────────────────────────
// In production the backend serves the frontend build.
// In development the frontend runs on its own port (5173).
const frontendDist = path.join(__dirname, "../frontend");
app.use(express.static(frontendDist));

// ── API Routes ─────────────────────────────────────────────────
app.use("/api/claude",      claudeRouter);
app.use("/api/report",      reportRouter);
app.use("/api/mom",         momRouter);
app.use("/api/standup",     standupRouter);
app.use("/api/sharepoint",  sharepointRouter);
app.use("/api/teams",       teamsRouter);
app.use("/api/outlook",     outlookRouter);
app.use("/api/mscalendar",  msCalendarRouter);
app.use("/api/ews",         ewsRouter);
app.use("/api/documents",  documentsRouter);

// ── Microsoft OAuth login redirect ─────────────────────────────
app.get("/api/auth/microsoft", (req, res) => {
  const { tenantId, clientId } = config.sharepoint;
  if (!tenantId || !clientId)
    return res.status(500).json({ error: "SP_TENANT_ID and SP_CLIENT_ID must be set in .env" });
  const origin      = process.env.VERCEL_URL ? `https://${process.env.VERCEL_URL}` : `http://localhost:${config.port}`;
  const redirectUri = `${origin}/api/sharepoint/callback`;
  const scope       = "Files.ReadWrite.All Sites.ReadWrite.All User.Read offline_access Calendars.Read Mail.Read Mail.Send Chat.Read ChannelMessage.Read.All";
  res.redirect(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?client_id=${clientId}&response_type=code&redirect_uri=${encodeURIComponent(redirectUri)}&scope=${encodeURIComponent(scope)}&response_mode=query`);
});

// ── Health check ───────────────────────────────────────────────
app.get("/api/health", (_req, res) =>
  res.json({ status: "ok", time: new Date().toISOString() })
);

// ── Catch-all: serve frontend index.html ──────────────────────
app.get("*", (_req, res) => {
  res.sendFile(path.join(frontendDist, "index.html"));
});

// ── Global error guards ────────────────────────────────────────
process.on("unhandledRejection", (reason) => {
  console.error("[unhandledRejection]", reason);
});
process.on("uncaughtException", (err) => {
  console.error("[uncaughtException]", err);
});

// ── Start ──────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\n✅  BA Hub backend running at: http://localhost:${PORT}`);
  console.log(`   Frontend served at:          http://localhost:${PORT}`);
  console.log(`   API health:                  http://localhost:${PORT}/api/health\n`);
});
