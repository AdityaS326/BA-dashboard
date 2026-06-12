// backend/controllers/teamsController.js
import { callGroq } from "../utils/groq.js";

function bearer(token) {
  return token.startsWith("Bearer ") ? token : `Bearer ${token}`;
}

// GET /api/teams/meetings
// Try Teams-native /me/onlineMeetings first (cloud, no Exchange dependency),
// fall back to calendarView for Exchange Online accounts.
export async function getMeetings(req, res) {
  const token = req.query.token || req.headers["x-ms-token"] || "";
  if (!token) return res.status(400).json({ error: "No Microsoft token. Connect via Microsoft 365 first." });

  const auth = { Authorization: bearer(token) };

  try {
    // Attempt 1: Teams-native endpoint (works even with on-premise Exchange)
    const teamsResp = await fetch(
      "https://graph.microsoft.com/v1.0/me/onlineMeetings?$top=20",
      { headers: auth }
    );

    if (teamsResp.ok) {
      const data = await teamsResp.json();
      const meetings = (data.value || []).map((m) => ({
        subject:         m.subject || "Teams meeting",
        start:           { dateTime: m.startDateTime },
        end:             { dateTime: m.endDateTime },
        attendees:       (m.participants?.attendees || []).map((a) => ({
          emailAddress:  { name: a.identity?.user?.displayName, address: a.identity?.user?.id }
        })),
        isOnlineMeeting: true,
        onlineMeetingUrl: m.joinWebUrl,
        source:          "teams",
      }));
      return res.json({ meetings, source: "teams" });
    }

    const teamsErr = await teamsResp.json();
    const teamsCode = teamsErr.error?.code || "";

    if (teamsResp.status === 403 || teamsCode === "Forbidden") {
      return res.status(403).json({
        error: "Missing OnlineMeetings.Read permission. In Graph Explorer click your avatar -> Consent to permissions -> add OnlineMeetings.Read, then copy a fresh token.",
        hint: "permission"
      });
    }

    // Attempt 2: Calendar view (Exchange Online only)
    const now   = new Date();
    const start = new Date(now.getTime() - 7  * 24 * 60 * 60 * 1000).toISOString();
    const end   = new Date(now.getTime() + 14 * 24 * 60 * 60 * 1000).toISOString();

    const calResp = await fetch(
      `https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=${start}&endDateTime=${end}&$filter=isOnlineMeeting eq true&$select=subject,start,end,attendees,isOnlineMeeting,onlineMeetingUrl&$orderby=start/dateTime desc&$top=20`,
      { headers: auth }
    );

    if (calResp.ok) {
      const data = await calResp.json();
      return res.json({ meetings: data.value || [], source: "calendar" });
    }

    const calErr = await calResp.json();
    const calMsg = calErr.error?.message || "";

    if (calMsg.toLowerCase().includes("on-premise") || calMsg.toLowerCase().includes("inactive") || calMsg.toLowerCase().includes("soft-deleted")) {
      return res.status(200).json({
        meetings: [],
        onPremise: true,
        error: "Your mailbox is hosted on-premise (not Exchange Online). Calendar sync is unavailable, but you can still use the MOM Generator manually."
      });
    }

    return res.status(calResp.status).json({ error: calMsg || "Graph API error" });

  } catch (err) {
    res.status(502).json({ error: err.message });
  }
}

// GET /api/teams/chats
export async function getChats(req, res) {
  const token = req.query.token || req.headers["x-ms-token"] || "";
  if (!token) return res.status(400).json({ error: "No Microsoft 365 token. Connect via Graph Explorer first." });
  const auth = { Authorization: bearer(token) };
  try {
    const resp = await fetch(
      "https://graph.microsoft.com/v1.0/me/chats?$expand=members&$top=30&$orderby=lastUpdatedDateTime desc",
      { headers: auth }
    );
    if (resp.ok) {
      const data = await resp.json();
      return res.json({ chats: data.value || [] });
    }
    const err = await resp.json();
    const msg = err.error?.message || "Graph API error";
    return res.status(resp.status).json({ error: msg });
  } catch (err) {
    res.status(502).json({ error: err.message });
  }
}

// GET /api/teams/chats/:chatId/messages
export async function getChatMessages(req, res) {
  const token  = req.query.token || req.headers["x-ms-token"] || "";
  const chatId = req.params.chatId;
  if (!token)  return res.status(400).json({ error: "No Microsoft 365 token." });
  if (!chatId) return res.status(400).json({ error: "Missing chatId." });
  const auth = { Authorization: bearer(token) };
  try {
    const resp = await fetch(
      `https://graph.microsoft.com/v1.0/me/chats/${chatId}/messages?$top=40`,
      { headers: auth }
    );
    if (resp.ok) {
      const data = await resp.json();
      return res.json({ messages: data.value || [] });
    }
    const err = await resp.json();
    return res.status(resp.status).json({ error: err.error?.message || "Graph API error" });
  } catch (err) {
    res.status(502).json({ error: err.message });
  }
}

// POST /api/teams/mom
export async function generateTeamsMOM(req, res) {
  const { subject, date, attendees, duration, context, transcript } = req.body;

  const hasTranscript = transcript && transcript.trim().length > 20;

  const prompt = hasTranscript
    ? `Generate a formal Minutes of Meeting (MOM) STRICTLY from the transcript below. Do NOT add anything not in the transcript.

Meeting  : ${subject}
Date     : ${date}
Attendees: ${attendees}
Duration : ${duration || "N/A"}

TRANSCRIPT:
${transcript}

Using ONLY what is in the transcript above, write the MOM:
1. Meeting header (title, date, attendees, duration)
2. Discussion points — closely based on what was discussed in the transcript
3. Decisions made — only decisions explicitly mentioned
4. Action items table (Action | Owner | Due Date | Priority) — only items explicitly assigned
5. Next steps — only if mentioned
6. Sign-off section`
    : `Generate a formal Minutes of Meeting for the meeting below. No transcript is available — mark all inferred content as [To be confirmed].

Meeting  : ${subject}
Date     : ${date}
Attendees: ${attendees}
Duration : ${duration || "N/A"}
Context  : ${context || "Regular team meeting"}

Include: header, discussion points [To be confirmed], decisions [To be confirmed], action items, next steps, sign-off.`;

  try {
    const text = await callGroq(
      prompt,
      "You are a professional business analyst writing formal Minutes of Meeting documents.",
      1200
    );
    res.json({ text });
  } catch (err) {
    console.error("[teamsController]", err.message);
    res.status(502).json({ error: err.message });
  }
}
