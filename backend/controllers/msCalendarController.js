// backend/controllers/msCalendarController.js
// Microsoft 365 calendar events via Graph API.

function bearer(token) {
  return token.startsWith("Bearer ") ? token : `Bearer ${token}`;
}

// GET /api/mscalendar/events â€" calendar events for next 30 days
export async function getEvents(req, res) {
  const token = req.query.token || req.headers["x-ms-token"] || "";
  if (!token) return res.status(400).json({ error: "No Microsoft token. Connect via Microsoft 365 first." });

  const now   = new Date();
  const start = now.toISOString();
  const end   = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000).toISOString();

  try {
    const resp = await fetch(
      `https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=${start}&endDateTime=${end}&$select=subject,start,end,location,attendees,isOnlineMeeting,bodyPreview&$orderby=start/dateTime&$top=50`,
      { headers: { Authorization: bearer(token) } }
    );
    if (!resp.ok) {
      const err = await resp.json();
      return res.status(resp.status).json({ error: err.error?.message || "Graph API error" });
    }
    const data = await resp.json();
    res.json({ events: data.value || [] });
  } catch (err) {
    res.status(502).json({ error: err.message });
  }
}

