// frontend/src/utils/api.js
// Central API client — all HTTP calls to the backend go through here.

const BASE = window.location.hostname === "localhost"
  ? "http://localhost:3000"   // Dev: backend runs separately
  : "";                        // Prod: same origin (backend serves frontend)

function msToken() {
  return localStorage.getItem("spToken") || "";
}

function isExpiredTokenError(status, msg) {
  return status === 401 || /expired|InvalidAuthenticationToken|token.*invalid/i.test(msg || "");
}

async function get(endpoint, params = {}) {
  const qs = new URLSearchParams(params).toString();
  const url = `${BASE}${endpoint}${qs ? "?" + qs : ""}`;
  try {
    const res = await fetch(url);
    if (!res.ok) {
      const err = await res.json().catch(() => ({ error: `HTTP ${res.status}` }));
      const msg = err.error || `Request failed (${res.status})`;
      if (isExpiredTokenError(res.status, msg)) window.dispatchEvent(new CustomEvent("ms-token-expired"));
      return { error: msg };
    }
    return res.json();
  } catch (e) {
    return { error: `Network error: ${e.message}` };
  }
}

async function post(endpoint, body) {
  try {
    const res = await fetch(`${BASE}${endpoint}`, {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify(body),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({ error: `HTTP ${res.status}` }));
      const msg = err.error || `Request failed (${res.status})`;
      if (isExpiredTokenError(res.status, msg)) window.dispatchEvent(new CustomEvent("ms-token-expired"));
      return { error: msg };
    }
    return res.json();
  } catch (e) {
    return { error: `Network error: ${e.message}. Is the backend running on port 3000?` };
  }
}

export const api = {
  // General assistant chat — pass provider so backend routes to correct AI
  chat: (prompt, system, provider) => post("/api/claude", { prompt, system, provider: provider || localStorage.getItem("ai_provider") || "groq" }),

  // Weekly report generation
  generateReport: (data) => post("/api/report", { ...data, provider: data.provider || localStorage.getItem("ai_provider") || "groq" }),

  // MOM (Minutes of Meeting)
  generateMom: (data) => post("/api/mom", data),

  // Token refresh
  refreshToken: (refreshToken) => post("/api/sharepoint/refresh", { refreshToken }),

  // Daily stand-up
  generateStandup: (data) => post("/api/standup",      data),
  standupQA:       (data) => post("/api/standup/qa",   data),

  // SharePoint / OneDrive
  spTest:   (token)  => post("/api/sharepoint/test",   { token }),
  spExport: (data)   => post("/api/sharepoint/export", data),

  // Microsoft 365 — Teams
  teamsMeetings:     ()         => get("/api/teams/meetings", { token: msToken() }),
  teamsMOM:          (data)     => post("/api/teams/mom", data),
  teamsChats:        ()         => get("/api/teams/chats",    { token: msToken() }),
  teamsChatMessages: (chatId)   => get(`/api/teams/chats/${chatId}/messages`, { token: msToken() }),

  // Microsoft 365 — Outlook
  outlookEmails:    ()    => get("/api/outlook/emails",           { token: msToken() }),
  outlookEmailBody: (id)  => get(`/api/outlook/emails/${id}/body`, { token: msToken() }),
  outlookDraft:     (data) => post("/api/outlook/draft", data),
  outlookSend:      (data) => post("/api/outlook/send", { ...data, token: msToken() }),

  // Microsoft 365 — Calendar
  msCalendar: () => get("/api/mscalendar/events", { token: msToken() }),

  // Exchange Web Services (on-premise Exchange)
  ewsMeetings:     (creds) => post("/api/ews/meetings",        creds),
  ewsEmails:       (creds) => post("/api/ews/emails",           creds),
  ewsEmailBody:    (creds) => post("/api/ews/email-body",       creds),
  ewsSendEmail:    (data)  => post("/api/ews/send-email",       data),
  ewsCreateMeeting:(data)  => post("/api/ews/create-meeting",   data),
  ewsDiscover:     (email) => get("/api/ews/discover",          { email }),

  // Document upload to SharePoint
  uploadDoc: (formData) => {
    return fetch(`${BASE}/api/documents/upload`, { method: "POST", body: formData })
      .then((r) => r.json())
      .catch((e) => ({ error: e.message }));
  },
};
