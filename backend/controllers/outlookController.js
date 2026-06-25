// backend/controllers/outlookController.js
import { callGroq } from "../utils/groq.js";

function bearer(token) {
  return token.startsWith("Bearer ") ? token : `Bearer ${token}`;
}

function isOnPremiseError(msg, code) {
  const m = (msg || "").toLowerCase();
  return m.includes("on-premise") || m.includes("inactive") || m.includes("soft-deleted") || code === "MailboxNotEnabledForRESTAPI";
}

// GET /api/outlook/emails/:id/body
export async function getEmailBody(req, res) {
  const token = req.query.token || req.headers["x-ms-token"] || "";
  const { id } = req.params;
  if (!token) return res.status(400).json({ error: "No Microsoft token." });
  if (!id)    return res.status(400).json({ error: "Missing email id." });

  const auth = { Authorization: bearer(token) };

  try {
    // Fetch body and inline attachments in parallel
    const [bodyResp, attResp] = await Promise.all([
      fetch(`https://graph.microsoft.com/v1.0/me/messages/${id}?$select=body`, { headers: auth }),
      fetch(`https://graph.microsoft.com/v1.0/me/messages/${id}/attachments`, { headers: auth }),
    ]);

    if (!bodyResp.ok) {
      const err = await bodyResp.json();
      return res.status(bodyResp.status).json({ error: err.error?.message || "Graph API error" });
    }

    const bodyData  = await bodyResp.json();
    const bodyType  = bodyData.body?.contentType || "text";
    let   html      = bodyData.body?.content || "";

    // Replace cid: inline image references with base64 data URLs
    if (bodyType === "html" && attResp.ok) {
      const attData    = await attResp.json();
      const attachments = (attData.value || []).filter(a => a.isInline && a["@odata.type"] === "#microsoft.graph.fileAttachment");
      for (const att of attachments) {
        if (!att.contentId || !att.contentBytes) continue;
        const cid     = att.contentId.replace(/^<|>$/g, "");
        const dataUrl = `data:${att.contentType};base64,${att.contentBytes}`;
        html = html.replace(new RegExp(`cid:${cid.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}`, "gi"), dataUrl);
      }
    }

    res.json({ body: html, bodyType });
  } catch (err) {
    res.status(502).json({ error: err.message });
  }
}

// GET /api/outlook/emails
export async function getEmails(req, res) {
  const token = req.query.token || req.headers["x-ms-token"] || "";
  if (!token) return res.status(400).json({ error: "No Microsoft token. Connect via Microsoft 365 first." });

  try {
    const resp = await fetch(
      "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=20&$orderby=receivedDateTime desc&$select=id,subject,from,receivedDateTime,isRead,bodyPreview",
      { headers: { Authorization: bearer(token) } }
    );

    if (!resp.ok) {
      const err  = await resp.json();
      const msg  = err.error?.message || "";
      const code = err.error?.code    || "";

      if (isOnPremiseError(msg, code)) {
        return res.status(200).json({
          emails:    [],
          onPremise: true,
          error:     "Your Outlook mailbox is on-premise (not Exchange Online). Microsoft Graph cannot access on-premise mailboxes directly. Ask your IT admin whether Outlook Web Access (OWA) is available, or use the AI draft reply feature manually."
        });
      }
      return res.status(resp.status).json({ error: msg || "Graph API error" });
    }

    const data = await resp.json();
    res.json({ emails: data.value || [] });
  } catch (err) {
    res.status(502).json({ error: err.message });
  }
}

// POST /api/outlook/draft
export async function draftReply(req, res) {
  const { subject, from, bodyPreview } = req.body;
  if (!subject) return res.status(400).json({ error: "subject is required" });

  const prompt = `Draft a professional reply to this email:

From    : ${from}
Subject : ${subject}
Preview : ${bodyPreview || "(no preview)"}

Write a concise, courteous reply. Sign off as:
Aditya S | System Analyst / Solution Architect | ESDS Software Solution Pvt. Ltd.`;

  try {
    const text = await callGroq(
      prompt,
      "You write concise, professional workplace email replies in formal English.",
      600
    );
    res.json({ text });
  } catch (err) {
    console.error("[outlookController]", err.message);
    res.status(502).json({ error: err.message });
  }
}

// POST /api/outlook/send
export async function sendEmail(req, res) {
  const { token, to, subject, body } = req.body;
  if (!token)               return res.status(400).json({ error: "No Microsoft token." });
  if (!to || !subject || !body) return res.status(400).json({ error: "to, subject, body are required." });

  try {
    const resp = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
      method: "POST",
      headers: { Authorization: bearer(token), "Content-Type": "application/json" },
      body: JSON.stringify({
        message: {
          subject,
          body: { contentType: "Text", content: body },
          toRecipients: [{ emailAddress: { address: to } }],
        },
      }),
    });

    if (!resp.ok) {
      const err = await resp.json();
      return res.status(resp.status).json({ error: err.error?.message });
    }
    res.json({ ok: true });
  } catch (err) {
    res.status(502).json({ error: err.message });
  }
}
