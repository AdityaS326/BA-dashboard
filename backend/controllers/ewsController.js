// backend/controllers/ewsController.js
// Uses httpntlm to handle the 3-way NTLM handshake that ESDS Exchange requires.
// Also accepts self-signed certs and TLS 1.0 (common on on-premise Exchange).

import httpntlm from "httpntlm";

const EWS_URL = "https://owa.esds.co.in/EWS/Exchange.asmx";

function buildSOAP(startDate, endDate) {
  return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013_SP1"/>
  </soap:Header>
  <soap:Body>
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject"/>
          <t:FieldURI FieldURI="calendar:Start"/>
          <t:FieldURI FieldURI="calendar:End"/>
          <t:FieldURI FieldURI="calendar:Location"/>
          <t:FieldURI FieldURI="calendar:IsOnlineMeeting"/>
          <t:FieldURI FieldURI="calendar:OnlineMeetingInternalLinks"/>
          <t:FieldURI FieldURI="calendar:RequiredAttendees"/>
          <t:FieldURI FieldURI="calendar:OptionalAttendees"/>
          <t:FieldURI FieldURI="calendar:MyResponseType"/>
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:CalendarView MaxReturnsTotal="50"
        StartDate="${startDate}"
        EndDate="${endDate}"/>
      <m:ParentFolderIds>
        <t:DistinguishedFolderId Id="calendar"/>
      </m:ParentFolderIds>
    </m:FindItem>
  </soap:Body>
</soap:Envelope>`;
}

function parseXML(xml) {
  const meetings = [];
  const blocks   = xml.match(/<t:CalendarItem>[\s\S]*?<\/t:CalendarItem>/g) || [];

  for (const block of blocks) {
    const tag = (name) => {
      const m = block.match(new RegExp(`<t:${name}[^>]*>([\\s\\S]*?)<\\/t:${name}>`, "i"));
      return m ? m[1].trim() : "";
    };
    const allTags = (name) => {
      const re  = new RegExp(`<t:${name}[^>]*>([\\s\\S]*?)<\\/t:${name}>`, "gi");
      const out = [];
      let m;
      while ((m = re.exec(block)) !== null) out.push(m[1].trim());
      return out;
    };

    const subject  = tag("Subject");
    const start    = tag("Start");
    const end      = tag("End");
    const location = tag("Location");
    const isOnline = tag("IsOnlineMeeting") === "true";
    const response = tag("MyResponseType");

    if (!subject || !start) continue;

    // Parse Teams join URL from OnlineMeetingInternalLinks (Exchange 2013+)
    const linksBlock = block.match(/<t:OnlineMeetingInternalLinks[\s\S]*?<\/t:OnlineMeetingInternalLinks>/i);
    const joinUrlFromLinks = linksBlock
      ? (linksBlock[0].match(/<t:Url[^>]*>([\s\S]*?)<\/t:Url>/i) || [])[1] || ""
      : "";
    // Fallback: Teams URL embedded in the Location field
    const joinUrlFromLocation = (location.match(/https:\/\/teams\.microsoft\.com\/[^\s<>"]+/) || [])[0] || "";
    const joinUrl = joinUrlFromLinks || joinUrlFromLocation;

    const attendeeNames = allTags("Name")
      .filter((n) => n && !n.includes("<"))
      .slice(0, 6);

    const startDt = new Date(start);
    const endDt   = new Date(end);
    const dur     = isNaN(startDt) || isNaN(endDt)
      ? "N/A"
      : `${Math.round((endDt - startDt) / 60000)} min`;

    meetings.push({ subject, start, end, dur, location, isOnline, joinUrl, attendees: attendeeNames, response });
  }

  return meetings.sort((a, b) => new Date(a.start) - new Date(b.start));
}

// Promisify httpntlm.post
function ntlmPost(opts) {
  return new Promise((resolve, reject) => {
    httpntlm.post(opts, (err, res) => {
      if (err) return reject(err);
      resolve(res);
    });
  });
}

// POST /api/ews/meetings
export async function getMeetings(req, res) {
  const { ewsUrl, username, password, days = 30 } = req.body;

  if (!username || !password) {
    return res.status(400).json({ error: "username and password are required." });
  }

  const url = ewsUrl || EWS_URL;

  // Split domain\username or use email as-is
  let domain   = "";
  let user     = username;
  if (username.includes("\\")) {
    [domain, user] = username.split("\\");
  }

  const now      = new Date();
  const start    = new Date(now.getTime() - 7  * 24 * 60 * 60 * 1000);
  const end      = new Date(now.getTime() + days * 24 * 60 * 60 * 1000);
  const startISO = start.toISOString().split(".")[0];
  const endISO   = end.toISOString().split(".")[0];
  const body     = buildSOAP(startISO, endISO);

  try {
    const response = await ntlmPost({
      url,
      username:  user,
      password,
      domain,
      workstation: "",
      body,
      headers: {
        "Content-Type": "text/xml; charset=utf-8",
        "SOAPAction":   '"http://schemas.microsoft.com/exchange/services/2006/messages/FindItem"',
      },
      // Accept self-signed certs and old TLS versions common on corporate Exchange
      rejectUnauthorized: false,
    });

    if (response.statusCode === 401) {
      return res.status(401).json({ error: "Invalid credentials — wrong username or password." });
    }
    if (response.statusCode === 403) {
      return res.status(403).json({ error: "Access denied (403). Your account may not have EWS access enabled." });
    }
    if (response.statusCode >= 400) {
      return res.status(response.statusCode).json({ error: `Exchange returned HTTP ${response.statusCode}. Check the EWS URL.` });
    }

    const xml = response.body;

    // SOAP fault check
    if (xml.includes("<faultstring>") || xml.includes("soap:Fault")) {
      const fault = xml.match(/<faultstring[^>]*>([\s\S]*?)<\/faultstring>/i);
      return res.status(500).json({ error: "Exchange SOAP error: " + (fault ? fault[1] : "Unknown fault") });
    }

    const meetings = parseXML(xml);
    res.json({ meetings, total: meetings.length, url });

  } catch (err) {
    const msg = err.message || "";
    if (msg.includes("ENOTFOUND") || msg.includes("getaddrinfo")) {
      return res.status(502).json({ error: `Cannot find Exchange server. Are you on the ESDS network or VPN?` });
    }
    if (msg.includes("ECONNREFUSED")) {
      return res.status(502).json({ error: "Connection refused by Exchange server. EWS may be disabled." });
    }
    if (msg.includes("ETIMEDOUT") || msg.includes("timed out")) {
      return res.status(504).json({ error: "Connection timed out. Connect to ESDS office WiFi or VPN first." });
    }
    res.status(502).json({ error: msg });
  }
}

// ── EWS Email helpers ────────────────────────────────────────────────────────

function buildEmailSOAP(maxRows = 25) {
  return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header><t:RequestServerVersion Version="Exchange2013_SP1"/></soap:Header>
  <soap:Body>
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject"/>
          <t:FieldURI FieldURI="message:From"/>
          <t:FieldURI FieldURI="item:DateTimeReceived"/>
          <t:FieldURI FieldURI="message:IsRead"/>
          <t:FieldURI FieldURI="message:ToRecipients"/>
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:IndexedPageItemView MaxEntriesReturned="${maxRows}" Offset="0" BasePoint="Beginning"/>
      <m:SortOrder>
        <t:FieldOrder Order="Descending">
          <t:FieldURI FieldURI="item:DateTimeReceived"/>
        </t:FieldOrder>
      </m:SortOrder>
      <m:ParentFolderIds>
        <t:DistinguishedFolderId Id="inbox"/>
      </m:ParentFolderIds>
    </m:FindItem>
  </soap:Body>
</soap:Envelope>`;
}

function buildGetItemSOAP(itemId, changeKey) {
  return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header><t:RequestServerVersion Version="Exchange2013_SP1"/></soap:Header>
  <soap:Body>
    <m:GetItem>
      <m:ItemShape>
        <t:BaseShape>Default</t:BaseShape>
        <t:BodyType>Text</t:BodyType>
      </m:ItemShape>
      <m:ItemIds>
        <t:ItemId Id="${itemId}" ChangeKey="${changeKey}"/>
      </m:ItemIds>
    </m:GetItem>
  </soap:Body>
</soap:Envelope>`;
}

function parseEmails(xml) {
  const emails = [];
  const blocks = xml.match(/<t:Message>[\s\S]*?<\/t:Message>/g) || [];
  for (const block of blocks) {
    const tag = (name) => {
      const m = block.match(new RegExp(`<t:${name}[^>]*>([\\s\\S]*?)<\\/t:${name}>`, "i"));
      return m ? m[1].trim() : "";
    };
    const idMatch  = block.match(/<t:ItemId Id="([^"]+)" ChangeKey="([^"]+)"/);
    emails.push({
      id:        idMatch ? idMatch[1] : "",
      changeKey: idMatch ? idMatch[2] : "",
      subject:   tag("Subject") || "(no subject)",
      from:      { name: tag("Name"), address: tag("EmailAddress") },
      received:  tag("DateTimeReceived"),
      isRead:    tag("IsRead") === "true",
    });
  }
  return emails;
}

// POST /api/ews/emails
export async function getEmails(req, res) {
  const { ewsUrl, username, password } = req.body;
  if (!username || !password) return res.status(400).json({ error: "username and password are required." });

  const url  = ewsUrl || EWS_URL;
  let domain = "", user = username;
  if (username.includes("\\")) [domain, user] = username.split("\\");

  try {
    const response = await ntlmPost({
      url, username: user, password, domain, workstation: "",
      body: buildEmailSOAP(25),
      headers: { "Content-Type": "text/xml; charset=utf-8", "SOAPAction": '"http://schemas.microsoft.com/exchange/services/2006/messages/FindItem"' },
      rejectUnauthorized: false,
    });
    if (response.statusCode === 401) return res.status(401).json({ error: "Invalid credentials." });
    if (response.statusCode >= 400) return res.status(response.statusCode).json({ error: `Exchange returned HTTP ${response.statusCode}` });
    const emails = parseEmails(response.body);
    res.json({ emails, total: emails.length });
  } catch (err) {
    const msg = err.message || "";
    if (msg.includes("ENOTFOUND") || msg.includes("ETIMEDOUT")) return res.status(502).json({ error: "Cannot reach Exchange. Are you on the ESDS network?" });
    res.status(502).json({ error: msg });
  }
}

// POST /api/ews/email-body
export async function getEmailBody(req, res) {
  const { ewsUrl, username, password, itemId, changeKey } = req.body;
  if (!username || !password || !itemId) return res.status(400).json({ error: "username, password and itemId required." });

  const url  = ewsUrl || EWS_URL;
  let domain = "", user = username;
  if (username.includes("\\")) [domain, user] = username.split("\\");

  try {
    const response = await ntlmPost({
      url, username: user, password, domain, workstation: "",
      body: buildGetItemSOAP(itemId, changeKey),
      headers: { "Content-Type": "text/xml; charset=utf-8", "SOAPAction": '"http://schemas.microsoft.com/exchange/services/2006/messages/GetItem"' },
      rejectUnauthorized: false,
    });
    const bodyMatch = response.body.match(/<t:Body[^>]*>([\s\S]*?)<\/t:Body>/i);
    const raw = bodyMatch ? bodyMatch[1] : "";
    const bodyText = raw
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/p>/gi, "\n")
      .replace(/<\/div>/gi, "\n")
      .replace(/<[^>]+>/g, "")
      .replace(/&#xD;/gi, "")
      .replace(/&#xA;/gi, "\n")
      .replace(/&nbsp;/gi, " ")
      .replace(/&amp;/gi, "&")
      .replace(/&lt;/gi, "<")
      .replace(/&gt;/gi, ">")
      .replace(/&quot;/gi, '"')
      .replace(/&#x27;/gi, "'")
      .replace(/\r/g, "")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
    res.json({ body: bodyText.slice(0, 3000) });
  } catch (err) {
    res.status(502).json({ error: err.message });
  }
}

// POST /api/ews/send-email
function buildSendEmailSOAP(to, cc, subject, body) {
  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const toXml = to.split(",").map((a) => `<t:Mailbox><t:EmailAddress>${esc(a.trim())}</t:EmailAddress></t:Mailbox>`).join("");
  const ccXml = cc
    ? cc.split(",").filter(Boolean).map((a) => `<t:Mailbox><t:EmailAddress>${esc(a.trim())}</t:EmailAddress></t:Mailbox>`).join("")
    : "";
  return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header><t:RequestServerVersion Version="Exchange2013_SP1"/></soap:Header>
  <soap:Body>
    <m:CreateItem MessageDisposition="SendAndSaveCopy">
      <m:SavedItemFolderId><t:DistinguishedFolderId Id="sentitems"/></m:SavedItemFolderId>
      <m:Items>
        <t:Message>
          <t:Subject>${esc(subject)}</t:Subject>
          <t:Body BodyType="Text">${esc(body)}</t:Body>
          <t:ToRecipients>${toXml}</t:ToRecipients>
          ${ccXml ? `<t:CcRecipients>${ccXml}</t:CcRecipients>` : ""}
        </t:Message>
      </m:Items>
    </m:CreateItem>
  </soap:Body>
</soap:Envelope>`;
}

export async function sendEmail(req, res) {
  const { ewsUrl, username, password, to, cc, subject, body } = req.body;
  if (!username || !password) return res.status(400).json({ error: "username and password are required." });
  if (!to)       return res.status(400).json({ error: "Recipient (to) is required." });
  if (!subject)  return res.status(400).json({ error: "Subject is required." });

  const url  = ewsUrl || EWS_URL;
  let domain = "", user = username;
  if (username.includes("\\")) [domain, user] = username.split("\\");

  try {
    const response = await ntlmPost({
      url, username: user, password, domain, workstation: "",
      body: buildSendEmailSOAP(to, cc || "", subject, body || ""),
      headers: { "Content-Type": "text/xml; charset=utf-8", "SOAPAction": '"http://schemas.microsoft.com/exchange/services/2006/messages/CreateItem"' },
      rejectUnauthorized: false,
    });
    if (response.statusCode === 401) return res.status(401).json({ error: "Invalid Exchange credentials." });
    if (response.statusCode >= 400)  return res.status(response.statusCode).json({ error: `Exchange returned HTTP ${response.statusCode}` });
    if (response.body.includes("soap:Fault")) {
      const fault = response.body.match(/<faultstring[^>]*>([\s\S]*?)<\/faultstring>/i);
      return res.status(500).json({ error: "Exchange SOAP error: " + (fault ? fault[1] : "Unknown") });
    }
    res.json({ ok: true });
  } catch (err) {
    const msg = err.message || "";
    if (msg.includes("ENOTFOUND") || msg.includes("ETIMEDOUT")) return res.status(502).json({ error: "Cannot reach Exchange. Are you on the ESDS network?" });
    res.status(502).json({ error: msg });
  }
}

// GET /api/ews/discover
export async function discoverEWS(req, res) {
  const { email } = req.query;
  if (!email || !email.includes("@")) {
    return res.status(400).json({ error: "email is required" });
  }
  const domain = email.split("@")[1];
  // owa.esds.co.in confirmed reachable (HTTP 401 = server exists, auth required)
  const candidates = [
    `https://owa.${domain}/EWS/Exchange.asmx`,
    `https://mail.${domain}/EWS/Exchange.asmx`,
    `https://webmail.${domain}/EWS/Exchange.asmx`,
    `https://exchange.${domain}/EWS/Exchange.asmx`,
    `https://outlook.${domain}/EWS/Exchange.asmx`,
  ];
  res.json({ candidates, domain, recommended: `https://owa.${domain}/EWS/Exchange.asmx` });
}
