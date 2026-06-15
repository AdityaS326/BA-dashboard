// backend/controllers/ewsController.js
// Uses httpntlm to handle the 3-way NTLM handshake required by on-premise Exchange.
// Also accepts self-signed certs and TLS 1.0 (common on on-premise Exchange).

import httpntlm from "httpntlm";

function buildSOAP(startDate, endDate) {
  return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2010_SP2"/>
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
          <t:FieldURI FieldURI="calendar:RequiredAttendees"/>
          <t:FieldURI FieldURI="calendar:OptionalAttendees"/>
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

    if (!subject || !start) continue;

    // Detect Teams meeting from URL in Location field
    const joinUrl = (location.match(/https:\/\/teams\.microsoft\.com\/[^\s<>"]+/) || [])[0] || "";
    const isOnline = joinUrl.length > 0;

    const attendeeNames = allTags("Name")
      .filter((n) => n && !n.includes("<"))
      .slice(0, 6);

    const startDt = new Date(start);
    const endDt   = new Date(end);
    const dur     = isNaN(startDt) || isNaN(endDt)
      ? "N/A"
      : `${Math.round((endDt - startDt) / 60000)} min`;

    meetings.push({ subject, start, end, dur, location, isOnline, joinUrl, attendees: attendeeNames });
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

  if (!ewsUrl) return res.status(400).json({ error: "Exchange server URL is required." });
  const url = ewsUrl;

  // Split domain\username or use email as-is
  let domain   = "";
  let user     = username;
  if (username.includes("\\")) {
    [domain, user] = username.split("\\");
  }

  const now      = new Date();
  const start    = new Date(now.getTime() - 7  * 24 * 60 * 60 * 1000);
  const end      = new Date(now.getTime() + days * 24 * 60 * 60 * 1000);
  const startISO = start.toISOString().replace(/\.\d{3}Z$/, "Z");
  const endISO   = end.toISOString().replace(/\.\d{3}Z$/, "Z");
  const body     = buildSOAP(startISO, endISO);

const soapBody = buildCreateMeetingSOAP(
    subject,
    toISO(startDate),
    toISO(endDate),
    attendees || [],
    body,
    location,
    startDate.toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' }),
    endDate.toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })
  );

  console.log('[createMeeting] Sending to', ewsUrl, '| subject:', subject, '| attendees:', (attendees || []).join(', ') || '(none)');

  try {
    const response = await ntlmPost({
      url: ewsUrl, username: user, password, domain, workstation: '',
      body: soapBody,
      headers: { 'Content-Type': 'text/xml; charset=utf-8', 'SOAPAction': '"http://schemas.microsoft.com/exchange/services/2006/messages/CreateItem"' },
      rejectUnauthorized: false,
    });

    console.log('[createMeeting] Exchange HTTP status:', response.statusCode);

    if (response.statusCode === 401) {
      console.log('[createMeeting] Auth failed - invalid credentials');
      return res.status(401).json({ error: 'Invalid Exchange credentials. Check your username and password.' });
    }
    if (response.statusCode >= 400) {
      console.log('[createMeeting] Exchange error body:', response.body?.substring(0, 500));
      return res.status(response.statusCode).json({ error: 'Exchange returned HTTP ' + response.statusCode });
    }
    if (response.body.includes('soap:Fault')) {
      const fault = response.body.match(/<faultstring[^>]*>([\s\S]*?)<\/faultstring>/i);
      const detail = fault ? fault[1] : 'Unknown SOAP fault';
      console.log('[createMeeting] SOAP fault:', detail);
      return res.status(500).json({ error: 'Exchange error: ' + detail });
    }
    if (response.body.includes('ResponseCode')) {
      const code = response.body.match(/<m:ResponseCode>([^<]+)<\/m:ResponseCode>/i);
      if (code && code[1] !== 'NoError') {
        const msg2 = response.body.match(/<m:MessageText>([^<]+)<\/m:MessageText>/i);
        const detail2 = msg2 ? msg2[1] : code[1];
        console.log('[createMeeting] EWS ResponseCode error:', detail2);
        return res.status(500).json({ error: 'Exchange: ' + detail2 });
      }
    }

    console.log('[createMeeting] Success - invite sent to', (attendees || []).join(', ') || 'organizer only');
    const sentTo = (attendees || []).filter(Boolean);
    res.json({ ok: true, start: toISO(startDate), end: toISO(endDate), sentTo, subject });
  } catch (err) {
    const msg = err.message || '';
    console.log('[createMeeting] Exception:', msg);
    if (msg.includes('ENOTFOUND') || msg.includes('ETIMEDOUT') || msg.includes('ECONNREFUSED'))
      return res.status(502).json({ error: 'Cannot reach Exchange server at ' + ewsUrl + '. Check the URL.' });
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
        <t:BodyType>HTML</t:BodyType>
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

  if (!ewsUrl) return res.status(400).json({ error: "Exchange server URL is required." });
  const url  = ewsUrl;
  let domain = "", user = username;
  if (username.includes("\\")) [domain, user] = username.split("\\");

const soapBody = buildCreateMeetingSOAP(
    subject,
    toISO(startDate),
    toISO(endDate),
    attendees || [],
    body,
    location,
    startDate.toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' }),
    endDate.toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })
  );

  console.log('[createMeeting] Sending to', ewsUrl, '| subject:', subject, '| attendees:', (attendees || []).join(', ') || '(none)');

  try {
    const response = await ntlmPost({
      url: ewsUrl, username: user, password, domain, workstation: '',
      body: soapBody,
      headers: { 'Content-Type': 'text/xml; charset=utf-8', 'SOAPAction': '"http://schemas.microsoft.com/exchange/services/2006/messages/CreateItem"' },
      rejectUnauthorized: false,
    });

    console.log('[createMeeting] Exchange HTTP status:', response.statusCode);

    if (response.statusCode === 401) {
      console.log('[createMeeting] Auth failed - invalid credentials');
      return res.status(401).json({ error: 'Invalid Exchange credentials. Check your username and password.' });
    }
    if (response.statusCode >= 400) {
      console.log('[createMeeting] Exchange error body:', response.body?.substring(0, 500));
      return res.status(response.statusCode).json({ error: 'Exchange returned HTTP ' + response.statusCode });
    }
    if (response.body.includes('soap:Fault')) {
      const fault = response.body.match(/<faultstring[^>]*>([\s\S]*?)<\/faultstring>/i);
      const detail = fault ? fault[1] : 'Unknown SOAP fault';
      console.log('[createMeeting] SOAP fault:', detail);
      return res.status(500).json({ error: 'Exchange error: ' + detail });
    }
    if (response.body.includes('ResponseCode')) {
      const code = response.body.match(/<m:ResponseCode>([^<]+)<\/m:ResponseCode>/i);
      if (code && code[1] !== 'NoError') {
        const msg2 = response.body.match(/<m:MessageText>([^<]+)<\/m:MessageText>/i);
        const detail2 = msg2 ? msg2[1] : code[1];
        console.log('[createMeeting] EWS ResponseCode error:', detail2);
        return res.status(500).json({ error: 'Exchange: ' + detail2 });
      }
    }

    console.log('[createMeeting] Success - invite sent to', (attendees || []).join(', ') || 'organizer only');
    const sentTo = (attendees || []).filter(Boolean);
    res.json({ ok: true, start: toISO(startDate), end: toISO(endDate), sentTo, subject });
  } catch (err) {
    const msg = err.message || '';
    console.log('[createMeeting] Exception:', msg);
    if (msg.includes('ENOTFOUND') || msg.includes('ETIMEDOUT') || msg.includes('ECONNREFUSED'))
      return res.status(502).json({ error: 'Cannot reach Exchange server at ' + ewsUrl + '. Check the URL.' });
    res.status(502).json({ error: msg });
  }
}

function buildGetAttachmentSOAP(attachmentId) {
  return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header><t:RequestServerVersion Version="Exchange2013_SP1"/></soap:Header>
  <soap:Body>
    <m:GetAttachment>
      <m:AttachmentShape/>
      <m:AttachmentIds>
        <t:AttachmentId Id="${attachmentId}"/>
      </m:AttachmentIds>
    </m:GetAttachment>
  </soap:Body>
</soap:Envelope>`;
}

// POST /api/ews/email-body
export async function getEmailBody(req, res) {
  const { ewsUrl, username, password, itemId, changeKey } = req.body;
  if (!username || !password || !itemId) return res.status(400).json({ error: "username, password and itemId required." });
  if (!ewsUrl) return res.status(400).json({ error: "Exchange server URL is required." });

  const url = ewsUrl;
  let domain = "", user = username;
  if (username.includes("\\")) [domain, user] = username.split("\\");
  const ntlmOpts = { url, username: user, password, domain, workstation: "", rejectUnauthorized: false };

const soapBody = buildCreateMeetingSOAP(
    subject,
    toISO(startDate),
    toISO(endDate),
    attendees || [],
    body,
    location,
    startDate.toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' }),
    endDate.toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })
  );

  console.log('[createMeeting] Sending to', ewsUrl, '| subject:', subject, '| attendees:', (attendees || []).join(', ') || '(none)');

  try {
    const response = await ntlmPost({
      ...ntlmOpts,
      body: buildGetItemSOAP(itemId, changeKey),
      headers: { "Content-Type": "text/xml; charset=utf-8", "SOAPAction": '"http://schemas.microsoft.com/exchange/services/2006/messages/GetItem"' },
    });

    const bodyMatch = response.body.match(/<t:Body[^>]*>([\s\S]*?)<\/t:Body>/i);
    const raw = bodyMatch ? bodyMatch[1].trim() : "";
    let html = raw
      .replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"').replace(/&#xD;/g, "").replace(/&#xA;/g, "\n");

    // Parse inline attachments and replace cid: references with data: URLs
    const cidRefs = [...html.matchAll(/cid:([^\s"'>\)]+)/gi)].map(m => m[1]);
    if (cidRefs.length > 0) {
      const attBlocks = response.body.match(/<t:FileAttachment>[\s\S]*?<\/t:FileAttachment>/gi) || [];
      for (const block of attBlocks) {
        const attIdMatch  = block.match(/<t:AttachmentId Id="([^"]+)"/i);
        const contentId   = (block.match(/<t:ContentId>([^<]+)<\/t:ContentId>/i) || [])[1] || "";
        const contentType = (block.match(/<t:ContentType>([^<]+)<\/t:ContentType>/i) || [])[1] || "image/png";
        const isInline    = block.includes("<t:IsInline>true</t:IsInline>");
        if (!attIdMatch || !isInline) continue;

        const cleanCid = contentId.replace(/^<|>$/g, "");
        const matched  = cidRefs.some(c => c.replace(/^<|>$/g, "") === cleanCid);
        if (!matched) continue;

        try {
          const attResp = await ntlmPost({
            ...ntlmOpts,
            body: buildGetAttachmentSOAP(attIdMatch[1]),
            headers: { "Content-Type": "text/xml; charset=utf-8", "SOAPAction": '"http://schemas.microsoft.com/exchange/services/2006/messages/GetAttachment"' },
          });
          const b64 = (attResp.body.match(/<t:Content>([^<]+)<\/t:Content>/i) || [])[1] || "";
          if (b64) {
            const dataUrl = `data:${contentType};base64,${b64}`;
            html = html.replace(new RegExp(`cid:${cleanCid.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}`, "gi"), dataUrl);
          }
        } catch (_) { /* skip failed attachment */ }
      }
    }

    res.json({ body: html, bodyType: "html" });
  } catch (err) {
    const msg = err.message || '';
    console.log('[createMeeting] Exception:', msg);
    if (msg.includes('ENOTFOUND') || msg.includes('ETIMEDOUT') || msg.includes('ECONNREFUSED'))
      return res.status(502).json({ error: 'Cannot reach Exchange server at ' + ewsUrl + '. Check the URL.' });
    res.status(502).json({ error: msg });
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

  if (!ewsUrl) return res.status(400).json({ error: "Exchange server URL is required." });
  const url  = ewsUrl;
  let domain = "", user = username;
  if (username.includes("\\")) [domain, user] = username.split("\\");

const soapBody = buildCreateMeetingSOAP(
    subject,
    toISO(startDate),
    toISO(endDate),
    attendees || [],
    body,
    location,
    startDate.toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' }),
    endDate.toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })
  );

  console.log('[createMeeting] Sending to', ewsUrl, '| subject:', subject, '| attendees:', (attendees || []).join(', ') || '(none)');

  try {
    const response = await ntlmPost({
      url: ewsUrl, username: user, password, domain, workstation: '',
      body: soapBody,
      headers: { 'Content-Type': 'text/xml; charset=utf-8', 'SOAPAction': '"http://schemas.microsoft.com/exchange/services/2006/messages/CreateItem"' },
      rejectUnauthorized: false,
    });

    console.log('[createMeeting] Exchange HTTP status:', response.statusCode);

    if (response.statusCode === 401) {
      console.log('[createMeeting] Auth failed - invalid credentials');
      return res.status(401).json({ error: 'Invalid Exchange credentials. Check your username and password.' });
    }
    if (response.statusCode >= 400) {
      console.log('[createMeeting] Exchange error body:', response.body?.substring(0, 500));
      return res.status(response.statusCode).json({ error: 'Exchange returned HTTP ' + response.statusCode });
    }
    if (response.body.includes('soap:Fault')) {
      const fault = response.body.match(/<faultstring[^>]*>([\s\S]*?)<\/faultstring>/i);
      const detail = fault ? fault[1] : 'Unknown SOAP fault';
      console.log('[createMeeting] SOAP fault:', detail);
      return res.status(500).json({ error: 'Exchange error: ' + detail });
    }
    if (response.body.includes('ResponseCode')) {
      const code = response.body.match(/<m:ResponseCode>([^<]+)<\/m:ResponseCode>/i);
      if (code && code[1] !== 'NoError') {
        const msg2 = response.body.match(/<m:MessageText>([^<]+)<\/m:MessageText>/i);
        const detail2 = msg2 ? msg2[1] : code[1];
        console.log('[createMeeting] EWS ResponseCode error:', detail2);
        return res.status(500).json({ error: 'Exchange: ' + detail2 });
      }
    }

    console.log('[createMeeting] Success - invite sent to', (attendees || []).join(', ') || 'organizer only');
    const sentTo = (attendees || []).filter(Boolean);
    res.json({ ok: true, start: toISO(startDate), end: toISO(endDate), sentTo, subject });
  } catch (err) {
    const msg = err.message || '';
    console.log('[createMeeting] Exception:', msg);
    if (msg.includes('ENOTFOUND') || msg.includes('ETIMEDOUT') || msg.includes('ECONNREFUSED'))
      return res.status(502).json({ error: 'Cannot reach Exchange server at ' + ewsUrl + '. Check the URL.' });
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
  // HTTP 401 = server exists and requires auth — candidate URLs to probe
  const candidates = [
    `https://owa.${domain}/EWS/Exchange.asmx`,
    `https://mail.${domain}/EWS/Exchange.asmx`,
    `https://webmail.${domain}/EWS/Exchange.asmx`,
    `https://exchange.${domain}/EWS/Exchange.asmx`,
    `https://outlook.${domain}/EWS/Exchange.asmx`,
  ];
  res.json({ candidates, domain, recommended: `https://owa.${domain}/EWS/Exchange.asmx` });
}

// POST /api/ews/create-meeting
function buildCreateMeetingSOAP(subject, start, end, attendees, userBody, location, startDisplay, endDisplay) {
  const esc = (s) => String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const attXml = (attendees || []).filter(Boolean).map(
    (a) => '<t:Attendee><t:Mailbox><t:EmailAddress>' + esc(a.trim()) + '</t:EmailAddress></t:Mailbox></t:Attendee>'
  ).join('');

  // Build proper HTML email body so attendees see a formatted meeting invite
  const agendaSection = userBody
    ? '<p><strong>Agenda / Notes:</strong><br>' + esc(userBody).replace(/\n/g, '<br>') + '</p>'
    : '';
  const locRow = location
    ? '<tr><td style="padding:2px 14px 2px 0;color:#555;white-space:nowrap"><strong>Location</strong></td><td>' + esc(location) + '</td></tr>'
    : '';
  const attRow = (attendees && attendees.length)
    ? '<tr><td style="padding:2px 14px 2px 0;color:#555;white-space:nowrap"><strong>Attendees</strong></td><td>' + attendees.filter(Boolean).map((a) => esc(a.trim())).join(', ') + '</td></tr>'
    : '';

  const htmlBody = '<html><body style="font-family:Calibri,Arial,sans-serif;font-size:14px;color:#222">'
    + '<p>You have been invited to the following meeting:</p>'
    + '<table cellpadding="2" cellspacing="0" style="border-collapse:collapse;margin-bottom:16px">'
    + '<tr><td style="padding:2px 14px 2px 0;color:#555;white-space:nowrap"><strong>Subject</strong></td><td>' + esc(subject) + '</td></tr>'
    + '<tr><td style="padding:2px 14px 2px 0;color:#555;white-space:nowrap"><strong>Start</strong></td><td>' + esc(startDisplay || start) + '</td></tr>'
    + '<tr><td style="padding:2px 14px 2px 0;color:#555;white-space:nowrap"><strong>End</strong></td><td>' + esc(endDisplay || end) + '</td></tr>'
    + locRow + attRow
    + '</table>'
    + agendaSection
    + '</body></html>';

  return '<?xml version="1.0" encoding="utf-8"?>'
    + '<soap:Envelope'
    + ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"'
    + ' xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"'
    + ' xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"'
    + ' xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">'
    + '<soap:Header><t:RequestServerVersion Version="Exchange2013_SP1"/></soap:Header>'
    + '<soap:Body>'
    + '<m:CreateItem MessageDisposition="SendAndSaveCopy" SendMeetingInvitations="SendToAllAndSaveCopy">'
    + '<m:SavedItemFolderId><t:DistinguishedFolderId Id="calendar"/></m:SavedItemFolderId>'
    + '<m:Items>'
    + '<t:CalendarItem>'
    + '<t:Subject>' + esc(subject) + '</t:Subject>'
    + '<t:Body BodyType="HTML">' + esc(htmlBody) + '</t:Body>'
    + '<t:Start>' + start + '</t:Start>'
    + '<t:End>' + end + '</t:End>'
    + '<t:IsAllDayEvent>false</t:IsAllDayEvent>'
    + (location ? '<t:Location>' + esc(location) + '</t:Location>' : '')
    + (attXml ? '<t:RequiredAttendees>' + attXml + '</t:RequiredAttendees>' : '')
    + '</t:CalendarItem>'
    + '</m:Items>'
    + '</m:CreateItem>'
    + '</soap:Body>'
    + '</soap:Envelope>';
}

export async function createMeeting(req, res) {
  const { ewsUrl, username, password, subject, date, time, duration, attendees, location, body } = req.body;
  if (!username || !password) return res.status(400).json({ error: "Exchange credentials required. Connect in the Teams panel first." });
  if (!ewsUrl)   return res.status(400).json({ error: "Exchange server URL is required." });
  if (!subject)  return res.status(400).json({ error: "Meeting title is required." });
  if (!date)     return res.status(400).json({ error: "Meeting date is required." });

  // Parse date + time → start ISO string
  const timeMatch = (time || "09:00 AM").match(/(\d+):(\d+)\s*(AM|PM)?/i);
  let h  = timeMatch ? parseInt(timeMatch[1]) : 9;
  let mi = timeMatch ? parseInt(timeMatch[2]) : 0;
  if (timeMatch?.[3]?.toUpperCase() === "PM" && h !== 12) h += 12;
  if (timeMatch?.[3]?.toUpperCase() === "AM" && h === 12) h = 0;

  const startDate = new Date(`${date}T00:00:00`);
  startDate.setHours(h, mi, 0, 0);
  const durMins = parseInt((duration || "30 min").match(/\d+/)?.[0] || "30");
  const endDate = new Date(startDate.getTime() + durMins * 60000);

  const toISO = (d) => d.toISOString().replace(/\.\d{3}Z$/, "");

  let domain = "", user = username;
  if (username.includes("\\")) [domain, user] = username.split("\\");

const soapBody = buildCreateMeetingSOAP(
    subject,
    toISO(startDate),
    toISO(endDate),
    attendees || [],
    body,
    location,
    startDate.toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' }),
    endDate.toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })
  );

  console.log('[createMeeting] Sending to', ewsUrl, '| subject:', subject, '| attendees:', (attendees || []).join(', ') || '(none)');

  try {
    const response = await ntlmPost({
      url: ewsUrl, username: user, password, domain, workstation: '',
      body: soapBody,
      headers: { 'Content-Type': 'text/xml; charset=utf-8', 'SOAPAction': '"http://schemas.microsoft.com/exchange/services/2006/messages/CreateItem"' },
      rejectUnauthorized: false,
    });

    console.log('[createMeeting] Exchange HTTP status:', response.statusCode);

    if (response.statusCode === 401) {
      console.log('[createMeeting] Auth failed - invalid credentials');
      return res.status(401).json({ error: 'Invalid Exchange credentials. Check your username and password.' });
    }
    if (response.statusCode >= 400) {
      console.log('[createMeeting] Exchange error body:', response.body?.substring(0, 500));
      return res.status(response.statusCode).json({ error: 'Exchange returned HTTP ' + response.statusCode });
    }
    if (response.body.includes('soap:Fault')) {
      const fault = response.body.match(/<faultstring[^>]*>([\s\S]*?)<\/faultstring>/i);
      const detail = fault ? fault[1] : 'Unknown SOAP fault';
      console.log('[createMeeting] SOAP fault:', detail);
      return res.status(500).json({ error: 'Exchange error: ' + detail });
    }
    if (response.body.includes('ResponseCode')) {
      const code = response.body.match(/<m:ResponseCode>([^<]+)<\/m:ResponseCode>/i);
      if (code && code[1] !== 'NoError') {
        const msg2 = response.body.match(/<m:MessageText>([^<]+)<\/m:MessageText>/i);
        const detail2 = msg2 ? msg2[1] : code[1];
        console.log('[createMeeting] EWS ResponseCode error:', detail2);
        return res.status(500).json({ error: 'Exchange: ' + detail2 });
      }
    }

    console.log('[createMeeting] Success - invite sent to', (attendees || []).join(', ') || 'organizer only');
    const sentTo = (attendees || []).filter(Boolean);
    res.json({ ok: true, start: toISO(startDate), end: toISO(endDate), sentTo, subject });
  } catch (err) {
    const msg = err.message || '';
    console.log('[createMeeting] Exception:', msg);
    if (msg.includes('ENOTFOUND') || msg.includes('ETIMEDOUT') || msg.includes('ECONNREFUSED'))
      return res.status(502).json({ error: 'Cannot reach Exchange server at ' + ewsUrl + '. Check the URL.' });
    res.status(502).json({ error: msg });
  }
}
