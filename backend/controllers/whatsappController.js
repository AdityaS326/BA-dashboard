// backend/controllers/whatsappController.js
// Meta WhatsApp Business Cloud API
// Docs: https://developers.facebook.com/docs/whatsapp/cloud-api

import { getCredential } from './credentialsController.js';

const GRAPH = 'https://graph.facebook.com/v19.0';

function getCreds() {
  return {
    phoneNumberId:     getCredential('WA_PHONE_NUMBER_ID'),
    accessToken:       getCredential('WA_ACCESS_TOKEN'),
    webhookVerifyToken: getCredential('WA_WEBHOOK_VERIFY_TOKEN') || 'ba_dashboard_verify',
  };
}

// In-memory message store  (keyed by contact phone number)
// { "919876543210": [ { id, from, to, body, timestamp, direction } ] }
const _messages = {};
let   _profile  = null;   // { name, phone }

// ── GET /api/whatsapp/status ──────────────────────────────────────────────────
export async function getStatus(req, res) {
  const { phoneNumberId, accessToken } = getCreds();
  if (!phoneNumberId || !accessToken) {
    return res.json({ connected: false, reason: 'not_configured' });
  }
  try {
    const r = await fetch(
      `${GRAPH}/${phoneNumberId}?fields=display_phone_number,verified_name`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    const d = await r.json();
    if (d.error) return res.json({ connected: false, reason: d.error.message });
    _profile = { name: d.verified_name || d.display_phone_number, phone: d.display_phone_number };
    res.json({ connected: true, profile: _profile });
  } catch (err) {
    res.json({ connected: false, reason: err.message });
  }
}

// ── POST /api/whatsapp/send ───────────────────────────────────────────────────
export async function sendMessage(req, res) {
  const { phoneNumberId, accessToken } = getCreds();
  if (!phoneNumberId || !accessToken) {
    return res.status(503).json({ error: 'WhatsApp credentials not configured.' });
  }
  const { to, message } = req.body || {};
  if (!to || !message) return res.status(400).json({ error: '"to" and "message" are required.' });

  const toClean = to.replace(/\D/g, '');
  try {
    const r = await fetch(`${GRAPH}/${phoneNumberId}/messages`, {
      method: 'POST',
      headers: {
        Authorization:  `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        messaging_product: 'whatsapp',
        to:   toClean,
        type: 'text',
        text: { body: message },
      }),
    });
    const d = await r.json();
    if (d.error) return res.status(400).json({ error: d.error.message });

    // Store sent message locally
    if (!_messages[toClean]) _messages[toClean] = [];
    _messages[toClean].push({
      id:        d.messages?.[0]?.id || Date.now().toString(),
      from:      'me',
      to:        toClean,
      body:      message,
      timestamp: Math.floor(Date.now() / 1000),
      direction: 'outbound',
    });

    res.json({ success: true, messageId: d.messages?.[0]?.id });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}

// ── GET /api/whatsapp/conversations ──────────────────────────────────────────
export function getConversations(req, res) {
  const contacts = Object.entries(_messages).map(([phone, msgs]) => {
    const last = msgs[msgs.length - 1];
    const unread = msgs.filter(m => m.direction === 'inbound' && !m.read).length;
    return { phone, name: last.fromName || ('+' + phone), lastMessage: last, unread };
  });
  contacts.sort((a,b) => b.lastMessage.timestamp - a.lastMessage.timestamp);
  res.json({ conversations: contacts, profile: _profile });
}

// ── GET /api/whatsapp/messages/:phone ─────────────────────────────────────────
export function getMessages(req, res) {
  const phone = req.params.phone.replace(/\D/g,'');
  const msgs  = _messages[phone] || [];
  // Mark as read
  msgs.forEach(m => { if (m.direction === 'inbound') m.read = true; });
  res.json({ messages: msgs });
}

// ── GET /api/whatsapp/webhook  (Meta verification challenge) ──────────────────
export function webhookVerify(req, res) {
  const { 'hub.mode': mode, 'hub.verify_token': token, 'hub.challenge': challenge } = req.query;
  const { webhookVerifyToken } = getCreds();
  if (mode === 'subscribe' && token === webhookVerifyToken) {
    console.log('[WhatsApp] Webhook verified');
    return res.status(200).send(challenge);
  }
  res.status(403).json({ error: 'Verification failed' });
}

// ── POST /api/whatsapp/webhook  (incoming messages from Meta) ─────────────────
export function webhookReceive(req, res) {
  res.sendStatus(200); // Always acknowledge immediately

  try {
    const body = req.body;
    if (body.object !== 'whatsapp_business_account') return;

    for (const entry of body.entry || []) {
      for (const change of entry.changes || []) {
        const val = change.value;
        if (!val) continue;

        // Incoming messages
        for (const msg of val.messages || []) {
          const from    = msg.from;
          const text    = msg.text?.body || msg.type;
          const name    = val.contacts?.find(c => c.wa_id === from)?.profile?.name || ('+' + from);

          if (!_messages[from]) _messages[from] = [];
          // Avoid duplicates
          if (!_messages[from].find(m => m.id === msg.id)) {
            _messages[from].push({
              id:        msg.id,
              from,
              fromName:  name,
              body:      text,
              timestamp: Number(msg.timestamp),
              direction: 'inbound',
              read:      false,
            });
          }
        }
      }
    }
  } catch (err) {
    console.error('[WhatsApp webhook]', err.message);
  }
}

// ── GET /api/whatsapp/profile ─────────────────────────────────────────────────
export async function getProfile(req, res) {
  const { phoneNumberId, accessToken } = getCreds();
  if (!phoneNumberId || !accessToken) return res.status(503).json({ error: 'Not configured' });
  try {
    const r = await fetch(
      `${GRAPH}/${phoneNumberId}?fields=display_phone_number,verified_name`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    const d = await r.json();
    if (d.error) return res.status(400).json({ error: d.error.message });
    res.json({ phone: d.display_phone_number, name: d.verified_name });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
