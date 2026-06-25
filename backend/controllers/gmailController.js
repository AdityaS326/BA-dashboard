// backend/controllers/gmailController.js
// Full Gmail OAuth 2.0 — uses real Google credentials from .env

import { google } from 'googleapis';

const CLIENT_ID     = process.env.GMAIL_CLIENT_ID;
const CLIENT_SECRET = process.env.GMAIL_CLIENT_SECRET;
const REDIRECT_URI  = process.env.GMAIL_REDIRECT_URI || 'http://localhost:3000/api/gmail/callback';

const SCOPES = [
  'https://www.googleapis.com/auth/gmail.readonly',
  'https://www.googleapis.com/auth/gmail.send',
  'https://www.googleapis.com/auth/gmail.modify',
  'https://www.googleapis.com/auth/userinfo.email',
  'https://www.googleapis.com/auth/userinfo.profile',
];

// In-memory token + user store (cleared on restart)
let _tokens   = null;
let _userInfo  = null;   // { email, name, picture }

function makeOAuth2() {
  return new google.auth.OAuth2(CLIENT_ID, CLIENT_SECRET, REDIRECT_URI);
}

// ── GET /api/gmail/auth-url ───────────────────────────────────────────────────
export function getAuthUrl(req, res) {
  if (!CLIENT_ID || !CLIENT_SECRET) {
    return res.status(503).json({ error: 'Gmail credentials not configured in .env' });
  }
  const oauth2 = makeOAuth2();
  const opts   = {
    access_type: 'offline',
    scope:       SCOPES,
    prompt:      'consent',
  };
  if (req.query.hint) opts.login_hint = req.query.hint;
  res.json({ url: oauth2.generateAuthUrl(opts) });
}

// ── GET /api/gmail/callback ───────────────────────────────────────────────────
export async function handleCallback(req, res) {
  const { code, error } = req.query;
  if (error) return res.redirect('/?panel=gm&auth=cancelled');
  if (!code)  return res.status(400).send('Missing code');
  try {
    const oauth2 = makeOAuth2();
    const { tokens } = await oauth2.getToken(code);
    _tokens = tokens;
    oauth2.setCredentials(tokens);

    const oauth2Info = google.oauth2({ version: 'v2', auth: oauth2 });
    const { data }   = await oauth2Info.userinfo.get();
    _userInfo = { email: data.email, name: data.name, picture: data.picture || '' };

    res.redirect('/?panel=gm');
  } catch (err) {
    console.error('[Gmail callback]', err.message);
    res.redirect(`/?panel=gm&auth=error&msg=${encodeURIComponent(err.message)}`);
  }
}

// ── GET /api/gmail/status ─────────────────────────────────────────────────────
export async function getStatus(req, res) {
  if (!_tokens || !_userInfo) return res.json({ authenticated: false });

  try {
    const oauth2 = makeOAuth2();
    oauth2.setCredentials(_tokens);

    // Auto-refresh if expired
    if (_tokens.expiry_date && Date.now() > _tokens.expiry_date - 60_000) {
      const { credentials } = await oauth2.refreshAccessToken();
      _tokens = credentials;
      oauth2.setCredentials(_tokens);
    }

    // Re-fetch user info to confirm still valid
    const o2info = google.oauth2({ version: 'v2', auth: oauth2 });
    const { data } = await o2info.userinfo.get();
    _userInfo = { email: data.email, name: data.name, picture: data.picture || '' };

    res.json({ authenticated: true, ..._userInfo });
  } catch {
    _tokens   = null;
    _userInfo = null;
    res.json({ authenticated: false });
  }
}

// ── GET /api/gmail/emails?label=INBOX ────────────────────────────────────────
export async function getEmails(req, res) {
  if (!_tokens) return res.status(401).json({ error: 'Not authenticated' });
  try {
    const oauth2 = makeOAuth2();
    oauth2.setCredentials(_tokens);
    const gmail  = google.gmail({ version: 'v1', auth: oauth2 });

    const labelId = req.query.label || 'INBOX';
    const listRes = await gmail.users.messages.list({
      userId:     'me',
      maxResults: 25,
      labelIds:   labelId !== 'ALL' ? [labelId] : undefined,
      q:          req.query.q || '',
    });

    if (!listRes.data.messages?.length) return res.json({ emails: [] });

    // Use allSettled so one failing fetch doesn't break the whole list
    const results = await Promise.allSettled(
      listRes.data.messages.map(({ id }) =>
        gmail.users.messages.get({
          userId: 'me', id, format: 'metadata',
          metadataHeaders: ['Subject', 'From', 'To', 'Date'],
        })
      )
    );

    const emails = results
      .filter(r => r.status === 'fulfilled')
      .map(r => {
        const msg  = r.value;
        const hdrs = msg.data.payload?.headers || [];
        const h    = (n) => hdrs.find(x => x.name === n)?.value || '';
        return {
          id:       msg.data.id,
          threadId: msg.data.threadId,
          subject:  h('Subject') || '(no subject)',
          from:     h('From'),
          to:       h('To'),
          date:     h('Date'),
          snippet:  msg.data.snippet || '',
          isUnread: (msg.data.labelIds || []).includes('UNREAD'),
        };
      });

    res.json({ emails });
  } catch (err) {
    // Detect scope error and force re-auth
    if (/insufficient authentication scopes/i.test(err.message)) {
      _tokens   = null;
      _userInfo = null;
      return res.status(401).json({
        error: 'SCOPE_ERROR',
        message: 'Gmail permission not granted. Please sign out and sign in again — make sure to allow all Gmail permissions when Google asks.',
      });
    }
    if (/API.*disabled|has not been used/i.test(err.message)) {
      return res.status(503).json({
        error: 'API_DISABLED',
        message: 'Gmail API is not enabled in your Google Cloud project. Go to console.cloud.google.com → APIs & Services → Library → search "Gmail API" → Enable.',
      });
    }
    res.status(500).json({ error: err.message });
  }
}

// ── GET /api/gmail/email/:id ──────────────────────────────────────────────────
export async function getEmailById(req, res) {
  if (!_tokens) return res.status(401).json({ error: 'Not authenticated' });
  try {
    const oauth2 = makeOAuth2();
    oauth2.setCredentials(_tokens);
    const gmail = google.gmail({ version: 'v1', auth: oauth2 });

    const msg  = await gmail.users.messages.get({ userId: 'me', id: req.params.id, format: 'full' });
    const hdrs = msg.data.payload.headers;
    const h    = (n) => hdrs.find(x => x.name === n)?.value || '';

    res.json({
      id:      msg.data.id,
      subject: h('Subject'),
      from:    h('From'),
      to:      h('To'),
      date:    h('Date'),
      body:    _extractBody(msg.data.payload),
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}

// ── POST /api/gmail/send ──────────────────────────────────────────────────────
export async function sendEmail(req, res) {
  if (!_tokens) return res.status(401).json({ error: 'Not authenticated' });
  const { to, subject, body } = req.body || {};
  if (!to || !subject || !body) return res.status(400).json({ error: 'to, subject, body are required' });
  try {
    const oauth2 = makeOAuth2();
    oauth2.setCredentials(_tokens);
    const gmail  = google.gmail({ version: 'v1', auth: oauth2 });

    const raw = Buffer.from(
      `To: ${to}\r\nSubject: ${subject}\r\nContent-Type: text/plain; charset=utf-8\r\n\r\n${body}`
    ).toString('base64url');

    await gmail.users.messages.send({ userId: 'me', requestBody: { raw } });
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}

// ── POST /api/gmail/logout ────────────────────────────────────────────────────
export function gmailLogout(req, res) {
  _tokens   = null;
  _userInfo = null;
  res.json({ success: true });
}

// ── Stubs kept for route compatibility ───────────────────────────────────────
export async function gmailLogin(req, res) {
  res.status(410).json({ error: 'Use OAuth flow: GET /api/gmail/auth-url' });
}

// ── Body extractor ────────────────────────────────────────────────────────────
function _extractBody(payload) {
  if (!payload) return '<p>No content</p>';
  if (payload.body?.data)
    return Buffer.from(payload.body.data, 'base64url').toString('utf8');
  if (payload.parts) {
    // Prefer HTML
    for (const part of payload.parts) {
      if (part.mimeType === 'text/html' && part.body?.data)
        return Buffer.from(part.body.data, 'base64url').toString('utf8');
    }
    // Plain text fallback
    for (const part of payload.parts) {
      if (part.mimeType === 'text/plain' && part.body?.data)
        return `<pre style="white-space:pre-wrap;font-family:inherit;font-size:13px;padding:8px">${
          Buffer.from(part.body.data, 'base64url').toString('utf8')
            .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
        }</pre>`;
      if (part.parts) { const r = _extractBody(part); if (r) return r; }
    }
  }
  return '<p style="color:#aaa;padding:16px">No readable content</p>';
}
