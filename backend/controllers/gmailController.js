// backend/controllers/gmailController.js
import { google } from 'googleapis';

function createClient() {
  return new google.auth.OAuth2(
    process.env.GMAIL_CLIENT_ID,
    process.env.GMAIL_CLIENT_SECRET,
    process.env.GMAIL_REDIRECT_URI || 'http://localhost:3000/api/gmail/callback'
  );
}

// In-memory token store (single-user; restart clears it)
let tokenStore = null;

export async function getAuthUrl(req, res) {
  if (!process.env.GMAIL_CLIENT_ID || !process.env.GMAIL_CLIENT_SECRET) {
    return res.status(503).json({ error: 'GMAIL_CLIENT_ID and GMAIL_CLIENT_SECRET are not set in .env' });
  }
  const client = createClient();
  const url = client.generateAuthUrl({
    access_type: 'offline',
    scope: [
      'https://www.googleapis.com/auth/gmail.readonly',
      'https://www.googleapis.com/auth/gmail.send',
      'https://www.googleapis.com/auth/userinfo.email',
      'https://www.googleapis.com/auth/userinfo.profile',
    ],
    prompt: 'consent',
  });
  res.json({ url });
}

export async function handleCallback(req, res) {
  const { code, error } = req.query;
  if (error) return res.redirect('/?panel=gm&auth=cancelled');
  if (!code) return res.status(400).send('Missing code');

  try {
    const client = createClient();
    const { tokens } = await client.getToken(code);
    tokenStore = tokens;
    client.setCredentials(tokens);

    const oauth2 = google.oauth2({ version: 'v2', auth: client });
    const { data } = await oauth2.userinfo.get();

    // Redirect back into the app at the Gmail panel
    res.redirect(`/?panel=gm`);
  } catch (err) {
    res.redirect(`/?panel=gm&auth=error&msg=${encodeURIComponent(err.message)}`);
  }
}

export async function getStatus(req, res) {
  if (!tokenStore) return res.json({ authenticated: false });
  try {
    const client = createClient();
    client.setCredentials(tokenStore);
    if (tokenStore.expiry_date && Date.now() > tokenStore.expiry_date - 60000) {
      const { credentials } = await client.refreshAccessToken();
      tokenStore = credentials;
    }
    const oauth2 = google.oauth2({ version: 'v2', auth: client });
    const { data } = await oauth2.userinfo.get();
    res.json({ authenticated: true, email: data.email, name: data.name, picture: data.picture });
  } catch {
    tokenStore = null;
    res.json({ authenticated: false });
  }
}

export async function getEmails(req, res) {
  if (!tokenStore) return res.status(401).json({ error: 'Not authenticated' });
  try {
    const client = createClient();
    client.setCredentials(tokenStore);
    const gmail = google.gmail({ version: 'v1', auth: client });

    const labelId = req.query.label || 'INBOX';
    const list = await gmail.users.messages.list({
      userId: 'me', maxResults: 25,
      labelIds: labelId !== 'ALL' ? [labelId] : undefined,
      q: req.query.q || '',
    });

    if (!list.data.messages?.length) return res.json({ emails: [] });

    const emails = await Promise.all(
      list.data.messages.map(async ({ id }) => {
        const msg = await gmail.users.messages.get({
          userId: 'me', id, format: 'metadata',
          metadataHeaders: ['Subject', 'From', 'To', 'Date'],
        });
        const headers = msg.data.payload.headers;
        const h = (name) => headers.find(x => x.name === name)?.value || '';
        return {
          id,
          threadId: msg.data.threadId,
          subject: h('Subject') || '(no subject)',
          from: h('From'),
          to: h('To'),
          date: h('Date'),
          snippet: msg.data.snippet || '',
          isUnread: (msg.data.labelIds || []).includes('UNREAD'),
        };
      })
    );
    res.json({ emails });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}

export async function getEmailById(req, res) {
  if (!tokenStore) return res.status(401).json({ error: 'Not authenticated' });
  try {
    const client = createClient();
    client.setCredentials(tokenStore);
    const gmail = google.gmail({ version: 'v1', auth: client });

    const msg = await gmail.users.messages.get({ userId: 'me', id: req.params.id, format: 'full' });
    const headers = msg.data.payload.headers;
    const h = (name) => headers.find(x => x.name === name)?.value || '';

    const body = extractBody(msg.data.payload);

    res.json({
      id: msg.data.id,
      subject: h('Subject'),
      from: h('From'),
      to: h('To'),
      date: h('Date'),
      body,
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}

export async function sendEmail(req, res) {
  if (!tokenStore) return res.status(401).json({ error: 'Not authenticated' });
  const { to, subject, body } = req.body || {};
  if (!to || !subject || !body) return res.status(400).json({ error: 'to, subject, body are required' });

  try {
    const client = createClient();
    client.setCredentials(tokenStore);
    const gmail = google.gmail({ version: 'v1', auth: client });

    const raw = Buffer.from(
      `To: ${to}\r\nSubject: ${subject}\r\nContent-Type: text/plain; charset=utf-8\r\n\r\n${body}`
    ).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

    await gmail.users.messages.send({ userId: 'me', requestBody: { raw } });
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}

export function gmailLogout(req, res) {
  tokenStore = null;
  res.json({ success: true });
}

// ── Helper ────────────────────────────────────────────────────────────────────

function extractBody(payload) {
  if (payload.body?.data) return Buffer.from(payload.body.data, 'base64').toString('utf8');
  if (payload.parts) {
    for (const part of payload.parts) {
      if (part.mimeType === 'text/html' && part.body?.data)
        return Buffer.from(part.body.data, 'base64').toString('utf8');
    }
    for (const part of payload.parts) {
      if (part.mimeType === 'text/plain' && part.body?.data)
        return `<pre style="white-space:pre-wrap;font-family:inherit">${Buffer.from(part.body.data, 'base64').toString('utf8')}</pre>`;
      if (part.parts) { const r = extractBody(part); if (r) return r; }
    }
  }
  return '<p style="color:#aaa">No readable content</p>';
}

