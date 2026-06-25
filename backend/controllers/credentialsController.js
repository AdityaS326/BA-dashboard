// backend/controllers/credentialsController.js
// Save and load API credentials from SQLite (so users can configure via dashboard UI)

import db from '../utils/db.js';

const ALLOWED_KEYS = new Set([
  'GMAIL_CLIENT_ID',
  'GMAIL_CLIENT_SECRET',
  'GMAIL_REDIRECT_URI',
  'WA_PHONE_NUMBER_ID',
  'WA_ACCESS_TOKEN',
  'WA_WEBHOOK_VERIFY_TOKEN',
]);

// GET /api/credentials?keys=GMAIL_CLIENT_ID,GMAIL_CLIENT_SECRET
export function getCredentials(req, res) {
  const keys = (req.query.keys || '').split(',').map(k => k.trim()).filter(Boolean);
  const result = {};
  for (const key of keys) {
    if (!ALLOWED_KEYS.has(key)) continue;
    const row = db.prepare('SELECT value FROM credentials WHERE key = ?').get(key);
    // Mask secrets — only return whether they are set, not the actual value
    if (row) {
      result[key] = key.toLowerCase().includes('secret') ? '••••••••' : row.value;
      result[`${key}_set`] = true;
    } else {
      result[`${key}_set`] = false;
    }
  }
  res.json(result);
}

// POST /api/credentials  { GMAIL_CLIENT_ID: '...', GMAIL_CLIENT_SECRET: '...' }
export function saveCredentials(req, res) {
  const body = req.body || {};
  const saved = [];
  const upsert = db.prepare(`
    INSERT INTO credentials (key, value, updated_at)
    VALUES (?, ?, datetime('now'))
    ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = excluded.updated_at
  `);

  for (const [key, value] of Object.entries(body)) {
    if (!ALLOWED_KEYS.has(key)) continue;
    if (typeof value !== 'string' || !value.trim()) continue;
    // Skip masked placeholders sent back from the UI
    if (value === '••••••••') continue;
    upsert.run(key, value.trim());
    // Also update process.env so current session picks them up without restart
    process.env[key] = value.trim();
    saved.push(key);
  }

  res.json({ saved });
}

// Helper — used by gmailController to read credentials from DB if not in env
export function getCredential(key) {
  if (process.env[key] && !process.env[key].startsWith('your-')) return process.env[key];
  const row = db.prepare('SELECT value FROM credentials WHERE key = ?').get(key);
  return row?.value || null;
}
