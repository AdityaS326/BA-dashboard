// backend/controllers/whatsappController.js
// Requires: npm install whatsapp-web.js qrcode
// (whatsapp-web.js pulls Puppeteer + ~150 MB Chromium on first install)

let Client, LocalAuth, qrcode;
let libLoaded = false;

async function loadLibs() {
  if (libLoaded) return true;
  try {
    const ww = await import('whatsapp-web.js');
    Client    = ww.Client;
    LocalAuth = ww.LocalAuth;
    const qr  = await import('qrcode');
    qrcode    = qr.default;
    libLoaded = true;
    return true;
  } catch {
    return false;
  }
}

// ── State ─────────────────────────────────────────────────────────────────────
let waClient   = null;
let waStatus   = 'disconnected'; // disconnected | initializing | qr | connecting | ready
let currentQR  = null;
let chatsCache = [];
let profile    = null;

function resetState() {
  waClient   = null;
  waStatus   = 'disconnected';
  currentQR  = null;
  chatsCache = [];
  profile    = null;
}

// ── Controllers ───────────────────────────────────────────────────────────────

export async function initWhatsApp(req, res) {
  if (waStatus === 'ready')  return res.json({ status: 'ready' });
  if (waStatus !== 'disconnected') return res.json({ status: waStatus });

  const ok = await loadLibs();
  if (!ok) {
    return res.status(503).json({
      error: 'whatsapp-web.js is not installed. Run: npm install whatsapp-web.js',
    });
  }

  waStatus = 'initializing';

  waClient = new Client({
    authStrategy: new LocalAuth({ dataPath: './.wwebjs_auth' }),
    puppeteer: {
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--disable-gpu'],
    },
  });

  waClient.on('qr', async (qr) => {
    waStatus  = 'qr';
    currentQR = await qrcode.toDataURL(qr);
  });

  waClient.on('authenticated', () => {
    waStatus  = 'connecting';
    currentQR = null;
  });

  waClient.on('ready', async () => {
    waStatus = 'ready';
    try {
      const info = waClient.info;
      profile = { name: info.pushname, phone: info.wid.user };
      await refreshChatsInternal();
    } catch (e) {
      console.error('[WhatsApp] ready error:', e.message);
    }
  });

  waClient.on('disconnected', () => resetState());
  waClient.on('auth_failure', () => resetState());

  waClient.initialize().catch(err => {
    console.error('[WhatsApp] init error:', err.message);
    resetState();
  });

  res.json({ status: 'initializing' });
}

export function getStatus(req, res) {
  const resp = { status: waStatus };
  if (waStatus === 'qr' && currentQR)    resp.qr      = currentQR;
  if (waStatus === 'ready' && profile)   resp.profile = profile;
  res.json(resp);
}

export function getChats(req, res) {
  if (waStatus !== 'ready') return res.status(401).json({ error: 'Not connected' });
  res.json({ chats: chatsCache, profile });
}

export async function refreshChats(req, res) {
  if (waStatus !== 'ready') return res.status(401).json({ error: 'Not connected' });
  try {
    await refreshChatsInternal();
    res.json({ chats: chatsCache });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}

export async function logoutWhatsApp(req, res) {
  try {
    if (waClient) {
      await waClient.logout().catch(() => {});
      await waClient.destroy().catch(() => {});
    }
  } catch { /* ignore */ }
  resetState();
  res.json({ success: true });
}

// ── Internal helpers ──────────────────────────────────────────────────────────

async function refreshChatsInternal() {
  const all = await waClient.getChats();
  chatsCache = all.slice(0, 40).map(c => ({
    id:          c.id._serialized,
    name:        c.name || c.id.user,
    isGroup:     c.isGroup,
    unreadCount: c.unreadCount || 0,
    lastMessage: c.lastMessage
      ? {
          body:      (c.lastMessage.body || '').substring(0, 120),
          timestamp: c.lastMessage.timestamp,
          fromMe:    c.lastMessage.fromMe,
        }
      : null,
  }));
}
