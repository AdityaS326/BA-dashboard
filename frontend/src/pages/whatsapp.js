// frontend/src/pages/whatsapp.js
// WhatsApp Web integration via whatsapp-web.js (QR-code based login)

let _waPollTimer = null;

// ── Called by nav() when WhatsApp panel activates ─────────────────────────────
export async function waInit() {
  try {
    const data = await waFetch('/api/whatsapp/status');
    _waHandleStatus(data);
  } catch {
    _waRenderDisconnected();
  }
}

// ── Connect button ────────────────────────────────────────────────────────────
window.waConnect = async function () {
  _waRenderLoading('Initialising WhatsApp client…');
  try {
    const data = await waFetch('/api/whatsapp/init', 'POST');
    if (data.error) {
      _waRenderInstallPrompt(data.error);
    } else {
      _waStartPoll();
    }
  } catch (err) {
    _waRenderDisconnected('Could not reach backend: ' + err.message);
  }
};

// ── Disconnect ────────────────────────────────────────────────────────────────
window.waDisconnect = async function () {
  _waStopPoll();
  try { await waFetch('/api/whatsapp/logout', 'POST'); } catch { /* ignore */ }
  _waRenderDisconnected();
  showToast('WhatsApp disconnected', 'success');
};

// ── Refresh chat list ─────────────────────────────────────────────────────────
window.waRefreshChats = async function () {
  const btn = document.querySelector('#wa-refresh-btn');
  if (btn) { btn.disabled = true; btn.innerHTML = '<i class="ti ti-loader-2" style="animation:spin 1s linear infinite"></i>'; }
  try {
    await waFetch('/api/whatsapp/refresh', 'POST');
    const { chats, profile } = await waFetch('/api/whatsapp/chats');
    _waRenderChats(chats, profile);
    showToast('Chats refreshed', 'success');
  } catch (err) {
    showToast('Refresh failed: ' + err.message, 'error');
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ti ti-refresh" style="font-size:13px"></i> Refresh'; }
  }
};

// ── Status poller ─────────────────────────────────────────────────────────────
function _waStartPoll() {
  _waStopPoll();
  _waPollTimer = setInterval(async () => {
    try {
      const data = await waFetch('/api/whatsapp/status');
      _waHandleStatus(data);
      if (data.status === 'ready') _waStopPoll();
    } catch { /* keep trying */ }
  }, 2200);
}

function _waStopPoll() {
  if (_waPollTimer) { clearInterval(_waPollTimer); _waPollTimer = null; }
}

function _waHandleStatus(data) {
  if (data.status === 'disconnected') _waRenderDisconnected();
  else if (data.status === 'initializing' || data.status === 'connecting') _waRenderLoading('Connecting to WhatsApp…');
  else if (data.status === 'qr')    _waRenderQR(data.qr);
  else if (data.status === 'ready') {
    waFetch('/api/whatsapp/chats')
      .then(({ chats, profile }) => _waRenderChats(chats, profile))
      .catch(() => _waRenderChats([], null));
  }
}

// ── Render states ─────────────────────────────────────────────────────────────

function _waRenderDisconnected(warn) {
  const body = _wqs('#wa-dynamic-body');
  if (!body) return;
  body.innerHTML = `
    <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;padding:56px 24px;text-align:center">
      <div style="width:80px;height:80px;background:#fff;border-radius:20px;display:flex;align-items:center;justify-content:center;box-shadow:0 4px 20px rgba(0,0,0,.10);margin-bottom:24px">
        <i class="ti ti-brand-whatsapp" style="font-size:48px;color:#25D366"></i>
      </div>
      <h2 style="font-size:20px;font-weight:700;color:var(--text);margin:0 0 8px">Connect WhatsApp</h2>
      <p style="font-size:13px;color:var(--muted);max-width:320px;margin:0 0 28px;line-height:1.55">
        Scan a QR code with your phone to link this dashboard to your WhatsApp account.
      </p>
      ${warn ? `<div style="padding:10px 14px;background:rgba(234,67,53,.06);border:1px solid rgba(234,67,53,.2);border-radius:var(--r-sm);font-size:12px;color:var(--muted);margin-bottom:20px;max-width:380px;text-align:left">
        <i class="ti ti-alert-circle" style="color:#EA4335"></i> ${wEsc(warn)}</div>` : ''}
      <button onclick="waConnect()"
        style="background:#25D366;color:#fff;border:none;border-radius:8px;padding:12px 28px;font-size:14px;font-weight:600;cursor:pointer;display:flex;align-items:center;gap:8px;box-shadow:0 2px 8px rgba(37,211,102,.3);transition:box-shadow .15s"
        onmouseenter="this.style.boxShadow='0 4px 16px rgba(37,211,102,.4)'"
        onmouseleave="this.style.boxShadow='0 2px 8px rgba(37,211,102,.3)'">
        <i class="ti ti-device-mobile" style="font-size:16px"></i> Connect WhatsApp
      </button>
      <p style="font-size:11px;color:var(--muted);margin-top:18px;line-height:1.5">
        Requires <code>whatsapp-web.js</code> installed on the backend.<br>
        Run: <code>cd backend &amp;&amp; npm install whatsapp-web.js</code>
      </p>
    </div>`;
}

function _waRenderInstallPrompt(errMsg) {
  const body = _wqs('#wa-dynamic-body');
  if (!body) return;
  body.innerHTML = `
    <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;padding:56px 24px;text-align:center">
      <div style="width:80px;height:80px;background:#fff8f0;border-radius:20px;display:flex;align-items:center;justify-content:center;box-shadow:0 4px 20px rgba(0,0,0,.08);margin-bottom:24px">
        <i class="ti ti-package" style="font-size:44px;color:#FBBC05"></i>
      </div>
      <h2 style="font-size:18px;font-weight:700;color:var(--text);margin:0 0 10px">Package Not Installed</h2>
      <p style="font-size:13px;color:var(--muted);max-width:360px;margin:0 0 24px;line-height:1.55">
        The <strong>whatsapp-web.js</strong> library needs to be installed first.<br>
        It also downloads Chromium (~150 MB) on first install.
      </p>
      <div style="background:var(--surface2);border:1px solid var(--border);border-radius:var(--r-sm);padding:12px 20px;font-family:monospace;font-size:13px;color:var(--text);margin-bottom:20px;text-align:left">
        cd backend<br>npm install whatsapp-web.js
      </div>
      <button onclick="waConnect()" style="background:#25D366;color:#fff;border:none;border-radius:8px;padding:10px 24px;font-size:13px;font-weight:600;cursor:pointer">
        Retry after installing
      </button>
    </div>`;
}

function _waRenderLoading(msg) {
  const body = _wqs('#wa-dynamic-body');
  if (!body) return;
  body.innerHTML = `
    <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;padding:56px 24px;text-align:center">
      <i class="ti ti-loader-2" style="font-size:48px;color:#25D366;animation:spin 1s linear infinite;margin-bottom:20px"></i>
      <p style="font-size:14px;color:var(--muted)">${wEsc(msg)}</p>
      <p style="font-size:12px;color:var(--muted);margin-top:8px">This may take 10–30 seconds on first run…</p>
    </div>`;
}

function _waRenderQR(qrDataUrl) {
  const body = _wqs('#wa-dynamic-body');
  if (!body) return;
  body.innerHTML = `
    <div style="display:flex;flex-direction:column;align-items:center;padding:32px 24px;text-align:center">
      <h2 style="font-size:18px;font-weight:700;color:var(--text);margin:0 0 6px">Scan with WhatsApp</h2>
      <p style="font-size:13px;color:var(--muted);max-width:320px;margin:0 0 20px;line-height:1.5">
        Open WhatsApp on your phone → tap <strong>⋮</strong> menu → <strong>Linked devices</strong> → <strong>Link a device</strong>
      </p>
      <div style="background:#fff;padding:12px;border-radius:12px;box-shadow:0 4px 20px rgba(0,0,0,.10);margin-bottom:20px">
        <img src="${qrDataUrl}" alt="WhatsApp QR Code" style="width:220px;height:220px;display:block">
      </div>
      <div style="display:flex;align-items:center;gap:8px;padding:10px 14px;background:rgba(37,211,102,.06);border:1px solid rgba(37,211,102,.2);border-radius:var(--r-sm);font-size:12px;color:var(--muted)">
        <i class="ti ti-loader-2" style="color:#25D366;font-size:14px;animation:spin 1s linear infinite"></i>
        Waiting for you to scan…
      </div>
      <button onclick="waDisconnect()" style="margin-top:16px;background:none;border:1px solid var(--border);border-radius:var(--r-sm);padding:7px 16px;font-size:12px;color:var(--muted);cursor:pointer">
        Cancel
      </button>
    </div>`;
}

function _waRenderChats(chats, profile) {
  const body = _wqs('#wa-dynamic-body');
  if (!body) return;

  const profileHtml = profile
    ? `<div style="display:flex;align-items:center;gap:8px;padding:10px 14px;background:rgba(37,211,102,.06);border:1px solid rgba(37,211,102,.15);border-radius:var(--r-sm);font-size:12px;color:var(--muted);margin-bottom:12px">
        <i class="ti ti-brand-whatsapp" style="color:#25D366;font-size:16px"></i>
        <span>Connected as <strong style="color:var(--text)">${wEsc(profile.name || profile.phone)}</strong>
        ${profile.phone ? `<span style="opacity:.6"> (${wEsc(profile.phone)})</span>` : ''}</span>
        <div style="flex:1"></div>
        <button id="wa-refresh-btn" onclick="waRefreshChats()" style="background:none;border:1px solid var(--border);border-radius:var(--r-sm);padding:4px 10px;font-size:11px;cursor:pointer;display:flex;align-items:center;gap:4px;color:var(--text)">
          <i class="ti ti-refresh" style="font-size:13px"></i> Refresh
        </button>
        <button onclick="waDisconnect()" style="background:none;border:1px solid var(--border);border-radius:var(--r-sm);padding:4px 10px;font-size:11px;cursor:pointer;color:var(--muted)">
          Disconnect
        </button>
       </div>` : '';

  if (!chats?.length) {
    body.innerHTML = profileHtml + `<div style="text-align:center;padding:40px 20px;color:var(--muted);font-size:13px"><i class="ti ti-messages" style="font-size:28px;display:block;margin-bottom:8px"></i>No recent chats found</div>`;
    return;
  }

  body.innerHTML = profileHtml + `
    <div style="border:1px solid var(--border);border-radius:var(--r-sm);overflow:hidden;overflow-y:auto;max-height:calc(100vh - 280px)">
      ${chats.map(c => {
        const ts = c.lastMessage?.timestamp
          ? new Date(c.lastMessage.timestamp * 1000).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
          : '';
        const lastMsg = c.lastMessage?.body || '';
        const prefix = c.lastMessage?.fromMe ? '↪ ' : '';
        return `<div style="display:flex;align-items:center;gap:12px;padding:12px 14px;border-bottom:1px solid var(--border);cursor:default;transition:background .12s"
          onmouseenter="this.style.background='var(--surface2)'" onmouseleave="this.style.background='transparent'">
          <div style="width:42px;height:42px;border-radius:50%;background:${c.isGroup ? 'rgba(37,211,102,.12)' : 'rgba(66,133,244,.1)'};display:flex;align-items:center;justify-content:center;flex-shrink:0">
            <i class="ti ${c.isGroup ? 'ti-users' : 'ti-user'}" style="font-size:18px;color:${c.isGroup ? '#25D366' : '#4285F4'}"></i>
          </div>
          <div style="flex:1;min-width:0">
            <div style="display:flex;justify-content:space-between;align-items:baseline">
              <div style="font-size:13px;font-weight:${c.unreadCount ? '600' : '500'};color:var(--text);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${wEsc(c.name)}</div>
              <div style="font-size:11px;color:var(--muted);white-space:nowrap;margin-left:8px">${ts}</div>
            </div>
            <div style="font-size:12px;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${wEsc(prefix + lastMsg) || 'No messages'}</div>
          </div>
          ${c.unreadCount ? `<div style="min-width:18px;height:18px;border-radius:9px;background:#25D366;color:#fff;font-size:10px;font-weight:700;display:flex;align-items:center;justify-content:center;padding:0 4px">${c.unreadCount}</div>` : ''}
        </div>`;
      }).join('')}
    </div>`;
}

// ── Utilities ─────────────────────────────────────────────────────────────────

function _wqs(sel) { return document.querySelector(sel); }

function wEsc(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

async function waFetch(url, method = 'GET', body) {
  const opts = { method, headers: {} };
  if (body) { opts.headers['Content-Type'] = 'application/json'; opts.body = JSON.stringify(body); }
  const r = await fetch(url, opts);
  if (!r.ok) { const e = await r.json().catch(() => ({ error: r.statusText })); throw new Error(e.error || r.statusText); }
  return r.json();
}
