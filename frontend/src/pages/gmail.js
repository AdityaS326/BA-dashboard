// frontend/src/pages/gmail.js
// Gmail OAuth + Inbox panel logic

const BASE = '';  // same-origin relative URLs

let _gm = {
  auth: false,
  email: '',
  name: '',
  activeLabel: 'INBOX',
};

// ── Called by nav() when Gmail panel activates ─────────────────────────────────
export async function gmailInit() {
  // Check for auth return params in URL (e.g. /?panel=gm&auth=error)
  const params = new URLSearchParams(window.location.search);
  const authResult = params.get('auth');
  const authMsg    = params.get('msg');
  if (authResult) window.history.replaceState({}, '', '/');  // clean URL

  if (authResult === 'cancelled') {
    _gmRenderSignIn('Sign-in was cancelled. Try again when ready.');
    return;
  }
  if (authResult === 'error') {
    _gmRenderSignIn(authMsg || 'Authentication failed. Check your credentials.');
    return;
  }

  try {
    const data = await apiFetch('/api/gmail/status');
    if (data.authenticated) {
      _gm.auth  = true;
      _gm.email = data.email;
      _gm.name  = data.name;
      _gmRenderAuth();
      gmailLoadEmails('INBOX');
    } else {
      _gmRenderSignIn(null);
    }
  } catch {
    _gmRenderSignIn('Backend not reachable. Make sure the server is running.');
  }
}

// ── Sign-in: same-tab redirect (no popup, no dialog) ─────────────────────────
window.gmailSignIn = async function () {
  try {
    const data = await apiFetch('/api/gmail/auth-url');
    // Navigate the current tab — Google redirects back to /?panel=gm
    window.location.href = data.url;
  } catch (err) {
    // Credentials not configured — show setup guide inside the panel
    _gmRenderSetupGuide(err.message);
  }
};

// ── Load emails by label ──────────────────────────────────────────────────────
window.gmailLoadEmails = async function (label = 'INBOX') {
  _gm.activeLabel = label;

  document.querySelectorAll('.gm-tab').forEach(t =>
    t.style.cssText = tabStyle(t.dataset.label === label));

  const listEl = _qs('#gm-email-list');
  if (listEl) listEl.innerHTML = loadingHtml('Loading emails…');

  // Hide viewer when switching labels
  const viewer = _qs('#gm-email-viewer');
  if (viewer) viewer.style.display = 'none';

  try {
    const { emails } = await apiFetch(`/api/gmail/emails?label=${label}`);
    _renderEmailList(emails || []);
  } catch (err) {
    if (listEl) listEl.innerHTML = errorHtml('Failed to load emails: ' + err.message);
  }
};

// ── View full email ────────────────────────────────────────────────────────────
window.gmailViewEmail = async function (id) {
  const viewer = _qs('#gm-email-viewer');
  if (!viewer) return;
  viewer.style.display = 'flex';
  viewer.innerHTML = loadingHtml('Loading email…');

  try {
    const email = await apiFetch(`/api/gmail/email/${id}`);
    viewer.innerHTML = `
      <div style="display:flex;flex-direction:column;height:100%">
        <div style="padding:14px 16px;border-bottom:1px solid var(--border);flex-shrink:0">
          <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:8px">
            <h3 style="margin:0;font-size:14px;font-weight:600;color:var(--text);line-height:1.4;flex:1">${esc(email.subject)}</h3>
            <button onclick="document.getElementById('gm-email-viewer').style.display='none'"
              style="background:none;border:none;cursor:pointer;color:var(--muted);font-size:18px;line-height:1;padding:0;flex-shrink:0">✕</button>
          </div>
          <div style="margin-top:8px;font-size:12px;color:var(--muted);display:flex;flex-direction:column;gap:2px">
            <div><strong>From:</strong> ${esc(email.from)}</div>
            <div><strong>To:</strong> ${esc(email.to)}</div>
            <div><strong>Date:</strong> ${esc(email.date)}</div>
          </div>
        </div>
        <div style="flex:1;overflow:auto;padding:0">
          <iframe sandbox="allow-same-origin allow-popups" srcdoc="${esc(email.body)}"
            style="width:100%;height:100%;min-height:300px;border:none;display:block"
            onload="this.style.height=Math.max(300,this.contentDocument.body.scrollHeight+32)+'px'"></iframe>
        </div>
      </div>`;
  } catch (err) {
    viewer.innerHTML = errorHtml('Error loading email: ' + err.message);
  }
};

// ── Compose ───────────────────────────────────────────────────────────────────
window.gmailCompose = function () {
  const m = _qs('#gm-compose-modal');
  if (m) { m.style.display = 'flex'; _qs('#gm-compose-to')?.focus(); }
};
window.gmailCloseCompose = function () {
  const m = _qs('#gm-compose-modal');
  if (m) m.style.display = 'none';
};
window.gmailSend = async function () {
  const to      = _qs('#gm-compose-to')?.value?.trim();
  const subject = _qs('#gm-compose-subject')?.value?.trim();
  const body    = _qs('#gm-compose-body')?.value?.trim();
  if (!to || !subject || !body) return showToast('Fill in To, Subject, and Body', 'warning');

  const btn = _qs('#gm-send-btn');
  if (btn) { btn.disabled = true; btn.textContent = 'Sending…'; }

  try {
    await apiFetch('/api/gmail/send', 'POST', { to, subject, body });
    showToast('Email sent!', 'success');
    gmailCloseCompose();
    ['gm-compose-to', 'gm-compose-subject', 'gm-compose-body'].forEach(id => {
      const el = _qs('#' + id); if (el) el.value = '';
    });
  } catch (err) {
    showToast('Send failed: ' + err.message, 'error');
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = 'Send'; }
  }
};

// ── Logout ────────────────────────────────────────────────────────────────────
window.gmailLogout = async function () {
  await apiFetch('/api/gmail/logout', 'POST').catch(() => {});
  _gm = { auth: false, email: '', name: '', activeLabel: 'INBOX' };
  _gmRenderSignIn(null);
  showToast('Signed out from Gmail', 'success');
};

// ── Render helpers ─────────────────────────────────────────────────────────────

function _gmRenderSetupGuide(serverMsg) {
  const body = _qs('#gm-dynamic-body');
  if (!body) return;
  body.innerHTML = `
    <div style="max-width:560px;margin:0 auto;padding:32px 16px">
      <div style="display:flex;align-items:center;gap:12px;padding:14px 16px;background:rgba(234,67,53,.06);border:1px solid rgba(234,67,53,.2);border-radius:var(--r-sm);margin-bottom:20px">
        <i class="ti ti-alert-circle" style="color:#EA4335;font-size:20px;flex-shrink:0"></i>
        <div style="font-size:13px;color:var(--text)"><strong>Google OAuth credentials not configured.</strong><br>
        <span style="color:var(--muted)">${esc(serverMsg || 'Set GMAIL_CLIENT_ID and GMAIL_CLIENT_SECRET in backend/.env')}</span></div>
      </div>

      <div style="font-size:14px;font-weight:600;color:var(--text);margin-bottom:14px">
        <i class="ti ti-tool" style="color:#4285F4"></i> One-time setup — 3 steps
      </div>

      ${[
        { n:1, title:'Go to Google Cloud Console', sub:'console.cloud.google.com → APIs & Services → Credentials', link:'https://console.cloud.google.com/apis/credentials' },
        { n:2, title:'Create OAuth 2.0 Client ID', sub:'Application type: <strong>Web application</strong><br>Authorised redirect URI: <code>http://localhost:3000/api/gmail/callback</code>' },
        { n:3, title:'Enable the Gmail API', sub:'APIs & Services → Library → search <strong>Gmail API</strong> → Enable' },
      ].map(s => `
        <div style="display:flex;gap:12px;margin-bottom:12px;padding:12px 14px;background:var(--surface2);border:1px solid var(--border);border-radius:var(--r-sm)">
          <div style="width:26px;height:26px;border-radius:50%;background:#4285F4;color:#fff;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0">${s.n}</div>
          <div>
            <div style="font-size:13px;font-weight:600;color:var(--text);margin-bottom:2px">
              ${s.link ? `<a href="${s.link}" target="_blank" rel="noopener" style="color:#4285F4;text-decoration:none">${s.title} ↗</a>` : s.title}
            </div>
            <div style="font-size:12px;color:var(--muted);line-height:1.5">${s.sub}</div>
          </div>
        </div>`).join('')}

      <div style="margin:16px 0;padding:12px 16px;background:var(--surface2);border:1px solid var(--border);border-radius:var(--r-sm)">
        <div style="font-size:12px;font-weight:600;color:var(--muted);margin-bottom:6px">Then add to <code>backend/.env</code>:</div>
        <pre style="margin:0;font-size:12px;color:var(--text);white-space:pre-wrap">GMAIL_CLIENT_ID=your-client-id.apps.googleusercontent.com
GMAIL_CLIENT_SECRET=your-client-secret
GMAIL_REDIRECT_URI=http://localhost:3000/api/gmail/callback</pre>
      </div>

      <div style="font-size:12px;color:var(--muted);margin-bottom:16px">Restart the backend after saving .env, then:</div>
      <button onclick="gmailSignIn()" style="background:#4285F4;color:#fff;border:none;border-radius:8px;padding:10px 22px;font-size:13px;font-weight:600;cursor:pointer;display:flex;align-items:center;gap:6px">
        <i class="ti ti-refresh" style="font-size:14px"></i> Try signing in again
      </button>
    </div>`;
}

function _gmRenderSignIn(warning) {
  const body = _qs('#gm-dynamic-body');
  if (!body) return;
  body.innerHTML = `
    <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;padding:56px 24px;text-align:center">
      <div style="width:80px;height:80px;background:#fff;border-radius:20px;display:flex;align-items:center;justify-content:center;box-shadow:0 4px 20px rgba(0,0,0,.10);margin-bottom:24px">
        <i class="ti ti-brand-gmail" style="font-size:48px;color:#EA4335"></i>
      </div>
      <h2 style="font-size:20px;font-weight:700;color:var(--text);margin:0 0 8px">Sign in to Gmail</h2>
      <p style="font-size:13px;color:var(--muted);max-width:320px;margin:0 0 28px;line-height:1.55">
        Connect your Google account to read and send emails directly from this dashboard.
      </p>
      ${warning ? `<div style="padding:10px 14px;background:rgba(251,188,5,.08);border:1px solid rgba(251,188,5,.3);border-radius:var(--r-sm);font-size:12px;color:var(--muted);margin-bottom:20px;max-width:360px;text-align:left">
        <i class="ti ti-alert-triangle" style="color:#FBBC05"></i> ${esc(warning)}</div>` : ''}
      <button onclick="gmailSignIn()"
        style="display:flex;align-items:center;gap:10px;background:#fff;color:#3c4043;border:1px solid #dadce0;border-radius:8px;padding:12px 26px;font-size:14px;font-weight:500;cursor:pointer;box-shadow:0 1px 4px rgba(0,0,0,.08);transition:box-shadow .15s"
        onmouseenter="this.style.boxShadow='0 2px 10px rgba(0,0,0,.14)'"
        onmouseleave="this.style.boxShadow='0 1px 4px rgba(0,0,0,.08)'">
        <svg width="18" height="18" viewBox="0 0 18 18"><path d="M17.64 9.2c0-.637-.057-1.251-.164-1.84H9v3.481h4.844c-.209 1.125-.843 2.078-1.796 2.717v2.258h2.908c1.702-1.567 2.684-3.875 2.684-6.615z" fill="#4285F4"/><path d="M9 18c2.43 0 4.467-.806 5.956-2.18l-2.908-2.259c-.806.54-1.837.86-3.048.86-2.344 0-4.328-1.584-5.036-3.711H.957v2.332A8.997 8.997 0 0 0 9 18z" fill="#34A853"/><path d="M3.964 10.71A5.41 5.41 0 0 1 3.682 9c0-.593.102-1.17.282-1.71V4.958H.957A8.996 8.996 0 0 0 0 9c0 1.452.348 2.827.957 4.042l3.007-2.332z" fill="#FBBC05"/><path d="M9 3.58c1.321 0 2.508.454 3.44 1.345l2.582-2.58C13.463.891 11.426 0 9 0A8.997 8.997 0 0 0 .957 4.958L3.964 6.29C4.672 4.163 6.656 3.58 9 3.58z" fill="#EA4335"/></svg>
        Sign in with Google
      </button>
      <p style="font-size:11px;color:var(--muted);margin-top:18px;line-height:1.5">
        Needs <code>GMAIL_CLIENT_ID</code> &amp; <code>GMAIL_CLIENT_SECRET</code> in <code>.env</code>.<br>
        <a href="https://console.cloud.google.com/apis/credentials" target="_blank" rel="noopener" style="color:#4285F4">
          Set up at Google Cloud Console →
        </a>
      </p>
    </div>`;
}

function _gmRenderAuth() {
  const body = _qs('#gm-dynamic-body');
  if (!body) return;

  const labels = [
    { id: 'INBOX',   icon: 'ti-inbox',     label: 'Inbox'   },
    { id: 'SENT',    icon: 'ti-send',      label: 'Sent'    },
    { id: 'DRAFT',   icon: 'ti-file-text', label: 'Drafts'  },
    { id: 'STARRED', icon: 'ti-star',      label: 'Starred' },
    { id: 'SPAM',    icon: 'ti-alert-circle', label: 'Spam' },
  ];

  body.innerHTML = `
    <!-- Account + toolbar row -->
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;flex-wrap:wrap">
      <div style="display:flex;align-items:center;gap:6px;padding:6px 10px;background:rgba(66,133,244,.06);border:1px solid rgba(66,133,244,.15);border-radius:var(--r-sm);font-size:12px;color:var(--muted)">
        <i class="ti ti-user-check" style="color:#4285F4;font-size:14px"></i>
        <span style="color:var(--text);font-weight:500">${esc(_gm.email)}</span>
      </div>
      <div style="flex:1"></div>
      <button onclick="gmailCompose()"
        style="background:#EA4335;color:#fff;border:none;border-radius:var(--r-sm);padding:6px 14px;font-size:12px;font-weight:600;cursor:pointer;display:flex;align-items:center;gap:5px">
        <i class="ti ti-pencil" style="font-size:13px"></i> Compose
      </button>
      <button onclick="gmailLogout()" title="Sign out"
        style="background:none;border:1px solid var(--border);border-radius:var(--r-sm);padding:6px 10px;cursor:pointer;color:var(--muted);font-size:12px;display:flex;align-items:center;gap:4px"
        onmouseenter="this.style.color='var(--text)'" onmouseleave="this.style.color='var(--muted)'">
        <i class="ti ti-logout" style="font-size:13px"></i> Sign out
      </button>
    </div>

    <!-- Label tabs -->
    <div style="display:flex;gap:4px;margin-bottom:10px;overflow-x:auto;padding-bottom:2px">
      ${labels.map(l => `
        <button class="gm-tab" data-label="${l.id}" onclick="gmailLoadEmails('${l.id}')"
          style="${tabStyle(l.id === 'INBOX')}">
          <i class="ti ${l.icon}" style="font-size:12px"></i> ${l.label}
        </button>`).join('')}
    </div>

    <!-- Split: list + viewer -->
    <div style="display:flex;gap:10px;height:calc(100vh - 290px);min-height:280px">
      <div id="gm-email-list"
        style="width:320px;flex-shrink:0;overflow-y:auto;border:1px solid var(--border);border-radius:var(--r-sm);background:var(--surface)">
        ${loadingHtml('Loading emails…')}
      </div>
      <div id="gm-email-viewer"
        style="flex:1;overflow:hidden;border:1px solid var(--border);border-radius:var(--r-sm);background:var(--surface);display:none;flex-direction:column">
      </div>
    </div>

    <!-- Compose modal -->
    <div id="gm-compose-modal"
      style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:2000;align-items:flex-end;justify-content:flex-end;padding:20px">
      <div style="background:var(--surface);border-radius:var(--r);box-shadow:0 8px 40px rgba(0,0,0,.25);width:500px;max-width:calc(100vw - 40px);overflow:hidden">
        <div style="background:#EA4335;padding:12px 16px;display:flex;align-items:center;justify-content:space-between">
          <span style="color:#fff;font-size:14px;font-weight:600;display:flex;align-items:center;gap:8px">
            <i class="ti ti-mail" style="font-size:16px"></i> New Message
          </span>
          <button onclick="gmailCloseCompose()" style="background:none;border:none;color:rgba(255,255,255,.8);cursor:pointer;font-size:20px;line-height:1;padding:0">✕</button>
        </div>
        <input id="gm-compose-to"      type="email" placeholder="To"      style="${composeInputStyle()}">
        <input id="gm-compose-subject" type="text"  placeholder="Subject" style="${composeInputStyle()}border-top:none">
        <textarea id="gm-compose-body" placeholder="Write your message…" style="width:100%;height:180px;padding:12px 16px;border:none;border-top:1px solid var(--border);background:transparent;font-size:13px;color:var(--text);resize:vertical;box-sizing:border-box;outline:none;font-family:inherit"></textarea>
        <div style="padding:10px 16px;border-top:1px solid var(--border);display:flex;justify-content:flex-end;gap:8px">
          <button onclick="gmailCloseCompose()" style="background:var(--surface2);color:var(--text);border:1px solid var(--border);border-radius:var(--r-sm);padding:8px 16px;font-size:13px;cursor:pointer">Discard</button>
          <button id="gm-send-btn" onclick="gmailSend()" style="background:#EA4335;color:#fff;border:none;border-radius:var(--r-sm);padding:8px 22px;font-size:13px;font-weight:600;cursor:pointer">Send</button>
        </div>
      </div>
    </div>`;
}

function _renderEmailList(emails) {
  const listEl = _qs('#gm-email-list');
  if (!listEl) return;

  if (!emails.length) {
    listEl.innerHTML = '<div style="text-align:center;padding:40px 20px;color:var(--muted);font-size:13px">No emails found</div>';
    return;
  }

  listEl.innerHTML = emails.map(e => {
    const from = e.from?.replace(/<[^>]+>/g, '').replace(/"/g, '').trim() || e.from;
    const date = fmtDate(e.date);
    const unread = e.isUnread;
    return `<div onclick="gmailViewEmail('${e.id}')"
      style="padding:11px 14px;border-bottom:1px solid var(--border);cursor:pointer;background:${unread ? 'rgba(66,133,244,.04)' : 'transparent'};transition:background .12s"
      onmouseenter="this.style.background='var(--surface2)'"
      onmouseleave="this.style.background='${unread ? 'rgba(66,133,244,.04)' : 'transparent'}'">
      <div style="display:flex;justify-content:space-between;align-items:baseline;gap:6px;margin-bottom:1px">
        <div style="font-size:13px;font-weight:${unread ? '600' : '400'};color:var(--text);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1">${esc(from)}</div>
        <div style="font-size:11px;color:var(--muted);white-space:nowrap;flex-shrink:0">${date}</div>
      </div>
      <div style="font-size:12px;font-weight:${unread ? '600' : '400'};color:var(--text);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;margin-bottom:1px">${esc(e.subject)}</div>
      <div style="font-size:11px;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(e.snippet)}</div>
    </div>`;
  }).join('');
}

// ── Micro-utilities ────────────────────────────────────────────────────────────

function _qs(sel) { return document.querySelector(sel); }

function esc(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function fmtDate(str) {
  if (!str) return '';
  try {
    const d   = new Date(str);
    const now = new Date();
    const diff = Math.floor((now - d) / 86400000);
    if (diff === 0) return d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    if (diff < 7)   return d.toLocaleDateString([], { weekday: 'short' });
    return d.toLocaleDateString([], { month: 'short', day: 'numeric' });
  } catch { return ''; }
}

function tabStyle(active) {
  return `background:${active ? '#EA4335' : 'var(--surface2)'};color:${active ? '#fff' : 'var(--text)'};border:1px solid ${active ? '#EA4335' : 'var(--border)'};border-radius:var(--r-sm);padding:5px 11px;font-size:12px;font-weight:500;cursor:pointer;white-space:nowrap;display:inline-flex;align-items:center;gap:4px;transition:all .15s`;
}

function composeInputStyle() {
  return 'width:100%;padding:10px 16px;border:none;border-bottom:1px solid var(--border);background:transparent;font-size:13px;color:var(--text);box-sizing:border-box;outline:none;font-family:inherit;';
}

function loadingHtml(msg) {
  return `<div style="text-align:center;padding:40px 20px;color:var(--muted)"><i class="ti ti-loader-2" style="font-size:22px;animation:spin 1s linear infinite"></i><div style="margin-top:8px;font-size:13px">${msg}</div></div>`;
}

function errorHtml(msg) {
  return `<div style="text-align:center;padding:40px 20px;color:var(--muted);font-size:13px"><i class="ti ti-alert-circle" style="font-size:22px;color:#EA4335;display:block;margin-bottom:8px"></i>${esc(msg)}</div>`;
}

async function apiFetch(url, method = 'GET', body) {
  const opts = { method, headers: {} };
  if (body) { opts.headers['Content-Type'] = 'application/json'; opts.body = JSON.stringify(body); }
  const r = await fetch(url, opts);
  if (!r.ok) { const e = await r.json().catch(() => ({ error: r.statusText })); throw new Error(e.error || r.statusText); }
  return r.json();
}
