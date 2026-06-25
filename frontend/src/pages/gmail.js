// frontend/src/pages/gmail.js — Gmail OAuth 2.0 with full inbox

let _gm = { auth: false, email: '', name: '', picture: '', label: 'INBOX' };

// ── Entry point ──────────────────────────────────────────────────────────────
export async function gmailInit() {
  // Handle return from Google OAuth redirect
  const p   = new URLSearchParams(window.location.search);
  const ret = p.get('auth');
  const msg = p.get('msg');
  if (ret) window.history.replaceState({}, '', '/');

  try {
    const s = await _gFetch('/api/gmail/status');
    if (s.authenticated) {
      _gm = { auth: true, email: s.email, name: s.name, picture: s.picture || '', label: 'INBOX' };
      _gmRenderInbox();
      gmailLoadEmails('INBOX');
      return;
    }
  } catch (e) {
    // 401 with scope error — clear state and show helpful message
    if (e.message?.includes('insufficient') || e.message?.includes('SCOPE')) {
      errMsg = 'Gmail permissions were not fully granted. Please sign in again and click "Allow" for all permissions.';
    }
  }

  let errMsg = ret === 'cancelled' ? 'Sign-in was cancelled. Try again.'
             : ret === 'error'     ? (msg || 'Authentication failed.')
             : null;
  _gmRenderLogin(errMsg);
}

// ── Show/hide password toggle (called from onclick) ──────────────────────────
window.gmailTogglePass = function (btn) {
  const inp = document.getElementById('gm-pass');
  if (!inp) return;
  inp.type = inp.type === 'password' ? 'text' : 'password';
  const ic = btn.querySelector('i');
  if (ic) ic.className = inp.type === 'password' ? 'ti ti-eye' : 'ti ti-eye-off';
};

// ── Login form ───────────────────────────────────────────────────────────────
function _gmRenderLogin(errMsg) {
  const el = document.getElementById('gm-dynamic-body');
  if (!el) return;
  el.innerHTML = `
  <div style="display:flex;align-items:center;justify-content:center;height:100%;padding:24px;overflow-y:auto">
    <div style="width:100%;max-width:400px;background:var(--surface);border:1px solid var(--border);border-radius:16px;overflow:hidden;box-shadow:0 6px 32px rgba(0,0,0,.1)">

      <!-- ── Header ── -->
      <div style="padding:32px 32px 24px;text-align:center;border-bottom:1px solid var(--border)">
        <div style="width:52px;height:52px;background:#fff;border-radius:14px;margin:0 auto 16px;box-shadow:0 2px 10px rgba(0,0,0,.14);display:flex;align-items:center;justify-content:center">
          <svg width="32" height="32" viewBox="0 0 48 48">
            <path d="M24 19.6c2.6 0 4.4 1.1 5.4 2l4-3.9C30.9 15.5 27.8 14 24 14c-5.3 0-9.9 3-12.2 7.5l4.7 3.6C17.8 21.5 20.7 19.6 24 19.6z" fill="#EA4335"/>
            <path d="M11.8 24c0-1 .14-1.98.4-2.9L8.1 17.5A14 14 0 0 0 7 24c0 2.3.57 4.5 1.6 6.4l4.7-3.6A8.1 8.1 0 0 1 11.8 24z" fill="#FBBC05"/>
            <path d="M24 34c-3.3 0-6.2-1.8-7.7-4.5l-4.7 3.6C13.9 37.2 18.7 40 24 40c3.6 0 6.8-1.3 9.2-3.5l-4.4-3.4C27.6 33.5 25.9 34 24 34z" fill="#34A853"/>
            <path d="M40.2 24c0-1-.12-2-.34-2.95H24v5.95h9.1c-.4 2-1.6 3.6-3.4 4.7l4.4 3.4C37.8 32.7 40.2 28.7 40.2 24z" fill="#4285F4"/>
          </svg>
        </div>
        <div style="font-size:22px;font-weight:700;color:var(--text);margin-bottom:4px">Sign in to Gmail</div>
        <div style="font-size:13px;color:var(--muted)">Use your Google Account</div>
      </div>

      <!-- ── Form ── -->
      <div style="padding:24px 32px 32px">
        ${errMsg ? `<div style="padding:10px 14px;background:rgba(234,67,53,.07);border:1px solid rgba(234,67,53,.2);border-radius:9px;font-size:12px;color:#c62828;margin-bottom:18px;display:flex;align-items:flex-start;gap:8px;line-height:1.5">
          <i class="ti ti-alert-circle" style="font-size:15px;margin-top:1px;flex-shrink:0"></i>${_ge(errMsg)}</div>` : ''}

        <!-- Email -->
        <div style="margin-bottom:14px">
          <label style="display:block;font-size:11px;font-weight:700;color:var(--muted);letter-spacing:.6px;margin-bottom:7px">EMAIL OR PHONE</label>
          <input id="gm-email" type="email" placeholder="you@gmail.com" autocomplete="email"
            style="width:100%;padding:12px 16px;border:1.5px solid var(--border);border-radius:9px;font-size:14px;color:var(--text);background:var(--surface);box-sizing:border-box;outline:none;transition:border-color .15s"
            onfocus="this.style.borderColor='#4285F4'" onblur="this.style.borderColor='var(--border)'"
            onkeydown="if(event.key==='Enter')document.getElementById('gm-pass').focus()">
        </div>

        <!-- Password (note: entered on Google's page after redirect) -->
        <div style="margin-bottom:10px">
          <label style="display:block;font-size:11px;font-weight:700;color:var(--muted);letter-spacing:.6px;margin-bottom:7px">PASSWORD</label>
          <div style="position:relative">
            <input id="gm-pass" type="password" placeholder="Your Google password" autocomplete="current-password"
              style="width:100%;padding:12px 46px 12px 16px;border:1.5px solid var(--border);border-radius:9px;font-size:14px;color:var(--text);background:var(--surface);box-sizing:border-box;outline:none;transition:border-color .15s"
              onfocus="this.style.borderColor='#4285F4'" onblur="this.style.borderColor='var(--border)'"
              onkeydown="if(event.key==='Enter')gmailSignIn()">
            <button onclick="gmailTogglePass(this)" type="button"
              style="position:absolute;right:12px;top:50%;transform:translateY(-50%);background:none;border:none;cursor:pointer;color:var(--muted);padding:4px;display:flex">
              <i class="ti ti-eye" style="font-size:18px"></i>
            </button>
          </div>
          <div style="margin-top:6px;font-size:11px;color:var(--muted);display:flex;align-items:center;gap:5px">
            <i class="ti ti-shield-lock" style="color:#34A853;font-size:12px"></i>
            Your password is entered securely on Google's page — never stored here.
          </div>
        </div>

        <div style="margin-bottom:22px">
          <a href="#" onclick="event.preventDefault()" style="font-size:12px;color:#4285F4;text-decoration:none">Forgot password?</a>
        </div>

        <!-- Sign in button -->
        <button id="gm-signin-btn" onclick="gmailSignIn()"
          style="width:100%;padding:13px;background:#4285F4;color:#fff;border:none;border-radius:9px;font-size:15px;font-weight:600;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:10px;transition:background .15s;box-shadow:0 2px 8px rgba(66,133,244,.3)"
          onmouseenter="this.style.background='#1a73e8';this.style.boxShadow='0 4px 14px rgba(26,115,232,.4)'"
          onmouseleave="this.style.background='#4285F4';this.style.boxShadow='0 2px 8px rgba(66,133,244,.3)'">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none">
            <path d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z" fill="#fff" opacity=".9"/>
            <path d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z" fill="#fff" opacity=".9"/>
            <path d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l3.66-2.84z" fill="#fff" opacity=".9"/>
            <path d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z" fill="#fff" opacity=".9"/>
          </svg>
          Sign in with Google
        </button>
      </div>
    </div>
  </div>`;
}

// ── Sign in — redirect to Google OAuth (same tab) ────────────────────────────
window.gmailSignIn = async function () {
  const email = (document.getElementById('gm-email')?.value || '').trim();
  const btn   = document.getElementById('gm-signin-btn');
  if (btn) { btn.disabled = true; btn.innerHTML = '<i class="ti ti-loader-2" style="animation:spin 1s linear infinite"></i> Redirecting to Google…'; }
  try {
    const hint = email ? `?hint=${encodeURIComponent(email)}` : '';
    const d    = await _gFetch(`/api/gmail/auth-url${hint}`);
    window.location.href = d.url;          // same-tab OAuth redirect
  } catch (err) {
    if (btn) { btn.disabled = false; btn.innerHTML = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none"><path d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z" fill="#fff" opacity=".9"/></svg> Sign in with Google'; }
    showToast(err.message, 'error');
  }
};

// ── Inbox layout ─────────────────────────────────────────────────────────────
function _gmRenderInbox() {
  const el = document.getElementById('gm-dynamic-body');
  if (!el) return;

  const avatar = _gm.picture
    ? `<img src="${_ge(_gm.picture)}" style="width:36px;height:36px;border-radius:50%;object-fit:cover;flex-shrink:0">`
    : `<div style="width:36px;height:36px;border-radius:50%;background:#4285F4;color:#fff;font-size:15px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0">${(_gm.name||_gm.email||'?')[0].toUpperCase()}</div>`;

  const LABELS = [
    { id:'INBOX',   icon:'ti-inbox',         label:'Inbox'   },
    { id:'SENT',    icon:'ti-send',           label:'Sent'    },
    { id:'DRAFT',   icon:'ti-file-text',      label:'Drafts'  },
    { id:'STARRED', icon:'ti-star',           label:'Starred' },
    { id:'SPAM',    icon:'ti-alert-octagon',  label:'Spam'    },
    { id:'TRASH',   icon:'ti-trash',          label:'Trash'   },
  ];

  // Outer div uses flex-direction:column + flex:1 so it stretches to fill
  // the full height of #gm-dynamic-body (which itself is height:100% of .panel-full)
  el.innerHTML = `
    <div style="display:flex;flex-direction:column;height:100%;padding:12px 16px 0;box-sizing:border-box">

      <!-- ── Account bar ── -->
      <div style="display:flex;align-items:center;gap:10px;padding:8px 12px;background:var(--surface2);border:1px solid var(--border);border-radius:var(--r-sm);flex-shrink:0;margin-bottom:8px">
        ${avatar}
        <div style="flex:1;min-width:0">
          <div style="font-size:13px;font-weight:600;color:var(--text);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${_ge(_gm.name)}</div>
          <div style="font-size:11px;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${_ge(_gm.email)}</div>
        </div>
        <button onclick="gmailCompose()"
          style="background:#EA4335;color:#fff;border:none;border-radius:var(--r-sm);padding:7px 14px;font-size:12px;font-weight:600;cursor:pointer;display:flex;align-items:center;gap:5px;white-space:nowrap;flex-shrink:0">
          <i class="ti ti-pencil" style="font-size:13px"></i> Compose
        </button>
        <button onclick="gmailLogout()"
          style="background:none;border:1px solid var(--border);border-radius:var(--r-sm);padding:7px 12px;cursor:pointer;color:var(--muted);font-size:12px;display:flex;align-items:center;gap:5px;white-space:nowrap;flex-shrink:0;transition:all .15s"
          onmouseenter="this.style.borderColor='#EA4335';this.style.color='#EA4335'"
          onmouseleave="this.style.borderColor='var(--border)';this.style.color='var(--muted)'">
          <i class="ti ti-logout" style="font-size:13px"></i> Sign out
        </button>
      </div>

      <!-- ── Label tabs ── -->
      <div style="display:flex;gap:4px;flex-shrink:0;margin-bottom:8px;overflow-x:auto;scrollbar-width:none">
        ${LABELS.map(l => `
          <button class="gm-tab" data-label="${l.id}" onclick="gmailLoadEmails('${l.id}')"
            style="${_tabStyle(l.id==='INBOX')}">
            <i class="ti ${l.icon}" style="font-size:12px"></i> ${l.label}
          </button>`).join('')}
      </div>

      <!-- ── Split pane — flex:1 fills ALL remaining height ── -->
      <div style="display:flex;gap:8px;flex:1;min-height:0;padding-bottom:12px">
        <div id="gm-list" style="width:300px;flex-shrink:0;overflow-y:auto;border:1px solid var(--border);border-radius:var(--r-sm);background:var(--surface)">
          ${_spin('Loading inbox…')}
        </div>
        <div id="gm-viewer" style="flex:1;overflow:hidden;border:1px solid var(--border);border-radius:var(--r-sm);background:var(--surface);display:none;flex-direction:column"></div>
      </div>
    </div>

    <!-- ── Compose modal ── -->
    <div id="gm-compose" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:2000;align-items:flex-end;justify-content:flex-end;padding:20px">
      <div style="background:var(--surface);border-radius:14px;box-shadow:0 10px 50px rgba(0,0,0,.28);width:520px;max-width:calc(100vw - 40px);overflow:hidden">
        <div style="background:linear-gradient(135deg,#EA4335,#c62828);padding:14px 18px;display:flex;align-items:center;justify-content:space-between">
          <span style="color:#fff;font-size:14px;font-weight:600;display:flex;align-items:center;gap:8px">
            <i class="ti ti-mail-forward" style="font-size:17px"></i> New Message
          </span>
          <button onclick="gmailCloseCompose()" style="background:none;border:none;color:rgba(255,255,255,.85);cursor:pointer;font-size:24px;line-height:1;padding:0">×</button>
        </div>
        <div style="padding:0">
          <input id="gm-to"      type="email" placeholder="To"
            style="width:100%;padding:11px 18px;border:none;border-bottom:1px solid var(--border);background:var(--surface);color:var(--text);font-size:13px;box-sizing:border-box;outline:none">
          <input id="gm-subject" type="text"  placeholder="Subject"
            style="width:100%;padding:11px 18px;border:none;border-bottom:1px solid var(--border);background:var(--surface);color:var(--text);font-size:13px;box-sizing:border-box;outline:none;font-weight:500">
          <textarea id="gm-body" placeholder="Write your message…"
            style="width:100%;height:200px;padding:14px 18px;border:none;background:var(--surface);color:var(--text);font-size:13px;resize:vertical;box-sizing:border-box;outline:none;font-family:inherit;line-height:1.6"></textarea>
        </div>
        <div style="padding:10px 18px;border-top:1px solid var(--border);display:flex;justify-content:space-between;align-items:center">
          <div style="font-size:11px;color:var(--muted)">From: ${_ge(_gm.email)}</div>
          <div style="display:flex;gap:8px">
            <button onclick="gmailCloseCompose()" style="background:var(--surface2);color:var(--text);border:1px solid var(--border);border-radius:var(--r-sm);padding:8px 16px;font-size:13px;cursor:pointer">Discard</button>
            <button id="gm-send-btn" onclick="gmailSend()" style="background:#EA4335;color:#fff;border:none;border-radius:var(--r-sm);padding:8px 24px;font-size:13px;font-weight:600;cursor:pointer;display:flex;align-items:center;gap:6px">
              <i class="ti ti-send" style="font-size:13px"></i> Send
            </button>
          </div>
        </div>
      </div>
    </div>`;
}

// ── Load email list ───────────────────────────────────────────────────────────
window.gmailLoadEmails = async function (label = 'INBOX') {
  _gm.label = label;
  document.querySelectorAll('.gm-tab').forEach(t => t.style.cssText = _tabStyle(t.dataset.label === label));
  const list   = document.getElementById('gm-list');
  const viewer = document.getElementById('gm-viewer');
  if (list)   list.innerHTML = _spin('Loading…');
  if (viewer) viewer.style.display = 'none';
  try {
    const data = await _gFetch(`/api/gmail/emails?label=${label}`);
    // Scope error — backend cleared the session, show login again
    if (data.error === 'SCOPE_ERROR') {
      _gm = { auth: false, email: '', name: '', picture: '', label: 'INBOX' };
      _gmRenderLogin(data.message);
      return;
    }
    if (data.error === 'API_DISABLED') {
      if (list) list.innerHTML = _scopeError(
        'Gmail API is not enabled',
        data.message,
        'https://console.cloud.google.com/apis/library/gmail.googleapis.com',
        'Enable Gmail API'
      );
      return;
    }
    _renderList(data.emails || []);
  } catch (e) {
    // HTTP 401 = scope/auth problem — force re-login
    if (e.message?.includes('SCOPE_ERROR') || e.message?.includes('insufficient')) {
      _gm = { auth: false, email: '', name: '', picture: '', label: 'INBOX' };
      _gmRenderLogin('Gmail permission not granted. Please sign in again and allow ALL Gmail permissions.');
      return;
    }
    if (list) list.innerHTML = _err(e.message);
  }
};

// ── Open an email ─────────────────────────────────────────────────────────────
window.gmailViewEmail = async function (id) {
  const viewer = document.getElementById('gm-viewer');
  if (!viewer) return;
  viewer.style.display = 'flex';
  viewer.innerHTML = _spin('Loading email…');
  try {
    const m = await _gFetch(`/api/gmail/email/${id}`);
    viewer.innerHTML = `
      <div style="display:flex;flex-direction:column;height:100%;overflow:hidden">
        <div style="padding:14px 18px;border-bottom:1px solid var(--border);flex-shrink:0">
          <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:8px;margin-bottom:10px">
            <h3 style="margin:0;font-size:15px;font-weight:700;color:var(--text);line-height:1.4;flex:1">${_ge(m.subject)}</h3>
            <button onclick="document.getElementById('gm-viewer').style.display='none'"
              style="background:none;border:none;cursor:pointer;color:var(--muted);font-size:22px;line-height:1;padding:0;flex-shrink:0">×</button>
          </div>
          <div style="font-size:12px;color:var(--muted);display:grid;grid-template-columns:auto 1fr;gap:2px 8px">
            <strong>From:</strong><span>${_ge(m.from)}</span>
            <strong>To:</strong>  <span>${_ge(m.to)}</span>
            <strong>Date:</strong><span>${_ge(m.date)}</span>
          </div>
        </div>
        <div style="flex:1;overflow:auto;background:#fff">
          <iframe sandbox="allow-same-origin allow-popups" srcdoc="${_ge(m.body)}"
            style="width:100%;min-height:300px;height:100%;border:none;display:block"
            onload="try{this.style.height=Math.max(300,this.contentDocument.body.scrollHeight+24)+'px'}catch{}"></iframe>
        </div>
      </div>`;
  } catch (e) {
    viewer.innerHTML = _err('Failed to load: ' + e.message);
  }
};

// ── Compose ───────────────────────────────────────────────────────────────────
window.gmailCompose = function () {
  const m = document.getElementById('gm-compose');
  if (m) { m.style.display = 'flex'; document.getElementById('gm-to')?.focus(); }
};
window.gmailCloseCompose = function () {
  const m = document.getElementById('gm-compose');
  if (m) m.style.display = 'none';
};
window.gmailSend = async function () {
  const to      = (document.getElementById('gm-to')?.value      || '').trim();
  const subject = (document.getElementById('gm-subject')?.value  || '').trim();
  const body    = (document.getElementById('gm-body')?.value     || '').trim();
  if (!to)      return showToast('Enter a recipient', 'warning');
  if (!subject) return showToast('Enter a subject',   'warning');
  if (!body)    return showToast('Write a message',   'warning');
  const btn = document.getElementById('gm-send-btn');
  if (btn) { btn.disabled = true; btn.innerHTML = '<i class="ti ti-loader-2" style="animation:spin 1s linear infinite"></i> Sending…'; }
  try {
    await _gFetch('/api/gmail/send', 'POST', { to, subject, body });
    showToast('Email sent successfully!', 'success');
    gmailCloseCompose();
    ['gm-to','gm-subject','gm-body'].forEach(id => { const e = document.getElementById(id); if (e) e.value = ''; });
  } catch (e) {
    showToast('Send failed: ' + e.message, 'error');
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ti ti-send" style="font-size:13px"></i> Send'; }
  }
};

// ── Logout ────────────────────────────────────────────────────────────────────
window.gmailLogout = async function () {
  await _gFetch('/api/gmail/logout', 'POST').catch(() => {});
  _gm = { auth: false, email: '', name: '', picture: '', label: 'INBOX' };
  showToast('Signed out from Gmail', 'success');
  _gmRenderLogin(null);
};

// ── Internal helpers ──────────────────────────────────────────────────────────
function _renderList(emails) {
  const el = document.getElementById('gm-list');
  if (!el) return;
  if (!emails.length) { el.innerHTML = '<div style="text-align:center;padding:48px 16px;color:var(--muted);font-size:13px"><i class="ti ti-inbox" style="font-size:28px;display:block;margin-bottom:8px;opacity:.4"></i>No emails here</div>'; return; }
  el.innerHTML = emails.map(e => {
    const from = (e.from || '').replace(/<[^>]+>/g,'').replace(/"/g,'').trim() || 'Unknown';
    const date = _fmtDate(e.date);
    const bold = e.isUnread ? '600' : '400';
    const bg   = e.isUnread ? 'rgba(66,133,244,.04)' : 'transparent';
    return `<div onclick="gmailViewEmail('${e.id}')"
      style="padding:11px 14px;border-bottom:1px solid var(--border);cursor:pointer;background:${bg};transition:background .1s"
      onmouseenter="this.style.background='var(--surface2)'"
      onmouseleave="this.style.background='${bg}'">
      <div style="display:flex;justify-content:space-between;gap:6px;align-items:baseline">
        <div style="font-size:13px;font-weight:${bold};color:var(--text);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1">${_ge(from)}</div>
        <div style="font-size:11px;color:var(--muted);white-space:nowrap">${date}</div>
      </div>
      <div style="font-size:12px;font-weight:${bold};color:var(--text);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;margin-top:2px">${_ge(e.subject)}</div>
      <div style="font-size:11px;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;margin-top:1px">${_ge(e.snippet)}</div>
    </div>`;
  }).join('');
}

function _ge(s) { return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

function _fmtDate(s) {
  if (!s) return '';
  try {
    const d = new Date(s), now = new Date(), diff = Math.floor((now-d)/86400000);
    if (diff === 0) return d.toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'});
    if (diff  < 7)  return d.toLocaleDateString([],{weekday:'short'});
    return d.toLocaleDateString([],{month:'short',day:'numeric'});
  } catch { return ''; }
}

function _tabStyle(active) {
  return `background:${active?'#EA4335':'var(--surface2)'};color:${active?'#fff':'var(--text)'};border:1px solid ${active?'#EA4335':'var(--border)'};border-radius:var(--r-sm);padding:5px 12px;font-size:12px;font-weight:500;cursor:pointer;white-space:nowrap;display:inline-flex;align-items:center;gap:4px;transition:all .15s`;
}

function _spin(msg) {
  return `<div style="display:flex;flex-direction:column;align-items:center;padding:48px 20px;color:var(--muted);gap:10px"><i class="ti ti-loader-2" style="font-size:24px;animation:spin 1s linear infinite"></i><span style="font-size:13px">${msg}</span></div>`;
}

function _err(msg) {
  return `<div style="text-align:center;padding:40px 20px;font-size:13px;color:var(--muted)"><i class="ti ti-alert-circle" style="font-size:24px;color:#EA4335;display:block;margin-bottom:8px"></i>${_ge(msg)}</div>`;
}

function _scopeError(title, detail, link, linkLabel) {
  return `<div style="padding:20px;font-size:13px">
    <div style="background:rgba(234,67,53,.07);border:1px solid rgba(234,67,53,.2);border-radius:10px;padding:16px">
      <div style="font-weight:700;color:#c62828;margin-bottom:8px;display:flex;align-items:center;gap:6px">
        <i class="ti ti-lock-open" style="font-size:16px"></i> ${_ge(title)}
      </div>
      <div style="color:var(--text);line-height:1.7;margin-bottom:12px">${_ge(detail)}</div>
      ${link ? `<a href="${link}" target="_blank"
        style="display:inline-flex;align-items:center;gap:6px;padding:8px 14px;background:#4285F4;color:#fff;border-radius:7px;text-decoration:none;font-size:12px;font-weight:600">
        <i class="ti ti-external-link" style="font-size:13px"></i> ${_ge(linkLabel)}
      </a>` : ''}
    </div>
    <div style="margin-top:14px;padding:12px;background:var(--surface2);border-radius:8px;font-size:12px;color:var(--muted);line-height:1.8">
      <strong style="color:var(--text)">After enabling, sign out and sign in again:</strong><br>
      1. Sign out using the button above<br>
      2. Click "Sign in with Google" again<br>
      3. When Google shows permissions — click <strong>Allow ALL</strong>
    </div>
    <button onclick="gmailLogout()" style="margin-top:12px;padding:8px 18px;background:#EA4335;color:#fff;border:none;border-radius:7px;font-size:13px;font-weight:600;cursor:pointer;display:flex;align-items:center;gap:6px">
      <i class="ti ti-logout" style="font-size:13px"></i> Sign out &amp; re-authorize
    </button>
  </div>`;
}

async function _gFetch(url, method = 'GET', body) {
  const opts = { method, headers: {} };
  if (body) { opts.headers['Content-Type'] = 'application/json'; opts.body = JSON.stringify(body); }
  const r = await fetch(url, opts);
  if (!r.ok) { const e = await r.json().catch(()=>({error:r.statusText})); throw new Error(e.error || r.statusText); }
  return r.json();
}
