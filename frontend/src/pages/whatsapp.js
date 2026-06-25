// frontend/src/pages/whatsapp.js — Meta WhatsApp Business Cloud API

let _waActivePoll  = null;
let _waActivePhone = null;  // currently open conversation

// ── Entry ────────────────────────────────────────────────────────────────────
export async function waInit() {
  _waStopPoll();
  _waActivePhone = null;
  const body = document.getElementById('wa-dynamic-body');
  if (!body) return;
  body.innerHTML = `<div style="display:flex;align-items:center;justify-content:center;height:100%;color:var(--muted)">
    <i class="ti ti-loader-2" style="animation:spin 1s linear infinite;font-size:28px"></i></div>`;

  try {
    const data = await _waApi('/api/whatsapp/status');
    if (data.connected) {
      _waRenderInbox(data.profile);
      _waStartPoll();
    } else {
      _waCheckCredentials();
    }
  } catch {
    _waCheckCredentials();
  }
}

async function _waCheckCredentials() {
  try {
    const r = await _waApi('/api/credentials?keys=WA_PHONE_NUMBER_ID,WA_ACCESS_TOKEN');
    if (r.WA_PHONE_NUMBER_ID_set && r.WA_ACCESS_TOKEN_set) {
      _waRenderSetup('Credentials saved but verification failed. Check they are correct.');
    } else {
      _waRenderSetup(null);
    }
  } catch {
    _waRenderSetup(null);
  }
}

// ── Polling ──────────────────────────────────────────────────────────────────
function _waStartPoll() {
  _waStopPoll();
  _waActivePoll = setInterval(async () => {
    try {
      const { conversations } = await _waApi('/api/whatsapp/conversations');
      _waUpdateConvList(conversations);
      if (_waActivePhone) {
        const { messages } = await _waApi('/api/whatsapp/messages/' + _waActivePhone);
        _waUpdateMessages(messages);
      }
    } catch { /* ignore */ }
  }, 4000);
}
function _waStopPoll() { if (_waActivePoll) { clearInterval(_waActivePoll); _waActivePoll = null; } }

// ── Save credentials ─────────────────────────────────────────────────────────
window.waSaveCredentials = async function () {
  const phoneId = document.getElementById('wa-phone-id')?.value?.trim();
  const token   = document.getElementById('wa-access-token')?.value?.trim();
  const verify  = document.getElementById('wa-verify-token')?.value?.trim();

  if (!phoneId || !token) {
    _waSetupError('Phone Number ID and Access Token are required.');
    return;
  }

  const btn = document.getElementById('wa-save-btn');
  if (btn) { btn.disabled = true; btn.textContent = 'Saving…'; }

  try {
    const body = { WA_PHONE_NUMBER_ID: phoneId, WA_ACCESS_TOKEN: token };
    if (verify) body.WA_WEBHOOK_VERIFY_TOKEN = verify;
    await _waApi('/api/credentials', 'POST', body);

    // Test connection immediately
    const status = await _waApi('/api/whatsapp/status');
    if (status.connected) {
      showToast('WhatsApp connected!', 'success');
      _waRenderInbox(status.profile);
      _waStartPoll();
    } else {
      _waSetupError('Credentials saved but connection failed: ' + (status.reason || 'unknown error'));
      if (btn) { btn.disabled = false; btn.textContent = 'Save & Connect'; }
    }
  } catch (err) {
    _waSetupError('Error: ' + err.message);
    if (btn) { btn.disabled = false; btn.textContent = 'Save & Connect'; }
  }
};

function _waSetupError(msg) {
  const el = document.getElementById('wa-setup-err');
  if (el) { el.textContent = msg; el.style.display = 'flex'; }
}

// ── Open conversation ────────────────────────────────────────────────────────
window.waOpenConv = async function (phone) {
  _waActivePhone = phone;

  // Highlight selected in sidebar
  document.querySelectorAll('[data-waphone]').forEach(el => {
    el.style.background = el.dataset.waphone === phone ? 'var(--surface2)' : 'transparent';
  });

  const viewer = document.getElementById('wa-viewer');
  if (!viewer) return;
  viewer.innerHTML = `<div style="display:flex;align-items:center;justify-content:center;height:100%;color:var(--muted)">
    <i class="ti ti-loader-2" style="animation:spin 1s linear infinite;font-size:22px"></i></div>`;

  try {
    const { messages } = await _waApi('/api/whatsapp/messages/' + phone);
    const name = document.querySelector(`[data-waphone="${phone}"]`)?.dataset.waname || ('+' + phone);
    _waRenderThread(phone, name, messages);
  } catch (err) {
    viewer.innerHTML = `<div style="padding:20px;color:#EA4335;font-size:13px">Failed to load messages: ${_waE(err.message)}</div>`;
  }
};

// ── Send message ─────────────────────────────────────────────────────────────
window.waSend = async function () {
  if (!_waActivePhone) return;
  const inp = document.getElementById('wa-msg-input');
  const msg = inp?.value?.trim();
  if (!msg) return;

  inp.value = '';
  inp.disabled = true;

  try {
    await _waApi('/api/whatsapp/send', 'POST', { to: _waActivePhone, message: msg });
    const { messages } = await _waApi('/api/whatsapp/messages/' + _waActivePhone);
    _waUpdateMessages(messages);
  } catch (err) {
    showToast('Send failed: ' + err.message, 'error');
    if (inp) inp.value = msg; // restore
  } finally {
    if (inp) { inp.disabled = false; inp.focus(); }
  }
};

// ── Disconnect ───────────────────────────────────────────────────────────────
window.waDisconnect = function () {
  _waStopPoll();
  _waActivePhone = null;
  _waRenderSetup('Disconnected. Re-enter your credentials to reconnect.');
};

// ── New conversation ─────────────────────────────────────────────────────────
window.waNewConv = function () {
  const viewer = document.getElementById('wa-viewer');
  if (!viewer) return;
  _waActivePhone = null;
  viewer.innerHTML = `
    <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;gap:16px;padding:32px">
      <i class="ti ti-message-plus" style="font-size:40px;color:#25D366"></i>
      <div style="font-size:15px;font-weight:500;color:var(--text)">Start a new conversation</div>
      <div style="font-size:12px;color:var(--muted);text-align:center;max-width:280px;line-height:1.6">
        Enter a phone number with country code to send a WhatsApp message
      </div>
      <div style="display:flex;gap:8px;width:100%;max-width:360px">
        <input id="wa-new-phone" type="tel" placeholder="919876543210"
          style="flex:1;padding:11px 13px;border:1.5px solid var(--border);border-radius:8px;font-size:14px;color:var(--text);background:var(--surface);outline:none;font-family:monospace"
          onfocus="this.style.borderColor='#25D366'" onblur="this.style.borderColor='var(--border)'"
          onkeydown="if(event.key==='Enter')waStartNew()">
        <button onclick="waStartNew()"
          style="padding:11px 16px;background:#25D366;color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer">
          Open
        </button>
      </div>
      <div style="font-size:11px;color:var(--muted)">No + or spaces — e.g. <code style="background:var(--surface2);padding:1px 5px;border-radius:3px">919876543210</code></div>
    </div>`;
  setTimeout(() => document.getElementById('wa-new-phone')?.focus(), 60);
};

window.waStartNew = function () {
  const phone = (document.getElementById('wa-new-phone')?.value || '').replace(/\D/g,'');
  if (phone.length < 7) return showToast('Enter a valid number with country code', 'warning');
  waOpenConv(phone);
};

// ═══════════════════════════════════════════════════════════════════════════════
// RENDER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

// ── Setup / credentials screen ───────────────────────────────────────────────
function _waRenderSetup(errMsg) {
  const body = document.getElementById('wa-dynamic-body');
  if (!body) return;

  body.innerHTML = `
    <div style="display:flex;height:100%;overflow-y:auto">
      <div style="flex:1;display:flex;align-items:flex-start;justify-content:center;padding:32px 20px;min-height:100%">
        <div style="width:100%;max-width:520px">

          <!-- Header -->
          <div style="display:flex;align-items:center;gap:14px;margin-bottom:28px">
            <div style="width:52px;height:52px;background:#25D366;border-radius:14px;display:flex;align-items:center;justify-content:center;flex-shrink:0">
              <i class="ti ti-brand-whatsapp" style="font-size:30px;color:#fff"></i>
            </div>
            <div>
              <div style="font-size:20px;font-weight:700;color:var(--text)">WhatsApp Business</div>
              <div style="font-size:12px;color:var(--muted)">Configure via Meta Cloud API</div>
            </div>
          </div>

          <!-- Error -->
          <div id="wa-setup-err" style="display:${errMsg?'flex':'none'};align-items:center;gap:8px;padding:12px;background:rgba(234,67,53,.06);border:1px solid rgba(234,67,53,.2);border-radius:8px;font-size:12px;color:#d93025;margin-bottom:20px">
            <i class="ti ti-alert-circle" style="flex-shrink:0"></i>
            <span>${_waE(errMsg || '')}</span>
          </div>

          <!-- Credentials form -->
          <div style="background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:24px;margin-bottom:20px">
            <div style="font-size:14px;font-weight:600;color:var(--text);margin-bottom:18px;display:flex;align-items:center;gap:8px">
              <i class="ti ti-key" style="color:#25D366"></i> API Credentials
            </div>

            <label style="font-size:12px;font-weight:500;color:var(--muted);display:block;margin-bottom:5px">
              Phone Number ID <span style="color:#EA4335">*</span>
            </label>
            <input id="wa-phone-id" type="text" placeholder="e.g. 123456789012345"
              style="width:100%;padding:10px 12px;border:1.5px solid var(--border);border-radius:8px;font-size:13px;color:var(--text);background:var(--surface);box-sizing:border-box;outline:none;font-family:monospace;margin-bottom:14px"
              onfocus="this.style.borderColor='#25D366'" onblur="this.style.borderColor='var(--border)'">

            <label style="font-size:12px;font-weight:500;color:var(--muted);display:block;margin-bottom:5px">
              Permanent Access Token <span style="color:#EA4335">*</span>
            </label>
            <div style="position:relative;margin-bottom:14px">
              <input id="wa-access-token" type="password" placeholder="EAAxxxxx…"
                style="width:100%;padding:10px 38px 10px 12px;border:1.5px solid var(--border);border-radius:8px;font-size:13px;color:var(--text);background:var(--surface);box-sizing:border-box;outline:none;font-family:monospace"
                onfocus="this.style.borderColor='#25D366'" onblur="this.style.borderColor='var(--border)'">
              <button onclick="waToggleToken()" title="Show/hide"
                style="position:absolute;right:8px;top:50%;transform:translateY(-50%);background:none;border:none;cursor:pointer;color:var(--muted);font-size:16px;padding:0;line-height:1"
                id="wa-tok-eye"><i class="ti ti-eye"></i></button>
            </div>

            <label style="font-size:12px;font-weight:500;color:var(--muted);display:block;margin-bottom:5px">
              Webhook Verify Token <span style="color:var(--muted)">(optional — any string)</span>
            </label>
            <input id="wa-verify-token" type="text" placeholder="ba_dashboard_verify"
              style="width:100%;padding:10px 12px;border:1.5px solid var(--border);border-radius:8px;font-size:13px;color:var(--text);background:var(--surface);box-sizing:border-box;outline:none;font-family:monospace;margin-bottom:6px"
              onfocus="this.style.borderColor='#25D366'" onblur="this.style.borderColor='var(--border)'">
            <div style="font-size:11px;color:var(--muted);margin-bottom:18px">Leave blank to use the default: <code style="background:var(--surface2);padding:1px 4px;border-radius:3px">ba_dashboard_verify</code></div>

            <button id="wa-save-btn" onclick="waSaveCredentials()"
              style="width:100%;padding:13px;background:#25D366;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:600;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:8px"
              onmouseenter="this.style.background='#1da851'" onmouseleave="this.style.background='#25D366'">
              <i class="ti ti-plug" style="font-size:16px"></i> Save & Connect
            </button>
          </div>

          <!-- Where to find credentials -->
          <div style="background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:24px;margin-bottom:20px">
            <div style="font-size:14px;font-weight:600;color:var(--text);margin-bottom:14px;display:flex;align-items:center;gap:8px">
              <i class="ti ti-help-circle" style="color:#4285F4"></i> Where to find these
            </div>
            <ol style="margin:0;padding-left:18px;display:flex;flex-direction:column;gap:10px;font-size:12px;color:var(--text);line-height:1.7">
              <li>Go to <strong>developers.facebook.com</strong> → <strong>My Apps</strong> → your app</li>
              <li>Sidebar: <strong>WhatsApp → API Setup</strong></li>
              <li><strong>Phone Number ID</strong> — shown under "From" phone number section</li>
              <li><strong>Access Token</strong> — copy the temporary token (for production: create a permanent System User token in <em>Business Settings → System Users</em>)</li>
              <li>Scroll to <strong>"Step 3: Configure webhooks"</strong> to set your callback URL</li>
            </ol>
          </div>

          <!-- Webhook setup instructions -->
          <div style="background:rgba(37,211,102,.04);border:1px solid rgba(37,211,102,.2);border-radius:12px;padding:24px">
            <div style="font-size:14px;font-weight:600;color:var(--text);margin-bottom:14px;display:flex;align-items:center;gap:8px">
              <i class="ti ti-webhook" style="color:#25D366"></i> Webhook Setup (to receive messages)
            </div>
            <div style="font-size:12px;color:var(--text);line-height:1.8">
              <div style="margin-bottom:10px">To receive incoming messages, configure your webhook in the Meta dashboard:</div>
              <div style="font-weight:600;margin-bottom:4px">1. Expose your local server with ngrok:</div>
              <div style="background:var(--surface2);border:1px solid var(--border);border-radius:6px;padding:10px 12px;font-family:monospace;font-size:12px;color:var(--text);margin-bottom:12px;overflow-x:auto">ngrok http 3000</div>
              <div style="font-weight:600;margin-bottom:4px">2. In Meta API Setup → Configure Webhooks, enter:</div>
              <div style="background:var(--surface2);border:1px solid var(--border);border-radius:6px;padding:10px 12px;font-family:monospace;font-size:12px;color:var(--text);margin-bottom:4px;overflow-x:auto;word-break:break-all">
                Callback URL: <strong>https://&lt;ngrok-id&gt;.ngrok.io/api/whatsapp/webhook</strong>
              </div>
              <div style="background:var(--surface2);border:1px solid var(--border);border-radius:6px;padding:10px 12px;font-family:monospace;font-size:12px;color:var(--text);margin-bottom:12px;overflow-x:auto">
                Verify Token: <strong>ba_dashboard_verify</strong> (or whatever you set above)
              </div>
              <div style="font-weight:600;margin-bottom:4px">3. Subscribe to the <strong>messages</strong> field under the <strong>whatsapp_business_account</strong> object.</div>
              <div style="font-size:11px;color:var(--muted);margin-top:10px">
                Note: For production, deploy the backend to a public URL (e.g. Render, Railway) instead of using ngrok.
              </div>
            </div>
          </div>

        </div>
      </div>
    </div>`;
}

window.waToggleToken = function () {
  const inp = document.getElementById('wa-access-token');
  const eye = document.getElementById('wa-tok-eye');
  if (!inp) return;
  const show = inp.type === 'password';
  inp.type = show ? 'text' : 'password';
  if (eye) eye.innerHTML = show ? '<i class="ti ti-eye-off"></i>' : '<i class="ti ti-eye"></i>';
};

// ── Main inbox layout ─────────────────────────────────────────────────────────
function _waRenderInbox(profile) {
  const body = document.getElementById('wa-dynamic-body');
  if (!body) return;

  const initials = profile ? (profile.name || '+')[0].toUpperCase() : '?';

  body.innerHTML = `
    <div style="display:flex;flex-direction:column;height:100%;overflow:hidden">

      <!-- Top bar -->
      <div style="display:flex;align-items:center;gap:10px;padding:10px 14px;background:var(--surface2);border-bottom:1px solid var(--border);flex-shrink:0">
        <div style="width:36px;height:36px;border-radius:50%;background:#25D366;color:#fff;font-size:15px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0">${_waE(initials)}</div>
        <div style="flex:1;min-width:0">
          <div style="font-size:13px;font-weight:600;color:var(--text)">${_waE(profile?.name || 'WhatsApp')}</div>
          <div style="font-size:11px;color:var(--muted)">${profile?.phone ? '+' + _waE(profile.phone) : 'Business Account'}</div>
        </div>
        <div style="display:flex;align-items:center;gap:4px;font-size:11px;color:#25D366;background:rgba(37,211,102,.08);border:1px solid rgba(37,211,102,.2);padding:4px 10px;border-radius:20px;white-space:nowrap;margin-right:4px">
          <i class="ti ti-circle-filled" style="font-size:8px"></i> Connected
        </div>
        <button onclick="waDisconnect()"
          style="background:none;border:1px solid var(--border);border-radius:8px;padding:6px 12px;cursor:pointer;font-size:12px;color:var(--muted);display:flex;align-items:center;gap:5px"
          onmouseenter="this.style.borderColor='#EA4335';this.style.color='#EA4335'"
          onmouseleave="this.style.borderColor='var(--border)';this.style.color='var(--muted)'">
          <i class="ti ti-settings" style="font-size:13px"></i> Settings
        </button>
      </div>

      <!-- Split pane -->
      <div style="display:flex;flex:1;min-height:0;overflow:hidden">

        <!-- Sidebar: conversations list -->
        <div style="width:300px;flex-shrink:0;border-right:1px solid var(--border);display:flex;flex-direction:column;overflow:hidden">
          <div style="padding:10px;border-bottom:1px solid var(--border);flex-shrink:0;display:flex;gap:8px">
            <div style="flex:1;position:relative">
              <i class="ti ti-search" style="position:absolute;left:9px;top:50%;transform:translateY(-50%);color:var(--muted);font-size:13px"></i>
              <input id="wa-search" type="text" placeholder="Search conversations…"
                style="width:100%;padding:8px 8px 8px 30px;border:1px solid var(--border);border-radius:8px;font-size:12px;color:var(--text);background:var(--surface);box-sizing:border-box;outline:none"
                oninput="waFilterConvs(this.value)">
            </div>
            <button onclick="waNewConv()" title="New conversation"
              style="padding:8px 10px;background:#25D366;color:#fff;border:none;border-radius:8px;cursor:pointer;font-size:14px;flex-shrink:0">
              <i class="ti ti-edit"></i>
            </button>
          </div>
          <div id="wa-conv-list" style="flex:1;overflow-y:auto">
            <div style="padding:32px;text-align:center;color:var(--muted);font-size:12px">
              <i class="ti ti-loader-2" style="animation:spin 1s linear infinite;font-size:20px;display:block;margin-bottom:8px"></i>
              Loading conversations…
            </div>
          </div>
        </div>

        <!-- Right: message thread -->
        <div id="wa-viewer" style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;min-width:0;background:var(--surface)">
          <div style="text-align:center;color:var(--muted)">
            <i class="ti ti-brand-whatsapp" style="font-size:48px;color:rgba(37,211,102,.3);display:block;margin-bottom:12px"></i>
            <div style="font-size:14px;font-weight:500;color:var(--text)">Select a conversation</div>
            <div style="font-size:12px;margin-top:4px">or start a new one with <button onclick="waNewConv()" style="background:none;border:none;color:#25D366;cursor:pointer;font-size:12px;font-weight:600;padding:0">+ New Message</button></div>
          </div>
        </div>

      </div>
    </div>`;

  // Load conversations immediately
  _waLoadConversations();
}

async function _waLoadConversations() {
  try {
    const { conversations } = await _waApi('/api/whatsapp/conversations');
    _waUpdateConvList(conversations);
  } catch {
    const el = document.getElementById('wa-conv-list');
    if (el) el.innerHTML = `<div style="padding:20px;text-align:center;font-size:12px;color:var(--muted)">No conversations yet.<br>Send a message to get started.</div>`;
  }
}

function _waUpdateConvList(conversations) {
  const el = document.getElementById('wa-conv-list');
  if (!el) return;

  if (!conversations || conversations.length === 0) {
    el.innerHTML = `<div style="padding:32px;text-align:center;font-size:12px;color:var(--muted)">
      <i class="ti ti-messages" style="font-size:24px;display:block;margin-bottom:8px"></i>
      No conversations yet.<br>Use the <strong>+</strong> button to send a message.
    </div>`;
    return;
  }

  el.innerHTML = conversations.map(c => {
    const last = c.lastMessage;
    const ts   = last?.timestamp ? _waFmtTime(last.timestamp) : '';
    const preview = last?.body ? _waE(last.body.slice(0, 48)) : '';
    const isActive = c.phone === _waActivePhone;
    const unreadBadge = c.unread > 0
      ? `<div style="min-width:18px;height:18px;border-radius:9px;background:#25D366;color:#fff;font-size:10px;font-weight:700;display:flex;align-items:center;justify-content:center;padding:0 4px">${c.unread}</div>`
      : '';
    return `<div data-waphone="${c.phone}" data-waname="${_waE(c.name)}"
      onclick="waOpenConv('${c.phone}')"
      style="display:flex;align-items:center;gap:10px;padding:12px 14px;cursor:pointer;border-bottom:1px solid var(--border);background:${isActive?'var(--surface2)':'transparent'}"
      onmouseenter="this.style.background='var(--surface2)'" onmouseleave="this.style.background='${isActive?'var(--surface2)':'transparent'}'">
      <div style="width:40px;height:40px;border-radius:50%;background:rgba(37,211,102,.15);display:flex;align-items:center;justify-content:center;flex-shrink:0;font-size:15px;font-weight:700;color:#25D366">
        ${_waE((c.name||'+')[0].toUpperCase())}
      </div>
      <div style="flex:1;min-width:0">
        <div style="display:flex;justify-content:space-between;align-items:baseline;gap:4px">
          <div style="font-size:13px;font-weight:${c.unread?'600':'500'};color:var(--text);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${_waE(c.name)}</div>
          <div style="font-size:10px;color:var(--muted);white-space:nowrap;flex-shrink:0">${ts}</div>
        </div>
        <div style="display:flex;justify-content:space-between;align-items:center;gap:4px">
          <div style="font-size:11px;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${preview || 'No messages'}</div>
          ${unreadBadge}
        </div>
      </div>
    </div>`;
  }).join('');
}

window.waFilterConvs = function (query) {
  const q = query.toLowerCase();
  document.querySelectorAll('[data-waphone]').forEach(el => {
    const name = (el.dataset.waname || '').toLowerCase();
    el.style.display = name.includes(q) || el.dataset.waphone.includes(q) ? '' : 'none';
  });
};

// ── Thread rendering ──────────────────────────────────────────────────────────
function _waRenderThread(phone, name, messages) {
  const viewer = document.getElementById('wa-viewer');
  if (!viewer) return;

  viewer.innerHTML = `
    <div style="display:flex;flex-direction:column;height:100%;width:100%">
      <!-- Thread header -->
      <div style="display:flex;align-items:center;gap:10px;padding:10px 16px;border-bottom:1px solid var(--border);flex-shrink:0;background:var(--surface2)">
        <div style="width:38px;height:38px;border-radius:50%;background:rgba(37,211,102,.15);display:flex;align-items:center;justify-content:center;font-size:15px;font-weight:700;color:#25D366">
          ${_waE(name[0].toUpperCase())}
        </div>
        <div style="flex:1">
          <div style="font-size:13px;font-weight:600;color:var(--text)">${_waE(name)}</div>
          <div style="font-size:11px;color:var(--muted)">+${_waE(phone)}</div>
        </div>
      </div>

      <!-- Messages area -->
      <div id="wa-thread-msgs" style="flex:1;overflow-y:auto;padding:16px;display:flex;flex-direction:column;gap:8px"></div>

      <!-- Input bar -->
      <div style="padding:10px 14px;border-top:1px solid var(--border);display:flex;gap:8px;align-items:flex-end;flex-shrink:0;background:var(--surface)">
        <textarea id="wa-msg-input" placeholder="Type a message…" rows="1"
          style="flex:1;padding:10px 12px;border:1.5px solid var(--border);border-radius:20px;font-size:13px;color:var(--text);background:var(--surface2);resize:none;outline:none;font-family:inherit;line-height:1.5;max-height:120px;overflow-y:auto"
          onfocus="this.style.borderColor='#25D366'" onblur="this.style.borderColor='var(--border)'"
          onkeydown="if(event.key==='Enter'&&!event.shiftKey){event.preventDefault();waSend()}"
          oninput="this.style.height='auto';this.style.height=Math.min(this.scrollHeight,120)+'px'"></textarea>
        <button onclick="waSend()"
          style="padding:10px 16px;background:#25D366;color:#fff;border:none;border-radius:20px;cursor:pointer;font-size:14px;display:flex;align-items:center;gap:6px;flex-shrink:0"
          onmouseenter="this.style.background='#1da851'" onmouseleave="this.style.background='#25D366'">
          <i class="ti ti-send"></i>
        </button>
      </div>
    </div>`;

  _waUpdateMessages(messages);
  setTimeout(() => document.getElementById('wa-msg-input')?.focus(), 60);
}

function _waUpdateMessages(messages) {
  const el = document.getElementById('wa-thread-msgs');
  if (!el) return;

  if (!messages || messages.length === 0) {
    el.innerHTML = `<div style="text-align:center;color:var(--muted);font-size:12px;margin:auto">No messages yet. Say hello!</div>`;
    return;
  }

  el.innerHTML = messages.map(m => {
    const out = m.direction === 'outbound';
    const ts  = _waFmtTime(m.timestamp);
    return `<div style="display:flex;justify-content:${out?'flex-end':'flex-start'}">
      <div style="max-width:70%;padding:9px 12px;border-radius:${out?'16px 16px 4px 16px':'16px 16px 16px 4px'};background:${out?'#25D366':'var(--surface2)'};color:${out?'#fff':'var(--text)'};font-size:13px;line-height:1.5;box-shadow:0 1px 3px rgba(0,0,0,.08)">
        <div>${_waE(m.body)}</div>
        <div style="font-size:10px;opacity:.65;margin-top:4px;text-align:right">${ts}</div>
      </div>
    </div>`;
  }).join('');

  // Scroll to bottom
  el.scrollTop = el.scrollHeight;
}

// ── Utilities ─────────────────────────────────────────────────────────────────
function _waFmtTime(ts) {
  const d = new Date(ts * 1000);
  const now = new Date();
  const sameDay = d.toDateString() === now.toDateString();
  if (sameDay) return d.toLocaleTimeString([], {hour:'2-digit',minute:'2-digit'});
  return d.toLocaleDateString([], {month:'short',day:'numeric'});
}

function _waE(s) {
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

async function _waApi(url, method = 'GET', body) {
  const opts = { method, headers: {} };
  if (body) { opts.headers['Content-Type'] = 'application/json'; opts.body = JSON.stringify(body); }
  const r = await fetch(url, opts);
  if (!r.ok) {
    const e = await r.json().catch(() => ({ error: r.statusText }));
    throw new Error(e.error || r.statusText);
  }
  return r.json();
}
